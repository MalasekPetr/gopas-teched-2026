import { Router } from "express";
import { graphGet, graphPatch, graphPost } from "../services/graphService.js";
import { getActiveDefinitions } from "../services/definitionService.js";
import { primeDeltaLink } from "../services/changeService.js";

export const manageSubscriptionsRouter = Router();

interface GraphSubscription {
  id: string;
  resource: string;
  notificationUrl: string;
  expirationDateTime: string;
}

// Graph caps SharePoint list subscriptions at 42,300 minutes (~29.4 days).
// We use 28 to leave headroom; a real app would re-renew before expiry on a cron.
const EXPIRATION_DAYS = 28;

// STEP 2: Ensure a Graph webhook subscription exists for every list referenced
// by an active definition. Idempotent — safe to run on a cron or by hand.
manageSubscriptionsRouter.get("/", async (_req, res) => {
  try {
    const defs = await getActiveDefinitions();

    // One subscription per distinct (site, list), not per definition.
    const targets = new Map<string, { siteUrl: string; listId: string }>();
    for (const d of defs) {
      targets.set(`${d.watchedSiteUrl}::${d.watchedListId}`, {
        siteUrl: d.watchedSiteUrl,
        listId: d.watchedListId,
      });
    }

    const results: unknown[] = [];
    for (const t of targets.values()) {
      results.push(await ensureSubscription(t.siteUrl, t.listId));
    }
    res.json({ ensured: results.length, results });
  } catch (err: any) {
    console.error("manage-subscriptions failed", err?.response?.data ?? err);
    res.status(500).json({ error: String(err?.message ?? err) });
  }
});

async function ensureSubscription(siteUrl: string, listId: string) {
  // Graph needs the composite site id, not the URL — resolve it first.
  const url = new URL(siteUrl);
  const path = url.pathname.replace(/\/$/, "");
  const site = await graphGet<{ id: string }>(`/sites/${url.hostname}:${path}`);
  const resource = `sites/${site.id}/lists/${listId}`;

  const expiration = new Date(Date.now() + EXPIRATION_DAYS * 24 * 60 * 60 * 1000).toISOString();
  const notificationUrl = process.env.WEBHOOK_NOTIFICATION_URL!;

  // Graph subscriptions are app-scoped — list all, filter for ours.
  const existing = await graphGet<{ value: GraphSubscription[] }>("/subscriptions");
  const ours = existing.value.find(
    (s) => s.resource === resource && s.notificationUrl === notificationUrl
  );

  if (ours) {
    await graphPatch(`/subscriptions/${ours.id}`, { expirationDateTime: expiration });
    await primeDeltaLink(site.id, listId);
    return { resource, action: "renewed", id: ours.id };
  }

  // Graph synchronously POSTs a validation token to notificationUrl before this returns.
  const created = await graphPost<GraphSubscription>("/subscriptions", {
    changeType: "updated",
    notificationUrl,
    resource,
    expirationDateTime: expiration,
    clientState: "sp-alert-demo",
  });
  await primeDeltaLink(site.id, listId);
  return { resource, action: "created", id: created.id };
}
