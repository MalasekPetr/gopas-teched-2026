import { Router } from "express";
import axios from "axios";
import { getAppToken } from "../services/graphService.js";
import { getActiveDefinitions } from "../services/definitionService.js";

export const manageSubscriptionsRouter = Router();

interface SpSubscription {
  id: string;
  resource: string;
  notificationUrl: string;
  expirationDateTime: string;
}

// STEP 2: Ensure a webhook subscription exists on every list referenced by an
// active definition. Idempotent — safe to run on a cron or by hand.
manageSubscriptionsRouter.get("/", async (_req, res) => {
  try {
    const defs = await getActiveDefinitions();

    // One subscription per distinct (site,list) pair, not per definition.
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
  const token = await getAppToken("sharepoint");
  const headers = {
    Authorization: `Bearer ${token}`,
    Accept: "application/json;odata=nometadata",
    "Content-Type": "application/json",
  };

  // NOTE for presentation: subscriptions expire after a max of 6 months —
  // you MUST renew (PATCH) or recreate them, otherwise notifications stop silently.
  const expiration = new Date(Date.now() + 180 * 24 * 60 * 60 * 1000).toISOString();
  const listEndpoint = `${siteUrl}/_api/web/lists(guid'${listId}')/subscriptions`;

  const existing = await axios.get<{ value: SpSubscription[] }>(listEndpoint, { headers });
  const ours = existing.data.value.find((s) => s.notificationUrl === process.env.WEBHOOK_NOTIFICATION_URL);

  if (ours) {
    await axios.patch(
      `${listEndpoint}('${ours.id}')`,
      { expirationDateTime: expiration },
      { headers }
    );
    return { listId, action: "renewed", id: ours.id };
  }

  const created = await axios.post<SpSubscription>(
    listEndpoint,
    {
      resource: `${siteUrl}/_api/web/lists(guid'${listId}')`,
      notificationUrl: process.env.WEBHOOK_NOTIFICATION_URL,
      expirationDateTime: expiration,
    },
    { headers }
  );
  return { listId, action: "created", id: created.data.id };
}
