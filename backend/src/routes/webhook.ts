import { Router } from "express";
import { getActiveDefinitions } from "../services/definitionService.js";
import { processNotification } from "../services/changeService.js";

export const webhookRouter = Router();

interface GraphNotification {
  subscriptionId: string;
  resource: string; // "sites/{site-id}/lists/{list-id}"
  clientState?: string;
  tenantId: string;
}

webhookRouter.post("/", async (req, res) => {
  // STEP 1: Handshake — Graph POSTs with ?validationToken=... and expects the
  // raw token echoed back as text/plain within 10 seconds.
  const validationToken = req.query.validationToken;
  if (typeof validationToken === "string") {
    res.set("Content-Type", "text/plain").status(200).send(validationToken);
    return;
  }

  // Real notification — ack FAST, then process asynchronously.
  res.status(202).end();

  const notifications: GraphNotification[] = req.body?.value ?? [];
  if (!notifications.length) return;

  try {
    const defs = await getActiveDefinitions();
    for (const n of notifications) {
      const parsed = parseResource(n.resource);
      if (!parsed) continue;
      const matching = defs.filter(
        (d) => d.watchedListId.toLowerCase() === parsed.listId.toLowerCase()
      );
      if (!matching.length) continue;
      await processNotification(parsed.siteId, parsed.listId, matching);
    }
  } catch (err) {
    console.error("webhook processing failed", err);
  }
});

// Graph subscription resource looks like
// "sites/contoso.sharepoint.com,GUID,GUID/lists/GUID"
function parseResource(resource: string): { siteId: string; listId: string } | null {
  const m = /^sites\/([^/]+)\/lists\/([^/]+)/.exec(resource);
  return m ? { siteId: m[1], listId: m[2] } : null;
}
