import { Router } from "express";
import { getActiveDefinitions } from "../services/definitionService.js";
import { processNotification } from "../services/changeService.js";

export const webhookRouter = Router();

interface SpNotification {
  subscriptionId: string;
  resource: string; // list GUID
  siteUrl: string; // server-relative path of the site
  tenantId: string;
}

webhookRouter.post("/", async (req, res) => {
  // STEP 1: Handshake — SharePoint POSTs with ?validationToken=... and expects
  // the raw token echoed back as text/plain within 5 seconds. No auth, no body.
  const validationToken = req.query.validationToken;
  if (typeof validationToken === "string") {
    res.set("Content-Type", "text/plain").status(200).send(validationToken);
    return;
  }

  // Real notification — ack FAST, then process asynchronously.
  // SharePoint retries if we don't return within 5 seconds.
  res.status(202).end();

  const notifications: SpNotification[] = req.body?.value ?? [];
  if (!notifications.length) return;

  try {
    const defs = await getActiveDefinitions();
    for (const n of notifications) {
      // siteUrl in the payload is server-relative — combine with tenant host.
      const absoluteSiteUrl = `${process.env.SHAREPOINT_TENANT}${n.siteUrl}`;
      const matching = defs.filter(
        (d) => d.watchedListId.toLowerCase() === n.resource.toLowerCase()
      );
      if (!matching.length) continue;
      await processNotification(absoluteSiteUrl, n.resource, matching);
    }
  } catch (err) {
    console.error("webhook processing failed", err);
  }
});
