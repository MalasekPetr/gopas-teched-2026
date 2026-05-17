import axios from "axios";
import { getAppToken } from "./graphService.js";
import type { AlertDefinition } from "./definitionService.js";
import { sendAlertEmail } from "./mailService.js";

// STEP 3: Webhook payloads only say "list X changed" — they carry NO item data.
// We must call SharePoint's getchanges API ourselves with the last ChangeToken.

// In-memory ChangeToken store. Production note: this MUST be persisted
// (SPO list item, Azure Table, Cosmos) — restarts will replay or miss changes.
const tokenStore = new Map<string, string>();

interface SpChange {
  ChangeType: number; // 1=Add, 2=Update, 3=DeleteObject
  ItemId: number;
  Time: string;
  Editor?: { LookupValue?: string };
}

const CHANGE_TYPE_MAP: Record<number, "New items" | "Modified" | "Deleted"> = {
  1: "New items",
  2: "Modified",
  3: "Deleted",
};

export async function processNotification(
  siteUrl: string,
  listId: string,
  matchingDefs: AlertDefinition[]
): Promise<void> {
  const token = await getAppToken("sharepoint");

  // First poll: no token yet — start watching from "now" so we don't flood on first install.
  const storeKey = `${siteUrl}::${listId}`;
  const lastToken = tokenStore.get(storeKey);

  const body = lastToken
    ? { query: { __metadata: { type: "SP.ChangeQuery" }, Item: true, ChangeTokenStart: { StringValue: lastToken } } }
    : { query: { __metadata: { type: "SP.ChangeQuery" }, Item: true } };

  const url = `${siteUrl}/_api/web/lists(guid'${listId}')/getchanges`;
  const { data } = await axios.post(url, body, {
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: "application/json;odata=verbose",
      "Content-Type": "application/json;odata=verbose",
    },
  });

  const changes: SpChange[] = data?.d?.results ?? [];
  if (!changes.length) return;

  // Last change carries the token we hand back on the next poll.
  const lastChange = changes[changes.length - 1] as any;
  if (lastChange?.ChangeToken?.StringValue) {
    tokenStore.set(storeKey, lastChange.ChangeToken.StringValue);
  }

  for (const change of changes) {
    const mapped = CHANGE_TYPE_MAP[change.ChangeType];
    if (!mapped) continue;

    // Fan out to every definition that wants this change type for this list.
    for (const def of matchingDefs) {
      if (def.notifyOn !== "All" && def.notifyOn !== mapped) continue;
      await sendAlertEmail(def.userEmail, listId, mapped, change.ItemId, change.Time).catch((e) =>
        console.error("sendAlertEmail failed", e?.response?.data ?? e)
      );
    }
  }
}
