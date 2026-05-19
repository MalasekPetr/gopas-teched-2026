import { graphGet } from "./graphService.js";
import type { AlertDefinition } from "./definitionService.js";
import { sendAlertEmail } from "./mailService.js";

// Graph subscription notifications carry NO item data — just "this list changed".
// We call /items/delta, compare against the deltaLink we saved last time,
// and figure out what actually changed.

interface DeltaItem {
  id: string;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  ["@removed"]?: { reason: string };
}

interface DeltaPage {
  value: DeltaItem[];
  ["@odata.deltaLink"]?: string;
  ["@odata.nextLink"]?: string;
}

// In-memory deltaLink store. Production note: this MUST be persisted
// (SPO list item, Azure Table, Cosmos) — container restarts will reset it.
const deltaLinkStore = new Map<string, string>();
const keyFor = (siteId: string, listId: string): string => `${siteId}::${listId}`;

// Call once per (site, list) right after subscribing. Otherwise the very first
// edit produces a notification we resolve to 0 items, because delta?token=latest
// returns a deltaLink anchored at "now" — and "now" is after the edit.
export async function primeDeltaLink(siteId: string, listId: string): Promise<void> {
  const page = await graphGet<DeltaPage>(
    `/sites/${siteId}/lists/${listId}/items/delta?token=latest`
  );
  if (page["@odata.deltaLink"]) {
    deltaLinkStore.set(keyFor(siteId, listId), page["@odata.deltaLink"]);
  }
}

export async function processNotification(
  siteId: string,
  listId: string,
  matchingDefs: AlertDefinition[]
): Promise<void> {
  const saved = deltaLinkStore.get(keyFor(siteId, listId));

  // First call after a cold start: ask for "now" so we don't flood with backfill.
  const path = saved
    ? saved.replace(/^https:\/\/graph\.microsoft\.com\/v1\.0/, "")
    : `/sites/${siteId}/lists/${listId}/items/delta?token=latest`;

  const page = await graphGet<DeltaPage>(path);

  if (page["@odata.deltaLink"]) {
    deltaLinkStore.set(keyFor(siteId, listId), page["@odata.deltaLink"]);
  }

  for (const item of page.value) {
    const change = classifyChange(item);
    if (!change) continue;

    for (const def of matchingDefs) {
      if (def.notifyOn !== "All" && def.notifyOn !== change) continue;
      await sendAlertEmail(
        def.userEmail,
        listId,
        change,
        Number(item.id),
        item.lastModifiedDateTime ?? new Date().toISOString()
      ).catch((e) => console.error("sendAlertEmail failed", e?.response?.data ?? e));
    }
  }
}

// Graph delta has no explicit change-type field — we infer from timestamps.
// createdDateTime ≈ lastModifiedDateTime ⇒ the item was just created.
function classifyChange(item: DeltaItem): "New items" | "Modified" | "Deleted" | null {
  if (item["@removed"]) return "Deleted";
  if (!item.createdDateTime || !item.lastModifiedDateTime) return "Modified";
  const created = Date.parse(item.createdDateTime);
  const modified = Date.parse(item.lastModifiedDateTime);
  return Math.abs(modified - created) < 5000 ? "New items" : "Modified";
}
