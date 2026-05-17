import { graphGet } from "./graphService.js";

// One row in SPAlerts_Definitions. See README for the column schema.
export interface AlertDefinition {
  id: string;
  watchedListId: string;
  watchedSiteUrl: string;
  userEmail: string;
  notifyOn: "All" | "New items" | "Modified" | "Deleted";
  isActive: boolean;
}

// Read all active definitions. The definitions list lives on a single known site
// (DEFINITIONS_SITE_URL) — we resolve its Graph site-id once per call.
export async function getActiveDefinitions(): Promise<AlertDefinition[]> {
  const site = await resolveSite();
  const list = await resolveList(site.id);

  // expand=fields gets the column values without a second round-trip.
  const res = await graphGet<{ value: any[] }>(
    `/sites/${site.id}/lists/${list.id}/items?expand=fields&$top=999`
  );

  return res.value
    .map((it) => ({
      id: it.id,
      watchedListId: it.fields.WatchedListId,
      watchedSiteUrl: it.fields.WatchedSiteUrl,
      userEmail: it.fields.UserEmail,
      notifyOn: it.fields.NotifyOn,
      isActive: it.fields.IsActive === true,
    }))
    .filter((d) => d.isActive && d.watchedListId && d.watchedSiteUrl);
}

async function resolveSite(): Promise<{ id: string }> {
  // Graph wants {hostname}:{server-relative-path}, not the absolute URL.
  const url = new URL(process.env.DEFINITIONS_SITE_URL!);
  const path = url.pathname.replace(/\/$/, "");
  return graphGet<{ id: string }>(`/sites/${url.hostname}:${path}`);
}

async function resolveList(siteId: string): Promise<{ id: string }> {
  const title = process.env.DEFINITIONS_LIST_TITLE!;
  const res = await graphGet<{ value: { id: string; displayName: string }[] }>(
    `/sites/${siteId}/lists?$filter=displayName eq '${title}'`
  );
  if (!res.value.length) throw new Error(`List '${title}' not found`);
  return { id: res.value[0].id };
}
