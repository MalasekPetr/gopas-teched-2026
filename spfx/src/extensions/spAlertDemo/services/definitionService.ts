import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";

// One row in SPAlerts_Definitions — kept here so the panel doesn't import server types.
export interface AlertDefinitionItem {
  Id: number;
  Title: string;
  WatchedListId: string;
  WatchedSiteUrl: string;
  UserEmail: string;
  NotifyOn: "All" | "New items" | "Modified" | "Deleted";
  IsActive: boolean;
}

const LIST_TITLE = "SPAlerts_Definitions";

export class DefinitionService {
  public constructor(private http: SPHttpClient, private siteUrl: string) {}

  // STEP B: Look up an existing definition for THIS list + user combo.
  // Same user setting an alert twice on the same list edits the same row.
  public async findForList(listId: string, userEmail: string): Promise<AlertDefinitionItem | null> {
    const filter =
      `WatchedListId eq '${listId}' and UserEmail eq '${userEmail.replace(/'/g, "''")}'`;
    const url =
      `${this.siteUrl}/_api/web/lists/getbytitle('${LIST_TITLE}')/items` +
      `?$filter=${encodeURIComponent(filter)}&$top=1`;

    const res = await this.http.get(url, SPHttpClient.configurations.v1);
    const json = await this.readJson(res);
    return (json.value && json.value[0]) || null;
  }

  public async create(item: Omit<AlertDefinitionItem, "Id">): Promise<void> {
    const url = `${this.siteUrl}/_api/web/lists/getbytitle('${LIST_TITLE}')/items`;
    await this.http.post(url, SPHttpClient.configurations.v1, {
      headers: { "Content-Type": "application/json;odata=nometadata", Accept: "application/json;odata=nometadata" },
      body: JSON.stringify(item),
    });
  }

  public async remove(id: number): Promise<void> {
    const url = `${this.siteUrl}/_api/web/lists/getbytitle('${LIST_TITLE}')/items(${id})`;
    await this.http.post(url, SPHttpClient.configurations.v1, {
      headers: { "X-HTTP-Method": "DELETE", "IF-MATCH": "*" },
    });
  }

  private async readJson(res: SPHttpClientResponse): Promise<any> {
    if (!res.ok) throw new Error(`SP request failed: ${res.status} ${res.statusText}`);
    return res.json();
  }
}
