# Step-by-step: Build a modern replacement for SharePoint alerts

A hands-on lab that walks you through building the entire app in this repo from an empty folder. The target reader is a developer who:

- Knows TypeScript / JavaScript and has used `npm`.
- Has a Microsoft 365 tenant and an Azure subscription (free tiers are fine).
- Has **never** built an SPFx extension, never written a SharePoint webhook receiver, and has never used Microsoft Graph for anything more than `/me`.

By the end you'll have everything in this repo running in your own tenant, and you'll know *why* each piece exists.

> **Time budget:** ~3 hours the first time. Skip a section by `git checkout`-ing the final code in this repo and reading just the narrative.

---

## Table of contents

1. [Background: why we're rebuilding alerts](#1-background-why-were-rebuilding-alerts)
2. [Architecture in one diagram](#2-architecture-in-one-diagram)
3. [Lab prerequisites](#3-lab-prerequisites)
4. [Provision the SharePoint site and lists](#4-provision-the-sharepoint-site-and-lists)
5. [Register the Entra app](#5-register-the-entra-app)
6. [Scaffold the SPFx solution](#6-scaffold-the-spfx-solution)
7. [Build the "Set Alert" ListView Command Set](#7-build-the-set-alert-listview-command-set)
8. [Build the React panel UI](#8-build-the-react-panel-ui)
9. [Scaffold the Node / Express backend](#9-scaffold-the-node--express-backend)
10. [App-only auth and the Graph helper](#10-app-only-auth-and-the-graph-helper)
11. [The webhook receiver and the validation handshake](#11-the-webhook-receiver-and-the-validation-handshake)
12. [Subscribe via Microsoft Graph](#12-subscribe-via-microsoft-graph)
13. [Detecting changes with `/items/delta`](#13-detecting-changes-with-itemsdelta)
14. [Sending the email with Graph `sendMail`](#14-sending-the-email-with-graph-sendmail)
15. [Local end-to-end with Dev Tunnels](#15-local-end-to-end-with-dev-tunnels)
16. [Deploy to Azure App Service](#16-deploy-to-azure-app-service)
17. [Ship the SPFx package tenant-wide](#17-ship-the-spfx-package-tenant-wide)
18. [What production would add](#18-what-production-would-add)

---

## 1. Background: why we're rebuilding alerts

The "Alert me when this list changes" feature in classic SharePoint sent emails directly from the SharePoint farm using its own SMTP path and a built-in subscription store. SharePoint Online has [announced its replacement](https://techcommunity.microsoft.com/discussions/sharepoint-developer/the-future-of-sharepoint-alerts/4400993) — newly created sites no longer expose it by default, and the long-term direction is "build your own with the modern platform pieces."

The replacement we're going to build follows the same conceptual flow but uses four standard, supported pieces:

| Job | Classic alerts | Our version |
| --- | --- | --- |
| Trigger UI | Out-of-box button in the list ribbon | [SPFx ListView Command Set](https://learn.microsoft.com/sharepoint/dev/spfx/extensions/get-started/building-simple-cmdset-with-dialog-api) — a button we add ourselves |
| Where the subscription lives | SP internal table | A SharePoint list named `SPAlerts_Definitions` |
| Detecting list changes | SP internal job | A [Microsoft Graph webhook subscription](https://learn.microsoft.com/graph/api/resources/webhooks) + `/items/delta` |
| Sending the email | SP's own SMTP | [Graph `sendMail`](https://learn.microsoft.com/graph/api/user-sendmail) from a service mailbox |

The trade-off: more moving parts, but every piece is documented, observable, and replaceable.

**Useful background reading:**

- [SharePoint Framework overview](https://learn.microsoft.com/sharepoint/dev/spfx/sharepoint-framework-overview)
- [Microsoft Graph change notifications (webhooks) overview](https://learn.microsoft.com/graph/webhooks)
- [OAuth 2.0 client credentials flow](https://learn.microsoft.com/entra/identity-platform/v2-oauth2-client-creds-grant-flow)

---

## 2. Architecture in one diagram

```text
        ┌─────────────────────┐ 1. user clicks "Set Alert" (delegated context)
        │ SPFx Command Set    │─────────┐
        │ (runs in the page)  │         │
        └─────────────────────┘         ▼
                                ┌─────────────────────────┐
                                │ SPAlerts_Definitions    │ (SP list)
                                │ row written via REST    │
                                └────────────┬────────────┘
                                             │ 2. backend reads via Graph
                                             ▼
                                ┌─────────────────────────┐
                                │ Azure App Service       │
                                │ (Node 22 / Express)     │
                                │                         │
                                │  /manage-subscriptions  │ 3. POST /v1.0/subscriptions
                                │  /webhook               │     to Microsoft Graph
                                └────────────┬────────────┘
                                             │
                                             ▼
                          ┌──────────────────────────────────┐
                          │ Microsoft Graph                  │
                          │  - validates /webhook URL        │
                          │  - sends notifications on change │
                          └────────────┬─────────────────────┘
                                       │ 4. POST /webhook
                                       ▼
                            ┌─────────────────────┐
                            │ Backend             │ 5. GET /items/delta
                            │                     │────────────────────────┐
                            └────────────┬────────┘                        │
                                         │ 6. POST /users/{x}/sendMail     ▼
                                         ▼                       ┌─────────────────────┐
                                  📧 Email to user               │ Microsoft Graph     │
                                                                 └─────────────────────┘
```

Three things to internalize before you start coding:

- **The SPFx panel runs as the signed-in user** (delegated). The backend runs **app-only**. Those are two entirely different security contexts.
- **Graph won't accept a brand-new subscription unless your `/webhook` URL is publicly reachable** and responds to a synchronous `validationToken` challenge within 10 seconds.
- **SharePoint REST `/_api/web/.../subscriptions` does NOT accept Entra (Azure AD) app-only tokens.** That's a footgun — there's a lot of old documentation and sample code that tells you to use it. The modern, supported path is Graph. We'll discuss it more in section 12.

---

## 3. Lab prerequisites

| Thing you need | How to get it |
| --- | --- |
| Microsoft 365 tenant (admin) | Sign up for a free [Microsoft 365 Developer Program](https://developer.microsoft.com/microsoft-365/dev-program) tenant — it includes 25 user licenses and a SharePoint site collection. |
| Azure subscription | Free trial: <https://azure.microsoft.com/free/>. We'll use one App Service plan (B1, ~€13/month — well below the free credit). |
| Node.js 22 LTS, x64 | <https://nodejs.org/> — SPFx 1.22 requires Node 22.14+ and ships heft (no Gulp, no Yeoman wrapper needed). |
| PowerShell 7+ | Bundled on Windows 11. Otherwise [install pwsh](https://learn.microsoft.com/powershell/scripting/install/installing-powershell). |
| Azure CLI 2.60+ | `winget install Microsoft.AzureCLI` ([docs](https://learn.microsoft.com/cli/azure/install-azure-cli)). Sign in with `az login`. |
| PnP PowerShell 3+ | `Install-Module PnP.PowerShell -Scope CurrentUser` ([docs](https://pnp.github.io/powershell/)). |
| A code editor | [VS Code](https://code.visualstudio.com/) with the [TypeScript](https://marketplace.visualstudio.com/items?itemName=ms-vscode.vscode-typescript-next) and [Azure App Service](https://marketplace.visualstudio.com/items?itemName=ms-azuretools.vscode-azureappservice) extensions. |
| Microsoft Graph Explorer | <https://developer.microsoft.com/graph/graph-explorer> — for poking at Graph endpoints while you build. |

Pin these two reference tabs — you'll come back to them constantly:

- [Microsoft Graph REST API v1.0 reference](https://learn.microsoft.com/graph/api/overview)
- [SharePoint Framework reference (1.22)](https://learn.microsoft.com/javascript/api/sp-listview-extensibility)

---

## 4. Provision the SharePoint site and lists

You need two lists on a brand new site collection. One holds the alert subscriptions, the other is the list users will actually edit during the demo.

```powershell
# Connect using the modern interactive flow (no client ID needed for this scope)
Connect-PnPOnline -Url "https://<tenant>-admin.sharepoint.com" -Interactive

# Create the demo site
New-PnPSite -Type TeamSite -Title "AlertsDemo" -Alias "AlertsDemo"

# Switch to the site and create the lists
Connect-PnPOnline -Url "https://<tenant>.sharepoint.com/sites/AlertsDemo" -Interactive

# "Demo" — the watched list (audience edits items here on stage)
New-PnPList -Title "Demo" -Template GenericList
Add-PnPField -List Demo -DisplayName "Description" -InternalName "Description" -Type Text

# "SPAlerts_Definitions" — one row per active alert
New-PnPList -Title "SPAlerts_Definitions" -Template GenericList
Add-PnPField -List SPAlerts_Definitions -DisplayName "WatchedListId"  -InternalName "WatchedListId"  -Type Text
Add-PnPField -List SPAlerts_Definitions -DisplayName "WatchedSiteUrl" -InternalName "WatchedSiteUrl" -Type Text
Add-PnPField -List SPAlerts_Definitions -DisplayName "UserEmail"      -InternalName "UserEmail"      -Type Text
Add-PnPField -List SPAlerts_Definitions -DisplayName "NotifyOn"       -InternalName "NotifyOn"       -Type Choice -Choices "All","New items","Modified","Deleted"
Add-PnPField -List SPAlerts_Definitions -DisplayName "IsActive"       -InternalName "IsActive"       -Type Boolean
```

**Why a SharePoint list and not a database?**

For a 1-hour demo, a list gives you a free admin UI, free audit log, no separate hosting, and zero database management. For production you'd swap it for [Azure Cosmos DB](https://learn.microsoft.com/azure/cosmos-db/), [Azure Table Storage](https://learn.microsoft.com/azure/storage/tables/), or a Dataverse table — the point is the architecture is the same.

References: [PnP PowerShell list cmdlets](https://pnp.github.io/powershell/cmdlets/New-PnPList.html), [SharePoint internal column names](https://learn.microsoft.com/sharepoint/dev/general-development/column-and-list-field-internal-names).

---

## 5. Register the Entra app

Go to [entra.microsoft.com](https://entra.microsoft.com/) → **Identity → Applications → App registrations → New registration**.

1. **Name** `sp-webhook-demo-teched`. Single-tenant. No redirect URI (the backend never receives an interactive login).
2. **Certificates & secrets → New client secret**, 6 months. **Copy the value immediately** — Entra never shows it again.
3. **API permissions → Add a permission**. Add both as **Application** permissions and click **Grant admin consent**:

    | API | Permission |
    | --- | --- |
    | Microsoft Graph | `Sites.ReadWrite.All` |
    | Microsoft Graph | `Mail.Send` |

4. **Overview** → copy **Application (client) ID** and **Directory (tenant) ID**.

You now have three secrets in your password manager: tenant ID, client ID, client secret.

> **Why no SharePoint API permission?**
> Earlier versions of this demo also asked for `SharePoint → Sites.FullControl.All`. We don't need it: the modern, supported path for subscribing to a list and reading deltas is Microsoft Graph. The legacy SP REST endpoints (`/_api/web/.../subscriptions`, `getchanges`) refuse Entra app-only tokens — they only accept ACS principals, which Microsoft is retiring.

References:

- [Quickstart: register an application with the Microsoft identity platform](https://learn.microsoft.com/entra/identity-platform/quickstart-register-app)
- [Limit application access to specific mailboxes](https://learn.microsoft.com/graph/auth-limit-mailbox-access) — strongly recommended before any production use of `Mail.Send` (Application).

---

## 6. Scaffold the SPFx solution

SPFx 1.22 dropped the Yeoman generator in favour of [Heft](https://heft.rushstack.io/), which is faster and avoids the global-install Yeoman tax. From an empty `spfx/` folder:

```powershell
mkdir spfx; cd spfx
npm init -y
npm install --save-exact @microsoft/sp-core-library@1.22.2 `
                         @microsoft/sp-http@1.22.2 `
                         @microsoft/sp-listview-extensibility@1.22.2 `
                         react@17.0.1 react-dom@17.0.1 `
                         @fluentui/react@^8.122.0 tslib@2.3.1
npm install --save-dev   @microsoft/eslint-config-spfx@1.22.2 `
                         @microsoft/spfx-heft-plugins@1.22.2 `
                         @microsoft/spfx-web-build-rig@1.22.2 `
                         @rushstack/heft@1.1.2 `
                         @types/react@17.0.45 @types/react-dom@17.0.17 `
                         typescript@~5.8.0
```

Easier alternative: copy the [spfx/](../spfx/) folder from this repo and run `npm install`.

The pieces of an SPFx 1.22 project that matter for this lab:

| Path | Purpose |
| --- | --- |
| [spfx/config/package-solution.json](../spfx/config/package-solution.json) | Solution metadata. `skipFeatureDeployment: true` enables tenant-wide deploy. |
| [spfx/config/config.json](../spfx/config/config.json) | Bundle definition — which component goes in which bundle file. |
| [spfx/src/extensions/spAlertDemo/SpAlertDemoCommandSet.manifest.json](../spfx/src/extensions/spAlertDemo/SpAlertDemoCommandSet.manifest.json) | Manifest of the command set itself — declares the `SET_ALERT` command and its icon. |

References:

- [Get started building your first SPFx extension](https://learn.microsoft.com/sharepoint/dev/spfx/extensions/get-started/build-a-hello-world-extension)
- [SPFx 1.22 release notes](https://learn.microsoft.com/sharepoint/dev/spfx/release-1.22)
- [SPFx samples on GitHub (PnP)](https://github.com/pnp/sp-dev-fx-extensions) — hundreds of working extension samples.

---

## 7. Build the "Set Alert" ListView Command Set

A [ListView Command Set](https://learn.microsoft.com/sharepoint/dev/spfx/extensions/get-started/building-simple-cmdset-with-dialog-api) is the SPFx way to add a button to the command bar of any SharePoint list. The whole entry-point class lives in [spfx/src/extensions/spAlertDemo/SpAlertDemoCommandSet.ts](../spfx/src/extensions/spAlertDemo/SpAlertDemoCommandSet.ts) — about 40 lines.

```typescript
public onInit(): Promise<void> {
  const cmd: Command = this.tryGetCommand("SET_ALERT");
  if (cmd) cmd.visible = true;
  return Promise.resolve();
}

public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
  if (event.itemId !== "SET_ALERT") return;

  const pageCtx = this.context.pageContext;
  ReactDOM.render(
    React.createElement(AlertPanel, {
      siteUrl: pageCtx.web.absoluteUrl,
      listId: pageCtx.list!.id.toString(),
      listTitle: pageCtx.list!.title,
      currentUserEmail: pageCtx.user.email,
      spHttpClient: this.context.spHttpClient,
      onClose: () => ReactDOM.unmountComponentAtNode(this.panelContainer!),
    }),
    this.panelContainer
  );
}
```

Things to notice:

- `this.context.pageContext` gives you the **current site URL, list id, user email** — exactly what we need to write into `SPAlerts_Definitions`.
- `this.context.spHttpClient` is an authenticated HTTP client that handles the SharePoint request digest, OAuth token, and CSRF for you. Use it for any call to `/_api/`. ([reference](https://learn.microsoft.com/javascript/api/sp-http/sphttpclient))

References: [PageContext API](https://learn.microsoft.com/javascript/api/sp-page-context), [Command Set base class](https://learn.microsoft.com/javascript/api/sp-listview-extensibility/baselistviewcommandset).

---

## 8. Build the React panel UI

We use a [Fluent UI Panel](https://developer.microsoft.com/fluentui#/controls/web/panel) for the form — it gives you keyboard focus management, mobile-friendly layout, and free accessibility. Code: [spfx/src/extensions/spAlertDemo/components/AlertPanel.tsx](../spfx/src/extensions/spAlertDemo/components/AlertPanel.tsx).

The three-field form:

```typescript
<Panel isOpen headerText="Set Alert" type={PanelType.medium} onDismiss={props.onClose}>
  <Stack tokens={{ childrenGap: 12 }}>
    <TextField label="List" value={props.listTitle} readOnly />
    <Dropdown
      label="Notify me about"
      selectedKey={notifyOn}
      options={NOTIFY_OPTIONS}
      onChange={(_, o) => o && setNotifyOn(o.key as AlertDefinitionItem["NotifyOn"])}
    />
    <TextField label="Your email" value={email} onChange={(_, v) => setEmail(v || "")} />
    <Stack horizontal tokens={{ childrenGap: 8 }}>
      <PrimaryButton text="Save Alert" onClick={save} />
      {existing && <DefaultButton text="Remove Alert" onClick={remove} />}
      <DefaultButton text="Cancel" onClick={props.onClose} />
    </Stack>
  </Stack>
</Panel>
```

The save flow is intentionally simple: look up an existing definition for this `(list, user)`, delete it if there is one, then create a fresh row. Idempotent and easy to demo. See [DefinitionService](../spfx/src/extensions/spAlertDemo/services/definitionService.ts).

References:

- [Fluent UI React Northstar / v8 docs](https://developer.microsoft.com/fluentui)
- [React 17 patterns SPFx still uses](https://learn.microsoft.com/sharepoint/dev/spfx/release-1.18) — note that SPFx is *not* on React 18 yet; don't reach for `createRoot`.

---

## 9. Scaffold the Node / Express backend

Backend [package.json](../backend/package.json) is intentionally tiny — three runtime dependencies:

```json
{
  "type": "module",
  "scripts": { "build": "tsc", "start": "node dist/server.js", "dev": "tsx watch src/server.ts" },
  "dependencies": {
    "axios": "^1.7.7",
    "dotenv": "^16.4.5",
    "express": "^4.21.1"
  }
}
```

[server.ts](../backend/src/server.ts) — 20 lines, two routes plus a root health check:

```typescript
import "dotenv/config";
import express from "express";
import { webhookRouter } from "./routes/webhook.js";
import { manageSubscriptionsRouter } from "./routes/manageSubscriptions.js";

const app = express();
app.use(express.json());

app.use("/webhook", webhookRouter);
app.use("/manage-subscriptions", manageSubscriptionsRouter);
app.get("/", (_req, res) => res.send("SP Webhook Demo backend up"));

const port = Number(process.env.PORT) || 8080;
app.listen(port, () => console.log(`Listening on ${port}`));
```

References: [Express 4 routing guide](https://expressjs.com/en/guide/routing.html), [tsx — TypeScript Execute](https://github.com/privatenumber/tsx).

---

## 10. App-only auth and the Graph helper

The whole backend talks to Microsoft Graph using **one** app-only token. The full helper is [backend/src/services/graphService.ts](../backend/src/services/graphService.ts) — 50 lines total.

The core is the OAuth 2 [client-credentials grant](https://learn.microsoft.com/entra/identity-platform/v2-oauth2-client-creds-grant-flow):

```typescript
async function getAppToken(): Promise<string> {
  const now = Date.now();
  if (cachedToken && cachedToken.expiresAt > now + 60_000) return cachedToken.value;

  const url = `https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    client_id: process.env.CLIENT_ID!,
    client_secret: process.env.CLIENT_SECRET!,
    scope: "https://graph.microsoft.com/.default",
    grant_type: "client_credentials",
  });

  const { data } = await axios.post(url, body.toString(), {
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
  });

  cachedToken = { value: data.access_token, expiresAt: now + data.expires_in * 1000 };
  return data.access_token;
}
```

Things to know:

- `scope: "https://graph.microsoft.com/.default"` asks for **every** consented Application permission this app has on Graph. There is no fine-grained per-request scoping in the client-credentials flow.
- Tokens last 1 hour. We cache with a 1-minute safety margin so we never use a token that's about to expire mid-request.
- Production note: don't ship the client secret in source. Use a certificate, then a [Managed Identity](https://learn.microsoft.com/entra/identity/managed-identities-azure-resources/overview), and store anything else in [Key Vault](https://learn.microsoft.com/azure/key-vault/general/overview).

References:

- [OAuth 2.0 client credentials flow](https://learn.microsoft.com/entra/identity-platform/v2-oauth2-client-creds-grant-flow)
- [Microsoft Graph access tokens](https://learn.microsoft.com/graph/auth-v2-service)
- [Microsoft Authentication Library (MSAL) for Node](https://learn.microsoft.com/entra/msal/node/) — what you'd use in production instead of hand-rolled `axios.post`.

---

## 11. The webhook receiver and the validation handshake

Microsoft Graph's [webhook validation flow](https://learn.microsoft.com/graph/webhooks#notification-endpoint-validation) is simple but strict:

1. When you `POST /v1.0/subscriptions`, Graph immediately makes an HTTP POST to your `notificationUrl` with a query parameter `validationToken=<opaque>`.
2. Your endpoint must reply `200 OK`, `Content-Type: text/plain`, body = the **URL-decoded** token, **within 10 seconds**.
3. If you fail to echo the token, the subscription creation fails — Graph returns a 400 to *your* `/manage-subscriptions` call.

[backend/src/routes/webhook.ts](../backend/src/routes/webhook.ts) handles both the handshake and real notifications in the same `POST` handler:

```typescript
webhookRouter.post("/", async (req, res) => {
  // STEP 1: Handshake
  const validationToken = req.query.validationToken;
  if (typeof validationToken === "string") {
    res.set("Content-Type", "text/plain").status(200).send(validationToken);
    return;
  }

  // Real notification — ack FAST, then process asynchronously.
  res.status(202).end();

  const notifications: GraphNotification[] = req.body?.value ?? [];
  // ... look up definitions, call /items/delta, send email ...
});
```

The `202` is critical: Graph retries if you take longer than 10 seconds, but it does NOT care what your response body is. Acknowledge the receipt, then do real work in the background.

References:

- [Receive change notifications through webhooks](https://learn.microsoft.com/graph/webhooks)
- [Express body parsers — `express.json()`](https://expressjs.com/en/api.html#express.json)

---

## 12. Subscribe via Microsoft Graph

This is the section where most blog posts you'll find are out of date.

**The old way** (still in many samples): `POST {siteUrl}/_api/web/lists(guid'X')/subscriptions` with an Entra app-only token → `401 Unsupported app only token`. That endpoint *only* accepts legacy [ACS app principals](https://learn.microsoft.com/sharepoint/dev/sp-add-ins/creating-sharepoint-add-ins-that-use-app-only-permissions), which Microsoft is retiring (no new ACS principals on new tenants).

**The modern way** ([Graph webhooks on a SharePoint list](https://learn.microsoft.com/graph/webhooks-with-resource-data)):

```http
POST https://graph.microsoft.com/v1.0/subscriptions
Authorization: Bearer <app-only-graph-token>
Content-Type: application/json

{
  "changeType": "updated",
  "notificationUrl": "https://<your-app>.azurewebsites.net/webhook",
  "resource": "sites/{site-id}/lists/{list-id}",
  "expirationDateTime": "2026-06-16T00:00:00Z",
  "clientState": "sp-alert-demo"
}
```

Important constraints:

- `resource` for a SharePoint list is `sites/{site-id}/lists/{list-id}` where `site-id` is the **composite Graph site id** (`hostname,siteGuid,webGuid`), not the SP URL or the site collection GUID. Resolve it with `GET /sites/{hostname}:/sites/<sitePath>`.
- `changeType` for SharePoint resources must be `updated` — there's no separate `created`/`deleted` notification stream.
- **Max expiration is 42,300 minutes (~29.4 days).** Renew before then. We use 28 days in the code to leave headroom.
- `clientState` is echoed back in every notification — you can use it as a shared secret to defend against spoofed callbacks.

Full code: [backend/src/routes/manageSubscriptions.ts](../backend/src/routes/manageSubscriptions.ts). The route is idempotent — it lists existing subscriptions and either `PATCH`es one to renew or `POST`s a new one.

References:

- [Subscription resource type](https://learn.microsoft.com/graph/api/resources/subscription) — table of `resource` formats and max lifetimes.
- [Create a subscription](https://learn.microsoft.com/graph/api/subscription-post-subscriptions)
- [Update (renew) a subscription](https://learn.microsoft.com/graph/api/subscription-update)
- [Subscription lifecycle notifications](https://learn.microsoft.com/graph/webhooks-lifecycle) — what to do when Graph tells you a subscription is about to expire or needs reauthorization.

---

## 13. Detecting changes with `/items/delta`

A Graph webhook notification body looks like this:

```json
{
  "value": [{
    "subscriptionId": "...",
    "clientState": "sp-alert-demo",
    "resource": "sites/contoso.sharepoint.com,...,.../lists/abc...",
    "changeType": "updated",
    "subscriptionExpirationDateTime": "...",
    "tenantId": "..."
  }]
}
```

Note what's **missing**: there's no item ID, no field values, no "what changed". By design — payloads are tiny and the same shape regardless of how much actually changed. You have to ask.

The "ask" is [`GET /sites/{site-id}/lists/{list-id}/items/delta`](https://learn.microsoft.com/graph/api/listitem-delta). The first call gives you a `@odata.deltaLink`. Save it. Next call, hit that deltaLink — you get only items that have changed since.

The wrinkle: **on the very first call after creating a subscription, you don't have a deltaLink yet.** If you let it default, Graph returns **every** existing item in the list, which floods the email mailbox. The fix is `?token=latest` — returns 0 items plus a deltaLink anchored at "now". We call this `primeDeltaLink` and run it inside `/manage-subscriptions` so the very first edit afterwards is the first thing the delta returns.

Code: [backend/src/services/changeService.ts](../backend/src/services/changeService.ts). The classifier maps delta items to the demo's four `NotifyOn` choices:

```typescript
function classifyChange(item: DeltaItem): "New items" | "Modified" | "Deleted" | null {
  if (item["@removed"]) return "Deleted";
  if (!item.createdDateTime || !item.lastModifiedDateTime) return "Modified";
  const created = Date.parse(item.createdDateTime);
  const modified = Date.parse(item.lastModifiedDateTime);
  return Math.abs(modified - created) < 5000 ? "New items" : "Modified";
}
```

Delta has no per-change-type field — we infer "new" from "was created and modified within the same 5 seconds." Good enough for a demo; production may want a more authoritative source.

References:

- [Track changes for items in a list (`delta`)](https://learn.microsoft.com/graph/api/listitem-delta)
- [Microsoft Graph deltaLink / delta token concept](https://learn.microsoft.com/graph/delta-query-overview)
- [Best practices for discovering files and detecting changes at scale](https://learn.microsoft.com/graph/files-detect-changes) — written for OneDrive but the patterns map 1:1 to SharePoint lists.

---

## 14. Sending the email with Graph `sendMail`

[backend/src/services/mailService.ts](../backend/src/services/mailService.ts) — about 30 lines:

```typescript
await graphPost(`/users/${sender}/sendMail`, {
  message: {
    subject,
    body: { contentType: "HTML", content: html },
    toRecipients: [{ emailAddress: { address: recipient } }],
  },
  saveToSentItems: false,
});
```

Things to know:

- `sender` is the **mailbox to send AS** — the app reg needs `Mail.Send` (Application) and the [Exchange Application Access Policy](https://learn.microsoft.com/graph/auth-limit-mailbox-access) for that mailbox.
- Without the access policy, `Mail.Send` (Application) lets the app send as **any** user. That's why you scope it.
- `saveToSentItems: false` keeps your service mailbox clean during demos.

References:

- [`POST /users/{id}/sendMail`](https://learn.microsoft.com/graph/api/user-sendmail)
- [Limiting application permissions to specific mailboxes](https://learn.microsoft.com/graph/auth-limit-mailbox-access) — read it twice.

---

## 15. Local end-to-end with Dev Tunnels

Graph won't subscribe to `http://localhost:8080` — it needs a public HTTPS URL it can validate. Microsoft Dev Tunnels is the easiest way:

```powershell
# Install (one-time)
winget install Microsoft.devtunnel

# Sign in (uses your work/school account)
devtunnel user login

# Forward your local backend
devtunnel host -p 8080 --allow-anonymous
```

It prints a `https://<random>.devtunnels.ms` URL. Update your `.env`:

```env
WEBHOOK_NOTIFICATION_URL=https://<random>.devtunnels.ms/webhook
```

Run the backend with `npm run dev` in `backend/`. Test the validation handshake before subscribing for real:

```powershell
Invoke-RestMethod -Method POST "https://<random>.devtunnels.ms/webhook?validationToken=hello"
# expect: hello
```

Then create the subscription:

```powershell
Invoke-RestMethod "https://<random>.devtunnels.ms/manage-subscriptions"
# expect: { ensured: 1, results: [{ action: "created", ... }] }
```

Edit an item in the watched SharePoint list — within ~30 seconds you should see a `POST /webhook` line in your `npm run dev` console followed by `sendMail` and an email in your inbox.

Alternatives if Dev Tunnels isn't available: [ngrok](https://ngrok.com/), [Cloudflare Tunnel](https://developers.cloudflare.com/cloudflare-one/connections/connect-networks/), [localtunnel](https://localtunnel.github.io/www/).

References:

- [Microsoft Dev Tunnels docs](https://learn.microsoft.com/azure/developer/dev-tunnels/overview)
- [Debug SPFx extensions locally](https://learn.microsoft.com/sharepoint/dev/spfx/extensions/get-started/serving-your-extension-from-localhost) — required if you also want hot-reload on the SPFx side.

---

## 16. Deploy to Azure App Service

```powershell
$rg     = "rg-spwebhook-teched"
$plan   = "plan-spwebhook-teched"
$app    = "spwebhook-teched-$([Guid]::NewGuid().ToString('N').Substring(0,6))"
$region = "westeurope"

az group create -n $rg -l $region
az appservice plan create -g $rg -n $plan --sku B1 --is-linux
az webapp create -g $rg -p $plan -n $app --runtime "NODE:22-lts"
```

Set every env var [in its own `az` call](../docs/setup.md#4-azure-app-service) — bundling them with PowerShell backtick line continuation is the #1 cause of "Invalid URL" errors on this app.

Build, prune, zip, deploy:

```powershell
cd backend
npm install
npm run build
npm prune --production
Compress-Archive -Path dist,node_modules,package.json,package-lock.json `
                 -DestinationPath backend.zip -Force
az webapp deploy -g $rg -n $app --src-path backend.zip --type zip
```

Smoke tests:

```powershell
Invoke-RestMethod "https://$app.azurewebsites.net/"
# expect: "SP Webhook Demo backend up"

Invoke-RestMethod "https://$app.azurewebsites.net/manage-subscriptions"
# expect: { ensured: 0, results: @() }  -- before any alerts are saved
```

If something doesn't work, the fastest debug path is `az webapp log tail -g $rg -n $app` and trigger the failing endpoint in another window.

References:

- [Quickstart: deploy a Node.js web app](https://learn.microsoft.com/azure/app-service/quickstart-nodejs)
- [Configure a Node.js app for Azure App Service](https://learn.microsoft.com/azure/app-service/configure-language-nodejs)
- [Enable diagnostic logs in App Service](https://learn.microsoft.com/azure/app-service/troubleshoot-diagnostic-logs)

---

## 17. Ship the SPFx package tenant-wide

```powershell
cd ..\spfx
npm install
npm run build      # produces sharepoint/solution/sp-alert-demo.sppkg
```

1. Open `https://<tenant>.sharepoint.com/sites/appcatalog`.
2. **Apps for SharePoint** → drag `sp-alert-demo.sppkg` in → tick **Make this solution available to all sites in the organization** → **Deploy**.
3. Open `https://<tenant>.sharepoint.com/sites/appcatalog/Lists/TenantWideExtensions/AllItems.aspx` and add a row:

    | Field | Value |
    | --- | --- |
    | Title | `Set Alert (TechEd 2026 demo)` |
    | Component Id | the `id` from [SpAlertDemoCommandSet.manifest.json](../spfx/src/extensions/spAlertDemo/SpAlertDemoCommandSet.manifest.json) |
    | Component Properties | `{}` |
    | List Template | `100` *(generic lists; add a second row with `101` for doc libs)* |
    | Location | `ClientSideExtension.ListViewCommandSet` |
    | Sequence | `100` |
    | Disabled | `No` |

4. Hard-refresh the watched list page in a private window — **Set Alert** appears in the toolbar within ~5 minutes.

**Why the manual row?** SPFx will auto-create that row for you *if* your package contains a `sharepoint/assets/ClientSideInstance.xml`. Adding that asset and re-shipping the package is the cleaner long-term answer — left as an exercise. See [Tenant-scoped deployment of SharePoint Framework solutions](https://learn.microsoft.com/sharepoint/dev/spfx/tenant-scoped-deployment).

References:

- [Manage app permissions in modern SharePoint](https://learn.microsoft.com/sharepoint/manage-app-permissions)
- [Set up your SharePoint Framework development environment](https://learn.microsoft.com/sharepoint/dev/spfx/set-up-your-development-environment)

---

## 18. What production would add

Everything in this repo is sized for a 1-hour conference demo. To run it for real you'd want:

| Concern | What to add |
| --- | --- |
| Subscription renewal | A timer ([Azure Function timer](https://learn.microsoft.com/azure/azure-functions/functions-bindings-timer), [GitHub Actions schedule](https://docs.github.com/actions/using-workflows/events-that-trigger-workflows#schedule), Logic App recurrence) that calls `/manage-subscriptions` every ~28 days. |
| Durable deltaLink storage | Move `deltaLinkStore` from in-memory to [Cosmos DB](https://learn.microsoft.com/azure/cosmos-db/), [Azure Table Storage](https://learn.microsoft.com/azure/storage/tables/table-storage-overview), or a hidden SP list. |
| Auth | Replace client secret with a [certificate](https://learn.microsoft.com/entra/identity-platform/howto-create-self-signed-certificate) or a [Managed Identity](https://learn.microsoft.com/entra/identity/managed-identities-azure-resources/overview). Store secrets in [Key Vault](https://learn.microsoft.com/azure/key-vault/general/overview). |
| Reliability | Put a [Service Bus](https://learn.microsoft.com/azure/service-bus-messaging/) or [Storage Queue](https://learn.microsoft.com/azure/storage/queues/) between `/webhook` and the email pipeline so retries are durable. |
| Throttling | Respect Graph's `Retry-After` header. Pattern: [Avoid throttling on Microsoft Graph](https://learn.microsoft.com/graph/throttling). |
| Multi-tenant | Move from single-tenant to multi-tenant app reg + per-tenant Graph clients. [Convert your single-tenant app](https://learn.microsoft.com/entra/identity-platform/howto-convert-app-to-be-multi-tenant). |
| Observability | [Application Insights](https://learn.microsoft.com/azure/azure-monitor/app/app-insights-overview) for the backend, [SPFx telemetry](https://learn.microsoft.com/sharepoint/dev/spfx/logging) for the panel. |
| CI/CD | GitHub Actions workflow: `npm test` → `npm run build` → `az webapp deploy`. Sample: [pnp/sp-dev-fx-extensions actions](https://github.com/pnp/sp-dev-fx-extensions/tree/main/.github/workflows). |
| Spoofing defense | Validate `clientState` on every `/webhook` notification — reject anything that doesn't match the constant you used at subscribe time. |
| Email templating | Replace the inline HTML with [MJML](https://mjml.io/) or [react-email](https://react.email/) so the layout doesn't break in Outlook. |
| Lifecycle notifications | Subscribe to [subscription lifecycle events](https://learn.microsoft.com/graph/webhooks-lifecycle) so you can react proactively when a subscription is about to be removed. |

---

## Further reading

- [Microsoft Graph training](https://learn.microsoft.com/training/paths/m365-msgraph-associate/)
- [SharePoint Framework training](https://learn.microsoft.com/training/paths/m365-sharepoint-associate/)
- [SharePoint REST API reference](https://learn.microsoft.com/sharepoint/dev/sp-add-ins/get-to-know-the-sharepoint-rest-service)
- [Microsoft 365 Patterns and Practices (PnP)](https://pnp.github.io/) — community-maintained samples and tools, including PnP PowerShell and the PnPjs library.
- [Microsoft Graph PowerShell SDK](https://learn.microsoft.com/powershell/microsoftgraph/overview) — the same Graph calls from PowerShell, ideal for exploration.
- [Graph Explorer](https://developer.microsoft.com/graph/graph-explorer) — try any Graph endpoint in your browser, with auth.
- [Microsoft Q&A — SharePoint Framework tag](https://learn.microsoft.com/answers/tags/235/sharepoint-framework)
- [Stack Overflow — `office365` + `sharepoint-online` + `microsoft-graph`](https://stackoverflow.com/questions/tagged/microsoft-graph)
- [@OfficeDev / @MS365Dev on X](https://x.com/OfficeDev) — release notes and breaking-change announcements.

You've now built the same app that ships in this repo. Star the repo, [open an issue](https://github.com/) (with the actual URL when you fork) if something doesn't reproduce on your tenant, and good luck with your replacement-for-classic-alerts rollout.
