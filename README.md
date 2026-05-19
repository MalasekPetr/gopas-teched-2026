# gopas-teched-2026 — SP Webhook Demo

Conference demo app for the talk *"Jak nahradit oblíbené alerty v SharePoint Online"* (TechEd 2026). Shows a minimal pro-code replacement for the deprecated SharePoint Online alerts feature.

> Looking for a production-ready solution instead of rolling your own? Try [Alerts 365](https://malachis.eu/alerts-365) — free 6-month trial with code `TECHED26`.

**Want to build this from scratch as a lab?** Follow the full step-by-step tutorial in [docs/step-by-step.md](docs/step-by-step.md).
**Want to just stand the demo up once?** Use the checklist in [docs/setup.md](docs/setup.md).

## What this is

The classic *"Alert me when this list changes"* feature in SharePoint is being phased out. This repo rebuilds the same user-visible behavior with a modern, supported architecture:

| Layer | Tech |
| --- | --- |
| Trigger UI | SPFx 1.22 ListView Command Set + Fluent UI React panel ([source](spfx/src/extensions/spAlertDemo/)) |
| Configuration store | A SharePoint list called `SPAlerts_Definitions` |
| Logic | Node.js 22 + Express on Azure App Service ([source](backend/src/)) |
| Change detection | Microsoft Graph webhook subscriptions on SharePoint lists |
| Delivery | Microsoft Graph `sendMail` from a service mailbox |

## Architecture

```text
┌────────────────────┐  1. user clicks "Set Alert"
│ SPFx Command Set   │ ─────────────────────────────┐
│ + React panel      │                              │
│ (runs as the user) │                              ▼
└────────────────────┘                  ┌────────────────────────┐
                                        │ SPAlerts_Definitions   │
                                        │ list (SP)              │
                                        └───────────┬────────────┘
                                                    │ 2. /manage-subscriptions
                                                    │    reads active definitions
                                                    ▼
┌────────────────────┐  3. POST /v1.0/subscriptions  ┌────────────────────────┐
│ Microsoft Graph    │ ◀─────────────────────────────│ Azure App Service      │
│                    │                               │ (Node 22 / Express)    │
│                    │  4. validationToken handshake │                        │
│                    │ ────────────────────────────▶ │ /webhook               │
└──────────┬─────────┘                               └───────────┬────────────┘
           │                                                     │
           │ 5. notification (POST /webhook) when                │
           │    items in the watched list change                 │
           ▼                                                     │
┌────────────────────┐                                           │
│ Watched SP list    │   6. GET /sites/X/lists/Y/items/delta     │
│ (user edits items) │ ◀─────────────────────────────────────────┤
└────────────────────┘                                           │
                                                                 │ 7. POST /users/{sender}/sendMail
                                                                 ▼
                                                          📧 Email to recipient
```

Key design choices:

- **One Entra app registration**, app-only auth (OAuth 2 `client_credentials`). The backend never acts as a user.
- **Microsoft Graph** for both the webhook subscription AND the change-detection (`/items/delta`). The older SharePoint REST `/_api/web/.../subscriptions` and `getchanges` endpoints don't accept Entra app-only tokens — they require legacy ACS principals, which Microsoft is retiring. Graph is the path forward.
- **The SPFx panel runs as the signed-in user** (delegated context via `SPHttpClient`), so it doesn't need any backend permission to write a row.

## Repo layout

| Path | What |
| --- | --- |
| [spfx/](spfx/) | SPFx 1.22 ListView Command Set + React panel (`Set Alert` button) |
| [backend/](backend/) | Node 22 / Express App Service — webhook receiver + subscription manager |
| [docs/setup.md](docs/setup.md) | Live-demo setup checklist (the one I run before each conference) |
| [docs/step-by-step.md](docs/step-by-step.md) | **Full beginner lab — build this from scratch** |

## SPAlerts_Definitions — list schema

Create this list **before** running anything. The list name `SPAlerts_Definitions` is hard-coded in both the SPFx panel and the backend.

| Column | Type | Required | Notes |
| --- | --- | --- | --- |
| `Title` | Single line of text | yes | Auto-set to `{listTitle} – {userEmail}` by the panel |
| `WatchedListId` | Single line of text | yes | GUID of the list being watched |
| `WatchedSiteUrl` | Single line of text | yes | Absolute URL of the site that holds the watched list |
| `UserEmail` | Single line of text | yes | Alert recipient |
| `NotifyOn` | Choice | yes | `All` / `New items` / `Modified` / `Deleted` |
| `IsActive` | Yes/No | yes | Default `Yes` |

## Entra app registration

One single-tenant app, app-only. Admin-consented Application permissions:

| API | Permission | Why |
| --- | --- | --- |
| [Microsoft Graph](https://learn.microsoft.com/graph/permissions-reference) | `Sites.ReadWrite.All` | Read `SPAlerts_Definitions`, manage webhook subscriptions, call `/items/delta` |
| [Microsoft Graph](https://learn.microsoft.com/graph/permissions-reference) | `Mail.Send` | Send the alert email |

Production tip: scope `Mail.Send` to a single mailbox with an [Exchange Application Access Policy](https://learn.microsoft.com/graph/auth-limit-mailbox-access). Otherwise `Mail.Send` (Application) lets the app send as any user.

## Backend — env vars

Copy [backend/.env.example](backend/.env.example) to `backend/.env` for local dev, or set them in **App Service → Configuration** for cloud.

| Var | Description |
| --- | --- |
| `TENANT_ID` / `CLIENT_ID` / `CLIENT_SECRET` | App reg credentials |
| `DEFINITIONS_SITE_URL` | Absolute URL of the site hosting `SPAlerts_Definitions` |
| `DEFINITIONS_LIST_TITLE` | Defaults to `SPAlerts_Definitions` |
| `SENDER_EMAIL` | Mailbox used as sender for alert emails |
| `WEBHOOK_NOTIFICATION_URL` | Public HTTPS URL of the App Service + `/webhook` |
| `PORT` | Local dev port (default 8080) |

## Run locally

```powershell
# Backend
cd backend
npm install
npm run dev          # tsx watch — exposes :8080

# In a second terminal: expose to the internet so Microsoft Graph can call /webhook
# Pick any tunnel: Microsoft Dev Tunnels, ngrok, Cloudflare Tunnel
devtunnel host -p 8080 --allow-anonymous

# SPFx (heft-based, no gulp)
cd ..\spfx
npm install
npm start            # opens the Workbench / debug URL
```

## Deploy

1. **Backend** — `npm run build` then ZIP-deploy `dist/`, `node_modules/`, `package.json`, `package-lock.json` to Azure App Service (Node 22 runtime). Set the env vars in **App Service → Configuration**.
2. **SPFx** — `npm run build` produces [spfx/sharepoint/solution/sp-alert-demo.sppkg](spfx/sharepoint/solution/). Upload to the tenant **App Catalog** and deploy tenant-wide. *(Until the package ships a `ClientSideInstance.xml`, you also have to add a row to the **Tenant Wide Extensions** list — see [docs/setup.md](docs/setup.md) step 6.)*
3. **First subscription** — call `GET https://<your-app>.azurewebsites.net/manage-subscriptions` once. Schedule it on a timer (Azure Function, Logic App, GitHub Actions cron) at least every **28 days** so subscriptions never expire — Graph caps SharePoint list subscriptions at ~29.4 days.

## Demo flow

1. Open a SharePoint list → click **Set Alert** in the toolbar.
2. The panel writes a row to `SPAlerts_Definitions`.
3. Call `/manage-subscriptions` — the backend creates (or renews) a Graph webhook on the watched list.
4. Edit an item in the watched list → Graph posts to `/webhook` → backend calls `/items/delta` → email goes out via Graph `sendMail`.

## What this demo deliberately skips

Durable deltaLink storage, retry queues, Managed Identity, multi-tenant, throttling / 429 retry, Application Insights, SPFx property pane for site-pinning, unit tests, CI/CD, Teams notifications, Copilot, content type binding for the definitions list. See [docs/step-by-step.md](docs/step-by-step.md) "Section 14 — what production would add" for the full hardening list.
