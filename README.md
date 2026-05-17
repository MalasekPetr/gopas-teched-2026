# gopas-teched-2026 — SP Webhook Demo

TechEd 2026 demo: replacing SharePoint Online alerts with SPFx + Azure App Service + webhooks. Step-by-step tutorial included.

Conference demo app for the talk *"Jak nahradit oblíbené alerty v SharePoint Online"* (TechEd 2026).

Shows the minimal production-pattern architecture for replacing native SharePoint alerts:
SPFx ListView Command Set → SPO List → Azure App Service (Node.js) → SharePoint Webhooks → Email via Graph API.

> Looking for a production-ready solution? Try [Alerts 365](https://malachis.eu/alerts-365) — free 6-month trial with code `TECHED26`.

## Architecture

```text
┌─────────────────┐    1. user clicks       ┌───────────────────┐
│ SPFx Command    │ ─── "Set Alert" ──────▶ │ SPAlerts_         │
│ Set + Panel     │     writes a row        │ Definitions list  │
└─────────────────┘                         └────────┬──────────┘
                                                     │ 2. cron / manual
                                                     ▼
┌─────────────────┐    4. webhook callback  ┌───────────────────┐
│ SharePoint List │ ────────────────────▶   │ Azure App Service │
│ (watched)       │  (item created/changed) │ /webhook          │
└─────────────────┘                         └────────┬──────────┘
        ▲                                            │ 3. ensure subscriptions
        │ 5. getchanges                              ▼
        └─────────────────────────────────  ┌───────────────────┐
                                            │ Graph sendMail    │
                                            └───────────────────┘
```

## Repo layout

| Path       | What                                                           |
|------------|----------------------------------------------------------------|
| `spfx/`    | SPFx 1.20 ListView Command Set + React panel                   |
| `backend/` | Node.js 22 / Express App Service (webhook + subscription mgmt) |
| `docs/`    | `step-by-step.md` tutorial narrative                           |

## SPAlerts_Definitions — list schema

Create this list manually on the demo site **before** running anything. The list name (`SPAlerts_Definitions`) is hard-coded in both the SPFx panel and the backend.

| Column           | Type   | Required | Notes                                        |
|------------------|--------|----------|----------------------------------------------|
| `Title`          | Text   | yes      | Auto-set to `{listTitle} – {userEmail}`      |
| `WatchedListId`  | Text   | yes      | GUID of the list being watched               |
| `WatchedSiteUrl` | Text   | yes      | Absolute URL of the site                     |
| `UserEmail`      | Text   | yes      | Alert recipient                              |
| `NotifyOn`       | Choice | yes      | `All` / `New items` / `Modified` / `Deleted` |
| `IsActive`       | Yes/No | yes      | Default `Yes`                                |

## Azure AD app registration

One app reg, app-only (client_credentials). Required permissions (admin consent):

| API             | Permission              | Why                                                   |
|-----------------|-------------------------|-------------------------------------------------------|
| Microsoft Graph | `Sites.ReadWrite.All`   | Read definitions, future-proof writes                 |
| Microsoft Graph | `Mail.Send`             | Send alert emails via Graph sendMail                  |
| SharePoint      | `Sites.FullControl.All` | Create/renew webhook subscriptions, call `getchanges` |

Production note: scope `Mail.Send` to the sender mailbox with an Exchange Application Access Policy.

## Backend — env vars

Copy `backend/.env.example` to `backend/.env` and fill in:

| Var                                         | Description                                              |
|---------------------------------------------|----------------------------------------------------------|
| `TENANT_ID` / `CLIENT_ID` / `CLIENT_SECRET` | App reg credentials                                      |
| `SHAREPOINT_TENANT`                         | e.g. `https://contoso.sharepoint.com` (no trailing slash)|
| `DEFINITIONS_SITE_URL`                      | Absolute URL of the site hosting `SPAlerts_Definitions`  |
| `DEFINITIONS_LIST_TITLE`                    | Defaults to `SPAlerts_Definitions`                       |
| `SENDER_EMAIL`                              | Mailbox used as sender for alert emails                  |
| `WEBHOOK_NOTIFICATION_URL`                  | Public HTTPS URL of this App Service + `/webhook`        |
| `PORT`                                      | Local dev port (default 8080)                            |

## Run locally

```powershell
# Backend
cd backend
npm install
npm run dev          # tsx watch — exposes :8080

# In a second terminal: expose to the internet so SharePoint can call /webhook
# (any tunnel works — devtunnel, ngrok, Azure Dev Tunnels)
devtunnel host -p 8080 --allow-anonymous

# SPFx (SPFx 1.22.2 — heft-based, no gulp)
cd ..\spfx
npm install
npm start            # heft start --clean; opens the Workbench / debug URL from serve.json
```

## Setup for the live demo

The full pre-flight checklist (tenant prep, app reg, Azure App Service, mail policy, dry run, troubleshooting) lives in [docs/setup.md](docs/setup.md). What follows here is the short summary.

## Deploy

1. **Backend** — `npm run build`, then ZIP-deploy `dist/` + `package.json` + `node_modules` to Azure App Service (Node 22 runtime). Set the env vars in App Service Configuration.
2. **SPFx** — `npm run build` produces `sharepoint/solution/sp-alert-demo.sppkg`. Upload to the tenant App Catalog and deploy.
3. **First subscription** — hit `GET https://<your-app>.azurewebsites.net/manage-subscriptions` once. Schedule it (Azure Function timer, or any cron) every ~5 months so subscriptions never expire.

## Demo flow

1. Open a SharePoint list → click **Set Alert** in the toolbar.
2. The panel writes a row to `SPAlerts_Definitions`.
3. Call `/manage-subscriptions` — the backend creates a webhook on the watched list.
4. Edit an item in the watched list → SharePoint posts to `/webhook` → backend calls `getchanges` → email goes out via Graph.

## What this demo deliberately skips

Storage queues, multi-tenant, Managed Identity, Teams notifications, Copilot, SPFx property pane, unit tests, CI/CD, rate-limit handling, durable ChangeToken storage. See `CLAUDE.md` for the full non-goals list.
