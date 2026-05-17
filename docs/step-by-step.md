# Step-by-step: Replacing SharePoint alerts with SPFx + Azure

Tutorial narrative for the 1-hour TechEd 2026 session. Each section is a slide / live coding beat. Headings are the skeleton — fill in voiceover and screenshots later.

## 0. Why the native alerts are dying

- What classic SharePoint alerts did, and why Microsoft is phasing them out
- The four ingredients we'll rebuild ourselves: trigger, store, transport, delivery

## 1. The architecture in one slide

- SPFx (UI) → SPO list (store) → Azure App Service (logic) → Webhook + Graph (transport + delivery)
- Why one Azure App Service is enough for the demo (and what production would add)

## 2. The SPO list — `SPAlerts_Definitions`

- Columns: `WatchedListId`, `WatchedSiteUrl`, `UserEmail`, `NotifyOn`, `IsActive`
- Why we picked a list over a database (governance, no infra, easy to inspect)

## 3. SPFx ListView Command Set — the "Set Alert" button

- Scaffold: `yo @microsoft/sharepoint` → ListView Command Set, React, TS
- The `onExecute` handler — opening a React panel from a command set
- Reading list context: `pageContext.list.id`, `pageContext.user.email`

## 4. The React panel (FluentUI)

- Why a Panel and not a dialog (full keyboard focus, mobile-friendly)
- The minimal form: list (readonly), notify-on (dropdown), email (textbox)
- Save / Remove flow against the definitions list

## 5. Talking to SharePoint from SPFx — `SPHttpClient`

- Why `SPHttpClient` (auth handled, digest handled)
- POST / DELETE patterns against `/_api/web/lists/getbytitle(...)/items`
- The "find existing definition for this list+user" lookup

## 6. The Azure App Service backend

- Node 22 + Express, TypeScript, `tsx` for dev
- Project structure: `routes/` + `services/`
- App-only auth via `client_credentials` — and why this is OK for a backend job

## 7. The webhook handshake — STEP 1

- SharePoint POSTs `?validationToken=...` and expects the token echoed back as `text/plain` within 5 seconds
- Live demo: ngrok / devtunnel, hit the endpoint with curl, observe the handshake

## 8. Creating the subscription — STEP 2

- `POST /_api/web/lists(id)/subscriptions` with `notificationUrl` + `expirationDateTime`
- The 6-month maximum — and the renewal pattern (idempotent PATCH)
- Why a scheduler is mandatory (Azure Function timer, Logic App, GitHub Action — any cron)

## 9. Reacting to a notification — STEP 3

- The notification payload carries NO item data — just `subscriptionId`, `resource`, `siteUrl`
- The `getchanges` API and the ChangeToken loop
- Production callout: ChangeToken **must** be persisted (SPO list, Azure Table, Cosmos)

## 10. Sending the email — STEP 4

- Graph `sendMail` from a service mailbox
- App-only `Mail.Send` + Exchange Application Access Policy
- Inline HTML body — keep it minimal, no template engine for the demo

## 11. End-to-end demo

- Click "Set Alert" → row appears in `SPAlerts_Definitions`
- Hit `/manage-subscriptions` → subscription created
- Edit an item → email arrives within seconds

## 12. What production would add

- Durable ChangeToken storage
- Retry queue (Storage Queue / Service Bus)
- Managed Identity instead of client secret
- Multi-tenant + per-tenant rate-limit handling
- Observability (App Insights, distributed tracing)

## 13. Q&A and pointers

- Link to Alerts 365 (the productized version)
- Microsoft Learn: SharePoint webhooks, Graph sendMail, SPFx extensions
