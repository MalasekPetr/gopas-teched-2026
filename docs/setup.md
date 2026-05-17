# Live demo setup checklist

Everything you need to do **before** walking on stage. Targets ~45 minutes the first time, ~10 minutes on a re-run.

Order matters — each step depends on the previous one.

## 0. Prerequisites

| Tool / account               | Notes                                                           |
|------------------------------|-----------------------------------------------------------------|
| Microsoft 365 tenant         | Tenant admin role required (app catalog, app reg, mail policy) |
| Azure subscription           | Pay-as-you-go is fine — one App Service B1 for the talk         |
| Node.js 22.14+               | x64 build recommended (heft + SPFx 1.22 tested there)           |
| PowerShell 7                 | For Exchange policy commands                                    |
| Azure CLI 2.60+              | `winget install Microsoft.AzureCLI`                             |
| `az login` done              | Logged into the tenant that hosts the App Service               |

## 1. SharePoint site + lists

1. Create the demo site: SharePoint admin → **Active sites** → **+ Create** → Team site → `/sites/AlertsDemo`.
2. On that site, create the **watched** list — e.g. `Demo` — with a default `Title` and a `Description` text column. This is the list the audience will edit.
3. Create the **definitions** list:
   - Name: `SPAlerts_Definitions` (exact match — hard-coded in code)
   - Columns (all text unless noted):
     - `WatchedListId` — Single line of text
     - `WatchedSiteUrl` — Single line of text
     - `UserEmail` — Single line of text
     - `NotifyOn` — Choice (`All`, `New items`, `Modified`, `Deleted`)
     - `IsActive` — Yes/No, default Yes
4. Note the absolute URL of the site (`https://<tenant>.sharepoint.com/sites/AlertsDemo`) — you'll paste it into `.env`.

## 2. Azure AD app registration

In Entra admin center → **App registrations** → **New registration**:

1. Name: `sp-webhook-demo-teched`. Single tenant. No redirect URI.
2. **Certificates & secrets** → **New client secret** → 6 months. Copy the **value** (you'll never see it again).
3. **API permissions** → add and grant admin consent for:

   | API             | Permission              | Type         |
   |-----------------|-------------------------|--------------|
   | Microsoft Graph | `Sites.ReadWrite.All`   | Application  |
   | Microsoft Graph | `Mail.Send`             | Application  |
   | SharePoint      | `Sites.FullControl.All` | Application  |

4. From the **Overview** page, copy the **Application (client) ID** and **Directory (tenant) ID**.

## 3. Scope Mail.Send to one mailbox

Without this, `Mail.Send` lets the app send as **any** user. For a demo this is fine in a sandbox tenant; in any real environment, scope it down:

```powershell
Connect-ExchangeOnline
New-DistributionGroup -Name "AlertsSenderScope" -Type Security
Add-DistributionGroupMember -Identity "AlertsSenderScope" -Member alerts@<tenant>.onmicrosoft.com
New-ApplicationAccessPolicy `
  -AppId "<client-id-from-step-2>" `
  -PolicyScopeGroupId "AlertsSenderScope" `
  -AccessRight RestrictAccess `
  -Description "TechEd demo - restrict Mail.Send to alerts mailbox"
```

The mailbox `alerts@<tenant>.onmicrosoft.com` will be your `SENDER_EMAIL`. Create it (or use any existing licensed mailbox) before this step.

## 4. Azure App Service

```powershell
$rg     = "rg-spwebhook-teched"
$plan   = "plan-spwebhook-teched"
$app    = "spwebhook-teched-$([Guid]::NewGuid().ToString('N').Substring(0,6))"
$region = "westeurope"

az group create -n $rg -l $region
az appservice plan create -g $rg -n $plan --sku B1 --is-linux
az webapp create -g $rg -p $plan -n $app --runtime "NODE:22-lts"
```

Note the resulting hostname — `https://$app.azurewebsites.net`. The webhook URL will be `https://$app.azurewebsites.net/webhook`.

Set app settings (these become env vars in the container):

```powershell
az webapp config appsettings set -g $rg -n $app --settings `
  TENANT_ID="<tenant-id>" `
  CLIENT_ID="<client-id>" `
  CLIENT_SECRET="<secret-value>" `
  SHAREPOINT_TENANT="https://<tenant>.sharepoint.com" `
  DEFINITIONS_SITE_URL="https://<tenant>.sharepoint.com/sites/AlertsDemo" `
  DEFINITIONS_LIST_TITLE="SPAlerts_Definitions" `
  SENDER_EMAIL="alerts@<tenant>.onmicrosoft.com" `
  WEBHOOK_NOTIFICATION_URL="https://$app.azurewebsites.net/webhook"
```

## 5. Deploy the backend

```powershell
cd backend
npm install
npm run build

# Bundle everything App Service needs at runtime
Compress-Archive -Path package.json,dist,node_modules -DestinationPath backend.zip -Force

az webapp deploy -g $rg -n $app --src-path backend.zip --type zip
```

Smoke test:

```powershell
curl "https://$app.azurewebsites.net/"
# expect: "SP Webhook Demo backend up"
```

## 6. Build and deploy the SPFx package

```powershell
cd ..\spfx
npm install            # heft postinstall pulls both sass-embedded arches
npm run build          # produces sharepoint/solution/sp-alert-demo.sppkg
```

Upload `sp-alert-demo.sppkg` to your **tenant app catalog** (SharePoint admin → **More features** → **Apps** → **App Catalog** → **Apps for SharePoint**). Trust + deploy. Tick "Make this solution available to all sites in the organization" so the command appears on every list.

Wait 5–10 minutes for the extension to propagate, then refresh your demo list — **Set Alert** should appear in the toolbar.

## 7. Bootstrap the first webhook subscription

```powershell
curl "https://$app.azurewebsites.net/manage-subscriptions"
```

Expected response: `{"ensured":0,"results":[]}` initially (no definitions yet). After you save your first alert on stage, run it again and it should return `{"ensured":1,"results":[{...,"action":"created",...}]}`.

> In production you'd schedule this every ~5 months via an Azure Function timer or Logic App so subscriptions never expire (6-month max).

## 8. Pre-demo dry run (do this the night before)

1. Open `https://<tenant>.sharepoint.com/sites/AlertsDemo/Lists/Demo`.
2. Click **Set Alert** → panel opens → save it (your email, `All changes`).
3. Verify a row exists in `SPAlerts_Definitions`.
4. `curl https://$app.azurewebsites.net/manage-subscriptions` → confirm `action:"created"`.
5. Add a new item to `Demo`. Within ~30 seconds, an email lands. **If not, check App Service logs** (`az webapp log tail -g $rg -n $app`).
6. **Tear down for the live run:**
   - Delete the row in `SPAlerts_Definitions` (so you can demo the panel creating it).
   - Leave the webhook subscription in place — it'll keep working for the demo even if the definitions list is empty during setup.

## 9. On stage — happy path

| Step | Action                                                              | Time |
|------|---------------------------------------------------------------------|------|
| 1    | Open `/sites/AlertsDemo/Lists/Demo`                                 | 5s   |
| 2    | Click **Set Alert** → panel opens                                   | 5s   |
| 3    | (Optional) Show the row appearing in `SPAlerts_Definitions`         | 10s  |
| 4    | Run `curl .../manage-subscriptions` in a terminal — show JSON       | 10s  |
| 5    | Edit an item in `Demo`                                              | 10s  |
| 6    | Switch to inbox — email arrives                                     | 30s  |

Total: ~70 seconds of live action.

## Troubleshooting cheatsheet

| Symptom                                              | Likely cause                                             | Fix                                                                                  |
|------------------------------------------------------|----------------------------------------------------------|--------------------------------------------------------------------------------------|
| `Set Alert` button doesn't appear                    | Extension not propagated, or wrong scope                 | Wait 10 min; re-deploy with "tenant-wide" ticked                                     |
| Panel opens but **Save** errors                      | `SPAlerts_Definitions` list missing or wrong column      | Verify list name + column internal names (no spaces; `WatchedListId`, not `Watched Id`) |
| `/manage-subscriptions` returns 500                  | `Sites.FullControl.All` not consented                    | Re-check admin consent in app reg                                                    |
| `/webhook` never fires after item edit               | Subscription URL doesn't match `WEBHOOK_NOTIFICATION_URL`| Delete subscription via Graph Explorer; re-run `/manage-subscriptions`               |
| Email never arrives, no error                        | `Mail.Send` policy blocks the sender                     | Verify `Test-ApplicationAccessPolicy` returns `Granted` for the sender mailbox       |
| heft build fails with `sass-embedded-win32-x64`      | Cross-arch Node setup                                    | `npm run` from `spfx/` — `postinstall` re-pulls both arches with `--force`           |

## Reset between dry-runs

```powershell
# Remove demo alert rows so the panel demos "create" cleanly
# (use PnP PowerShell or the SharePoint UI — list is small)

# To re-test webhook handshake, delete the subscription:
# GET /_api/web/lists(guid'<list>')/subscriptions
# DELETE /_api/web/lists(guid'<list>')/subscriptions('<sub-id>')
```
