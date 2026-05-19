# Live demo setup checklist

Everything you need to do **before** walking on stage. Targets ~45 minutes the first time, ~10 minutes on a re-run.

Order matters — each step depends on the previous one. For a full guided lab that explains *why* each piece exists, see [step-by-step.md](step-by-step.md).

## 0. Prerequisites

| Tool / account | Notes |
| --- | --- |
| Microsoft 365 tenant | Tenant admin role required. A free [Microsoft 365 Developer Program](https://developer.microsoft.com/microsoft-365/dev-program) tenant works perfectly. |
| Azure subscription | Pay-as-you-go is fine — one App Service B1 (~€13/mo). [Free trial](https://azure.microsoft.com/free/) credit covers this for the talk. |
| Node.js 22.14+ | [Download](https://nodejs.org/) — x64 build (SPFx 1.22 heft is tested there). |
| PowerShell 7+ | Bundled on Windows 11; otherwise [install pwsh](https://learn.microsoft.com/powershell/scripting/install/installing-powershell). |
| Azure CLI 2.60+ | `winget install Microsoft.AzureCLI` ([docs](https://learn.microsoft.com/cli/azure/install-azure-cli)). |
| PnP PowerShell 3+ | `Install-Module PnP.PowerShell -Scope CurrentUser` ([docs](https://pnp.github.io/powershell/)). |
| `az login` done | Logged into the tenant that hosts the App Service. |

## 1. SharePoint site + lists

1. SharePoint admin → **Active sites** → **+ Create** → Team site → `/sites/AlertsDemo`.
2. On that site, create the **watched** list — e.g. `Demo` — with the default `Title` column plus a `Description` text column. This is the list the audience will edit.
3. Create the **definitions** list:
    - Name: `SPAlerts_Definitions` (exact match — hard-coded in code).
    - Columns (all *Single line of text* unless noted):
        - `WatchedListId`
        - `WatchedSiteUrl`
        - `UserEmail`
        - `NotifyOn` — Choice (`All`, `New items`, `Modified`, `Deleted`)
        - `IsActive` — Yes/No, default `Yes`
4. Note the absolute URL of the site (`https://<tenant>.sharepoint.com/sites/AlertsDemo`) — you'll paste it into `.env`.

PnP one-liner if you want to skip the GUI:

```powershell
Connect-PnPOnline -Url "https://<tenant>.sharepoint.com/sites/AlertsDemo" -Interactive
New-PnPList -Title "SPAlerts_Definitions" -Template GenericList
Add-PnPField -List SPAlerts_Definitions -DisplayName "WatchedListId"  -InternalName "WatchedListId"  -Type Text
Add-PnPField -List SPAlerts_Definitions -DisplayName "WatchedSiteUrl" -InternalName "WatchedSiteUrl" -Type Text
Add-PnPField -List SPAlerts_Definitions -DisplayName "UserEmail"      -InternalName "UserEmail"      -Type Text
Add-PnPField -List SPAlerts_Definitions -DisplayName "NotifyOn"       -InternalName "NotifyOn"       -Type Choice -Choices "All","New items","Modified","Deleted"
Add-PnPField -List SPAlerts_Definitions -DisplayName "IsActive"       -InternalName "IsActive"       -Type Boolean
```

## 2. Entra app registration

In [Entra admin center](https://entra.microsoft.com/) → **Identity** → **Applications** → **App registrations** → **New registration**:

1. Name: `sp-webhook-demo-teched`. Single tenant. No redirect URI.
2. **Certificates & secrets** → **New client secret** → 6 months. Copy the **value** (you'll never see it again — store it in a password manager now).
3. **API permissions** → add and grant admin consent for:

    | API | Permission | Type |
    | --- | --- | --- |
    | Microsoft Graph | `Sites.ReadWrite.All` | Application |
    | Microsoft Graph | `Mail.Send` | Application |

4. From the **Overview** page, copy the **Application (client) ID** and **Directory (tenant) ID**.

Reference: [Microsoft identity platform — register an app](https://learn.microsoft.com/entra/identity-platform/quickstart-register-app).

## 3. Scope Mail.Send to one mailbox

Without this, `Mail.Send` (Application) lets the app send as **any** user in the tenant. In a sandbox / developer tenant that's fine; in any real environment, scope it down:

```powershell
Connect-ExchangeOnline
New-DistributionGroup -Name "AlertsSenderScope" -Type Security
Add-DistributionGroupMember -Identity "AlertsSenderScope" -Member alerts@<tenant>.onmicrosoft.com
New-ApplicationAccessPolicy `
  -AppId "<client-id-from-step-2>" `
  -PolicyScopeGroupId "AlertsSenderScope" `
  -AccessRight RestrictAccess `
  -Description "TechEd demo - restrict Mail.Send to alerts mailbox"
Test-ApplicationAccessPolicy -AppId "<client-id>" -Identity alerts@<tenant>.onmicrosoft.com
```

`Test-ApplicationAccessPolicy` should return `AccessCheckResult : Granted`.

Reference: [Limit app access to specific mailboxes](https://learn.microsoft.com/graph/auth-limit-mailbox-access).

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

Set app settings (these become env vars in the container). **Each setting in its own `az` call** — bundling many with `\`` line-continuation in PowerShell silently drops most of them:

```powershell
az webapp config appsettings set -g $rg -n $app --settings TENANT_ID=<tenant-id>
az webapp config appsettings set -g $rg -n $app --settings CLIENT_ID=<client-id>
az webapp config appsettings set -g $rg -n $app --settings CLIENT_SECRET="<secret-value>"
az webapp config appsettings set -g $rg -n $app --settings DEFINITIONS_SITE_URL=https://<tenant>.sharepoint.com/sites/AlertsDemo
az webapp config appsettings set -g $rg -n $app --settings DEFINITIONS_LIST_TITLE=SPAlerts_Definitions
az webapp config appsettings set -g $rg -n $app --settings SENDER_EMAIL=alerts@<tenant>.onmicrosoft.com
az webapp config appsettings set -g $rg -n $app --settings WEBHOOK_NOTIFICATION_URL="https://$app.azurewebsites.net/webhook"

# Tell App Service to run the prebuilt package as-is (no Oryx rebuild)
az webapp config appsettings set -g $rg -n $app --settings SCM_DO_BUILD_DURING_DEPLOYMENT=false
az webapp config appsettings set -g $rg -n $app --settings WEBSITE_RUN_FROM_PACKAGE=1
```

## 5. Deploy the backend

```powershell
cd backend
npm install
npm run build
npm prune --production      # strips devDeps from node_modules

Compress-Archive -Path dist,node_modules,package.json,package-lock.json `
                 -DestinationPath backend.zip -Force

az webapp deploy -g $rg -n $app --src-path backend.zip --type zip
```

Smoke tests:

```powershell
# 1. The backend is up
Invoke-RestMethod "https://$app.azurewebsites.net/"
# expect: "SP Webhook Demo backend up"

# 2. Subscription manager can read SPAlerts_Definitions (empty list = ensured:0)
Invoke-RestMethod "https://$app.azurewebsites.net/manage-subscriptions"
# expect: { ensured: 0, results: @() }
```

If `/manage-subscriptions` returns a 500, the error message is included in the response body. Common causes are listed in the [troubleshooting cheatsheet](#troubleshooting-cheatsheet) below.

## 6. Build and deploy the SPFx package

```powershell
cd ..\spfx
npm install            # heft postinstall pulls win32-x64 sass-embedded
npm run build          # produces sharepoint/solution/sp-alert-demo.sppkg
```

**Upload to the tenant app catalog:**

1. Open `https://<tenant>.sharepoint.com/sites/appcatalog`.
2. Go to **Apps for SharePoint** → drag `sp-alert-demo.sppkg` in.
3. Tick **Make this solution available to all sites in the organization** → **Deploy**.

**Add a Tenant Wide Extensions row** (required until the package ships a `ClientSideInstance.xml`):

1. Open `https://<tenant>.sharepoint.com/sites/appcatalog/Lists/TenantWideExtensions/AllItems.aspx`.
2. **+ new item** with these values:

    | Field | Value |
    | --- | --- |
    | Title | `Set Alert (TechEd 2026 demo)` |
    | Component Id | the `id` from [spfx/src/extensions/spAlertDemo/SpAlertDemoCommandSet.manifest.json](../spfx/src/extensions/spAlertDemo/SpAlertDemoCommandSet.manifest.json) |
    | Component Properties | `{}` |
    | List Template | `100` |
    | Location | `ClientSideExtension.ListViewCommandSet` |
    | Sequence | `100` |
    | Disabled | `No` |

3. Wait 5–10 minutes for propagation, hard-refresh the watched list — **Set Alert** appears in the toolbar.

Reference: [Tenant-scoped deployment of SPFx solutions](https://learn.microsoft.com/sharepoint/dev/spfx/tenant-scoped-deployment).

## 7. Bootstrap the first webhook subscription

```powershell
Invoke-RestMethod "https://$app.azurewebsites.net/manage-subscriptions"
```

Before you save any alert: `{ ensured: 0, results: @() }`.
After you save your first alert on stage, run it again:

```json
{
  "ensured": 1,
  "results": [
    {
      "resource": "sites/<host>,<siteGuid>,<webGuid>/lists/<listGuid>",
      "action": "created",
      "id": "<graph-subscription-id>"
    }
  ]
}
```

> Schedule this every ~28 days via Azure Function timer / Logic App / GitHub Actions cron so subscriptions never expire — Graph caps SharePoint list subscriptions at 42,300 minutes (~29.4 days). See [Subscription resource lifetime](https://learn.microsoft.com/graph/api/resources/subscription#properties).

## 8. Pre-demo dry run (do this the night before)

1. Open `https://<tenant>.sharepoint.com/sites/AlertsDemo/Lists/Demo`.
2. Click **Set Alert** → panel opens → save it (your email, `All changes`).
3. Verify a row exists in `SPAlerts_Definitions`.
4. `Invoke-RestMethod https://$app.azurewebsites.net/manage-subscriptions` → confirm `action: "created"`.
5. Edit an item in `Demo`. Within ~30 seconds, an email lands. **If not, check App Service logs** (`az webapp log tail -g $rg -n $app`).
6. **Tear down for the live run:**
    - Delete the row in `SPAlerts_Definitions` (so you can demo the panel creating it).
    - Leave the webhook subscription in place — it'll keep working for the demo even if the definitions list is briefly empty during setup.

## 9. On stage — happy path

| Step | Action | Time |
| --- | --- | --- |
| 1 | Open `/sites/AlertsDemo/Lists/Demo` | 5s |
| 2 | Click **Set Alert** → panel opens | 5s |
| 3 | *(Optional)* Show the row appearing in `SPAlerts_Definitions` | 10s |
| 4 | Run `Invoke-RestMethod .../manage-subscriptions` in a terminal — show JSON | 10s |
| 5 | Edit an item in `Demo` | 10s |
| 6 | Switch to inbox — email arrives | 30s |

Total: ~70 seconds of live action.

## Troubleshooting cheatsheet

| Symptom | Likely cause | Fix |
| --- | --- | --- |
| `Set Alert` button doesn't appear | Tenant Wide Extensions row missing, or extension not propagated yet | Re-check step 6; allow 5–10 min and hard-refresh the page |
| Panel opens but **Save** errors | `SPAlerts_Definitions` list missing or wrong column name | Verify list name + internal names (no spaces; `WatchedListId`, not `Watched Id`) |
| `/manage-subscriptions` returns `Invalid URL` | `DEFINITIONS_SITE_URL` env var unset on the App Service | `az webapp config appsettings list ... -o table` — set it, then `az webapp restart` |
| `/manage-subscriptions` returns 500 with `"Unsupported app only token"` | Code is still calling SP REST `/_api/...` with an Entra app-only token | Pull from `main` — the SP REST path was migrated to Graph |
| `/manage-subscriptions` returns 500 with `400 Subscription validation request failed` | Backend not reachable from Microsoft Graph during sub creation | Verify the public URL works (`Invoke-RestMethod /webhook?validationToken=test` should return `test`) |
| `/webhook` never fires after item edit | Subscription expired (>~29 days), or notificationUrl doesn't match | Re-run `/manage-subscriptions`; it will renew or recreate |
| Email never arrives, no error in logs | `Mail.Send` policy blocks the sender | `Test-ApplicationAccessPolicy -AppId <id> -Identity <sender>` — must say `Granted` |
| First edit after a fresh deploy doesn't email; second one does | deltaLink hasn't been primed yet | Call `/manage-subscriptions` once after each container restart — it primes the deltaLink |
| heft build fails with `sass-embedded-win32-x64` | Cross-arch Node setup | `npm install` from `spfx/` — postinstall re-pulls win32-x64 with `--force` |

## Reset between dry-runs

```powershell
# Drop demo alert rows so the panel demos "create" cleanly
Connect-PnPOnline -Url "https://<tenant>.sharepoint.com/sites/AlertsDemo" -Interactive
Get-PnPListItem -List SPAlerts_Definitions | Remove-PnPListItem -Force

# Inspect / delete a Graph subscription (run in Graph Explorer signed in as tenant admin)
# GET    https://graph.microsoft.com/v1.0/subscriptions
# DELETE https://graph.microsoft.com/v1.0/subscriptions/<id>
```
