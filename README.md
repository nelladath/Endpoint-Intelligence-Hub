# Endpoint Intelligence Hub

Endpoint Intelligence Hub is a centralized operational portal for Microsoft Intune, Windows Autopilot, AI-assisted driver automation, patch management, application delivery, device compliance, endpoint health, security risk, hardware intelligence, and reporting analytics.

A Python Azure Function collects current tenant data through Microsoft Graph, calculates health and risk metrics, and synchronizes twelve SharePoint lists. Modern SharePoint pages embed those lists into a navigable operations hub.

## What It Collects

- Managed device inventory, ownership, platform, model, serial number, and join type
- Encryption, storage, inactivity, threat, management certificate, and compliance health
- Windows OS currency and update-ring deployment status
- Application inventory classified as Win32, Microsoft Store, Microsoft 365 Apps, iOS, Android, or macOS
- Per-device and per-user application installation outcomes
- Windows Autopilot registration, enrollment, group tag, profile assignment, and last contact
- Compliance-policy and configuration-profile deployment outcomes
- Complete compliance-policy, configuration-profile, and Windows Update ring definitions, including policies with no deployment-status rows
- Device risk scores and actionable risk reasons
- Fleet-wide health score and operational distributions

## Architecture

```text
Daily timer (08:00 UTC) or function-key HTTP POST
                         |
                         v
Azure Function: function_app.py -> pipeline.py
                         |
                         v
MSAL client-credentials token -> Microsoft Graph
                         |
                         v
Intune + Autopilot + Intune Reports data collectors
                         |
                         v
Health, OS currency, status, and risk analytics
                         |
                         v
Keyed SharePoint upsert + stale-item deletion
                         |
                         v
Endpoint Intelligence Hub modern SharePoint pages
```

The SharePoint lists are a **current-state mirror**, not a historical warehouse. Every row receives a `LastRunId`. Records no longer returned by Microsoft Graph are deleted from SharePoint after each successful collection.

## Repository Contents

| File | Purpose |
|---|---|
| `function_app.py` | Azure Functions HTTP and timer triggers |
| `pipeline.py` | End-to-end collection, analytics, and synchronization orchestration |
| `graph_client.py` | MSAL authentication, Microsoft Graph paging, retries, and report calls |
| `intune_collectors.py` | Intune, application, Autopilot, compliance, configuration, patch, and device collectors |
| `intune_analytics.py` | Health score, OS currency, status normalization, and device risk logic |
| `sharepoint_sync.py` | SharePoint keyed upsert and stale-record deletion |
| `config.py` | Required environment-variable settings |
| `error_catalog.py` | Known application deployment errors and recommendations |
| `setup_sharepoint_lists.ps1` | Idempotent creation/repair of the twelve SharePoint lists |
| `setup_sharepoint_site.ps1` | Site branding, modern pages, navigation, views, libraries, and embedded list web parts |
| `set_sharepoint_field_labels.ps1` | Updates existing SharePoint column display labels to readable Title Case without changing internal names |
| `local.settings.json.example` | Secret-free local configuration template |
| `host.json` | Azure Functions host configuration |
| `requirements.txt` | Python dependencies |

## Health Score

The overall score is a weighted average from 0 to 100:

| Component | Weight |
|---|---:|
| Compliance | 25% |
| Encryption | 20% |
| Patch/OS currency | 20% |
| Threat-free | 15% |
| Device activity | 10% |
| Management health | 10% |

Ratings:

- **Excellent:** 95 or higher
- **Good:** 85 to less than 95
- **Fair:** 70 to less than 85
- **Poor:** below 70

Review the OS-build thresholds and risk rules in `intune_analytics.py` whenever Microsoft servicing policy changes.

## Device Risk Rules

- Jailbroken/rooted, high-severity threat, or failed wipe: 3 points
- Medium-severity threat or inactivity over 90 days: 2 points
- Low-severity threat, unencrypted, wipe pending, or non-compliant: 1 point
- High risk: 4 or more
- Medium risk: 2-3
- Low risk: 1
- Only devices scoring 2 or higher are written to the risk register

## Prerequisites

### Tenant and licensing

- Microsoft Intune tenant with managed endpoints
- SharePoint Online
- Azure subscription
- Permission to register Entra applications and grant admin consent
- SharePoint Administrator or target-site owner

### Administrator workstation

- Python 3.11
- Azure CLI
- Azure Functions Core Tools v4
- PowerShell 7.4 or later
- PnP.PowerShell

Install PnP.PowerShell:

```powershell
Install-Module PnP.PowerShell -Scope CurrentUser -Force -AllowClobber
```

## 1. Create the Function App Registration

In the Microsoft Entra admin center:

1. Go to **Identity > Applications > App registrations > New registration**.
2. Name it `Endpoint Intelligence Hub Automation`.
3. Select **Accounts in this organizational directory only**.
4. No redirect URI is required.
5. Record the **Application (client) ID** and **Directory (tenant) ID**.
6. Open **API permissions > Add a permission > Microsoft Graph > Application permissions**.
7. Add:
   - `DeviceManagementApps.Read.All`
   - `DeviceManagementConfiguration.Read.All`
   - `DeviceManagementManagedDevices.Read.All`
   - `DeviceManagementServiceConfig.Read.All`
   - `Sites.Selected`
8. Select **Grant admin consent**.
9. Create a certificate or client secret.

Prefer certificate authentication or an Azure Key Vault-managed secret. Never commit a secret value.

## 2. Prepare a PnP Provisioning App

SharePoint schema and page provisioning uses a delegated administrator, separate from the Function identity.

Create or reuse a tenant-approved public-client app:

1. Create an Entra app registration such as `Endpoint Intelligence Hub PnP Provisioning`.
2. Enable public client flows.
3. Add public-client redirect URIs:
   - `http://localhost`
   - `ms-appx-web://microsoft.aad.brokerplugin/<PNP-CLIENT-ID>`
4. Add SharePoint delegated `AllSites.FullControl`.
5. Grant admin consent.
6. Optionally require explicit enterprise-app assignment and assign approved administrators.

The provisioning scripts use `Connect-PnPOnline -OSLogin`. This preserves compliant-device context for Conditional Access. Device-code authentication may fail with `AADSTS53003` in tenants requiring a managed device.

## 3. Create the SharePoint Site

Create a modern SharePoint site, for example:

```text
https://contoso.sharepoint.com/sites/EndpointIntelligenceHub
```

The display title becomes **Endpoint Intelligence Hub**. The site URL may use another approved path.

Ensure the administrator running the scripts is a site owner.

## 4. Provision SharePoint Lists

Run PowerShell 7 from the repository folder:

```powershell
pwsh -NoProfile -File .\setup_sharepoint_lists.ps1 `
  -SiteUrl "https://contoso.sharepoint.com/sites/EndpointIntelligenceHub" `
  -TenantId "<tenant-id>" `
  -ClientId "<pnp-public-client-id>"
```

The script creates or repairs these lists:

1. `Intune Apps`
2. `Intune Application Inventory`
3. `Intune Autopilot`
4. `Intune Compliance Policies`
5. `Intune Config Profiles`
6. `Intune Patch Compliance`
7. `Intune Devices`
8. `Intune Device Risks`
9. `Intune Health Summary`
10. `Intune Compliance Policy Inventory`
11. `Intune Configuration Profile Inventory`
12. `Intune Update Ring Inventory`

Copy the twelve list IDs printed at the end. The script is idempotent and preserves existing data.

To update labels on an already provisioned site:

```powershell
pwsh -NoProfile -File .\set_sharepoint_field_labels.ps1 `
  -SiteUrl "https://contoso.sharepoint.com/sites/EndpointIntelligenceHub" `
  -TenantId "<tenant-id>" `
  -ClientId "<pnp-public-client-id>"
```

This changes only the visible column titles. Internal names such as `policyName`, `userPrincipalName`, and `lastModifiedDateTime` remain unchanged so the Python pipeline continues to work.

## 5. Provision the SharePoint Experience

```powershell
pwsh -NoProfile -File .\setup_sharepoint_site.ps1 `
  -SiteUrl "https://contoso.sharepoint.com/sites/EndpointIntelligenceHub" `
  -TenantId "<tenant-id>" `
  -ClientId "<pnp-public-client-id>"
```

This configures:

- Endpoint Intelligence Hub title and description
- Branded home page
- Device Management
- Application Management
- Policy Management
- Patch Management
- Endpoint Analytics
- Windows Autopilot
- Hardware Reports
- App Health
- Vulnerability Reports
- Power BI Dashboards library
- AI Driver Automation workspace and Automation library
- Documentation library
- Embedded SharePoint list views
- Hierarchical quick-launch navigation

### AI Driver Automation scope

The repository creates the workspace, hardware/device targeting data, embedded inventory, and document library for driver automation assets. It does **not** include an AI model or a driver-remediation engine. Integrate an approved AI service, Intune remediation package, or automation workflow separately if automated driver decisions and deployments are required.

## 6. Grant the Function App Access to One Site

`Sites.Selected` requires an explicit site grant. Connect with the PnP provisioning administrator:

```powershell
Connect-PnPOnline `
  -Url "https://contoso-admin.sharepoint.com" `
  -OSLogin `
  -ClientId "<pnp-public-client-id>" `
  -Tenant "<tenant-id>"

Grant-PnPAzureADAppSitePermission `
  -AppId "<function-app-client-id>" `
  -DisplayName "Endpoint Intelligence Hub Automation" `
  -Permissions Write `
  -Site "https://contoso.sharepoint.com/sites/EndpointIntelligenceHub"

Get-PnPAzureADAppSitePermission `
  -Site "https://contoso.sharepoint.com/sites/EndpointIntelligenceHub"
```

Do not replace `Sites.Selected` with `Sites.ReadWrite.All` unless a documented exception is approved.

## 7. Create the Azure Function App

Recommended portal configuration:

1. Create or select a resource group.
2. Create a **Function App**.
3. Choose **Flex Consumption**.
4. Select **Linux**.
5. Select **Python 3.11** and Functions runtime v4.
6. Choose the approved region.
7. Create or select the required storage account.
8. Enable Application Insights.
9. Enable HTTPS only.
10. Require TLS 1.2 or later.
11. Keep FTP disabled where policy allows.

For production, enable a system-assigned managed identity and use Key Vault references for credentials.

## 8. Configure Azure Function Settings

Required settings:

| Setting | Purpose |
|---|---|
| `TENANT_ID` | Target Entra tenant ID |
| `CLIENT_ID` | Confidential Function app client ID |
| `CLIENT_SECRET` | Secret value or Key Vault reference |
| `SHAREPOINT_SITE_ID` | Microsoft Graph composite site ID |
| `SP_LIST_APPS_ID` | Intune Apps list ID |
| `SP_LIST_COMPLIANCE_ID` | Compliance list ID |
| `SP_LIST_CONFIG_PROFILES_ID` | Config profiles list ID |
| `SP_LIST_PATCH_ID` | Patch compliance list ID |
| `SP_LIST_DEVICES_ID` | Devices list ID |
| `SP_LIST_APP_INVENTORY_ID` | Application inventory list ID |
| `SP_LIST_AUTOPILOT_ID` | Autopilot list ID |
| `SP_LIST_DEVICE_RISKS_ID` | Device risks list ID |
| `SP_LIST_HEALTH_SUMMARY_ID` | Health summary list ID |
| `SP_LIST_COMPLIANCE_INVENTORY_ID` | Complete compliance-policy definition list ID |
| `SP_LIST_CONFIGURATION_INVENTORY_ID` | Complete configuration-profile definition list ID |
| `SP_LIST_UPDATE_RING_INVENTORY_ID` | Complete Windows Update ring definition list ID |

Find the site ID:

```powershell
az rest --method GET `
  --url "https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/sites/EndpointIntelligenceHub"
```

Configure settings:

```powershell
az functionapp config appsettings set `
  --name <function-app-name> `
  --resource-group <resource-group> `
  --settings `
    TENANT_ID=<tenant-id> `
    CLIENT_ID=<function-app-client-id> `
    CLIENT_SECRET='@Microsoft.KeyVault(SecretUri=<secret-uri>)' `
    SHAREPOINT_SITE_ID=<site-id> `
    SP_LIST_APPS_ID=<list-id> `
    SP_LIST_COMPLIANCE_ID=<list-id> `
    SP_LIST_CONFIG_PROFILES_ID=<list-id> `
    SP_LIST_PATCH_ID=<list-id> `
    SP_LIST_DEVICES_ID=<list-id> `
    SP_LIST_APP_INVENTORY_ID=<list-id> `
    SP_LIST_AUTOPILOT_ID=<list-id> `
    SP_LIST_DEVICE_RISKS_ID=<list-id> `
    SP_LIST_HEALTH_SUMMARY_ID=<list-id> `
    SP_LIST_COMPLIANCE_INVENTORY_ID=<list-id> `
    SP_LIST_CONFIGURATION_INVENTORY_ID=<list-id> `
    SP_LIST_UPDATE_RING_INVENTORY_ID=<list-id>
```

Do not print secrets in tickets, chat, logs, or screenshots.

## 9. Local Development

```powershell
py -3.11 -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
Copy-Item local.settings.json.example local.settings.json
```

Populate `local.settings.json` locally. It is ignored by Git.

If Azure Storage Emulator/Azurite is unavailable, provide a valid development storage configuration before running Functions locally.

Run:

```powershell
func start
```

## 10. Deploy

```powershell
az login --tenant <tenant-id>
az account set --subscription <subscription-id>
func azure functionapp publish <function-app-name>
```

Successful output must show:

- `IntuneReportHttpTrigger`
- `IntuneReportTimerTrigger`
- Function host state `Running`

The legacy trigger names and `/api/intune-report` route remain for compatibility. The returned product name is `Endpoint Intelligence Hub`.

## 11. Trigger On Demand

```powershell
$key = az functionapp function keys list `
  --name <function-app-name> `
  --resource-group <resource-group> `
  --function-name IntuneReportHttpTrigger `
  --query default -o tsv

$result = Invoke-RestMethod `
  -Method POST `
  -Uri "https://<function-host>/api/intune-report?code=$key" `
  -TimeoutSec 300

$result | ConvertTo-Json -Depth 7
```

Expected response:

- `reportName`: `Endpoint Intelligence Hub`
- `endpointHealth`: current score and rating
- `syncResults`: twelve list results
- One `runId` shared by the completed synchronization

Treat function keys as secrets.

## 12. Daily Schedule

The timer is configured in `function_app.py`:

```python
@app.schedule(schedule="0 0 8 * * *", arg_name="timer", run_on_startup=False, use_monitor=True)
```

It runs **every day at 08:00 UTC** unless `WEBSITE_TIME_ZONE` or `TZ` is configured in the Function App.

Verify:

```powershell
az functionapp function show `
  --name <function-app-name> `
  --resource-group <resource-group> `
  --function-name IntuneReportTimerTrigger `
  --query "config.bindings[0].{schedule:schedule,useMonitor:useMonitor,runOnStartup:runOnStartup}" `
  -o json
```

Expected values:

```json
{
  "runOnStartup": false,
  "schedule": "0 0 8 * * *",
  "useMonitor": true
}
```

## 13. Validation

After deployment:

1. Trigger one on-demand run.
2. Confirm HTTP 200.
3. Confirm all nine entries are present in `syncResults`.
4. Confirm all SharePoint lists contain one common `LastRunId`.
5. Confirm `Intune Health Summary` contains `Overall Endpoint Health Score`.
6. Confirm device rows include encryption, storage, threat, OS currency, inactivity, and risk fields.
7. Confirm Autopilot, Application Inventory, Compliance Policy Inventory, Configuration Profile Inventory, and Update Ring Inventory contain data when the tenant has source records.
8. Open the SharePoint home page and verify navigation and embedded lists.
9. Review Application Insights for exceptions and throttling retries.

## 14. Security Guidance

- Use the five documented Graph application permissions only.
- Use `Sites.Selected` and grant access only to the reporting site.
- Prefer certificate authentication or Key Vault references.
- Never commit `local.settings.json`, `.env`, secrets, tokens, function keys, or generated report data.
- Restrict Function management through Azure RBAC.
- Protect production deployment environments with approvals.
- Review permissions and site grants quarterly.
- Track credential ownership and expiry.
- Rotate any credential exposed in chat, logs, tickets, or screenshots.

## 15. Troubleshooting

### `mobileApps/{id}/deviceStatuses` returns 400

That navigation property was removed. This project uses:

```text
POST /beta/deviceManagement/reports/retrieveDeviceAppInstallationStatusReport
```

with report name `DeviceInstallStatusByApp`.

### HTTP 504 Gateway Timeout

Do not use long asynchronous export jobs for per-app status in the HTTP request path. The project uses the synchronous streamed Intune report action. Use the timer for normal unattended execution.

### SharePoint query returns 400 for `fields(select=Key)`

Nested OData select requires:

```text
fields($select=Key)
```

### SharePoint returns transient 429, 503, or 504

The Graph client retries individual write requests with backoff. Do not treat an outer batch HTTP 200 as proof that every subrequest succeeded.

### Duplicate app deployment keys

App installation state may be user-specific. The business key includes:

```text
appId + deviceId + userPrincipalName
```

### PnP sign-in fails with `AADSTS90095`

Grant tenant admin consent to the PnP app's delegated SharePoint permission.

### PnP sign-in fails with `AADSTS53003`

Conditional Access blocked device-code authentication. Use `Connect-PnPOnline -OSLogin` on a compliant, Entra-joined Windows device.

### Function receives 403 from SharePoint

Verify:

- The Function app has Microsoft Graph `Sites.Selected` application permission with admin consent.
- The service principal has an explicit `Write` site grant.
- The configured `SHAREPOINT_SITE_ID` and list IDs belong to that site.

### Function cannot determine project language

For local Core Tools use, ensure `local.settings.json` contains:

```json
"FUNCTIONS_WORKER_RUNTIME": "python"
```

### Azure Functions Core Tools cannot connect to Azure

Ensure Azure CLI is installed, on `PATH`, signed into the correct tenant, and the intended subscription is selected.

## 16. Operational Notes

- This solution mirrors current state; it does not preserve historical snapshots.
- For trends, export each run to a data lake, database, Log Analytics, or Power BI dataset.
- Review OS build classifications whenever Windows servicing versions change.
- Review health weights and risk rules with security and endpoint-management owners.
- Validate the HTTP endpoint after changing permissions, Graph APIs, list schemas, or credentials.
- Monitor the daily timer in Application Insights.

## License

No license is granted by default. Add an organization-approved license before public redistribution.
