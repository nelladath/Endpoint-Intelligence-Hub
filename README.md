# Endpoint Intelligence Hub

Centralized portal for Microsoft Intune, Windows Autopilot, AI-assisted driver automation,
patch management, application delivery, device compliance, endpoint health, security risk,
hardware intelligence, and operational analytics.

The Python Azure Function collects a full Intune deployment and endpoint intelligence report
via Microsoft Graph. It includes app deployment status, application inventory and type,
Autopilot registrations, compliance/configuration/update-ring status, current-month Windows
Patch Tuesday compliance, enriched device health,
an actionable device risk register, and an auditable fleet health score. Sync uses an
upsert + mark-and-sweep pattern (`sharepoint_sync.py`) so SharePoint always mirrors current
Intune state -- no duplicates, no stale rows -- whether triggered on a schedule or on demand.

The overall score is the weighted average of compliance (25%), encryption (20%), OS currency
(20%), threat-free status (15%), device activity (10%), and management health (10%).

Patch compliance compares each recently checked-in Windows device's full Intune OS build with
the official Microsoft security cumulative-update baseline for the current month and servicing
branch. Before the month's Patch Tuesday it uses the previous month's released baseline. The
dedicated SharePoint Patch Compliance page shows a status chart and device-level evidence,
including required KB/build and the Microsoft source URL. This is a current-state snapshot; it
does not claim historical installation timing.

Every run also uploads a single EndpointIntelligenceHub_Latest.json file to the "JSON Exports"
document library, containing the full report: every domain's records, the health metrics, and
the sync results. Versioning is enabled on that library, so each run's JSON is retained as prior
file versions. Automation (Power Automate, scripts, Power BI) can read this one file instead of
querying every SharePoint list individually.

## Files
- `config.py` -- app settings loader
- `graph_client.py` -- MSAL client-secret auth, paging, retry, `$batch`
- `error_catalog.py` -- known Intune error code -> issue/recommendation mapping
- `intune_collectors.py` -- per-domain Graph data collection
- `intune_analytics.py` -- status normalization, success ratio, health classification
- `pipeline.py` -- reusable collection, analytics, and SharePoint orchestration
- `sharepoint_sync.py` -- upsert + stale-row cleanup sync to SharePoint lists
- `function_app.py` -- Azure Functions v2 entry points (HTTP for on-demand, Timer for schedule)

## Setup
1. Azure AD app registration with a client secret credential (Certificates & secrets -> New
   client secret; copy the value immediately, it's shown only once, and set an expiry reminder
   to rotate it before it lapses). Grant application permissions, admin-consented:
   `DeviceManagementApps.Read.All`, `DeviceManagementConfiguration.Read.All`,
   `DeviceManagementManagedDevices.Read.All`, and `Sites.Selected` (grant write access to the
   target SharePoint site only, via the Sites.Selected admin API).
2. Run `setup_sharepoint_lists.ps1` in PowerShell 7 to create or repair the reporting
   lists: Apps, Application Inventory, Autopilot, Compliance Policies, Config Profiles,
   Patch Compliance, Devices, Device Risks, and Health Summary. Run
   `setup_sharepoint_site.ps1` to publish the modern pages, embedded views, libraries, and
   hierarchical navigation; it also creates the "JSON Exports" document library and prints
   the `SP_LIST_JSON_EXPORTS_ID` value needed in step 3. Both scripts preserve existing
   reporting lists and data.
3. Copy `local.settings.json.example` to `local.settings.json` and fill in real values for
   local testing. In Azure, set the same keys as Function App Settings (use Key Vault
   references for secrets).
4. `pip install -r requirements.txt`, then run/deploy with Azure Functions Core Tools
   (`func start` locally, `func azure functionapp publish <name>` to deploy).

The timer trigger runs daily at 08:00 UTC (`0 0 8 * * *`).

## Triggering from Power Automate
Call the `IntuneReportHttpTrigger` HTTP endpoint (with its function key) from a Power Automate
flow whenever you want a fresh sync -- on a schedule, on a manual button press, or both. The
flow can simply call the Function and branch on the HTTP status/response body to notify on
success or failure; all parsing/upsert/cleanup logic runs inside the Function.
