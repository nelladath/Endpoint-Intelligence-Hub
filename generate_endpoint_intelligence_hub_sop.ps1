[CmdletBinding()]
param(
    [string]$OutputPath = "C:\Temp\Telemetry\Endpoint-Intelligence-Hub-SOP.docx"
)

$ErrorActionPreference = "Stop"
$word = $null
$document = $null

function Add-Paragraph {
    param([string]$Text, [string]$Style = "Normal", [switch]$Bold)
    $selection.Style = $Style
    $selection.Font.Bold = [int]$Bold.IsPresent
    $selection.TypeText($Text)
    $selection.TypeParagraph()
    $selection.Font.Bold = 0
}

function Add-Heading {
    param([string]$Text, [int]$Level = 1)
    Add-Paragraph -Text $Text -Style "Heading $Level"
}

function Add-Bullets {
    param([string[]]$Items)
    foreach ($item in $Items) {
        $selection.Style = "Normal"
        $selection.Range.ListFormat.ApplyBulletDefault()
        $selection.TypeText($item)
        $selection.TypeParagraph()
        $selection.Range.ListFormat.RemoveNumbers()
    }
}

function Add-Steps {
    param([string[]]$Items)
    for ($index = 0; $index -lt $Items.Count; $index++) {
        Add-Paragraph -Text "$($index + 1). $($Items[$index])"
    }
}

function Add-Code {
    param([string]$Text)
    $selection.Style = "Normal"
    $selection.Font.Name = "Consolas"
    $selection.Font.Size = 8
    $selection.ParagraphFormat.LeftIndent = 18
    $selection.Shading.BackgroundPatternColor = 15132390
    $selection.TypeText($Text.Trim())
    $selection.TypeParagraph()
    $selection.Shading.BackgroundPatternColor = 16777215
    $selection.ParagraphFormat.LeftIndent = 0
    $selection.Font.Name = "Aptos"
    $selection.Font.Size = 10.5
}

function Add-Table {
    param([string[]]$Headers, [object[][]]$Rows)
    $range = $selection.Range
    $table = $document.Tables.Add($range, $Rows.Count + 1, $Headers.Count)
    $table.Borders.Enable = 1
    $table.Rows.Item(1).Range.Bold = 1
    $table.Rows.Item(1).Shading.BackgroundPatternColor = 14277081
    for ($column = 0; $column -lt $Headers.Count; $column++) {
        $table.Cell(1, $column + 1).Range.Text = $Headers[$column]
    }
    for ($row = 0; $row -lt $Rows.Count; $row++) {
        for ($column = 0; $column -lt $Headers.Count; $column++) {
            $table.Cell($row + 2, $column + 1).Range.Text = [string]$Rows[$row][$column]
        }
    }
    $table.AutoFitBehavior(1)
    $selection.SetRange($table.Range.End, $table.Range.End)
    $selection.TypeParagraph()
}

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $document = $word.Documents.Add()
    $selection = $word.Selection

    $document.Styles.Item("Normal").Font.Name = "Aptos"
    $document.Styles.Item("Normal").Font.Size = 10.5
    $document.Styles.Item("Title").Font.Name = "Aptos Display"
    $document.Styles.Item("Title").Font.Size = 30
    $document.Styles.Item("Title").Font.Color = 8547344
    $document.Styles.Item("Heading 1").Font.Color = 8547344
    $document.Styles.Item("Heading 2").Font.Color = 12419407

    $selection.Style = "Title"
    $selection.TypeText("Endpoint Intelligence Hub")
    $selection.TypeParagraph()
    $selection.Style = "Subtitle"
    $selection.TypeText("Implementation and Operations Standard Operating Procedure")
    $selection.TypeParagraph()
    $selection.TypeParagraph()
    Add-Paragraph -Text "Centralized portal for Microsoft Intune, Windows Autopilot, AI-assisted driver automation, patch management, application delivery, device compliance, endpoint health, security risk, hardware intelligence, and operational analytics." -Bold
    $selection.TypeParagraph()
    Add-Table -Headers @("Document field", "Value") -Rows @(
        @("Version", "1.0"),
        @("Date", "17 September 2026"),
        @("Classification", "Internal - contains configuration identifiers but no credentials"),
        @("Current environment", "NTT Data Dynamic Workplace Services tenant"),
        @("Purpose", "Operate the current solution and reproduce it safely in another Microsoft 365 tenant")
    )
    $selection.InsertBreak(7)

    Add-Heading -Text "Table of Contents" -Level 1
    $tocRange = $selection.Range
    $document.TablesOfContents.Add($tocRange, $true, 1, 3) | Out-Null
    $selection.SetRange($tocRange.End, $tocRange.End)
    $selection.InsertBreak(7)

    Add-Heading -Text "1. How the Current Setup Works" -Level 1
    Add-Paragraph -Text "Endpoint Intelligence Hub is a serverless reporting pipeline. It authenticates to Microsoft Graph as an Entra application, gathers Intune and Windows Autopilot data, calculates health and risk metrics in Python, mirrors the latest state into nine SharePoint lists, and presents those lists through modern SharePoint pages."
    Add-Heading -Text "1.1 End-to-end data flow" -Level 2
    Add-Steps -Items @(
        "Azure starts the Function on the daily timer or a caller invokes the function-key protected HTTP endpoint.",
        "MSAL obtains an app-only Microsoft Graph token using TENANT_ID, CLIENT_ID, and CLIENT_SECRET.",
        "Collectors read managed devices, applications, Autopilot registrations, compliance policies, configuration profiles, update rings, and per-app deployment reports.",
        "The analytics layer calculates storage health, inactivity, OS currency, certificate state, device risk, status normalization, and the weighted endpoint health score.",
        "The synchronization layer upserts each record into its target SharePoint list using a stable Key and stamps every row with one LastRunId.",
        "Items absent from the current source data are deleted, so SharePoint remains a current-state mirror rather than an accumulating history table.",
        "Modern SharePoint pages embed the list views for operations, health, Autopilot, applications, patching, risk, and AI driver automation."
    )
    Add-Heading -Text "1.2 Architecture" -Level 2
    Add-Code -Text @'
Daily Timer (08:00 UTC) or HTTP POST
                  |
                  v
Azure Function: intune-report-func
  function_app.py -> pipeline.py
                  |
                  v
MSAL app-only token -> Microsoft Graph
  Intune + Autopilot + Intune Reports API
                  |
                  v
Collectors -> Analytics -> Risk/Health metrics
                  |
                  v
SharePoint current-state synchronization
                  |
                  v
Endpoint Intelligence Hub modern pages and embedded views
'@
    Add-Heading -Text "1.3 Trigger behavior" -Level 2
    Add-Bullets -Items @(
        "Scheduled trigger: IntuneReportTimerTrigger.",
        "Schedule: 0 0 8 * * * (NCRONTAB), every day at 08:00 UTC.",
        "runOnStartup is false; useMonitor is true.",
        "No WEBSITE_TIME_ZONE or TZ override is configured, so UTC is authoritative.",
        "On-demand trigger: IntuneReportHttpTrigger, POST /api/intune-report, Function authorization level.",
        "The established trigger names and route are intentionally retained for compatibility even though the product display name is Endpoint Intelligence Hub."
    )

    Add-Heading -Text "2. Where the Code and Services Are Hosted" -Level 1
    Add-Heading -Text "2.1 Source code" -Level 2
    Add-Table -Headers @("Item", "Current value") -Rows @(
        @("Editable source folder", "C:\Temp\Telemetry"),
        @("Source control", "No local .git metadata detected; currently not backed by a Git repository"),
        @("Deployment method", "Azure Functions Core Tools: func azure functionapp publish intune-report-func"),
        @("Python dependencies", "azure-functions, msal, requests"),
        @("Local configuration", "local.settings.json; never commit or distribute this file")
    )
    Add-Paragraph -Text "Operational recommendation: move C:\Temp\Telemetry into a private Azure DevOps or GitHub repository, add local.settings.json and generated package folders to .gitignore, and deploy through a protected CI/CD environment. Azure contains the deployed package, but it should not be treated as the authoritative source repository."

    Add-Heading -Text "2.2 Current Azure deployment" -Level 2
    Add-Table -Headers @("Setting", "Current value") -Rows @(
        @("Subscription", "NTT Data Dynamic Workplace Services"),
        @("Subscription ID", "96c73ca6-1ee3-4076-b5ea-5d9358e302b3"),
        @("Resource group", "WPS_Sandbox"),
        @("Function App", "intune-report-func"),
        @("Region", "East US 2"),
        @("Operating system", "Linux"),
        @("Hosting plan", "ASP-WPSSandbox-9f86"),
        @("SKU", "FC1 - Flex Consumption"),
        @("State", "Running"),
        @("HTTPS only", "Enabled"),
        @("Minimum TLS", "1.2"),
        @("Hostname", "intune-report-func-abdcbya8fzcrc7gq.eastus2-01.azurewebsites.net"),
        @("Application Insights", "intune-report-func, East US 2, 90-day retention")
    )

    Add-Heading -Text "2.3 Current Entra application" -Level 2
    Add-Table -Headers @("Setting", "Current value") -Rows @(
        @("Display name", "intune-report-function"),
        @("Application/client ID", "81b0f329-a882-44d4-9755-7c34938047e1"),
        @("Service principal object ID", "f6508445-7328-448a-9da4-55c4f0a0792f"),
        @("Authentication", "Confidential client using app-only client credentials"),
        @("SharePoint scope", "Sites.Selected plus an explicit site write grant")
    )
    Add-Paragraph -Text "Granted Microsoft Graph application permissions:"
    Add-Bullets -Items @(
        "DeviceManagementApps.Read.All",
        "DeviceManagementConfiguration.Read.All",
        "DeviceManagementManagedDevices.Read.All",
        "DeviceManagementServiceConfig.Read.All",
        "Sites.Selected"
    )

    Add-Heading -Text "2.4 Current SharePoint deployment" -Level 2
    Add-Table -Headers @("Setting", "Current value") -Rows @(
        @("Site title", "Endpoint Intelligence Hub"),
        @("Site URL", "https://nttdsdws.sharepoint.com/sites/IntuneReporting"),
        @("Site ID", "nttdsdws.sharepoint.com,615acc00-a25f-455a-aa0a-930cced2ed52,38b21869-85a4-407d-bc4f-9bc099165f40"),
        @("Home page", "SitePages/Endpoint-Intelligence-Hub.aspx"),
        @("Provisioning authentication", "PnP PowerShell OSLogin using a tenant-approved public client"),
        @("Presentation", "15 modern pages, embedded list views, hierarchical quick launch")
    )

    Add-Heading -Text "3. Code Responsibilities" -Level 1
    Add-Table -Headers @("File", "Responsibility") -Rows @(
        @("function_app.py", "Azure Functions HTTP and daily timer entry points"),
        @("pipeline.py", "Endpoint Intelligence Hub identity and end-to-end orchestration"),
        @("graph_client.py", "MSAL client credentials, Graph paging, retries, writes, and Intune streamed app report"),
        @("intune_collectors.py", "Devices, apps, Autopilot, compliance, configuration, and patch collection"),
        @("intune_analytics.py", "Status normalization, OS currency, device risk, health score, and distributions"),
        @("sharepoint_sync.py", "Keyed upsert, LastRunId stamp, stale-item deletion, and write failure enforcement"),
        @("config.py", "Required Azure/local environment settings"),
        @("error_catalog.py", "Known Intune error explanations and recommendations"),
        @("setup_sharepoint_lists.ps1", "Idempotent nine-list schema provisioning"),
        @("setup_sharepoint_site.ps1", "Site branding, pages, navigation, libraries, views, and embedded web parts"),
        @("verify_sharepoint_site.ps1", "Read-only validation of branding, pages, libraries, web parts, and navigation")
    )

    Add-Heading -Text "4. Data Sets and SharePoint Lists" -Level 1
    Add-Table -Headers @("List", "Purpose") -Rows @(
        @("Intune Devices", "Inventory plus ownership, encryption, storage, threat, inactivity, OS currency, certificate, join, Autopilot, and risk fields"),
        @("Intune Application Inventory", "Application-level inventory and Win32, Store, Microsoft 365, iOS, Android, and macOS classification"),
        @("Intune Apps", "Per-device and per-user application deployment state, errors, and remediation"),
        @("Intune Autopilot", "Registration, serial, group tag, enrollment state, profile assignment, and contact state"),
        @("Intune Compliance Policies", "Per-device compliance-policy outcomes"),
        @("Intune Config Profiles", "Per-device configuration-profile outcomes"),
        @("Intune Patch Compliance", "Windows Update ring deployment status"),
        @("Intune Device Risks", "Devices with risk score 2 or greater and actionable reasons"),
        @("Intune Health Summary", "Overall score, weighted components, fleet indicators, and distributions")
    )
    Add-Paragraph -Text "Each list requires a Key field and LastRunId field. Key is the business identifier used for upsert. LastRunId proves which execution last refreshed a record and supports stale-item removal."

    Add-Heading -Text "5. Analytics Methodology" -Level 1
    Add-Heading -Text "5.1 Overall health score" -Level 2
    Add-Paragraph -Text "The score is a weighted average from 0 to 100:"
    Add-Code -Text "Score = Compliance x 25% + Encryption x 20% + Patch/OS Currency x 20% + Threat-Free x 15% + Device Activity x 10% + Management Health x 10%"
    Add-Bullets -Items @(
        "Excellent: 95 or higher.",
        "Good: 85 to less than 95.",
        "Fair: 70 to less than 85.",
        "Poor: below 70.",
        "Current validated score at document creation: 74.7, Fair. This changes with source telemetry."
    )
    Add-Heading -Text "5.2 Device health rules" -Level 2
    Add-Bullets -Items @(
        "Stale: last sync is missing or older than 30 days.",
        "Critical storage: below 2 percent free; Low storage: below 10 percent free.",
        "OS currency uses Windows build thresholds in intune_analytics.py; review these thresholds when Microsoft servicing policy changes.",
        "Expired management certificates are determined by comparing certificate expiration with collection time.",
        "Threat state is read from partnerReportedThreatState.",
        "Risk scoring: jailbreak/high threat/wipe failed = 3; medium threat or inactivity over 90 days = 2; low threat, unencrypted, wipe pending, or non-compliant = 1.",
        "Risk level: High at 4 or more, Medium at 2-3, Low at 1, None at 0. Only score 2 or higher enters the risk register."
    )

    Add-Heading -Text "6. Prerequisites for a New Tenant" -Level 1
    Add-Bullets -Items @(
        "Microsoft Intune tenant with managed-device and application data.",
        "SharePoint Online site for the hub.",
        "Azure subscription with permission to create a resource group, storage, Function App, and Application Insights.",
        "Entra role able to create app registrations and grant application admin consent, normally Global Administrator or Privileged Role Administrator with appropriate application administration rights.",
        "SharePoint Administrator or site owner for hub provisioning and Sites.Selected grant.",
        "Python 3.11, Azure CLI, Azure Functions Core Tools v4, PowerShell 7.4 or later, and PnP.PowerShell.",
        "A tenant-approved PnP public-client app for interactive SharePoint provisioning. Under managed-device Conditional Access, use OSLogin rather than device-code login."
    )

    Add-Heading -Text "7. New Tenant Implementation Procedure" -Level 1
    Add-Heading -Text "7.1 Establish naming and record identifiers" -Level 2
    Add-Steps -Items @(
        "Choose the Azure subscription, region, resource group, Function App name, storage account name, and SharePoint site URL.",
        "Record the tenant ID and target SharePoint URL in an implementation worksheet.",
        "Use separate development and production app registrations and Function Apps where production change control is required."
    )

    Add-Heading -Text "7.2 Create the Function Entra application" -Level 2
    Add-Steps -Items @(
        "Open Entra admin center > Identity > Applications > App registrations > New registration.",
        "Name it Endpoint Intelligence Hub Automation. Select Accounts in this organizational directory only. No redirect URI is required for the confidential Function app.",
        "Record the Application (client) ID and Directory (tenant) ID.",
        "Under API permissions, add Microsoft Graph Application permissions: DeviceManagementApps.Read.All, DeviceManagementConfiguration.Read.All, DeviceManagementManagedDevices.Read.All, DeviceManagementServiceConfig.Read.All, and Sites.Selected.",
        "Grant tenant-wide admin consent and verify every permission shows Granted.",
        "Prefer a certificate or a Key Vault-managed secret. If using a client secret, create it, capture the value once, set an owner and expiration reminder, and never put it in source control or documentation."
    )

    Add-Heading -Text "7.3 Prepare the PnP provisioning application" -Level 2
    Add-Paragraph -Text "The Function app uses Sites.Selected. Schema and page provisioning therefore runs under a separate delegated administrative identity. Reuse a tenant-approved PnP app or create a public-client app."
    Add-Steps -Items @(
        "Create an Entra app registration such as Endpoint Intelligence Hub PnP Provisioning.",
        "Enable public client flows.",
        "Add public-client redirect URIs http://localhost and ms-appx-web://microsoft.aad.brokerplugin/<PNP-CLIENT-ID>.",
        "Add SharePoint delegated AllSites.FullControl and grant admin consent. Restrict assignment to approved administrators if organizational policy requires it.",
        "Use Connect-PnPOnline -OSLogin on compliant Windows devices. DeviceLogin can fail with Conditional Access error 53003 because it lacks managed-device context."
    )

    Add-Heading -Text "7.4 Create the SharePoint site" -Level 2
    Add-Steps -Items @(
        "Create a modern Team site or Communication site named Endpoint Intelligence Hub.",
        "Use a stable URL such as https://<tenant>.sharepoint.com/sites/EndpointIntelligenceHub. The display name can change later without changing the URL.",
        "Add the implementing administrator as site owner.",
        "Update SiteUrl, TenantId, and ClientId parameters when invoking the provisioning scripts."
    )
    Add-Code -Text @'
pwsh -NoProfile -File .\setup_sharepoint_lists.ps1 `
  -SiteUrl "https://<tenant>.sharepoint.com/sites/EndpointIntelligenceHub" `
  -TenantId "<tenant-id>" `
  -ClientId "<pnp-public-client-id>"

pwsh -NoProfile -File .\setup_sharepoint_site.ps1 `
  -SiteUrl "https://<tenant>.sharepoint.com/sites/EndpointIntelligenceHub" `
  -TenantId "<tenant-id>" `
  -ClientId "<pnp-public-client-id>"

pwsh -NoProfile -File .\verify_sharepoint_site.ps1 `
  -SiteUrl "https://<tenant>.sharepoint.com/sites/EndpointIntelligenceHub" `
  -TenantId "<tenant-id>" `
  -ClientId "<pnp-public-client-id>"
'@
    Add-Paragraph -Text "Copy the nine list IDs printed by setup_sharepoint_lists.ps1. Rerunning the script is supported; it adds missing lists/fields without deleting existing data."

    Add-Heading -Text "7.5 Grant Sites.Selected write access" -Level 2
    Add-Paragraph -Text "Grant the Function app write access only to the reporting site. Do not grant Sites.ReadWrite.All unless a documented exception is approved."
    Add-Code -Text @'
Connect-PnPOnline `
  -Url "https://<tenant>-admin.sharepoint.com" `
  -OSLogin `
  -ClientId "<pnp-public-client-id>" `
  -Tenant "<tenant-id>"

Grant-PnPAzureADAppSitePermission `
  -AppId "<function-app-client-id>" `
  -DisplayName "Endpoint Intelligence Hub Automation" `
  -Permissions Write `
  -Site "https://<tenant>.sharepoint.com/sites/EndpointIntelligenceHub"

Get-PnPAzureADAppSitePermission `
  -Site "https://<tenant>.sharepoint.com/sites/EndpointIntelligenceHub"
'@

    Add-Heading -Text "7.6 Create Azure resources" -Level 2
    Add-Steps -Items @(
        "In Azure portal, create or select a resource group.",
        "Create a Function App using Flex Consumption, Linux, Functions runtime v4, Python 3.11, and the chosen region.",
        "Create or select the required storage account.",
        "Enable Application Insights and choose a retention period aligned with operational and compliance requirements.",
        "Enable HTTPS only and require TLS 1.2 or later.",
        "For the HTTP trigger, retain Function authorization. Do not make it anonymous.",
        "For production, connect the Function App to Key Vault using managed identity and store the Graph credential as a Key Vault reference."
    )

    Add-Heading -Text "7.7 Configure Function App settings" -Level 2
    Add-Table -Headers @("Setting", "Meaning") -Rows @(
        @("TENANT_ID", "Target Entra tenant ID"),
        @("CLIENT_ID", "Endpoint Intelligence Hub confidential app client ID"),
        @("CLIENT_SECRET", "Secret value or Key Vault reference; never document the value"),
        @("SHAREPOINT_SITE_ID", "Graph composite SharePoint site ID"),
        @("SP_LIST_APPS_ID", "Intune Apps list ID"),
        @("SP_LIST_COMPLIANCE_ID", "Intune Compliance Policies list ID"),
        @("SP_LIST_CONFIG_PROFILES_ID", "Intune Config Profiles list ID"),
        @("SP_LIST_PATCH_ID", "Intune Patch Compliance list ID"),
        @("SP_LIST_DEVICES_ID", "Intune Devices list ID"),
        @("SP_LIST_APP_INVENTORY_ID", "Intune Application Inventory list ID"),
        @("SP_LIST_AUTOPILOT_ID", "Intune Autopilot list ID"),
        @("SP_LIST_DEVICE_RISKS_ID", "Intune Device Risks list ID"),
        @("SP_LIST_HEALTH_SUMMARY_ID", "Intune Health Summary list ID")
    )
    Add-Code -Text @'
az functionapp config appsettings set `
  --name <function-app-name> `
  --resource-group <resource-group> `
  --settings `
    TENANT_ID=<tenant-id> `
    CLIENT_ID=<function-app-client-id> `
    CLIENT_SECRET=@Microsoft.KeyVault(SecretUri=<secret-uri>) `
    SHAREPOINT_SITE_ID=<site-id> `
    SP_LIST_APPS_ID=<id> `
    SP_LIST_COMPLIANCE_ID=<id> `
    SP_LIST_CONFIG_PROFILES_ID=<id> `
    SP_LIST_PATCH_ID=<id> `
    SP_LIST_DEVICES_ID=<id> `
    SP_LIST_APP_INVENTORY_ID=<id> `
    SP_LIST_AUTOPILOT_ID=<id> `
    SP_LIST_DEVICE_RISKS_ID=<id> `
    SP_LIST_HEALTH_SUMMARY_ID=<id>
'@

    Add-Heading -Text "7.8 Prepare and deploy code" -Level 2
    Add-Steps -Items @(
        "Copy the complete source folder to the controlled workstation or clone it from the approved private repository.",
        "Create a Python 3.11 virtual environment and install requirements.txt.",
        "Create local.settings.json from local.settings.json.example only for local testing. Keep it out of source control.",
        "Sign in with az login --tenant <tenant-id> and select the correct subscription.",
        "From the source directory, publish using Azure Functions Core Tools.",
        "Confirm the deployment output lists both IntuneReportHttpTrigger and IntuneReportTimerTrigger and reports the host as Running."
    )
    Add-Code -Text @'
py -3.11 -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
az login --tenant <tenant-id>
az account set --subscription <subscription-id>
func azure functionapp publish <function-app-name>
'@

    Add-Heading -Text "7.9 Test the deployed Function" -Level 2
    Add-Code -Text @'
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
'@
    Add-Paragraph -Text "Expected result: reportName is Endpoint Intelligence Hub; endpointHealth contains a score/status; syncResults contains nine successful list summaries; no secrets appear in output."

    Add-Heading -Text "7.10 Verify the timer" -Level 2
    Add-Code -Text @'
az functionapp function show `
  --name <function-app-name> `
  --resource-group <resource-group> `
  --function-name IntuneReportTimerTrigger `
  --query "config.bindings[0].{schedule:schedule,useMonitor:useMonitor,runOnStartup:runOnStartup}" `
  -o json
'@
    Add-Paragraph -Text "Expected: schedule 0 0 8 * * *, useMonitor true, runOnStartup false. Azure uses UTC unless WEBSITE_TIME_ZONE or TZ is explicitly configured."

    Add-Heading -Text "8. Validation and Acceptance Criteria" -Level 1
    Add-Bullets -Items @(
        "Function host state is Running and both triggers are indexed.",
        "HTTP test returns HTTP 200 with reportName Endpoint Intelligence Hub.",
        "All nine syncResults entries are present and no write failure is reported.",
        "Every SharePoint list contains one common LastRunId from the test execution.",
        "Health Summary includes Overall Endpoint Health Score and all six weighted components.",
        "Devices contain encryption, storage, threat, patchStatus, risk, inactivity, join type, and Autopilot fields.",
        "Autopilot and Application Inventory lists contain live records where the tenant has source data.",
        "SharePoint site title and home page are Endpoint Intelligence Hub.",
        "Navigation includes Device Management, Application Management, Policy Management, Patch Management, Endpoint Analytics, Power BI Dashboards, AI Driver Automation, and Documentation.",
        "verify_sharepoint_site.ps1 reports the expected pages and embedded List web parts."
    )

    Add-Heading -Text "9. Daily Operations" -Level 1
    Add-Bullets -Items @(
        "The scheduled run starts daily at 08:00 UTC. Confirm business-local time before communicating the schedule.",
        "Review Application Insights failures and duration after each deployment and after credential changes.",
        "Use LastRunId across all lists to confirm a complete synchronized run.",
        "Treat SharePoint lists as current operational state, not historical audit storage. Export history to a data lake or Log Analytics if trend reporting is required.",
        "Review OS build thresholds and risk rules at least quarterly.",
        "Review client-secret or certificate expiry monthly and rotate before expiration.",
        "Test the on-demand endpoint after permission, schema, Graph API, or SharePoint changes."
    )

    Add-Heading -Text "10. Security and Governance" -Level 1
    Add-Bullets -Items @(
        "Use least privilege: Graph read permissions for Intune and Sites.Selected for one SharePoint site.",
        "Keep schema/page provisioning separate from the Function identity; use delegated PnP administration only during controlled changes.",
        "Prefer certificate authentication or Key Vault references over plain client secrets.",
        "Never place CLIENT_SECRET, function keys, storage connection strings, tokens, or exported list data in documentation or Git.",
        "Restrict Function App management with Azure RBAC and protect production deployments with approvals.",
        "Retain Application Insights according to policy and avoid logging access tokens or credentials.",
        "Review Sites.Selected grants and enterprise-app permissions at least quarterly.",
        "Rotate any credential disclosed in chat, terminal transcripts, tickets, or documents."
    )

    Add-Heading -Text "11. Troubleshooting" -Level 1
    Add-Table -Headers @("Symptom", "Cause and resolution") -Rows @(
        @("400 Resource not found for deviceStatuses", "mobileApps/{id}/deviceStatuses was removed. Use beta retrieveDeviceAppInstallationStatusReport with DeviceInstallStatusByApp."),
        @("504 Gateway Timeout", "Long asynchronous report exports exceeded the HTTP front-end timeout. Use the synchronous streamed Intune report action and keep the timer for routine runs."),
        @("400 fields(select=Key)", "Nested OData select requires fields($select=Key)."),
        @("503 serviceNotAvailable in SharePoint writes", "Retry individual writes through the Graph HTTP retry layer; do not trust a successful outer batch response without checking subrequests."),
        @("Duplicate SharePoint Key", "App status is user-specific. Use appId + deviceId + userPrincipalName."),
        @("403 while creating list columns", "The Function's Sites.Selected write grant is for data sync, not administrative provisioning. Run PnP scripts with an authorized delegated administrator."),
        @("AADSTS90095 admin consent required", "Grant tenant admin consent to the PnP delegated SharePoint permission."),
        @("AADSTS53003 Conditional Access blocked", "Use Connect-PnPOnline -OSLogin from a compliant Entra-joined Windows device, not DeviceLogin."),
        @("PnP page locked", "Update generated pages in place with ClearPage/Save rather than deleting and recreating coauthored files."),
        @("func cannot connect to Azure", "Ensure Azure CLI is installed, on PATH, signed in to the correct tenant, and the subscription is selected."),
        @("Can't determine project language", "Ensure local.settings.json specifies FUNCTIONS_WORKER_RUNTIME=python for local Core Tools usage."),
        @("Timer did not run", "Check function indexing, schedule, timezone overrides, useMonitor, storage connectivity, and Application Insights traces.")
    )

    Add-Heading -Text "12. Change, Backup, and Recovery" -Level 1
    Add-Steps -Items @(
        "Commit source and provisioning scripts to a private repository before production use.",
        "Export Function App settings without secret values and retain list IDs in controlled configuration.",
        "Keep setup_sharepoint_lists.ps1 and setup_sharepoint_site.ps1 as the repeatable infrastructure procedure.",
        "Before major schema changes, export SharePoint lists or confirm an approved restore path.",
        "Deploy to a non-production Function App and site first, run acceptance tests, then promote.",
        "Rollback code by redeploying the previous approved package. Do not delete SharePoint lists during rollback.",
        "After rollback, trigger one run and verify all nine lists share the new LastRunId."
    )

    Add-Heading -Text "13. Handover Checklist" -Level 1
    Add-Bullets -Items @(
        "Source repository URL and owners recorded.",
        "Azure subscription, resource group, Function App, plan, storage, and Application Insights owners recorded.",
        "Entra app owner, credential owner, and expiration date recorded in the approved secret-management system.",
        "Graph admin consent and Sites.Selected site grant verified.",
        "SharePoint site owners and support group assigned.",
        "Nine list IDs configured in Azure.",
        "HTTP and timer tests completed.",
        "Monitoring alert recipients configured.",
        "Quarterly permission, OS-threshold, risk-rule, and credential reviews scheduled."
    )

    Add-Heading -Text "Appendix A. Current Source Files" -Level 1
    Add-Bullets -Items @(
        "config.py",
        "error_catalog.py",
        "function_app.py",
        "graph_client.py",
        "host.json",
        "intune_analytics.py",
        "intune_collectors.py",
        "pipeline.py",
        "requirements.txt",
        "setup_sharepoint_lists.ps1",
        "setup_sharepoint_site.ps1",
        "sharepoint_sync.py",
        "verify_sharepoint_site.ps1"
    )

    Add-Heading -Text "Appendix B. Important Design Boundaries" -Level 1
    Add-Bullets -Items @(
        "This is an operational current-state hub, not a long-term data warehouse.",
        "AI Driver Automation currently provides a governed workspace and device targeting data; an AI model or driver remediation engine must be integrated separately if automated decision-making is required.",
        "Security Baselines and Expedite Updates pages exist as operational workspaces, but dedicated collectors must be added if deeper source-specific records are required.",
        "OS currency classifications are deterministic code rules and must be maintained as Windows servicing versions change.",
        "The HTTP trigger is useful for validation and orchestration, while the timer is the normal unattended execution path."
    )

    foreach ($section in $document.Sections) {
        $footer = $section.Footers.Item(1)
        $footer.Range.Text = "Endpoint Intelligence Hub SOP | Internal"
        $footer.Range.ParagraphFormat.Alignment = 1
        $footer.PageNumbers.Add() | Out-Null
    }

    $document.TablesOfContents.Item(1).Update()
    $document.Fields.Update() | Out-Null
    $document.SaveAs2($OutputPath, 12)
    $pageCount = $document.ComputeStatistics(2)
    $document.Close()
    $word.Quit()
    $document = $null
    $word = $null
    [pscustomobject]@{Path=$OutputPath;Pages=$pageCount;Created=$true} | ConvertTo-Json
} finally {
    if ($document) { $document.Close(0) }
    if ($word) { $word.Quit() }
    [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($selection) | Out-Null
    if ($document) { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($document) | Out-Null }
    if ($word) { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($word) | Out-Null }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}