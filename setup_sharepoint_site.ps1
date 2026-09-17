[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$SiteUrl,

    [Parameter(Mandatory = $true)]
    [string]$TenantId,

    [Parameter(Mandatory = $true)]
    [string]$ClientId
)

$ErrorActionPreference = "Stop"

if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    Install-Module PnP.PowerShell -Scope CurrentUser -Force -AllowClobber
}
Import-Module PnP.PowerShell

Connect-PnPOnline -Url $SiteUrl -OSLogin -ClientId $ClientId -Tenant $TenantId

$hubDescription = "Centralized portal for Microsoft Intune, Windows Autopilot, AI-assisted driver automation, patch management, application delivery, device compliance, endpoint health, security risk, hardware intelligence, and operational analytics."

function Ensure-DocumentLibrary {
    param([string]$Title, [string]$Url)

    $library = Get-PnPList -Identity $Title -ErrorAction SilentlyContinue
    if (-not $library) {
        New-PnPList -Title $Title -Template DocumentLibrary -Url $Url -OnQuickLaunch:$false | Out-Null
        $library = Get-PnPList -Identity $Title
        Write-Host "Created library: $Title"
    }
    Set-PnPList -Identity $Title -EnableVersioning $true -MajorVersions 50 -ListExperience NewExperience | Out-Null
    return "$SiteUrl/$($library.RootFolder.ServerRelativeUrl.TrimStart('/').Split('/', 3)[-1])"
}

function New-OperationsPage {
    param(
        [string]$Name,
        [string]$Title,
        [string]$Intro,
        [array]$Columns
    )

    $pageName = "$Name.aspx"
    $page = Get-PnPPage -Identity $pageName -ErrorAction SilentlyContinue
    if ($page) {
        $page.ClearPage() | Out-Null
        $page.Save() | Out-Null
    } else {
        $page = Add-PnPPage -Name $Name -LayoutType Article -Title $Title -HeaderLayoutType ColorBlock
    }
    Add-PnPPageSection -Page $page -SectionTemplate OneColumn -ZoneEmphasis 1 | Out-Null
    Add-PnPPageTextPart -Page $page -Section 1 -Column 1 -Text "<h2>$Title</h2><p>$Intro</p>" | Out-Null

    if ($Columns.Count -gt 0) {
        $template = if ($Columns.Count -eq 3) { "ThreeColumn" } else { "TwoColumn" }
        Add-PnPPageSection -Page $page -SectionTemplate $template | Out-Null
        for ($index = 0; $index -lt $Columns.Count; $index++) {
            $column = $Columns[$index]
            $links = ($column.Links | ForEach-Object {
                "<li><a href='$($_.Url)'><strong>$($_.Title)</strong></a><br/><span>$($_.Description)</span></li>"
            }) -join ""
            $html = "<h2>$($column.Title)</h2><p>$($column.Description)</p><ul>$links</ul>"
            Add-PnPPageTextPart -Page $page -Section 2 -Column ($index + 1) -Text $html | Out-Null
        }
    }

    Set-PnPPage -Identity $pageName -Title $Title -HeaderLayoutType ColorBlock -CommentsEnabled:$false -Publish | Out-Null
    return "$SiteUrl/SitePages/$pageName"
}

function Get-ListUrl {
    param([string]$Title)
    $list = Get-PnPList -Identity $Title
    return "$SiteUrl/$($list.RootFolder.ServerRelativeUrl.TrimStart('/').Split('/', 3)[-1])/AllItems.aspx"
}

function Set-ReportingView {
    param([string]$ListTitle, [string[]]$Fields)
    Set-PnPView -List $ListTitle -Identity "All Items" -Fields $Fields | Out-Null
    Set-PnPList -Identity $ListTitle -ListExperience NewExperience | Out-Null
}

function Add-ReportingListPart {
    param([string]$PageUrl, [string]$ListTitle)
    $pageName = $PageUrl.Split('/')[-1]
    $page = Get-PnPPage -Identity $pageName
    $sectionNumber = $page.Sections.Count + 1
    Add-PnPPageSection -Page $page -SectionTemplate OneColumn | Out-Null
    $list = Get-PnPList -Identity $ListTitle
    $view = Get-PnPView -List $ListTitle -Identity "All Items"
    Add-PnPPageWebPart -Page $page -DefaultWebPartType List -Section $sectionNumber -Column 1 -WebPartProperties @{
        isDocumentLibrary = $false
        selectedListId = $list.Id.ToString()
        selectedViewId = $view.Id.ToString()
    } | Out-Null
    Set-PnPPage -Identity $pageName -CommentsEnabled:$false -Publish | Out-Null
}

$urls = @{
    Devices = Get-ListUrl "Intune Devices"
    Apps = Get-ListUrl "Intune Apps"
    Compliance = Get-ListUrl "Intune Compliance Policies"
    Configuration = Get-ListUrl "Intune Config Profiles"
    Patches = Get-ListUrl "Intune Patch Compliance"
    ApplicationInventory = Get-ListUrl "Intune Application Inventory"
    Autopilot = Get-ListUrl "Intune Autopilot"
    DeviceRisks = Get-ListUrl "Intune Device Risks"
    HealthSummary = Get-ListUrl "Intune Health Summary"
    PowerBI = Ensure-DocumentLibrary "Power BI Dashboards" "PowerBIDashboards"
    Automation = Ensure-DocumentLibrary "Automation" "Automation"
    Documentation = Ensure-DocumentLibrary "Documentation" "Documentation"
}

Set-ReportingView "Intune Devices" @("deviceName", "userPrincipalName", "operatingSystem", "osVersion", "complianceState", "isEncrypted", "storagePercentFree", "threatState", "patchStatus", "riskLevel", "daysInactive", "lastSyncDateTime")
Set-ReportingView "Intune Apps" @("appName", "publisher", "deviceName", "userPrincipalName", "installState", "normalizedStatus", "errorCode", "lastSyncDateTime")
Set-ReportingView "Intune Compliance Policies" @("policyName", "deviceName", "userPrincipalName", "complianceState", "lastReportedDateTime")
Set-ReportingView "Intune Config Profiles" @("profileName", "profileType", "deviceName", "userPrincipalName", "status", "lastReportedDateTime")
Set-ReportingView "Intune Patch Compliance" @("ringName", "deviceName", "userPrincipalName", "status", "lastReportedDateTime")
Set-ReportingView "Intune Application Inventory" @("appName", "appType", "publisher", "displayVersion", "isAssigned", "publishingState", "lastModifiedDateTime")
Set-ReportingView "Intune Autopilot" @("serialNumber", "manufacturer", "model", "groupTag", "enrollmentState", "profileAssignmentStatus", "userPrincipalName", "lastContactedDateTime")
Set-ReportingView "Intune Device Risks" @("deviceName", "userPrincipalName", "riskLevel", "riskScore", "riskReasons", "daysInactive", "lastSyncDateTime")
Set-ReportingView "Intune Health Summary" @("category", "metricName", "metricValue", "percentage", "status", "description", "LastRunId")

$pageUrls = @{}
$pageUrls.Autopilot = New-OperationsPage "Autopilot" "Windows Autopilot" "Enrollment readiness, deployment profiles, and provisioning outcomes for corporate Windows devices." @(@{ Title = "Autopilot inventory"; Description = "Registration and profile assignment state."; Links = @(@{ Title = "Open Autopilot records"; Url = $urls.Autopilot; Description = "Serial number, enrollment state, group tag, profile assignment, and last contact." }) })
$pageUrls.Hardware = New-OperationsPage "Hardware-Reports" "Hardware Reports" "Hardware lifecycle, model distribution, operating system readiness, storage, encryption, and inventory insights." @(@{ Title = "Device health inventory"; Description = "Current hardware and lifecycle telemetry."; Links = @(@{ Title = "Open hardware records"; Url = $urls.Devices; Description = "Model, OS, storage, encryption, ownership, activity, and risk." }) })
$pageUrls.Win32 = New-OperationsPage "Win32-Apps" "Win32 Apps" "Package ownership, assignment readiness, deployment status, and remediation context for Win32 applications." @(@{ Title = "Application inventory"; Description = "Use App Type to filter Win32 packages."; Links = @(@{ Title = "Open application inventory"; Url = $urls.ApplicationInventory; Description = "Type, publisher, version, assignment, and publishing state." }) })
$pageUrls.Store = New-OperationsPage "Store-Apps" "Microsoft Store Apps" "Store application inventory, assignments, and deployment outcomes." @(@{ Title = "Store inventory"; Description = "Use App Type to filter Microsoft Store packages."; Links = @(@{ Title = "Open application inventory"; Url = $urls.ApplicationInventory; Description = "Type, publisher, version, assignment, and publishing state." }) })
$pageUrls.AppHealth = New-OperationsPage "App-Health" "Application Health" "A focused workspace for installation outcomes, errors, pending deployments, and remediation." @(@{ Title = "Deployment health"; Description = "Per-device and per-user installation state."; Links = @(@{ Title = "Open app deployment health"; Url = $urls.Apps; Description = "Installed, pending, failed, and not-applicable outcomes." }) })
$pageUrls.Baselines = New-OperationsPage "Security-Baselines" "Security Baselines" "Baseline assignment, deployment posture, conflicts, and endpoint security alignment." @()
$pageUrls.Expedite = New-OperationsPage "Expedite-Updates" "Expedite Updates" "Track expedited quality update assignments and deployment progress." @()
$pageUrls.Vulnerability = New-OperationsPage "Vulnerability-Reports" "Vulnerability Reports" "Risk-led endpoint vulnerability reporting and remediation prioritization." @(@{ Title = "Device risk register"; Description = "Devices with multiple actionable risk signals."; Links = @(@{ Title = "Open device risks"; Url = $urls.DeviceRisks; Description = "Risk score, reasons, OS, ownership, and inactivity." }) })
$pageUrls.Analytics = New-OperationsPage "Endpoint-Analytics" "Endpoint Analytics" "Fleet health score, compliance, encryption, OS currency, threat visibility, activity, and management health." @(@{ Title = "Endpoint health metrics"; Description = "Auditable health score and fleet distributions."; Links = @(@{ Title = "Open health summary"; Url = $urls.HealthSummary; Description = "Overall score, weighted components, fleet risks, OS, threat, join, Autopilot, and app metrics." }) })
$pageUrls.DriverAutomation = New-OperationsPage "AI-Driver-Automation" "AI Driver Automation" "AI-assisted driver lifecycle automation for hardware discovery, model targeting, deployment orchestration, exception handling, and operational evidence." @(@{ Title = "Automation workspace"; Description = "Store driver automation packages, runbooks, outputs, and supporting assets."; Links = @(@{ Title = "Open Automation library"; Url = $urls.Automation; Description = "Central repository for AI-assisted driver automation assets and runbooks." }) })

$pageUrls.Device = New-OperationsPage "Device-Management" "Device Management" "A single operational view of managed endpoints, compliance posture, enrollment, and hardware health." @(
    @{ Title = "Inventory and compliance"; Description = "Current managed endpoint posture."; Links = @(
        @{ Title = "Device Inventory"; Url = $urls.Devices; Description = "Managed device ownership and platform details." },
        @{ Title = "Compliance"; Url = $urls.Compliance; Description = "Per-policy device compliance outcomes." }
    ) },
    @{ Title = "Lifecycle"; Description = "Provisioning and hardware planning."; Links = @(
        @{ Title = "Autopilot"; Url = $pageUrls.Autopilot; Description = "Windows enrollment and provisioning." },
        @{ Title = "Hardware Reports"; Url = $pageUrls.Hardware; Description = "Hardware inventory and lifecycle insights." }
    ) }
)

$pageUrls.Application = New-OperationsPage "Application-Management" "Application Management" "Application inventory, deployment outcomes, installation health, and remediation in one workspace." @(
    @{ Title = "Application estate"; Description = "Manage application types and ownership."; Links = @(
        @{ Title = "Win32 Apps"; Url = $pageUrls.Win32; Description = "Win32 package operations." },
        @{ Title = "Store Apps"; Url = $pageUrls.Store; Description = "Microsoft Store application operations." }
    ) },
    @{ Title = "Deployment operations"; Description = "Monitor delivery and installation health."; Links = @(
        @{ Title = "Deployments"; Url = $urls.Apps; Description = "Per-device and per-user deployment status." },
        @{ Title = "App Health"; Url = $pageUrls.AppHealth; Description = "Failures and remediation context." }
    ) }
)

$pageUrls.Policy = New-OperationsPage "Policy-Management" "Policy Management" "Configuration, compliance, and security policy deployment posture." @(
    @{ Title = "Configuration"; Description = "Endpoint configuration delivery."; Links = @(
        @{ Title = "Configuration Profiles"; Url = $urls.Configuration; Description = "Profile deployment status by device." },
        @{ Title = "Compliance Policies"; Url = $urls.Compliance; Description = "Compliance policy outcomes." }
    ) },
    @{ Title = "Security"; Description = "Security configuration posture."; Links = @(
        @{ Title = "Security Baselines"; Url = $pageUrls.Baselines; Description = "Baseline deployment and conflicts." }
    ) }
)

$pageUrls.Patch = New-OperationsPage "Patch-Management" "Patch Management" "Windows servicing posture, expedited deployments, and vulnerability-led remediation." @(
    @{ Title = "Update operations"; Description = "Monitor Windows servicing."; Links = @(
        @{ Title = "Windows Updates"; Url = $urls.Patches; Description = "Update ring status by device." },
        @{ Title = "Expedite Updates"; Url = $pageUrls.Expedite; Description = "Expedited update delivery." }
    ) },
    @{ Title = "Risk"; Description = "Prioritize remediation using endpoint risk."; Links = @(
        @{ Title = "Vulnerability Reports"; Url = $pageUrls.Vulnerability; Description = "Vulnerability and remediation reporting." }
    ) }
)

$pageUrls.Home = New-OperationsPage "Endpoint-Intelligence-Hub" "Endpoint Intelligence Hub" $hubDescription @(
    @{ Title = "Endpoint estate"; Description = "Inventory and endpoint health."; Links = @(
        @{ Title = "Device Management"; Url = $pageUrls.Device; Description = "Inventory, compliance, Autopilot, and hardware." },
        @{ Title = "Endpoint Analytics"; Url = $pageUrls.Analytics; Description = "Health score, security posture, OS currency, and risks." },
        @{ Title = "Overall Health Score"; Url = $urls.HealthSummary; Description = "Current weighted endpoint health metrics." }
    ) },
    @{ Title = "Configuration and delivery"; Description = "Policy and application operations."; Links = @(
        @{ Title = "Application Management"; Url = $pageUrls.Application; Description = "Apps, deployments, and health." },
        @{ Title = "Policy Management"; Url = $pageUrls.Policy; Description = "Profiles, compliance, and baselines." },
        @{ Title = "Patch Management"; Url = $pageUrls.Patch; Description = "Updates and vulnerability reporting." }
    ) },
    @{ Title = "Insights and enablement"; Description = "Dashboards, automation, and knowledge."; Links = @(
        @{ Title = "Power BI Dashboards"; Url = $urls.PowerBI; Description = "Published operational dashboards." },
        @{ Title = "AI Driver Automation"; Url = $pageUrls.DriverAutomation; Description = "Driver discovery, targeting, orchestration, and automation assets." },
        @{ Title = "Documentation"; Url = $urls.Documentation; Description = "Operating procedures and reference material." }
    ) }
)

Add-ReportingListPart $pageUrls.Home "Intune Health Summary"
Add-ReportingListPart $pageUrls.Device "Intune Devices"
Add-ReportingListPart $pageUrls.Autopilot "Intune Autopilot"
Add-ReportingListPart $pageUrls.Hardware "Intune Devices"
Add-ReportingListPart $pageUrls.Application "Intune Application Inventory"
Add-ReportingListPart $pageUrls.Win32 "Intune Application Inventory"
Add-ReportingListPart $pageUrls.Store "Intune Application Inventory"
Add-ReportingListPart $pageUrls.AppHealth "Intune Apps"
Add-ReportingListPart $pageUrls.Patch "Intune Patch Compliance"
Add-ReportingListPart $pageUrls.Vulnerability "Intune Device Risks"
Add-ReportingListPart $pageUrls.Analytics "Intune Health Summary"
Add-ReportingListPart $pageUrls.DriverAutomation "Intune Devices"

Set-PnPHomePage -RootFolderRelativeUrl "SitePages/Endpoint-Intelligence-Hub.aspx"

Get-PnPNavigationNode -Location QuickLaunch | ForEach-Object {
    Remove-PnPNavigationNode -Identity $_ -Force
}

Add-PnPNavigationNode -Location QuickLaunch -Title "Home" -Url $pageUrls.Home | Out-Null

$deviceNode = Add-PnPNavigationNode -Location QuickLaunch -Title "Device Management" -Url $pageUrls.Device
Add-PnPNavigationNode -Location QuickLaunch -Parent $deviceNode -Title "Device Inventory" -Url $urls.Devices | Out-Null
Add-PnPNavigationNode -Location QuickLaunch -Parent $deviceNode -Title "Compliance" -Url $urls.Compliance | Out-Null
Add-PnPNavigationNode -Location QuickLaunch -Parent $deviceNode -Title "Autopilot" -Url $pageUrls.Autopilot | Out-Null
Add-PnPNavigationNode -Location QuickLaunch -Parent $deviceNode -Title "Hardware Reports" -Url $pageUrls.Hardware | Out-Null

$appNode = Add-PnPNavigationNode -Location QuickLaunch -Title "Application Management" -Url $pageUrls.Application
Add-PnPNavigationNode -Location QuickLaunch -Parent $appNode -Title "Win32 Apps" -Url $pageUrls.Win32 | Out-Null
Add-PnPNavigationNode -Location QuickLaunch -Parent $appNode -Title "Store Apps" -Url $pageUrls.Store | Out-Null
Add-PnPNavigationNode -Location QuickLaunch -Parent $appNode -Title "Deployments" -Url $urls.Apps | Out-Null
Add-PnPNavigationNode -Location QuickLaunch -Parent $appNode -Title "App Health" -Url $pageUrls.AppHealth | Out-Null

$policyNode = Add-PnPNavigationNode -Location QuickLaunch -Title "Policy Management" -Url $pageUrls.Policy
Add-PnPNavigationNode -Location QuickLaunch -Parent $policyNode -Title "Configuration Profiles" -Url $urls.Configuration | Out-Null
Add-PnPNavigationNode -Location QuickLaunch -Parent $policyNode -Title "Compliance Policies" -Url $urls.Compliance | Out-Null
Add-PnPNavigationNode -Location QuickLaunch -Parent $policyNode -Title "Security Baselines" -Url $pageUrls.Baselines | Out-Null

$patchNode = Add-PnPNavigationNode -Location QuickLaunch -Title "Patch Management" -Url $pageUrls.Patch
Add-PnPNavigationNode -Location QuickLaunch -Parent $patchNode -Title "Windows Updates" -Url $urls.Patches | Out-Null
Add-PnPNavigationNode -Location QuickLaunch -Parent $patchNode -Title "Expedite Updates" -Url $pageUrls.Expedite | Out-Null
Add-PnPNavigationNode -Location QuickLaunch -Parent $patchNode -Title "Vulnerability Reports" -Url $pageUrls.Vulnerability | Out-Null

Add-PnPNavigationNode -Location QuickLaunch -Title "Endpoint Analytics" -Url $pageUrls.Analytics | Out-Null
Add-PnPNavigationNode -Location QuickLaunch -Title "Power BI Dashboards" -Url $urls.PowerBI | Out-Null
Add-PnPNavigationNode -Location QuickLaunch -Title "AI Driver Automation" -Url $pageUrls.DriverAutomation | Out-Null
Add-PnPNavigationNode -Location QuickLaunch -Title "Documentation" -Url $urls.Documentation | Out-Null

Write-Host "`nEndpoint Intelligence Hub site design completed."
Write-Host "Home: $($pageUrls.Home)"
$web = Get-PnPWeb -Includes WelcomePage
Write-Host "Configured home page: $($web.WelcomePage)"
Write-Host "Quick launch:"
Get-PnPNavigationNode -Location QuickLaunch | ForEach-Object {
    Write-Host "- $($_.Title)"
    Get-PnPProperty -ClientObject $_ -Property Children | Out-Null
    $_.Children | ForEach-Object { Write-Host "  - $($_.Title)" }
}

Set-PnPWeb -Title "Endpoint Intelligence Hub" -Description $hubDescription