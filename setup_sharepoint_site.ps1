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

function Add-HubDashboardPart {
    param([string]$PageUrl)
    $pageName = $PageUrl.Split('/')[-1]
    $page = Get-PnPPage -Identity $pageName
    $summaryItems = @(Get-PnPListItem -List "Intune Health Summary" -PageSize 1000 -Fields "metricName", "metricValue", "percentage", "status")
    $metrics = @{}
    foreach ($item in $summaryItems) {
        $fields = $item.FieldValues
        $metrics[[string]$fields.metricName] = $fields
    }
    function Get-Metric([string]$Name, [string]$Fallback = "0") {
        if ($metrics.ContainsKey($Name)) { return [string]$metrics[$Name].metricValue }
        return $Fallback
    }
    function Get-Percent([string]$Name) {
        if ($metrics.ContainsKey($Name)) { return [double]$metrics[$Name].percentage }
        return 0
    }
    function Get-Description([string]$Name) {
        if ($metrics.ContainsKey($Name)) { return [string]$metrics[$Name].description }
        return ""
    }
    function Get-Bar([string]$Name, [string]$Label) {
        $value = Get-Percent $Name
        $width = [math]::Max(0, [math]::Min(100, $value))
        $color = if ($value -ge 95) { "#22c55e" } elseif ($value -ge 80) { "#eab308" } else { "#ef4444" }
        return "<tr><td style='width:190px;padding:9px 12px;color:#cbd5e1;font-size:14px;font-weight:600'>$Label</td><td style='padding:9px 12px'><div style='background:#1e2a37;height:16px;border-radius:6px'><div style='background:$color;height:16px;border-radius:6px;width:$width%'></div></div><div style='color:#94a3b8;font-size:11px;margin-top:4px'>$(Get-Description $Name)</div></td><td style='width:70px;padding:9px 12px;text-align:right;color:#f8fafc;font-size:14px;font-weight:700'>$([math]::Round($value,1))%</td></tr>"
    }
    $score = Get-Metric "Overall Endpoint Health Score"
    $scoreStatus = if ($metrics.ContainsKey("Overall Endpoint Health Score")) { $metrics["Overall Endpoint Health Score"].status } else { "Unknown" }
    $risks = @(Get-PnPListItem -List "Intune Device Risks" -PageSize 1000 -Fields "riskLevel", "riskScore")
    $highRisk = @($risks | Where-Object { $_.FieldValues.riskLevel -eq "High" }).Count
    $mediumRisk = @($risks | Where-Object { $_.FieldValues.riskLevel -eq "Medium" }).Count
    $html = @"
<div style='font-family:Segoe UI,Arial;background:#0f1720;color:#f8fafc;padding:24px;border:1px solid #2a3742;border-radius:10px'>
<div style='font-size:13px;text-transform:uppercase;letter-spacing:.06em;color:#cbd5e1;margin-bottom:12px'>Endpoint Intelligence Hub | Current scope: $(Get-Metric 'Total Devices') managed devices</div>
<table style='width:100%;border-collapse:separate;border-spacing:10px;margin:-10px'><tr>
<td style='width:25%;vertical-align:top;background:#17212b;border:1px solid #334155;padding:18px;border-radius:10px'><div style='font-size:12px;text-transform:uppercase;color:#cbd5e1'>Overall Endpoint Health</div><div style='font-size:42px;font-weight:700;color:#ffffff;margin:6px 0'>$score<span style='font-size:18px;color:#cbd5e1'>/100</span></div><div style='display:inline-block;background:#166534;color:#dcfce7;padding:5px 14px;border-radius:16px;font-weight:700'>$scoreStatus</div></td>
<td style='width:25%;vertical-align:top;background:#17212b;border:1px solid #334155;padding:18px;border-radius:10px'><div style='font-size:12px;text-transform:uppercase;color:#cbd5e1'>Risk-flagged devices</div><div style='font-size:30px;font-weight:700;color:#fbbf24;margin:6px 0'>$($highRisk + $mediumRisk)</div><div style='color:#e2e8f0;line-height:1.5'>High (score 4+): $highRisk<br>Medium (score 2-3): $mediumRisk</div></td>
<td style='width:25%;vertical-align:top;background:#17212b;border:1px solid #334155;padding:18px;border-radius:10px'><div style='font-size:12px;text-transform:uppercase;color:#cbd5e1'>Critical storage</div><div style='font-size:30px;font-weight:700;color:#fb923c;margin:6px 0'>$(Get-Metric 'Critical Storage (<2%)')</div><div style='color:#e2e8f0'>Devices below 2% free space</div></td>
<td style='width:25%;vertical-align:top;background:#17212b;border:1px solid #334155;padding:18px;border-radius:10px'><div style='font-size:12px;text-transform:uppercase;color:#cbd5e1'>Autopilot registered</div><div style='font-size:30px;font-weight:700;color:#60a5fa;margin:6px 0'>$(Get-Metric 'Autopilot Registered')</div><div style='color:#e2e8f0'>Current registrations</div></td>
</tr></table>
<table style='width:100%;border-collapse:collapse;margin-top:22px'><tr><td style='width:50%;vertical-align:top;padding-right:18px'><h3 style='color:#ffffff;margin:0 0 8px'>Weighted health components</h3><table style='width:100%;border-collapse:collapse'>$(Get-Bar 'Compliance (25%)' 'Compliance (25%)')$(Get-Bar 'Encryption (20%)' 'Encryption (20%)')$(Get-Bar 'Patch/OS currency (20%)' 'Patch/OS currency (20%)')</table></td><td style='width:50%;vertical-align:top;padding-left:18px'><h3 style='color:#ffffff;margin:0 0 8px'>Coverage and risk signals</h3><table style='width:100%;border-collapse:collapse'>$(Get-Bar 'Threat-free (15%)' 'Threat-free (15%)')$(Get-Bar 'Device activity (10%)' 'Device activity (10%)')$(Get-Bar 'Management health (10%)' 'Management health (10%)')</table></td></tr></table>
<div style='margin-top:22px;padding:16px;background:#17212b;border:1px solid #334155;border-radius:8px;color:#cbd5e1;line-height:1.6'><strong style='color:#ffffff'>What these numbers mean:</strong> The score is weighted across the current managed-device scope. High and Medium Risk are counts of devices, not users. High means a device risk score of 4 or higher; Medium means 2-3. Risk points combine threat severity, jailbreak/root status, encryption, compliance, wipe state, and inactivity. Unknown threat telemetry is not counted as threat-free.</div>
</div>
"@
    $sectionNumber = $page.Sections.Count + 1
    Add-PnPPageSection -Page $page -SectionTemplate OneColumn | Out-Null
    Add-PnPPageTextPart -Page $page -Section $sectionNumber -Column 1 -Text $html | Out-Null
    Set-PnPPage -Identity $pageName -CommentsEnabled:$false -Publish | Out-Null
}

$urls = @{
    Devices = Get-ListUrl "Intune Devices"
    Apps = Get-ListUrl "Intune Apps"
    Compliance = Get-ListUrl "Intune Compliance Policies"
    Configuration = Get-ListUrl "Intune Config Profiles"
    Patches = Get-ListUrl "Intune Patch Compliance"
    ApplicationInventory = Get-ListUrl "Intune Application Inventory"
    ComplianceInventory = Get-ListUrl "Intune Compliance Policy Inventory"
    ConfigurationInventory = Get-ListUrl "Intune Configuration Profile Inventory"
    UpdateRingInventory = Get-ListUrl "Intune Update Ring Inventory"
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
        @{ Title = "Configuration Profile Inventory"; Url = $urls.ConfigurationInventory; Description = "Every configuration profile definition, including profiles with no status rows." },
        @{ Title = "Compliance Policies"; Url = $urls.Compliance; Description = "Compliance policy outcomes." },
        @{ Title = "Compliance Policy Inventory"; Url = $urls.ComplianceInventory; Description = "Every compliance policy definition." }
    ) },
    @{ Title = "Security"; Description = "Security configuration posture."; Links = @(
        @{ Title = "Security Baselines"; Url = $pageUrls.Baselines; Description = "Baseline deployment and conflicts." },
        @{ Title = "Update Ring Inventory"; Url = $urls.UpdateRingInventory; Description = "Every Windows Update ring definition." }
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
Add-HubDashboardPart $pageUrls.Home
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
Add-ReportingListPart $pageUrls.Policy "Intune Configuration Profile Inventory"

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