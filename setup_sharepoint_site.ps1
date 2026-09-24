[CmdletBinding()]
param(
    [string]$SiteUrl = "https://nttdsdws.sharepoint.com/sites/IntuneReporting",
    [string]$TenantId = "e815be3c-0e2d-4721-9249-1f5dee892931",
    [string]$ClientId = "a71fa672-4cdf-4e00-bcf8-b634d40bac9e"
)

$ErrorActionPreference = "Stop"

if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    Install-Module PnP.PowerShell -Scope CurrentUser -Force -AllowClobber
}
Import-Module PnP.PowerShell

Write-Host "Connecting to SharePoint with the cached OS account..."
Connect-PnPOnline -Url $SiteUrl -OSLogin -ClientId $ClientId -Tenant $TenantId
Write-Host "Connected. Publishing site pages and navigation..."

Write-Host "Removing superseded pages that duplicate other reporting views..."
Remove-PnPPage -Identity "Endpoint-Analytics.aspx" -Force -ErrorAction SilentlyContinue

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
        $page = Add-PnPPage -Name $Name -LayoutType Article -Title $Title
    }
    $page.RemovePageHeader()
    $page.Save() | Out-Null
    $introSectionNumber = 0
    if ($Intro) {
        Add-PnPPageSection -Page $page -SectionTemplate OneColumn -ZoneEmphasis 1 | Out-Null
        $introSectionNumber = 1
        Add-PnPPageTextPart -Page $page -Section 1 -Column 1 -Text "<h2>$Title</h2><p>$Intro</p>" | Out-Null
    }

    if ($Columns.Count -gt 0) {
        $template = if ($Columns.Count -eq 3) { "ThreeColumn" } else { "TwoColumn" }
        Add-PnPPageSection -Page $page -SectionTemplate $template | Out-Null
        $columnsSectionNumber = $introSectionNumber + 1
        for ($index = 0; $index -lt $Columns.Count; $index++) {
            $column = $Columns[$index]
            $links = ($column.Links | ForEach-Object {
                "<li><a href='$($_.Url)'><strong>$($_.Title)</strong></a><br/><span>$($_.Description)</span></li>"
            }) -join ""
            $html = "<h2>$($column.Title)</h2><p>$($column.Description)</p><ul>$links</ul>"
            Add-PnPPageTextPart -Page $page -Section $columnsSectionNumber -Column ($index + 1) -Text $html | Out-Null
        }
    }

    Set-PnPPage -Identity $pageName -Title $Title -CommentsEnabled:$false -Publish | Out-Null
    return "$SiteUrl/SitePages/$pageName"
}

function Get-ListUrl {
    param([string]$Title)
    $list = Get-PnPList -Identity $Title
    return "$SiteUrl/$($list.RootFolder.ServerRelativeUrl.TrimStart('/').Split('/', 3)[-1])/AllItems.aspx"
}

function Set-ReportingView {
    param([string]$ListTitle, [string[]]$Fields)
    Set-PnPView -List $ListTitle -Identity "All Items" -Fields $Fields -Values @{ RowLimit = [uint32]999; Paged = $true } | Out-Null
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

<# Superseded dashboard block retained temporarily for traceability.
function Add-HubDashboardPart {

    function Add-PatchComplianceDashboardPart {
        param([string]$PageUrl)
        $pageName = $PageUrl.Split('/')[-1]
        $page = Get-PnPPage -Identity $pageName
        $items = @(Get-PnPListItem -List "Intune Patch Compliance" -PageSize 5000 -Fields "complianceStatus", "reportingMonth", "requiredKB", "patchTuesdayDate")
        $counts = @{
            "Compliant" = 0
            "Non-Compliant" = 0
            "Stale" = 0
            "OS End of Service" = 0
            "Unknown" = 0
        }
        foreach ($item in $items) {
            $status = [string]$item.FieldValues.complianceStatus
            if ($counts.ContainsKey($status)) { $counts[$status]++ } else { $counts["Unknown"]++ }
        }
        $total = $items.Count
        $evaluated = $counts["Compliant"] + $counts["Non-Compliant"]
        $percentage = if ($evaluated -gt 0) { [math]::Round(($counts["Compliant"] / $evaluated) * 100, 1) } else { 0 }
        $compliantDegrees = if ($total -gt 0) { [math]::Round(($counts["Compliant"] / $total) * 360, 2) } else { 0 }
        $nonCompliantDegrees = if ($total -gt 0) { [math]::Round(($counts["Non-Compliant"] / $total) * 360, 2) } else { 0 }
        $staleDegrees = if ($total -gt 0) { [math]::Round(($counts["Stale"] / $total) * 360, 2) } else { 0 }
        $eolDegrees = if ($total -gt 0) { [math]::Round(($counts["OS End of Service"] / $total) * 360, 2) } else { 0 }
        $stopCompliant = $compliantDegrees
        $stopNonCompliant = $stopCompliant + $nonCompliantDegrees
        $stopStale = $stopNonCompliant + $staleDegrees
        $stopEol = $stopStale + $eolDegrees
        $sample = $items | Select-Object -First 1
        $month = if ($sample) { [string]$sample.FieldValues.reportingMonth } else { "Not available" }
        $requiredKb = if ($sample) { [string]$sample.FieldValues.requiredKB } else { "Not available" }
        $patchTuesday = if ($sample) { [string]$sample.FieldValues.patchTuesdayDate } else { "Not available" }
        $html = @"
    <div style='font-family:Segoe UI,Arial;background:#0f1720;color:#f8fafc;padding:24px;border:1px solid #2a3742;border-radius:8px'>
    <div style='font-size:20px;font-weight:700;margin-bottom:6px'>Current Patch Compliance Snapshot</div>
    <div style='color:#cbd5e1;margin-bottom:22px'>Reporting month: <strong style='color:#fff'>$month</strong> &nbsp;|&nbsp; Patch Tuesday: <strong style='color:#fff'>$patchTuesday</strong> &nbsp;|&nbsp; Baseline example: <strong style='color:#fff'>$requiredKb</strong></div>
    <table style='width:100%;border-collapse:collapse'><tr>
    <td style='width:34%;text-align:center;vertical-align:middle'><div style='width:210px;height:210px;margin:auto;border-radius:50%;background:conic-gradient(#22c55e 0deg $stopCompliant`deg,#ef4444 $stopCompliant`deg $stopNonCompliant`deg,#eab308 $stopNonCompliant`deg $stopStale`deg,#f97316 $stopStale`deg $stopEol`deg,#64748b $stopEol`deg 360deg);position:relative'><div style='position:absolute;inset:42px;background:#17212b;border-radius:50%;display:flex;align-items:center;justify-content:center;flex-direction:column'><span style='font-size:34px;font-weight:700'>$percentage%</span><span style='font-size:12px;color:#cbd5e1'>evaluable devices</span></div></div></td>
    <td style='width:66%;vertical-align:middle'><table style='width:100%;border-collapse:separate;border-spacing:8px'><tr><td style='background:#17212b;border-left:5px solid #22c55e;padding:14px'><strong>Compliant</strong><div style='font-size:28px;font-weight:700'>$($counts["Compliant"])</div></td><td style='background:#17212b;border-left:5px solid #ef4444;padding:14px'><strong>Non-Compliant</strong><div style='font-size:28px;font-weight:700'>$($counts["Non-Compliant"])</div></td></tr><tr><td style='background:#17212b;border-left:5px solid #eab308;padding:14px'><strong>Stale</strong><div style='font-size:28px;font-weight:700'>$($counts["Stale"])</div></td><td style='background:#17212b;border-left:5px solid #f97316;padding:14px'><strong>OS End of Service</strong><div style='font-size:28px;font-weight:700'>$($counts["OS End of Service"])</div></td></tr><tr><td style='background:#17212b;border-left:5px solid #64748b;padding:14px'><strong>Unknown</strong><div style='font-size:28px;font-weight:700'>$($counts["Unknown"])</div></td><td style='background:#17212b;border-left:5px solid #60a5fa;padding:14px'><strong>Total Windows devices</strong><div style='font-size:28px;font-weight:700'>$total</div></td></tr></table></td>
    </tr></table>
    <div style='margin-top:20px;padding:15px;background:#17212b;color:#cbd5e1;line-height:1.6'><strong style='color:#fff'>Method:</strong> Compliant means the Intune-reported OS revision meets or exceeds the official Microsoft security cumulative-update baseline for that Windows servicing branch. The percentage is Compliant / (Compliant + Non-Compliant). Stale (no check-in within 14 days), unsupported, and unknown devices remain visible but are excluded from the percentage.</div>
    </div>
    "@
        $sectionNumber = $page.Sections.Count + 1
        Add-PnPPageSection -Page $page -SectionTemplate OneColumn | Out-Null
        Add-PnPPageTextPart -Page $page -Section $sectionNumber -Column 1 -Text $html | Out-Null
        Set-PnPPage -Identity $pageName -CommentsEnabled:$false -Publish | Out-Null
    }
    Set-ReportingView "Intune Patch Compliance" @("deviceName", "userPrincipalName", "osBranch", "installedBuild", "requiredKB", "requiredBuild", "complianceStatus", "lastCheckInDays", "lastCheckIn", "reportingMonth", "baselineSourceUrl")
    param([string]$PageUrl)
    Add-PatchComplianceDashboardPart {
        param([string]$PageUrl)
        $pageName = $PageUrl.Split('/')[-1]
        $page = Get-PnPPage -Identity $pageName
        $items = @(Get-PnPListItem -List "Intune Patch Compliance" -PageSize 5000 -Fields "complianceStatus", "reportingMonth", "requiredKB", "patchTuesdayDate")
        $counts = @{ "Compliant" = 0; "Non-Compliant" = 0; "Stale" = 0; "OS End of Service" = 0; "Unknown" = 0 }
        foreach ($item in $items) {
            $status = [string]$item.FieldValues.complianceStatus
            if ($counts.ContainsKey($status)) { $counts[$status]++ } else { $counts["Unknown"]++ }
        }
        $total = $items.Count
        $evaluated = $counts["Compliant"] + $counts["Non-Compliant"]
        $percentage = if ($evaluated -gt 0) { [math]::Round(($counts["Compliant"] / $evaluated) * 100, 1) } else { 0 }
        $stops = @()
        $running = 0
        foreach ($status in @("Compliant", "Non-Compliant", "Stale", "OS End of Service")) {
            $running += if ($total -gt 0) { ($counts[$status] / $total) * 360 } else { 0 }
            $stops += [math]::Round($running, 2)
        }
        $sample = $items | Select-Object -First 1
        $month = if ($sample) { [string]$sample.FieldValues.reportingMonth } else { "Not available" }
        $requiredKb = if ($sample) { [string]$sample.FieldValues.requiredKB } else { "Not available" }
        $patchTuesday = if ($sample) { [string]$sample.FieldValues.patchTuesdayDate } else { "Not available" }
        $html = @"
<div style='font-family:Segoe UI,Arial;background:#0f1720;color:#f8fafc;padding:24px;border:1px solid #2a3742;border-radius:8px'>
<div style='font-size:20px;font-weight:700;margin-bottom:6px'>Current Patch Compliance Snapshot</div>
<div style='color:#cbd5e1;margin-bottom:22px'>Reporting month: <strong style='color:#fff'>$month</strong> | Patch Tuesday: <strong style='color:#fff'>$patchTuesday</strong> | Baseline example: <strong style='color:#fff'>$requiredKb</strong></div>
<table style='width:100%;border-collapse:collapse'><tr><td style='width:34%;text-align:center;vertical-align:middle'><div style='width:210px;height:210px;margin:auto;border-radius:50%;background:conic-gradient(#22c55e 0deg $($stops[0])deg,#ef4444 $($stops[0])deg $($stops[1])deg,#eab308 $($stops[1])deg $($stops[2])deg,#f97316 $($stops[2])deg $($stops[3])deg,#64748b $($stops[3])deg 360deg)'></div><div style='font-size:34px;font-weight:700;margin-top:-125px;margin-bottom:84px'>$percentage%</div></td>
<td style='width:66%;vertical-align:middle'><table style='width:100%;border-collapse:separate;border-spacing:8px'><tr><td style='background:#17212b;border-left:5px solid #22c55e;padding:14px'><strong>Compliant</strong><div style='font-size:28px;font-weight:700'>$($counts["Compliant"])</div></td><td style='background:#17212b;border-left:5px solid #ef4444;padding:14px'><strong>Non-Compliant</strong><div style='font-size:28px;font-weight:700'>$($counts["Non-Compliant"])</div></td></tr><tr><td style='background:#17212b;border-left:5px solid #eab308;padding:14px'><strong>Stale</strong><div style='font-size:28px;font-weight:700'>$($counts["Stale"])</div></td><td style='background:#17212b;border-left:5px solid #f97316;padding:14px'><strong>OS End of Service</strong><div style='font-size:28px;font-weight:700'>$($counts["OS End of Service"])</div></td></tr><tr><td style='background:#17212b;border-left:5px solid #64748b;padding:14px'><strong>Unknown</strong><div style='font-size:28px;font-weight:700'>$($counts["Unknown"])</div></td><td style='background:#17212b;border-left:5px solid #60a5fa;padding:14px'><strong>Total Windows devices</strong><div style='font-size:28px;font-weight:700'>$total</div></td></tr></table></td></tr></table>
<div style='margin-top:20px;padding:15px;background:#17212b;color:#cbd5e1;line-height:1.6'><strong style='color:#fff'>Method:</strong> Compliant means the Intune-reported OS revision meets or exceeds the official Microsoft security cumulative-update baseline for that servicing branch. The percentage excludes stale, unsupported, and unknown devices.</div></div>
    "@
        $sectionNumber = $page.Sections.Count + 1
        Add-PnPPageSection -Page $page -SectionTemplate OneColumn | Out-Null
        Add-PnPPageTextPart -Page $page -Section $sectionNumber -Column 1 -Text $html | Out-Null
        Set-PnPPage -Identity $pageName -CommentsEnabled:$false -Publish | Out-Null
    Add-PnPPageTextPart -Page $page -Section $sectionNumber -Column 1 -Text $html | Out-Null
    Set-PnPPage -Identity $pageName -CommentsEnabled:$false -Publish | Out-Null
}

#>

$script:DashboardContainerStyle = "font-family:Segoe UI,Arial;background:#0f1720;color:#f8fafc;padding:24px;border:1px solid #2a3742;border-radius:8px"

function New-DashboardCard {
    param([string]$Label, [string]$Value, [string]$Subtitle, [string]$AccentColor, [string]$ValueSuffix = "")
    $suffixHtml = if ($ValueSuffix) { "<span style='font-size:18px;color:#cbd5e1'>$ValueSuffix</span>" } else { "" }
    return "<div style='flex:1;background:#17212b;border-top:5px solid $AccentColor;padding:18px'><div style='font-size:12px;color:#cbd5e1'>$Label</div><div style='font-size:42px;font-weight:700'>$Value$suffixHtml</div><div style='color:#cbd5e1'>$Subtitle</div></div>"
}

function New-DashboardBarRow {
    param([string]$Label, [string]$ValueText, [double]$Percent, [string]$Color)
    $width = [math]::Max(1, [math]::Min(100, $Percent))
    return "<div style='display:flex;align-items:center;background-color:#17212b;border-radius:6px;padding:10px 14px;margin-bottom:8px'><div style='width:180px;flex-shrink:0;color:#e2e8f0;font-weight:600'>$Label</div><div style='flex:1;height:18px;background-color:#263442;border-radius:4px;overflow:hidden;margin:0 14px'><div style='height:18px;width:$width%;background-color:$Color'></div></div><div style='width:110px;flex-shrink:0;text-align:right;color:#ffffff;font-weight:700'>$ValueText</div></div>"
}

function New-DashboardHeader {
    param([string]$Title, [string]$Subtitle = "")
    if ($Subtitle) {
        return "<div style='font-size:20px;font-weight:700;margin-bottom:6px'>$Title</div><div style='color:#cbd5e1;margin-bottom:22px'>$Subtitle</div>"
    }
    return "<div style='font-size:20px;font-weight:700;margin-bottom:18px'>$Title</div>"
}

function New-DashboardFooter {
    param([string]$Label, [string]$Text)
    return "<div style='margin-top:20px;padding:15px;background:#17212b;color:#cbd5e1;line-height:1.6'><strong style='color:#fff'>$Label`:</strong> $Text</div>"
}

function Get-HubDashboardHtml {
    $summaryItems = @(Get-PnPListItem -List "Intune Health Summary" -PageSize 1000 -Fields "metricName", "metricValue", "percentage", "status")
    $metrics = @{}
    foreach ($item in $summaryItems) { $metrics[[string]$item.FieldValues.metricName] = $item.FieldValues }
    function Get-MetricValue([string]$Name, [string]$Fallback = "0") {
        if ($metrics.ContainsKey($Name)) { return [string]$metrics[$Name].metricValue }
        return $Fallback
    }
    function Get-MetricPercent([string]$Name) {
        if ($metrics.ContainsKey($Name) -and $metrics[$Name].percentage) { return [double]$metrics[$Name].percentage }
        return 0
    }
    $score = Get-MetricValue "Overall Endpoint Health Score"
    $scoreStatus = if ($metrics.ContainsKey("Overall Endpoint Health Score")) { [string]$metrics["Overall Endpoint Health Score"].status } else { "Unknown" }
    $cards = (New-DashboardCard -Label "HEALTH SCORE" -Value $score -ValueSuffix "/100" -Subtitle $scoreStatus -AccentColor "#22c55e") +
        (New-DashboardCard -Label "MANAGED DEVICES" -Value (Get-MetricValue "Total Devices") -Subtitle "Current Intune scope" -AccentColor "#60a5fa") +
        (New-DashboardCard -Label "MEDIUM / HIGH RISK" -Value (Get-MetricValue "High/Medium Risk Devices") -Subtitle "Devices requiring review" -AccentColor "#f97316") +
        (New-DashboardCard -Label "STALE DEVICES" -Value (Get-MetricValue "Stale / Inactive Devices") -Subtitle "Outside activity window" -AccentColor "#eab308")
    $componentRows = foreach ($component in @(
        "Compliance (25%)", "Encryption (20%)", "Patch/OS currency (20%)",
        "Threat-free (15%)", "Device activity (10%)", "Management health (10%)"
    )) {
        $value = Get-MetricPercent $component
        $color = if ($value -ge 95) { "#22c55e" } elseif ($value -ge 80) { "#eab308" } else { "#ef4444" }
        New-DashboardBarRow -Label $component -ValueText "$([math]::Round($value,1))%" -Percent $value -Color $color
    }
    $html = @"
<div style='$($script:DashboardContainerStyle)'>
$(New-DashboardHeader -Title "Overall Endpoint Health")
<div style='display:flex;gap:10px;margin-bottom:20px'>$cards</div>
<div style='font-size:17px;font-weight:700;margin:8px 0'>Weighted health components</div><div>$($componentRows -join "")</div>
$(New-DashboardFooter -Label "Score weights" -Text "Compliance 25%, encryption 20%, patch/OS currency 20%, threat-free 15%, device activity 10%, and management health 10%.")
</div>
"@
    return $html
}

function Add-HubDashboardPart {
    param([string]$PageUrl)
    $pageName = $PageUrl.Split('/')[-1]
    $page = Get-PnPPage -Identity $pageName
    $html = Get-HubDashboardHtml
    $sectionNumber = $page.Sections.Count + 1
    Add-PnPPageSection -Page $page -SectionTemplate OneColumn | Out-Null
    Add-PnPPageTextPart -Page $page -Section $sectionNumber -Column 1 -Text $html | Out-Null
    Set-PnPPage -Identity $pageName -CommentsEnabled:$false -Publish | Out-Null
}

function Publish-HomePageBody {
    param([string]$PageUrl, [array]$QuickLinkColumns)
    $pageName = $PageUrl.Split('/')[-1]
    $page = Get-PnPPage -Identity $pageName

    $dashboardHtml = Get-HubDashboardHtml
    Add-PnPPageTextPart -Page $page -Section 1 -Column 1 -Text $dashboardHtml | Out-Null

    $linksSection = $page.Sections.Count + 1
    $template = if ($QuickLinkColumns.Count -eq 3) { "ThreeColumn" } else { "TwoColumn" }
    Add-PnPPageSection -Page $page -SectionTemplate $template | Out-Null
    for ($index = 0; $index -lt $QuickLinkColumns.Count; $index++) {
        $column = $QuickLinkColumns[$index]
        $links = ($column.Links | ForEach-Object {
            "<li><a href='$($_.Url)'><strong>$($_.Title)</strong></a><br/><span>$($_.Description)</span></li>"
        }) -join ""
        $columnHtml = "<h2>$($column.Title)</h2><p>$($column.Description)</p><ul>$links</ul>"
        Add-PnPPageTextPart -Page $page -Section $linksSection -Column ($index + 1) -Text $columnHtml | Out-Null
    }

    Set-PnPPage -Identity $pageName -CommentsEnabled:$false -Publish | Out-Null
}

function Add-PatchComplianceDashboardPart {
    param([string]$PageUrl)
    $pageName = $PageUrl.Split('/')[-1]
    $page = Get-PnPPage -Identity $pageName
    $items = @(Get-PnPListItem -List "Intune Patch Compliance" -PageSize 5000 -Fields "complianceStatus", "reportingMonth", "requiredKB", "patchTuesdayDate")
    $managedDevices = @(Get-PnPListItem -List "Intune Devices" -PageSize 5000 -Fields "deviceId")
    $counts = @{ "Compliant" = 0; "Non-Compliant" = 0; "Stale" = 0; "OS End of Service" = 0; "Unknown" = 0; "Not Applicable" = 0 }
    foreach ($item in $items) {
        $status = [string]$item.FieldValues.complianceStatus
        if ($counts.ContainsKey($status)) { $counts[$status]++ } else { $counts["Unknown"]++ }
    }
    $total = $managedDevices.Count
    $counts["Not Applicable"] = [math]::Max(0, $total - $items.Count)
    $percentage = if ($total -gt 0) { [math]::Round(($counts["Compliant"] / $total) * 100, 1) } else { 0 }
    $coverage = if ($total -gt 0) { [math]::Round(($items.Count / $total) * 100, 1) } else { 0 }
    $cards = (New-DashboardCard -Label "OVERALL FLEET COMPLIANCE" -Value "$percentage%" -Subtitle "$($counts["Compliant"]) of $total managed devices" -AccentColor "#22c55e") +
        (New-DashboardCard -Label "ACTION REQUIRED" -Value "$($counts["Non-Compliant"])" -Subtitle "Non-compliant devices" -AccentColor "#ef4444") +
        (New-DashboardCard -Label "TELEMETRY RISK" -Value "$($counts["Stale"] + $counts["Unknown"])" -Subtitle "Stale or unknown" -AccentColor "#eab308") +
        (New-DashboardCard -Label "WINDOWS COVERAGE" -Value "$coverage%" -Subtitle "$($items.Count) of $total evaluated" -AccentColor "#60a5fa")
    $chartRows = foreach ($definition in @(
        @{ Status = "Compliant"; Color = "#22c55e" },
        @{ Status = "Non-Compliant"; Color = "#ef4444" },
        @{ Status = "Stale"; Color = "#eab308" },
        @{ Status = "Unknown"; Color = "#64748b" },
        @{ Status = "OS End of Service"; Color = "#f97316" },
        @{ Status = "Not Applicable"; Color = "#60a5fa" }
    )) {
        $status = $definition.Status
        $count = $counts[$status]
        $statusPercentage = if ($total -gt 0) { [math]::Round(($count / $total) * 100, 1) } else { 0 }
        New-DashboardBarRow -Label $status -ValueText "$count ($statusPercentage%)" -Percent $statusPercentage -Color $definition.Color
    }
    $sample = $items | Select-Object -First 1
    $month = "Not available"
    if ($sample -and $sample.FieldValues.reportingMonth) {
        try {
            $monthDate = [datetime]::ParseExact([string]$sample.FieldValues.reportingMonth, "yyyy-MM", [System.Globalization.CultureInfo]::InvariantCulture)
            $month = $monthDate.ToString("MMMM yyyy", [System.Globalization.CultureInfo]::InvariantCulture)
        } catch {
            $month = [string]$sample.FieldValues.reportingMonth
        }
    }
    $requiredKb = if ($sample) { [string]$sample.FieldValues.requiredKB } else { "Not available" }
    $patchTuesday = if ($sample) { [string]$sample.FieldValues.patchTuesdayDate } else { "Not available" }
    $subtitle = "Reporting month: <strong style='color:#fff'>$month</strong> | Patch Tuesday: <strong style='color:#fff'>$patchTuesday</strong> | Baseline example: <strong style='color:#fff'>$requiredKb</strong>"
    $html = @"
<div style='$($script:DashboardContainerStyle)'>
$(New-DashboardHeader -Title "Patch Compliance Overview" -Subtitle $subtitle)
<div style='display:flex;gap:10px;margin-bottom:20px'>$cards</div>
<div style='font-size:17px;font-weight:700;margin:8px 0'>Fleet status distribution</div><div>$($chartRows -join "")</div>
$(New-DashboardFooter -Label "Management calculation" -Text "Overall compliance is Compliant / All managed devices. Every device remains in the denominator; non-Windows devices appear as Not Applicable, and stale or unknown telemetry is not counted as compliant.")
</div>
"@
    Add-PnPPageTextPart -Page $page -Section 1 -Column 1 -Text $html | Out-Null
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
    JsonExports = Ensure-DocumentLibrary "JSON Exports" "JSONExports"
}

Set-ReportingView "Intune Devices" @("deviceName", "userPrincipalName", "operatingSystem", "osVersion", "complianceState", "isEncrypted", "storagePercentFree", "threatState", "patchStatus", "riskLevel", "daysInactive", "lastSyncDateTime")
Set-ReportingView "Intune Apps" @("appName", "publisher", "deviceName", "userPrincipalName", "installState", "normalizedStatus", "errorCode", "lastSyncDateTime")
Set-ReportingView "Intune Compliance Policies" @("policyName", "deviceName", "userPrincipalName", "complianceState", "lastReportedDateTime")
Set-ReportingView "Intune Config Profiles" @("profileName", "profileType", "deviceName", "userPrincipalName", "status", "lastReportedDateTime")
Set-ReportingView "Intune Patch Compliance" @("deviceName", "userPrincipalName", "osBranch", "installedBuild", "requiredKB", "requiredBuild", "complianceStatus", "lastCheckInDays", "lastCheckIn", "reportingMonth", "baselineSourceUrl")
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
$pageUrls.PatchCompliance = New-OperationsPage "Patch-Compliance" "Patch Compliance" "Current-month Windows quality-update compliance measured against Microsoft's official Patch Tuesday baseline, evaluated across every managed device." @()
$pageUrls.Vulnerability = New-OperationsPage "Vulnerability-Reports" "Vulnerability Reports" "Risk-led endpoint vulnerability reporting and remediation prioritization." @(@{ Title = "Device risk register"; Description = "Devices with multiple actionable risk signals."; Links = @(@{ Title = "Open device risks"; Url = $urls.DeviceRisks; Description = "Risk score, reasons, OS, ownership, and inactivity." }) })
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
        @{ Title = "Patch Compliance"; Url = $pageUrls.PatchCompliance; Description = "Current-month security update compliance, baseline evidence, and device details." },
        @{ Title = "Update Ring Inventory"; Url = $urls.UpdateRingInventory; Description = "Windows Update ring definitions and assignments." },
        @{ Title = "Expedite Updates"; Url = $pageUrls.Expedite; Description = "Expedited update delivery." }
    ) },
    @{ Title = "Risk"; Description = "Prioritize remediation using endpoint risk."; Links = @(
        @{ Title = "Vulnerability Reports"; Url = $pageUrls.Vulnerability; Description = "Vulnerability and remediation reporting." }
    ) }
)

$homeQuickLinks = @(
    @{ Title = "Endpoint estate"; Description = "Inventory and endpoint health."; Links = @(
        @{ Title = "Device Management"; Url = $pageUrls.Device; Description = "Inventory, compliance, Autopilot, and hardware." },
        @{ Title = "Overall Health Score"; Url = $urls.HealthSummary; Description = "Current weighted endpoint health metrics." }
    ) },
    @{ Title = "Configuration and delivery"; Description = "Policy and application operations."; Links = @(
        @{ Title = "Application Management"; Url = $pageUrls.Application; Description = "Apps, deployments, and health." },
        @{ Title = "Policy Management"; Url = $pageUrls.Policy; Description = "Profiles, compliance, and baselines." },
        @{ Title = "Patch Management"; Url = $pageUrls.Patch; Description = "Updates and vulnerability reporting." }
    ) },
    @{ Title = "Insights and enablement"; Description = "Dashboards, automation, and knowledge."; Links = @(
        @{ Title = "Power BI Dashboards"; Url = $urls.PowerBI; Description = "Published operational dashboards." },
        @{ Title = "Full Report (JSON)"; Url = $urls.JsonExports; Description = "Download the complete health-check dataset for automation, Power Automate, or Power BI." },
        @{ Title = "AI Driver Automation"; Url = $pageUrls.DriverAutomation; Description = "Driver discovery, targeting, orchestration, and automation assets." },
        @{ Title = "Documentation"; Url = $urls.Documentation; Description = "Operating procedures and reference material." }
    ) }
)

$pageUrls.Home = New-OperationsPage "Endpoint-Intelligence-Hub" "Endpoint Intelligence Hub" $hubDescription @()

Publish-HomePageBody $pageUrls.Home $homeQuickLinks
Add-ReportingListPart $pageUrls.Home "Intune Health Summary"
Add-ReportingListPart $pageUrls.Device "Intune Devices"
Add-ReportingListPart $pageUrls.Autopilot "Intune Autopilot"
Add-ReportingListPart $pageUrls.Hardware "Intune Devices"
Add-ReportingListPart $pageUrls.Application "Intune Application Inventory"
Add-ReportingListPart $pageUrls.Win32 "Intune Application Inventory"
Add-ReportingListPart $pageUrls.Store "Intune Application Inventory"
Add-ReportingListPart $pageUrls.AppHealth "Intune Apps"
Add-PatchComplianceDashboardPart $pageUrls.PatchCompliance
Add-ReportingListPart $pageUrls.PatchCompliance "Intune Patch Compliance"
Add-ReportingListPart $pageUrls.Vulnerability "Intune Device Risks"
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
Add-PnPNavigationNode -Location QuickLaunch -Parent $patchNode -Title "Patch Compliance" -Url $pageUrls.PatchCompliance | Out-Null
Add-PnPNavigationNode -Location QuickLaunch -Parent $patchNode -Title "Update Ring Inventory" -Url $urls.UpdateRingInventory | Out-Null
Add-PnPNavigationNode -Location QuickLaunch -Parent $patchNode -Title "Expedite Updates" -Url $pageUrls.Expedite | Out-Null
Add-PnPNavigationNode -Location QuickLaunch -Parent $patchNode -Title "Vulnerability Reports" -Url $pageUrls.Vulnerability | Out-Null

Add-PnPNavigationNode -Location QuickLaunch -Title "Power BI Dashboards" -Url $urls.PowerBI | Out-Null
Add-PnPNavigationNode -Location QuickLaunch -Title "Full Report (JSON)" -Url $urls.JsonExports | Out-Null
Add-PnPNavigationNode -Location QuickLaunch -Title "AI Driver Automation" -Url $pageUrls.DriverAutomation | Out-Null
Add-PnPNavigationNode -Location QuickLaunch -Title "Documentation" -Url $urls.Documentation | Out-Null

Write-Host "`nEndpoint Intelligence Hub site design completed."
Write-Host "Home: $($pageUrls.Home)"
$jsonExportsList = Get-PnPList -Identity "JSON Exports"
Write-Host "JSON Exports library ready. Add this Function App setting: SP_LIST_JSON_EXPORTS_ID=$($jsonExportsList.Id)"
$web = Get-PnPWeb -Includes WelcomePage
Write-Host "Configured home page: $($web.WelcomePage)"
Write-Host "Quick launch:"
Get-PnPNavigationNode -Location QuickLaunch | ForEach-Object {
    Write-Host "- $($_.Title)"
    Get-PnPProperty -ClientObject $_ -Property Children | Out-Null
    $_.Children | ForEach-Object { Write-Host "  - $($_.Title)" }
}

Set-PnPWeb -Title "Endpoint Intelligence Hub" -Description $hubDescription