[CmdletBinding()]
param(
    [string]$SiteUrl = "https://nttdsdws.sharepoint.com/sites/IntuneReporting",
    [string]$TenantId = "e815be3c-0e2d-4721-9249-1f5dee892931",
    [string]$ClientId = "a71fa672-4cdf-4e00-bcf8-b634d40bac9e"
)

$ErrorActionPreference = "Stop"
Import-Module PnP.PowerShell
Connect-PnPOnline -Url $SiteUrl -OSLogin -ClientId $ClientId -Tenant $TenantId

$web = Get-PnPWeb -Includes WelcomePage,Title,Description
$pageNames = @(
    "Autopilot.aspx", "Hardware-Reports.aspx", "Win32-Apps.aspx", "Store-Apps.aspx",
    "App-Health.aspx", "Security-Baselines.aspx", "Expedite-Updates.aspx", "Patch-Compliance.aspx",
    "Vulnerability-Reports.aspx", "Device-Management.aspx",
    "Application-Management.aspx", "Policy-Management.aspx", "Patch-Management.aspx",
    "Endpoint-Intelligence-Hub.aspx", "AI-Driver-Automation.aspx"
)
$pages = @($pageNames | ForEach-Object {
    Get-PnPPage -Identity $_ -ErrorAction SilentlyContinue
})
$pageWebParts = @($pages | ForEach-Object {
    [pscustomobject]@{
        Page = $_.Name
        ListWebParts = @($_.Controls | Where-Object { $_.Title -eq "List" }).Count
    }
})
$pagesWithNativeHeader = @($pages | Where-Object { $_.PageHeader.Type -ne "None" } | ForEach-Object { $_.Name })
$libraries = @("Power BI Dashboards", "Automation", "Documentation", "JSON Exports") | ForEach-Object {
    $list = Get-PnPList -Identity $_ -ErrorAction SilentlyContinue
    [pscustomobject]@{
        Title = $_
        Exists = [bool]$list
        Url = if ($list) { $list.RootFolder.ServerRelativeUrl } else { $null }
    }
}
$navigation = @(Get-PnPNavigationNode -Location QuickLaunch | ForEach-Object {
    Get-PnPProperty -ClientObject $_ -Property Children | Out-Null
    [pscustomobject]@{
        Title = $_.Title
        Url = $_.Url
        Children = @($_.Children | ForEach-Object { $_.Title })
    }
})
$patchItems = @(Get-PnPListItem -List "Intune Patch Compliance" -PageSize 5000 -Fields "complianceStatus", "requiredKB", "reportingMonth")
$patchPage = Get-PnPPage -Identity "Patch-Compliance.aspx"
$patchPageText = ($patchPage.Controls | ForEach-Object { $_.Text }) -join " "
$patchTextControls = @($patchPage.Controls | Where-Object { $_.Text })
$homePage = Get-PnPPage -Identity "Endpoint-Intelligence-Hub.aspx"
$homePageText = ($homePage.Controls | ForEach-Object { $_.Text }) -join " "

$patchDistributionStart = $patchPageText.IndexOf("Fleet status distribution")
$patchDistributionEnd = $patchPageText.IndexOf("Management calculation:")
$patchDistributionHtml = if ($patchDistributionStart -ge 0 -and $patchDistributionEnd -gt $patchDistributionStart) { $patchPageText.Substring($patchDistributionStart, $patchDistributionEnd - $patchDistributionStart) } else { "" }
$homeComponentsStart = $homePageText.IndexOf("Weighted health components")
$homeComponentsEnd = $homePageText.IndexOf("Score weights:")
$homeComponentsHtml = if ($homeComponentsStart -ge 0 -and $homeComponentsEnd -gt $homeComponentsStart) { $homePageText.Substring($homeComponentsStart, $homeComponentsEnd - $homeComponentsStart) } else { "" }
$endpointAnalyticsPage = Get-PnPPage -Identity "Endpoint-Analytics.aspx" -ErrorAction SilentlyContinue
$jsonExportsLibrary = Get-PnPList -Identity "JSON Exports" -ErrorAction SilentlyContinue
$jsonExportFile = if ($jsonExportsLibrary) { Get-PnPFile -Url "$($jsonExportsLibrary.RootFolder.ServerRelativeUrl)/EndpointIntelligenceHub_Latest.json" -ErrorAction SilentlyContinue } else { $null }
$driverAutomationPage = Get-PnPPage -Identity "AI-Driver-Automation.aspx"
$driverAutomationListWebParts = @($driverAutomationPage.Controls | Where-Object { $_.Title -eq "List" }).Count

[pscustomobject]@{
    SiteTitle = $web.Title
    SiteDescription = $web.Description
    HomePage = $web.WelcomePage
    PagesFound = $pages.Count
    ExpectedPages = $pageNames.Count
    PageWebParts = $pageWebParts
    Libraries = $libraries
    Navigation = $navigation
    EndpointAnalyticsRemoved = -not [bool]$endpointAnalyticsPage
    DriverAutomationHasNoDuplicateList = $driverAutomationListWebParts -eq 0
    JsonExportsLibraryExists = [bool]$jsonExportsLibrary
    JsonExportFileExists = [bool]$jsonExportFile
    NoDuplicateNativeHeaders = $pagesWithNativeHeader.Count -eq 0
    PagesWithNativeHeader = $pagesWithNativeHeader
    PatchCompliance = [pscustomobject]@{
        PageControls = $patchPage.Controls.Count
        ListWebParts = @($patchPage.Controls | Where-Object { $_.Title -eq "List" }).Count
        ListItems = $patchItems.Count
        PopulatedStatuses = @($patchItems | Where-Object { $_.FieldValues.complianceStatus }).Count
        HasManagementTitle = $patchPageText.Contains("Management Patch Compliance Overview")
        HasFleetRatio = $patchPageText.Contains("37.5%")
        HasDistributionGraph = $patchPageText.Contains("Fleet status distribution")
        UsesAllManagedDevices = $patchPageText.Contains("All managed devices")
        OverviewIsFirstTextControl = $patchTextControls.Count -gt 0 -and $patchTextControls[0].Text.Contains("Patch Compliance")
        DashboardImmediatelyAfterIntro = $patchTextControls.Count -gt 1 -and $patchTextControls[1].Text.Contains("Patch Compliance Overview")
        HasWrittenReportingMonth = $patchPageText.Contains("September 2026")
        OverviewHasNumericReportingMonth = $patchTextControls.Count -gt 1 -and $patchTextControls[1].Text -match "Reporting month:.{0,150}>2026-09<"
        HomeHealthDashboardRestored = $homePageText.Contains("Overall Endpoint Health")
        HomeWeightedChartRestored = $homePageText.Contains("Weighted health components")
        PatchDistributionUsesDivRows = $patchDistributionHtml -and (-not $patchDistributionHtml.Contains("<tr>"))
        HomeComponentsUseDivRows = $homeComponentsHtml -and (-not $homeComponentsHtml.Contains("<tr>"))
        PatchSectionCount = $patchPage.Sections.Count
        PatchDashboardIsSection1 = $patchTextControls.Count -gt 1 -and $patchTextControls[1].Text.Contains("Patch Compliance Overview")
        HomeDashboardBeforeQuickLinks = $homePageText.IndexOf("Overall Endpoint Health") -lt $homePageText.IndexOf("Endpoint estate")
        HomeDashboardImmediatelyAfterIntro = $homePageText.IndexOf("Overall Endpoint Health") -lt $homePageText.IndexOf("Weighted health components")
    }
} | ConvertTo-Json -Depth 6