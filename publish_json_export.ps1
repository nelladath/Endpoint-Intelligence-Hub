[CmdletBinding()]
param(
    [string]$SiteUrl = "https://nttdsdws.sharepoint.com/sites/IntuneReporting",
    [string]$TenantId = "e815be3c-0e2d-4721-9249-1f5dee892931",
    [string]$ClientId = "a71fa672-4cdf-4e00-bcf8-b634d40bac9e"
)

# Populates the JSON Exports library directly from the already-synced SharePoint reporting
# lists, so a full data export is available immediately without waiting for an Azure Function
# redeploy. The Azure Function's own pipeline run will overwrite this file with fresh data.

$ErrorActionPreference = "Stop"
Import-Module PnP.PowerShell
Connect-PnPOnline -Url $SiteUrl -OSLogin -ClientId $ClientId -Tenant $TenantId

$listFields = [ordered]@{
    "Intune Apps" = @("appId", "appName", "publisher", "deviceId", "deviceName", "userPrincipalName", "installState", "errorCode", "issue", "recommendedSolution", "lastSyncDateTime", "normalizedStatus")
    "Intune Compliance Policies" = @("policyId", "policyName", "deviceId", "deviceName", "userPrincipalName", "complianceState", "issue", "recommendedSolution", "lastReportedDateTime", "normalizedStatus")
    "Intune Compliance Policy Inventory" = @("policyId", "policyName", "policyType", "odataType", "description", "platforms", "technologies", "isAssigned", "roleScopeTagIds", "createdDateTime", "lastModifiedDateTime")
    "Intune Config Profiles" = @("profileId", "profileName", "profileType", "deviceId", "deviceName", "userPrincipalName", "status", "lastReportedDateTime", "normalizedStatus")
    "Intune Configuration Profile Inventory" = @("policyId", "policyName", "policyType", "odataType", "description", "platforms", "technologies", "isAssigned", "roleScopeTagIds", "createdDateTime", "lastModifiedDateTime")
    "Intune Patch Compliance" = @("deviceId", "azureAdDeviceId", "deviceName", "userPrincipalName", "serialNumber", "manufacturer", "model", "osBranch", "installedBuild", "installedRevision", "reportingMonth", "patchTuesdayDate", "requiredKB", "requiredBuild", "requiredRevision", "baselineReleaseDate", "complianceStatus", "lastCheckIn", "lastCheckInDays", "baselineSourceUrl", "evaluatedAtUtc")
    "Intune Update Ring Inventory" = @("policyId", "policyName", "policyType", "odataType", "description", "platforms", "technologies", "isAssigned", "roleScopeTagIds", "createdDateTime", "lastModifiedDateTime")
    "Intune Devices" = @("deviceId", "deviceName", "userId", "userPrincipalName", "operatingSystem", "osVersion", "manufacturer", "model", "serialNumber", "complianceState", "managementState", "enrollmentType", "jailBroken", "isEncrypted", "totalStorageSpaceInBytes", "freeStorageSpaceInBytes", "storagePercentFree", "storageHealth", "threatState", "daysInactive", "isStale", "patchStatus", "riskScore", "riskLevel", "riskReasons", "gracePeriodExpirationDateTime", "managementCertExpirationDate", "certificateStatus", "azureAdDeviceId", "ownerType", "managementAgent", "autopilotEnrolled", "approvalPending", "joinType", "physicalMemoryInBytes", "processorArchitecture", "enrollmentProfileName", "enrolledDateTime", "lastSyncDateTime")
    "Intune Application Inventory" = @("appId", "appName", "appType", "publisher", "displayVersion", "owner", "developer", "publishingState", "isAssigned", "createdDateTime", "lastModifiedDateTime")
    "Intune Autopilot" = @("autopilotId", "serialNumber", "manufacturer", "model", "groupTag", "enrollmentState", "profileAssignmentStatus", "profileAssignmentDetail", "profileAssignedDateTime", "lastContactedDateTime", "userPrincipalName", "managedDeviceId", "azureAdDeviceId", "userlessEnrollmentStatus")
    "Intune Device Risks" = @("deviceId", "deviceName", "userPrincipalName", "operatingSystem", "osVersion", "riskLevel", "riskScore", "riskReasons", "daysInactive", "lastSyncDateTime")
    "Intune Health Summary" = @("category", "metricName", "metricValue", "percentage", "status", "description")
}

function Get-ListRecords {
    param([string]$ListTitle, [string[]]$Fields)
    $items = @(Get-PnPListItem -List $ListTitle -PageSize 5000 -Fields $Fields)
    return @($items | ForEach-Object {
        $record = [ordered]@{}
        foreach ($field in $Fields) { $record[$field] = $_.FieldValues[$field] }
        [pscustomobject]$record
    })
}

Write-Host "Reading reporting lists..."
$records = [ordered]@{}
foreach ($listTitle in $listFields.Keys) {
    $key = ($listTitle -replace "^Intune ", "") -replace " ", ""
    $key = [char]::ToLower($key[0]) + $key.Substring(1)
    $records[$key] = Get-ListRecords -ListTitle $listTitle -Fields $listFields[$listTitle]
    Write-Host "  $listTitle : $($records[$key].Count) records"
}

$healthMetrics = $records["healthSummary"]
$overallHealth = $healthMetrics | Where-Object { $_.metricName -eq "Overall Endpoint Health Score" } | Select-Object -First 1
$runIdMetric = $healthMetrics | Where-Object { $_.metricName -eq "Run ID" } | Select-Object -First 1
$refreshedMetric = $healthMetrics | Where-Object { $_.metricName -eq "Last Refreshed (UTC)" } | Select-Object -First 1

$exportedAtUtc = (Get-Date).ToUniversalTime().ToString("o")
$jsonExport = [ordered]@{
    reportName = "Endpoint Intelligence Hub"
    description = "Centralized portal for Microsoft Intune, Windows Autopilot, AI-assisted driver automation, patch management, application delivery, device compliance, endpoint health, security risk, hardware intelligence, and operational analytics."
    lastPipelineRunId = if ($runIdMetric) { $runIdMetric.metricValue } else { $null }
    lastPipelineGeneratedAtUtc = if ($refreshedMetric) { $refreshedMetric.metricValue } else { $null }
    jsonExportedAtUtc = $exportedAtUtc
    jsonExportSource = "publish_json_export.ps1 (direct SharePoint read; superseded automatically by the next Azure Function pipeline run)"
    overallEndpointHealthScore = if ($overallHealth) { $overallHealth.metricValue } else { $null }
    apps = $records["apps"]
    compliancePolicies = $records["compliancePolicies"]
    complianceInventory = $records["compliancePolicyInventory"]
    configProfiles = $records["configProfiles"]
    configurationInventory = $records["configurationProfileInventory"]
    patchCompliance = $records["patchCompliance"]
    updateRingInventory = $records["updateRingInventory"]
    devices = $records["devices"]
    applicationInventory = $records["applicationInventory"]
    autopilot = $records["autopilot"]
    deviceRisks = $records["deviceRisks"]
    healthMetrics = $healthMetrics
}

$jsonBytes = [System.Text.Encoding]::UTF8.GetBytes(($jsonExport | ConvertTo-Json -Depth 8))
$library = Get-PnPList -Identity "JSON Exports"
Add-PnPFile -FileName "EndpointIntelligenceHub_Latest.json" -Folder $library.RootFolder.ServerRelativeUrl -Stream ([System.IO.MemoryStream]::new($jsonBytes)) | Out-Null

$totalRecords = ($records.Values | ForEach-Object { $_.Count } | Measure-Object -Sum).Sum
Write-Host "Uploaded EndpointIntelligenceHub_Latest.json with $totalRecords total records across $($records.Keys.Count) domains."
