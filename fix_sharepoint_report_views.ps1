[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)] [string]$SiteUrl,
    [Parameter(Mandatory = $true)] [string]$TenantId,
    [Parameter(Mandatory = $true)] [string]$ClientId
)

$ErrorActionPreference = "Stop"
Import-Module PnP.PowerShell
Connect-PnPOnline -Url $SiteUrl -OSLogin -ClientId $ClientId -Tenant $TenantId

$views = [ordered]@{
    "Intune Apps" = @("appName", "publisher", "deviceName", "userPrincipalName", "installState", "normalizedStatus", "errorCode", "lastSyncDateTime")
    "Intune Application Inventory" = @("appName", "appType", "publisher", "displayVersion", "isAssigned", "publishingState", "lastModifiedDateTime")
    "Intune Autopilot" = @("serialNumber", "manufacturer", "model", "groupTag", "enrollmentState", "profileAssignmentStatus", "userPrincipalName", "lastContactedDateTime")
    "Intune Compliance Policies" = @("policyName", "deviceName", "userPrincipalName", "complianceState", "lastReportedDateTime")
    "Intune Config Profiles" = @("profileName", "profileType", "deviceName", "userPrincipalName", "status", "lastReportedDateTime")
    "Intune Patch Compliance" = @("ringName", "deviceName", "userPrincipalName", "status", "lastReportedDateTime")
    "Intune Devices" = @("deviceName", "userPrincipalName", "operatingSystem", "osVersion", "complianceState", "isEncrypted", "storagePercentFree", "threatState", "patchStatus", "riskLevel", "daysInactive", "lastSyncDateTime")
    "Intune Device Risks" = @("deviceName", "userPrincipalName", "riskLevel", "riskScore", "riskReasons", "daysInactive", "lastSyncDateTime")
    "Intune Health Summary" = @("category", "metricName", "metricValue", "percentage", "status", "description", "LastRunId")
    "Intune Compliance Policy Inventory" = @("policyName", "policyType", "description", "platforms", "technologies", "isAssigned", "createdDateTime", "lastModifiedDateTime")
    "Intune Configuration Profile Inventory" = @("policyName", "policyType", "description", "platforms", "technologies", "isAssigned", "createdDateTime", "lastModifiedDateTime")
    "Intune Update Ring Inventory" = @("policyName", "policyType", "description", "platforms", "technologies", "isAssigned", "createdDateTime", "lastModifiedDateTime")
}

foreach ($listName in $views.Keys) {
    Set-PnPView -List $listName -Identity "All Items" -Fields $views[$listName] -Values @{ RowLimit = [uint32]999; Paged = $true } | Out-Null
    Write-Host "Updated first-click view: $listName"
}
