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
Import-Module PnP.PowerShell
Connect-PnPOnline -Url $SiteUrl -OSLogin -ClientId $ClientId -Tenant $TenantId

$listDefinitions = [ordered]@{
    "Intune Apps" = @(
        "appId", "appName", "publisher", "deviceId", "deviceName", "userPrincipalName",
        "installState", "errorCode", "issue", "recommendedSolution", "lastSyncDateTime", "normalizedStatus"
    )
    "Intune Compliance Policies" = @(
        "policyId", "policyName", "deviceId", "deviceName", "userPrincipalName",
        "complianceState", "issue", "recommendedSolution", "lastReportedDateTime", "normalizedStatus"
    )
    "Intune Config Profiles" = @(
        "profileId", "profileName", "profileType", "deviceId", "deviceName", "userPrincipalName",
        "status", "lastReportedDateTime", "normalizedStatus"
    )
    "Intune Patch Compliance" = @(
        "ringId", "ringName", "deviceId", "deviceName", "userPrincipalName",
        "status", "lastReportedDateTime", "normalizedStatus"
    )
    "Intune Devices" = @(
        "deviceId", "deviceName", "userId", "userPrincipalName", "operatingSystem", "osVersion",
        "manufacturer", "model", "serialNumber", "complianceState", "managementState", "enrollmentType",
        "jailBroken", "isEncrypted", "totalStorageSpaceInBytes", "freeStorageSpaceInBytes",
        "storagePercentFree", "storageHealth", "threatState", "daysInactive", "isStale", "patchStatus",
        "riskScore", "riskLevel", "riskReasons", "gracePeriodExpirationDateTime",
        "managementCertExpirationDate", "certificateStatus", "azureAdDeviceId", "ownerType",
        "managementAgent", "autopilotEnrolled", "approvalPending", "joinType", "physicalMemoryInBytes",
        "processorArchitecture", "enrollmentProfileName", "enrolledDateTime", "lastSyncDateTime"
    )
    "Intune Application Inventory" = @(
        "appId", "appName", "appType", "publisher", "displayVersion", "owner", "developer",
        "publishingState", "isAssigned", "createdDateTime", "lastModifiedDateTime"
    )
    "Intune Autopilot" = @(
        "autopilotId", "serialNumber", "manufacturer", "model", "groupTag", "enrollmentState",
        "profileAssignmentStatus", "profileAssignmentDetail", "profileAssignedDateTime",
        "lastContactedDateTime", "userPrincipalName", "managedDeviceId", "azureAdDeviceId",
        "userlessEnrollmentStatus"
    )
    "Intune Device Risks" = @(
        "deviceId", "deviceName", "userPrincipalName", "operatingSystem", "osVersion", "riskLevel",
        "riskScore", "riskReasons", "daysInactive", "lastSyncDateTime"
    )
    "Intune Health Summary" = @(
        "category", "metricName", "metricValue", "percentage", "status", "description"
    )
}

$listIds = [ordered]@{}
foreach ($listName in $listDefinitions.Keys) {
    $list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
    if (-not $list) {
        New-PnPList -Title $listName -Template GenericList -OnQuickLaunch:$false | Out-Null
        $list = Get-PnPList -Identity $listName
        Write-Host "Created list '$listName'"
    }
    $listIds[$listName] = $list.Id

    $existingFields = @(Get-PnPField -List $listName | ForEach-Object { $_.InternalName })
    foreach ($fieldName in @("Key", "LastRunId") + $listDefinitions[$listName]) {
        if ($existingFields -contains $fieldName) { continue }
        Add-PnPField -List $listName -DisplayName $fieldName -InternalName $fieldName -Type Text -AddToDefaultView | Out-Null
        Write-Host "Added '$fieldName' to '$listName'"
    }
    Set-PnPList -Identity $listName -ListExperience NewExperience | Out-Null
}

Write-Host "`n--- Azure Function App settings ---"
Write-Host "SP_LIST_APPS_ID=$($listIds['Intune Apps'])"
Write-Host "SP_LIST_COMPLIANCE_ID=$($listIds['Intune Compliance Policies'])"
Write-Host "SP_LIST_CONFIG_PROFILES_ID=$($listIds['Intune Config Profiles'])"
Write-Host "SP_LIST_PATCH_ID=$($listIds['Intune Patch Compliance'])"
Write-Host "SP_LIST_DEVICES_ID=$($listIds['Intune Devices'])"
Write-Host "SP_LIST_APP_INVENTORY_ID=$($listIds['Intune Application Inventory'])"
Write-Host "SP_LIST_AUTOPILOT_ID=$($listIds['Intune Autopilot'])"
Write-Host "SP_LIST_DEVICE_RISKS_ID=$($listIds['Intune Device Risks'])"
Write-Host "SP_LIST_HEALTH_SUMMARY_ID=$($listIds['Intune Health Summary'])"
