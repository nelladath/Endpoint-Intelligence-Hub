[CmdletBinding()]
param(
    [string]$SiteUrl = "https://nttdsdws.sharepoint.com/sites/IntuneReporting",
    [string]$TenantId = "e815be3c-0e2d-4721-9249-1f5dee892931",
    [string]$ClientId = "a71fa672-4cdf-4e00-bcf8-b634d40bac9e"
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
        "deviceId", "azureAdDeviceId", "deviceName", "userPrincipalName", "serialNumber",
        "manufacturer", "model", "osBranch", "installedBuild", "installedRevision",
        "reportingMonth", "patchTuesdayDate", "requiredKB", "requiredBuild", "requiredRevision",
        "baselineReleaseDate", "complianceStatus", "lastCheckIn", "lastCheckInDays",
        "baselineSourceUrl", "evaluatedAtUtc"
    )
    "Intune Devices" = @(
        "deviceId", "deviceName", "userId", "userPrincipalName", "operatingSystem", "osVersion",
        "manufacturer", "model", "serialNumber", "complianceState", "managementState", "enrollmentType",
        "jailBroken", "isEncrypted", "totalStorageSpaceInBytes", "freeStorageSpaceInBytes",
        "storagePercentFree", "storageHealth", "threatState", "daysInactive", "isStale",
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
    "Intune Compliance Policy Inventory" = @(
        "policyId", "policyName", "policyType", "odataType", "description", "platforms",
        "technologies", "isAssigned", "roleScopeTagIds", "createdDateTime", "lastModifiedDateTime"
    )
    "Intune Configuration Profile Inventory" = @(
        "policyId", "policyName", "policyType", "odataType", "description", "platforms",
        "technologies", "isAssigned", "roleScopeTagIds", "createdDateTime", "lastModifiedDateTime"
    )
    "Intune Update Ring Inventory" = @(
        "policyId", "policyName", "policyType", "odataType", "description", "platforms",
        "technologies", "isAssigned", "roleScopeTagIds", "createdDateTime", "lastModifiedDateTime"
    )
}

$listIds = [ordered]@{}

function Convert-FieldNameToDisplayName {
    param([string]$FieldName)

    $name = $FieldName -creplace "([a-z0-9])([A-Z])", '$1 $2'
    $name = $name -creplace "([A-Z]+)([A-Z][a-z])", '$1 $2'
    $name = $name -replace "[_-]", " "
    $name = (Get-Culture).TextInfo.ToTitleCase($name.ToLower())
    $replacements = @{
        "Id" = "ID"
        "Os" = "OS"
        "Ad" = "AD"
        "Upn" = "UPN"
        "Odata" = "OData"
        "Ios" = "iOS"
        "Macos" = "macOS"
        "Win32" = "Win32"
        "Mdm" = "MDM"
        "Esp" = "ESP"
    }
    foreach ($replacement in $replacements.GetEnumerator()) {
        $name = $name -replace "\b$($replacement.Key)\b", $replacement.Value
    }
    return $name
}

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
        $displayName = Convert-FieldNameToDisplayName $fieldName
        if ($existingFields -notcontains $fieldName) {
            Add-PnPField -List $listName -DisplayName $displayName -InternalName $fieldName -Type Text -AddToDefaultView | Out-Null
            Write-Host "Added '$displayName' ('$fieldName') to '$listName'"
        } else {
            Set-PnPField -List $listName -Identity $fieldName -Values @{Title = $displayName} -UpdateExistingLists | Out-Null
        }
    }
    Set-PnPList -Identity $listName -ListExperience NewExperience | Out-Null
}

$obsoleteDeviceField = Get-PnPField -List "Intune Devices" -Identity "patchStatus" -ErrorAction SilentlyContinue
if ($obsoleteDeviceField) {
    Remove-PnPField -List "Intune Devices" -Identity "patchStatus" -Force
    Write-Host "Removed obsolete 'patchStatus' from 'Intune Devices'"
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
Write-Host "SP_LIST_COMPLIANCE_INVENTORY_ID=$($listIds['Intune Compliance Policy Inventory'])"
Write-Host "SP_LIST_CONFIGURATION_INVENTORY_ID=$($listIds['Intune Configuration Profile Inventory'])"
Write-Host "SP_LIST_UPDATE_RING_INVENTORY_ID=$($listIds['Intune Update Ring Inventory'])"
