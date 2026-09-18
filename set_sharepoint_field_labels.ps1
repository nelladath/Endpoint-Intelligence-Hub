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

$listNames = @(
    "Intune Apps", "Intune Compliance Policies", "Intune Config Profiles", "Intune Patch Compliance",
    "Intune Devices", "Intune Application Inventory", "Intune Autopilot", "Intune Device Risks",
    "Intune Health Summary", "Intune Compliance Policy Inventory",
    "Intune Configuration Profile Inventory", "Intune Update Ring Inventory"
)

function Convert-FieldNameToDisplayName {
    param([string]$FieldName)
    $name = $FieldName -creplace "([a-z0-9])([A-Z])", '$1 $2'
    $name = $name -creplace "([A-Z]+)([A-Z][a-z])", '$1 $2'
    $name = $name -replace "[_-]", " "
    $name = (Get-Culture).TextInfo.ToTitleCase($name.ToLower())
    $replacements = @{
        "Id" = "ID"; "Os" = "OS"; "Ad" = "AD"; "Upn" = "UPN"; "Odata" = "OData"
        "Ios" = "iOS"; "Macos" = "macOS"; "Win32" = "Win32"; "Mdm" = "MDM"; "Esp" = "ESP"
    }
    foreach ($replacement in $replacements.GetEnumerator()) {
        $name = $name -replace "\b$($replacement.Key)\b", $replacement.Value
    }
    $name
}

foreach ($listName in $listNames) {
    $fields = Get-PnPField -List $listName
    foreach ($field in $fields) {
        if ($field.InternalName -in @("ContentType", "Attachments", "_UIVersionString")) { continue }
        $displayName = Convert-FieldNameToDisplayName $field.InternalName
        Set-PnPField -List $listName -Identity $field.InternalName -Values @{Title = $displayName} -UpdateExistingLists | Out-Null
    }
    Write-Host "Updated labels: $listName"
}