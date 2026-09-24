[CmdletBinding()]
param(
    [string]$SiteUrl = "https://nttdsdws.sharepoint.com/sites/IntuneReporting",
    [string]$TenantId = "e815be3c-0e2d-4721-9249-1f5dee892931",
    [string]$ClientId = "a71fa672-4cdf-4e00-bcf8-b634d40bac9e",
    [string]$PythonExe = "C:\IntuneMCP\EndpointOps-MCP\.venv\Scripts\python.exe"
)

$ErrorActionPreference = "Stop"
Import-Module PnP.PowerShell
Connect-PnPOnline -Url $SiteUrl -OSLogin -ClientId $ClientId -Tenant $TenantId

$sourceFields = @(
    "deviceId", "azureAdDeviceId", "deviceName", "userPrincipalName", "serialNumber",
    "manufacturer", "model", "operatingSystem", "osVersion", "lastSyncDateTime"
)
$sourceItems = @(Get-PnPListItem -List "Intune Devices" -PageSize 5000 -Fields $sourceFields)
$devices = @($sourceItems | ForEach-Object {
    $fields = $_.FieldValues
    [ordered]@{
        deviceId = [string]$fields.deviceId
        azureAdDeviceId = [string]$fields.azureAdDeviceId
        deviceName = [string]$fields.deviceName
        userPrincipalName = [string]$fields.userPrincipalName
        serialNumber = [string]$fields.serialNumber
        manufacturer = [string]$fields.manufacturer
        model = [string]$fields.model
        operatingSystem = [string]$fields.operatingSystem
        osVersion = [string]$fields.osVersion
        lastSyncDateTime = [string]$fields.lastSyncDateTime
    }
})

if ($devices.Count -eq 0) {
    throw "The Intune Devices list is empty; patch compliance cannot be evaluated."
}

$scriptPath = Join-Path $PSScriptRoot "patch_compliance.py"
$json = $devices | ConvertTo-Json -Depth 4 -Compress
$resultJson = $json | & $PythonExe $scriptPath
if ($LASTEXITCODE -ne 0) {
    throw "Patch baseline evaluation failed with exit code $LASTEXITCODE."
}
$records = @($resultJson | ConvertFrom-Json)
if ($records.Count -eq 0) {
    throw "Patch baseline evaluation returned no Windows device records."
}

$existing = @(Get-PnPListItem -List "Intune Patch Compliance" -PageSize 5000)
foreach ($item in $existing) {
    Remove-PnPListItem -List "Intune Patch Compliance" -Identity $item.Id -Force
}

$runId = [guid]::NewGuid().ToString()
foreach ($record in $records) {
    $values = @{ Key = [string]$record.key; LastRunId = $runId }
    foreach ($property in $record.PSObject.Properties) {
        if ($property.Name -ne "key") {
            $values[$property.Name] = if ($null -eq $property.Value) { "" } else { [string]$property.Value }
        }
    }
    Add-PnPListItem -List "Intune Patch Compliance" -Values $values | Out-Null
}

$counts = $records | Group-Object complianceStatus | Sort-Object Name | ForEach-Object {
    [pscustomobject]@{ Status = $_.Name; Count = $_.Count }
}
Write-Host "Published $($records.Count) patch compliance records from $($devices.Count) device inventory rows."
$counts | Format-Table -AutoSize