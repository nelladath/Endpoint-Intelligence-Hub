[CmdletBinding()]
param(
    [string]$SiteUrl = "https://nttdsdws.sharepoint.com/sites/IntuneReporting",
    [string]$TenantId = "e815be3c-0e2d-4721-9249-1f5dee892931",
    [string]$ClientId = "a71fa672-4cdf-4e00-bcf8-b634d40bac9e",
    [string]$FilePath = "C:\Temp\Telemetry\Endpoint-Intelligence-Hub-SOP.docx"
)

$ErrorActionPreference = "Stop"
Import-Module PnP.PowerShell
Connect-PnPOnline -Url $SiteUrl -OSLogin -ClientId $ClientId -Tenant $TenantId

$library = Get-PnPList -Identity "Documentation"
Add-PnPFile -Path $FilePath -Folder $library.RootFolder.ServerRelativeUrl | Out-Null
Write-Host "Uploaded $(Split-Path $FilePath -Leaf) to Documentation library."
