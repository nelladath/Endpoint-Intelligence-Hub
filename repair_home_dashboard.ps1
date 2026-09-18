[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)] [string]$SiteUrl,
    [Parameter(Mandatory = $true)] [string]$TenantId,
    [Parameter(Mandatory = $true)] [string]$ClientId
)
$ErrorActionPreference = "Stop"
Import-Module PnP.PowerShell
Connect-PnPOnline -Url $SiteUrl -OSLogin -ClientId $ClientId -Tenant $TenantId
$description = "Centralized portal for Microsoft Intune, Windows Autopilot, AI-assisted driver automation, patch management, application delivery, device compliance, endpoint health, security risk, hardware intelligence, and operational analytics."
Set-PnPWeb -Title "Endpoint Intelligence Hub" -Description $description
$pageName = "Endpoint-Intelligence-Hub.aspx"
$page = Get-PnPPage -Identity $pageName
$page.ClearPage() | Out-Null
$page.Save() | Out-Null
Add-PnPPageSection -Page $page -SectionTemplate OneColumn -ZoneEmphasis 1 | Out-Null
$summaryItems = @(Get-PnPListItem -List "Intune Health Summary" -PageSize 1000 -Fields "metricName", "metricValue", "percentage", "status", "description")
$metrics = @{}
foreach ($item in $summaryItems) { $metrics[[string]$item.FieldValues.metricName] = $item.FieldValues }
function Metric([string]$Name) { if ($metrics.ContainsKey($Name)) { return [string]$metrics[$Name].metricValue }; return "0" }
function Percent([string]$Name) { if ($metrics.ContainsKey($Name)) { return [double]$metrics[$Name].percentage }; return 0 }
function Description([string]$Name) { if ($metrics.ContainsKey($Name)) { return [string]$metrics[$Name].description }; return "" }
function Bar([string]$Name,[string]$Label) { $v=Percent $Name;$w=[math]::Max(0,[math]::Min(100,$v));$c=if($v -ge 95){"#22c55e"}elseif($v -ge 80){"#eab308"}else{"#ef4444"};return "<tr><td style='width:190px;padding:9px 12px;background:#17212b;border:1px solid #334155;color:#f8fafc;font-size:14px;font-weight:600'>$Label</td><td style='padding:9px 12px;background:#17212b;border:1px solid #334155'><div style='background:#1e2a37;height:16px;border-radius:6px'><div style='background:$c;height:16px;border-radius:6px;width:$w%'></div></div><div style='color:#cbd5e1;font-size:11px;margin-top:4px'>$(Description $Name)</div></td><td style='width:70px;padding:9px 12px;text-align:right;background:#17212b;border:1px solid #334155;color:#ffffff;font-size:14px;font-weight:700'>$([math]::Round($v,1))%</td></tr>" }
$risks=@(Get-PnPListItem -List "Intune Device Risks" -PageSize 1000 -Fields "riskLevel");$high=@($risks|Where-Object{$_.FieldValues.riskLevel -eq "High"}).Count;$medium=@($risks|Where-Object{$_.FieldValues.riskLevel -eq "Medium"}).Count
$score=Metric "Overall Endpoint Health Score";$scoreStatus=if($metrics.ContainsKey("Overall Endpoint Health Score")){$metrics["Overall Endpoint Health Score"].status}else{"Unknown"}
$html=@"
<div style='font-family:Segoe UI,Arial;background:#0f1720;color:#f8fafc;padding:24px;border:1px solid #2a3742;border-radius:10px'>
<div style='font-size:18px;font-weight:700;text-transform:uppercase;letter-spacing:.04em;color:#f8fafc;margin:0 0 22px 0'>Endpoint Intelligence Hub | Current scope: $(Metric 'Total Devices') Managed Devices</div>
<table style='width:100%;border-collapse:separate;border-spacing:10px;margin:-10px'><tr><td style='width:25%;vertical-align:top;background:#17212b;border:1px solid #334155;padding:18px;border-radius:10px'><div style='font-size:12px;text-transform:uppercase;color:#cbd5e1'>Overall Endpoint Health</div><div style='font-size:42px;font-weight:700;color:#ffffff;margin:6px 0'>$score<span style='font-size:18px;color:#cbd5e1'>/100</span></div><div style='display:inline-block;background:#166534;color:#dcfce7;padding:5px 14px;border-radius:16px;font-weight:700'>$scoreStatus</div></td><td style='width:25%;vertical-align:top;background:#17212b;border:1px solid #334155;padding:18px;border-radius:10px'><div style='font-size:12px;text-transform:uppercase;color:#cbd5e1'>Risk-flagged devices</div><div style='font-size:30px;font-weight:700;color:#fbbf24;margin:6px 0'>$($high+$medium)</div><div style='color:#e2e8f0;line-height:1.5'>High (score 4+): $high<br>Medium (score 2-3): $medium</div></td><td style='width:25%;vertical-align:top;background:#17212b;border:1px solid #334155;padding:18px;border-radius:10px'><div style='font-size:12px;text-transform:uppercase;color:#cbd5e1'>Critical storage</div><div style='font-size:30px;font-weight:700;color:#fb923c;margin:6px 0'>$(Metric 'Critical Storage (<2%)')</div><div style='color:#e2e8f0'>Devices below 2% free space</div></td><td style='width:25%;vertical-align:top;background:#17212b;border:1px solid #334155;padding:18px;border-radius:10px'><div style='font-size:12px;text-transform:uppercase;color:#cbd5e1'>Autopilot registered</div><div style='font-size:30px;font-weight:700;color:#60a5fa;margin:6px 0'>$(Metric 'Autopilot Registered')</div><div style='color:#e2e8f0'>Current registrations</div></td></tr></table>
<div style='margin:22px 0 18px 0;padding:16px 18px;background:#17212b;border:1px solid #334155;border-radius:8px;color:#cbd5e1;line-height:1.65;font-size:13px'><strong style='color:#ffffff;font-size:15px'>How The Score Works</strong><br><strong style='color:#86efac'>Compliance (25%)</strong>: proportion reporting compliant health.<br><strong style='color:#86efac'>Encryption (20%)</strong>: proportion reporting encryption enabled.<br><strong style='color:#86efac'>Patch/OS currency (20%)</strong>: proportion classified Current or Recent.<br><strong style='color:#86efac'>Threat-free (15%)</strong>: proportion explicitly reporting secured; unknown telemetry is not safe.<br><strong style='color:#86efac'>Device activity (10%)</strong>: proportion synchronized within 30 days.<br><strong style='color:#86efac'>Management health (10%)</strong>: proportion whose state is managed.<br><span style='color:#94a3b8'>Each percentage is healthy devices divided by the current managed-device scope. Weights total 100%.</span></div>
<table style='width:100%;border-collapse:collapse;margin-top:18px'><tr><td style='width:50%;vertical-align:top;padding-right:18px'><h3 style='color:#ffffff;margin:0 0 8px'>Weighted health components</h3><table style='width:100%;border-collapse:collapse'>$(Bar 'Compliance (25%)' 'Compliance (25%)')$(Bar 'Encryption (20%)' 'Encryption (20%)')$(Bar 'Patch/OS currency (20%)' 'Patch/OS currency (20%)')</table></td><td style='width:50%;vertical-align:top;padding-left:18px'><h3 style='color:#ffffff;margin:0 0 8px'>Coverage and risk signals</h3><table style='width:100%;border-collapse:collapse'>$(Bar 'Threat-free (15%)' 'Threat-free (15%)')$(Bar 'Device activity (10%)' 'Device activity (10%)')$(Bar 'Management health (10%)' 'Management health (10%)')</table></td></tr></table>
<div style='margin-top:22px;padding:16px;background:#17212b;border:1px solid #334155;border-radius:8px;color:#cbd5e1;line-height:1.6'><strong style='color:#ffffff'>Risk definitions:</strong> High and Medium are device counts, not users. High means score 4 or higher; Medium means score 2-3. Risk points combine threat severity, jailbreak/root status, encryption, compliance, wipe state, and inactivity.</div>
</div>
"@
Add-PnPPageTextPart -Page $page -Section 1 -Column 1 -Text $html | Out-Null
Set-PnPPage -Identity $pageName -Title "Endpoint Intelligence Hub" -HeaderLayoutType ColorBlock -CommentsEnabled:$false -Publish | Out-Null
Write-Output "Home dashboard repaired: $SiteUrl/SitePages/$pageName"
