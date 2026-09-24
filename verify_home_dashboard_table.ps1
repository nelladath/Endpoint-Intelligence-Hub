Import-Module PnP.PowerShell
Connect-PnPOnline -Url $args[0] -OSLogin -ClientId $args[1] -Tenant $args[2]
$page=Get-PnPPage -Identity 'Endpoint-Intelligence-Hub.aspx'
[pscustomobject]@{Sections=$page.Sections.Count;Controls=$page.Controls.Count;ControlTitles=(@($page.Controls|ForEach-Object{$_.Title})-join ',')}|ConvertTo-Json
