Import-Module PnP.PowerShell
Connect-PnPOnline -Url $args[0] -OSLogin -ClientId $args[1] -Tenant $args[2]
$web=Get-PnPWeb -Includes Title,Description,WelcomePage
$page=Get-PnPPage -Identity 'Endpoint-Intelligence-Hub.aspx'
$html=$page.ToHtml()
[pscustomobject]@{Title=$web.Title;Description=$web.Description;HomePage=$web.WelcomePage;ContainsDarkBackground=$html.Contains('#0f1720');ContainsPanelCells=$html.Contains("background:#17212b;border:1px solid #334155;color:#f8fafc");ContainsScopeLine=$html.Contains('Current scope');ContainsMethodology=$html.Contains('How The Score Works');ContainsRiskDefinitions=$html.Contains('Risk definitions');Sections=$page.Sections.Count;Controls=$page.Controls.Count}|ConvertTo-Json
