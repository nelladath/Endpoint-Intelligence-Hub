Import-Module PnP.PowerShell
Connect-PnPOnline -Url $args[0] -OSLogin -ClientId $args[1] -Tenant $args[2]
foreach ($listName in @("Intune Apps", "Intune Application Inventory", "Intune Autopilot", "Intune Compliance Policies", "Intune Config Profiles", "Intune Patch Compliance")) {
    $view = Get-PnPView -List $listName -Identity "All Items"
    [pscustomobject]@{ List = $listName; RowLimit = $view.RowLimit; Paged = $view.Paged }
}
