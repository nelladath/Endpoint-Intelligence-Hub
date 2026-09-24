Import-Module PnP.PowerShell
Connect-PnPOnline -Url $args[0] -OSLogin -ClientId $args[1] -Tenant $args[2]
foreach ($listName in @("Intune Apps", "Intune Application Inventory", "Intune Autopilot", "Intune Devices", "Intune Device Risks")) {
    $view = Get-PnPView -List $listName -Identity "All Items"
    [pscustomobject]@{ List = $listName; RowLimit = $view.RowLimit; Paged = $view.Paged; Fields = ($view.ViewFields | Select-Object -ExpandProperty Name) -join "," }
}
