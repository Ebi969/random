$path = "D:\UserData\Ibraaheem\Scripts\VMWare\All ACI Hosts"

$virtualSwitches = Get-VirtualSwitch | Where {$_.Name -match "ACI"}

foreach($vSwitch in $virtualSwitches){
    foreach($VMhost in $vSwitch | Get-VMHost){
        $output = [pscustomobject] @{
            vmHost = $VMhost.Name
            cluster = $VMhost | Get-Cluster | Select -ExpandProperty Name
            vSwitch = $vSwitch.Name
            Datacenter = $vSwitch.Datacenter
        }
        $output | Export-Excel -Path "$path\ACIHosts.xlsx" -WorksheetName "Hosts" -Append -BoldTopRow -AutoSize -AutoFilter -FreezeTopRow
    }
}


<#
$pGs = Get-VirtualPortGroup | Where {$_.Name -match "EPG"-and $_.ExtensionData.Host -ne $null} | Select -First 1
foreach($pg in $pGs){
    $pg
    foreach($moRefHost in $pg.ExtensionData.Host){
        $VMhost = Get-VMHost -Id $moRefHost
        $output = [pscustomobject] @{
            vmHost = $VMhost.Name
            portGroup = $pg.Name
            portGroupDatacenter = $pg.Datacenter
            virtualSwitch = $pg.VirtualSwitch
        }

        $output | Export-Excel -Path "$path\ACIHosts.xlsx" -WorksheetName "Hosts" -Append -BoldTopRow -AutoSize -AutoFilter -FreezeTopRow

    }
}
#>