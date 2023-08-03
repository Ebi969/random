foreach($vmHost in Get-VMHost){
    $VmsOnHost = $vmHost | Get-VM
    if($VmsOnHost.count -eq "0"){
        $outExport = [pscustomobject] @{
            #Cluster = $vmHost | Get-Cluster | Select -ExpandProperty Name
            Parent = $vmHost.Parent
            'Host Name' = $vmHost.Name
            'Connection State' = $vmHost.ConnectionState
            'Model' = $vmHost.Model
            'ESXi Version' = $vmHost.Version
        } | Export-Excel -Path "D:\UserData\Ibraaheem\Scripts\VMWare\Empty VMHosts\EmptyHosts.xlsx" -Append -BoldTopRow -FreezeTopRow -AutoSize -AutoFilter
    }
}