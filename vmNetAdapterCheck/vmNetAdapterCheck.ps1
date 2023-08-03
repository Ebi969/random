foreach($vm in Get-VM){
    $netAdapterInfo = $vm | Get-NetworkAdapter

    $vmCluster = $vm | Get-Cluster | Select -ExpandProperty Name

    foreach($adapter in $netAdapterInfo){
        if($adapter.Type -notlike "*vmxnet3*"){
            $outputInfo = [pscustomobject] @{
                Cluster = $vmCluster
                VM = $vm.Name
                OS = $vm.ExtensionData.Guest.GuestFullName
                PowerState = $vm.PowerState
                AdapterName = $adapter.Name
                Type = $adapter.Type
                vLan = $adapter.NetworkName
                MacAddress = $adapter.MacAddress
                'WakeOnLan Enabled' = $adapter.WakeOnLanEnabled
            }
            $outputInfo
            $outputInfo | Export-Excel -Path "D:\UserData\Ibraaheem\Scripts\VMWare\vmNetAdapterCheck\E1000E_vmNetAdapter.xlsx" -WorksheetName "E1000E" -Append -AutoSize -AutoFilter -BoldTopRow
        }
    }
}