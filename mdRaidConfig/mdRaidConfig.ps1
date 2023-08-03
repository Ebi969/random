$path = "D:\UserData\Ibraaheem\Scripts\VMWare\mdRaidConfig"
$vmList = Get-Content "$path\list.txt"

foreach($vmName in $vmList){
    $drName = $vmName + "_DR"
    
    $vm = Get-VM $vmName
    $vmNic = $vm | Get-NetworkAdapter | Select Name, NetworkName
    $vmHost = $vm | Get-VMHost | Select Name
    $vmCluster = $vm | Get-Cluster | Select Name
    $vmDisks = $vm | Get-HardDisk | Select @{N='VM';E={$_.Parent.Name}}, Name, CapacityGB, FileName, @{N='SCSIid';E={
            $hd = $_
            $ctrl = $hd.Parent.Extensiondata.Config.Hardware.Device | where{$_.Key -eq $hd.ExtensionData.ControllerKey}
            "$($ctrl.BusNumber):$($_.ExtensionData.UnitNumber)"
            }} | Sort-Object SCSIID

    $drvm = Get-VM $drName
    $drvmNic = $drvm | Get-NetworkAdapter | Select Name, NetworkName
    $drvmHost = $drvm | Get-VMHost | Select Name
    $drvmCluster = $drvm | Get-Cluster | Select Name
    $drvmDisks = $drvm | Get-HardDisk | Select @{N='VM';E={$_.Parent.Name}}, Name, CapacityGB, FileName, @{N='SCSIid';E={
            $hd = $_
            $ctrl = $hd.Parent.Extensiondata.Config.Hardware.Device | where{$_.Key -eq $hd.ExtensionData.ControllerKey}
            "$($ctrl.BusNumber):$($_.ExtensionData.UnitNumber)"
            }} | Sort-Object SCSIID

    $vmObject = [pscustomObject] @{
        VM = $vm.Name
        VMHost = $vmHost.Name
        VMCluster = $vmCluster.Name
        VMnicName = $vmNic.Name
        VMnicVLAN = $vmNic.NetworkName
    }
    $drvmObject = [pscustomObject] @{
        VM = $drvm.Name
        VMHost = $drvmHost.Name
        VMCluster = $drvmCluster.Name
        VMnicName = $drvmNic.Name
        VMnicVLAN = $drvmNic.NetworkName
    }
    
    $vmObject | Export-Excel -Path "$path\MDRAID_Configs.xlsx" -WorksheetName "VMDetails" -Append -BoldTopRow -AutoSize -FreezeTopRow -AutoFilter
    $drvmObject | Export-Excel -Path "$path\MDRAID_Configs.xlsx" -WorksheetName "VMDetails" -Append -BoldTopRow -AutoSize -FreezeTopRow -AutoFilter
    
    $vmDisks | Export-Excel -Path "$path\MDRAID_Configs.xlsx" -WorksheetName "$vmName" -Append -BoldTopRow -AutoSize -FreezeTopRow -AutoFilter
    $drvmDisks | Export-Excel -Path "$path\MDRAID_Configs.xlsx" -WorksheetName "$vmName" -Append -BoldTopRow -AutoSize -FreezeTopRow -AutoFilter
}