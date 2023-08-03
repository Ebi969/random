$vms = Get-VM SRV005027_DR

$vms
#$vm | Get-NetworkAdapter | Select Name, NetworkName
#$vm | Get-VMHost | Select Name

foreach($vm in $vms | Sort-Object Name){
    $vm | Get-HardDisk | Sort-Object Name | Select @{N='VM';E={$_.Parent.Name}}, Name, CapacityGB, FileName, @{N='SCSIid';E={

            $hd = $_
            $ctrl = $hd.Parent.Extensiondata.Config.Hardware.Device | where{$_.Key -eq $hd.ExtensionData.ControllerKey}
            "$($ctrl.BusNumber):$($_.ExtensionData.UnitNumber)"

    }} | Sort-Object SCSIID | Export-Excel -Path "D:\UserData\Ibraaheem\Scripts\VMWare\SCSI ID\SRV005027_DR.xlsx" -Append
}
