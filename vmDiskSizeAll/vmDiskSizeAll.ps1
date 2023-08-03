$allVMList = Get-VM

Function msFunctionDisk{
param(
    $server,
    $poweredState
)

    try{
        $listDisks =  Get-CimInstance -ClassName Win32_LogicalDisk -ComputerName $server -ErrorAction Stop | Where {$_.DriveType -eq "3"}
            foreach($disk in $listDisks){
                $diskName = $disk.DeviceID
                $diskCapacityGB = [Math]::Round(($disk.Size/1gb),2)
                $diskFreeSpaceGB = [Math]::Round(($disk.FreeSpace/1gb),2)

                $outObject = [pscustomobject]@{
                    VMName = $server
                    'Disk Name' = $diskName
                    'CapacityGB' = $diskCapacityGB
                    'FreeSpaceGB' = $diskFreeSpaceGB                    
                } | Export-Excel -Path "D:\UserData\Ibraaheem\Scripts\VMWare\vmDiskSizeAll\vmDiskSizes.xlsx" -WorksheetName "Windows VMs" -Append -BoldTopRow -AutoSize -AutoFilter
            }
    }catch{

        if($poweredState -eq "PoweredOff"){
            $FailureReason = "VM is powered off"
        }else{
            $FailureReason = "Unable to access OS - Check Domain/Network"
        }

        $outObject = [pscustomobject]@{            
            Inaccessible = $server
            PoweredState = $poweredState
            'Failure Reason' = $FailureReason                   
        } | Export-Excel -Path "D:\UserData\Ibraaheem\Scripts\VMWare\vmDiskSizeAll\vmDiskSizes.xlsx" -WorksheetName "Inaccessible Windows VMs" -Append -BoldTopRow -AutoSize -AutoFilter
    }
}


foreach($vm in $allVMList){

    $vmName = $vm.Name
    $os = $vm.extensiondata.guest.guestfullname
    $powerstate = $vm.PowerState

    $vmName

    if($os -like "*Microsoft*"){
        msFunctionDisk -server $vmName -poweredState $powerstate
    }else{
        $listDisks = $vm | Get-HardDisk
            foreach($disk in $listDisks){
                $diskName = $disk.Name
                $diskCapacityGB = $disk.CapacityGB

                $outObject = [pscustomobject]@{
                    VMName = $vmName
                    OS = $os
                    PoweredState = $powerstate
                    'Disk Name' = $diskName
                    'CapacityGB' = $diskCapacityGB                 
                } | Export-Excel -Path "D:\UserData\Ibraaheem\Scripts\VMWare\vmDiskSizeAll\vmDiskSizes.xlsx" -WorksheetName "Non-Windows VMs" -Append -BoldTopRow -AutoSize -AutoFilter
            }
    }
}