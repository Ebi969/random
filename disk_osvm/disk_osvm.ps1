$vmName = Read-Host "Please input VM name"

$vmHardDisks = Get-VM $vmName | Get-HardDisk

$osHardDisks = Get-CimInstance -ClassName Win32_LogicalDisk -CimSession $vmName | Where {$_.DriveType -eq 3}

foreach($osDisk in $osHardDisks){
    
    $osDiskLetter = $osDisk.DeviceID
    $osDiskFreeSpace = [Math]::Round($osDisk.FreeSpace / 1gb , 2)
    $osDiskSize = [Math]::Round($osDisk.Size / 1gb , 0)

    foreach($vmDisk in $vmHardDisks){
    
        $vmDiskSize = [Math]::Round($vmDisk.CapacityGB , 0)

        if($vmDiskSize -eq $osDiskSize){
            $output = [pscustomobject] @{
                'OS Drive Letter' = $osDiskLetter
                'OS Drive FreeSpace' = $osDiskFreeSpace
                'OS Drive Size' = $osDiskSize
                'VM Hard Disk Number' = $vmDisk.Name
                'VM Disk FileName' = $vmDisk.Filename
                'VM Disk Capacity' = $vmDiskSize
            }
            $output
        }
    }
}