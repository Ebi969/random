$path = "D:\UserData\Ibraaheem\Scripts\VMWare\RevisedMigrationProjects"
$prepPath = "D:\UserData\Ibraaheem\Scripts\VMWare\RevisedMigrationProjects\Prep\VDI"
$actionPlanPath = "$path\ActionPlans\VDI"

$listToWorkWith = Import-Excel -Path $prepPath\TargetList.xlsx -WorksheetName "Target List"

$diskTierID = "RUB|QUA|AMB|EME|DIA"

if(Test-Path $actionPlanPath\*.xlsx){
    Remove-Item $actionPlanPath\*.xlsx
}
$uniqueClusters = $listToWorkWith | Where-Object {$_.Accessible -eq $true} | Select-Object -ExpandProperty Cluster -Unique
foreach($uniqueCluster in $uniqueClusters){
    $newFileName = "$uniqueCluster.xlsx"

    $clusterHosts = Get-Cluster $uniqueCluster | Get-VMHost

    $collect = @()
    foreach($vmHost in $clusterHosts.Name){

        $collect += Get-VMHosthba -VMHost $vmHost -type FibreChannel | Where-Object {$_.Status -eq 'online'} |
        Select-Object  @{N="Host";E={$vmHost}}, @{N="HBA";E={$_.Name}} ,@{N='WWN';E={"{0:x}" -f $_.PortWorldWideName}} | Sort-Object Host
    }

    $collect | Export-Excel -Path "$actionPlanPath\$newFileName" -WorksheetName "Hosts and LUNs" -Append -AutoSize -BoldTopRow -AutoFilter

}

$uniqueDatastores = $listToWorkWith | Where-Object {$_.Accessible -eq $true} | Select-Object -ExpandProperty DatastoreName -Unique

foreach($uniqueDat in $uniqueDatastores){

    if($uniqueDat -match "NML|NMP"){
        $datType = "NML"
    }elseif($uniqueDat -match "MSP"){
        $datType = "MSP"
    }elseif($uniqueDat -match "MAP"){
        $datType = "MAP"
    }elseif($uniqueDat -match "VDI"){
        $datType = "VDI"
    }

    $datVMDetails = Get-Datastore $uniqueDat | Get-VM

        foreach($vmDetails in $datVMDetails){
            $disksToMigrate = @()
            $vmdkNames = @()
            $totalGBToMigrate = $null
            $migrateType = $null

            $vmCluster = $vmDetails | Get-Cluster | Select-Object -ExpandProperty Name
            $vmDatAll = $vmDetails | Get-Datastore
            $countLUNS = ($vmDatAll | Measure-Object -Line).Lines
            $vmDisksAll = $vmDetails | Get-HardDisk 

            $totalGBToMigrate = $vm.MemoryGB
            $totalDisksAttached = ($vmDisksAll | Measure-Object -Line).Lines
            $vmDisks = $vmDisksAll | Where-Object {$_.FileName -match $diskTierID}
                foreach($vmDisk in $vmDisks){
                    $totalGBToMigrate += $vmDisk.CapacityGB
                    $disksToMigrate += $vmDisk.Name
                        $splitNames = ($vmDisk.Filename).Split("/")
                        $finalVMDK = $splitNames[1].Replace(".vmdk", "")
                    $vmdkNames += $finalVMDK
                }
            if($totalGBToMigrate -gt 3400){
                $toBeDedicated = $true
                $migrateType = $datType + "-Indiv"
            }else{
                $toBeDedicated = $false
                $migrateType = $datType + "-Shared"
            }
            $lunNeeded = 0
            if($migrateType -match "Indiv"){
                $lunNeeded = [MATH]::Ceiling($totalGBToMigrate/1024)
                while(($totalGBToMigrate/($lunNeeded*1024)) -gt 0.89){
                    $lunNeeded += 1
                }
            }else{
                $lunNeeded = "Shared"
            }

            $vmOutput = [pscustomobject] @{
                'VMName' = $vmDetails.Name
                'CurrentCluster' = $vmCluster
                'Migrate(GB)' = $totalGBToMigrate
                'Num LUNs' = $countLUNS
                'LUN Names' = ($vmDatAll | Select-Object -ExpandProperty Name) -join ", "
                'Total VMDKS' = $totalDisksAttached
                'Disks to Migrate' = $disksToMigrate -join ", "
                'VMDKNames' = $vmdkNames -join ", "
                'Lun Needed' = $lunNeeded
            }

            $vmOutput | Export-Excel -Path "$actionPlanPath\$vmCluster.xlsx" -WorksheetName $migrateType -Append -BoldTopRow -FreezeTopRow -AutoFilter
        }
}