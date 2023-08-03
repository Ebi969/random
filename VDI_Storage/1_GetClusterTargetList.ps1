$path = "D:\UserData\Ibraaheem\Scripts\VMWare\RevisedMigrationProjects"
$prepPath = "$path\Prep\VDI"

$UIdentifiers = "60050763808104BB00000|600507681081018E90000"
#$ownStorageClients = "SKY|CCC"
#$tierIdentifiers = "DIA|SAP"
#$identifiers = "VDI"

$masterList = Get-Datastore | Where-Object {$_.extensiondata.info.vmfs.Extent[0].DiskName -match $UIdentifiers}
#$masterList = Get-Datastore | Where-Object {$_.Name -match $identifiers}

foreach($datInformation in $masterList){
    $datClusters = $null
    $datNumClustersAssigned = $null
    $usedSpaceGB = $null

    $usedSpaceGB = $datInformation.CapacityGB - [Math]::Round($datInformation.FreeSpaceGB, 2)
    $datClusters = $datInformation | Get-VMHost | Get-Cluster
    $datNumClustersAssigned = ($datClusters | Measure-Object -Line).Lines

    foreach($datCluster in $datClusters){
    $indivWorkWith = $null
        $indivWorkWith = [pscustomobject] @{
            DatastoreName = $datInformation.Name
            CapacityGB = $datInformation.CapacityGB
            FreeSpaceGB = [Math]::Round($datInformation.FreeSpaceGB, 2)
            UsedSpaceGB = [Math]::Round($usedSpaceGB, 2)
            VMFS = [math]::Floor($datInformation.FileSystemVersion)
            Accessible = $datInformation.Accessible
            ClustersAssigned = $datNumClustersAssigned
            Cluster = $datCluster
            vCenter = $datCluster.uid.split("@")[1].split(":")[0]
        }
        $indivWorkWith
        if($indivWorkWith.Accessible -eq $false){
            $indivWorkWith | Export-Excel -Path $prepPath\TargetList.xlsx -WorksheetName "Target List" -Append -AutoSize -BoldTopRow -AutoFilter
            $indivWorkWith | Export-Excel -Path $prepPath\TargetList.xlsx -WorksheetName "To be Reclaimed" -Append -AutoSize -BoldTopRow -AutoFilter
        }else{
            $indivWorkWith | Export-Excel -Path $prepPath\TargetList.xlsx -WorksheetName "Target List" -Append -AutoSize -BoldTopRow -AutoFilter
        }
    }
}