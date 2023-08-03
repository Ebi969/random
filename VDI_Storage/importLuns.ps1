$fileToRead = "D:\USERDATA\Ibraaheem\Scripts\VMWare\RevisedMigrationProjects\ActionPlans\VDI\CDC VDI Replacement.xlsx"
$importFile = Import-Excel -Path $fileToRead -WorksheetName "LunList"

$allLunInfo = @()
foreach($cluster in $importFile.Cluster | Select -Unique){
    $cluster

    $vmhost = Get-Cluster $cluster | Get-VMHost | Select-Object -First 1
    $unmappedLUNs = Get-SCSILun -VMhost $vmhost -LunType Disk | Select-Object CanonicalName

    foreach($newLun in $importFile | Where-Object {$_.Cluster -match $cluster}){
        $naa = ""
        $naa = $unmappedLUNs | Where-Object {$_.CanonicalName -eq "naa."+$($newLun.UID)}
        $lunIndiv = [pscustomobject]@{
            Cluster = $newLun.Cluster
            vmHost = $vmhost.Name
            DatastoreName = $newLun.Name
            UID = $naa.CanonicalName
        }
        $allLunInfo += $lunIndiv
    }

}

$totalLunsInImport = $importFile.Count
$totalLunsWithUID = ($allLunInfo | Where-Object {$_.UID}).count
Write-Host "$totalLunsWithUID of $totalLunsInImport found!"

$failed = @()
foreach($cluster in $importFile.Cluster | Select -Unique){

    Write-Host "Starting $cluster"
    $amountToImport = $allLunInfo | Where-Object {$_.Cluster -eq $cluster -and $_.UID}.count

    $completed = 0

    foreach($toImport in $allLunInfo | Where-Object {$_.Cluster -eq $cluster -and $_.UID}){
    Write-Host "$($toImport.DatastoreName) - $($toImport.UID)"
    
        Try{
            $vmhost = Get-Cluster $cluster | Get-VMHost | Select-Object -First 1
            New-Datastore -VMHost $vmhost -Name $toImport.DatastoreName -Path $toImport.UID -Vmfs -BlockSizeMB 1MB -FileSystemVersion 6

            $Datastore = Get-Datastore $toImport.DatastoreName
            $vmhostView = ($Datastore | Get-VMHost).ExtensionData
            $storageSystem = Get-View $vmhostView.ConfigManager.StorageSystem

            $enableUNMAP = "none"
            $reconfigMessage = "Disabling Automatic VMFS Unmap for $Datastore"

            $uuid = $datastore.ExtensionData.Info.Vmfs.Uuid

            Write-Host "$reconfigMessage ..."    
            $storageSystem.UpdateVmfsUnmapPriority($uuid,$enableUNMAP)

            $completed += 1
            Write-Host "$($toimport.DatastoreName) complete, going next"
        }catch{
            $failed += $toImport
            Write-Host "$($toimport.DatastoreName) failed, going next"            

        }
    
    }
    
    Get-Cluster -name $cluster | Get-VMhost | Get-VMHostStorage -RescanAllHBA
    Write-Host "$cluster complete"

}

#Get-SCSILun -VMhost 192.168.1.103 -LunType Disk | Select CanonicalName,Capacity
#New-Datastore -VMHost Host -Name Datastore -Path CanonicalName -VMFS
#Get-Cluster -name Cluster | Get-VMhost | Get-VMHostStorage -RescanAllHBA