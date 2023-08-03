$inputPath = "D:\UserData\Ibraaheem\Scripts\VMWare\VM_Datastore"
$outputPath = "D:\UserData\Ibraaheem\Scripts\VMWare\VM_Datastore"

$vmList = Get-Content $inputPath\serverList.txt

if(Test-Path $outputPathdsvms15.xlsx{
    Remove-Item $outputPath\dsvms15.xlsx
}

$finalOut = @()

foreach($vm in $vmList){

$vmData = Get-VM $vm
$clusterData = $vmData | Get-Cluster
$dataStoreData = $vmData | Get-Datastore

    foreach($dataStore in $dataStoreData){
        $percFifteen = [Math]::Round($dataStore.CapacityGB * 0.15, 2)

        if($percFifteen -gt $dataStore.FreeSpaceGB){
            $neededCap = $percFifteen - $dataStore.FreeSpaceGB
            $neededCap = [Math]::Round($neededCap,2)
        }else{
            $neededCap = "Not Needed"
        }
        
        $out = [pscustomobject] @{
            "VM Name" = $vmData.Name
            "VM Size" = ([Math]::Round($vmData.UsedSpaceGB,2))
            "VM Cluster" = $clusterData.Name
            "VM Datastore" = $dataStore.Name
            "Datastore Max Capacity" = ([Math]::Round($dataStore.CapacityGB,2))
            "Datastore Free Space" = ([Math]::Round($dataStore.FreeSpaceGB,2))
            "15% of Datastore Capacity" = $percFifteen
            "Capacity needed to be 15%" = $neededCap
        }
        $finalOut += $out
        
    }
}

$finalOut | Export-Excel -Path $outputPath\dsvms15.xlsx -AutoSize -AutoFilter -BoldTopRow