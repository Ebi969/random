$inputPath = "D:\UserData\Ibraaheem\Scripts\VMWare\VM_Datastore"
$outputPath = "D:\UserData\Ibraaheem\Scripts\VMWare\VM_Datastore"

$vmList = Get-Content $inputPath\serverList.txt

if(Test-Path $outputPath\dsvms.xlsx){
    Remove-Item $outputPath\dsvms.xlsx
}

$finalOut = @()

foreach($vm in $vmList){

$vmData = Get-VM $vm
$clusterData = $vmData | Get-Cluster
$dataStoreData = ($vmData | Get-Datastore | select -ExpandProperty Name) -join ", "

        $out = [pscustomobject] @{
            "VM Name" = $vmData.Name
            "VM Size" = ([Math]::Round($vmData.UsedSpaceGB,2))
            "VM Cluster" = $clusterData.Name
            "VM Datastore" = $dataStoreData
        }
        $finalOut += $out
}

$finalOut | Export-Excel -Path $outputPath\dsvms.xlsx -AutoSize -AutoFilter -BoldTopRow