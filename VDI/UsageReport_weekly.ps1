Measure-Command{
$outputPath = "D:\UserData\Ibraaheem\Scripts\VMWare\VDI"

$timeStamp = Get-Date -Format 'dd-MMM-yyyy'

$fileName = "WeeklyVDIUsageCheck_" + $timeStamp + ".xlsx"

$exportName = $outputPath + "\" + $fileName

$vms = Get-View -ViewType VirtualMachine  

$allVMOutput = @()
foreach($vm in $vms){
    $hostName = Get-view -Id $vm.runtime.host
    $cluster = Get-view -Id $hostName.Parent
    $datacenter = Get-Datacenter -Cluster $cluster.Name
    $fullURL = $vm.Client.ServiceUrl
    $splitWhack = $fullURL.split("/")
    $vc = ($splitWhack[2].Split("."))[0]

    $rowOutput = [pscustomobject] @{    
        VM = $vm.Name
        PowerState = $vm.Summary.Runtime.PowerState
        Template = $vm.Config.Template
        CPUs = $vm.Config.Hardware.NumCPU        
        "Memory MB" = $vm.Config.Hardware.MemoryMB
        "Capacity MB" = [MATH]::Round(($vm.Storage.PerDatastoreUsage.Uncommitted + $vm.Storage.PerDatastoreUsage.committed)/1024/1024/1024 , 2)
        Datacenter = $dataCenter.Name
        Cluster = $cluster.Name
        Host = $hostName.Name
        "OS according to the configuration file" = $vm.Config.GuestFullName
        "vCenter" = $vc
    }
    $allVMOutput += $rowOutput
}

$allVMOutput | Where vCenter -eq "SRV006270" | Sort-Object Host | Export-Excel -Path $exportName -WorkSheetname "OldEnvironment" -ClearSheet `
    -IncludePivotTable -PivotTableName "Summary_OldEnvironment" -PivotRows @("Datacenter", "Host") -PivotDataToColumn -PivotData @{"VM"="Count"; "CPUs"="Sum"; "Memory MB"="Sum"; "Capacity MB"="Sum"} `
    -PivotFilter "PowerState" -AutoSize -BoldTopRow -FreezeTopRow -AutoFilter

$allVMOutput | Where vCenter -eq "SRV008146" | Sort-Object Host | Export-Excel -Path $exportName -WorkSheetname "NewEnvironment" -ClearSheet `
    -IncludePivotTable -PivotTableName "Summary_NewEnvironment" -PivotRows @("Datacenter", "Host") -PivotDataToColumn -PivotData @{"VM"="Count"; "CPUs"="Sum"; "Memory MB"="Sum"; "Capacity MB"="Sum"} `
    -PivotFilter "PowerState" -AutoSize -BoldTopRow -FreezeTopRow -AutoFilter
    #-IncludePivotTable -PivotTableName Summary -PivotRows @("Datacenter", "Host") -PivotData @{"CPUs"="Sum"; "Size MB"="Sum"; "Capacity"="Sum"; "VM"="Count" } -PivotDataToColumn `
}