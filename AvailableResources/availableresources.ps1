$outputPath = "D:\UserData\Ibraaheem\Scripts\VMWare\AvailableResources"
$fileName = "AvailableResources.xlsx"
if(Test-Path $outputPath\$fileName){
    Remove-Item $outputPath\$fileName -Force
}


$clusters = Get-Cluster | Where {$_.Name -match "-06-SQL"}

foreach($cluster in $clusters){

$vmHosts = $cluster | Get-VMHost

if(!($vmHosts)){
continue 
}

$vms = $cluster | Get-VM

$totalCPU = ($vmHosts | Measure-Object NumCpu -Sum).Sum
$totalMemGB = [MATH]::Round(($vmHosts | Measure-Object MemoryTotalGB -Sum).Sum ,2)

$vmCPUs = ($vms | Measure-Object NumCpu -Sum).Sum
$vmMem = ($vms | Measure-Object MemoryGB -Sum).Sum

$cpuRatio = [MATH]::Round(($vmCPUs/$totalCPU), 2)
$memPerc = [MATH]::Round(($vmMem/$totalMemGB)*100, 2)

$clusView = $cluster | Get-View

$vcURL = $clusView.Client.ServiceUrl
$split = $vcURL.Split(".")
$vcName = ($split[0].Split("/"))[2]

$toBeLimitMem = (0.8 * $totalMemGB) - $vmMem
$toBeLimitCpu = ($totalCPU*2) - $vmCPUs

    $clusterOutput = [pscustomobject] @{
        Cluster = $cluster.Name
        vCenter = $vcName
        TotalCPUinCluster = $totalCPU
        CPUallocation = $vmCPUs
        CPURatio = $cpuRatio
        AvailableCpu = $toBeLimitCpu
        TotalMEMinCluster = $totalMemGB
        MEMallocation = [MATH]::Round($vmMem, 2)
        MemPercAllocated = $memPerc
        AvailableMem = [MATH]::Round($toBeLimitMem, 2)
    }

foreach($vmHost in $vmHosts){

    $totalHostCPU = ($vmHost | Measure-Object NumCpu -Sum).Sum
    $totalHostMemGB = [MATH]::Round(($vmHost | Measure-Object MemoryTotalGB -Sum).Sum ,2)

    $hostVMs = $vmHost | Get-VM
    $hostVMCPUs = ($hostVMs | Measure-Object NumCpu -Sum).Sum
    $hostVMMem = ($hostVMs | Measure-Object MemoryGB -Sum).Sum

    $hostcpuRatio = [MATH]::Round(($hostVMCPUs/$totalHostCPU), 2)
    $hostmemPerc = [MATH]::Round(($hostVMMem/$totalHostMemGB)*100, 2)
    
    $hosttoBeLimitCpu = ($totalHostCPU*2) - $hostVMCPUs
    $hosttoBeLimitMem = (0.8 * $totalHostMemGB) - $hostVMMem

    $hostOutput = [pscustomobject] @{
        Cluster = $cluster.Name
        vCenter = $vcName
        host = $vmHost.Name
        TotalCPUinCluster = $totalHostCPU
        CPUallocation = $hostVMCPUs
        CPURatio = $hostcpuRatio
        AvailableCpu = $hosttoBeLimitCpu
        TotalMEMinCluster = $totalHostMemGB
        MEMallocation = [MATH]::Round($hostVMMem, 2)
        MemPercAllocated = $hostmemPerc
        AvailableMem = [MATH]::Round($hosttoBeLimitMem, 2)
    }

    $hostOutput | Export-Excel -Path $outputPath\$fileName -WorksheetName Host -Append -BoldTopRow -AutoFilter -FreezeTopRow -AutoSize

    foreach($vm in $hostVMs){

        $vmView = $vm | Get-View
        $cps = $vmView.config.hardware.NumCoresPerSocket
        $sockets = $vmView.config.hardware.NumCPU / $cps
        $vmOutput = [pscustomobject] @{
            Host = $vmHost.Name
            VM = $vm.Name
            vCPU = $vm.NumCpu
            Sockets = $sockets 
            CPS = $cps
            Memory = $vm.MemoryGB
        }
        $vmOutput | Export-Excel -Path $outputPath\$fileName -WorksheetName VMs -Append -BoldTopRow -AutoFilter -FreezeTopRow -AutoSize
    }

}

$clusterOutput | Export-Excel -Path $outputPath\$fileName -WorksheetName Cluster -Append -BoldTopRow -AutoFilter -FreezeTopRow -AutoSize

}