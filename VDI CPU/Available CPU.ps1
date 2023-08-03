$clusters = Get-Cluster

$finalOutput = @()
foreach($cluster in $clusters){
$getCluster = $cluster
$vc = $getCluster.Uid.Split(":")[0].Split("@")[1]
$getVMHost = $getCluster | Get-VMHost
$getVM = $getCluster | Get-VM

#Current Totals for VMs and Host

$vmTotalNumCpu = (($getVM | Measure-Object NumCpu -Sum).Sum)
$hostTotalNumCpu = ($getVMHost | Measure-Object NumCpu -Sum).Sum

#Ratio Calculations

$cpuRatio = $vmTotalNumCpu / $hostTotalNumCpu
$availableCPU = ($hostTotalNumCpu * 6) - $vmTotalNumCpu
#Output

    $calcOutput = [pscustomobject] @{
        Cluster = $cluster
        vCenter = $vc
        "TotalCPUinCluster" = $hostTotalNumCpu
        "CPUAllocation" = $vmTotalNumCpu
        "CPU Ratio" = $cpuRatio
        "AvailableCPU" = $availableCPU
    }

    $finalOutput += $calcOutput

}

$finalOutput | Out-GridView