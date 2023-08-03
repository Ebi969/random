$clusters = Read-Host "Cluster to be added to"

$addCPU = Read-Host "Number of CPU to be added to each Cluster"
$addMem = Read-Host "Memory in GB to be added to each Cluster"

$finalOutput = @()

foreach($cluster in $clusters){
$cluster
$getCluster = Get-Cluster $cluster
$getVMHost = $getCluster | Get-VMHost
$getVM = $getCluster | Get-VM
$numHosts = ($getVMHost | Measure-Object).Count
$hostEffAfterOneFail = ($numHosts - 1) / $numHosts 

#Current Totals for VMs and Host

$vmTotalNumCpu = (($getVM | Measure-Object NumCpu -Sum).Sum)
$vmTotalMem = (($getVM | Measure-Object MemoryGB -Sum).Sum)

$hostTotalNumCpu = ($getVMHost | Measure-Object NumCpu -Sum).Sum
$hostTotalMem = ($getVMHost | Measure-Object MemoryTotalGB -Sum).Sum

#Add Resources / Host Failure Calculations

$vmTotalNumCpuAfterAdd = $vmTotalNumCpu + $addCPU
$vmTotalMemAfterAdd = $vmTotalMem + $addMem

$hostTotalNumCpuAfterOneFail = $hostTotalNumCpu * $hostEffAfterOneFail
$hostTotalMemAfterOneFail = $hostTotalMem * $hostEffAfterOneFail

#Ratio Calculations

$cpuRatio = $vmTotalNumCpu / $hostTotalNumCpu
$memoryPercent = ($vmTotalMem / $hostTotalMem).ToString("P")

$cpuRatioAfterAdd = $vmTotalNumCpuAfterAdd / $hostTotalNumCpu
$memoryPercentAfterAdd = ($vmTotalMemAfterAdd / $hostTotalMem).ToString("P")

$cpuRatioAfterOneFail = $vmTotalNumCpuAfterAdd / $hostTotalNumCpuAfterOneFail
$memoryPercentAfterOneFail = ($vmTotalMemAfterAdd / $hostTotalMemAfterOneFail).ToString("P")

#Output

    $calcOutput = [pscustomobject] @{
        Cluster = $cluster
        "Current CPU Ratio" = $cpuRatio    
        "Current Mem %" = $memoryPercent
        "After Addition CPU Ratio" = $cpuRatioAfterAdd
        "After Addition Mem %" = $memoryPercentAfterAdd
        "New CPU Ratio after 1 host failure" = $cpuRatioAfterOneFail
        "New Mem % after 1 host failure" = $memoryPercentAfterOneFail
    }

    $finalOutput += $calcOutput

}

$finalOutput | Out-GridView

#mem add 288
#cpu add 18
