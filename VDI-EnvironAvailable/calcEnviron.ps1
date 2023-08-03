$clusters = Get-Cluster

$finalOutput = @()
foreach($cluster in $clusters){
$cluster
$getCluster = $cluster
$vc = $getCluster.Uid.Split(":")[0].Split("@")[1]
$getVMHost = $getCluster | Get-VMHost
$getVM = $getCluster | Get-VM
$numHosts = ($getVMHost | Measure-Object).Count
$hostEffAfterOneFail = ($numHosts - 1) / $numHosts 

#Current Totals for VMs and Host

$vmTotalNumCpu = (($getVM | Measure-Object NumCpu -Sum).Sum)
$vmTotalMem = (($getVM | Measure-Object MemoryGB -Sum).Sum)

$hostTotalNumCpu = ($getVMHost | Measure-Object NumCpu -Sum).Sum
$hostTotalMem = ($getVMHost | Measure-Object MemoryTotalGB -Sum).Sum

#Ratio Calculations

$cpuRatio = $vmTotalNumCpu / $hostTotalNumCpu
$memoryPercent = ($vmTotalMem / $hostTotalMem).ToString("P")

#Output

    $calcOutput = [pscustomobject] @{
        VC = $vc
        Cluster = $cluster
        VMtotCPU = $vmTotalNumCpu
        VMtotMem = $vmTotalMem
        hostotCpu = $hostTotalNumCpu
        hosttotMem = $hostTotalMem
        cpuRatio = $cpuRatio
        AvailvCPU = ($hostTotalNumCpu*4.74)-$vmTotalNumCpu
        memPerc = $memoryPercent
    }

    $finalOutput += $calcOutput

}

$finalOutput | Export-Excel -Path D:\UserData\Ibraaheem\Scripts\VMWare\VDI-EnvironAvailable\calcEnviron.xlsx -AutoSize -AutoFilter -BoldTopRow
