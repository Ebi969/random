$path = "D:\UserData\Ibraaheem\Scripts\VMWare\Ratio"
$clusters = @("SKY-BDC-GOLD-03-DATABASE")

foreach($cluster in $clusters){
$cluster
$getCluster = Get-Cluster $cluster
$getVMHost = $getCluster | Get-VMHost
$getVM = $getCluster | Get-VM

$cpuRatio = (($getVM | Measure-Object NumCpu -Sum).Sum) / ($getVMHost | Measure-Object NumCpu -Sum).Sum
$memoryPercent = ((($getVM | Measure-Object MemoryGB -Sum).Sum) / ($getVMHost | Measure-Object MemoryTotalGB -Sum).Sum).ToString("P")

    $output = [pscustomobject] @{
        Cluster = $cluster
        cpuRatio = $cpuRatio
        memPerc = $memoryPercent
    } | Export-Excel -Path $path\vmSpecs.xlsx -WorksheetName "Ratio" -Append -BoldTopRow -AutoSize -AutoFilter

    $split = $cluster.Split("-")

$getVM | Select Name, PowerState, NumCpu, MemoryGB | Export-Excel -Path $path\vmSpecs.xlsx -WorksheetName $split[1] -Append -BoldTopRow -AutoSize -AutoFilter

}

#mem 288
#mem 18