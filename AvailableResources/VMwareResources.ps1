$outputPath = "D:\Reports\AutoDeploy"
$clusterMatrix = Import-Excel "D:\Scripts\VMware\AvailableResources\ClusterMatrix.xlsx"
$fileName = "AvailableResources.xlsx"
if(Test-Path $outputPath\$fileName){
    Remove-Item $outputPath\$fileName -Force
}

$VMwareVC = @('SRV007281','SRV007282','SRV008097')

# Get Cred store
$creds7281 = Get-VICredentialStoreItem -file  C:\Users\svcVMWareScriptAcc\CredStore7281.creds
$creds7282 = Get-VICredentialStoreItem -file  C:\Users\svcVMWareScriptAcc\CredStore7282.creds
$creds8097 = Get-VICredentialStoreItem -file  C:\Users\svcVMWareScriptAcc\CredStore8097.creds

# Connect-VIServer
Connect-VIServer -Server SRV007281.mud.internal.co.za -User $Creds7281.User -Password $Creds7281.Password
Connect-VIServer -Server SRV007282.mud.internal.co.za -User $Creds7282.User -Password $Creds7282.Password
Connect-VIServer -Server SRV008097.mud.internal.co.za -User $creds8097.User -Password $creds8097.Password

$clusters = Get-Cluster <#| Where {($_| Get-Datacenter) -notmatch "PHYSICAL|CCC|REGIONAL"}#> | Sort-Object Name

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

$matrixCpu = $null
$matrixMem = $null
$matrixCpu = $clusterMatrix | Where {$_.Cluster -eq $cluster.Name} | Select -ExpandProperty vCPURatioLimit
$matrixMem = $clusterMatrix | Where {$_.Cluster -eq $cluster.Name} | Select -ExpandProperty MemPercMax

if($matrixCpu -eq $null){
$matrixCpu = 4
}

if($matrixMem -eq $null){
$matrixMem = 80
}

$cluster.Name
$toBeLimitMem = (($matrixMem/100) * $totalMemGB) - $vmMem
$toBeLimitCpu = ($totalCPU*$matrixCpu) - $vmCPUs

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

$clusterOutput | Export-Excel -Path $outputPath\$fileName -WorksheetName Cluster -Append -BoldTopRow -AutoFilter -FreezeTopRow -AutoSize
        $datastores = $cluster | Get-Datastore | Where {$_.Name -notmatch "datastore|pamco|stage|local|scratch|NFS|DNU|SQLDEV|srv|-0|-1|-3|SRM|DRP|DRT|MAP|PDF|ARC|STR"} | Select @{n="ClusterName"; e={$cluster.Name}}, @{n="vCenter"; e={$vcName}}, Name, @{n="FreeSpaceGB"; e={[MATH]::Round($_.FreeSpaceGB,2)}}, @{n="CapacityGB"; e={[MATH]::Round($_.CapacityGB,2)}}, @{n="AvailableStorage"; e={[MATH]::Round(($_.FreeSpaceGB - ($_.CapacityGB * 0.1)),2)}}

$datastores | Export-Excel -Path $outputPath\$fileName -WorksheetName Storage -Append -BoldTopRow -AutoFilter -FreezeTopRow -AutoSize

}
