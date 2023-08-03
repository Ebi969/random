$datastores = Get-Cluster SLM-CDC-GOLD-03-LINUX_OLD | Get-Datastore | Where {$_.Name -notmatch "STAGE|local"}
$collect = @()
foreach($ds in $datastores){
$ds
    $clusters = $ds | Get-VMHost | Get-Cluster
    $clusterList = $clusters -join ", "
    $clusterCount = ($clusters | Measure-Object -Line).Lines
    $uid = $ds.ExtensionData.Info.Vmfs.Extent[0].DiskName

    $data = [pscustomobject] @{
        Datastore = $ds.Name
        Uid = $uid
        ClusterCount = $clusterCount
        ClustersMapped = $clusterList
    }

    $collect += $data

}