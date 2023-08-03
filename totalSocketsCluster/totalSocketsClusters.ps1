$path = "D:\UserData\Ibraaheem\Scripts\VMWare\totalSocketsCluster"
$vmClusters = Get-Content $path\ClusterList.txt
$totSockets = 0
foreach($cluster in $vmClusters){
    foreach($vmHost in (Get-Cluster $cluster | Get-VMHost)){        
        $sockets = $vmHost.ExtensionData.Hardware.CpuInfo.NumCpuPackages
        $output = [pscustomobject] @{
            Cluster = ($vmHost.Parent).Name
            vmHost = $vmHost.Name
            sockets = $sockets
        } | Export-Excel -Path $path\TotalSockets.xlsx -WorksheetName "Socket Count" -Append -BoldTopRow -AutoSize -AutoFilter
        $totSockets += $sockets
    }
}

$output = [pscustomobject] @{
    Cluster = "Total Sockets"
    vmHost = ""
    sockets = $totSockets
} | Export-Excel -Path $path\TotalSockets.xlsx -WorksheetName "Socket Count" -Append -BoldTopRow -AutoSize -AutoFilter