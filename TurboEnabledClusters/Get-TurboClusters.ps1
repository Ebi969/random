$path = "D:\UserData\Ibraaheem\Scripts\VMWare\TurboEnabledClusters"
$clusters = Get-Cluster

foreach($cluster in $clusters){
    $perms = $null
    $perms = $cluster | Get-VIPermission | Where {$_.Principal -match "turbo"}

    if($perms -ne $null){
        $rowOutput = [pscustomobject] @{
            Cluster = $cluster.Name
            Principal = $perms.Principal
            Role = $perms.Role
        }
        $rowOutput | Export-Excel -Path "$path\TurboClusterDetails.xlsx" -WorksheetName "Details" -AutoSize -AutoFilter -BoldTopRow -Append -FreezeTopRow
    }else{
        $rowOutput = [pscustomobject] @{
            Cluster = $cluster.Name
            Message = "Turbo account not present"
        }
        $rowOutput | Export-Excel -Path "$path\TurboClusterDetails.xlsx" -WorksheetName "NotPresent" -AutoSize -AutoFilter -BoldTopRow -Append -FreezeTopRow
    }
}