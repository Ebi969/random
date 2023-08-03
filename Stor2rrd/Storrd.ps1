$stor2rrdPath = "D:\UserData\Ibraaheem\Scripts\VMWare\Stor2rrd"
if(Test-Path $stor2rrdPath\Storrd.csv){
    Remove-Item $stor2rrdPath\Storrd.csv
}

Invoke-WebRequest -Uri http://srv005326.mud.internal.co.za/stor2rrd_reports/LATEST/Report_volume_storage.csv -OutFile "$stor2rrdPath\Storrd.csv"

$storrd = Import-CSV $stor2rrdPath\Storrd.csv -delimiter ';'
$collect = @()
foreach($dsRow in $storrd | Where {$_.'Storage System name'-match "SVC" -and $_.'Storage Tier' -eq 3}){
$hostList = @()
    $hostSplit = ($dsRow.'Host Mappings').Split(":")
        foreach($hostName in $hostSplit){
            $vmhost = $hostName + ".mud.internal.co.za"
            $hostList += $vmhost
        }
            $hostCluster = Get-VMHost $hostList | Get-Cluster | Select -ExpandProperty Name

    $dsDetail = [pscustomobject] @{
        "Volume Name" = $dsRow.'Volume Name'
        "Uid" = $dsRow.'LUN UID'
        "Capacity" = $dsRow.'Capacity (GiB)'
        "Storage Pool ID" = $dsRow.'Storage Pool ID'
        "Storage Pool Name" = $dsRow.'Storage Pool Name'
        "Storage System name" = $dsRow.'Storage System name'
        "Storage Tier" = $dsRow.'Storage Tier'
        "Host Mappings" = $hostSplit -join ", "
        "Cluster Mappings" = $hostCluster -join ", "
    }
    $dsDetail
    #$dsDetail | Export-Excel -Path "D:\UserData\Ibraaheem\FSLuns.xlsx" -AutoSize -AutoFilter -BoldTopRow -Append
}