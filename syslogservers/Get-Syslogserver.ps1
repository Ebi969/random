$outputLocation = "D:\UserData\Ibraaheem\Scripts\VMWare\syslogservers"
$vmHosts = Get-VMHost

$fullDetail = @()
foreach($vmHost in $vmHosts){

$cluster = $null
$vCenter = $null
$syslogDetails = $null
$hostDetails = $null

    $cluster = $vmHost.Parent
    $vCenter = $vmHost.Uid.Split('@')[1].Split(':')[0].Split(".")[0]
    $syslogDetails = $vmHost | Get-VMHostSysLogServer

    $hostDetails = [pscustomobject] @{    
        vCenter = $vCenter
        Cluster = $cluster.Name
        vmHost = $vmHost.Name
        syslogserver = $syslogDetails.Host
    }
    $hostDetails
    $fullDetail += $hostDetails

}

$fullDetail | Export-Excel -Path $outputLocation\vmHostList.xlsx -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow