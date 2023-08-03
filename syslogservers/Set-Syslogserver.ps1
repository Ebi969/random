$outputLocation = "D:\UserData\Ibraaheem\Scripts\VMWare\syslogservers"
$newSysLogServer = "SRV006712.mud.internal.co.za"

$vmHostList = Get-content -Path "D:\UserData\Ibraaheem\Scripts\VMWare\syslogservers\serverList.txt" 

$beforefullDetail = @()
$afterfullDetail = @()

foreach($vmHostName in $vmHostList){

    $vmHost = Get-VMHost $vmHostName 

    $cluster = $null
    $vCenter = $null
    $syslogDetails = $null
    $hostDetailsRow = $null

    $cluster = $vmHost.Parent
    $vCenter = $vmHost.Uid.Split('@')[1].Split(':')[0].Split(".")[0]
    $syslogDetails = $vmHost | Get-VMHostSysLogServer

    $hostDetailsRow = [pscustomobject] @{    
        vCenter = $vCenter
        Cluster = $cluster.Name
        vmHost = $vmHost.Name
        syslogserver = $syslogDetails.Host
    }

    $beforefullDetail += $hostDetailsRow

    $vmHost | Set-VMHostSysLogServer -SysLogServer $newSysLogServer

    $syslogDetails = $null
    $syslogDetails = $vmHost | Get-VMHostSysLogServer

    $afterhostDetailsRow = [pscustomobject] @{    
        vCenter = $vCenter
        Cluster = $cluster.Name
        vmHost = $vmHost.Name
        syslogserver = $syslogDetails.Host
    }

    $afterfullDetail += $afterhostDetailsRow

}

$beforefullDetail | Export-Excel -Path $outputLocation\hostSysLogs.xlsx -WorksheetName "Before" -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow -Append
$afterfullDetail | Export-Excel -Path $outputLocation\hostSysLogs.xlsx -WorksheetName "After" -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow -Append