$outputPath = "D:\UserData\Ibraaheem\Scripts\VMWare"
$vmHosts = Get-VMHost

$collect = @()

foreach($vmHost in $vmHosts){

    $vmHostView = $vmHost | Get-View

    $exportLine = [pscustomobject] @{
        "Host" = $vmhost.Name
        "Version" = $vmhost.Version
        "Build" = $vmhost.Build
        "Firmware" = $vmHostView.Hardware.BiosInfo.BiosVersion
        "Release" = $vmHostView.Hardware.BiosInfo.ReleaseDate
        "Cluster" = $vmhost | Get-Cluster
    }

    $exportLine
    $collect += $exportLine

}

$collect | Export-Excel -path "$outputPath\patchLevel.xlsx"
