$datastores = Get-Content "\\slmcdc01phy001\D$\UserData\Ibraaheem\Scripts\VMWare\list.txt"

    $datListDetail = @()
foreach($dat in $datastores){

    $datDetails = Get-Datastore $dat.replace(" ", "")
    $clusters = $datDetails | Get-VMHost | Get-Cluster | Select-Object -ExpandProperty Name
    $uid = $datDetails.ExtensionData.Info.Vmfs.Extent[0].DiskName
    $numClusters = ($clusters | Measure-Object -Line).Lines

    $datastoreDetails = [pscustomobject] @{
        'Name' = $dat
        'Capacity' = $datDetails.CapacityGB
        'UID' = $uid
        'Clusters' = $clusters -join ", "
        'Num Cluster' = $numClusters
    }

    $datListDetail += $datastoreDetails
}

$datListDetail | Export-Excel "\\slmcdc01phy001\D$\UserData\Ibraaheem\Scripts\VMWare\LUNDetails.xlsx"