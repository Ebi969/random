$datastore = Read-Host "Please Enter Datastore Name"

$dataRealName = $datastore.Replace(" ","")

$dataDetails = Get-Datastore $dataRealName

$naaID = $dataDetails.ExtensionData.Info.Vmfs.Extent[0].DiskName

$vmHosts = $dataDetails | Get-VMHost | Sort-Object Name
$clusters = $vmHosts | Get-Cluster | Select -ExpandProperty Name

$vmHosts = $vmHosts.Name  -join ", "
$clusters = $clusters -join ", "

"Name: " + $dataRealName
"LUN Naa: " + $naaID
"Clusters assigned: `n" + $clusters
"Hosts assigned: `n" + $vmHosts

