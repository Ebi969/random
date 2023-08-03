$allv6Datastores = Get-Datastore | Where {$_.ExtensionData.Info.Vmfs.MajorVersion -eq 6}
$collect = @()
foreach($dat in $allv6Datastores){
    $details = $dat.ExtensionData.Info.Vmfs
    $collect += $details
}
$collect | Out-GridView