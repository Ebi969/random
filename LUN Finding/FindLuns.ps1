$path = "D:\UserData\Ibraaheem\Scripts\VMWare\LUN Finding"
$UIdentifiers = "6b46e081005270eb|6cc05771007c3f00"
$LunName = Get-Datastore | Where {$_.ExtensionData.Info.Vmfs.Extent[0].DiskName -notmatch $UIdentifiers}

$collect = @()
foreach($Lun in $LunName){
    $collect += Get-Datastore $Lun | Select Name, @{N='Naa';E={$_.ExtensionData.Info.Vmfs.Extent[0].DiskName}}, CapacityGB, Accessible
}
$collect | Export-Excel "$path\export.xlsx" -BoldTopRow -FreezeTopRow -AutoSize -AutoFilter