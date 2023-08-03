#$LunName = Get-Content "D:\UserData\Ibraaheem\Scripts\VMWare\LUN Naa\LunList.txt"
#$LunName = Get-Datastore SKY05-GLD01-SAP-NMP-DAT14 
$LunName = Get-Cluster SLM-CDC-GOLD-04-WAS | Get-Datastore  #Datastore name or host name - host name gives all datastores assigned to the host, datastore is specific to one DS
$LunName = Get-vmhost (get-content D:\UserData\Siyah\VMWareScripts\GetLUNs\LUNList.txt) | Get-Datastore  #Datastore name or host name - host name gives all datastores assigned to the host, datastore is specific to one DS

$collect = @()
foreach($Lun in $LunName){

$collect += Get-Datastore $Lun | Select Name, @{N='Naa';E={$_.ExtensionData.Info.Vmfs.Extent[0].DiskName}}, CapacityGB, Accessible
}
$collect | Out-GridView