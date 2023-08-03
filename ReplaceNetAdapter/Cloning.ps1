$Template = Get-Template -Name 'PAMCOTMPWIN2019'
$vmHost = Get-VMHost "slm01es0xx.mud.internal.co.za"
$datastore = Get-Datastore "xx-xxx-xx-xxxx"
$vmName = "SRV00xxxx"

New-VM -Name $vmName -Template $Template -VMHost $vmHost -Datastore $datastore