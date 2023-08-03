$datastore = Get-Datastore SLM-06097-DIA01-NML-CMPTHN01

$esxi= $datastore | Get-VMHost
$cluster = $esxi | Get-Cluster

$datDetails = $datastore | Select Name, CapacityGB, @{N='CanonicalName';E={$_.ExtensionData.Info.Vmfs.Extent[0].DiskName}}

#[MATH]::Round(($datDetails.CapacityGB/1024),0)
$datDetails.Name
$datDetails.CanonicalName
$cluster.Name
$hostNames = $esxi.Name | Sort-Object
$hostNames -join ", "

$collection = @()

foreach($vmHost in $esxi){
    $collection += Get-VMHosthba -VMHost $vmHost -type FibreChannel | where{$_.Status -eq 'online'} |
    Select  @{N="Host";E={$vmHost}},
        @{N="HBA";E={$_.Name}},
        @{N='WWN';E={$wwn = "{0:X}" -f $_.NodeWorldWideName; (0..7 | %{$wwn.Substring($_*2,2)}) -join ':'}} | Sort-Object Host
}

$collection | Sort-Object Host