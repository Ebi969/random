$path = "D:\UserData\Ibraaheem\Scripts\VMWare\Lun Requests\Using existing LUN"
$list = Get-Content $path\list.txt

$allDS = Get-Datastore | Where Name -Match "SCRATCH"

foreach($ds in $list){

    $datastoreDetails = $allDS | Where Name -Match "$ds"
    $vmHosts = $datastoreDetails | Get-VMHost
    $cluster = ( $vmHosts | Get-Cluster | Select -ExpandProperty Name) -join ", "
    <##>
    $hbaCollect = @()
    foreach($vmHost in $vmHosts){
        $hbaCollect += Get-VMHosthba -VMHost $vmHost -type FibreChannel | where{$_.Status -eq 'online'} |
        Select  @{N="Host";E={$vmHost}},
            @{N="HBA";E={$_.Name}},
            @{N='WWN';E={$wwn = "{0:X}" -f $_.NodeWorldWideName; (0..7 | %{$wwn.Substring($_*2,2)}) -join ':'}} | Sort-Object Host
    }
    $hbaCollect = $hbaCollect | Sort-Object Host
    <##>
    $datastoreDetails
    $cluster
    $hbaCollect

    ""
}