#$wwn = "{0:X}" -f $_.NodeWorldWideName; (0..7 | %{$wwn.Substring($_*2,2)}) -join ':'
$cluster = "STM-BDC-GOLD-02-LINUX"

$clusterHosts = Get-Cluster $cluster | Get-VMHost
"Cluster: " + ($cluster)
"Hosts: " + ($clusterHosts.Name -join ", " | Sort-Object)

$collect = @()
foreach($esxi in $clusterHosts.Name){

    $collect += Get-VMHosthba -VMHost $esxi -type FibreChannel | where{$_.Status -eq 'online'} |
    Select  @{N="Host";E={$esxi}}, @{N="HBA";E={$_.Name}} ,@{N='WWN';E={"{0:x}" -f $_.PortWorldWideName}} | Sort-Object Host
}

$collect | Sort-Object Host

#Get-Datastore | Where {$_.Name -like "SLM02-BRN04-QUA-NMP-DAT*"} | Sort-Object Name