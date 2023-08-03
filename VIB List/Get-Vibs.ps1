$VMHosts = Get-VMHost stmcdc02esx009.mud.internal.co.za

foreach ($VMHost in $VMHosts) {

    $VMHostName = $VMhost.Name

    $esxcli = $VMHost | Get-EsxCli

    $List += $esxcli.software.vib.list() | Select-Object @{N="VMHostName"; E={$VMHostName}}, *

}