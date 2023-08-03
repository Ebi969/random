$vmHosts = Get-VMHost | Where-Object {$_.ConnectionState -like "Connected" -or "Maintenance"}    

$allDetails = @()

foreach($vmhost in $vmHosts){

    $NetSystem = Get-View $VMHost.ExtensionData.ConfigManager.NetworkSystem
    $esxDetials = Get-Esxcli -vmhost $vmHost.Name -V2
    $nicList = $esxDetials.network.nic.list.invoke()
    $allSwitches = $vmHost | Get-VirtualSwitch
    $distributedSwitches = $vmHost | Get-VDSwitch
    foreach($hostNic in $nicList){
        foreach ($Pnic in $VMHost.ExtensionData.Config.Network.Pnic) {            
            if($Pnic.Device -eq $hostNic.Name){

                $switch = $allSwitches | Where {$_.Nic -eq $hostNic.Name}
                $switchName = $switch.Name
                if($switchName -eq $null){
                    foreach($vds in $distributedSwitches){
                        $vdsName = $vds.Name
                        $vdsNics = $null
                        $vdsNics = $vmhost | Get-VMHostNetworkAdapter -DistributedSwitch $vdsName
                        if($vdsNics.Name -contains $hostNic.Name){
                            $switchName = $vdsName
                        }
                    }
                }

                $PnicInfo = $NetSystem.QueryNetworkHint($Pnic.Device)
                #$PnicInfo.ConnectedSwitchPort
                     $nic = [PSCustomObject] @{
                        'Cluster' = $vmhost.Parent
                        'Host' = $vmhost.Name
                        'Device' = $Pnic.Device
                        'vSwitch' = $switchName
                        'LinkStatus' = $hostNic.LinkStatus
                        'PortID' = $PnicInfo.ConnectedSwitchPort.PortId
                        'DevID' = $PnicInfo.ConnectedSwitchPort.DevId
                        'SwitchIP' = $PnicInfo.ConnectedSwitchPort.Address
                        'MacAddress' = $hostNic.MACAddress
                        'Speed' = $hostNic.Speed
                        'Duplex' = $hostNic.Duplex
                        'Vendor' = $hostNic.Description
                        'Driver' = $hostNic.Driver
                        'PCIDevice' = $hostNic.PCIDevice
                        'MTU' = $hostNic.MTU
                    }
                $allDetails += $nic
            }
        }

    }

}

$outputPath = "D:\UserData\Ibraaheem\Scripts\VMWare\CDP on Host"
#$outputPath = "\\srv005879\D$\Reports\VMware\DailyChecks"
$timeStamp = Get-Date -Format yyyy-MM-dd
$fileName = $outputpath + "\VMHostPhysicalNicDetails_" + $timeStamp + ".xlsx"

#$allDetails | Export-Excel -Path $fileName -WorksheetName "Nic Details" -AutoSize -BoldTopRow -AutoFilter -Append

<#

$report = @()
foreach($sw in ($vmHosts | Get-VirtualSwitch -Distributed)){
    $uuid = $sw.ExtensionData.Summary.Uuid
    $sw.ExtensionData.Config.Host | %{
        $esx = Get-View $_.Config.Host
        $netSys = Get-View $esx.ConfigManager.NetworkSystem
        $netSys.NetworkConfig.ProxySwitch | where {$_.Uuid -eq $uuid} | %{
            $_.Spec.Backing.PnicSpec | %{
                $row = "" | Select Host,dvSwitch,PNic
                $row.Host = $esx.Name
                $row.dvSwitch = $sw.Name
                $row.PNic = $_.PnicDevice
                $report += $row            }
        }
    }
}
$report 

#>