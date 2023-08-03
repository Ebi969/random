# EXAMPLE Host
$exampleHost = "slm02esx230.mud.internal.co.za"
# NEW HOST
$vmhosts = "slm01esx230.mud.internal.co.za"

# EXAMPLE Host
$vlans1 = Get-VMHost $exampleHost | Get-VirtualSwitch -Name vSwitch1 | Get-VirtualPortGroup | Select Name,VLanID
$sec1 = Get-VMHost $exampleHost | Get-VirtualSwitch -Name vSwitch1 | Get-VirtualPortGroup | Get-SecurityPolicy 
$vlans2 = Get-VMHost $exampleHost | Get-VirtualSwitch -Name vSwitch2 | Get-VirtualPortGroup | Select Name,VLanID
$sec2 = Get-VMHost $exampleHost | Get-VirtualSwitch -Name vSwitch2 | Get-VirtualPortGroup | Get-SecurityPolicy 


foreach($vmhost in $vmhosts)
{ 
    $vhost = Get-VMHost $vmhost
    Get-VMHost $vhost.Name | New-VirtualSwitch -Name "vSwitch1"
    Start-Sleep -Seconds 2 
    Get-VMHost $vhost.Name | Get-VirtualSwitch -Name vSwitch0 | Get-SecurityPolicy | Set-SecurityPolicy -MacChanges:$false -AllowPromiscuous:$false -ForgedTransmits:$false
    Get-VMHost $vhost.Name | Get-VirtualSwitch -Name vSwitch1 | Get-SecurityPolicy | Set-SecurityPolicy -MacChanges:$false -AllowPromiscuous:$false -ForgedTransmits:$false

    foreach($vlan in $vlans1) 
     { 
        Get-VMHost $vhost.Name | Get-VirtualSwitch -Name vSwitch1 | New-VirtualPortGroup -Name $vlan.Name -VLanId $vlan.VLanID
     }
  
    Get-VMHost $vhost.Name | New-VirtualSwitch -Name "vSwitch2"

    foreach($vlan in $vlans2) 
     { 
        Get-VMHost $vhost.Name | Get-VirtualSwitch -Name vSwitch2 | New-VirtualPortGroup -Name $vlan.Name -VLanId $vlan.VLanID
     }

    foreach($sec in $sec1) 
     { 
     
        Get-VMHost $vhost.Name | Get-VirtualSwitch -Name vSwitch1 | Get-VirtualPortGroup -Name $sec.VirtualPortGroup.Name | Get-SecurityPolicy | Set-SecurityPolicy -MacChanges $sec.MacChanges -AllowPromiscuous $sec.AllowPromiscuous -ForgedTransmits $sec.ForgedTransmits
     }

    foreach($sec in $sec2) 
     { 
     
        Get-VMHost $vhost.Name | Get-VirtualSwitch -Name vSwitch2 | Get-VirtualPortGroup -Name $sec.VirtualPortGroup.Name | Get-SecurityPolicy | Set-SecurityPolicy -MacChanges $sec.MacChanges -AllowPromiscuous $sec.AllowPromiscuous -ForgedTransmits $sec.ForgedTransmits
     }


   Get-VMHost $vhost.Name | Add-VMHostNtpServer -NtpServer "ntptime1.sanlam.co.za","ntptime2.sanlam.co.za"

        Get-VMHostFirewallException -VMHost $vhost.Name | where {$_.Name -eq "NTP client"} | Set-VMHostFirewallException -Enabled:$true
        Start-Sleep -s 2
        Get-VmHostService -VMHost $vhost.Name | Where-Object {$_.key -eq "ntpd"} | Start-VMHostService
        Start-Sleep -s 2
        Get-VmHostService -VMHost $vhost.Name | Where-Object {$_.key -eq "ntpd"} | Set-VMHostService -policy "on"
        Start-Sleep -s 2
        Get-VMHostFirewallException -VMHost $vhost.Name | where {$_.Name -eq "syslog"} | Set-VMHostFirewallException -Enabled:$true

        Start-Sleep -s 2

        Get-VMHost $vmhosts  | Get-AdvancedSetting -Name VMFS3.HardwareAcceleratedLocking | Set-AdvancedSetting -Value 1 -Confirm:$false
        Start-Sleep -s 2

        Get-VMHost $vmhosts | Get-AdvancedSetting -Name VMFS3.UseATSForHBOnVMFS5 | Set-AdvancedSetting -Value 0 -Confirm:$false
        Start-Sleep -s 2


        Get-VMHost $vmhosts  | Get-AdvancedSetting -Name VMFS3.MaxHeapSizeMB | Set-AdvancedSetting -Value 256 -Confirm:$false
        Start-Sleep -s 2

        Get-VMHost $vmhosts| Get-AdvancedSetting -Name Syslog.global.logDir | Set-AdvancedSetting -Value "[] /scratch/log" -Confirm:$false
        Start-Sleep -s 2

        Get-VMHost $vmhosts | Get-AdvancedSetting -Name Syslog.global.logHost | Set-AdvancedSetting -Value "srv006711.mud.internal.co.za" -Confirm:$false
        Start-Sleep -s 2


        }