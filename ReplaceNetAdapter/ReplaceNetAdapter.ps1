$vm = "VLAN Test - Santam"

<#
Remove Nic - server requires powered off
Get-VM $vm | Get-NetworkAdapter | Where {$_.Name -eq "Network adapter 2"} | Remove-NetworkAdapter 
#>

#add network adapter
Get-Vm $vm | New-NetworkAdapter -NetworkName "Internal" -StartConnected -WakeOnLan:$true -Type Vmxnet3