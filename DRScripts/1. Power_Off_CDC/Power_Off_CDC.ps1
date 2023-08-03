$vms = Import-Csv  "D:\UserData\Ibraaheem\Scripts\VMWare\DRScripts\1. Power_Off_CDC\vms.csv"

foreach($vm in $vms) 
 { 

   $vmname = $vm.Name
   Write-Host "Shutting Down" $vmname
   Shutdown-VMGuest $vmname -Confirm:$false

 }

"T minus 60 seconds till VMs are offline" 
start-sleep 60
get-vm $vms.name | Select Name, PowerState, NumCpu, MemoryGB | Sort-Object PowerState -Descending | ft -AutoSize