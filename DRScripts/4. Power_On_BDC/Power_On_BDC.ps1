$vms = Import-Csv "D:\UserData\Ibraaheem\Scripts\VMWare\DRScripts\4. Power_On_BDC\vms.csv"

foreach($vm in $vms)
{

   $vmname = $vm.Name
   $vm = Get-VM $vmname
   Write-Host "Powering On" $vmname
   start-sleep 1
   $vm | Start-VM -Confirm:$false -ErrorAction SilentlyContinue
   $vm | Export-Csv "D:\Scripts\VMware\DR\DRScripts\4. Power_On_BDC\Started VMS.csv" -Append 
   $vm
   $vm | Get-VMQuestion | Set-VMQuestion -Option ‘button.uuid.movedTheVM’ -Confirm:$false
   
   Start-Sleep -Seconds 2
 }

"T minus 15 seconds till VMs are running..." 

start-sleep 15

get-vm $vms.name -ErrorAction SilentlyContinue | Sort-Object PowerState