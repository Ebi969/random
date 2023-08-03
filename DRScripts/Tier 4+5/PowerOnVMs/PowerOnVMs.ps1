#import VM List
$path = "D:\UserData\Ibraaheem\Scripts\VMWare\DRScripts\Tier 4+5\PowerOnVMs"
$vms = Get-content "$path\vmList.txt"
$outputPath = "$path\ShutdownReport.xlsx"

# Get total count of VMs
$totalVMs = $vms.Count

$chunkSize = 50
$chunks = for($i=0; $i -lt $vms.Length; $i += $chunkSize){
    , ($vms | select -Skip $i -First $chunkSize)
}

foreach($chunk in $chunks){ 

    foreach($vm in $chunk){
       Write-Host "Powering up" $vm
       #Get-VM $vm | Start-VM -ErrorAction SilentlyContinue -Confirm:$false
    }  
    if($chunk.lenth -eq 50){  
        Write-host "Sleep 10 sec before next batch"
        #Start-Sleep -Seconds 10
    }
}

"T minus 60 seconds till VMs are online" 
Start-Sleep -Seconds 60
$vmDetails = Get-VM $vms -ErrorAction SilentlyContinue | Select Name, PowerState, NumCpu, MemoryGB, VMHost
While(($vmDetails | Where {$_.PowerState -match "Off"}.length -gt 0)){
    Write-Host "VMs Still powering up:`n"
    $vmDetails | Where {$_.PowerState -match "On"} | Select -ExpandProperty Name
    ""
    Read-Host “Press ENTER to Check For remaining servers...”
    $vmDetails = Get-VM $vms -ErrorAction SilentlyContinue | Select Name, PowerState, NumCpu, MemoryGB, VMHost
}

$vmDetails | Export-Excel $outputPath -Append -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow
Write-Host "All VMs successfully powered on, Report exported to: $outputPath"