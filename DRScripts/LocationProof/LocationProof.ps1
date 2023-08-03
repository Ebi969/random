$vms = get-content "D:\UserData\Ibraaheem\Scripts\VMWare\DRScripts\LocationProof\vms.txt"
$alloutput =@()
foreach($vm in $vms){

    $vmDets = Get-VM $vm
    $date = Get-date
    $cluster = $vmDets | Get-Cluster

    $Output = [pscustomobject] @{

        VM = $vmDets.Name
        PowerState = $vmDets.PowerState
        Cluster = $cluster 
        Time = $date

    }
   $alloutput += $Output 
}
$alloutput | Out-Gridview