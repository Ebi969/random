$toinventory = Import-Excel "D:\UserData\Ibraaheem\Scripts\VMWare\DRScripts\7. Inventory_CDC\ToInventoryCDC.xlsx"

foreach($vm in $toinventory) 
{  
    $vmname = ($vm.VMName).ToUpper() 
    $path = $vm.VMXPath
    $Cluster = $vm.newCluster
    if($Cluster -match "STM-CDC-GOLD-05-SQL"){
        $vmhost = Get-VMHost stm02esx080.mud.internal.co.za
    }else{
        $vmhost = Get-Cluster $Cluster | Get-VMHost | Get-Random
    }
    $hostname = $vmhost.Name

    Write-Host "Inventorying" $vmname
    $inventory = New-VM -VMFilePath $path -VMHost $hostname -Name $vmname
    $inventory | Export-Csv "D:\Scripts\VMware\DR\DRScripts\7. Inventory_CDC\ReinventoriedVMs.csv" -Append
    $inventory
    Start-Sleep 2

    $allnics = ($vm.Nics).split(";")
    foreach($nic in $allnics){
        $nicSplit = $nic.Split(":")
        $nicName = $nicSplit[0]
        $newVLAN = $nicSplit[1]
        Get-VM $vm.VMName | Get-NetworkAdapter | Where {$_.Name -eq $nicName} | Set-NetworkAdapter -NetworkName $newVLAN -Confirm:$false
    }    
}