$vms = Import-Csv  "D:\UserData\Ibraaheem\Scripts\VMWare\DRScripts\2. Uninventory_CDC\vms.csv"

foreach($vm in $vms)
 { 
    $vmname = $vm.Name
    $vmDetails = Get-VM $vmname
    $vmxPath = $vmDetails.ExtensionData.Config.Files.VmPathName
    $currentcluster = $vmDetails | Get-cluster
     
    $vmNics = $vmDetails | Get-NetworkAdapter
    $allNics = @()
    foreach($vmNic in $vmNics){

        $singleNic = $vmNic.Name + ":" + $vmNic.NetworkName
        $allNics += $singleNic
    }

    $collectiveNics = $allNics -join ";"

    if( $currentcluster.name -match "BDC01"){
        $cluster = $currentcluster.name.replace("BDC01","CDC02")
    }elseif( $currentcluster.name -match "BDC"){
        $cluster = $currentcluster.name.replace("BDC","CDC")
    }elseif( $currentcluster.name -match "CDC02"){
        $cluster = $currentcluster.name.replace("CDC02","BDC01")
    }elseif( $currentcluster.name -match "CDC"){
        $cluster = $currentcluster.name.replace("CDC","BDC")
    }

    $vmoutput = [PSCustomObject] @{
        VMName	= $vmname
        VMXPath	= $vmxPath
        Nics = $collectiveNics
        oldCluster =$currentcluster.Name
        newCluster = $cluster
    }
    $vmoutput | Export-Excel -path "D:\Scripts\VMware\DR\DRScripts\3. Inventory_BDC\toInventoryBDC.xlsx" -Append

   ## Remove-VM $vmname -Confirm:$false -ErrorAction Inquire
 }  

"T minus 15 seconds till VMs are removed from VC" 

start-sleep 15

get-vm $vms.name -ErrorAction SilentlyContinue | Sort-Object PowerState -Descending