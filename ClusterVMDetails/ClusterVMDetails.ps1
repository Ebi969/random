$vmhosts = Get-cluster SLM-BDC-BRONZE-04-WINDOWS, SLM-CDC-BRONZE-04-WINDOWS | Get-VMHost

$allDetails = @()
foreach($vmHost in $vmhosts){

    $vms = $vmHost | Get-VM

    foreach($vm in $vms){

        $vmView = $vm | Get-View
        $cps = $vmView.config.hardware.NumCoresPerSocket
        $sockets = $vmView.config.hardware.NumCPU / $cps
        $vmObject = [pscustomobject] @{

            Cluster = $vmHost.Parent
            Host = $vmHost.Name
            VM = $vm.Name
            vCPU = $vm.NumCpu
            Sockets = $sockets 
            CPS = $cps
            Memory = $vm.MemoryGB
        }

        $vmObject

        $allDetails += $vmObject

    }
}

$allDetails | Export-Excel "D:\UserData\Ibraaheem\Scripts\VMWare\ClusterVMDetails\Details.xlsx" -AutoSize -AutoFilter -BoldTopRow
