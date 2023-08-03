$clusterList = Get-Cluster

$allVMs = @()
foreach($cluster in $clusterList){

    $vms = $cluster | Get-VM | Select Name, @{l="ID"; e={$_.ExtensionData.Config.GuestID}}, @{l="OS"; e={$_.ExtensionData.Guest.GuestFullName}}

        foreach($vm in $vms){
        $os = $null

        if($vm.ID -match "win"){
            $os = "Windows"
        }elseif($vm.ID -match "rhel|sles|linux"){
            $os = "Linux"
        }elseif($vm.ID -match "Cento"){
            $os = "CentOS"
        }elseif($vm.ID -match "freebsd"){
            $os = "FreeBSD"
        }elseif($vm.ID -match "ubuntu"){
            $os = "Ubuntu"
        }elseif($vm.ID -match "Solaris"){
            $os = "Solaris"
        }else{
            $os = "Other"
        }

            
            $outputRow = [pscustomobject] @{

                Cluster = $cluster.Name
                vmName = $vm.Name
                GuestID = $vm.ID
                OS = $os

            }
            $allVMs += $outputRow
        }
}

$allVMs | Export-Excel -Path "D:\UserData\Ibraaheem\Scripts\VMWare\OSTypes\ClusterOSCount.xlsx" -AutoSize -BoldTopRow -AutoFilter