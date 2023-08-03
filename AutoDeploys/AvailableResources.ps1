Get-date
#################################################

    ### Variables ###

#################################################

# General Variables #
$outputPath = "D:\UserData\Ibraaheem\Scripts\VMWare\AutoDeploys"
$clusterMatrix = Import-Excel "\\srv005879\Reports\AutoDeploy\ConfigManagement.xlsx" -WorksheetName "HostingConfig"
$fileName = "$outputPath\AvailableResources.xlsx"
$clusSheet = "Cluster"
$storSheet = "Storage"
$hostSheet = "Host"
$defaultCPUMatrix = 4
$matrixMem = 80

# HyperV Variables #
$HostGroup = "All Hosts"
$HyperVC = "CLS000288"

# VMware Variables #
$VMwareVC = @('SRV007281','SRV007282','SRV008097')

# Remove current resource file #
if(Test-Path $fileName){Remove-Item $fileName}

#################################################

    ### VMware resource calculations ###

#################################################

<#
### Connect To VMware VCs ###
foreach($vCenter in $VMwareVC){
    $vcCreds = Get-VICredentialStoreItem -file  "C:\Users\svcVMWareScriptAcc\CredStore$($vCenter.substring($vCenter.length - 4, 4)).creds"
    Connect-VIServer -Server "$vCenter.mud.internal.co.za" -User $vcCreds.User -Password $vcCreds.Password
}
#>
$clusters = Get-Cluster | Where {($_| Get-Datacenter) -notmatch "PHYSICAL|CCC|REGIONAL"} | Sort-Object Name
foreach($cluster in $clusters){
    # get hosts in the cluster, if no hosts, skip cluster #
    $vmHosts = $cluster | Get-VMHost
    if(!($vmHosts)){
        continue 
    }
    
    # import clustermatrix to get current clusters ratio #
    $matrixCpu = $clusterMatrix | Where {$_.Cluster -eq $cluster.Name} | Select -ExpandProperty cpu_ratio -Unique
    # default to 4 ratio if config file doesn't have cluster #
    $matrixCpu = $defaultCPUMatrix  

        # get host allocations #
        $vmwareHostOutput = @()
        foreach($vmHost in $vmHosts){
            # get all vms on host #
            $hostVMs = $vmHost | Get-VM

            # get host vms total allocations #
            $hostvmCPUs = ($hostVMs | Measure-Object NumCpu -Sum).Sum
            $hostvmMem = ($hostVMs | Measure-Object MemoryGB -Sum).Sum

            # get host total mem and cpu #
            $hostTotalCPU = $vmHost.NumCpu
            $hostTotalMemGB = [MATH]::Round($vmHost.MemoryTotalGB ,2)

            # calculate current ratios #
            $hostCpuRatio = [MATH]::Round(($hostvmCPUs/$hostTotalCPU), 2)
            $hostMemPerc = [MATH]::Round(($hostvmMem/$hostTotalMemGB)*100, 2)   

            # calculate available resources #
            $hostToBeLimitCpu = ($hostTotalCPU*$matrixCpu) - $hostvmCPUs    
            $hostToBeLimitMem = (($matrixMem/100) * $hostTotalMemGB) - $hostvmMem 

                $hostRow = [pscustomobject] @{
                    Cluster = $cluster.Name
                    Host = $vmHost.Name
                    Status = $vmhost.ConnectionState
                    HostTotalCPU = $hostTotalCPU
                    CPUallocation = $hostvmCPUs
                    CPURatio = $hostCpuRatio
                    AvailableCpu = $hostToBeLimitCpu
                    HostTotalMem = [MATH]::Round($hostTotalMemGB, 2)
                    MEMallocation = [MATH]::Round($hostvmMem, 2)
                    MemPercAllocated = $hostMemPerc
                    AvailableMem = [MATH]::Round($hostToBeLimitMem, 2)
                }
                $vmwareHostOutput += $hostRow
        }

    # get all vms in cluster #
    $vms = $cluster | Get-VM

    # get clusters vms total allocations #
    $vmCPUs = ($vms | Measure-Object NumCpu -Sum).Sum
    $vmMem = ($vms | Measure-Object MemoryGB -Sum).Sum

    # get clusters total mem and cpu #
    $totalCPU = ($vmHosts | Measure-Object NumCpu -Sum).Sum
    $totalMemGB = [MATH]::Round(($vmHosts | Measure-Object MemoryTotalGB -Sum).Sum ,2)

    # calculate current ratios #
    $cpuRatio = [MATH]::Round(($vmCPUs/$totalCPU), 2)
    $memPerc = [MATH]::Round(($vmMem/$totalMemGB)*100, 2)

    # get views to get VC name #
    $clusView = $cluster | Get-View
    $vcURL = $clusView.Client.ServiceUrl
    $split = $vcURL.Split(".")
    $vcName = ($split[0].Split("/"))[2] 

    # calculate available resources #
    $toBeLimitCpu = ($totalCPU*$matrixCpu) - $vmCPUs
    $toBeLimitMem = (($matrixMem/100) * $totalMemGB) - $vmMem

    # get datastore information and calculations #
    $vmwareDatastoreOutput = $cluster | Get-Datastore | Where {$_.Name -notmatch "datastore|pamco|stage|local|scratch|NFS|DNU|SQLDEV|srv|-0|-1|-3|SRM|DRP|DRT|MAP|PDF|ARC|STR"} | Select @{n="ClusterName"; e={$cluster.Name}}, 
    @{n="vCenter"; e={$vcName}}, 
    @{n='Name';e={$_.Name}},
    @{n="FreeSpaceGB"; e={[MATH]::Round($_.FreeSpaceGB,2)}}, 
    @{n="CapacityGB"; e={[MATH]::Round($_.CapacityGB,2)}}, 
    @{n="AvailableStorage"; e={[MATH]::Round(($_.FreeSpaceGB - ($_.CapacityGB * 0.1)),2)}}

    $vmwareClusterOutput = [pscustomobject] @{
        Cluster = $cluster.Name
        vCenter = $vcName
        TotalCPUinCluster = $totalCPU
        CPUallocation = $vmCPUs
        CPURatio = $cpuRatio
        AvailableCpu = $toBeLimitCpu
        TotalMEMinCluster = [MATH]::Round($totalMemGB, 2)
        MEMallocation = [MATH]::Round($vmMem, 2)
        MemPercAllocated = $memPerc
        AvailableMem = [MATH]::Round($toBeLimitMem, 2)
    }

# Export to excel #
$vmwareClusterOutput | Export-Excel -Path $fileName -WorksheetName $clusSheet -Append -BoldTopRow -AutoFilter -FreezeTopRow -AutoSize
$vmwareHostOutput | Export-Excel -Path $fileName -WorksheetName $hostSheet -Append -BoldTopRow -AutoFilter -FreezeTopRow -AutoSize
$vmwareDatastoreOutput | Export-Excel -Path $fileName -WorksheetName $storSheet -Append -BoldTopRow -AutoFilter -FreezeTopRow -AutoSize

}

#################################################

    ### HyperV resource calculations ###

#################################################

# Import-Modules
Import-Module -Name VirtualMachineManager

$SCVMHostCluster = Get-SCVMHostCluster #-VMHostGroup $SCVMHostGroup
foreach ($cluster in $SCVMHostCluster) {
    # get hosts in the cluster, if no hosts, skip cluster #
    $nodes = $cluster.Nodes
    if(!($nodes)){
        continue 
    }

    # import clustermatrix to get current clusters ratio #
    $matrixCpu = $clusterMatrix | Where {$_.Cluster -eq $cluster.Name} | Select -ExpandProperty cpu_ratio -Unique
    # default to 4 ratio if config file doesn't have cluster #
    $matrixCpu = $defaultCPUMatrix
        
    $allocatedvCPU = $null
    $allocatedvMem = $null

        $hpvHostOutput = @()  
        foreach ($node in $nodes) {
            $nodeAllocatedvCPU = ((Get-SCVirtualMachine -VMHost $node | Measure-Object -Property CPUCount -Sum).Sum)
            $nodeAllocatedMem = ((Get-SCVirtualMachine -VMHost $node | Measure-Object -Property Memory -Sum).Sum/1kb)
            $allocatedvCPU += $nodeAllocatedvCPU
            $allocatedvMem += $nodeAllocatedMem

            # get clusters vms total allocations #
            $nodevmCPUs = $nodeAllocatedvCPU
            $nodevmMem = $nodeAllocatedMem

            # get clusters total mem and cpu #
            $nodeTotalCPU = $node.LogicalCPUCount
            $nodeTotalMemGB = [math]::Round(($node.TotalMemory)/1gb,2)

            # calculate current ratios #
            $nodeCpuRatio = [math]::Round(($nodevmCPUs/$nodeTotalCPU),2)
            $nodeMemPerc = [MATH]::Round(($nodevmMem/$nodeTotalMemGB)*100, 2) 

            # calculate available resources #
            $nodeToBeLimitCpu = ($nodeTotalCPU*$matrixCpu) - $nodevmCPUs
            $nodeToBeLimitMem = (($matrixMem/100) * $nodeTotalMemGB) - $nodevmMem

            $nodeRow = [pscustomobject] @{
                Cluster = $cluster.Name
                Host = $node.Name
                Status = $node.ComputerState
                HostTotalCPU = $nodeTotalCPU
                CPUallocation = $nodevmCPUs
                CPURatio = $nodeCpuRatio
                AvailableCpu = $nodeToBeLimitCpu
                HostTotalMem = [MATH]::Round($nodeTotalMemGB, 2)
                MEMallocation = [MATH]::Round($nodevmMem, 2)
                MemPercAllocated = $nodeMemPerc
                AvailableMem = [MATH]::Round($nodeToBeLimitMem, 2)
            }
            $hpvHostOutput += $nodeRow
        }

    # get clusters vms total allocations #
    $vmCPUs = $allocatedvCPU
    $vmMem = $allocatedvMem

    # get clusters total mem and cpu #
    $totalCPU = ((($cluster.Nodes).LogicalCPUCount | Measure-Object -sum).sum)
    $totalMemGB = [math]::Round(((($cluster.Nodes).TotalMemory | Measure-Object -sum).sum)/1gb,2)

    # calculate current ratios #
    $cpuRatio = [math]::Round(($vmCPUs/$totalCPU),2)
    $memPerc = [MATH]::Round(($vmMem/$totalMemGB)*100, 2) 

    # calculate available resources #
    $toBeLimitCpu = ($totalCPU*$matrixCpu) - $vmCPUs
    $toBeLimitMem = (($matrixMem/100) * $totalMemGB) - $vmMem
    
    $hpvClusterOutput = [pscustomobject] @{
        Cluster = $cluster.Name
        vCenter = $HyperVC
        TotalCPUinCluster = $totalCPU
        CPUallocation = $vmCPUs
        CPURatio = $cpuRatio
        AvailableCpu = $toBeLimitCpu
        TotalMEMinCluster = [MATH]::Round($totalMemGB, 2)
        MEMallocation = [MATH]::Round($vmMem, 2)
        MemPercAllocated = $memPerc
        AvailableMem = [MATH]::Round($toBeLimitMem, 2)
    }

    ### STORAGE INFO ###

    $ClusterVolumes = $Cluster.SharedVolumes
    $hpvDatastoreOutput = $ClusterVolumes.GetEnumerator() | Select @{n="ClusterName"; e={$cluster.Name}}, 
    @{n="vCenter"; e={$vcName}}, 
    @{n='Name';e={$_.Name}},
    @{n="FreeSpaceGB"; e={[MATH]::Round(($_.freespace/1gb),2)}}, 
    @{n="CapacityGB"; e={[MATH]::Round(($_.Capacity/1gb),2)}}, 
    @{n="AvailableStorage"; e={[MATH]::Round((($_.freespace/1gb) - (($_.Capacity/1gb) * 0.1)),2)}}

# Export to excel #
$hpvClusterOutput | Export-Excel -Path $fileName -WorksheetName $clusSheet -Append -BoldTopRow -AutoFilter -FreezeTopRow -AutoSize
$hpvHostOutput | Export-Excel -Path $fileName -WorksheetName $hostSheet -Append -BoldTopRow -AutoFilter -FreezeTopRow -AutoSize
$hpvDatastoreOutput | Export-Excel -Path $fileName -WorksheetName $storSheet -Append -BoldTopRow -AutoFilter -FreezeTopRow -AutoSize
        
}

Get-date