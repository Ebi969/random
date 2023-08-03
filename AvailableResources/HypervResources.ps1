#Modules
Import-Module -Name ImportExcel
Import-Module -Name VirtualMachineManager

#Variables
$HostGroup = "All Hosts"
$TxConso = 4
$ExcelFile = "D:\UserData\Ibraaheem\Scripts\HyperV\VMM.xlsx"
$ClusSheet = "Cluster"
$StorSheet = "Storage"
$HostSheet = "Host"
$HyperVC = "CLS000288"

Function Get-SCVMInfo{
param (
    $SCVMHostGroup
)
$SCVMHostCluster = Get-SCVMHostCluster -VMHostGroup "All Hosts" #$SCVMHostGroup
    foreach ($Cluster in $SCVMHostCluster) {

        ### CLUSTER INFO ###
        $ClusterResult = "" | Select-Object -Property ClusterName,LogicalCPUCount,AllocatedvCPU,AvailableCPU,CPURatio,NodeCount,TotalMemory,AllocatedvMEM,MemoryRatio,AvailableMemory
        $ClusterResult.ClusterName = $Cluster.Name
        $ClusterResult.LogicalCPUCount = ((($Cluster.Nodes).LogicalCPUCount | Measure-Object -sum).sum)
        $ClusterResult.AvailableCPU = [Math]::Round(($TxConso*(((($Cluster.Nodes).LogicalCPUCount | Measure-Object -sum).sum)/(($Cluster.Nodes).count)))*((($Cluster.Nodes).count)-1), 0)
        $ClusterResult.NodeCount = (($Cluster.Nodes).count)
        $ClusterResult.TotalMemory = [math]::Round(((($Cluster.Nodes).TotalMemory | Measure-Object -sum).sum)/1gb,2)
        $ClusterResult.AvailableMemory = [math]::Round((($Cluster.Nodes).AvailableMemory | Measure-Object -sum).sum/1kb,2)
    
            foreach ($Node in $Cluster.Nodes) {
                $nodeAllocatedvCPU = ((Get-SCVirtualMachine -VMHost $Node | Measure-Object -Property CPUCount -Sum).Sum)
                $nodeAllocatedMem = ((Get-SCVirtualMachine -VMHost $Node | Measure-Object -Property Memory -Sum).Sum/1kb)
                $ClusterResult.AllocatedvCPU += $nodeAllocatedvCPU
                $ClusterResult.AllocatedvMEM += $nodeAllocatedMem

                $HostOutput | Export-Excel -Path $ExcelFile -WorksheetName $HostSheet -AutoSize -Append -BoldTopRow -AutoFilter
            }
    
        $ClusterResult.CPURatio = [math]::Round(($ClusterResult.AllocatedvCPU/$ClusterResult.LogicalCPUCount),2)
        $ClusterResult.MemoryRatio = [math]::Round(($ClusterResult.AllocatedvMem/$ClusterResult.TotalMemory),2)
    
        $ClusterOutput = [pscustomobject] @{
            Cluster = $ClusterResult.ClusterName
            vCenter = "CLS000288"
            TotalCPUinCluster = $ClusterResult.LogicalCPUCount
            CPUallocation = $ClusterResult.AllocatedvCPU
            CPURatio = $ClusterResult.CPURatio
            AvailableCpu = $ClusterResult.AvailableCPU
            TotalMEMinCluster = $ClusterResult.TotalMemory
            MEMallocation = [MATH]::Round($ClusterResult.AllocatedvMEM, 2)
            MemPercAllocated = $ClusterResult.MemoryRatio
            AvailableMem = [MATH]::Round($ClusterResult.AvailableMemory, 2)
        }
        $ClusterOutput | Export-Excel -Path $ExcelFile -WorksheetName $ClusSheet -AutoSize -Append -BoldTopRow -AutoFilter

        ### STORAGE INFO ###

        $ClusterVolumes = $Cluster.SharedVolumes
        
        $StorageOutput = [pscustomobject] @{
            ClusterName	=
            vCenter	= 
            Name = $ClusterVolumes.Name
            FreeSpaceGB	=
            CapacityGB =
            AvailableStorage =
        }
        $StorageOutput | Export-Excel -Path $ExcelFile -WorksheetName $StorSheet -AutoSize -Append -BoldTopRow -AutoFilter
    }
}

Function Get-SCVMStorageInfo{
param (
    $SCVMHostGroup
)
    $SCVMHostCluster = Get-SCVMHostCluster -VMHostGroup $SCVMHostGroup
    foreach ($Cluster in $SCVMHostCluster) {
        $ClusterVolumes = $Cluster.SharedVolumes
        $ClusterVolumes.GetEnumerator() | Select-Object @{n='ClusterName';e={$Cluster.name}},
        @{n='Volume Name';e={$_.Name}},
        @{n='Freespace(GB)';e={"{0:N2}"-f ($_.freespace/1gb)}},
        @{n='Capacity(GB)';e={"{0:N2}"-f ($_.Capacity/1gb)}},
        @{n='Percentage(Available)';e={"{0:N2}" -f ($_.freespace / ($_.Capacity/100))}},
        @{n='CSV-Volume';e={$_.IsClusterSharedVolume}},
        @{n='Available For Placement?';e={$_.IsAvailableForPlacement}},
        @{n='Volume Label';e={$_.VolumeLabel}}
    }
}

Get-SCVMClusterInfo -SCVMHostGroup $HostGroup | Export-Excel -Path $ExcelFile -WorksheetName $ClusSheet -AutoSize -Append -BoldTopRow -AutoFilter
Get-SCVMStorageInfo -SCVMHostGroup $HostGroup  | Export-Excel -Path $ExcelFile -WorksheetName $StorSheet -AutoSize -Append -BoldTopRow -AutoFilter