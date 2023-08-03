$outputPath = "D:\Reports\VMware\AvailableResources"

if(Test-Path $outputPath\AvailableResources.xlsx){
    Remove-Item $outputPath\AvailableResources.xlsx -Force
}

# Get Cred store
$creds7281 = Get-VICredentialStoreItem -file  C:\Users\svcVMWareScriptAcc\CredStore7281.creds
$creds7282 = Get-VICredentialStoreItem -file  C:\Users\svcVMWareScriptAcc\CredStore7282.creds

# Connect-VIServer
Connect-VIServer -Server SRV007281.mud.internal.co.za -User $Creds7281.User -Password $Creds7281.Password
Connect-VIServer -Server SRV007282.mud.internal.co.za -User $Creds7282.User -Password $Creds7282.Password

$clusters = Get-Cluster | Where {$_.Name -Match "06-SQL|CCC|BRONZE-03-POC|BRONZE-04-Windows|GOLD-03-Linux|GOLD-04-WAS|GOLD-04-Windows|GOLD-07-BACKUP|GOLD-01-BI|GOLD-09-SQL-GLD|HP-03-PM|HP-08-Database|BRONZE-01-Shared|BRONZE-02-POC|GOLD-01-Windows|GOLD-02-Linux|GOLD-05-SQL|HP-01-Database"}

$collect = @()
foreach($cluster in $clusters){

$vmHosts = $cluster | Get-VMHost

if(!($vmHosts)){
continue 
}

$vms = $cluster | Get-VM

$totalCPU = ($vmHosts | Measure-Object NumCpu -Sum).Sum
$totalMemGB = [MATH]::Round(($vmHosts | Measure-Object MemoryTotalGB -Sum).Sum ,2)

$vmCPUs = ($vms | Measure-Object NumCpu -Sum).Sum
$vmMem = ($vms | Measure-Object MemoryGB -Sum).Sum

$cpuRatio = [MATH]::Round(($vmCPUs/$totalCPU), 2)
$memPerc = ($vmMem/$totalMemGB).ToString("P")

$toBeLimitMem = (0.8 * $totalMemGB) - $vmMem

################
# Switch Cases #
################

switch -Wildcard ($cluster.Name){

#######
# SLM #
#######

       "*06-SQL*" {      
            $toBeLimitCpu = ($totalCPU*2) - $vmCPUs
            $toBeLimitMem = (0.9 * $totalMemGB) - $vmMem
       }

       "*CCC*" {
            $toBeLimitCpu = $totalCPU - $vmCPUs
       }

       "*BRONZE-03-POC*" {      
            $toBeLimitCpu = ($totalCPU*6) - $vmCPUs   
       }
   
       "*BRONZE-04-Windows*" {      
            $toBeLimitCpu = ($totalCPU*6) - $vmCPUs
       }

       "*GOLD-03-Linux*" {      
            $toBeLimitCpu = ($totalCPU*4) - $vmCPUs
       }

       "*GOLD-04-WAS*" {      
            $toBeLimitCpu = ($totalCPU*4) - $vmCPUs
       }

       "*GOLD-04-Windows*" {      
            $toBeLimitCpu = ($totalCPU*4) - $vmCPUs
       }

       "*GOLD-07-BACKUP*" {      
            $toBeLimitCpu = ($totalCPU*4) - $vmCPUs
       }
   
       "*GOLD-01-BI*" {      
            $toBeLimitCpu = ($totalCPU*4) - $vmCPUs
       }

       "*GOLD-09-SQL-GLD*" {      
            $toBeLimitCpu = ($totalCPU*2) - $vmCPUs
            $toBeLimitMem = (0.9 * $totalMemGB) - $vmMem
       }

       "*HP-03-PM*" {      
            $toBeLimitCpu = ($totalCPU*4) - $vmCPUs
       }

       "*HP-08-Database*" {      
            $toBeLimitCpu = ($totalCPU*4) - $vmCPUs
       }

#######
# STM #
#######

        "*BRONZE-01-Shared*" {      
            $toBeLimitCpu = ($totalCPU*4) - $vmCPUs
       }

       "*BRONZE-02-POC*" {      
            $toBeLimitCpu = ($totalCPU*6) - $vmCPUs
       }

       "*GOLD-01-Windows*" {      
            $toBeLimitCpu = ($totalCPU*4) - $vmCPUs
       }

       "*GOLD-02-Linux*" {      
            $toBeLimitCpu = ($totalCPU*4) - $vmCPUs
       }

       "*GOLD-05-SQL*" {      
            $toBeLimitCpu = ($totalCPU*2) - $vmCPUs
       }

       "*HP-01-Database*" {      
            $toBeLimitCpu = ($totalCPU*2) - $vmCPUs
       }
}


    $rowOutput = [pscustomobject] @{
        Cluster = $cluster.Name
        TotalCPUinCluster = $totalCPU
        CPUallocation = $vmCPUs
        CPURatio = $cpuRatio
        AvailableCpu = $toBeLimitCpu
        TotalMEMinCluster = $totalMemGB
        MEMallocation = $vmMem
        MemPercUsed = $memPerc
        AvailableMem = $toBeLimitMem
    }

$collect += $rowOutput

}

$collect | Sort-Object Name | Export-Excel $outputPath\AvailableResources.xlsx -WorksheetName "AvailableResources" -BoldTopRow -AutoSize -AutoFilter