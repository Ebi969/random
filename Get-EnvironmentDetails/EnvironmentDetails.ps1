$path = "D:\UserData\Ibraaheem\Scripts\VMWare\Get-EnvironmentDetails"
$serverList = Get-Content "$path\serverList.txt"

$vmList = @()

foreach($vm in $serverList){

    Try{
        $vmDetail = Get-VM $vm -ErrorAction Stop
        $Data = ($vmDetail | Get-Datastore) -join ", "
            $vmOutput = [pscustomobject] @{
                "vmName" = $vmDetail.Name
                "Cluster" = $vmDetail | Get-Cluster | Select -ExpandProperty Name
                "vCPU" = $vmDetail.numCPU
                "Memory" = $vmDetail.MemoryGB
                "Disk Space" = [MATH]::Round($vmDetail.ProvisionedSpaceGB, 2)
                "Datastores" = $Data
            }

    }catch{
            $vmOutput = [pscustomobject] @{
                "vmName" = $vm
                "Cluster" = "Not in VMware"
                "vCPU" = ""
                "Memory" = ""
                "Disk Space" = ""
            }
    }

    $vmList += $vmOutput

}

$vmList | Export-Excel -Path "$path\VMwareGlacierOS.xlsx" -AutoSize -BoldTopRow -AutoFilter