$path = "D:\UserData\Ibraaheem\Scripts\VMWare\RevisedMigrationProjects"
$actionPlanPath = "$path\ActionPlans\VDI"

$allFiles = Get-ChildItem $actionPlanPath\*.xlsx

foreach($clusterFile in $allFiles){
    $fileName = $clusterFile.Name
    
    $NMPShared = $null
    $MSPShared = $null
    $MAPShared = $null
    $VDIShared = $null

    Try{$NMPShared = Import-Excel "$actionPlanPath\$fileName" -WorksheetName "NML-Shared" -ErrorAction Stop}catch{}
    Try{$MSPShared = Import-Excel "$actionPlanPath\$fileName" -WorksheetName "MSP-Shared" -ErrorAction Stop}catch{}
    Try{$MAPShared = Import-Excel "$actionPlanPath\$fileName" -WorksheetName "MAP-Shared" -ErrorAction Stop}catch{}
    Try{$VDIShared = Import-Excel "$actionPlanPath\$fileName" -WorksheetName "VDI-Shared" -ErrorAction Stop}catch{}
    Try{$lunRequests = Import-Excel "$actionPlanPath\$fileName" -WorksheetName "Lun Requests" -ErrorAction Stop}catch{}
    
    $sharedNMPLuns = $lunRequests | Where-Object {$_.Type -eq "NML" -and $_.ForVM -eq "Shared"}
    $sharedMSPLuns = $lunRequests | Where-Object {$_.Type -eq "MSP" -and $_.ForVM -eq "Shared"}
    $sharedMAPLuns = $lunRequests | Where-Object {$_.Type -eq "MAP" -and $_.ForVM -eq "Shared"}
    $sharedVDILuns = $lunRequests | Where-Object {$_.Type -eq "VDI" -and $_.ForVM -eq "Shared"}

    <### NMP SECTION ###>
    if($NMPShared){
        $newNMPDatastores = @()
        $newNMPTab = @()
        $count = 1
        foreach($newLun in $sharedNMPLuns){
            $newLun = "Datastore" + $count
            $newNMPDatastores += $newLun
            $count += 1
        }
        $newVMDetails = @()
        foreach($vm in $NMPShared | Sort-Object 'Migrate(GB)' -Descending){
            $vmSize = $vm.'Migrate(GB)'
            $datNum = 0
            $possibleAssignTo = $newNMPDatastores[$datNum]
            $capacityOfPossibleDS = ($newVMDetails | Where-Object {$_.'AssignTo' -eq $possibleAssignTo} | Measure-Object 'Migrate(GB)' -Sum).Sum + $vmSize

            while(($capacityOfPossibleDS/4096) -gt 0.85){
                $datNum += 1
                $possibleAssignTo = $newNMPDatastores[$datNum]
                $capacityOfPossibleDS = ($newVMDetails | Where-Object {$_.'AssignTo' -eq $possibleAssignTo} | Measure-Object 'Migrate(GB)' -Sum).Sum + $vmSize
            }

            $finalAssignment = $possibleAssignTo
            $vm.VMName + " to be assigned to " + $finalAssignment
            $newVMDetails += $vm | Select-Object *, @{n="AssignTo"; e={$finalAssignment}}

        }
        $newNMPTab = $newVMDetails
        $newNMPTab | Export-Excel -Path "$actionPlanPath\$fileName" -WorksheetName "NML-Shared" -BoldTopRow -FreezeTopRow -AutoFilter
    }
    <### VDI SECTION ###>
    if($VDIShared){
        $newVDIDatastores = @()
        $newVDITab = @()
        $count = 1
        foreach($newLun in $sharedVDILuns){
            $newLun = "Datastore" + $count
            $newVDIDatastores += $newLun
            $count += 1
        }
        $newVMDetails = @()
        foreach($vm in $VDIShared | Sort-Object 'Migrate(GB)' -Descending){
            $vmSize = $vm.'Migrate(GB)'
            $datNum = 0
            $possibleAssignTo = $newVDIDatastores[$datNum]
            $capacityOfPossibleDS = ($newVMDetails | Where-Object {$_.'AssignTo' -eq $possibleAssignTo} | Measure-Object 'Migrate(GB)' -Sum).Sum + $vmSize

            while(($capacityOfPossibleDS/4096) -gt 0.85){
                $datNum += 1
                $possibleAssignTo = $newVDIDatastores[$datNum]
                $capacityOfPossibleDS = ($newVMDetails | Where-Object {$_.'AssignTo' -eq $possibleAssignTo} | Measure-Object 'Migrate(GB)' -Sum).Sum + $vmSize
            }

            $finalAssignment = $possibleAssignTo
            $vm.VMName + " to be assigned to " + $finalAssignment
            $newVMDetails += $vm | Select-Object *, @{n="AssignTo"; e={$finalAssignment}}

        }
        $newVDITab = $newVMDetails
        $newVDITab | Export-Excel -Path "$actionPlanPath\$fileName" -WorksheetName "VDI-Shared" -BoldTopRow -FreezeTopRow -AutoFilter
    }

    <### MSP SECTION ###>
    if($MSPShared){
        $newMSPDatastores = @()
        $newMSPTab = @()
        $count = 1
        foreach($newLun in $sharedMSPLuns){
            $newLun = "Datastore" + $count
            $newMSPDatastores += $newLun
            $count += 1
        }
        $newVMDetails = @()
        foreach($vm in $MSPShared | Sort-Object 'Migrate(GB)' -Descending){
            $vmSize = $vm.'Migrate(GB)'
            $datNum = 0
            $possibleAssignTo = $newMSPDatastores[$datNum]
            $capacityOfPossibleDS = ($newVMDetails | Where-Object {$_.'AssignTo' -eq $possibleAssignTo} | Measure-Object 'Migrate(GB)' -Sum).Sum + $vmSize

            while(($capacityOfPossibleDS/4096) -gt 0.85){
                $datNum += 1
                $possibleAssignTo = $newMSPDatastores[$datNum]
                $capacityOfPossibleDS = ($newVMDetails | Where-Object {$_.'AssignTo' -eq $possibleAssignTo} | Measure-Object 'Migrate(GB)' -Sum).Sum + $vmSize
            }

            $finalAssignment = $possibleAssignTo
            $vm.VMName + " to be assigned to " + $finalAssignment
            $newVMDetails += $vm | Select-Object *, @{n="AssignTo"; e={$finalAssignment}}
        }
        $newMSPTab = $newVMDetails
        $newMSPTab | Export-Excel -Path "$actionPlanPath\$fileName" -WorksheetName "MSP-Shared" -BoldTopRow -FreezeTopRow -AutoFilter
    }

    <### MAP SECTION ###>
    if($MAPShared){
        $newMAPDatastores = @()
        $newMAPTAB = @()
        $count = 1
        foreach($newLun in $sharedMAPLuns){
            $newLun = "Datastore" + $count
            $newMAPDatastores += $newLun
            $count += 1
        }
        $newVMDetails = @()
        foreach($vm in $MAPShared | Sort-Object 'Migrate(GB)' -Descending){
            $vmSize = $vm.'Migrate(GB)'
            $datNum = 0
            $possibleAssignTo = $newMAPDatastores[$datNum]
            $capacityOfPossibleDS = ($newVMDetails | Where-Object {$_.'AssignTo' -eq $possibleAssignTo} | Measure-Object 'Migrate(GB)' -Sum).Sum + $vmSize

            while(($capacityOfPossibleDS/4096) -gt 0.85){
                $datNum += 1
                $possibleAssignTo = $newMAPDatastores[$datNum]
                $capacityOfPossibleDS = ($newVMDetails | Where-Object {$_.'AssignTo' -eq $possibleAssignTo} | Measure-Object 'Migrate(GB)' -Sum).Sum + $vmSize
            }

            $finalAssignment = $possibleAssignTo
            $vm.VMName + " to be assigned to " + $finalAssignment
            $newVMDetails += $vm | Select-Object *, @{n="AssignTo"; e={$finalAssignment}}
        }
        $newMAPTAB = $newVMDetails
        $newMAPTAB | Export-Excel -Path "$actionPlanPath\$fileName" -WorksheetName "MAP-Shared" -BoldTopRow -FreezeTopRow -AutoFilter
    }
}