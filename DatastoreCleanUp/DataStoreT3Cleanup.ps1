$finalvmOutput = @()
$finaldsOutput = @()
$finalvmDiskOutput = @()

$outputPath = "D:\UserData\Ibraaheem\Scripts\VMWare\DatastoreCleanUp"
$timeStamp = Get-Date -Format yyyy-MM-dd
$exportFileName = $outputpath + "\DatastoreDetails_" + $timeStamp + ".xlsx"

foreach($cluster in Get-Cluster){

    $clusterName = $cluster.Name

    foreach($datastore in ($cluster | Get-Datastore | Where {$_.Name -match "AMB|RUB|T3|EME|QUA" -and $_.Name -notmatch "stage|pamco|local|scratch"})){
        
        $datastoreName = $datastore.Name

            $dsCapacity = [MATH]::Round($datastore.CapacityGB, 2)
            $dsFreeSpace = [MATH]::Round($datastore.FreeSpaceGB, 2)
            $dsPercentFree = [MATH]::Round(($dsFreeSpace/$dsCapacity) * 100, 2)
            $ds15Perc = [MATH]::Round((0.15)*$dsCapacity, 2)
            $dsPercRequiredTo15 = [MATH]::Round($ds15Perc - $dsPercentFree, 2)
            $dsGBRequired = [Math]::Round(($dsPercRequiredTo15 - $dsFreeSpace), 2)

            if($dsGBRequired -lt 0){
                $dsGBRequired = "No cleanup needed"
            }

            ##############################################
            # Per VMs
            ##############################################

                    foreach($vm in ($datastore | Get-VM)){

                        $allDsVM = $vm | Get-Datastore | Where {$_.Name -match "AMB|RUB|T3|EME|QUA" -and $_.Name -notmatch "stage|pamco|local|scratch"}

                            if($allDsVM.length -gt 1){
                                $singleDS = $false
                            }else{
                                $singleDS = $true
                            }
                            
                        $vmName = $vm.Name

                            $vmTotal = $null

                            foreach($vmDisk in $vm | Get-HardDisk){
                                if($vmDisk.FileName -match "AMB|RUB|T3|EME|QUA" -and $vmDisk.FileName -notmatch "stage|pamco|local|scratch"){
                                    $vmTotal += [MATH]::Round($vmDisk.CapacityGB, 2)

                                        $vmDiskOutput = [pscustomobject] @{
                                            "Cluster" = $clusterName
                                            "VM" = $vmName
                                            "Disk Capacity Used" = $vMDisk.CapacityGB
                                            "DiskName" = $vmDisk.Name
                                            "FullName" = $vmDisk.FileName
                                        }

                                        $finalvmDiskOutput += $vmDiskOutput

                                }
                            }

                        $vmDetail = [pscustomobject] @{
                            "Cluster" = $clusterName
                            "Datastore" = $datastoreName
                            "DS Capacity(GB)" = $dsCapacity
                            "DS GB required" = $dsGBRequired
                            "VM" = $vmName
                            "VM Capacity Used" = $vmTotal
                            "SingleDS" = $singleDS
                        }

                        $finalvmOutput += $vmDetail
                    }


            ##############################################

        $dsDetail = [pscustomobject] @{
            "Cluster" = $clusterName
            "Datastore" = $datastoreName
            "Capacity(GB)" = $dsCapacity
            "FreeSpace(GB)" = $dsFreeSpace
            "% Free" = $dsPercentFree
            "15% of Cap" = $ds15Perc
            "(GB)required" = $dsGBRequired
        }

        $finaldsOutput += $dsDetail
    }    
}

$finaldsOutput | Export-Excel -Path $exportFileName -WorksheetName "Datastores" -Append -FreezeTopRow -BoldTopRow -AutoSize -AutoFilter
$finalvmOutput | Export-Excel -Path $exportFileName -WorksheetName "VM Detail" -Append -FreezeTopRow -BoldTopRow -AutoSize -AutoFilter
$finalvmDiskOutput | Export-Excel -Path $exportFileName -WorksheetName "VM Disk Detail" -Append -FreezeTopRow -BoldTopRow -AutoSize -AutoFilter