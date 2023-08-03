$path = "D:\UserData\Ibraaheem\Scripts\VMWare\DatT3Low"

$datastoreLowDetails = Get-Datastore | Where {$_.Name -match "AMB|RUD|T3|EME|QUA"} | Foreach{                    
                        
                        $clusterNames = $null

                        if($_.'CapacityRequiredFor15%' -notlike "*-*"){
                        
                            $vmClusters = Get-Datastore $_.Name | Get-VMHost | Get-Cluster
                            $clusterNames = $vmClusters -join ", "

                            if($vmClusters.Length -ne 1){
                                $multiStore = "Yes"
                            }else{
                                $multiStore = "No"
                            }
                            
                            $output = [pscustomobject] @{
                                'DS Name' = $_.Name
                                'Cluster' = $clusterNames
                                'MultiCluster' = $multiStore
                                'FreeSpaceGB' = [Math]::Round($_.FreeSpaceGB, 2)
                                'CapacityGB' = [Math]::Round($_.CapacityGB, 2)
                                'FreePercentage' = [Math]::Round((($_.FreeSpaceGB/$_.CapacityGB)*100), 2)
                                ' ' = ""
                                '10%ofCap' = [Math]::Round((($_.CapacityGB * 0.1)), 2)
                                'CapacityRequiredFor10%' = [Math]::Round((($_.CapacityGB * 0.1) - $_.FreeSpaceGB), 2)
                                '  ' = ""
                                '15%ofCap' = [Math]::Round((($_.CapacityGB * 0.15)), 2)
                                'CapacityRequiredFor15%' = [Math]::Round((($_.CapacityGB * 0.15) - $_.FreeSpaceGB), 2)
                            }

                            $output | Export-Excel -Path $path\DatT3Low.xlsx -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow -Append

                        }
                     }


                     