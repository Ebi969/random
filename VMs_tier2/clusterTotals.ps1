$outputPath = "D:\UserData\Ibraaheem\Scripts\VMWare\VMs_tier2"

$allClusters = Get-Cluster

$totNMP = $null
$totMSP = $null
$totMAP = $null
$blankcol = ""


if(Test-Path $outputPath\clusterTotals.xlsx){
    Remove-Item $outputPath\clusterTotals.xlsx
}

foreach($cluster in $allClusters){

$cluster.Name

    if(($cluster|Get-VM).count -gt 0){

    $clusterTotNMP = $null
    $clusterTotMSP = $null
    $clusterTotMAP = $null
    
    foreach($vm in $cluster | Get-VM){

    $CountT2Luns = 0

    $CountT2Luns = ($vm|Get-Datastore|Where {$_.Name.Contains("-SAP-") -or $_.Name.Contains("-DIA-")}).count

    $vm.Name
    $vmTotT2 = 0
                
        foreach($vmDisk in $vm | Get-HardDisk){
                        
                $datastoreDisk = ($vmDisk.Filename).Split(" ")      
            
                $datastore = $datastoreDisk[0].Replace("[","")
                $datastore = $datastore.Replace("]","")
                
                if(($datastore -like '*SAP*' -or $datastore -like '*DIA*') -and $datastore -like '*NMP*' ){

                    $dataType = "T2 NMP"
                    $vmTotT2 += $vmDisk.CapacityGB
                    $clusterTotNMP += $vmDisk.CapacityGB
                
                }elseif(($datastore -like '*SAP*' -or $datastore -like '*DIA*') -and $datastore -like '*MSP*'){

                    $dataType = "T2 MSP"
                    $vmTotT2 += $vmDisk.CapacityGB
                    $clusterTotMSP += $vmDisk.CapacityGB

                }elseif(($datastore -like '*SAP*' -or $datastore -like '*DIA*') -and $datastore -like '*MAP*'){
            
                    $dataType = "T2 MAP"
                    $vmTotT2 += $vmDisk.CapacityGB
                    $clusterTotMAP += $vmDisk.CapacityGB
            
                }else{                   
                    $dataType = "Not T2"
                }

                if($vmTotT2 -ge 3500){
                    $ownDat = "over 3.5TB"
                }else{
                    $ownDat = "below 3.5TB"
                }


                $output = New-Object psobject
            
                $output | Add-Member -MemberType NoteProperty -Name "Cluster" -Value $cluster.Name
                $output | Add-Member -MemberType NoteProperty -Name "VM" -Value $vm.Name   
                $output | Add-Member -MemberType NoteProperty -Name "VM Disk" -Value $vmDisk.Name
                $output | Add-Member -MemberType NoteProperty -Name "VM Disk Size" -Value $vmDisk.CapacityGB
                $output | Add-Member -MemberType NoteProperty -Name "Datastore" -Value $datastore

                $output | Export-Excel -Path $outputPath\clusterTotals.xlsx -BoldTopRow -AutoFilter -WorksheetName "VM hard disk Details" -Append           
        } 

        foreach($vmDataStore in $vm|Get-Datastore){      
        
            if($vmDataStore.Name.Contains("-SAP-") -or $vmDataStore.Name.Contains("-DIA-")){
                            
                $output = New-Object psobject
            
                $output | Add-Member -MemberType NoteProperty -Name "Cluster" -Value $cluster.Name
                $output | Add-Member -MemberType NoteProperty -Name "VM" -Value $vm.Name   
                $output | Add-Member -MemberType NoteProperty -Name "Datastore" -Value $vmDataStore.Name        
                $output | Add-Member -MemberType NoteProperty -Name "Total T2 Size" -Value ([MATH]::Round($vmTotT2,2))                
                $output | Add-Member -MemberType NoteProperty -Name "Over 3.5TB?" -Value $ownDat
                $output | Add-Member -MemberType NoteProperty -Name "T2 Luns on VM" -Value $CountT2Luns

                $output | Export-Excel -Path $outputPath\clusterTotals.xlsx -BoldTopRow -AutoFilter -WorksheetName "VM Tier 2 Details" -Append
            }
        }

    }
          
          if(!(($clusterTotNMP -eq $null) -and ($clusterTotNMP -eq $null) -and ($clusterTotNMP -eq $null))){
           
                $output = New-Object psobject
            
                $output | Add-Member -MemberType NoteProperty -Name "Cluster" -Value $cluster.Name
                $output | Add-Member -MemberType NoteProperty -Name "Total NMP" -Value ([MATH]::Round($clusterTotNMP,2))
                $output | Add-Member -MemberType NoteProperty -Name "Total MSP" -Value ([MATH]::Round($clusterTotMSP,2))
                $output | Add-Member -MemberType NoteProperty -Name "Total MAP" -Value ([MATH]::Round($clusterTotMAP,2))
    
                $output | Export-Excel -Path $outputPath\clusterTotals.xlsx -BoldTopRow -AutoFilter -WorksheetName "ClusterTotals" -Append
          }              
    }
}
