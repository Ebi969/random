$path = "D:\UserData\Ibraaheem\Scripts\VMWare\vmT2onDSlist"
$allDataStores = Get-Content -Path $path\dsList.txt

foreach($ds in $allDataStores){
    
    $stores = $null
    $vmsOnDS = $null

    Try{
    $stores = Get-Datastore $ds -ErrorAction Stop    
        $stores.Name
    $vmsOnDS = Get-VM -Datastore $stores.Name

    if($vmsOnDS.length -lt 1){
        
        $vm = "No VMs on DS"
        $totT2 = "0"

        $storeNameSplit = ($stores.Name).Split("-")
            if($storeNameSplit[0] -like "*01"){          


                $out = New-Object psobject

                    $out | Add-Member -MemberType NoteProperty -Name "DS" -Value $ds
                    $out | Add-Member -MemberType NoteProperty -Name "VM" -Value $vm
                    $out | Add-Member -MemberType NoteProperty -Name "T2 Size" -Value $totT2

                $out | Export-Excel -path $path\OutputList.xlsx -Append -AutoSize -BoldTopRow -AutoFilter -WorksheetName "BDC"
            
            }elseif($storeNameSplit[0] -like "*02"){
                
                $out = New-Object psobject

                    $out | Add-Member -MemberType NoteProperty -Name "DS" -Value $ds
                    $out | Add-Member -MemberType NoteProperty -Name "VM" -Value $vm
                    $out | Add-Member -MemberType NoteProperty -Name "T2 Size" -Value $totT2

                $out | Export-Excel -path $path\OutputList.xlsx -Append -AutoSize -BoldTopRow -AutoFilter -WorksheetName "CDC"
            }    

    }
        
        foreach($vm in $vmsOnDS){
        
        $totT2 = $null
        $vm

            foreach($vmDisk in ($vm | Get-HardDisk)){
            
                $datastoreDisk = ($vmDisk.Filename).Split(" ")      
            
                $datastore = $datastoreDisk[0].Replace("[","")
                $datastore = $datastore.Replace("]","")

                if($datastore -like '*SAP*' -or $datastore -like '*DIA*'){

                    $totT2 += $vmDisk.CapacityGB

                }

            }  
            
            $storeNameSplit = ($stores.Name).Split("-")
            if($storeNameSplit[0] -like "*01"){          


                $out = New-Object psobject

                    $out | Add-Member -MemberType NoteProperty -Name "DS" -Value $ds
                    $out | Add-Member -MemberType NoteProperty -Name "VM" -Value $vm.Name
                    $out | Add-Member -MemberType NoteProperty -Name "T2 Size" -Value $totT2

                $out | Export-Excel -path $path\OutputList.xlsx -Append -AutoSize -BoldTopRow -AutoFilter -WorksheetName "BDC"
            
            }elseif($storeNameSplit[0] -like "*02"){
                
                $out = New-Object psobject

                    $out | Add-Member -MemberType NoteProperty -Name "DS" -Value $ds
                    $out | Add-Member -MemberType NoteProperty -Name "VM" -Value $vm.Name
                    $out | Add-Member -MemberType NoteProperty -Name "T2 Size" -Value $totT2

                $out | Export-Excel -path $path\OutputList.xlsx -Append -AutoSize -BoldTopRow -AutoFilter -WorksheetName "CDC"
            }        
                       
        }
        

    }catch{

        $vm = "DataStore not found"
        $totT2 = "0"

        $dsSplit = $ds.Split("-")

        if($dsSplit[0] -like "*01"){

            $out = New-Object psobject

                $out | Add-Member -MemberType NoteProperty -Name "DS" -Value $ds
                $out | Add-Member -MemberType NoteProperty -Name "VM" -Value $vm
                $out | Add-Member -MemberType NoteProperty -Name "T2 Size" -Value $totT2

            $out | Export-Excel -path $path\OutputList.xlsx -Append -AutoSize -BoldTopRow -AutoFilter -WorksheetName "BDC"      

        }elseif($dsSplit[0] -like "*02"){

            $out = New-Object psobject

                $out | Add-Member -MemberType NoteProperty -Name "DS" -Value $ds
                $out | Add-Member -MemberType NoteProperty -Name "VM" -Value $vm
                $out | Add-Member -MemberType NoteProperty -Name "T2 Size" -Value $totT2

            $out | Export-Excel -path $path\OutputList.xlsx -Append -AutoSize -BoldTopRow -AutoFilter  -WorksheetName "CDC"
                
        }
    }
}