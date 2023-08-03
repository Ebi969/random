$latestADDMextract = (Get-ChildItem -Path "\\srv003789\ADDMReports" | Sort CreationTime | Select -Last 1)
    $importFile = "\\srv003789\ADDMReports\" + $latestADDMextract
    $importData = Import-CSV $importFile

foreach ($row in $importData){

    $row.Name
    $row.sla
    $row.applicationtier
    $row.primarytechowner
    $row.appowner
    $row.serverdescription

    Try{

        $fullname = $row.Name

        $nameSplit = $fullname.Split(".")

        $dsAssingedperVM = Get-VM $nameSplit[0] -ErrorAction Stop | Get-Datastore | Select -ExpandProperty Name 

            foreach($datastore in $dsAssingedperVM){

                if($datastore -like "*SAP*" -or $datastore -like "*DIA*"){

                    $storageTier = "Tier 2"                    
                    break
                }else{
                    $storageTier = "Tier 3"
                }
                
            }

        $out = New-Object psobject

        $out | Add-Member -MemberType NoteProperty -Name "Name" -Value $nameSplit[0]
        $out | Add-Member -MemberType NoteProperty -Name "PTO" -Value $row.primarytechowner
        $out | Add-Member -MemberType NoteProperty -Name "App Owner" -Value $row.appowner
        $out | Add-Member -MemberType NoteProperty -Name "Server Description" -Value $row.serverdescription
        $out | Add-Member -MemberType NoteProperty -Name "SLA" -Value $row.sla
        $out | Add-Member -MemberType NoteProperty -Name "App Tier" -Value $row.applicationtier
        $out | Add-Member -MemberType NoteProperty -Name "Storage Tier" -Value $storageTier

        $out | Export-Excel -Path "D:\UserData\Ibraaheem\Scripts\VMWare\addmCheckVMWareStorageTier\checks.xlsx" -WorksheetName "VMWare" -Append -AutoSize -AutoFilter -BoldTopRow

    }catch{
        
        $storageTier = "Server is not in VMWare"

        $out = New-Object psobject

        $out | Add-Member -MemberType NoteProperty -Name "Name" -Value $nameSplit[0]
        $out | Add-Member -MemberType NoteProperty -Name "PTO" -Value $row.primarytechowner
        $out | Add-Member -MemberType NoteProperty -Name "App Owner" -Value $row.appowner
        $out | Add-Member -MemberType NoteProperty -Name "Server Description" -Value $row.serverdescription
        $out | Add-Member -MemberType NoteProperty -Name "SLA" -Value $row.sla
        $out | Add-Member -MemberType NoteProperty -Name "App Tier" -Value $row.applicationtier
        $out | Add-Member -MemberType NoteProperty -Name "Storage Tier" -Value $storageTier

        $out | Export-Excel -Path "D:\UserData\Ibraaheem\Scripts\VMWare\addmCheckVMWareStorageTier\checks.xlsx" -WorksheetName "Non VMWare" -Append -AutoSize -AutoFilter -BoldTopRow

    }    
    
}