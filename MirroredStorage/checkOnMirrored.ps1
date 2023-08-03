$path = "D:\UserData\Ibraaheem\Scripts\VMWare\MirroredStorage"

$latestADDMextract = (Get-ChildItem -Path "\\srv003789\ADDMReports" | Sort CreationTime | Select -Last 1)
    $importFile = "\\srv003789\ADDMReports\" + $latestADDMextract
    $importData = Import-CSV $importFile


Function LoopThrough{
    foreach($row in $importData){

                $server = $row.Name
                $app = $row.'applicationtier'

                Get_VM_lunData -vm $server -appTiering $app -pto $row.primarytechowner -appOwner $row.appowner -desc $row.serverdescription -sla $row.sla        
    }
}

Function Get_VM_lunData{
Param(
    $vm,
    $pto,
    $appOwner,
    $desc,    
    $sla,
    $appTiering
)

    $nameSplit = $vm.Split(".")
    $shortName = $nameSplit[0]
    $vmNoSpaces = $shortName.Replace(" ","")

    $vmNoSpaces
    $appTiering

        Try{
            $dataStores = Get-VM $vmNoSpaces -ErrorAction Stop | Get-Datastore
            $vmCluster = Get-VM $vmNoSpaces | Get-Cluster | Select -ExpandProperty Name

            foreach($ds in $dataStores){
                $isMirrored = $null
                    if(($ds.Name) -like "*MSP*"){
                        $isMirrored = "Mirrored"
                        break
                    }else{
                        $isMirrored = "Not Mirrored"
                    }
            }

            $dataStoresJoin = $dataStores -join ","

            export_Data_VMWare -vmName $vmNoSpaces -vmCluster $vmCluster -appTier $appTiering -dataStoreName $dataStoresJoin -Mirrored $isMirrored -pto $pto -appOwner $appOwner -desc $desc -sla $sla
        }catch{            
            export_Data_NonVMWare -vmName $vmNoSpaces -appTier $appTiering -pto $pto -appOwner $appOwner -desc $desc -sla $sla
        }

}

Function export_Data_VMWare{
Param(
    [String] $vmName,
    $vmCluster,
    $pto,
    $appOwner,
    $desc,    
    $sla,
    $appTier,
    [String] $dataStoreName,
    [String] $Mirrored
)

    if($Mirrored -eq "Not Mirrored"){
        $anyIssues = "Issue over Here"
    }else{
        $anyIssues = ""
    }

    if($appTier -eq "AT3" -or $appTier -eq "ATIII" -or $appTier -eq "III"){

    $out = New-Object psobject

    $out | Add-Member -MemberType NoteProperty -Name "VM Name" -Value $vmName
    $out | Add-Member -MemberType NoteProperty -Name "VM Host" -Value $vmCluster
    $out | Add-Member -MemberType NoteProperty -Name "PTO" -Value $pto
    $out | Add-Member -MemberType NoteProperty -Name "App Owner" -Value $appOwner
    $out | Add-Member -MemberType NoteProperty -Name "Server Description" -Value $desc
    $out | Add-Member -MemberType NoteProperty -Name "SLA" -Value $sla
    $out | Add-Member -MemberType NoteProperty -Name "AppTier" -Value $appTier
    $out | Add-Member -MemberType NoteProperty -Name "DS Name" -Value $dataStoreName
    $out | Add-Member -MemberType NoteProperty -Name "Mirrored" -Value $Mirrored
    $out | Add-Member -MemberType NoteProperty -Name "Issues" -Value $anyIssues

    $out | Export-Excel -Path $path\AppTier3_Mirrored.xlsx -WorksheetName "ADDMTier3" -Append -BoldTopRow -AutoSize -AutoFilter

    }else{
    

    $out = New-Object psobject

    $out | Add-Member -MemberType NoteProperty -Name "VM Name" -Value $vmName
    $out | Add-Member -MemberType NoteProperty -Name "VM Host" -Value $vmCluster
    $out | Add-Member -MemberType NoteProperty -Name "PTO" -Value $pto
    $out | Add-Member -MemberType NoteProperty -Name "App Owner" -Value $appOwner
    $out | Add-Member -MemberType NoteProperty -Name "Server Description" -Value $desc
    $out | Add-Member -MemberType NoteProperty -Name "SLA" -Value $sla
    $out | Add-Member -MemberType NoteProperty -Name "AppTier" -Value $appTier
    $out | Add-Member -MemberType NoteProperty -Name "DS Name" -Value $dataStoreName
    $out | Add-Member -MemberType NoteProperty -Name "Mirrored" -Value $Mirrored

    $out | Export-Excel -Path $path\AppTier3_Mirrored.xlsx -WorksheetName "ADDMNotTier3" -Append -BoldTopRow -AutoSize -AutoFilter
    
    }
}

Function export_Data_NonVMWare{
Param(
    [String] $vmName,
    $pto,
    $appOwner,
    $desc,    
    $sla,
    $appTier
)

    $anyIssues = "Not in VMWare"

    $out = New-Object psobject

    $out | Add-Member -MemberType NoteProperty -Name "VM Name" -Value $vmName
    $out | Add-Member -MemberType NoteProperty -Name "PTO" -Value $pto
    $out | Add-Member -MemberType NoteProperty -Name "App Owner" -Value $appOwner
    $out | Add-Member -MemberType NoteProperty -Name "Server Description" -Value $desc
    $out | Add-Member -MemberType NoteProperty -Name "SLA" -Value $sla
    $out | Add-Member -MemberType NoteProperty -Name "AppTier" -Value $appTier
    $out | Add-Member -MemberType NoteProperty -Name "Issues" -Value $anyIssues

    $out | Export-Excel -Path $path\AppTier3_Mirrored.xlsx -WorksheetName "Non VMWare" -Append -BoldTopRow -AutoSize -AutoFilter

}

LoopThrough