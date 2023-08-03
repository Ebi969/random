$latestADDMextract = (Get-ChildItem -Path "\\srv003789\ADDMReports" | Sort CreationTime | Select -Last 1)
    $importFile = "\\srv003789\ADDMReports\" + $latestADDMextract
    $importData = Import-CSV $importFile


$clusterToQuery = @("SLM-BDC-GOLD-06-SQL","SLM-CDC-GOLD-06-SQL")

foreach($cluster in $clusterToQuery){

    $vmsInClust = Get-Cluster $cluster | Get-VM | Select -ExpandProperty Name


    foreach($vm in $vmsInClust){
        foreach($row in $importData){
            if($row.Name -eq $vm){
                
                    $export = [pscustomobject]@{
                        Name = $vm
                        Competency = $row.'Competency '
                        OS = $row.'Discovered OS'
                        SLA = $row.sla
                        application = $row.application
                        serverDescription = $row.serverdescription
                        Client = $row.client
                        "App Owner" = $row.appowner
                        PTO = $row.primarytechowner
                        STO = $row.secondarytechowner
                        "IT Exec" = $row.itexec
                        "App Tier" = $row.applicationtier
                    }

                break
            }else{

                $export = [pscustomobject]@{
                        Name = $vn
                        Competency = "N/A"
                        OS = "N/A"
                        SLA = "N/A"
                        application = "N/A"
                        serverDescription = "N/A"
                        Client = "N/A"
                        "App Owner" = "N/A"
                        PTO = "N/A"
                        STO = "N/A"
                        "IT Exec" = "N/A"
                        "App Tier" = "N/A"
                }
            }
        }

        $export
        $export | Export-Excel -Path "D:\UserData\Ibraaheem\Scripts\VMWare\vmDetailsUsingCluster\sqlClusterVMs.xlsx" -WorksheetName $cluster -Append -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow

    }
}