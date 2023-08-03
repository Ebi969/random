$latestADDMextract = (Get-ChildItem -Path "\\srv003789\ADDMReports" | Sort CreationTime | Select -Last 1)
$importFile = "\\srv003789\ADDMReports\" + $latestADDMextract
$importData = Import-CSV $importFile

$path = "D:\UserData\Ibraaheem\Scripts\VMWare\Get-HostVMs"
$hostList = Get-Content "$path\serverList.txt"

foreach($vmHostName in $hostList){
    if(!($vmHostName -like ".mud.internal.co.za")){
        $vmHostName += ".mud.internal.co.za"
    }

    Try{
        $vmHostDetails = Get-VMHost $vmHostName -ErrorAction Stop
        $vmHostDetails.Name
        $vmList = $vmHostDetails | Get-VM
        $vmCount = ($vmList | Measure-Object -Line).Lines
            $summaryObject = [pscustomObject] @{
                "HostName" = $vmHostDetails.Name
                "ClusterName" = $vmHostDetails.Parent
                "VMs Impacted" = $vmCount
            } | Export-Excel -Path "$path\Details.xlsx" -WorksheetName "Summary" -Append -AutoSize -BoldTopRow -AutoFilter -FreezeTopRow
    
        
        $vmHostName

        $vmList | % {

            $competency = $null
            $osType = $null
            $sla = $null
            $app = $null
            $desc = $null
            $appOwner = $null
            $priTech = $null
            $secTech = $null

            foreach($row in $importData){
                if($row.Name -like "*$($_.Name)*"){
                    $competency = $row.'competency '
                    $osType = $row.'Discovered OS Type'
                    $sla = $row.sla
                    $app = $row.application
                    $desc = $row.serverdescription
                    $appOwner = $row.appowner
                    $priTech = $row.primarytechowner
                    $secTech = $row.secondarytechowner
                    break
                }else{
                    $competency = "Not in ADDM"
                    $osType = $null
                    $sla = $null
                    $app = $null
                    $desc = $null
                    $appOwner = $null
                    $priTech = $null
                    $secTech = $null
                }
            }            

            $vmOutput = [pscustomobject] @{
                "vmName" = $_.Name
                "HostName" = $vmHostDetails.Name
                "ClusterName" = $vmHostDetails.Parent
                "Competency" = $competency
                "OS" = $osType
                "SLA" = $sla
                "Application" = $app
                "Description" = $desc
                "App Owner" = $appOwner
                "PriTechOwner" = $priTech
                "SecTechOwner" = $secTech
            } | Export-Excel -Path "$path\Details.xlsx" -WorksheetName "VMDetails" -Append -AutoSize -BoldTopRow -AutoFilter -FreezeTopRow

        }

    }catch{
        "oops - $vmHostName"
            $hostIssuesObject = [pscustomObject] @{
                "HostName" = $vmHostName
            } | Export-Excel -Path "$path\Details.xlsx" -WorksheetName "Hosts not Found" -Append -AutoSize -BoldTopRow -AutoFilter -FreezeTopRow
    }

}