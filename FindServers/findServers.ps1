$completeList = Get-Content "C:\Users\DA003089\Desktop\Scripts\VMWare\FindServers\serverList.txt"

$latestADDMextract = (Get-ChildItem -Path "\\srv003789\ADDMReports" | Sort CreationTime | Select -Last 1)
    $importFile = "\\srv003789\ADDMReports\" + $latestADDMextract
    $importData = Import-CSV $importFile

$finalOut = @()

foreach($row in $importData){
    foreach($server in $completeList){
        $serverName = $server.Replace(" ", "")

        if($row.Name -like "*$serverName*"){
            $output = [pscustomobject] @{
                    'Name' = $row.Name
                    'Ip Address' = $row.'IP Address'
                    'Vendor' = $row.'Hardware Vendor'
                    'Last Communicated with ADDM' = $row.'Last Update Success'
            }
            $finalOut += $output
        }
    }
}

$finalOut | Out-GridView