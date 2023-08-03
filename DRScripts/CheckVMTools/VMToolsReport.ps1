$path = "D:\UserData\Ibraaheem\Scripts\VMWare\DRScripts\CheckVMTools"
$vms = Get-Content "$path\VMList.txt"
$outputPath = "$path\VMToolsReport.xlsx"

### TWSDEV = SRV008866 ###

foreach($server in $vms){
    
    if($server -match "TWSDEV"){
        $serverName = "SRV008866"
    }else{
        $serverName = $server.Replace(" ", "")
    }

    Try{
        $vm = Get-VM $serverName -ErrorAction Stop
        $toolsOut = [pscustomobject] @{
            VM = $server
            ToolsStatus = $vm.ExtensionData.Guest.ToolsStatus
            ToolsRunning = $vm.ExtensionData.Guest.ToolsRunningStatus
            ToolsVersion = $vm.ExtensionData.Guest.ToolsVersion
            OS = $vm.ExtensionData.Guest.GuestFamily
        } | Export-Excel $outputPath -WorksheetName ToolsCheck -Append -BoldTopRow -AutoSize -AutoFilter

    }catch{

        $errorMessage = ($Error[0].exception).Message
        $errOut = [pscustomobject] @{
            VM = $server
            Error = $errorMessage
        } | Export-Excel $outputPath -WorksheetName Errors -Append -BoldTopRow -AutoSize -AutoFilter

    }
}