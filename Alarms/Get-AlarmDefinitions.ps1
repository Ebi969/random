$vcs = $global:DefaultVIServers

foreach($vc in $vcs){
$vc
    $alarms = Get-AlarmDefinition -Server $vc.Name
    $alarmReport = @()
    foreach($alarm in $alarms){

        $action = $alarm | Get-AlarmAction | Where {$_.ActionType -match "email"} | Select ActionType, To, Cc

        $row = [pscustomobject] @{
        
            Alarm = $alarm.Name
            VC = $vc.Name
            ActionType = $action.ActionType
            To = $action.To
            Cc = $action.Cc
        }
        $alarmReport += $row
    }
}

$alarmReport | Export-Excel -Path "D:\UserData\Ibraaheem\Scripts\VMWare\Alarms\All-Alarms.xlsx" -AutoFilter -BoldTopRow -Append