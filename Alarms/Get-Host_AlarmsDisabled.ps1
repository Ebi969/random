$outputPath = "C:\Users\DA003089\Desktop\Scripts\VMWare\Alarms"

if(Test-Path $outputPath\Host_Disabled_Alarms.csv){
    Remove-Item $outputPath\Host_Disabled_Alarms.csv
}

Foreach($hostsName in (Get-VMHost| where{!$_.ExtensionData.AlarmActionsEnabled})){

$disabledAlarmHost = $hostsName.Name
$cluster = Get-VMHost $hostsName.Name | Get-Cluster | Select -ExpandProperty Name
$disabledAlarmHost
$cluster

        $out = New-Object psobject

        $out | Add-Member -MemberType NoteProperty -Name "Host Name" -Value $disabledAlarmHost
        $out | Add-Member -MemberType NoteProperty -Name "Cluster Name" -Value $cluster

        $out | Export-Csv -Path $outputPath\Host_Disabled_Alarms.csv -NoTypeInformation -Append

}

