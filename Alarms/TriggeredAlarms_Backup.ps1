
$cred = get-credential

$vcList = @('srv006383', 'srv006384', 'srv007280', 'srv007281', 'srv007282')

$timestamp=Get-Date -Format s | foreach {$_ -replace "-", ""} | foreach {$_ -replace "T", "_"} | foreach {$_ -replace ":",""} 

$fileName = "D:\Scripts\VMware_Alarms\AlarmOutput\TriggeredAlarms-$timestamp.xlsx" 

if(Test-Path $fileName){
    Remove-Item $fileName
}

foreach($oneVC in $vcList){

    Connect-viserver $oneVC -credential $cred        
    
    $vc = $global:DefaultVIServer

   $si = Get-View ServiceInstance -Server $vc
   $root = Get-View -Id $si.Content.RootFolder
   $root.TriggeredAlarmState |
   Select @{N='vCenter';E={$vc.Name}},
     @{N='Entity';E={(Get-View -id $_.Entity -Property Name).Name}},
     @{N='Alarm';E={(Get-View -Id $_.Alarm -Property Info).Info.Name}},
     Time,OverallStatus,Acknowledged | Export-Excel -Path $fileName -WorkSheetname "$vc" -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow
          
    disconnect-viserver $oneVC -Force -Confirm:$false

}

