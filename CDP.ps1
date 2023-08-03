<# CDP for Absolute losers #>
$vmhosts = Get-VMHost sky02esx110.mud.internal.co.za
$CDPARRAY = @() 

 

foreach($VMHost in $vmhosts) 
{ 
   
$vmh = Get-VMHost $VMHost
If ($vmh.State -ne "Connected") 
 {
   Write-Output "Host $($vmh) state is not connected, skipping."
  }
Else 
   {
    Get-View $vmh.ID | `
    % { $esxname = $_.Name; Get-View $_.ConfigManager.NetworkSystem} | `
    %  { foreach ($physnic in $_.NetworkInfo.Pnic) 
        {
        $pnicInfo = $_.QueryNetworkHint($physnic.Device)
        foreach( $hint in $pnicInfo ){
            if ( $hint.ConnectedSwitchPort ){
                    $cdpinformation = $hint.ConnectedSwitchPort | select @{n="VMHost";e={$esxname}},@{n="VMNic";e={$physnic.Device}},DevId,Address,PortId,HardwarePlatform
                    $CDPDATA = New-Object System.Object
                    $CDPDATA | Add-Member -MemberType NoteProperty -Name "HostName" -Value $cdpinformation.VMHost
                    $CDPDATA | Add-Member -MemberType NoteProperty -Name "PNIC" -Value $cdpinformation.VMNic
                    $CDPDATA | Add-Member -MemberType NoteProperty -Name "Switch DeviceID" -Value $cdpinformation.DevId
                    $CDPDATA | Add-Member -MemberType NoteProperty -Name "Switch Port Number" -Value $cdpinformation.PortId
                    $CDPDATA | Add-Member -MemberType NoteProperty -Name "Switch hardware Platform" -Value $cdpinformation.HardwarePlatform
                    $CDPARRAY += $CDPDATA
            }else{
                    Write-Host "No CDP information available."
            }
        }
     }   
   }
 }
}

$CDPDATA | Export-Excel 