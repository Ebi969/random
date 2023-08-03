Get-VM | Where{$_.PowerState -eq "PoweredOff"} | foreach{
    $_.Name

    Get-VIEvent -Entity $_.Name | Where {$_.FullFormattedMessage -like "*Task: Initiate guest OS shutdown*"} | Select UserName, FullFormattedMessage

}