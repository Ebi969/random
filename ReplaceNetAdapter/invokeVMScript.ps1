$vm = "VLAN Test - Santam"

$ip = "10.12.75.199"
$gateWay = "10.12.75.1"
$subnetMask = "255.255.255.0"
$dnsPreferred = "10.11.21.121"
$dnsAlternate = ""
$newName = "Campus"

$newDomainName = "mud.internal.co.za"
$domainUserName = ""
$domainPassword = ""

$newDescription = "PAMCO TMP WIN2019"
$newServerName = "SRV008263"

if($subnetMask -eq "255.255.255.0"){
    $prefixLength = "24"
}elseif($subnetMask -eq "255.255.240.0"){
    $prefixLength = "20"
}elseif($subnetMask -eq "255.255.0.0"){
    $prefixLength = "16"
}

$addIPscript = @'
    set-netipinterface -interfacealias "Ethernet0" –dhcp disabled
    New-NetIPAddress -InterfaceAlias "Ethernet0" -IPAddress #ipAddress# -AddressFamily IPv4 -DefaultGateway #gateway# -PrefixLength #prefixLength#
    Set-DnsClientServerAddress -InterfaceAlias "Ethernet0" -ServerAddresses #dnsPreferred#
    Disable-NetAdapterBinding -Name "Ethernet0" -ComponentID ms_tcpip6
'@
$addIPscript = $addIPscript.Replace("#newNicName#", $newName)
$addIPscript = $addIPscript.Replace("#ipAddress#", $ip)
$addIPscript = $addIPscript.Replace("#gateway#", $gateWay)
$addIPscript = $addIPscript.Replace("#prefixLength#", $prefixLength)
$addIPscript = $addIPscript.Replace("#dnsPreferred#", $dnsPreferred)
$addIPscript = $addIPscript.Replace("#dnsAlternate#", $dnsAlternate)

$addToDomainScript = @'
    $domainUser = "#uName#"
    $domainPword = ConvertTo-SecureString -String "#pWord#" -AsPlainText -Force
    $domainCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $DomainUser, $DomainPWord
    add-computer -domainname #NewDomainName# -credential $domainCredential -restart -force
'@
$addToDomainScript = $addToDomainScript.Replace("#NewDomainName#", $newDomainName)
$addToDomainScript = $addToDomainScript.Replace("#pWord#", $domainUserName)
$addToDomainScript = $addToDomainScript.Replace("#uName#", $domainPassword)

$changeServerDetails = @'
    Get-CimInstance -ClassName Win32_OperatingSystem | Set-CimInstance -Property @{Description = '#NewDescription#'}

    $localUser = "Administrator"
    $localPword = ConvertTo-SecureString -String "asdfQWER1234" -AsPlainText -Force
    $localCrendential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $localUser, $localPword

    Rename-Computer -NewName #NewServerName# -LocalCredential $localCrendential -PassThru

'@
$changeServerDetails = $changeServerDetails.Replace("#NewDescription#", $newDescription)
$changeServerDetails = $changeServerDetails.Replace("#NewServerName#", $newServerName)

$enableFirewallRules = @'
    Set-NetFirewallRule -DisplayName "File and Printer Sharing (Echo Request - ICMPv4-In)" -Enabled True
    
    Set-NetFirewallRule -DisplayName "Remote Service Management (NP-In)" -Enabled True
    Set-NetFirewallRule -DisplayName "Remote Service Management (RPC)" -Enabled True
    Set-NetFirewallRule -DisplayName "Remote Service Management (RPC-EPMAP)" -Enabled True
    
    Set-NetFirewallRule -DisplayName "Windows Management Instrumentation (ASync-In)" -Enabled True
    Set-NetFirewallRule -DisplayName "Windows Management Instrumentation (DCOM-In)" -Enabled True
    Set-NetFirewallRule -DisplayName "Windows Management Instrumentation (WMI-In)" -Enabled True
'@

$changeCDLetter = @'
    Get-WmiObject -Class Win32_Volume -Filter 'DriveType=5' | Select-Object -First 1 | Set-WmiInstance -Arguments @{DriveLetter='Z:'}
'@

#Invoke-VMScript -VM $vm -ScriptText $addIPscript -GuestUser Administrator -GuestPassword asdfQWER1234
#Invoke-VMScript -VM $vm -ScriptText $changeServerDetails -GuestUser Administrator -GuestPassword asdfQWER1234
#Invoke-VMScript -VM $vm -ScriptText $addToDomainScript -GuestUser Administrator -GuestPassword asdfQWER1234

#Invoke-VMScript -VM $vm -ScriptText $enableFirewallRules -GuestUser Administrator -GuestPassword asdfQWER1234
#Invoke-VMScript -VM $vm -ScriptText $changeCDLetter -GuestUser Administrator -GuestPassword asdfQWER1234
