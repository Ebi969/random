
$creds = Get-Credential

import-module VMware.VimAutomation.Core
import-module VMware.VimAutomation.Storage

connect-viserver srv006383, srv006384, srv007281, srv007282 -credential $creds

$objHtml = $null

$serverName = Read-Host -Prompt 'Input your server  name'

#Declaration of Variables
$objHost = $serverName
$objFormattedDate = get-date -f "dd-MM-yyyy HH:mm:ss"
$objTxtDate = get-date -f "yyyy-MM-dd_HHmmss"

#Start of HTML Document format
$objHTML =	"<html>"
$objHTML +=	"<head>"
$objHTML +=	"<Title>" + $objHost + "</Title>"
$objHTML +=	"<Style>"
$objHTML +=	" table{  border: 1px solid black;}`
			 td {  border-bottom: 1px solid #ddd;text-align: left;font-size:12}`
			 .label {  border-bottom: 1px solid #ddd;text-align: left;font-weight: bold;color:blue;font-size:18}`
			 th {  border-bottom: 1px solid #ddd;text-align: left;font-size:15;}`
            "
$objHTML +=	"</Style>"
$objHTML +=	"</head>"
$objHTML +=	"<body>"
$objHTML +=	"<h1><b>"+ $objHost+"_"+$objFormattedDate+"</h1></b><th>"

#####################################################################
#Computer information
#####################################################################

Write-Host "Collecting computer information..." -ForegroundColor Green

$osInfo = Get-CimInstance -ClassName Win32_OperatingSystem -CimSession $serverName

$totalVisibleMemory = [math]::round(((($osInfo.TotalVisibleMemorySize) / 1024) /1024), 2)

$procInfo = Get-CimInstance -ClassName Win32_Processor -CimSession $serverName 

$computerOU = Get-ADComputer -Identity $serverName -Properties canonicalname | select -ExpandProperty canonicalname

Write-Host "Adding computer information to html..." -ForegroundColor Green

#Set the Table and first header
$objHTML +=	"<table width=100%>"
$objHTML += "<tr> <br> </tr>" 
$objHTML += "<tr>"
$objHTML += "<th colspan=2 class=""label""> Computer Information  </th>"
$objHTML += "</tr>"

#Begin the information dump
$objHTML +=	"<tr><td><b> Server Name </b></td>"
$objHTML +=	"<td>" + $osInfo.PSComputerName + "</td></tr>"
$objHTML +=	"<tr><td><b> OS Version </b></td>"
$objHTML +=	"<td>" + $osInfo.Caption + "</td></tr>"
$objHTML +=	"<tr><td><b> Architecture </b></td>"
$objHTML +=	"<td>" + $osInfo.OSArchitecture + "</td></tr>"
$objHTML +=	"<tr><td ><b> Total Visible Memory </b></td>"
$objHTML +=	"<td>" + $totalVisibleMemory + "GB</td></tr>"
$objHTML +=	"<tr><td><b> Number of CPU </b></td>"
$objHTML +=	"<td>" + $procInfo.NumberOfLogicalProcessors + "</td></tr>"
$objHTML +=	"<tr><td><b> ComputerOU </b></td>"
$objHTML +=	"<td>" + $computerOU + "</td></tr>"

#####################################################################
#Network Information
#####################################################################

$ipInfo = Test-Connection $serverName -Count 1 | Select IPV4Address

$ipV4Info =  $ipInfo.IPV4Address

$actualIp = $ipV4Info.IpAddressToString

$objHTML +=	"<tr><td><b> IP Address </b></td>"
$objHTML +=	"<td>" + $actualIp + "</td></tr>"

$defaultGatewayInfo = Get-NetIPConfiguration -CimSession $serverName | Select -ExpandProperty IPv4DefaultGateway
$actualDefaultGateway = $defaultGatewayInfo.NextHop

$objHTML +=	"<tr><td><b> Default Gateway </b></td>"
$objHTML +=	"<td>" + $actualDefaultGateway + "</td></tr>"

#####################################################################
#NLA Information
#####################################################################

$nlaValue = (Get-CimInstance -Class "Win32_TSGeneralSetting" -NameSpace root\cimv2\terminalservices -ComputerName $serverName -Filter "TerminalName='RDP-tcp'").UserAuthenticationRequired
if($nlaValue -eq "1"){
    $nlaEnabled = "True"
}else{
    $nlaEnabled = "False"
}
 
$objHTML +=	"<tr><td><b> NLA Enabled </b></td>"
$objHTML +=	"<td>" + $nlaEnabled + "</td></tr>"
$objHTML +=	"</table>"

Write-Host "Computer information complete..." -ForegroundColor Green

#####################################################################
#VMconfig information
#####################################################################

Write-Host "Collecting VMware information..." -ForegroundColor Green

#Set the Table and first header
$objHTML +=	"<table width=100%>"
$objHTML += "<tr> <br> </tr>" 
$objHTML += "<tr>"
$objHTML += "<th colspan=2 class=""label""> VMware Information  </th>"
$objHTML += "</tr>"

Write-Host "Adding VMware information to html..." -ForegroundColor Green

try{
$vm = Get-VM $serverName.Replace(" ","")  -ErrorAction Stop
        $vmView = $vm | Get-View
        $vmHost = $vm | Get-VMHost
        $VMToolsVer = $vm.ExtensionData.Guest.ToolsVersion
        $VmNetworkAdapters = $vm | Get-NetworkAdapter
        $vmDatastore = $vm | Get-Datastore | select -ExpandProperty Name
            if($VMToolsVer -eq 11265){
                $VMToolsVer += " (Version OK)"
            }else{
                $VMToolsVer += " (Requires Version update)"
            }
        $cores = $vmView.config.hardware.NumCPU
        $coresPerSocket = $vmView.config.hardware.NumCoresPerSocket
        $numSockets = $cores/$coresPerSocket
        if($numSockets -gt 1){
            $issueMsg = "Check core config"
        }else{
            $issueMsg = ""
        }
        $hostCoresPerSocket = $vmHost.ExtensionData.Hardware.CpuInfo.NumCpuCores / $vmHost.ExtensionData.Hardware.CpuInfo.NumCpuPackages
                
        #Begin the information dump
            $objHTML +=	"<tr><td><b> VMtools Version </b></td>"
            $objHTML +=	"<td colspan=4>" + $VMToolsVer + "</td></tr>"
            $objHTML +=	"<tr><td><b> Total Cores </b></td>"
            $objHTML +=	"<td colspan=4>" + $cores + "</td></tr>"
            $objHTML +=	"<tr><td><b> Cores Per Socket </b></td>"
            $objHTML +=	"<td colspan=4>" + $coresPerSocket + "</td></tr>"
            $objHTML +=	"<tr><td ><b> Total Sockets </b></td>"
            $objHTML +=	"<td colspan=4>" + $numSockets + "</td></tr>"
            $objHTML +=	"<tr><td><b> Host core per socket </b></td>"
            $objHTML +=	"<td colspan=4>" + $hostCoresPerSocket + "</td></tr>"
            $objHTML +=	"<tr><td><b> Datastore </b></td>"
            $objHTML +=	"<td colspan=4>" + $vmDatastore + "</td></tr>"
                           
                $objHTML +=	"<tr style=`"outline:thin solid`"><td colspan=5; style=`"text-align:center`"><b> Network Adapters </b></td></tr>"

                $objHTML +=	"<tr><td><b> Name </b></td>"               
                $objHTML +=	"<td><b> Type </b></td>"               
                $objHTML +=	"<td><b> vLAN </b></td>"               
                $objHTML +=	"<td><b> MacAddress </b></td>"  
                $objHTML +=	"<td><b> Direct Path I/O </b></td></tr>"


            foreach($nic in $VmNetworkAdapters){

                if($nic.Type -notmatch "Vmxnet3" -or $nic.ExtensionData.UptCompatibilityEnabled -eq $true){
                $objHTML +=	"<tr style=`"color:#FF0000`"><td>" + $nic.Name + "</td>"
                $objHTML +=	"<td>" + $nic.Type + "</td>"
                $objHTML +=	"<td>" + $nic.NetworkName + "</td>"
                $objHTML +=	"<td>" + $nic.MacAddress + "</td>"
                $objHTML +=	"<td>" + $nic.ExtensionData.UptCompatibilityEnabled + "</td></tr>"
                }else{
                $objHTML +=	"<tr><td>" + $nic.Name + "</td>"
                $objHTML +=	"<td>" + $nic.Type + "</td>"
                $objHTML +=	"<td>" + $nic.NetworkName + "</td>"
                $objHTML +=	"<td>" + $nic.MacAddress + "</td>" 
                $objHTML +=	"<td>" + $nic.ExtensionData.UptCompatibilityEnabled + "</td></tr>"             
                }
            }

            $objHTML +=	"</table>"

}catch{
        #Begin the information dump
            $objHTML +=	"<tr><td><b> Server is not in the VMware Environment </b></td>"
            $objHTML +=	"</table>"
}

Write-Host "VMware information complete..." -ForegroundColor Green

#####################################################################
#Disk information
#####################################################################
Write-Host "Collecting Disk information..." -ForegroundColor Green

#Set the Table and first header
$objHTML +=	"<table width=100%>"
$objHTML += "<tr> <br> </tr>" 
$objHTML += "<tr>"
$objHTML += "<th colspan=2 class=""label""> Disk Information  </th>"
$objHTML += "</tr>"

#Set Headers
$objHTML+= "<tr>"
$objHTML+= "<th><b> Disk Number  </b></th>"
$objHTML+= "<th><b> Drive Letter  </b></th>"
$objHTML+= "<th><b> Drive Size  </b></th>"
$objHTML+= "<th><b> Format Style  </b></th>"
$objHTML+= "<th><b> Partition Style  </b></th>"
$objHTML+= "<th><b> Operational Status  </b></th>"
$objHTML+= "<th><b> Allocation Unit Size  </b></th>"
$objHTML+= "</tr>"

$serverVolumes = Invoke-Command -Credential $creds -ComputerName $serverName -Scriptblock {
    Get-Volume | Where-Object {$_.DriveType -eq "Fixed"} | select DriveLetter, FileSystem, AllocationUnitSize -Unique
}

Write-Host "Adding Disk information to html..." -ForegroundColor Green
$finalCollect = @()
foreach($volume in $serverVolumes){

    if($volume.DriveLetter -ne $null){
        $partitionInfo = Invoke-Command -Credential $creds -ComputerName $serverName -Scriptblock {
            Get-Partition -DriveLetter $using:volume.DriveLetter | Select -Unique
        }

        $diskInfo = Invoke-Command -Credential $creds -ComputerName $serverName -Scriptblock {
            Get-Disk -Number $using:partitionInfo.DiskNumber | select DiskNumber, PartitionStyle, Size
        }


        $indivdriveFinal = [pscustomobject] @{
            DiskNumber = $partitionInfo.DiskNumber
            DriveLetter = $volume.DriveLetter
            DriveSize = [Math]::Round(($diskInfo.Size/1024/1024/1024), 2)
            FormatStyle = $volume.FileSystem
            PartitionType = $diskInfo.PartitionStyle
            OperationalStatus = $partitionInfo.OperationalStatus
            AllocationUnitSize = $volume.AllocationUnitSize
        }
        $finalCollect += $indivdriveFinal
    }
}

foreach($driveFinal in $finalCollect | Sort-Object DriveLetter){
#Begin the information dump
    $objHTML+=	"<tr>"
	$objHTML+=	"<td>" +	$driveFinal.DiskNumber		+ "</td>"
	$objHTML+=	"<td>" +	$driveFinal.DriveLetter	+ "</td>"
	$objHTML+=	"<td>" +	$driveFinal.DriveSize	+ "GB</td>"
	$objHTML+=	"<td>" +	$driveFinal.FormatStyle			+ "</td>"
	$objHTML+=	"<td>" +	$driveFinal.PartitionType		+ "</td>"
	$objHTML+=	"<td>" +	$driveFinal.OperationalStatus		+ "</td>"
	$objHTML+=	"<td>" +	$driveFinal.AllocationUnitSize		+ "</td>"
	$objHTML+=	"</tr>"
}

$objHTML +=	"</table>"

Write-Host "Disk information complete..." -ForegroundColor Green

#####################################################################
#App information
#####################################################################

Write-Host "Collecting App information..." -ForegroundColor Green

$installedPrograms = Get-CimInstance Win32_product -ComputerName $serverName | Select Name, Version, InstallDate, Vendor | Sort-Object -Property Name

#Set the Table and first header
$objHTML +=	"<table width=100%>"
$objHTML += "<tr> <br> </tr>" 
$objHTML += "<tr>"
$objHTML += "<th colspan=4 class=""label""> App Information  </th>"
$objHTML += "</tr>"


#Set Headers
$objHTML+= "<tr>"
$objHTML+= "<th><b> Name  </b></th>"
$objHTML+= "<th><b> Vendor  </b></th>"
$objHTML+= "<th><b> Version  </b></th>"
$objHTML+= "<th><b> Installed Date  </b></th>"
$objHTML+= "</tr>"

foreach($app in $installedPrograms){

$convertDate =[datetime]::ParseExact($app.InstallDate,'yyyyMMdd', $null)
    $dateInstalled = Get-Date $convertDate -Format "yyyy-MM-dd"

#Begin the information dump
    $objHTML+=	"<tr>"
	$objHTML+=	"<td>" +	$app.Name		+ "</td>"
	$objHTML+=	"<td>" +	$app.Vendor	+ "</td>"
	$objHTML+=	"<td>" +	$app.Version			+ "</td>"
	$objHTML+=	"<td>" +	$dateInstalled			+ "</td>"
	$objHTML+=	"</tr>"
}

$objHTML +=	"</table>"

Write-Host "App information complete..." -ForegroundColor Green

#End of HTML
$objHTML +=	"</body>"
$objHTML +=	"</html>"

if(!(Test-Path "D:\Reports\Server\DeploymentCheck\$serverName")){
New-Item -Path "D:\Reports\Server\DeploymentCheck" -ItemType Directory -Name $serverName
}

$gpName = "GPResults_" + $serverName + "_" + $objTxtDate

$gpExport = "\\SLM02S2D005\Reports\Server\DeploymentCheck\$serverName\$gpName.html"

Invoke-Command -Credential $creds -ComputerName $serverName -Scriptblock {
    if(!(Test-Path C:\Temp\$using:serverName)){
        New-Item -Path "C:\Temp" -ItemType Directory -Name $using:serverName
    }
    gpresult /h C:\Temp\$using:serverName\GPResults_$using:serverName.html /f
}

Copy-Item -Path "\\$serverName\C$\Temp\$serverName\GPResults_$serverName.html" -Destination  $gpExport -Force

$objFilename = "CompCheck_" + $objHost + "_" + $objTxtDate

$objHTML | out-file "D:\Reports\Server\DeploymentCheck\$serverName\$objFilename.html"

Copy-Item -Path "\\SLM02S2D005\Reports\Server\DeploymentCheck\$serverName" -Destination "\\SRV005879\Reports\Server\DeploymentCheck\$serverName" -Force
Copy-Item -Path "\\SLM02S2D005\Reports\Server\DeploymentCheck\$serverName\$gpName.html" -Destination "\\SRV005879\Reports\Server\DeploymentCheck\$serverName\$gpName.html" -Force
Copy-Item -Path "\\SLM02S2D005\Reports\Server\DeploymentCheck\$serverName\$objFilename.html" -Destination "\\SRV005879\Reports\Server\DeploymentCheck\$serverName\$objFilename.html" -Force