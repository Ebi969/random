$objHtml = $null

$serverName = $env:COMPUTERNAME

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
#Disk information
#####################################################################
Write-Host "Collecting Disk information..." -ForegroundColor Green

#Set the Table and first header
$objHTML +=	"<table width=100%>"
$objHTML += "<tr> <br> </tr>" 
$objHTML += "<tr>"
$objHTML += "<th colspan=7 class=""label""> Disk Information  </th>"
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

$serverVolumes = Get-Volume | Where-Object {$_.DriveType -eq "Fixed"} | select DriveLetter, FileSystem, AllocationUnitSize -Unique

$finalCollect = @()

foreach($volume in $serverVolumes){

    if($volume.DriveLetter -ne $null){
        $partitionInfo = Get-Partition -DriveLetter $volume.DriveLetter | Select -Unique

        $diskInfo = Get-Disk -Number $partitionInfo.DiskNumber | select DiskNumber, PartitionStyle, Size

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

Write-Host "Adding Disk information to html..." -ForegroundColor Green

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

if(!(Test-Path "C:\Temp\Server\DeploymentCheck\$serverName")){
New-Item -Path "C:\Temp\Server\DeploymentCheck" -ItemType Directory -Name $serverName
}

    gpresult /h C:\Temp\Server\DeploymentCheck\$serverName\GPResults_$serverName.html /f

$objFilename = "CompCheck_" + $objHost + "_" + $objTxtDate

$objHTML | out-file "C:\Temp\Server\DeploymentCheck\$serverName\$objFilename.html"