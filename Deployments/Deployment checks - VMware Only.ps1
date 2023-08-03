<#
$creds = Get-Credential

import-module VMware.VimAutomation.Core
import-module VMware.VimAutomation.Storage

connect-viserver srv006383, srv006384, srv007281, srv007282 -credential $creds
#>
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

#End of HTML
$objHTML +=	"</body>"
$objHTML +=	"</html>"

if(!(Test-Path "D:\Reports\Server\DeploymentCheck\$serverName")){
New-Item -Path "D:\Reports\Server\DeploymentCheck" -ItemType Directory -Name $serverName
}

$objFilename = "VMwareCheck_" + $objHost + "_" + $objTxtDate

$objHTML | out-file "D:\Reports\Server\DeploymentCheck\$serverName\$objFilename.html"