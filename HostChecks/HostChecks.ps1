$path = "D:\UserData\Ibraaheem\Scripts\VMWare\HostChecks"
$hostList = Get-Content $path\hostlist.txt

foreach($hostName in $hostList){

$objHTML = $null

$objFormattedDate = get-date -f "dd-MM-yyyy HH:mm:ss"
$objTxtDate = get-date -f "yyy-MM-dd_HHmmss"

    if($hostName.Contains(".mud.internal.co.za")){
        $VMhost = $hostName
        $vmHostNoMud = $VMhost.Replace(".mud.internal.co.za","")
    }else{
        $VMhost = $hostName + ".mud.internal.co.za"
        $vmHostNoMud = $hostName
    }

    $VMhostDetail = Get-VMHost $VMhost
    $esxc = Get-EsxCli -VMHost $VMhost
    $hostVMs = $VMhostDetail | Get-VM
    $hostDatastores = $VMhostDetail | Get-Datastore
    $hostHBAs = Get-VMHostHba -VMHost $VMhost -Type "FibreChannel" | where Status -EQ online

    $objHTML =	"<html>"
    $objHTML +=	"<head>"
    $objHTML +=	"<Title>" + $vmHostNoMud + "</Title>"
    $objHTML +=	"<Style>"
    $objHTML +=	" table{  border: 1px solid black;}`
			     td {  border-bottom: 1px solid #ddd;text-align: left;font-size:12}`
			     .label {  border-bottom: 1px solid #ddd;text-align: left;font-weight: bold;color:blue;font-size:18}`
			     th {  border-bottom: 1px solid #ddd;text-align: left;font-size:15;}`
    "
    $objHTML +=	"</Style>"
    $objHTML +=	"</head>"
    $objHTML +=	"<body>"
    $objHTML +=	"<h1><b>"+ $vmHostNoMud+"_"+$objFormattedDate+"</h1></b><th>"

#Host Information#
#Set the Table and first header
    $objHTML +=	"<table width=100%>"
    $objHTML += "<tr> <br> </tr>" 
    $objHTML += "<tr>"
    $objHTML += "<th class=""label"" colspan=2> Host Information  </th>"
    $objHTML += "</tr>"

#Begin the information dump
    $objHTML +=	"<tr><td><b> ConnectionState </b></td>"
    $objHTML +=	"<td>" + $VMhostDetail.ConnectionState + "</td></tr>"
    $objHTML +=	"<tr><td><b> PowerState </b></td>"
    $objHTML +=	"<td>" + $VMhostDetail.PowerState + "</td></tr>"
    $objHTML +=	"<tr><td><b> NumCpu </b></td>"
    $objHTML +=	"<td>" + $VMhostDetail.NumCpu + "</td></tr>"
    $objHTML +=	"<tr><td><b> CpuUsageMhz </b></td>"
    $objHTML +=	"<td>" + $VMhostDetail.CpuUsageMhz + "</td></tr>"
    $objHTML +=	"<tr><td><b> CpuTotalMhz </b></td>"
    $objHTML +=	"<td>" + $VMhostDetail.CpuTotalMhz + "</td></tr>"
    $objHTML +=	"<tr><td><b> MemoryUsageGB </b></td>"
    $objHTML +=	"<td>" + [math]::Round($VMhostDetail.MemoryUsageGB,0) + "</td></tr>"
    $objHTML +=	"<tr><td><b> MemoryTotalGB </b></td>"
    $objHTML +=	"<td>" + [math]::Round($VMhostDetail.MemoryTotalGB,0) + "</td></tr>"
    $objHTML +=	"<tr><td><b> Version </b></td>"
    $objHTML +=	"<td>" + $VMhostDetail.Version + "</td></tr>"
    $objHTML +=	"</table>"

#VM Information#
    $objHTML +=	"<table width=100%>"
    $objHTML += "<tr> <br> </tr>" 
    $objHTML += "<tr>"
    $objHTML += "<th class=""label"" colspan=5> VM Information  </th>"
    $objHTML += "</tr>"    
    
    #Set Headers
    $objHTML += "<tr>"
    $objHTML += "<th><b> VM Name  </b></th>"
    $objHTML += "<th><b> PowerState  </b></th>"
    $objHTML += "<th><b> Num Cpus  </b></th>"
    $objHTML += "<th><b> MemoryGB  </b></th>"
    $objHTML += "<th><b> ProvisionedSpaceGB  </b></th>"
        foreach($vm in $hostVMs){
            $objHTML += "<tr>"
            $objHTML +=	"<td>" +	$vm.Name				 	+ "</td>"
	        $objHTML +=	"<td>" +	$vm.PowerState					+ "</td>"
	        $objHTML +=	"<td>" +	$vm.NumCpu		+ "</td>"
	        $objHTML +=	"<td>" +	[math]::Round($vm.MemoryGB,0)	+ "GB</td>"
            $objHTML +=	"<td>" +	[math]::Round($vm.ProvisionedSpaceGB,2)	+ "</td>"
	        $objHTML +=	"</tr>"
        }
    $objHTML += "</table>"

#Datastore Information#
    $objHTML +=	"<table width=100%>"
    $objHTML += "<tr> <br> </tr>" 
    $objHTML += "<tr>"
    $objHTML += "<th class=""label"" colspan=3> Datastore Information - " + $hostDatastores.Count +" </th>"
    $objHTML += "</tr>"    
    
    #Set Headers
    $objHTML += "<tr>"
    $objHTML += "<th><b> DS Name  </b></th>"
    $objHTML += "<th><b> FreeSpaceGB  </b></th>"
    $objHTML += "<th><b> CapacityGB  </b></th>"
        foreach($ds in $hostDatastores){
            $objHTML += "<tr>"
            $objHTML +=	"<td>" +	$ds.Name				 	+ "</td>"
	        $objHTML +=	"<td>" +	[math]::Round($ds.FreeSpaceGB,2)					+ "GB</td>"
	        $objHTML +=	"<td>" +	[math]::Round($ds.CapacityGB,2)		+ "GB</td>"
	        $objHTML +=	"</tr>"
        }
    $objHTML += "</table>"

#HBA Information#
    $objHTML +=	"<table width=100%>"
    $objHTML += "<tr> <br> </tr>" 
    $objHTML += "<tr>"
    $objHTML += "<th class=""label"" colspan=5> HBA Information  </th>"
    $objHTML += "</tr>"    
    
    #Set Headers
    $objHTML += "<tr>"
    $objHTML += "<th><b> HBA Name  </b></th>"
    $objHTML += "<th><b> HBA Status  </b></th>"
    $objHTML += "<th><b> Targets  </b></th>"
    $objHTML += "<th><b> Devices  </b></th>"
    $objHTML += "<th><b> Paths  </b></th>"
        foreach($hba in $hostHBAs){
            $hbaname = $hba.Name
            $target = ((Get-View $hba.VMhost).Config.StorageDevice.ScsiTopology.Adapter | where {$_.Adapter -eq $hba.Key}).Target
            $luns = Get-ScsiLun -Hba $hba  -LunType "disk"  -ErrorAction SilentlyContinue
            $wwn = $esxc.storage.core.adapter.list() | where HBAName -eq $hbaname | select UID
            $nrPaths = ($target | %{$_.Lun.Count} | Measure-Object -Sum).Sum

            $objHTML += "<tr>"
            $objHTML +=	"<td>" +	$hbaname				 	+ "</td>"
	        $objHTML +=	"<td>" +	$hba.Status					+ "</td>"
	        $objHTML +=	"<td>" +	$target.Count		+ "</td>"
	        $objHTML +=	"<td>" +	$luns.Count		+ "</td>"
	        $objHTML +=	"<td>" +	$nrPaths		+ "</td>"
	        $objHTML +=	"</tr>"
        }
    $objHTML += "</table>"

#End of HTML
$objHTML +=	"</body>"
$objHTML +=	"</html>"

$objFilename = $vmHostNoMud + "_" + $objTxtDate

#Write File
$objHTML | out-file "$path\$objFilename.html"
}