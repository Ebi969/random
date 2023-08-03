$path = "D:\UserData\Ibraaheem\Scripts\VMWare\HostVMs"
$hostList = Get-Content $path\hostlist.txt

foreach($hostName in $hostList){
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
    $hostVMs = $VMhostDetail | Get-VM

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

#End of HTML
$objHTML +=	"</body>"
$objHTML +=	"</html>"

$objFilename = $vmHostNoMud + "_" + $objTxtDate

#Write File
$objHTML | out-file "$path\$objFilename.html"
}