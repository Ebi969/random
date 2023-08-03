$path = "D:\UserData\Ibraaheem\Scripts\VMWare\HostVMs"
$hostList = Get-Content $path\hostlist.txt

$allInfo = @()
foreach($hostName in $hostList){
    if($hostName.Contains(".mud.internal.co.za")){
        $VMhost = $hostName
        $vmHostNoMud = $VMhost.Replace(".mud.internal.co.za","")
    }else{
        $VMhost = $hostName + ".mud.internal.co.za"
        $vmHostNoMud = $hostName
    }

    $VMhostDetail = Get-VMHost $VMhost
    $hostVMs = $VMhostDetail | Get-VM

    foreach ($vm in $hostVMs){
        $output = [pscustomobject] @{

            cluster = $VMhostDetail.Parent
            vmHost = $VMhostDetail.Name
            vm = $vm.Name

        }

        $allInfo += $output

    }
}

$allInfo | Export-Excel $path\VMList.xlsx -Append 