
$vmlist = Get-content "D:\UserData\Ibraaheem\Scripts\VMWare\VDI-Migs\serverList.txt"

$vmReport = @()
foreach($vm in $vmlist){
    $vmDetails = Get-VM $vm | select Name,@{E={$_.ExtensionData.Config.Files.VmPathName};L=”VM Path”}
    $vmReport += $vmDetails
}