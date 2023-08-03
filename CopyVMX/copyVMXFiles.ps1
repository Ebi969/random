$vmlist = Get-content "D:\UserData\Ibraaheem\Scripts\VMWare\CopyVMX\vmList.txt"
$destination = "D:\UserData\Ibraaheem\Scripts\VMWare\CopyVMX\VMXFiles\"

foreach($vmName in $vmlist){
$vmName
$vm = Get-VM $($vmName.Replace(" ", "")) | Select Name, @{N="VMX";E={$_.Extensiondata.Summary.Config.VmPathName}}
$pathSplit = $vm.VMX.split("]")
$dsName = ($pathSplit)[0].replace("[","")
$restPath = ($pathSplit[1].Replace("/","\")).Replace(" ", "")
$datastore = Get-Datastore $dsName

New-PSDrive -Location $datastore -Name ds -PSProvider VimDatastore -Root "\"

Copy-DatastoreItem -Item "ds:\$restPath" -Destination $destination

Remove-PSDrive -name ds

}