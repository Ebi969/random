$directPath = $false
$path = "D:\UserData\Ibraaheem\Scripts\VMWare\vmNicStuff"
$vms = Get-Content $path\ServerList.txt

foreach($vmName in $vms){
    #$vm = Get-VM $vmName.replace(" ","")
    $vm = Get-VM $vmName
    $vm
    
    
    $vm | Get-NetworkAdapter | Foreach {
        $_ | Select Parent, Name, Type, @{n="Direct Path I/O"; e={$_.ExtensionData.UptCompatibilityEnabled}}
    } | Export-Excel -Path "$path\dpioFIX\dpioFix.xlsx" -WorksheetName "Before" -Append -AutoFilter -AutoSize -BoldTopRow
        $nics = $vm | Get-NetworkAdapter | Where {$_.ExtensionData.UptCompatibilityEnabled -eq $true}
        foreach($nic in $nics){
            $vmConfigSpec = New-Object VMware.Vim.VirtualMachineConfigSpec
            $deviceConfigSpec = New-Object VMware.Vim.VirtualDeviceConfigSpec

            $deviceConfigSpec.Operation = [VMware.Vim.VirtualDeviceConfigSpecOperation]::edit
            $deviceConfigSpec.Device = [VMware.Vim.VirtualDevice]$nic.ExtensionData 
            $deviceConfigSpec.Device.UptCompatibilityEnabled = $directPath
            $vmConfigSpec.DeviceChange += $deviceConfigSpec
            $vm.ExtensionData.ReconfigVM($vmConfigSpec)
        }
    $vm | Get-NetworkAdapter | Foreach {
        $_ | Select Parent, Name, Type, @{n="Direct Path I/O"; e={$_.ExtensionData.UptCompatibilityEnabled}}
    } | Export-Excel -Path "$path\dpioFIX\dpioFix.xlsx" -WorksheetName "After" -Append -AutoFilter -AutoSize -BoldTopRow
    
    
}