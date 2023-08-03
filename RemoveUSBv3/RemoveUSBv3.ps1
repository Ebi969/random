$creds = Get-Credential

# Connect-VIServer -Credential $creds
Connect-VIServer -Server SRV007281.mud.internal.co.za -Credential $creds
Connect-VIServer -Server SRV007282.mud.internal.co.za -Credential $creds

$outputLocation = "D:\UserData\Ibraaheem\Scripts\VMWare\RemoveUSBv3"

$deviceNotWanted = "USB" 
$ConfigureVM = $true 
$importList = Import-Excel $outputLocation\ProdList.xlsx 
$VMs = Get-VM $importList.'Virtual Machines with USBXHCIController'

foreach ($vm in $VMs){
  $allToRemove = @()
  $vmView = $vm | Get-View
  $allDevicesBefore = $vmView.Config.Hardware.Device

  $beforeOutput = @()
  foreach($indivDevice in $allDevicesBefore){

    $beforeRow = [pscustomobject] @{
        VM = $vm.Name
        Device = $indivDevice.DeviceInfo.Label
    }
    
    $beforeOutput += $beforeRow    

  }

  $beforeOutput | Export-Excel -Path $outputLocation\USBVMsRemoved.xlsx -WorksheetName "BeforeView" -Append -AutoFilter -BoldTopRow -AutoSize

  $devicesToRemove = $allDevicesBefore | where {$_.DeviceInfo.Label -match $deviceNotWanted} 
  $devicesToRemove | %{
    $deviceDetail = "" | select DeviceName, RemoveDev, Device
    $deviceDetail.DeviceName = $_.DeviceInfo.Label
    $deviceDetail.Device = $_
    if ($_.DeviceInfo.Label -match "xHCI"){
        $deviceDetail.RemoveDev = $true
    }else{
        $deviceDetail.RemoveDev = $false
    }
        $allToRemove += $deviceDetail | Sort Hardware
  }

  $allToRemove | Select @{N="VMName"; E={$vm.Name}}, DeviceName, @{N="ToBeRemoved?";E="RemoveDev"} | Export-Excel -Path $outputLocation\USBVMsRemoved.xlsx -WorksheetName "ToBeRemoved" -Append -AutoFilter -BoldTopRow -AutoSize
  
  if($allToRemove -and $ConfigureVM){
    # Unwanted Hardware is configured for removal
    $vmConfigSpec = New-Object VMware.Vim.VirtualMachineConfigSpec

    foreach($device in $allToRemove){
      if($deviceDetail.RemoveDev -eq $true){
        $vmConfigSpec.DeviceChange += New-Object VMware.Vim.VirtualDeviceConfigSpec
        $vmConfigSpec.DeviceChange[-1].device = $deviceDetail.Device
        $vmConfigSpec.DeviceChange[-1].operation = "remove"
        Write-Host "Removed $($deviceDetail.DeviceName) on $($vm.Name)"
      }
    }

    $vmView.ReconfigVM_Task($vmConfigSpec)
    Sleep -Seconds 10
  }# Unwanted Hardware is configured for removal
    
  $vmView = $vm | Get-View
  $allDevicesAfter = $vmView.Config.Hardware.Device

  $afterOutput = @()
  foreach($indivDevice in $allDevicesAfter){

    $afterRow = [pscustomobject] @{
        VM = $vm.Name
        Device = $indivDevice.DeviceInfo.Label
    }
    
    $afterOutput += $afterRow    

  }
  $afterOutput | Export-Excel -Path $outputLocation\USBVMsRemoved.xlsx -WorksheetName "afterView" -Append -AutoFilter -BoldTopRow -AutoSize
}