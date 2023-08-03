(Get-View -ViewType VirtualMachine -Property Name,'Config.Hardware' | Where-Object { $_.Config.Hardware.Device.Where({$_.gettype().name -match 'VirtualUSBXHCIController'}) } ).count
