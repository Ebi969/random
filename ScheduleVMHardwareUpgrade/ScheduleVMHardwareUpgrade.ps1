$vms = Get-Cluster SKY-CDC02-GLD02-D-DATABASE | Get-VM | Where {$_.HardwareVersion -notmatch "15"}

$do = New-Object -TypeName VMware.Vim.VirtualMachineConfigSpec
$do.ScheduledHardwareUpgradeInfo = New-Object -TypeName VMware.Vim.ScheduledHardwareUpgradeInfo
$do.ScheduledHardwareUpgradeInfo.UpgradePolicy = “always”
$do.ScheduledHardwareUpgradeInfo.VersionKey = “vmx-15”
$vms.ExtensionData.ReconfigVM_Task($do)