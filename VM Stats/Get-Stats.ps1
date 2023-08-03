$cpuStats = Get-Stat -Entity SRV005879 -Start 07/11/2019 -Finish 07/18/2019 -Stat net.usage.average -IntervalMins 75
[Array]::Reverse($cpuStats)
$cpuStats