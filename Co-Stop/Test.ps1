Measure-Command{

$path = "D:\UserData\Ibraaheem\Scripts\VMWare\Co-Stop"

$stat = 'cpu.costop.summation','cpu.ready.summation', 'mem.swapped.average', 'mem.vmmemctl.average'
#$stat = 'cpu.costop.summation'
$vmList = Get-VM | Where {$_.PowerState -eq "PoweredOn"} | Select -First 1

foreach($vm in $vmList){    

    $hostDir = (($vm.VMHost.Name).Split("."))[0]
    $clusterName = $vm.VMHost.Parent
    $fileName = $vm.Name + "_" + $(Get-Date -Format ddMMyyyy)
    $dailyOutputPath = "$path\Exports\VMs\$fileName.xlsx"
    $archiveOutputPath = "$path\Exports\Hosts\$hostDir\$fileName.xlsx"

    $vm

    $vmCollect = @()
    $hostDir = (($vm.VMHost.Name).Split("."))[0]
    $clusterName = $vm.VMHost.Parent

    $statistics = Get-Stat -Entity $vm.Name -Realtime -Stat $stat -ErrorAction SilentlyContinue | Group-Object TimeStamp | %{

            $costopSummation = [MATH]::Round((($_.Group | Where {$_.MetricId -eq "cpu.costop.summation" } | Select -ExpandProperty Value) | Measure-Object -Average).Average ,2)
            $readySumamation = [MATH]::Round((($_.Group | Where {$_.MetricId -eq "cpu.ready.summation" } | Select -ExpandProperty Value) | Measure-Object -Average).Average , 2)
            $memSwappedAverage = [MATH]::Round((($_.Group | Where {$_.MetricId -eq "mem.swapped.average" } | Select -ExpandProperty Value) | Measure-Object -Average).Average , 2)
            $memballoonedAverage = [MATH]::Round((($_.Group | Where {$_.MetricId -eq "mem.vmmemctl.average" } | Select -ExpandProperty Value) | Measure-Object -Average).Average ,2)

    ### CPU READY % calculation
    $cpuReadyPerc = ($readySumamation/ (20 * 1000)) * 100

        $output =  New-Object PSObject -Property ([ordered]@{
            "vmName" = $vm.Name
            "vmHost" = $hostDir
            "cluster" = $clusterName
            "TimeStamp" = $_.Name
            <##>
            "cpu.costop.summation" = $costopSummation
            "cpu.ready.summation" = $readySumamation
            "cpu.ready.%" = $cpuReadyPerc
            "mem.swapped.average" = $memSwappedAverage
            "mem.vmmemctl.average" = $memballoonedAverage
            <##>
        })
        $vmCollect += $output
    }

    $vmCollect | Export-Excel $dailyOutputPath -Append -WorksheetName "CoStop" -BoldTopRow -AutoSize -FreezeTopRow -AutoFilter
}
}