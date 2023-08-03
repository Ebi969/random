Measure-Command{
Param(
    $batchNumber
)

$path = "D:\UserData\Ibraaheem\Scripts\VMWare\Co-Stop"

connect-viserver srv006383.mud.internal.co.za, srv006384.mud.internal.co.za, srv007281.mud.internal.co.za, srv007282.mud.internal.co.za -User Mud\DA003089 -Password BlackLightStudios@123

$stat = 'cpu.costop.summation','cpu.ready.summation', 'mem.swapped.average', 'mem.vmmemctl.average'
$batchNumber += 1
$vmList = Import-Excel "$path\Batches\BatchLoad.xlsx" -WorksheetName "Batch_$batchNumber"

foreach($vm in $vmList){    

    $vm

    $hostDir = (($vm.VMHost).Split("."))[0]
    $clusterName = (Get-VMHost $vm.VMHost).Parent.Name

    $date = (Get-Date).AddHours(-1)
    $format = "$($date.Day)"+"$($date.Month)"+"$($date.Year)"
    $fileName = $vm.Name + "_" + $format

    $dailyOutputPath = "$path\Exports\VMs\$fileName.xlsx"
    $archiveOutputPath = "$path\Exports\Hosts\$hostDir\$fileName.xlsx"

    $vmCollect = @()

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

    if($vmCollect){
        $vmCollect | select * -Unique | Sort-Object Timestamp | Export-Excel $dailyOutputPath -Append -WorksheetName "CoStop" -BoldTopRow -AutoSize -FreezeTopRow -AutoFilter
        $vmCollect | select * -Unique | Sort-Object Timestamp | Export-Excel $archiveOutputPath -Append -WorksheetName "CoStop" -BoldTopRow -AutoSize -FreezeTopRow -AutoFilter
    }
}
}