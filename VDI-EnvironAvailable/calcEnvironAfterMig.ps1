$test = Import-Excel -Path D:\UserData\Ibraaheem\Scripts\VMWare\VDI-EnvironAvailable\calcEnviron.xlsx

$BDColdCPU = $null
$BDColdMem = $null
$BDCNewCpu = $null
$BDCNewMem = $null
$CDColdCPU = $null
$CDColdMem = $null
$CDCNewCpu = $null
$CDCNewMem = $null

$newBDCTotCpu = $null
$newBDCTotMem = $null

$newCDCTotCpu = $null
$newCDCTotMem = $null

foreach($cluster in $test){

    if($cluster.VC -eq "srv008146"){

        if($cluster.Cluster -match "BDC"){
            $BDCNewCpu += $cluster.VMtotCPU
            $BDCNewMem += $cluster.VMtotMem
            $newBDCTotCpu = $cluster.hostotCpu
            $newBDCTotMem = $cluster.hosttotMem
        }elseif($cluster.Cluster -match "CDC"){
            $CDCNewCpu += $cluster.VMtotCPU
            $CDCNewMem += $cluster.VMtotMem
            $newCDCTotCpu = $cluster.hostotCpu
            $newCDCTotMem = $cluster.hosttotMem
        }

    }else{

        if($cluster.Cluster -match "BDC"){
            $BDColdCpu += $cluster.VMtotCPU
            $BDColdMem += $cluster.VMtotMem
        }elseif($cluster.Cluster -match "CDC"){
            $CDColdCpu += $cluster.VMtotCPU
            $CDColdMem += $cluster.VMtotMem
        }

    }

}

$afterBDCTotCPU = $BDColdCPU + $BDCNewCpu
$afterBDCTotMem = $BDColdMem + $BDCNewMem
$afterCDCTotCPU = $CDColdCPU + $CDCNewCpu
$afterCDCTotMem = $CDColdMem + $CDCNewMem

#Ratio Calculations



$cpuCDCRatio = $afterCDCTotCPU / $newCDCTotCpu
$memoryCDCPercent = ($afterCDCTotMem / $newCDCTotMem).ToString("P")


$finalOutput = @()
foreach($cluster in $test | Where {$_.vc -eq "srv008146"}){

    if($cluster.Cluster -match "BDC"){
        $cpuBDCRatio = $afterBDCTotCPU / $newBDCTotCpu
        $memoryBDCPercent = ($afterBDCTotMem / $newBDCTotMem).ToString("P")
        $availableCPU = ($newBDCTotCpu * 5) - $afterBDCTotCPU
        $availableMem = ($newBDCTotMem * 0.9) - $afterBDCTotMem

        $outPut = [pscustomobject] @{
            vc = $cluster.VC
            cluster = $cluster.Cluster
            CurrentCPURatio = $cluster.cpuRation
            CurrentMerPerc = $cluster.memPerc
            AfterMigCPURatio = $cpuBDCRatio
            AfterMigMemPerc = $memoryBDCPercent
            AvailableCPUAfter = $availableCPU
            AvailableMemAfter = $availableMem
        }

    }elseif($cluster.Cluster -match "CDC"){
        $cpuCDCRatio = $afterCDCTotCPU / $newCDCTotCpu
        $memoryCDCPercent = ($afterCDCTotMem / $newCDCTotMem).ToString("P")
        $availableCPU = ($newCDCTotCpu * 5) - $afterCDCTotCPU
        $availableMem = ($newCDCTotMem * 0.9) - $afterCDCTotMem

        $outPut = [pscustomobject] @{
            vc = $cluster.VC
            cluster = $cluster.Cluster
            CurrentCPURatio = $cluster.cpuRation
            CurrentMerPerc = $cluster.memPerc
            AfterMigCPURatio = $cpuCDCRatio
            AfterMigMemPerc = $memoryCDCPercent
            AvailableCPUAfter = $availableCPU
            AvailableMemAfter = $availableMem
        }

    }

    $finalOutput += $outPut

}

$finalOutput | Export-Excel -Path "D:\UserData\Ibraaheem\Scripts\VMWare\VDI-EnvironAvailable\calcEnvironAfterMig.xlsx" -Append -AutoSize -AutoFilter -BoldTopRow