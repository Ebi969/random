Get-Date
$path = "D:\UserData\Ibraaheem\Scripts\VMWare\Co-Stop"

connect-viserver srv006383.mud.internal.co.za, srv006384.mud.internal.co.za, srv007281.mud.internal.co.za, srv007282.mud.internal.co.za -User Mud\DA003089 -Password BlackLightStudios@123

$vms = Get-VM | Where {$_.PowerState -eq "PoweredOn"}

$scriptPath = "$path\Get-Costop.ps1"

$totalVMs = $vms.Count
$numPerBatch = $totalVMs / 5

    $splitList = for ($i = 0; $i -lt $totalVMs; $i += $numPerBatch) {
        ,@($vms[$i..($i+($numPerBatch-1))]);
    }

if(Test-Path "$path\Batches\BatchLoad.xlsx"){
    Remove-Item "$path\Batches\BatchLoad.xlsx"
}

$splitList[0] | Select Name, NumCPU, MemoryGB, VMHost | Export-Excel "$path\Batches\BatchLoad.xlsx" -WorksheetName "Batch_1" -BoldTopRow -AutoSize
$splitList[1] | Select Name, NumCPU, MemoryGB, VMHost | Export-Excel "$path\Batches\BatchLoad.xlsx" -WorksheetName "Batch_2" -BoldTopRow -AutoSize
$splitList[2] | Select Name, NumCPU, MemoryGB, VMHost | Export-Excel "$path\Batches\BatchLoad.xlsx" -WorksheetName "Batch_3" -BoldTopRow -AutoSize
$splitList[3] | Select Name, NumCPU, MemoryGB, VMHost | Export-Excel "$path\Batches\BatchLoad.xlsx" -WorksheetName "Batch_4" -BoldTopRow -AutoSize
$splitList[4] | Select Name, NumCPU, MemoryGB, VMHost | Export-Excel "$path\Batches\BatchLoad.xlsx" -WorksheetName "Batch_5" -BoldTopRow -AutoSize

for($i = 0; $i -lt 5; $i++){
    start-process powershell -ArgumentList "$scriptPath $i"
    Start-Sleep -Seconds 2
}