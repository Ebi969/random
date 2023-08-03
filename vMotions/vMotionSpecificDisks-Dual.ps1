$path = "D:\UserData\Ibraaheem\Scripts\VMWare\RevisedMigrationProjects\vMotions\"
$migrateList = Import-Excel -Path $path\MigrationList-Dual.xlsx -WorksheetName "To Migrate"

foreach($row in $migrateList){
    $vmName = $row.VMName 
    $targetDatastore = $row.AssignTO

    $vm = Get-VM $vmName 
    $allDisks = Get-HardDisk -VM $vm
    $destinationDS = Get-Datastore $targetDatastore

    $selectDisks = $allDisks | Where {$_.FileName -match "DIA|SAP"}
    #$selectDisks = $allDisks | Where {$_.FileName -match "RUB|EME|QUA"}
    "Move " + $row.VMName + " - " + $($selectDisks.Name -join ", ") + " To " + $targetDatastore + "`n"

    $selectDisks | Move-HardDisk -Datastore $destinationDS -Confirm:$false
}