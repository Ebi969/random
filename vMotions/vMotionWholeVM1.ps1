$path = "D:\UserData\Ibraaheem\Scripts\VMWare\RevisedMigrationProjects\vMotions\"
$migrateList = $null
$migrateList = Import-Excel -Path $path\MigrationList1.xlsx -WorksheetName "To Migrate"

foreach($row in $migrateList){
    $vmName = $row.VMName 
    $targetDatastore = $row.AssignTO

    $vm = Get-VM $vmName
    $destinationDS = Get-Datastore $targetDatastore

    "Move " + $row.VMName + " To " + $destinationDS + "`n"
    #$vm.Name
    #$vm | Get-Datastore | Select -ExpandProperty Name
    $vm | Move-VM -datastore $destinationDS
}