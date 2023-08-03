$migrationList = Import-Excel "D:\UserData\Ibraaheem\Scripts\VMWare\Storage vMotion\MigrationList-T2.xlsx" -WorksheetName "Migration"

foreach ($row in $migrationList){
    if($row.VM -ne $null){
        ""
        $row.VM
         
        Get-VM $row.VM | Get-Datastore | Select -ExpandProperty Name
        #Get-VM $row.VM | Get-Cluster | Select -ExpandProperty Name
    }
}