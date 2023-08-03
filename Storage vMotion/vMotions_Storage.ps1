$migrationList = Import-Excel "D:\UserData\Ibraaheem\Scripts\VMWare\Storage vMotion\MigrationList-T2.xlsx" -WorksheetName "Migration"
$VMsToRelocate = @()
$DStoGoNow = $null
Get-Date
foreach ($row in $migrationList){ 
    if($row.VM -ne $null){
        $VMsToRelocate += $row.VM
    }else{
        $datastoreToGoTo = $row.'Move To'        
            $VMsMovingNow = $VMsToRelocate
            $DStoGoNow = $datastoreToGoTo
        $datastoreToGoTo = $null
        $VMsToRelocate = @()      
    }

    if($DStoGoNow -ne $null){
        $datastoreGO = Get-Datastore $DStoGoNow        
        foreach($vm in $VMsMovingNow){
            "Moving " + $vm + " to " + $datastoreGO
            Get-VM $vm | Move-VM -datastore $datastoreGO
        }
        $VMsMovingNow = @()
        $DStoGoNow = $null
    }
}
Get-Date