$path = "D:\UserData\Ibraaheem\Scripts\VMWare\RevisedMigrationProjects"
$actionPlanPath = "$path\ActionPlans\VDI"

$allFiles = Get-ChildItem $actionPlanPath\*.xlsx

foreach($clusterFile in $allFiles){
    $fileName = $clusterFile.Name
    $clusterName = $fileName.Replace(".xlsx","")
    
    Try{$lunRequests = Import-Excel "$actionPlanPath\$fileName" -WorksheetName "Lun Requests" -ErrorAction Stop}catch{}

    $indivNMPLuns = $lunRequests | Where-Object {$_.Type -eq "NML" -and $_.ForVM -ne "Shared"}
    $indivMSPLuns = $lunRequests | Where-Object {$_.Type -eq "MSP" -and $_.ForVM -ne "Shared"}
    $indivMAPLuns = $lunRequests | Where-Object {$_.Type -eq "MAP" -and $_.ForVM -ne "Shared"}
    $indivVDILuns = $lunRequests | Where-Object {$_.Type -eq "VDI" -and $_.ForVM -ne "Shared"}

    $clusterSplit = $clusterName.split("-")

    $lunClient = $clusterSplit[0]
    $lunLocation = $clusterSplit[1]

    if($lunLocation -match "BDC"){
        $lunLocCode = "BDC01"
    }elseif($lunLocation -match "CDC"){
        $lunLocCode = "CDC01"
    }else{
        $lunLocCode = "STR01"
    }

    
    if($lunSLA -match "GOLD|GLD"){
        $tier = "DIA31"
    }else{
        $tier = "DIA31"
    }    

    $clusterName
    $newLuns = @()
    foreach($newLun in $lunRequests){
        $type = $newLun.Type

        if($newLun.ForVM -eq "Shared"){
            $identity = "CMPTHN"
            $lunName = $lunClient + "-" + $lunLocCode + "-" + $tier + "-" + $type + "-" + $identity
        }else{
            $identity = ($newLun.ForVM).Substring(($newLun.ForVM).Length - 5)
            $lunName = $lunClient + "-" + $identity + "-" + $tier + "-" + $type + "-" + "CMPTHN"
        }

        $newLun.Name = $lunName
        $newLuns += $newLun
    }
        $newLuns | Export-Excel "$actionPlanPath\$fileName" -WorksheetName "Lun Requests" -BoldTopRow -FreezeTopRow -AutoFilter
}