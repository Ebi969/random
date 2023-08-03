$path = "D:\UserData\Ibraaheem\Scripts\VMWare\LUN Renaming"

$importList = Import-Excel $path\FinalList.xlsx -WorksheetName "Datastore"

$lun = $null
foreach($lun in $importList){

    $nowName = $lun.CurrentName
    $newDatName = $lun.'NewName'
    $clusterName =  $lun.Cluster

    if($clusterName -notmatch $tempClusterName){
        $getCluster = Get-Cluster $tempClusterName
        ### Rescan Cluster Storage ###
        $getCluster | Get-VMHost | Get-VMHostStorage -RescanAllHba -RescanVmfs
    }
    
    ### Rename Datastore ###
    $dsTarget = Get-datastore -name $nowName 
    #$dsTarget.Name
    $dsTarget | Set-datastore -Name $newDatName

    $tempClusterName = $clusterName
}

    