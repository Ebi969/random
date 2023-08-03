$path = "D:\UserData\Ibraaheem\Scripts\VMWare\LUN Renaming"

$datastoreList = Get-Datastore | Where-Object {$_.Name -match "CDC02" -and $_.Name -notmatch "local"} | Sort-Object Name
$namesAlreadyExisting = Get-Datastore | Where-Object {$_.Name -match "CDC01"} | Select -ExpandProperty Name
$count = 1

foreach($datastore in $datastoreList){

    $clusters = $datastore | Get-VMHost | Get-Cluster
    
    $replaceName = ($datastore.Name).Replace("CDC02", "CDC01")
    if($datastore.Name -notmatch "scratch"){
        $noIDName = $replaceName.subString(0,$replaceName.length-2)
    }else{
        $noIDName = $replaceName
    }

    foreach($cluster in $clusters){

        $dsOutput = [pscustomobject] @{
            Cluster = $cluster
            NumberClusters = $clusters.count
            CurrentName = $datastore.Name
            UID = $datastore.ExtensionData.Info.Vmfs.Extent.DiskName
            NewName = $noIDName
        }
        $dsOutput | Export-Excel "$path\NewNameList.xlsx" -Append -BoldTopRow -AutoSize -AutoFilter -FreezeTopRow

    }

}

$importList = Import-Excel "$path\NewNameList.xlsx" | Where-Object {$_.CurrentName -notmatch "scratch"} | Sort-Object Cluster

$clientID = @("SLM", "STM", "SKY", "SEM")
$storageID = @("DIA01", "QUA01", "QUA02", "EME01", "RUB01")
$mirrorID = @("MSP", "NML", "SRM")

foreach($client in $clientID){
    foreach($storageType in $storageID){
        foreach($mirror in $mirrorID){
            $count = 1
            foreach($dat in $importList | Where-Object {$_.CurrentName -match $client -and $_.CurrentName -match $storageType -and $_.CurrentName -match $mirror}){

                if($count -lt 10){
                    $newName = $dat.NewName + "0" + $count
                }else{
                    $newName = $dat.NewName + $count
                }

                if($namesAlreadyExisting -contains $dat.CurrentName){
                        $alreadyExists = $true
                    }else{
                        $alreadyExists = $false
                }   
                          
                while($alreadyExists){  
                    $count += 1               
                    if($count -lt 10){
                        $newName = $dat.NewName + "0" + $count
                    }else{
                        $newName = $dat.NewName + $count
                    }

                    if($namesAlreadyExisting -contains $dat.CurrentName){
                            $alreadyExists = $true
                        }else{
                            $alreadyExists = $false
                    }                
                }

                $dsOutput = [pscustomobject] @{
                    Cluster = $dat.Cluster
                    NumberClusters = $dat.NumberClusters
                    CurrentName = $dat.CurrentName
                    UID = $dat.UID
                    NewName = $newName
                }
                $dsOutput | Export-Excel "$path\FinalList.xlsx" -Append -BoldTopRow -AutoSize -AutoFilter -FreezeTopRow            

                $count += 1
            }
        }

    }    

}