$vmList = Get-Content "D:\UserData\Ibraaheem\Scripts\VMWare\currentDCandMirroredStorage\vmList.txt"

$finaloutput = @()

foreach($server in $vmList){

    $currentCluster = $null
    $datastores = $null
    $location = $null
    $mirrored = $null

    $vm = $server.replace(" ", "")

    Try{
        
        $datastores = Get-VM $vm -ErrorAction Stop | Get-Datastore
        
        $datastores = $datastores -join ", "

        if($datastores -like "*MSP*" -or $datastores -like "*MAP*"){
            $mirrored = "Mirrored Storage"
        }else{
            $mirrored = "Not Mirrored"
        }

        $currentCluster = Get-VM $vm -ErrorAction Stop | Get-Cluster

        $clusterSplit = ($currentCluster.Name).Split("-")
            $location = $clusterSplit[1]

    }catch{
        $location = "VM not found"
        $mirrored = "N/A"
    }

    $output = [pscustomobject] @{
        'VM' = $vm
        'Current Cluster' = $currentCluster.Name
        'Current DS' = $datastores
        'Location' = $location
        'On Mirrored' = $mirrored
    } | Export-excel -path "D:\UserData\Ibraaheem\Scripts\VMWare\currentDCandMirroredStorage\Export.xlsx" -Append -BoldTopRow -AutoSize -AutoFilter
    $finaloutput += $output
}

$finaloutput | Out-GridView