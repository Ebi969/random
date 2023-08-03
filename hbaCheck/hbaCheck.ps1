$path = "D:\UserData\Ibraaheem\Scripts\VMWare\hbaCheck"

#$vmhosts = Get-Cluster | Where {$_.Name -match "BDC"} | Get-VMHost 
$vmhosts = Get-Content D:\UserData\Ibraaheem\Scripts\VMWare\hbaCheck\list.txt
$collect = @()
 
    foreach ($vmhostName in $vmhosts){
  
        $vmhost = $vmhostName + ".mud.internal.co.za"
        Try{        
            $esx = Get-VMHost $vmhost -ErrorAction Stop
            $esxc = Get-EsxCli -VMHost $esx
    
            $esx.Name

                foreach($hba in (Get-VMHostHba -VMHost $esx -Type "FibreChannel" | where Status -EQ online )){
                     $data = New-Object System.Object
                     $hbaname = $hba.Name
                     $target = ((Get-View $hba.VMhost).Config.StorageDevice.ScsiTopology.Adapter | where {$_.Adapter -eq $hba.Key}).Target
                     $luns = Get-ScsiLun -Hba $hba  -LunType "disk"  -ErrorAction SilentlyContinue
                     $wwn = $esxc.storage.core.adapter.list() | where HBAName -eq $hbaname | select UID
                     $nrPaths = ($target | %{$_.Lun.Count} | Measure-Object -Sum).Sum

                            Write-Host $hba.Device "Targets:" $target.Count "Devices:" $luns.Count "Paths:" $nrPaths "WWN:" $wwn.UID

                     $data | Add-Member -MemberType NoteProperty -Name "HostName" -Value $vmhost
                     $data | Add-Member -MemberType NoteProperty -Name "HBA Name" -Value $hbaname
                     $data | Add-Member -MemberType NoteProperty -Name "HBA Status" -Value $hba.Status
                     $data | Add-Member -MemberType NoteProperty -Name "Targets" -Value $target.Count
                     $data | Add-Member -MemberType NoteProperty -Name "Luns" -Value $luns.Count
                     $data | Add-Member -MemberType NoteProperty -Name "Paths" -Value $nrPaths

                     $collect += $data
                }
        }catch{
            $data = New-Object System.Object
            $data | Add-Member -MemberType NoteProperty -Name "HostName" -Value $vmhost
            $data | Add-Member -MemberType NoteProperty -Name "HBA Name" -Value "Does not Exist"
            $data | Add-Member -MemberType NoteProperty -Name "HBA Status" -Value $null
            $data | Add-Member -MemberType NoteProperty -Name "Targets" -Value $null
            $data | Add-Member -MemberType NoteProperty -Name "Luns" -Value $null
            $data | Add-Member -MemberType NoteProperty -Name "Paths" -Value $null
            $collect += $data
        }

    }

$collect | Export-Excel -Path $path\currentState.xlsx -Append -BoldTopRow -AutoFilter