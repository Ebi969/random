$path = "D:\UserData\Ibraaheem\Scripts\VMWare\RevisedMigrationProjects"
$actionPlanPath = "$path\ActionPlans\VDI"

$allFiles = Get-ChildItem $actionPlanPath\*.xlsx

    $blankRow = [pscustomobject] @{}
    $noInfoRow = [pscustomobject] @{"Comment" =  "Nothing needed here"}

$lunsToRequest = @()
$sharedSize = "4"
$tier = "3"

foreach($clusterFile in $allFiles){
    $fileName = $clusterFile.Name

    $NMPShared = $null
    $NMPIndiv = $null
    $MSPShared = $null
    $MSPIndiv = $null
    $MAPShared = $null
    $MAPIndiv = $null
    $VDIShared = $null
    $VDIIndiv = $null

    $lunsToRequest = @()

    Try{$NMPShared = Import-Excel "$actionPlanPath\$fileName" -WorksheetName "NML-Shared" -ErrorAction Stop}catch{}
    Try{$NMPIndiv = Import-Excel "$actionPlanPath\$fileName" -WorksheetName "NML-Indiv" -ErrorAction Stop}catch{}
    Try{$MSPShared = Import-Excel "$actionPlanPath\$fileName" -WorksheetName "MSP-Shared" -ErrorAction Stop}catch{}
    Try{$MSPIndiv = Import-Excel "$actionPlanPath\$fileName" -WorksheetName "MSP-Indiv" -ErrorAction Stop}catch{}
    Try{$MAPShared = Import-Excel "$actionPlanPath\$fileName" -WorksheetName "MAP-Shared" -ErrorAction Stop}catch{}
    Try{$MAPIndiv = Import-Excel "$actionPlanPath\$fileName" -WorksheetName "MAP-Indiv" -ErrorAction Stop}catch{}
    Try{$VDIShared = Import-Excel "$actionPlanPath\$fileName" -WorksheetName "VDI-Shared" -ErrorAction Stop}catch{}
    Try{$VDIIndiv = Import-Excel "$actionPlanPath\$fileName" -WorksheetName "VDI-Indiv" -ErrorAction Stop}catch{}
    
    if($NMPShared){
        $totalNMPSharedCap = ($NMPShared | Measure-Object -Sum 'Migrate(GB)').Sum
        [int]$numNMPSharedLuns = [math]::Ceiling($totalNMPSharedCap / (4096 * 0.8))
            While($numNMPSharedLuns -ne 0){
                $lunReq = [pscustomobject] @{
                    Name = ""
                    Size = $sharedSize
                    Tier = $tier
                    Type = "NML"
                    ForVM = "Shared"
                }
                $numNMPSharedLuns -= 1
                $lunsToRequest += $lunReq
            }
    }    
    if($NMPIndiv){
        foreach($lun in $NMPIndiv | Select-Object * -Unique){
            $lunReq = [pscustomobject] @{
                Name = ""
                Size = $lun.'Lun Needed'
                Tier = $tier
                Type = "NML"
                ForVM = $lun.VMName
            }
            $lunsToRequest += $lunReq
        }        
    }

    if($MSPShared){
        $totalMSPSharedCap = ($MSPShared | Measure-Object -Sum 'Migrate(GB)').Sum
        [int]$numMSPSharedLuns = [math]::Ceiling($totalMSPSharedCap / (4096 * 0.8))
            While($numMSPSharedLuns -ne 0){
                $lunReq = [pscustomobject] @{
                    Name = ""
                    Size = $sharedSize
                    Tier = $tier
                    Type = "MSP"
                    ForVM = "Shared"
                }
                $numMSPSharedLuns -= 1
                $lunsToRequest += $lunReq
            }
    }    
    if($MSPIndiv){
        foreach($lun in $MSPIndiv | Select-Object * -Unique){
            $lunReq = [pscustomobject] @{
                Name = ""
                Size = $lun.'Lun Needed'
                Tier = $tier
                Type = "MSP"
                ForVM = $lun.VMName
            }
            $lunsToRequest += $lunReq
        }        
    }

    if($MAPShared){
        $totalMAPSharedCap = ($MAPShared | Measure-Object -Sum 'Migrate(GB)').Sum
        [int]$numMAPSharedLuns = [math]::Ceiling($totalMAPSharedCap / (4096 * 0.8))
            While($numMAPSharedLuns -ne 0){
                $lunReq = [pscustomobject] @{
                    Name = ""
                    Size = $sharedSize
                    Tier = $tier
                    Type = "MAP"
                    ForVM = "Shared"
                }
                $numMAPSharedLuns -= 1
                $lunsToRequest += $lunReq
            }
    }    
    if($MAPIndiv){
        foreach($lun in $MAPIndiv | Select-Object * -Unique){
            $lunReq = [pscustomobject] @{
                Name = ""
                Size = $lun.'Lun Needed'
                Tier = $tier
                Type = "MAP"
                ForVM = $lun.VMName
            }
            $lunsToRequest += $lunReq
        }        
    }

    if($VDIShared){
        $totalVDISharedCap = ($VDIShared | Measure-Object -Sum 'Migrate(GB)').Sum
        [int]$numVDISharedLuns = [math]::Ceiling($totalVDISharedCap / (4096 * 0.8))
            While($numVDISharedLuns -ne 0){
                $lunReq = [pscustomobject] @{
                    Name = ""
                    Size = $sharedSize
                    Tier = $tier
                    Type = "VDI"
                    ForVM = "Shared"
                }
                $numVDISharedLuns -= 1
                $lunsToRequest += $lunReq
            }
    }    
    if($VDIIndiv){
        foreach($lun in $VDIIndiv | Select-Object * -Unique){
            $lunReq = [pscustomobject] @{
                Name = ""
                Size = $lun.'Lun Needed'
                Tier = $tier
                Type = "VDI"
                ForVM = $lun.VMName
            }
            $lunsToRequest += $lunReq
        }        
    }

    $lunsToRequest | Export-Excel -Path "$actionPlanPath\$fileName" -WorksheetName "Lun Requests" -Append -BoldTopRow -FreezeTopRow -AutoFilter
}