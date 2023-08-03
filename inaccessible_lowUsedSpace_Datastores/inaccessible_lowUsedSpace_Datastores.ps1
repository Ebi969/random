Get-Datastore | Select Name, 
                @{l="Free Space(GB)"; e={[MATH]::Round($_.FreeSpaceGB, 2)}}, 
                @{l="Capacity(GB)"; e={[MATH]::Round($_.CapacityGB, 2)}}, 
                @{l="Used Spaced(GB)"; e={[MATH]::Round(($_.CapacityGB - $_.FreeSpaceGB), 2)}} | Where {$_.Name -notlike "*local*"} | Export-Excel -Path "D:\UserData\Ibraaheem\Scripts\VMWare\inaccessible_lowUsedSpace_Datastores\Export.xlsx" -WorksheetName "Low Used Datastores" -AutoSize -AutoFilter


Get-Datastore | Select * Parent, Name, Accessible | Where {$_.Accessible -like "False"} | Export-Excel -Path "D:\UserData\Ibraaheem\Scripts\VMWare\inaccessible_lowUsedSpace_Datastores\Export.xlsx" -WorksheetName "Inaccessible Datastores" -AutoSize -AutoFilter
