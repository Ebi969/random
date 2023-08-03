$serverList = Get-Content "D:\UserData\Ibraaheem\Scripts\VMWare\Tier2List\list.txt"

foreach($vmSpaces in $serverList){
    $vm = $vmSpaces.replace(" ","")
    $totT2Storage = $null

Try{

    $hardDisks = Get-VM $vm | Get-HardDisk -ErrorAction Stop

    foreach($disk in $hardDisks){

        $datastoreDisk = ($disk.Filename).Split(" ")      
            
                $datastore = $datastoreDisk[0].Replace("[","")
                $datastore = $datastore.Replace("]","")

                if($datastore -like '*SAP*' -or $datastore -like '*DIA*'){

                    $totT2Storage += $disk.CapacityGB

                }
    }

}catch{
    $totT2Storage = "VM not found"
}

$vm
$tots = [Math]::Round($totT2Storage,2)
$tots

    $out = New-Object PSobject

    $out | Add-Member -MemberType NoteProperty -Name "VM" -Value $vm
    $out | Add-Member -MemberType NoteProperty -Name "Total T2 (GB)" -Value ([MATH]::Round($tots,2))

    $out | Export-Excel -Path "D:\UserData\Ibraaheem\Scripts\VMWare\Tier2List\Tier2Totals.xlsx" -FreezeTopRow -BoldTopRow -Append
}