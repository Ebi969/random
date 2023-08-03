<#
FS9150 BDC -     600507681081012ED
FS9150 CDC -    600507681081802D2
V5030 BDC -      600507638081041AD8
V5030 CDC -       600507638081041718
#>

$outputPath = "D:\UserData\Ibraaheem\Scripts\VMWare\DatastoresUUID\uuidFind.xlsx"

$naaList = @("600000E00D29000000293567004B0000", "600000E00D29000000293567004C0000", "600000E00D2900000029377B000C0000", "600000E00D290000002932FA000C0000", "60050768108100D918000000000000BA", "60050768108100D918000000000000BB", "60050768018200EC0000000000000661")

foreach($naa in $naaList){
    get-datastore | Where-Object {$_.extensiondata.info.vmfs.extent.diskname -like “*$naa*”} | foreach{
            $DSName = $_.Name
            $_ | select @{n="Name"; e={$_.Name}}, @{n="FreeSpaceGB"; e={[MATH]::Round($_.FreeSpaceGB,2)}}, @{n="CapacityGB"; e={[MATH]::Round($_.CapacityGB,2)}}, @{n="UUid"; e={$_.extensiondata.info.vmfs.extent.diskname}} | Export-Excel -Path $outputPath -WorksheetName "STMDSOnly" -Append -BoldTopRow -AutoFilter -AutoSize
    }
}