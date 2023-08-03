Get-VM | foreach {

    $_ | Get-NetworkAdapter | Foreach {
        $_ | Select Parent, Name, Type, @{n="Direct Path I/O"; e={$_.ExtensionData.UptCompatibilityEnabled}}
    }

} | Export-Excel -Path "D:\UserData\Ibraaheem\Scripts\VMWare\vmNicStuff\dpioNics.xlsx" -Append -AutoFilter -AutoSize -BoldTopRow