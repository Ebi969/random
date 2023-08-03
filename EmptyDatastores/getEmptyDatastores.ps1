$collect = @()
Get-Datastore | Where {$_.Name -notmatch "local|scratch|stag"} | Foreach{
    $vmList = $_ | Get-VM
    if($vmList -eq $null){
        $collect += $_
    }
    $vmList = $null
}

$collect | Sort-Object Name | ft -AutoSize