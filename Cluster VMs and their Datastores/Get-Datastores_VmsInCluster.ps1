$VmsInCluster = Get-Cluster SKY-BDC-GOLD-03-Database | Get-VM | Select Name
$blank = ""

$path = "D:\UserData\Ibraaheem\Scripts\VMWare\Cluster VMs and their Datastores"

if(Test-Path $path\output.csv){
    Remove-Item $path\output.csv
}


# BLOCK 1: Create and open runspace pool, setup runspaces array with min and max threads
$pool = [RunspaceFactory]::CreateRunspacePool(1, [int]$env:NUMBER_OF_PROCESSORS+1)
$pool.ApartmentState = "MTA"
$pool.Open()
$runspaces = @()
    
# BLOCK 2: Create reusable scriptblock. This is the workhorse of the runspace. Think of it as a function.
$scriptblock = {
    Param (
        [String]$blank,
        [String]$vmName,
        $disk
    )
     
#Stuff to do

        $Datastore = Get-VM $vmName | Get-HardDisk -name $disk | Select -ExpandProperty FileName
        $capacity = Get-VM $vmName | Get-HardDisk -name $disk | Select -ExpandProperty CapacityGB

        $DataSplit = $Datastore.Split(" ")
        $DataReplaced = $DataSplit[0].Replace("[","")
        $location = $DataReplaced.Replace("]","")

        $out = New-Object PSObject

        $out | Add-Member -MemberType NoteProperty -Name "VM Name" -Value $vmName
        $out | Add-Member -MemberType NoteProperty -Name "Disk Number" -Value $disk
        $out | Add-Member -MemberType NoteProperty -Name "Disk Capacity(GB)" -Value $capacity
        $out | Add-Member -MemberType NoteProperty -Name "Datastore" -Value $location
    
    # return whatever you want, or don't.
    return $out
}
 
# BLOCK 3: Create runspace and add to runspace pool
foreach($vm in $VmsInCluster){

    $vmName = $vm.Name

    $Disks = Get-VM $vmName | Get-HardDisk | Select -ExpandProperty Name

        foreach($disk in $Disks){

            $runspace = [PowerShell]::Create()
            $null = $runspace.AddScript($scriptblock)
            $null = $runspace.AddArgument($blank)
            $null = $runspace.AddArgument($path)
            $null = $runspace.AddArgument($vmName)
            $null = $runspace.AddArgument($disk)
            $runspace.RunspacePool = $pool
 
# BLOCK 4: Add runspace to runspaces collection and "start" it
        # Asynchronously runs the commands of the PowerShell object pipeline
            $runspaces += [PSCustomObject]@{ Pipe = $runspace; Status = $runspace.BeginInvoke() }
        }
}
 
# BLOCK 5: Wait for runspaces to finish
 while ($runspaces.Status.IsCompleted -notcontains $true) {}
 
# BLOCK 6: Clean up
foreach ($runspace in $runspaces ) {
    # EndInvoke method retrieves the results of the asynchronous call
    $results = $runspace.Pipe.EndInvoke($runspace.Status)
    $results
    $results | Export-Csv $path\output.csv -NoTypeInformation -Append
    $runspace.Pipe.Dispose()
}
    
$pool.Close() 
$pool.Dispose()
 
# Bonus block 7
# Look at $results to see any errors or whatever was returned from the runspaces
