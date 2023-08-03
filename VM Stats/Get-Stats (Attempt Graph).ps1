#Server list
$path = "C:\Users\DA003089\Desktop\Scripts\VMWare\VM Stats"
$outputPath = "C:\Users\DA003089\Desktop\Scripts\VMWare\VM Stats\Output"

$Servers = Get-Content $path\ServerList.txt
$startDate = (Get-Date).AddDays(-1) #mm/dd/yyyy
$endDate = (Get-Date).AddDays(1) #mm/dd/yyyy

$blankcol = ""

<#
if(Test-Path $path\Output){
    Remove-Item $path\Output -Recurse
    Start-Sleep -Seconds 5
    New-Item -Path $path\Output -ItemType Directory 
}else{
    New-Item -Path $path\Output -ItemType Directory
}

#remove Output Files
if(Test-Path $outputpath\memUsage.csv){
    Remove-Item $outputpath\memUsage.csv
}
if(Test-Path $outputpath\cpuUsage.csv){
    Remove-Item $outputpath\cpuUsage.csv
}
if(Test-Path $outputpath\diskUsage.csv){
    Remove-Item $outputpath\diskUsage.csv
}
#>
###############################################################################################################################################################

[void][Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
[void][Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms.DataVisualization')

###############################################################################################################################################################

#Looping servers
Foreach($Server in $Servers)
{ 

###############################################################################################################################################################
#Memory charts
###############################################################################################################################################################
  
$MemChart = New-Object -TypeName System.Windows.Forms.DataVisualization.Charting.Chart
$MemChart.Size = '1400,500'
 
$MemChartArea = New-Object -TypeName System.Windows.Forms.DataVisualization.Charting.ChartArea
$MemChartArea.AxisX.LabelStyle.Enabled = $true
$MemChartArea.AxisX.LabelStyle.Angle = 90
$MemChart.ChartAreas.Add($MemChartArea)
$MemChart.Series.Add('Memory')
$MemChartArea.AxisY.Title = 'Percentage %'
$MemChartArea.AxisY.Interval = '25'
$MemChartArea.AxisY.Maximum = '100'
$MemChartArea.AxisX.Title = 'Period'
$MemChartArea.AxisX.Interval = '24'

$MemChart.Series['Memory'].ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Line

    $Server = $Server.Trim()
    $memtimes = Get-Stat -Entity $Server -Start 11/29/2019 -Finish $endDate -Stat mem.usage.average -IntervalMins 30
    [Array]::Reverse($memtimes)
    Write-Host "Processing Memory Usage on $Server" -ForegroundColor Green
    
        $memOut = New-Object PSObject

        $memOut | Add-Member -MemberType NoteProperty -Name "Server" -Value $Server
        $memOut | Add-Member -MemberType NoteProperty -Name "Time Stamp" -Value $blankcol
        $memOut | Add-Member -MemberType NoteProperty -Name "Usage (%)" -Value $blankcol

        $memOut | Export-Csv "$outputpath\($server)memUsage.csv" -NoTypeInformation -Append

    foreach($memtime in $memtimes){
        $actualmemTime = $memtime.TimeStamp
        $actualMem = $memtime.Value

        $Value = $MemChart.Series['Memory'].Points.AddXY("$actualmemTime","$actualMem")

        $memOut = New-Object PSObject

        $memOut | Add-Member -MemberType NoteProperty -Name "Server" -Value $Server
        $memOut | Add-Member -MemberType NoteProperty -Name "Time Stamp" -Value $actualmemTime
        $memOut | Add-Member -MemberType NoteProperty -Name "Usage (%)" -Value $actualMem

        $memOut | Export-Csv "$outputpath\($server)memUsage.csv" -NoTypeInformation -Append
    }

        $Title = New-Object -TypeName System.Windows.Forms.DataVisualization.Charting.Title
        $MemChart.Titles.Add($Title)
        $MemChart.Titles[0].Text = "Memory usage ($Server)"
 
        #Saving PNG file on desktop
        $MemChart.SaveImage("$outputpath\($Server)MemoryUsage.png", "PNG")

###############################################################################################################################################################
#CPU charts
###############################################################################################################################################################
 
$cpuChart = New-Object -TypeName System.Windows.Forms.DataVisualization.Charting.Chart
$cpuChart.Size = '1400,500'
 
$cpuChartArea = New-Object -TypeName System.Windows.Forms.DataVisualization.Charting.ChartArea
$cpuChartArea.AxisX.LabelStyle.Enabled = $true
$cpuChartArea.AxisX.LabelStyle.Angle = 90
$cpuChart.ChartAreas.Add($cpuChartArea)
$cpuChart.Series.Add('CPU')
$cpuChartArea.AxisY.Title = 'Percentage %'
$cpuChartArea.AxisY.Interval = '25'
$cpuChartArea.AxisY.Maximum = '100'
$cpuChartArea.AxisX.Title = 'Period'
$cpuChartArea.AxisX.Interval = '24'

$cpuChart.Series['CPU'].ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Line

    $Server = $Server.Trim()
    $cputimes = Get-Stat -Entity $Server -Start 11/29/2019 -Finish $endDate -Stat cpu.usage.average -IntervalMins 30
    [Array]::Reverse($cputimes)
    Write-Host "Processing CPU Usage on $Server" -ForegroundColor Green
    
        $cpuOut = New-Object PSObject

        $cpuOut | Add-Member -MemberType NoteProperty -Name "Server" -Value $Server
        $cpuOut | Add-Member -MemberType NoteProperty -Name "Time Stamp" -Value $blankcol
        $cpuOut | Add-Member -MemberType NoteProperty -Name "Usage (%)" -Value $blankcol

        $cpuOut | Export-Csv "$outputpath\($server)cpuUsage.csv" -NoTypeInformation -Append

    foreach($cputime in $cputimes){
        $actualcpuTime = $cputime.TimeStamp
        $actualCpu = $cputime.Value

        $Value = $cpuChart.Series['CPU'].Points.AddXY("$actualcpuTime","$actualCpu")
                
        $cpuOut = New-Object PSObject

        $cpuOut | Add-Member -MemberType NoteProperty -Name "Server" -Value $Server
        $cpuOut | Add-Member -MemberType NoteProperty -Name "Time Stamp" -Value $actualcpuTime
        $cpuOut | Add-Member -MemberType NoteProperty -Name "Usage (%)" -Value $actualCpu

        $cpuOut | Export-Csv "$outputpath\($server)cpuUsage.csv" -NoTypeInformation -Append
    }

        $Title = New-Object -TypeName System.Windows.Forms.DataVisualization.Charting.Title
        $cpuChart.Titles.Add($Title)
        $cpuChart.Titles[0].Text = "CPU usage ($Server)"
 
        #Saving PNG file on desktop
        $cpuChart.SaveImage("$outputpath\($Server)cpuUsage.png", "PNG")

###############################################################################################################################################################
#Disk charts
###############################################################################################################################################################

$diskChart = New-Object -TypeName System.Windows.Forms.DataVisualization.Charting.Chart
$diskChart.Size = '1400,500'
 
$diskChartArea = New-Object -TypeName System.Windows.Forms.DataVisualization.Charting.ChartArea
$diskChartArea.AxisX.LabelStyle.Enabled = $true
$diskChartArea.AxisX.LabelStyle.Angle = 90
$diskChart.ChartAreas.Add($diskChartArea)
$diskChart.Series.Add('Disk')
#$diskChartArea.AxisY.Maximum = '20000'
$diskChartArea.AxisY.Title = 'KBps'
$diskChartArea.AxisX.Title = 'Period'
$diskChartArea.AxisX.Interval = '24'

$diskChart.Series['Disk'].ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Line

    $Server = $Server.Trim()
    $disktimes = Get-Stat -Entity $Server -Start 11/29/2019 -Finish $endDate -Stat disk.usage.average -IntervalMins 30
    [Array]::Reverse($disktimes)
    Write-Host "Processing Disk Usage on $Server" -ForegroundColor Green

        $diskOut = New-Object PSObject

        $diskOut | Add-Member -MemberType NoteProperty -Name "Server" -Value $Server
        $diskOut | Add-Member -MemberType NoteProperty -Name "Time Stamp" -Value $blankcol
        $diskOut | Add-Member -MemberType NoteProperty -Name "Usage (KBps)" -Value $blankcol

        $diskOut | Export-Csv "$outputpath\($server)diskUsage.csv" -NoTypeInformation -Append
    
    foreach($disktime in $disktimes){
        $actualdiskTime = $disktime.TimeStamp
        $actualdisk = $disktime.Value

        $Value = $diskChart.Series['Disk'].Points.AddXY("$actualdiskTime","$actualdisk")

        $diskOut = New-Object PSObject

        $diskOut | Add-Member -MemberType NoteProperty -Name "Server" -Value $Server
        $diskOut | Add-Member -MemberType NoteProperty -Name "Time Stamp" -Value $actualdiskTime
        $diskOut | Add-Member -MemberType NoteProperty -Name "Usage (KBps)" -Value $actualdisk

        $diskOut | Export-Csv "$outputpath\($server)diskUsage.csv" -NoTypeInformation -Append
    }

        $Title = New-Object -TypeName System.Windows.Forms.DataVisualization.Charting.Title
        $diskChart.Titles.Add($Title)
        $diskChart.Titles[0].Text = "Disk usage ($Server)"
 
        #Saving PNG file on desktop
        $diskChart.SaveImage("$outputpath\($Server)DiskUsage.png", "PNG")

 }