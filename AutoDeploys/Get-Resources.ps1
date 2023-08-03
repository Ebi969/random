param(
    [Parameter(Mandatory)]$jobID
)

$importsPaths = "\\srv005879\Reports\AutoDeploy"

$configList = Import-Excel "$importsPaths\ConfigManagement.xlsx"
$jobList = Import-Excel "$importsPaths\ExampleJobFile.xlsx"

# Get Cred store
$creds7281 = Get-VICredentialStoreItem -file  C:\Users\svcVMWareScriptAcc\CredStore7281.creds
$creds7282 = Get-VICredentialStoreItem -file  C:\Users\svcVMWareScriptAcc\CredStore7282.creds
$creds8097 = Get-VICredentialStoreItem -file  C:\Users\svcVMWareScriptAcc\CredStore8097.creds

# Connect-VIServer
Connect-VIServer -Server SRV007281.mud.internal.co.za -User $Creds7281.User -Password $Creds7281.Password
Connect-VIServer -Server SRV007282.mud.internal.co.za -User $Creds7282.User -Password $Creds7282.Password
Connect-VIServer -Server SRV008097.mud.internal.co.za -User $creds8097.User -Password $creds8097.Password

function Get-JobInfo {
param(
    [Parameter(Mandatory)]$jobID
)

    

}

function Get-ClusterLimits {
param(
    [Parameter(Mandatory)]$clusterName
)

    

}

function Get-ClusterStats {
param(
    [Parameter(Mandatory)]$clusterName,
    [Parameter(Mandatory)]$maxCPURatio,
    [Parameter(Mandatory)]$maxMemPercentage
)

    

}