#$clusters = (get-content C:\tmp\ClusterList.txt)
$clusters = "SKY-BDC-GOLD-03-DATABASE", "SKY-BDC01-BRN01-D-DATABASE", "SKY-BDC01-GLD01-D-DATABASE", "SKY-CDC-GOLD-03-DATABASE", "SKY-CDC02-GLD01-D-DATABASE"

Foreach($cluster in $clusters){

    Get-Cluster $cluster | Get-VM | Select @{n="cluster"; e={$cluster}},@{n="vmhost"; e={$_.vmhost}},@{n="vm"; e={$_.name}} | Export-Excel -append -path "D:\UserData\Ibraaheem\Scripts\VMWare\vmsInCluster\SKY-DBClusters-VMS.xlsx" #"D:\UserData\muhammad\VMHostPatched.xlsx"
}