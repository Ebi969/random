$vmList = Get-content "D:\UserData\Ibraaheem\Scripts\VMWare\VM-Power\vmList.txt"

Get-vm $vmList | Start-VM