$ErrorActionPreference = 'stop'
$Date = Get-Date
$LogPath = 'c:\Temp\Logs\GCN_WifiDevices.log'
$i=0;$j=0
if (!(Test-Path $LogPath)){New-Item -Path $LogPath -ItemType File -Force | Out-Null}
Add-Content -Value "################################## $Date ################################" -Path $LogPath

$GBL_802_Members = Get-ADGroupMember 'GBL_802.1x-authentication' # All member objects of the GBL_802.1x-authentication security group
$GCN_Devices = Get-ADComputer -SearchBase 'OU=GCN,OU=DCP,OU=Business,DC=ad,DC=dcd,DC=wa,DC=gov,DC=au' -SearchScope Subtree -Filter *

$DCP_Wireless_Members = Get-ADGroupMember 'GBL_DCP_Wireless_Devices' # All member objects of the GBL_DCP_Wireless_Devices security group
$All_Devices = ` # All devices contained within the below 4 OU's
'OU=Laptops and SurfacePro,OU=Pilot-Windows 10,OU=CPFS,OU=Business SOE V3,DC=ad,DC=dcd,DC=wa,DC=gov,DC=au', # 
'OU=Laptops,OU=CPFS,OU=Business SOE V3,DC=ad,DC=dcd,DC=wa,DC=gov,DC=au',# 
'OU=Laptops,OU=DCP,OU=Business,DC=ad,DC=dcd,DC=wa,DC=gov,DC=au',# 
'OU=SurfacePro,OU=GCN,OU=DCP,OU=Business,DC=ad,DC=dcd,DC=wa,DC=gov,DC=au'|
%{Get-ADComputer -SearchBase $_ -SearchScope Subtree -Filter *}

Add-Content -Value "Adding devices to Wireless_Devices" -Path $LogPath

foreach ($Device in $All_Devices)
{
    if ($DCP_Wireless_Members.Name -notcontains $Device.Name)
    {
        try
        {
            Add-ADGroupMember -Identity 'Wireless_Devices' -Members $Device.DistinguishedName
            Add-Content -Value "    * $($Device.Name) was added to Wireless_Devices" -Path $LogPath
            $i++
        }
        catch
        {
            Add-Content -Value ("    * Attempting $($Device.Name):" + $_) -Path $LogPath
        }
    }
}

Add-Content -Value "$i devices were added to Wireless_Devices" -Path $LogPath

Add-Content -Value "Adding devices to 802.1x-authentication" -Path $LogPath

foreach ($Device in $GCN_Devices)
{
    if ($GBL_802_Members.Name -notcontains $Device.Name)
    {
        try
        {
            Add-ADGroupMember -Identity '802.1x-authentication' -Members $Device.DistinguishedName
            Add-Content -Value "    * $($Device.Name) was added to 802.1x-authentication" -Path $LogPath
            $j++
        }
        catch
        {
            Add-Content -Value ("    * Attempting $($Device.Name):" + $_) -Path $LogPath
        }
    }  
}
Add-Content -Value "$j devices were added to 802.1x-authentication" -Path $LogPath
Add-Content -Value "###################################### FINISHED #######################################" -Path $LogPath