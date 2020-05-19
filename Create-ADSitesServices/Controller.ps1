param($CsvPath, $DefaultSite, $Cost, $ReplicationInterval, $Log='C:\Temp\SiteCreation.log')

WriteLog "Processing $CsvPath on $ENV:COMPUTERNAME" $Log -Start
$ErrorActionPreference='stop'
Write-Host "Writing log to $Log"

#region Dependecy check

try
{
    Import-Module ActiveDirectory                                     # The Active Directory module must be available for the script to run
    Import-Module "$((Get-Location).Path)\ADSiteDHCPManagement.ps1"   # This needs to point to the custom DHCP management module <--- IMPORTANT!!
    if (!(Check-CsvHeaders $CsvPath))                                 # If the CSV headers are not as expected the script will not continue
    {throw 'Error: Invalid CSV headers' ; return}  
    $csv = Import-Csv $CsvPath ; $count = $csv.count                  # Import CSV and count how many sites there are
    WriteLog "All dependecies were met"
}
catch [System.IO.FileNotFoundException]
{
    Write-Host "Error: A dependency was not met`n$_" -ForegroundColor Red
    return
}

#endregion

#region Create sites, site links, and subnets

$i=1
foreach ($Site in $Csv)
{
    try
    {
        New-ADSite -SiteName $Site.Name -Description $Site.Name
        New-ADSiteLink -DefaultSite $DefaultSite -SiteName $Site.Name -Cost $Cost -ReplicationInterval $ReplicationInterval
        New-ADSubnet -SiteName $Site.Name -Subnet $Site.Subnet -Description $Site.Name -Location $Site.Name
        Write-Progress -Activity 'Creating AD sites, site links, and subnets' -PercentComplete ([math]::Round((($i/$count)*100),2)) -Status "$([math]::Round((($i/$count)*100),2))% Complete"
        $i++
    }
    catch
    {
        Write-Host "Error: Unable to create site component for $($site.Name)" -ForegroundColor Red
        WriteLog ($Error[0].Message) $Log      
        $i++ 
    }
}
WriteLog "Finished creating sites, site links and subnets"

#endregion

#region Create DHCP scopes and reservations

$i=1
foreach ($site in $csv) # Currently the DHCP server is hard-coded, I will modify this to use parameter input.
{
    try
    {
        $s=($site.ScopeId).Split('.')   
        New-DHCPScope -Server $PrimaryDHCPServer -Address $site.ScopeID -SubnetMask $site.SubnetMask -Name $site.Name -Description $site.Description
        Add-DHCPIPRange -Server $PrimaryDHCPServer -Scope $site.ScopeID -StartAddress $site.StartRange -EndAddress $site.EndRange
        Set-DHCPOption -OptionID 003 -Owner "$PrimaryDHCPServer/$($site.ScopeId)" -DataType IPAddress -Value $site.Router
        Set-DHCPOption -OptionID 006 -Owner "$PrimaryDHCPServer/$($site.ScopeId)" -DataType IPAddress -Value $site.DnsServer
        Set-DHCPOption -OptionID 015 -Owner "$PrimaryDHCPServer/$($site.ScopeId)" -DataType STRING -Value $site.DnsDomain
        New-DHCPReservation -Server $PrimaryDHCPServer -Scope $site.ScopeId -IPAddress "$($s[0]).$($s[1]).$($s[2]).127" -MACAddress 00112233445566 -Name GovNext_PlaceHolder1 -Description Reservation1
        New-DHCPReservation -Server $PrimaryDHCPServer -Scope $site.ScopeId -IPAddress "$($s[0]).$($s[1]).$($s[2]).128" -MACAddress 00112233445577 -Name GovNext_PlaceHolder2 -Description Reservation1
        New-DHCPReservation -Server $PrimaryDHCPServer -Scope $site.ScopeId -IPAddress "$($s[0]).$($s[1]).$($s[2]).129" -MACAddress 00112233445588 -Name GovNext_PlaceHolder3 -Description Reservation3
        Write-Progress -Activity 'Creating DHCP Scopes and reservations' -PercentComplete ([math]::Round((($i/$count)*100),2)) -Status "$([math]::Round((($i/$count)*100),2)) Complete"
        $i++
    }
    catch
    {
        Write-Host "Error: Unable to create DHCP scope for $($site.Name)" -ForegroundColor Red
        WriteLog ($Error[0].Message) $Log 
        $i++    
    }
}
WriteLog "Finished creating DHCP scopes"

#endregion