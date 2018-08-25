# Import AD Module
Import-Module activedirectory
 
#Set variables for domain, filename, date, save path
$Domain   = (Get-ADDomain | Select-Object -ExpandProperty Name)
$Date     = Get-Date -UFormat "%d%m%Y" 
$FullDate = (Get-Date -Format g)
$SavePath = "C:\Temp\ADobjectsExport\$Date"
If (-not(Test-Path $SavePath)){New-Item -Path $SavePath -ItemType Directory}

Get-ADComputer -Filter * -Properties DistinguishedName,IPv4Address,DNSHostname,OperatingSystem,OperatingSystemServicePack,OperatingSystemVersion,LastLogonDate,WhenChanged |
Select-Object DistinguishedName,IPv4Address,DNSHostname,OperatingSystem,OperatingSystemServicePack,OperatingSystemVersion,LastLogonDate,WhenChanged, Enabled |
Export-CSV "$SavePath\All_Computers_$domain-$Date.csv"

Get-ADUser -Filter * -Properties DistinguishedName,EmployeeID,EmployeeNumber,SamAccountName,Name,GivenName,Initials,Surname,DisplayName,Description,Title,EmailAddress,Department,Company,CannotChangePassword,PasswordNeverExpires,PasswordNotRequired,LockedOut,AccountExpirationDate,LastLogonDate,PasswordLastSet,whenCreated,lastLogonTimestamp,msDS-UserPasswordExpiryTimeComputed,whenChanged |
Select-Object DistinguishedName,EmployeeID,EmployeeNumber,SamAccountName,Name,GivenName,Initials,Surname,DisplayName,Description,Title,EmailAddress,Department,Company,CannotChangePassword,PasswordNeverExpires,PasswordNotRequired,LockedOut,AccountExpirationDate,LastLogonDate,PasswordLastSet,whenCreated,@{n='LastLogon';e={[DateTime]::FromFileTime($_.LastLogon)}},@{n='msDS-UserPasswordExpiryTimeComputed';e={[DateTime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed")}},whenChanged, Enabled |
Export-CSV "$SavePath\All_Users_$domain-$Date.csv"

Send-MailMessage -To KineticIT_Security@cpfs.wa.gov.au -From ProductionDomain@cpfs.wa.gov.au -Subject "Users/Computers Export from $domain on $FullDate" -SmtpServer "DC01SV043.ad.dcd.wa.gov.au" `
-Attachments "$SavePath\All_Users_$domain-$Date.csv","$SavePath\All_Computers_$domain-$Date.csv"

$Attachment1 = "\\ServerName\C$\temp\ADobjectsExport\$Date\"+"All_Users_DMZ-$Date.csv"
$Attachment2 = "\\ServerName\C$\temp\ADobjectsExport\$Date\"+"All_Computers_DMZ-$Date.csv"
 
Send-MailMessage -To 'Email' -From 'Email' -Subject "Users/Computers Export from DMZ on $FullDate" -SmtpServer "DC01SV043.ad.dcd.wa.gov.au" `
-Attachments $Attachment1,$Attachment2
 


 