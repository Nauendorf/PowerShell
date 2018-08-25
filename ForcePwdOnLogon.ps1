
$OUs =  "OU=Users,", # Contractors OU
        "OU=Users,",         # Standard Users OU
        "OU=Users,"  # SOE V3 OU

# This query will only return user objects that are normal accounts where the password expires and is required, enabled, with passwords older 30 hours, contained within one of the above OU's
$TargetUsers = $OUs|%{Get-ADUser -Properties samAccountName,pwdLastSet,PasswordNotRequired,PasswordNeverExpires,PasswordLastSet -SearchBase $_ -SearchScope Subtree -LDAPFilter "(&(objectClass=User)(userAccountControl=544))"}|
?{($_.PasswordLastSet -gt (Get-Date).AddHours(-5))}

#|?{($_.PasswordLastSet -lt (Get-Date).AddHours(-30))}    Peter Harrison has requested for all accounts to be reset regardless of password age.

$TargetCount = $TargetUsers.Count # Count the number of accounts that will be modified

# Export modified accounts to csv, this can be used to reverse changes
$TargetUsers|Export-Csv -NoClobber -NoTypeInformation -Path 'C:\Temp\ModifiedUsers.csv' 

# Gives a final warning before modifying any objects
$message = "Are you sure you want to force ChangePasswordAtLogon for $TargetCount user accounts?"
$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes"
$no = New-Object System.Management.Automation.Host.ChoiceDescription "&No"
$options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
$result = $host.ui.PromptForChoice('', $message, $options, 0)

If ($result -eq '0') # If you selected yes then the captured objects will be modified. 0 represents the 'Yes' options since it is first in the options array.
{
    Try
    {
        $TargetUsers|%{Set-ADUser $_ -ChangePasswordAtLogon $true -ErrorAction stop}
    }
    Catch
    {
        $Error[0]
        return # If any errors occur the loop will immediately stop... there shouldn't be any errors...
    }
}


 Send-MailMessage -To 'Email' -From 'Email' -Subject "Success!!" -SmtpServer "DC01SV043.ad.dcd.wa.gov.au" -Attachments 'C:\Temp\ModifiedUsers.csv'

<#  ## Backout with this ##
    $TargetUsers = Import-Csv -Path 'C:\Temp\ModifiedUsers.csv'
    $TargetUsers|%{Set-ADUser $_ -ChangePasswordAtLogon $false -ErrorAction stop}
#>