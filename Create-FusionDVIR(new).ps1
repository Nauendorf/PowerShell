<#
New Fusion/DVIR account creation
Version 3.0
Author: David Nauendorf
requires -version 3
#>

################################ Start Functions ##################################

# Check for existing email address
Function Get-Existing {

Param
(
    [parameter(Mandatory=$true)]$Email
)

$Searcher = [adsisearcher]"(&(objectCategory=person)(objectClass=User)(mail=$Email))"
$Searcher_Contact = [adsisearcher]"(&(objectCategory=person)(objectClass=contact)(mail=$Email))"

    If ( ( ($Searcher.FindAll().Path) -eq $null) -and ( ($Searcher_Contact.FindAll().Path) -eq $null) )
    {
        return $false
    }
    Else
    {
        $Script:Path =  ($Searcher.FindAll().Path)
        return $true
    }
}

# Scrolls to the bottom of the output box
Function ScrollDown
{
$OutputTxt.SelectionStart = $OutputTxt.TextLength;
$OutputTxt.ScrollToCaret()
$OutputTxt.Focus()
}

# Append text and errors to a log file
Function AppendLog
{
Param
    (
    [Parameter(Mandatory=$true)]
    [string]$LogText,
    [Parameter(Mandatory=$false)]
    [switch]$Errors,
    [Parameter(Mandatory=$false)]
    [switch]$Time,
    [Parameter(Mandatory=$false)]
    [switch]$Header,
    [Parameter(Mandatory=$false)]
    [switch]$Footer
    )

$LogArray  = @()
$HeadStamp = @()
$FootStamp = @()
$ErrorMSG = $Error[0]
$DateStamp = Get-Date -Format G
$TimeStamp = Get-Date -Format T
$LogFile    = "C:\Temp\Logs\FusionDVIR_Creation.log"

$HeadStamp = @"

_______________________________________________________________________
Initiated by $ENV:USERNAME on $DateStamp
"@

$FootStamp = @"

$TimeStamp
________________________________________________________________________

"@
If (!($Header)){$HeadStamp = $null}
If (!($Footer)){$FootStamp = $null}
If (!($Errors)){$ErrorMSG = "None"}
If (!($Time)){$TimeStamp = $null}
If (!( Test-Path $LogFile)){New-Item -ItemType File -Path $LogFile -Force}

$LogArray=@"
$HeadStamp
$TimeStamp
$LogText
Errors:
$ErrorMSG
$FootStamp
"@

Add-Content -Value $LogArray -Path $LogFile -Force

}

# Generate a unique username
Function Generate-Username
{
    if ($FirstName.Length -gt 6)
    {
        $FirstNameShort = $FirstName.Substring(0,6)
    }
    else
    {
        $FirstNameShort = $FirstName.ToUpper()
    }
    
    if ($LastName.Length -gt 2)
    {
        $LastNameShort = $LastName.Substring(0,2)
    }
    else
    {
        $LastNameShort = $LastName.ToUpper()
    }
    
    $searcher = [adsisearcher]"(samaccountname=$FirstNameShort$LastNameShort)"
    $rtn = $searcher.findall()
    
    while ($rtn.count -ne 0)
    {
        if ($LastName.Length -gt 2)
        {
            [int]$Rand = Get-Random -Maximum ($LastName.Length -2)
            $LastNameShort = $LastName.Substring($Rand,2)
        }
        $searcher = [adsisearcher]"(samaccountname=$FirstNameShort$LastNameShort)"
        $rtn = $searcher.findall()
    }
    
    return "$FirstNameShort$LastNameShort"
}

# Generate a password
Function Generate-Password
{
    [int32[]]$ArrayofAscii=26,97,26,65,10,48,15,33
    
    $Upper = [char]((Get-Random 26) + 65)
    For($i=1; $i -le 6; $i++){$Lower = $Lower + [char]((Get-Random 26) + 97)}
    $number = [char]((Get-Random 10) +48)
    Return $Upper+$Lower+$number
}

# Creates the account. This function is called from the event handler for the Create button
Function CreateAccount
{
Param (
        [Parameter(Position=0)]
        [string]$AccountType,
        [Parameter(Position=1)]
        [string]$Firstname,
        [Parameter(Position=2)]
        [string]$Lastname,
        [Parameter(Position=3)]
        [string]$EmailAddress,
        [Parameter(Position=3)]
        [string]$RequestNumber
       )
Try{Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010} ## Adds exchange cmdlets to the session 
Catch{
        $OutputTxt.AppendText("Error: Unabel to load Exchange PSSnapin...")
        ScrollDown
        AppendLog -LogText "Error: Unabel to load Exchange PSSnapin..." -Errors -Time
      }

### Account creation logic

AppendLog -LogText "Creating AD object" -Time
$OutputTxt.AppendText("`n Creating the AD object")
ScrollDown
$ErrorActionPreference = 'silentlycontinue'
$Date = Get-Date -Format d
$NewUserName = $(Generate-Username)
$NewUserName = $NewUserName.ToUpper()
# $NewUserPassword = $(Generate-Password)
#$Passwd = convertto-securestring $NewUserPassword -asplaintext -force
$Displayname = "$Firstname $Lastname"
$Description += "Created $Date $RequestNumber by $ENV:USERNAME;"
    Try
    { # Create AD object
        If ($AccountType -eq "Fusion") # Set OU depending on account type
        { 
        $OU = "OU="
        $Memberof = ""
        $NewUserPassword = ""
        $SecurePasswd = ConvertTo-SecureString $NewUserPassword -AsPlainText -Force
        }
        ElseIf ($AccountType -eq "DVIR")
        { 
        $OU = "OU=" 
        $Memberof = ""
        $NewUserPassword = ""
        $SecurePasswd = ConvertTo-SecureString $NewUserPassword -AsPlainText -Force
        }

        New-ADUser -Name "$Firstname $Lastname" `
                   -GivenName $Firstname `
                   -Surname $Lastname `
                   -DisplayName "$Firstname $Lastname" `
                   -UserPrincipalName "$NewUserName@dcd.wa.gov.au" `
                   -samAccountName $NewUserName `
                   -Path $OU `
                   -ChangePasswordAtLogon $true
        Start-Sleep 3
        
        Set-ADAccountPassword -Identity $NewUserName `
                              -Reset `
                              -NewPassword $SecurePasswd

        Set-aduser $NewUserName -changepasswordatlogon $true -Description $Description

        $NewUserName | Enable-ADAccount
                              
        $OutputTxt.AppendText("`n Waiting for account")                      
        $Searcher = [adsiSearcher]"(samAccountName=$NewUserName)"
        $Result = $Searcher.FindAll()
        If ($Result.Count -lt 1){Start-Sleep 2}

        Add-ADGroupMember -Identity $Memberof -Members $NewUserName
        $OutputTxt.AppendText("`n Finished configuring AD account")
        ScrollDown
        AppendLog -LogText "Added user to $Memberof" -Time
    }
    Catch
    {
        AppendLog -LogText $Error[0] -Errors -Time
        $OutputTxt.AppendText("`n [Error] Something went wrong configuring the AD account. See logs for details")
        ScrollDown
        $ErrorCount = 1
    }

    # Create mailbox for new user
    Try
    {
        AppendLog -LogText "Creating Exchange Session" -Time
        $OutputTxt.AppendText("`n Creating the mailbox")
        ScrollDown
        $EXURI = "http://EschangeServerFQDN/Powershell"
        $PrimarySMTP = "$Firstname.$Lastname@dcd.wa.gov.au"
        $ExchangeSession = New-PSSession -ConfigurationName microsoft.exchange -ConnectionURI $EXURI 
        Import-PSSession $ExchangeSession -DisableNameChecking | Out-Null
        AppendLog "Imported"
        $SmallestDB = Get-MailboxDatabase -Status | 
                      sort DatabaseSize | 
                      Select-Object Name,DatabaseSize,AvailableNewMailboxSpace | 
                      where {$_.Name -notlike "Disabled users Database"} | 
                      where {$_.Name -notlike "RDB1"} | 
                      Select-Object -First 1 | 
                      Select-Object Name | 
                      ForEach {$_.Name}
        AppendLog -LogText "Found smallest DB is $SmallestDB" -Time
        Enable-Mailbox -Identity "$NewUserName" -Database $SmallestDB -Alias "$NewUserName" | Out-Null
        AppendLog -LogText "Enabled Mailbox" -Time
        $OutputTxt.AppendText("`n Mailbox successfully created")
       
    }
    Catch
    {
        AppendLog -LogText $Error[0] -Errors -Time
        $OutputTxt.AppendText("`n [Error] Something went wrong configuring the users mailbox. See logs for details")
        ScrollDown
        $ErrorCount = 1
    }
    Finally
    {
        Get-PSSession | Remove-PSSession
    }

    Try
    {
        $OU = "OU=" # OU for external contacts, used to create contact and set description
        AppendLog -LogText "Creating mail contact" -Time
        $OutputTxt.AppendText("`n Creating mail contact")
        ScrollDown
        New-MailContact -FirstName $FirstName `
                        -LastName $LastName `
                        -Name $Displayname `
                        -ExternalEmailAddress $EmailAddress `
                        -OrganizationalUnit $OU `
                        -Alias "$FirstName$LastName"
        $OutputTxt.AppendText("`n Mail contact successfully created")
        ScrollDown
        Start-Sleep 5

        $OutputTxt.AppendText("`n Configuring the mail contact")
        ScrollDown
        Set-MailContact "$Firstname$Lastname" -HiddenFromAddressListsEnabled $true # Hide contact from GAL
        # Set description for AD contact object (not the exchange contact)      

        $ADobject = Get-ADObject -LDAPFilter "ObjectCategory=Contact" -SearchBase $OU | Where-Object {$_.DistinguishedName -like "*$FirstName $LastName*"} # Variable containing AD contact object
        Set-ADObject -Identity $ADobject -Description $Description # Set description for AD contact object

        AppendLog -LogText "Set mailbox hidden, forward, primary SMTP" -Time
        $OutputTxt.AppendText("`n The mail contact has been configured")
        ScrollDown

        Start-Sleep 3
        $OutputTxt.AppendText("`n Configuring mailbox")
        Set-Mailbox -Identity $NewUserName -ForwardingAddress $EmailAddress # Set email forward
        $OutputTxt.AppendText("`n Setting mail forward")
        Set-Mailbox -Identity $NewUserName -EmailAddressPolicyEnabled $false -PrimarySMTPAddress $PrimarySMTP # set primary SMPT to DCD
        $OutputTxt.AppendText("`n Setting primary SMTP")
        Set-Mailbox -Identity $NewUserName -HiddenFromAddressListsEnabled $true # Hide mailbox from GAL
        $OutputTxt.AppendText("`n Hiding from GAL")

        Write-Host "Finished configuring mailbox"
        AppendLog -LogText "Finished configuring mail contact" -Time
        $OutputTxt.AppendText("`n Finished configuring mailbox")
        ScrollDown
        Get-PSSession | Remove-PSSession
    }
    Catch
    {
        $OutputTxt.AppendText("`n There was an error configuring the mail contact")
        ScrollDown
        AppendLog -LogText $Error[0] -Time -Errors
        $ErrorCount = 1
    }

If ($ErrorCount -gt 0)
{ $OutputTxt.AppendText("`n Errors occurred during the account creation. See logs for details.") }
Else
{ $OutputTxt.AppendText("`n SUCCESS! No errors occurred.") }

# Final account details for the new user are displayed on the output box and sent to the clipboard
$Finished_MSG = @"


The account for $Firstname $Lastname has been created with the following login details.

        Username: $NewUserName
        Password: $NewUserPassword

"@
Write-Host "Account creaion complete" -ForegroundColor Green
AppendLog -LogText $Finished_MSG -Time
AppendLog -LogText "Account creaion complete" -Footer
$OutputTxt.AppendText("`n $Finished_MSG")
ScrollDown
$Finished_MSG | clip.exe
### End CreateAccount ###
}

################################ End Functions #####################################


################################### Start Building GUI #####################################

# Load assemblies for GUI objects
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") 
[System.Windows.Forms.Application]::EnableVisualStyles()

# Create form objects 
$MainForm      = New-Object System.Windows.Forms.Form
$DetailGroup   = New-Object System.Windows.Forms.GroupBox
$accTypeGroup  = New-Object System.Windows.Forms.GroupBox
$ConsoleGroup  = New-Object System.Windows.Forms.GroupBox
$Fusion_Radio  = New-Object System.Windows.Forms.RadioButton
$DVIR_Radio    = New-Object System.Windows.Forms.RadioButton
$CreateButton  = New-Object System.Windows.Forms.Button
$CloseButton   = New-Object System.Windows.Forms.Button
$CopyButton    = New-Object System.Windows.Forms.Button
$emailTXT      = New-Object System.Windows.Forms.TextBox
$FirstnameTxt  = New-Object System.Windows.Forms.TextBox
$LastnameTxt   = New-Object System.Windows.Forms.TextBox
$RequestTxt    = New-Object System.Windows.Forms.TextBox
$OutputTxt     = New-Object System.Windows.Forms.RichTextBox
$RequestLabel  = New-Object System.Windows.Forms.Label
$FnameLabel    = New-Object System.Windows.Forms.Label
$LnameLabel    = New-Object System.Windows.Forms.Label
$EmailLabel    = New-Object System.Windows.Forms.Label
$Font          = New-Object System.Drawing.Font("Terminal",9)

# Properties of form objects
$MainForm.Size            = New-Object System.Drawing.Size(396,430)
$MainForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$MainForm.StartPosition   = [System.Windows.Forms.FormStartPosition]::CenterScreen
$MainForm.Text            = "DVIR / Fusion Account Creator"
#$MainForm.BackColor       = "Azure"
$MainForm.TopMost         = $true
$MainForm.MaximizeBox     = $false
$MainForm.ShowIcon        = $true

# Group properties
$DetailGroup.Text     = "Account Details"
$DetailGroup.Size     = New-Object System.Drawing.Size(272,195)
$DetailGroup.Location = New-Object System.Drawing.Size(13,7)

    # Request textbox properties
    $RequestTxt.Size       = New-Object System.Drawing.Size(100,20)
    $RequestTxt.Location   = New-Object System.Drawing.Size(80,20)

    $FirstnameTxt.Size     = New-Object System.Drawing.Size(180,20)
    $FirstnameTxt.Location = New-Object System.Drawing.Size(80,50)

    $LastnameTxt.Size      = New-Object System.Drawing.Size(180,20)
    $LastnameTxt.Location  = New-Object System.Drawing.Size(80,80)

    $emailTXT.Size         = New-Object System.Drawing.Size(180,20)
    $emailTXT.Location     = New-Object System.Drawing.Size(80,110)

    # Label properties
    $RequestLabel.Text     = "Request #"
    $RequestLabel.Size     = New-Object System.Drawing.Size(80,23)
    $RequestLabel.Location = New-Object System.Drawing.Size(10,22)
    $RequestLabel.Font = $Font

    $FnameLabel.Text       = "First Name"
    $FnameLabel.Size       = New-Object System.Drawing.Size(80,23)
    $FnameLabel.Location   = New-Object System.Drawing.Size(10,52)
    $FnameLabel.Font = $Font

    $LnameLabel.Text       = "Last Name"
    $LnameLabel.Size       = New-Object System.Drawing.Size(80,23)
    $LnameLabel.Location   = New-Object System.Drawing.Size(10,82)
    $LnameLabel.Font = $Font

    $EmailLabel.Text       = "Email"
    $EmailLabel.Size       = New-Object System.Drawing.Size(80,23)
    $EmailLabel.Location   = New-Object System.Drawing.Size(11,112)
    $EmailLabel.Font = $Font

    # Field group controls
    $DetailGroup.Controls.Add($RequestTxt)
    $DetailGroup.Controls.Add($FirstnameTxt)
    $DetailGroup.Controls.Add($LastnameTxt)
    $DetailGroup.Controls.Add($emailTXT)
    $DetailGroup.Controls.Add($accTypeGroup)
    $DetailGroup.Controls.Add($RequestLabel)
    $DetailGroup.Controls.Add($FnameLabel)
    $DetailGroup.Controls.Add($LnameLabel)
    $DetailGroup.Controls.Add($EmailLabel)

# Account type group properties
$accTypeGroup.Text     = "Account Type"
$accTypeGroup.Size     = New-Object System.Drawing.Size(125,50)
$accTypeGroup.Location = New-Object System.Drawing.Size(80,135)
    
    $Fusion_Radio.Text     = "Fusion"
    $Fusion_Radio.Location = New-Object System.Drawing.Size(10,20)
    $Fusion_Radio.Size     = New-Object System.Drawing.Size(58,20)
    $Fusion_Radio.Checked  = $true # Fusion radio checked by default

    $DVIR_Radio.Text     = "DVIR"
    $DVIR_Radio.Location = New-Object System.Drawing.Size(70,20)
    $DVIR_Radio.Size     = New-Object System.Drawing.Size(50,20)

    # Account type group controls
    $accTypeGroup.Controls.Add($Fusion_Radio)
    $accTypeGroup.Controls.Add($DVIR_Radio)

# Console group properties
$ConsoleGroup.Text     = "Console Output"
$ConsoleGroup.Size     = New-Object System.Drawing.Size(362,182)
$ConsoleGroup.Location = New-Object System.Drawing.Size(13,205)

    $OutputTxt.Size        = New-Object System.Drawing.Size(342,155)
    $OutputTxt.Location    = New-Object System.Drawing.Size(10,17)
    $OutputTxt.BorderStyle = 1
    $OutputTxt.ReadOnly    = $true
    $OutputTxt.WordWrap    = $true
    $OutputTxt.BackColor   = "White"
    $OutputTxt.Cursor      = "Hand"
    $OutputTxt.ScrollBars  = "ForcedVertical" 
    $OutputTxt.Font = New-Object System.Drawing.Font("Courier",9)

    # Console group controls
    $ConsoleGroup.Controls.Add($OutputTxt)

# Button properties
$CreateButton.Text      = "Create"
$CreateButton.Size      = New-Object System.Drawing.Size(80,25)
$CreateButton.Location  = New-Object System.Drawing.Size(296,14)
$CreateButton.FlatStyle = "popup"
$CreateButton.Font = $Font

$CloseButton.Text      = "Close"
$CloseButton.Size      = New-Object System.Drawing.Size(80,25)
$CloseButton.Location  = New-Object System.Drawing.Size(296,45)
$CloseButton.FlatStyle = "popup"
$CloseButton.Font = $Font

$CopyButton.Text      = "Copy"
$CopyButton.Size      = New-Object System.Drawing.Size(35,18)
$CopyButton.Location  = New-Object System.Drawing.Size(285,1)
$CopyButton.FlatStyle = "flat"
$CopyButton.Font = "Terminal,6"
#$OutputTxt.Controls.Add($CopyButton)

# Mainform controls
$MainForm.Controls.Add($DetailGroup)
$MainForm.Controls.Add($CreateButton)
$MainForm.Controls.Add($CloseButton)
$MainForm.Controls.Add($ConsoleGroup)

################################# End GUI ########################################


################################ Event handlers & click events ################################

$Create_Event = [System.EventHandler]{ # This event sets the required variables and calls the CreateAccount function

$OutputTxt.Clear()

If (-not($RequestTxt.Text)){ # Checks for mandatory inputs
    Write-Warning "You must enter a request number"
    [Microsoft.VisualBasic.Interaction]::MsgBox("You must enter a request number", "OKOnly,SystemModal,Exclamation", "Warning")
    }
    Elseif (-not($FirstnameTxt.Text)){
        Write-Warning "You must enter a first name"
        [Microsoft.VisualBasic.Interaction]::MsgBox("You must enter a first name", "OKOnly,SystemModal,Exclamation", "Warning")
        }
    Elseif (-not($LastnameTxt.Text)){
        Write-Warning "You must enter a last name"
        [Microsoft.VisualBasic.Interaction]::MsgBox("You must enter a last name", "OKOnly,SystemModal,Exclamation", "Warning")
        }        
    Elseif (-not($emailTXT.Text)){
        Write-Warning "You must enter an email address"
        [Microsoft.VisualBasic.Interaction]::MsgBox("You must enter an email address", "OKOnly,SystemModal,Exclamation", "Warning")
        }
    Elseif (($emailTXT.Text -notlike "*@*.*" )){
        Write-Warning "Invalid email address"
        [Microsoft.VisualBasic.Interaction]::MsgBox("Invalid email address", "OKOnly,SystemModal,Exclamation", "Warning")
        }
    Elseif ((Get-Existing -Email $emailTXT.Text) -eq $true )
        {
            Write-Warning 'The email address already exists as a user or mail contact'
            $mail = $emailTXT.Text
            [Microsoft.VisualBasic.Interaction]::MsgBox("An object with the email $mail already exists.", "OKOnly,SystemModal,Critical", "Error")
        }
    Else
        {
            If ($Fusion_Radio.Checked -eq $true){
                $AccountType = "Fusion"} Else {
                $AccountType = "DVIR"}
             # Sanitize user input
            $RequestTxt.Text.Trim()
            $RequestTxt.Text -replace '\s+'
            [string]$RequestNum = $RequestTxt.Text
            $RequestNum = $RequestNum.ToUpper()

            $FirstnameTxt.Text.Trim()
            $FirstnameTxt.Text -replace '\s+'
            [string]$Firstname = $FirstnameTxt.Text
            $Firstname = $Firstname.substring(0,1).toupper()+$Firstname.substring(1)

            $LastnameTxt.Text.Trim()
            $LastnameTxt.Text -replace '\s+'
            [string]$Lastname = $LastnameTxt.Text
            $Lastname = $Lastname.substring(0,1).toupper()+$Lastname.substring(1)

            $emailTXT.Text.Trim()
            $emailTXT.Text -replace '\s+'
            $EmailAddress = $emailTXT.Text

$CreatingAccountMSG = @"

Creating a $AccountType account for $Firstname $Lastname at $EmailAddress for Request $RequestNum

"@ 
            Write-Host $CreatingAccountMSG 
            $OutputTxt.AppendText($CreatingAccountMSG)
            $OutputTxt.AppendText("`n Please wait...")
            AppendLog -LogText $CreatingAccountMSG -Header
            CreateAccount -AccountType $AccountType -Firstname $Firstname -Lastname $Lastname -EmailAddress $EmailAddress -RequestNumber $RequestNum
        }

} # End of Create_Event

$CopyOutput_Event = [System.EventHandler]{ # Copies to console output to the clipboard
$OutputTxt.Text | clip.exe
} # End of CopyOutput_Event

$CreateButton.Add_Click($Create_Event)
$CopyButton.Add_Click($CopyOutput_Event)
$MainForm.CancelButton = $CloseButton

############################ End event handlers & click events #########################

$MainForm.ShowDialog()
