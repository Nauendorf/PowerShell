<# 
    User Profile Reset Tool
    Version 2.0
    Author: David Nauendorf
    requires -version 3
#>

If (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{   
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    [System.Windows.Forms.MessageBox]::Show("This is not an elevated session. You may not have the appropriate permissions to manage user profiles.","Are you sure?","ok","Question")
}

If ($Host.Version.Major -lt 3)
{
    If ([System.Windows.Forms.MessageBox]::Show("You are running PowerShell version $($Host.Version.Major). This version is not compatible.`n`nWould you like to download the latest Windows Management Framework?","Incompatible PowerShell Version","yesno","Warning")-eq'yes')
    {
        #$IE=new-object -com internetexplorer.application
        #$IE.navigate2("https://www.microsoft.com/en-us/download/details.aspx?id=50395")
        #$IE.visible=$true
        Start-Process "chrome.exe" "https://www.microsoft.com/en-us/download/details.aspx?id=50395"
    }
    return
}

#region 'Functions and dependencies
$ErrorActionPreference = 'stop'
# C# code for loading Shell32.dll icons

$code = @"
using System;
using System.Drawing;
using System.Runtime.InteropServices;

namespace System
{
	public class IconExtractor
	{

	 public static Icon Extract(string file, int number, bool largeIcon)
	 {
	  IntPtr large;
	  IntPtr small;
	  ExtractIconEx(file, number, out large, out small, 1);
	  try
	  {
	   return Icon.FromHandle(largeIcon ? large : small);
	  }
	  catch
	  {
	   return null;
	  }

	 }
	 [DllImport("Shell32.dll", EntryPoint = "ExtractIconExW", CharSet = CharSet.Unicode, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
	 private static extern int ExtractIconEx(string sFile, int iIndex, out IntPtr piLargeVersion, out IntPtr piSmallVersion, int amountIcons);

	}
}
"@ 

# Load assemblies for GUI objects
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[System.Windows.Forms.Application]::EnableVisualStyles()
Add-Type -TypeDefinition $code -ReferencedAssemblies System.Drawing # Loads C# assembly for icons

Function AppendLog
{
[cmdletbinding(DefaultParameterSetName=1)]
Param
(

[parameter(ParameterSetName=1)][string]$LogText,
[parameter(ParameterSetName=2)][switch]$Header,
[parameter(ParameterSetName=3)][switch]$Footer

)

    $ErrorActionPreference = 'silentlycontinue'
    $Error_Message = New-Object System.Collections.ArrayList
    $LogFile = "C:\Temp\Logtest.log"
    $csvLogFile = 'C:\Temp\Teslogfile.csv'
    If (!(Test-Path -Path $csvLogFile))
    {''|Select-Object Time,Action,Errors | Export-Csv -Path $csvLogFile -NoClobber -NoTypeInformation
    $Csvlog = Import-Csv -LiteralPath $csvLogFile}
    If (!(Test-Path -Path $LogFile))
    {New-Item -ItemType File -Path $LogFile -Force|Out-Null}
    $DateStamp = Get-Date -Format G
    $TimeStamp = Get-Date -Format T

        ################ HEADER ###############
        $HeadStamp = @"
*****************************************************************************
     ============= Started at $DateStamp by $ENV:Username ============
"@

        ############## FOOTER ###############
        $FootStamp = @"
     ================== Completed at $DateStamp ================
*****************************************************************************
"@

    If ($Header)
    {
        Add-Content -Value $HeadStamp -Path $LogFile -Force
    }
    Elseif ( ($Footer))# -or (Test-Path "C:\Temp\Logtest.log") )
    {
        Add-Content -Value $FootStamp -Path $LogFile -Force
    }
    Else
    {
        If (!($Error[0]))
        {
            $Error_Message.Add('None') | Out-Null
        }
        Else
        {
            If ($Error[0].GetType().Name -eq 'ErrorRecord')
            {
                $Error_Message.Add($Error[0].Exception.Message) | Out-Null
            }
        }

        $LogEntry = "$TimeStamp - $LogText - Errors: $Error_Message"
        Add-Content -Value $LogEntry -Path $LogFile -Force

        #$LogEntrycsv = New-Object PsCustomObject -Property @{ Time = $TimeStamp ; Action = $LogText ; Errors = $Error_Message }
        #Export-Csv -InputObject $LogEntrycsv -Path $csvLogFile -Append

    }
    $Error.Clear()
    $Error_Message.Clear()
}

Function Backup-UserProfile
{
Param
(
[string][Parameter( ParameterSetName='Restore', Position=0)][Parameter(ParameterSetName='Backup',Position=0, Mandatory=$True)]$Username,
[string][Parameter( ParameterSetName='Backup', Position=1)][Parameter( ParameterSetName='Restore', Position=1, Mandatory=$True)]$Computername,
[switch][Parameter( ParameterSetName='Backup', Position=2, Mandatory=$True)]$Backup,
[string][Parameter( ParameterSetName='Backup', Position=3, Mandatory=$True)]$BackupFrom,
[string][Parameter( ParameterSetName='Backup', Position=4, Mandatory=$True)]$BackupTo,
[string][Parameter( ParameterSetName='Restore', Position=3, Mandatory=$True)]$RestoreFrom,
[string][Parameter( ParameterSetName='Restore', Position=4, Mandatory=$True)]$RestoreTo,
[switch][Parameter( ParameterSetName='Restore', Position=2, Mandatory=$True)]$Restore
)
    $BackupTargets = `
    "Desktop",
    "Downloads",
    "Favorites",
    "Documents",
    "Music",
    "Pictures",
    "Videos",
    "Contacts",
    "AppData\Local\Mozilla",
    "AppData\Local\Google\Chrome\User Data",
    "AppData\Roaming\Mozilla",
    "AppData\Roaming\Microsoft\Internet Explorer\Quick Launch",
    "AppData\Roaming\Microsoft\Proof",
    "AppData\Roaming\Microsoft\Signatures",
    "AppData\Roaming\Microsoft\Sticky Notes"

    Try
    {
        #$UserProfile = (Get-WmiObject win32_userprofile -Property * -ComputerName $Computername).Where({$_.LocalPath -like "*\Users\$username*"})
        [array]$BackupSource=@()
        Foreach ($Target in $BackupTargets)
        {
            If ([IO.Directory]::Exists($BackupFrom + "\$Target\") -and ( (ls ($BackupFrom + "\$Target\") ).Count -gt '1') ) # Removed -recurse
            {[array]$BackupSource+=$BackupFrom + "\$Target\"}
        }
    }
    Catch
    {
        AppendLog -LogText "Testing which profile targets exist"       
    }

    If ($Backup)
    {   
        [int]$i=1
        $BackupFrom_Count = (Get-ChildItem $BackupSource -Recurse -File).Count
        $BackupFrom_List  = (Get-ChildItem $BackupSource -Recurse -File)
           
        Foreach ($File in $BackupFrom_List)
        {
            Try
            {
                $Destination = $BackupTo + ($File.FullName).Split(':')[1]
                $DirectoryName = $BackupTo + ($File.DirectoryName).Split(':')[1]
                If (![IO.Directory]::Exists($DirectoryName) ){[IO.Directory]::CreateDirectory($DirectoryName)}
                Copy-Item -Path $File.FullName -Destination $Destination -Force -ErrorAction Continue -Verbose
            }
            Catch
            {                    
                AppendLog -LogText "Trying to backup items from $($File.FullName) to $Destination"
            }      
            Finally
            {
                [float]$pct = ($i/$BackupFrom_Count)*100
                $progressbar.Value = ($pct)
                [void] [System.Windows.Forms.Application]::DoEvents()
                [int]$i++|Out-Null
            }                                                                    
        }
     
        $progressbar.Value = 100
        AppendLog -LogText "Completing backup"  
    }
    Elseif ($Restore)
    {  
        $progressbar.Value = 0       
        [int]$i=1
        $RestoreFrom_FileCount = (Get-ChildItem $RestoreFrom -Recurse -File).Count
        $RestoreFrom_FileList  = (Get-ChildItem $RestoreFrom -Recurse -File)
                  
        Foreach ($File in $RestoreFrom_FileList)
        {
            Try
            {
                $DirectoryName = $RestoreTo + ($File.Directory -split "(\S$Username)")[-1]
                $Destination = $RestoreTo + ($File.FullName -split "(\S$Username)")[-1]
                If (![IO.Directory]::Exists($DirectoryName) ){[IO.Directory]::CreateDirectory($DirectoryName)}                       
                Copy-Item -Path $File.FullName -Destination $Destination -Force -ErrorAction Continue -Verbose
            }
            Catch
            {
                AppendLog -LogText "Trying to restore items from $($File.FullName) to $RestoreTo"
                [System.Windows.Forms.MessageBox]::Show("$($Error[0].Exception.Message)","Copy error ¯\_(ツ)_/¯","OK", "Error")
            }
            Finally
            {
                [float]$pct = ($i/$RestoreFrom_FileCount)*100
                $progressbar.Value = ($pct)
                [void] [System.Windows.Forms.Application]::DoEvents()
                [int]$i++|Out-Null
            }
        }

        $progressbar.Value = 100    
    }
}

Function Remove-UserProfile 
{
[cmdletbinding()]
Param
(
    [string]$Computername,
    [string]$ProfilePath
)
    # Test connection to the host again before attempting to reset the selected profile
    If (Test-Connection $ComputerName -Count 2 -Quiet)
    {
        # Set variables...
        $Date = get-date -UFormat "%d%m%Y%S"
        $SelectedProfile = (Get-WmiObject -Class Win32_UserProfile -ComputerName $ComputerName).Where({$_.LocalPath -eq $ProfilePath})

        # Restart the computer if the user profile is in use
        If ( ($SelectedProfile.RefCount -gt 0) -and ($SelectedProfile.Loaded -eq $true) )
        { 
            AppendLog -LogText "$SelectedProfile is currently loaded. Restart required."
            $LockedProfErr = [System.Windows.Forms.MessageBox]::Show("The profile for '$($SelectedProfile.LocalPath)\' on $Computername is currently loaded. The computer must be restarted before the profile can be removed.`n`nWould you like to restart now?","Restart now?", "YesNo", "Information")
            If ($LockedProfErr -eq 'yes')
            {
                Try
                {
                    AppendLog -LogText "Restarting $Computername"
                    Restart-Computer -ComputerName $ComputerName -Wait -For WMI -Timeout 300 -Protocol dcom -Force -Verbose -Delay 30 -ErrorAction Stop
                    AppendLog -LogText "Successfully reconnected to $Computername"
                    [System.Windows.Forms.MessageBox]::Show("$Computername is back online. Try removing the profile now.","Restart complete","OK","Information")
                    return
                }
                Catch
                {
                    [System.Windows.Forms.MessageBox]::Show("Failed to find $Computername after initiating restart","Something went wrong...¯\_(ツ)_/¯","OK","Information")
                    return
                }
            }
            Elseif ($LockedProfErr -eq 'no')
            {return}
        }
        Else
        {
            Try
            {   ### DELETES the user profile
                $SelectedProfile   = (Get-WmiObject -Class Win32_UserProfile -ComputerName $ComputerName) | Where-Object -Property LocalPath -Like $ProfilePath -ErrorAction Stop 
                $OldProfilePath    = "\\$Computername\C$" + ($SelectedProfile.LocalPath).Split(':')[1]
                $OldProfileRenamed = $OldProfilePath + ".$($Date)"
                Rename-Item -Path $OldProfilePath $OldProfileRenamed -Force
                $SelectedProfile.Delete()
                [System.Windows.Forms.MessageBox]::Show("The profile was successfully removed","Success","OK","Information")
                AppendLog -LogText "Successfully removed $ProfilePath from $Computername"
            }
            Catch
            {
                AppendLog -LogText "Trying to retrieve and delete profile $SelectedProfile"
                [System.Windows.Forms.MessageBox]::Show("Something weird happened while renaming the path of $Username or removing the profile","Host unavailable ¯\_(ツ)_/¯","OK", "Error")
            }  
         }                     
    } 
    Else 
    { 
        AppendLog -LogText "Trying to connect to $Computername"
        [System.Windows.Forms.MessageBox]::Show("Unable to connect to host $Computername.","Host unavailable ¯\_(ツ)_/¯","OK", "Error") # This msg box will display if the Test-Connection fails
        return
    }   
}

#endregion

#region 'Graphical Interface
   
# Start instances of .Net form objects
$MainWindow            = New-Object System.Windows.Forms.Form
$GetProfiles_Button    = New-Object System.Windows.Forms.Button
$RemoveProfile_Button  = New-Object System.Windows.Forms.Button
$BackupProfile_Button  = New-Object System.Windows.Forms.Button
$RestoreProfile_Button = New-Object System.Windows.Forms.Button
$ProfileListBox        = New-Object System.Windows.Forms.ListBox 
$Hostname_Textbox      = New-Object System.Windows.Forms.TextBox
$progressBar           = New-Object System.Windows.Forms.ProgressBar
$GroupBox              = New-Object System.Windows.Forms.GroupBox
$Remove_ToolTip        = New-Object System.Windows.Forms.ToolTip
$Search_ToolTip        = New-Object System.Windows.Forms.ToolTip
$Backup_ToolTip        = New-Object System.Windows.Forms.ToolTip
$Restore_ToolTip       = New-Object System.Windows.Forms.ToolTip
$BackupToDialog        = New-Object System.Windows.Forms.FolderBrowserDialog
$RestoreFromDialog     = New-Object System.Windows.Forms.FolderBrowserDialog
$CurrentDir_Label      = New-Object System.Windows.Forms.Label
#$PictureBox            = New-Object System.Windows.Forms.PictureBox

# Define object properties
$MainWindow.Width           = 293
$MainWindow.Height          = 316
$MainWindow.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$MainWindow.StartPosition   = [System.Windows.Forms.FormStartPosition]::CenterScreen
$MainWindow.Text            = "Profile Reset Tool"
$MainWindow.Font            = "Arial,10"
$MainWindow.TopMost         = $true
$MainWindow.MaximizeBox     = $false

$GetProfiles_Button.Width                   = 25
$GetProfiles_Button.Height                  = 25
$GetProfiles_Button.Top                     = 15
$GetProfiles_Button.Left                    = 145
$GetProfiles_Button.UseVisualStyleBackColor = $true
$GetProfiles_Button.FlatStyle               = "Standard"
$GetProfiles_Button.Image = [System.IconExtractor]::Extract("imageres.dll", 7, $false)
#$GetProfiles_Button.Image = [System.IconExtractor]::Extract("wmploc.dll", 135, $true)
#$GetProfiles_Button.Image = [System.IconExtractor]::Extract("ieframe.dll", 70, $true)

    $Search_ToolTip.IsBalloon = $true
    $Search_ToolTip.SetToolTip($GetProfiles_Button, 'List user profiles')

$BackupProfile_Button.Width                   = 25
$BackupProfile_Button.Height                  = 25
$BackupProfile_Button.Top                     = 15
$BackupProfile_Button.Left                    = 172
$BackupProfile_Button.UseVisualStyleBackColor = $true
$BackupProfile_Button.FlatStyle               = "Standard"
#$BackupProfile_Button.Image = [System.IconExtractor]::Extract("Comres.dll", 4, $false)
$BackupProfile_Button.Image = [System.IconExtractor]::Extract("ieframe.dll", 12, $false)
#$BackupProfile_Button.Image = [System.IconExtractor]::Extract("ieframe.dll", 43, $false)

    $Backup_ToolTip.IsBalloon = $true
    $Backup_ToolTip.SetToolTip($BackupProfile_Button, 'Backup selected profile')

$RemoveProfile_Button.Width                   = 25
$RemoveProfile_Button.Height                  = 25
$RemoveProfile_Button.Top                     = 15
$RemoveProfile_Button.Left                    = 199
$RemoveProfile_Button.UseVisualStyleBackColor = $true
$RemoveProfile_Button.FlatStyle               = "Standard"
$RemoveProfile_Button.Image = [System.IconExtractor]::Extract("User32.dll", 3, $false)
#$RemoveProfile_Button.Image = [System.IconExtractor]::Extract("wmploc.dll", 135, $true)

    $Remove_ToolTip.IsBalloon = $true
    $Remove_ToolTip.SetToolTip($RemoveProfile_Button, 'Remove selected profile')

$RestoreProfile_Button.Width                   = 25
$RestoreProfile_Button.Height                  = 25
$RestoreProfile_Button.Top                     = 15
$RestoreProfile_Button.Left                    = 226
$RestoreProfile_Button.UseVisualStyleBackColor = $true
$RestoreProfile_Button.FlatStyle               = "Standard"
#$RestoreProfile_Button.Image = [System.IconExtractor]::Extract("ieframe.dll", 55, $false)
$RestoreProfile_Button.Image = [System.IconExtractor]::Extract("explorer.exe", 1, $false)

    $Restore_ToolTip.IsBalloon = $true
    $Restore_ToolTip.SetToolTip($RestoreProfile_Button, 'Restore selected profile')

$ProfileListBox.Location            = New-Object System.Drawing.Size(10,55)
$ProfileListBox.Size                = New-Object System.Drawing.Size(260,200)
$ProfileListBox.BackColor           = "white"
$ProfileListBox.ScrollAlwaysVisible = $true
$ProfileListBox.BorderStyle         = 2
$ProfileListBox.Font                = "Arial,11"

$Hostname_Textbox.Width       = 130
$Hostname_Textbox.Left        = 8
$Hostname_Textbox.Top         = 16
$Hostname_Textbox.TextAlign   = "Center"
$Hostname_Textbox.BorderStyle = 1
$Hostname_Textbox.Text        = $null
$Hostname_Textbox.ForeColor   = 'black'
$Hostname_Textbox.BackgroundImage = [System.IconExtractor]::Extract("imageres.dll", 236, $false)

$progressbar.Name      = 'ProgressBar'
$progressbar.Value     = 100
$progressbar.Style     = "blocks"
$progressbar.Location  = New-Object System.Drawing.Size(10,252) 
$progressbar.Size      = New-Object System.Drawing.Size(260,28)
$progressbar.Height    = 20
$progressbar.Visible   = $true
$progressBar.SendToBack()

#If (!([System.IO.File]::Exists("C:\Users\$env:USERNAME\AppData\Local\LoadingGif.jpg")))
#{[System.IO.File]::WriteAllBytes("C:\Users\$env:USERNAME\AppData\Local\LoadingGif.jpg",[Convert]::FromBase64String($LoadingGif))}
#$PictureBox.Image = [System.Drawing.Image]::FromFile("C:\Users\$env:USERNAME\AppData\Local\LoadingGif.jpg")
#$PictureBox.Size = New-Object System.Drawing.Size(350,350)
#$PictureBox.Location = New-Object System.Drawing.Size(0,10)
#$PictureBox.Visible = $True
#$PictureBox.Text = 'Doing some stuff...'

$GroupBox.Size     = New-Object System.Drawing.Size(260,50)
$GroupBox.location = New-Object System.Drawing.Size(10,0)

#endregion

#region 'Click events

#######
$GetProfiles_Button.ADD_Click({
    
    $ProfileListBox.Items.Clear()
    $Computername = $Hostname_Textbox.Text
    If ($Hostname_Textbox.Text -eq '')
    {
        $Computername = $env:COMPUTERNAME
        $Hostname_Textbox.Text = $env:COMPUTERNAME
    }

    If (!(Test-Connection -Quiet -ComputerName $Computername -Count 2))
    {
        [System.Windows.Forms.MessageBox]::Show("Could not connect to $Computername",'Invalid hostname   ¯\_(ツ)_/¯','ok','error')
        return
    }
    
    Try
    {  
        $UserProfiles = ((Get-WmiObject -Class Win32_UserProfile -ComputerName $ComputerName).Where({$_.LocalPath -like "*C:\Users\*"})).LocalPath
        # Populate list box with profile items
        Foreach ($Path in $UserProfiles) 
        {
            $Path = $Path -Replace ".*="
            $Path = $Path -Replace "}" 
            [void]$ProfileListBox.Items.Add("     $Path")
        }
    }
    Catch
    {
        [System.Windows.Forms.MessageBox]::Show("Could not connect to $Computername",'Invalid hostname   ¯\_(ツ)_/¯','ok','error')
    }
})
########
$RemoveProfile_Button.Add_Click({

    $Computername = $Hostname_Textbox.Text 
    If (!$ComputerName)
    {
        [System.Windows.Forms.MessageBox]::Show( "Who am I connecting to again...??","??????   ¯\_(ツ)_/¯","ok","Warning") 
        return 
    }
    Elseif (!$ProfileListBox.SelectedItem)
    {
        [System.Windows.Forms.MessageBox]::Show( "Select a user profile first","??????   ¯\_(ツ)_/¯","ok","Warning") 
        return         
    }

    AppendLog -Header
    AppendLog -LogText 'User Profile Remove_Event Started'
      
    If ([System.Windows.Forms.MessageBox]::Show("Are you sure you want to delete $($ProfileListBox.SelectedItem.Trim()) from $Computername?","Are you sure?","yesno","Question") -eq 'yes')
    {
        AppendLog -LogText "$env:username selected 'yes' to remove $SelectedProfile from $Computername"

        Try
        {
            $SelectedProfile = $ProfileListBox.SelectedItem.Trim()
            Remove-UserProfile -Computername $Computername -ProfilePath $SelectedProfile -ErrorAction Stop
            AppendLog -LogText "Sucessfully removed the profile"

            $UserProfiles = ((Get-WmiObject -Class Win32_UserProfile -ComputerName $ComputerName -ErrorAction Stop).Where({$_.LocalPath -like "*C:\Users\*"})).LocalPath        
            # Populate list box with profile items
            $ProfileListBox.Items.Clear()
            Foreach ($Path in $UserProfiles) 
            {               
                $Path = $Path -Replace ".*="
                $Path = $Path -Replace "}" 
                [void]$ProfileListBox.Items.Add("     $Path")
            }
        }
        Catch
        {
            [System.Windows.Forms.MessageBox]::Show("Failed to remove the profile $SelectedProfile from $Computername","Failed...   ¯\_(ツ)_/¯","ok","Error") 
            AppendLog -LogText "Failed to remove the profile"
        }
    }
    
    AppendLog -LogText 'User Profile Remove_Event Ended'
    AppendLog -Footer
})

$BackupProfile_Button.Add_Click({

    $Computername = $Hostname_Textbox.Text
    If (!$Hostname_Textbox.Text)
    {
        [System.Windows.Forms.MessageBox]::Show( "Who am I connecting to again...??","??????   ¯\_(ツ)_/¯","ok","Warning") 
        return 
    }
    Elseif (!$ProfileListBox.SelectedItem)
    {
        [System.Windows.Forms.MessageBox]::Show( "You must select a profile first","Invalid selection   ¯\_(ツ)_/¯","ok","Warning")       
        return
    }

    AppendLog -Header
    AppendLog -LogText 'User Profile Backup_Event Started'
    $SelectedProfile = $ProfileListBox.SelectedItem.Trim()
    $Username = $SelectedProfile.Split('\')[-1]
    $PSDrive = (ls function:[d-z]: -n|?{!(test-path $_)}|random)-replace':'

    Try
    {
        New-PSDrive -Name $PSDrive -PSProvider FileSystem -Root "\\$ComputerName\C$" -Persist -ErrorAction Stop
        AppendLog -LogText "Mapped drive $PSDrive to \\$Computername\C$"
        $BackupToDialog.RootFolder = "MyComputer"
        $BackupToDialog.SelectedPath = "$PSDrive"+":\"
        $BackupToDialog.Description = $PSDrive+": has been mapped to the root of $ComputerName.`n`nPlease select a save location local on this computer:"        
    }
    Catch
    {
        AppendLog -LogText "Trying to create remote connection on $ComputerName"
        [System.Windows.Forms.MessageBox]::Show( "$($Error[0].Exception.Message)","Backup failure   ¯\_(ツ)_/¯","ok","Error" )        
    }

    If ($BackupToDialog.ShowDialog() -eq 'Cancel'){return}
    $AreYouSure = [System.Windows.Forms.MessageBox]::Show( "Backup the profile for $Username on $Computername to $($BackupToDialog.SelectedPath)\Users\ ?","Are you sure?","YesNo","Information" )
    If ($AreYouSure -eq 'Yes')
    {        
        AppendLog -LogText "$nev:Username selected 'yes' to backup $SelectedProfile on $Computername to $($BackupToDialog.SelectedPath)"
        
        Try
        {           
            $BackupFrom = $PSDrive + ':' + (Split-Path -NoQualifier $SelectedProfile)
            Backup-UserProfile -Username $Username `
                                -Computername $Computername `
                                -Backup -BackupFrom $BackupFrom `
                                -BackupTo $BackupToDialog.SelectedPath `
                                -ErrorAction Stop

            [System.Windows.Forms.MessageBox]::Show( "The backup complete successfully","Backup complete!","ok","Information" ) 
            AppendLog -LogText "Successfully backed up the profile"
        }
        Catch
        {
            [System.Windows.Forms.MessageBox]::Show( "$($Error[0].Exception.Message)","Backup failure   ¯\_(ツ)_/¯","ok","Error" ) 
            AppendLog -LogText "Trying to backup $SelectedProfile on $Computername to $($BackupToDialog.SelectedPath)"
        }          
    }
    Else
    {
        AppendLog -LogText "$env:Username wasn't sure..."
    }

    Remove-PSDrive $PSDrive
    AppendLog -LogText 'User Profile Backup_Event Ended'
    AppendLog -Footer

})

$RestoreProfile_Button.Add_Click({

    $Computername = $Hostname_Textbox.Text
    If (!$Hostname_Textbox.Text)
    {      
        [System.Windows.Forms.MessageBox]::Show( "Enter a valid hostname","Hostname required   ¯\_(ツ)_/¯","ok","Warning" ) 
        return
    }
    Elseif (!(Test-Connection -Quiet -ComputerName $ComputerName -Count 1) )
    {
        AppendLog -LogText "Trying to connect to $ComputerName"
        [System.Windows.Forms.MessageBox]::Show( "Unable to connect to host $ComputerName","Connection error   ¯\_(ツ)_/¯","ok","Error" ) 
        return        
    }

    AppendLog -Header
    AppendLog -LogText 'User Profile Restore_Event Started'

    Try
    {
        $PSDrive = (ls function:[d-z]: -n|?{!(test-path $_)}|random)-replace':'
        AppendLog -LogText "Mapped $PSDrive on $Computername"
        New-PSDrive -Name $PSDrive -PSProvider FileSystem -Root "\\$ComputerName\C$" -Persist
        $RestoreFromDialog.RootFolder = 'MyComputer'
        $RestoreFromDialog.SelectedPath = "$PSDrive"+":\"
        $RestoreFromDialog.Description = $PSDrive + ": has been mapped to the root of $ComputerName.`n`n Select a profile backup to restore from:"
        If ($RestoreFromDialog.ShowDialog() -eq 'Cancel'){return}
        AppendLog -LogText "$nev:Username selected $($RestoreFromDialog.SelectedPath) to restore from"
        $UserProfiles  = ((Get-WmiObject -Class Win32_UserProfile -ComputerName $ComputerName).Where({$_.LocalPath -like "*C:\Users\*"})).LocalPath
        $Username      = ($RestoreFromDialog.SelectedPath).Split('\')[-1]
        $RestoreToPath = $( $UserProfiles.Where({$_ -like "*\$Username"})).Replace('C:',$($PSDrive + ':'))
        $RestoreTo = $RestoreToPath
    }
    Catch
    {
        AppendLog -LogText "Trying to create remote drive on $ComputerName"
        [System.Windows.Forms.MessageBox]::Show( "$($Error[0].Message)","Restore failure   ¯\_(ツ)_/¯","ok","Error" )      
        return  
    } 
      
    $AreYouSure = [System.Windows.Forms.MessageBox]::Show( "Restore $($RestoreFromDialog.SelectedPath) to $RestoreToPath`?`n`nClick NO to select another location","Are you sure?","YesNoCancel","Information" )
    If ($AreYouSure -eq 'no')
    {

       $RestoreToPath = New-Object System.Windows.Forms.FolderBrowserDialog
       $RestoreToPath.RootFolder = 'MyComputer'
       $RestoreToPath.SelectedPath = "$PSDrive"+":\"
       $RestoreToPath.Description = $PSDrive + ": has been mapped to the root of $ComputerName.`n`n Select a restore destination:"
       $RestoreToPath.ShowDialog()
       $RestoreTo = $RestoreToPath.SelectedPath
       AppendLog -LogText "Selected non-standard restore path $($RestoreToPath.SelectedPath)"

    }
    Elseif ($AreYouSure -eq 'cancel')
    {

        AppendLog -LogText "Cancelled by user"
        return

    }
    Elseif ( ($AreYouSure -eq 'yes'))
    {       
        Try
        {                                  
            AppendLog -LogText "Backing up $Username on $Computername to $RestoreToPath"

            Backup-UserProfile -Username $Username `
                               -Computername $Computername `
                               -Restore -RestoreFrom $RestoreFromDialog.SelectedPath `
                               -RestoreTo $RestoreTo `
                               -ErrorAction Continue

            AppendLog -LogText "Restore successful"
            [System.Windows.Forms.MessageBox]::Show("The restore completed successfully","Restore Complete","OK", "Information")

        }
        Catch
        {
            AppendLog -LogText "Trying to backup $Username on $ComputerName to $RestoreToPath"
            [System.Windows.Forms.MessageBox]::Show( "$($Error[0].Exception.Message)","Backup failure   ¯\_(ツ)_/¯","ok","Error" )            
        } 
    }
    Else
    {
        AppendLog -LogText "$env:Username wasn't sure..."
    }

    Remove-PSDrive $PSDrive
    AppendLog -LogText 'User Profile Restore_Event Ended'
    AppendLog -Footer

})
#endregion

#region 'Form controls
# Add controls to form objects
$GroupBox.Controls.Add($GetProfiles_Button)
$GroupBox.Controls.Add($RemoveProfile_Button)
$GroupBox.Controls.Add($Hostname_Textbox)
$GroupBox.Controls.Add($BackupProfile_Button)
$GroupBox.Controls.Add($RestoreProfile_Button)

#$MainWindow.Controls.Add($ProgressLabel)
$MainWindow.Controls.Add($progressBar)
$MainWindow.Controls.Add($ProfileListBox)
$MainWindow.Controls.Add($GroupBox)

$MainWindow.Controls.Add($PictureBox)

#$progressBar.Controls.Add($CurrentDir_Label)
$MainWindow.Controls.Add($CurrentDir_Label)

#endregion

$MainWindow.ShowDialog()