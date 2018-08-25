# Load assemblies for GUI objects
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") 
[System.Windows.Forms.Application]::EnableVisualStyles()
$DotNetVersion = (Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP' -recurse |
Get-ItemProperty -name Version,Release -EA 0 |
Where { ($_.PSChildName -match '^(?!S)\p{L}') -and ($_.PSChildName -eq 'Full') } ).Version
$VersionCheck=($PSVersionTable.PSVersion.Major)
If (($PSVersionTable.PSVersion.Major)-lt 4)
{[Microsoft.VisualBasic.Interaction]::MsgBox("This script has not been tested on PowerShell version 3 or lower. $env:COMPUTERNAME is using version $VersionCheck. You may experience errors.", "okonly,SystemModal,Exclamation", "Warning")}
#Elseif ($DotNetVersion -lt '4.6')
#{[Microsoft.VisualBasic.Interaction]::MsgBox("You are not running the .Net 4.5 Framework. You may experience errors.", "okonly,SystemModal,Exclamation", "Warning")}

# Get Domain Admin Authentication 
If (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))
{   
$arguments = "& '" + $myinvocation.mycommand.definition + "'"
Start-Process PowerShell.exe -Verb runAs -ArgumentList $arguments
Break
}

$Script:showWindowAsync = Add-Type -MemberDefinition @"
[DllImport("user32.dll")]
public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
"@ -Name "Win32ShowWindowAsync" -Namespace Win32Functions -PassThru
Function Show-Powershell()
{
$null = $showWindowAsync::ShowWindowAsync((Get-Process -Id $pid).MainWindowHandle, 10)
}
Function Hide-Powershell()
{
$null = $showWindowAsync::ShowWindowAsync((Get-Process -Id $pid).MainWindowHandle, 2)
}
Hide-Powershell

Function Get-InstalledSoftware
{
Param
(
[Alias('Computer','ComputerName','HostName')]
[Parameter(ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$true,Position=1)]
[string[]]$Name = $env:COMPUTERNAME
)
Begin
{
$LMkeys = "Software\Microsoft\Windows\CurrentVersion\Uninstall","SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
$LMtype = [Microsoft.Win32.RegistryHive]::LocalMachine
$CUkeys = "Software\Microsoft\Windows\CurrentVersion\Uninstall"
$CUtype = [Microsoft.Win32.RegistryHive]::CurrentUser
}
Process
{
ForEach($Computer in $Name)
{
$MasterKeys = @()
If(!(Test-Connection -ComputerName $Computer -count 1 -quiet))
{
Write-Error -Message "Unable to contact $Computer. Please verify its network connectivity and try again." -Category ObjectNotFound -TargetObject $Computer
Break
}
$CURegKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($CUtype,$computer)
$LMRegKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($LMtype,$computer)
ForEach($Key in $LMkeys)
{
$RegKey = $LMRegKey.OpenSubkey($key)
If($RegKey -ne $null)
{
ForEach($subName in $RegKey.getsubkeynames())
{
foreach($sub in $RegKey.opensubkey($subName))
{
$MasterKeys += (New-Object PSObject -Property @{
"ComputerName" = $Computer
"Name" = $sub.getvalue("displayname")
"SystemComponent" = $sub.getvalue("systemcomponent")
"ParentKeyName" = $sub.getvalue("parentkeyname")
"Version" = $sub.getvalue("DisplayVersion")
"UninstallCommand" = $sub.getvalue("UninstallString")
})
}
}
}
}
ForEach($Key in $CUKeys)
{
$RegKey = $CURegKey.OpenSubkey($Key)
If($RegKey -ne $null)
{
ForEach($subName in $RegKey.getsubkeynames())
{
foreach($sub in $RegKey.opensubkey($subName))
{
$MasterKeys += (New-Object PSObject -Property @{
"ComputerName" = $Computer
"Name" = $sub.getvalue("displayname")
"SystemComponent" = $sub.getvalue("systemcomponent")
"ParentKeyName" = $sub.getvalue("parentkeyname")
"Version" = $sub.getvalue("DisplayVersion")
"UninstallCommand" = $sub.getvalue("UninstallString")
})
}
}
}
}
$MasterKeys = ($MasterKeys | Where {$_.Name -ne $Null -AND $_.SystemComponent -ne "1" -AND $_.ParentKeyName -eq $Null} | select Name,Version,ComputerName,UninstallCommand | sort Name)
$MasterKeys
}
}
End
{
}
}

Function Compare-InstalledApps {
Param (
[Parameter(Mandatory=$true)]
[string]$BaselineFile,
[Parameter(Mandatory=$true)]
[string]$Computers,
[Parameter(Mandatory=$true)]
[string]$SavePath
)
$Date = Get-Date -Format ddMyyy
$table=@()
$ComputerList = Get-Content $Computers
$Baseline = Import-Csv $BaselineFile
$i=0

Foreach ($Computer in $ComputerList)
{
    If ($computer -eq $null){Write-Host "No object in text file" -ForegroundColor Red ; Break}
    ElseIf ((Test-Connection -ComputerName $Computer -Count 2 -Quiet -ErrorAction Continue)-eq $false)
        {
            Write-Host "Unable to connect to host $Computer" -ForegroundColor red
            Out-File -Append -Force -InputObject "$Computer" -FilePath "$SavePath\Offline_Computers-$Date.txt"
            Show-Powershell
        }
    Else{Write-Host "Processing $Computer" -ForegroundColor Green}

    # This is the list of software installed on the target computer(s)#    
    $CurrentList = Get-InstalledSoftware -Name $Computer

    Foreach ($Item in $CurrentList.Name)
    {
        If ($Baseline.Name -notcontains $Item)
        {
            $Table += [pscustomobject]@{
            Computer = $Computer
            Application = $Item
            }  
        }   
        $Table | Export-Csv -Path "$SavePath\$computer-$Date.csv" -Force -NoClobber -NoTypeInformation -Append
        $table=@()
    }   
    $progress.Value = ($i/$ComputerList.Count)*100
    $i++
}
$progress.Value = '100'
[Microsoft.VisualBasic.Interaction]::MsgBox("Operation complete!", "OKOnly,SystemModal,Information", "Finished!")
}

Add-Type -AssemblyName System.Windows.Forms
$ErrorActionPreference = 'SilentlyContinue'
$MainForm    = New-Object System.Windows.Forms.Form
$Groupbx     = New-Object System.Windows.Forms.GroupBox
$GenGroupbx  = New-Object System.Windows.Forms.GroupBox
$BaseBtn     = New-Object System.Windows.Forms.Button
$PCListBtn   = New-Object System.Windows.Forms.Button
$SaveBtn     = New-Object System.Windows.Forms.Button
$GoButton    = New-Object System.Windows.Forms.Button
$GenerateBtn = New-Object System.Windows.Forms.Button
$BasePthTxt  = New-Object System.Windows.Forms.RichTextBox
$PCPthTxt    = New-Object System.Windows.Forms.RichTextBox
$SavePthTxt  = New-Object System.Windows.Forms.RichTextBox
$GenerateTxt = New-Object System.Windows.Forms.RichTextBox
$Progress    = New-Object System.Windows.Forms.ProgressBar
$Font = New-Object System.Drawing.Font("System",10,[System.Drawing.FontStyle]::Italic)

$SaveFileDialog = New-Object windows.forms.FolderBrowserDialog
$BaseFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$PCFileDialog   = New-Object System.Windows.Forms.OpenFileDialog
$GenerateDialog = New-Object System.Windows.Forms.SaveFileDialog

# Properties of form objects
$MainForm.Width           = 375
$MainForm.Height          = 235
$MainForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D
$MainForm.StartPosition   = [System.Windows.Forms.FormStartPosition]::CenterScreen
$MainForm.Text            = "Compare Installed Software"
$MainForm.Font            = "Arial,10"
$MainForm.BackColor       = "#757170"
$MainForm.TopMost         = $true
$MainForm.MaximizeBox     = $false
$MainForm.ShowIcon        = $true

$Groupbx.Size     = New-Object System.Drawing.Size(294,120)
$Groupbx.Location = New-Object System.Drawing.Size(10,67)
#$Groupbx.Text     = "Compare Installed Software"
$Groupbx.SendToBack()

$GenGroupbx.Size     = New-Object System.Drawing.Size(294,58)
$GenGroupbx.Location = New-Object System.Drawing.Size(10,10)
$GenGroupbx.Text     = "Generate a baseline file"
$GenGroupbx.Font = New-Object System.Drawing.Font("System",9)
$GenGroupbx.ForeColor = 'White'
$GenGroupbx.SendToBack()

$BaseBtn.Text = 'Base File'
$BaseBtn.Location = New-Object System.Drawing.Size(10,15)
$BaseBtn.Size = New-Object System.Drawing.Size(80,25)
$BaseBtn.Font = 'Segoe UI,9'
$BaseBtn.FlatStyle = [System.Windows.Forms.FlatStyle]::System
$BaseBtn.BackColor = '#856259'

$BasePthTxt.Size     = New-Object System.Drawing.Size(190,25)
$BasePthTxt.Location = New-Object System.Drawing.Size(95,15)
$BasePthTxt.ReadOnly = $true

$PCListBtn.Text = 'Computers'
$PCListBtn.Location = New-Object System.Drawing.Size(10,50)
$PCListBtn.Size = New-Object System.Drawing.Size(80,25)
$PCListBtn.Font = 'Segoe UI,9'
$PCListBtn.FlatStyle = [System.Windows.Forms.FlatStyle]::System

$PCPthTxt.Size     = New-Object System.Drawing.Size(190,25)
$PCPthTxt.Location = New-Object System.Drawing.Size(95,50)
$PCPthTxt.ReadOnly = $true

$SaveBtn.Text = 'Save Path'
$SaveBtn.Location = New-Object System.Drawing.Size(10,85)
$SaveBtn.Size = New-Object System.Drawing.Size(80,25)
$SaveBtn.Font = 'Segoe UI,9'
$SaveBtn.FlatStyle = [System.Windows.Forms.FlatStyle]::System

$SavePthTxt.Size     = New-Object System.Drawing.Size(190,25)
$SavePthTxt.Location = New-Object System.Drawing.Size(95,85)
$SavePthTxt.ReadOnly = $true

$GenerateBtn.Text = 'Generate'
$GenerateBtn.Location = New-Object System.Drawing.Size(10,20)
$GenerateBtn.Size = New-Object System.Drawing.Size(80,25)
$GenerateBtn.Font = 'Segoe UI,9'
$GenerateBtn.FlatStyle = [System.Windows.Forms.FlatStyle]::System

$GenerateTxt.Size     = New-Object System.Drawing.Size(190,25)
$GenerateTxt.Location = New-Object System.Drawing.Size(95,21)
$GenerateTxt.ReadOnly = $true
$GenerateTxt.Text = 'Enter computername for baseline'
$GenerateTxt.Font = 'Segoe UI,9'

$GoButton.Text = 'GO!'
$GoButton.Size = New-Object System.Drawing.Size(40,35)
$GoButton.Location = New-Object System.Drawing.Size(310,16)
$GoButton.FlatStyle = [System.Windows.Forms.FlatStyle]::System

$Progress.Location = New-Object System.Drawing.Size(1,2)
$Progress.Size = New-Object System.Drawing.Size(358,5)
$Progress.BackColor = 'green'
$Progress.Maximum = '100'

$BaseFileDialog.initialDirectory = 'C:\Temp\'
$BaseFileDialog.filter = ".CSV files (*.csv)|*.csv"
$BaseFileDialog.ShowHelp = $true
$BaseFileDialog.Title = 'Select a baseline file'

$PCFileDialog.initialDirectory = 'C:\Temp\'
$PCFileDialog.filter = ".txt files (*.txt)|*.txt"
$PCFileDialog.ShowHelp = $true
$PCFileDialog.Title = 'Select a TXT list of computers to query'

$GenerateDialog.initialDirectory = 'C:\Temp\'
$GenerateDialog.filter = ".csv files (*.csv)|*.csv"
$GenerateDialog.Title = 'Select a location to save baseline file'

$SaveFileDialog.Description = 'Select a folder to save the output for each computer query'

# Event Handlers
$Generate_Event = [System.EventHandler]{

If (($GenerateDialog.ShowDialog()) -ne 'Cancel'){
    $SaveLocation = $GenerateDialog.FileName
    $Computername = $GenerateTxt.Text
    Get-InstalledSoftware -Name $Computername |
    Export-Csv -LiteralPath $SaveLocation -Force -NoClobber -NoTypeInformation
    }
}

$BaseFile_Event = [System.EventHandler]{
$BaseFileDialog.ShowDialog()
$BasePthTxt.Text = $BaseFileDialog.FileName
}

$PCList_Event = [System.EventHandler]{
$PCFileDialog.ShowDialog()
$PCPthTxt.Text = $PCFileDialog.FileName
}

$Save_Event = [System.EventHandler]{
$SaveFileDialog.ShowDialog()
$SavePthTxt.Text = $SaveFileDialog.SelectedPath
}

$GO_Event = [System.EventHandler]{
If (($BaseFileDialog.FileName) -eq '')
    {[Microsoft.VisualBasic.Interaction]::MsgBox("You must select a baseline file", "OKOnly,SystemModal,Exclamation", "Warning")}
ElseIf(($PCFileDialog.FileName)-eq '')
    {[Microsoft.VisualBasic.Interaction]::MsgBox("You must select a txt list of computers", "OKOnly,SystemModal,Exclamation", "Warning")}
ElseIf(($SaveFileDialog.SelectedPath)-eq '')
    {[Microsoft.VisualBasic.Interaction]::MsgBox("You must select a save location", "OKOnly,SystemModal,Exclamation", "Warning")}
Else
    {
        $Progress.Text = 'Processing...'
        Compare-InstalledApps -BaselineFile $BaseFileDialog.FileName -Computers $PCFileDialog.FileName -SavePath $SaveFileDialog.SelectedPath 
    }
}

$Txtbx_Event = [System.EventHandler]{
$GenerateTxt.Text = $null
$GenerateTxt.ReadOnly = $false
}

# Click Events
$BaseBtn.ADD_Click($BaseFile_Event)
$PCListBtn.ADD_Click($PCList_Event)
$SaveBtn.ADD_Click($Save_Event)
$GoButton.Add_Click($GO_Event)
$GenerateBtn.Add_Click($Generate_Event)
$GenerateTxt.Add_Click($Txtbx_Event)

# Controls
$Groupbx.Controls.Add($BaseBtn)
$Groupbx.Controls.Add($PCListBtn)
$Groupbx.Controls.Add($SaveBtn)
$Groupbx.Controls.Add($BasePthTxt)
$Groupbx.Controls.Add($PCPthTxt)
$Groupbx.Controls.Add($SavePthTxt)
$GenGroupbx.Controls.Add($GenerateBtn)
$GenGroupbx.Controls.Add($GenerateTxt)

$MainForm.controls.Add($Groupbx)
$MainForm.Controls.Add($GenGroupbx)
$MainForm.Controls.Add($GoButton)
$MainForm.Controls.Add($Progress)
$MainForm.ShowDialog()