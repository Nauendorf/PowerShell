#region Global Variables
$ErrorActionPreference='stop'
$CurrentDir = (Get-Location).Path  
$TempDir = "$env:USERPROFILE\AppData\Local\Temp"
$Date = Get-Date -Format ddMyyyy
# Clear old temp files
Get-ChildItem -Path $TempDir -Filter ~Commencements*.csv | Remove-Item 
Get-ChildItem -Path $TempDir -Filter ~Terminations*.csv | Remove-Item
#endregion

#region Functions

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
Add-Type -TypeDefinition $code -ReferencedAssemblies System.Drawing # Loads C# assembly for icons

Function ExportToCsv ([ValidateScript({if ((Test-Path $_)){return $true}else{throw 'Invalid file path'}})]$ImportPath, $ExportPath)
{
    try
    {
        $Com = New-Object -ComObject Excel.Application
        $Com.Visible = $false
        $Com.DisplayAlerts = $false
        $wb = $Com.Workbooks.Open($ImportPath)
        $wb.Worksheets|%{$_.SaveAs($ExportPath, 6)}
        $Com.Quit()
    }
    catch 
    {
        throw "Unable to access Excel.Application Com"
    }
}

function Get-DomainUser
{
    param
    (
        [ValidateSet('LDAP://DC=dhw ,DC=wa, DC=gov, DC=au','LDAP://DC=ad, DC=dcd, DC=wa, DC=gov, DC=au','LDAP://DC=dsc ,DC=wa, DC=gov, DC=au')]
        [string]$SearchRoot,
        [string]$Identity,
        [string[]]$Properties  
    )

    $LDAPFilter = "(samAccountName=$Identity)"
    $searcher = New-Object DirectoryServices.DirectorySearcher
    $searcher.Filter = $LDAPFilter
    $Searcher.SearchRoot = $SearchRoot
    [adsi]$result = ($searcher.FindAll()).Path
    
    switch ($result.userAccountControl )
    {
        '512'   {$Status = 'Enabled'}
        '514'   {$Status = 'Disabled'}
        '66048' {$Status = 'Enabled, password never expires'}
        '66050' {$Status = 'Disabled, password never expires'}
    }

    if ($Properties -eq '*')
    {
        $table = [pscustomobject]@{}
        ($result|gm).Name | %{ Add-Member -InputObject $table -MemberType NoteProperty -Name $_ -Value $result.$_.ToString() }
    }
    else
    {

        $table = [pscustomobject]@{
            
            FullName          = "$($result.GivenName.ToString()) $($result.sn.ToString())"
            Status            = $Status.ToString()
            SamAccountName    = $result.SamAccountName.ToString()
            Email             = $result.Mail.ToString()
            GivenName         = $result.GivenName.ToString()
            Surname           = $result.sn.ToString()                                                                                   
            UserPrincipalName = $result.UserPrincipalName.ToString() 
            DistinguishedName = $result.DistinguishedName.ToString()            
        }

        if ($Properties)
        {$Properties|%{Add-Member -InputObject $table -MemberType NoteProperty -Name $_ -Value $result.$_.ToString()}}
    }

    if ($table){return $table}
    else {'no user object found'}

}

function CrossCheckReports ($Termiantions, $Commencements)
{
    #$ErrorActionPreference='silentlycontinue'
    $Comm = Import-Csv $Commencements
    $Term = Import-Csv $Termiantions
    
    try  
    {
        $StillEnabled = $Term.'Web User I' | %{Get-DomainUser -Identity $_ -SearchRoot 'LDAP://DC=dhw ,DC=wa, DC=gov, DC=au' | where {$_.Status -eq 'Enabled'} } 
    }
    catch
    {'no user found'}

    [array]$ToBeTerminated=@()

    foreach ($User in $StillEnabled)
    {
        if ($Comm.'Web User I' -notcontains $User.SamAccountName)
        {
            [array]$ToBeTerminated += $User
        }
    }

    $ErrorActionPreference='stop'
    return $ToBeTerminated
}
 
#endregion

#region Declare GUI objects

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[System.Windows.Forms.Application]::EnableVisualStyles()
[System.Windows.Forms.Cursor]::Show()

$Form       = New-Object System.Windows.Forms.Form
$Commen_pic = New-Object System.Windows.Forms.PictureBox
$Termin_pic = New-Object System.Windows.Forms.PictureBox
$Output_txt = New-Object System.Windows.Forms.RichTextBox
$Import_Grp = New-Object System.Windows.Forms.GroupBox
$Output_Grp = New-Object System.Windows.Forms.GroupBox
$Go_btn     = New-Object System.Windows.Forms.Button
$Refr_btn   = New-Object System.Windows.Forms.Button
$Font       = New-Object System.Drawing.Font("Terminal",9)
$Icon       = New-Object system.drawing.icon("$CurrentDir\report_icon.ico", 192, 192)
$OutputMenu = New-Object System.Windows.Forms.ContextMenuStrip
$Commen_tip = New-Object System.Windows.Forms.ToolTip
$Term_tip   = New-Object System.Windows.Forms.ToolTip
$SaveAs     = [System.Windows.Forms.SaveFileDialog]::new()

#endregion

#region Set GUI attributes

$Form.Size            = New-Object System.Drawing.Size(490,205)
$Form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D
$Form.StartPosition   = [System.Windows.Forms.FormStartPosition]::CenterScreen
$Form.Text            = "User Status Report"
$Form.TopMost         = $true
$Form.MaximizeBox     = $false
$Form.ShowIcon        = $true
$Form.Icon            = $Icon
$Form.AllowDrop = $true

    $Refr_btn.Size      = New-Object System.Drawing.Size(25,25)
    $Refr_btn.Location  = New-Object System.Drawing.Size(262,50)
    $Refr_btn.FlatStyle = 'flat'
    $Refr_btn.BackgroundImage = [System.IconExtractor]::Extract("hgcpl.dll", 3, $true)
    $Refr_btn.BackgroundImageLayout = [System.Windows.Forms.ImageLayout]::Stretch
    $Refr_btn.FlatAppearance.BorderSize = 0

    $Go_btn.Size        = New-Object System.Drawing.Size(25,25)
    $Go_btn.Location    = New-Object System.Drawing.Size(263,90)
    $Go_btn.FlatStyle   = 'flat'
    $Go_btn.BackgroundImage = [System.IconExtractor]::Extract("ieframe.dll", 100, $true)
    $Go_btn.BackgroundImageLayout = [System.Windows.Forms.ImageLayout]::Stretch
    $Go_btn.FlatAppearance.BorderSize = 0

    $Form.Controls.Add($Import_Grp)
    $Form.Controls.Add($Output_Grp)
    $Form.Controls.Add($Go_btn)
    $Form.Controls.Add($Refr_btn)


$Import_Grp.Text     = 'Drag && Drop your spreadsheets here'
$Import_Grp.Size     = New-Object System.Drawing.Size(240,145)
$Import_Grp.Location = New-Object System.Drawing.Size(10,10)

    $Commen_pic.Text        = 'Commencements'
    $Commen_pic.Size        = New-Object System.Drawing.Size(100,100)
    $Commen_pic.Location    = New-Object System.Drawing.Size(15,27)
    $Commen_pic.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
    $Commen_pic.AllowDrop   = $true        
    $Commen_pic.Anchor      = [System.Windows.Forms.AnchorStyles]::Bottom
    $Comm_img               = [System.Drawing.Image]::FromFile("$CurrentDir\Commencements.png")
    $Commen_pic.Image       = $Comm_img

    $Commen_tip.SetToolTip($Commen_pic, 'Commencements Spreadsheet')

    $Termin_pic.Text        = 'Terminations'
    $Termin_pic.Size        = New-Object System.Drawing.Size(100,100)
    $Termin_pic.Location    = New-Object System.Drawing.Size(125,27)
    $Termin_pic.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
    $Termin_pic.AllowDrop   = $true
    $Termin_pic.Anchor      = [System.Windows.Forms.AnchorStyles]::Bottom
    $Term_img               = [System.Drawing.Image]::FromFile("$CurrentDir\Terminations.png")
    $Termin_pic.Image       = $Term_img

    $Term_tip.SetToolTip($Termin_pic, 'Terminations Spreadsheet')

    $Import_Grp.Controls.Add($Commen_pic)
    $Import_Grp.Controls.Add($Termin_pic)

$Output_Grp.Text     = "To be investigated"
$Output_Grp.Size     = New-Object System.Drawing.Size(160,145)
$Output_Grp.Location = New-Object System.Drawing.Size(300,10)

    [void]$OutputMenu.Items.Add('Save')

    $Output_txt.Size        = New-Object System.Drawing.Size(130,100)
    $Output_txt.Location    = New-Object System.Drawing.Size(15,27)
    $Output_txt.ReadOnly    = $true
    $Output_txt.ContextMenuStrip = $OutputMenu

    $Output_Grp.Controls.Add($Output_txt)

#endregion

#region Events

$Commen_DragOver=[System.Windows.Forms.DragEventHandler]{
#Event Argument: $_ = [System.Windows.Forms.DragEventArgs]
$_.Effect = [System.Windows.DragDropEffects]::Copy
    if ($_.Data.GetDataPresent([Windows.Forms.DataFormats]::FileDrop))
    {$_.Effect = [System.Windows.DragDropEffects]::Copy}
    else
    {$_.Effect = [System.Windows.DragDropEffects]::None}
}

$Commen_DragDrop=[System.Windows.Forms.DragEventHandler]{
#Event Argument: $_ = [System.Windows.Forms.DragEventArgs]
   
    $Commen_pic.Cursor = 'Hand'
    $Form.Text = 'Loading...'
    [string[]] $files = [string[]]$_.Data.GetData([Windows.Forms.DataFormats]::FileDrop)

    if ($files)
    {
        $Commen_pic.Cursor = 'WaitCursor'
        ExportToCsv -ImportPath $files[0] -ExportPath "$TempDir\~Commencements$Date.csv"          
        $Pic = Get-Item "$CurrentDir\excel-icon-sml.png"
        $img1 = [System.Drawing.Image]::FromFile($Pic)
        $Commen_pic.Image = $img1
        $Commen_pic.AllowDrop = $false           
    } 
    [System.Windows.Forms.Application]::DoEvents() 
    $Form.Text = 'User Status Report'  
    $Commen_pic.Cursor = 'Arrow'
}

$Terminate_DragOver=[System.Windows.Forms.DragEventHandler]{
#Event Argument: $_ = [System.Windows.Forms.DragEventArgs]
    $Termin_pic.Cursor = 'Hand'
    if ($_.Data.GetDataPresent([Windows.Forms.DataFormats]::FileDrop))
    {$_.Effect =  [System.Windows.DragDropEffects]::Copy}
    else
    {$_.Effect = [System.Windows.DragDropEffects]::None}
}

$Terminate_DragDrop=[System.Windows.Forms.DragEventHandler]{
#Event Argument: $_ = [System.Windows.Forms.DragEventArgs]

    $Termin_pic.Cursor = 'Hand'
    $Form.Text = 'Loading...'  
    [string[]] $files = [string[]]$_.Data.GetData([Windows.Forms.DataFormats]::FileDrop)
    if ($files)
    {      
        $Termin_pic.Cursor = 'WaitCursor'    
        ExportToCsv -ImportPath $files[0] -ExportPath "$TempDir\~Terminations$Date.csv"
        $Pic = Get-Item "$CurrentDir\excel-icon-sml.png"
        $img2 = [System.Drawing.Image]::FromFile($Pic)
        $Termin_pic.Image = $img2
        $Termin_pic.AllowDrop = $false           
    }    
    [System.Windows.Forms.Application]::DoEvents() 
    $Form.Text = 'User Status Report'
    $Termin_pic.Cursor = 'Arrow'
}

$Go_btn.Add_Click({
    
    $Form.Cursor = 'waitcursor'
    Get-ChildItem -Path $TempDir -Filter ~Commencements*.csv -Exclude ~Commencements$Date.csv | Remove-Item 
    Get-ChildItem -Path $TempDir -Filter ~Terminations*.csv -Exclude ~Terminations$Date.csv | Remove-Item  

    $Output_txt.Clear()
    $Form.Text = 'Loading...'
    if ( ([System.IO.File]::Exists("$TempDir\~Terminations$Date.csv") ) -and ([System.IO.File]::Exists("$TempDir\~Commencements$Date.csv") ) )
    {
        $Global:ToBeTerminated = CrossCheckReports "$TempDir\~Terminations$Date.csv" "$TempDir\~Commencements$Date.csv"
        $ToBeTerminated|%{
        
            $First = $_.GivenName
            $Last  = $_.Surname
            $Output_txt.AppendText("$First $Last`n")
            [System.Windows.Forms.Application]::DoEvents()
        }   
    }
    else
    {
        [System.Windows.Forms.MessageBox]::Show('No spreadsheets have been loaded','Error','OK',[System.Windows.Forms.MessageBoxIcon]::Error)
    }
    $Form.Text = 'User Status Report'
    $Form.Cursor = 'Arrow'
})

$Refr_btn.Add_Click({

    Get-ChildItem -Path $TempDir -Filter ~Commencements*.csv | Remove-Item 
    Get-ChildItem -Path $TempDir -Filter ~Terminations*.csv | Remove-Item  

    $Comm_img             = [System.Drawing.Image]::FromFile("$CurrentDir\Commencements.png")
    $Commen_pic.Image     = $Comm_img
    $Commen_pic.AllowDrop = $true   

    $Term_img             = [System.Drawing.Image]::FromFile("$CurrentDir\Terminations.png")
    $Termin_pic.Image     = $Term_img
    $Termin_pic.AllowDrop = $true 

    $Output_txt.Clear()
})

$Output_txt.Add_MouseDoubleClick({

    if (!$Output_txt.Text -eq '')
    { $ToBeTerminated | Out-GridView -Title 'User Status Report' }
})

$OutputMenu.Add_Click({

    if ($Output_txt.Text -ne '')
    {
        $SaveAs.Filter = "CSV Files (*.csv)|"
        if ($SaveAs.ShowDialog() -eq 'Cancel'){return}
        $ToBeTerminated | Export-Csv -Force -NoClobber -NoTypeInformation -Path "$($SaveAs.FileName).csv"
        Invoke-Item -Path "$($SaveAs.FileName).csv"
    }
})

$Commen_pic.Add_DragEnter($Commen_DragDrop)
$Commen_pic.Add_DragOver($Commen_DragOver)

$Termin_pic.Add_DragEnter($Terminate_DragDrop)
$Termin_pic.Add_DragOver($Terminate_DragOver)

#endregion

$Form.ShowDialog()
