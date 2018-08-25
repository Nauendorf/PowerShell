<#	
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2017 v5.4.143
	 Created on:   	15/10/2017 6:15 PM
	 Created by:   	David
	 Organization: 	
	 Filename:     	SD_Tools.psm1
	-------------------------------------------------------------------------
	 Module Name: SD_Tools
	===========================================================================
#>

. {
	
	Function Get-SDTools
	{
		$CommandList = @"
===============================================================================================

    Test-Admin                : Checks whether the current PowerShell session is elevated
                              
    Invoke-Admin              : Starts an elevated PowerShell sessions
                              
    Get-User                  : Retrieves the currently logged on user for a given hostname
                              
    Get-LastHost              : Retrieves the logon history for a given username
                              
    Get-LogonHistory          : Retrieves the logon history for a given hostname
                              
    Get-InstalledSoftware     : Retrieves installed software from a given hostname. Better than Win32_Product
                              
    Get-SystemReport          : Generates a system report containing hardware & software information
                              
    Test-Credential           : Tests validity of the supplied credentials 

    Get-BitLockerRecoveryKey  : Retrieves the bitlocker recovery key for a given device

===============================================================================================
"@
		
		
		Write-Host $CommandList
		
	}
	
	Function Test-Admin
	{
    <#
    .Description
    Checks whether the current session is running under as an Administrator

    .Synopsis
    Type Test-Admin to check whether the current session is running as an Administrator
    #>
		
		param
		(
			[CmdletBinding(DefaultParameterSetName = 'Domain')]
			[parameter(ParameterSetName = 'Domain')]
			[switch]$Domain,
			[switch]$Local
		)
		
		If (-NOT ([Security.Principal.WindowsPrincipal]`
				[Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))
		{
			Write-Host "This is [NOT] an elevated PowerShell session" -ForegroundColor Red
		}
		else
		{
			Write-Host "This [IS] an elevated PowerShell session" -ForegroundColor Green
		}
	}
	
	Function Invoke-Admin
	{
    <#
    .Description
    The Invoke-Admin cmdlet starts an elevated PowerShell session

    .Synopsis
    Type [su] to open an elevated PowerShell window
    #>
		
		Try
		{
			Start-Process -FilePath "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Accessories\Windows PowerShell\Windows PowerShell.lnk" -Verb Runas
		}
		Catch
		{
			$Error[0].Exception
			#Write-Host "The operation was cancelled by the user"
		}
	}
	
	Function Get-User
	{
    <#
    .Description
    The Get-User cmdlet will retrieve the currently logged on user for a given computer

    .Synopsis
    Type Get-User <hostname> to get the username for any accounts currently logged on to that computer
    #>
		
		Param
		(
			[CmdletBinding(DefaultParameterSetName = 'local')]`
			[Parameter(ValueFromPipeline = $true)]
			[String[]]$Computername
		)
		
		If (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))
		{
			Write-Host "This command requires an elevated command prompt" -ForegroundColor Red
			Break
		}
		
		If (!$Computername)
		{
			$Computername = $ENV:ComputerName
		}
		
		If (Test-Connection -ComputerName $Computername -Count 1 -Quiet)
		{
			
			Try
			{
				$User = Get-WmiObject -Computer $Computername -Class Win32_ComputerSystem
				$User = ($User.Username).Split('\')[1]
			}
			Catch
			{
				Write-Host "Error: Something went wrong!" -ForegroundColor Red
			}
			
			If (!($user))
			{
				Write-Host "No user is currently logged in..." -ForegroundColor Red
			}
			Else
			{
				Write-Host $User -ForegroundColor Green
			}
			
		}
		Else
		{
			Write-Host "Error: Unable to connect to host" -ForegroundColor Red
		}
	}
	
	Function Get-LastHost
	{
    <# 
    .Description 
    The Get-LastHost cmdlet retrieves the hostname for computers that were recently logged on to for a given account 
 
 
    .Synopsis 
    Type Get-LastHost <username> [number] to retrieve the last host that a specified user logged on to. Use the [number] parameter to list multiple hosts that the user has logged on to. 
    #>		
		
		Param ($User,
			$Num)
		If (!$User) { Write-Host "Error! You must enter a username" -ForegroundColor Red; Break }
		If (!$Num) { $Num = "1" }
		
		$usercsv = ($user + ".csv")
		New-PSDrive -Name P -PSProvider FileSystem -Root \\ServerName\UserLoggingData$\LOGON\ | Out-Null
		$csvPath = "P:\$usercsv"
		
		If (Test-Path $csvPath)
		{
			
			$Header = "Date", "Username", "Site Code", "Hostname", "IP Address"
			
			$hashTable = @{
				Date		  = $csvPath.Date
				Username	  = $csvPath.Username
				'Site Code'   = $csvPath.'Site Code'
				Hostname	  = $csvPath.Hostname
				'IP Address'  = $csvPath.'IP Address'
			}
			
			Import-Csv P:\$usercsv -Header $Header |
			Select-Object -Last $Num
			#Sort-Object -Property Date -Descending
			#Format-Table -AutoSize -Wrap
			
		}
		Else { Write-Host "Error: There is no logon history for $User" -ForegroundColor Red }
	}
	
	Function Get-LogonHistory
	{
    <# 
    .Description 
    The Get-LogonHistory cmdlet retrieves the date, time, and username of accounts that have recently logged on to a specified computer.
 
 
    .Synopsis 
    Type Get-LogonHistory <hostname> 
    #>		
		
		Param ([string]$Computername,
			[int]$num)
		If (!$Computername) { Write-Host "Error! You must enter a hostname" -ForegroundColor Red; Break }
		If (!$num) { $num = "1" }
		Write-Host "Retrieving data. Please wait... "
		$Original = Get-Location
		$FreeDrive = ls function:[d-z]: -n | ?{ !(test-path $_) } | random
		$Drive = $FreeDrive -replace ":"
		$Path = New-PSDrive -Name $Drive -PSProvider FileSystem -Root \\ServerName\UserLoggingData$\Logon | Out-Null; CD $FreeDrive
		$Directory = Get-ChildItem $FreeDrive
		$Header = "Date", "Username", "Site Code", "Hostname", "IP Address"
		
		ForEach ($CSV in $Directory)
		{
			$Import = Import-Csv $CSV -Header $Header
			$Import | Where-Object { $_.Hostname -like $Computername }
		}
		CD $Original
	}
	
	Function Get-SystemReport
	{
    <# 
    .Description 
    The Get-SystemReport generates a report of basic system data and recent system errors.
  
    .Synopsis 
    Type Get-SystemReport  <hostname> 
    #>		
		
		[CmdletBinding(DefaultParameterSetName = "NoSave")]
		[CmdletBinding(SupportsShouldProcess = $True, ConfirmImpact = 'Low')]
		Param
		(
			[parameter(ParameterSetName = 'NoSave', Position = 0)]
			[parameter(ParameterSetName = 'SavePath', Position = 0)]
			[Parameter(ValueFromPipeline = $true)]
			[string]$ComputerName,
			### Parameter 1    

			[parameter(ParameterSetName = 'SavePath', Mandatory = $True, HelpMessage = 'Output report to either csv or html', Position = 1)]
			[ValidateSet('csv', 'html')]
			[AllowNull()]
			[string]$ReportType,
			### Parameter 2

			[parameter(ParameterSetName = 'SavePath', Mandatory = $True, HelpMessage = 'Save path must contain a valid file extension', Position = 2)]
			[ValidatePattern('csv|html')]
			[ValidateScript({ Test-Path -Path $_ -PathType 'Any' -IsValid })]
			[string]$SavePath,
			### Parameter 3  

			[switch]$HotFixes,
			### Parameter 4

			[switch]$CriticalErrors ### Patameter 5
		)
		
		# Check for administrator rights, functions ends if not run from an elevated session
		If (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))
		{
			Write-Host "Error: this function must be run from an elevated PowerShell session" -ForegroundColor Red
			Break
		}
		
		# Validate dependencies
		$ErrorActionPreference = "SilentlyContinue"
		If (!$ComputerName) { $ComputerName = $env:COMPUTERNAME }
		If (($ComputerName -ne $env:COMPUTERNAME) -and ($ComputerName -ne 'localhost')) # If computername variable is local then skip connection test
		{
			If ((Test-Connection -Quiet -Count 1 -ComputerName $ComputerName) -eq $false)
			{
				Write-Error -ErrorId 'Invalid hostname' -Exception 'Connection error' -Message "Error: Unable to connect to host $ComputerName" -Category ConnectionError -TargetObject $ComputerName `
							-RecommendedAction 'Check the host is online'
				Break
			}
			
			Write-Warning "Starting WinRM service on $ComputerName"
			& "$env:windir\system32\sc.exe" \\$ComputerName Start WinRM | Out-Null
			
			If ($LASTEXITCODE -eq '1722')
			{
				Write-Error 'Error: The RPC service is unavailable'
				Write-Warning 'Some information cannot be retrieved'
			}
			Else
			{
				$WinRM = 'Enabled' # This variable is set only if the WinRM service starts. WinRM will be disabled at the end.
				Write-Host "Generating remote system report... Please wait... " -ForegroundColor Green
			}
		}
		Else
		{
			Write-Host "Generating local system report... Please wait... " -ForegroundColor Green
		}
		
		# Updates the progress bar for each Win32 object collected
		Function Update-Progress
		{
			param
			(
				[Parameter(Mandatory = $true)]
				[string]$Activity
			)
			Write-Progress -Activity $Activity -PercentComplete (($i/17) * 100)
		}
		
		# Function for retrieving last logged on sam user from remote registry
		Function Get-LastLoggedOnSamUser
		{
			Param
			(
				[parameter(Mandatory = $True)]
				[string]$ComputerName
			)
			If ($ComputerName -eq 'localhost') { $ComputerName = $ENV:COMPUTERNAME }
			$HKLM = [Microsoft.Win32.RegistryHive]::LocalMachine
			$RegPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI"
			$RemoteBaseKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($HKLM, $ComputerName)
			$RegKey = $RemoteBaseKey.OpenSubKey($RegPath)
			($RegKey.GetValue('LastLoggedOnSAMUser'))
		}
		
		# Array of WMI classes to retrieve data from
		$i = 0
		[array]$WMI_Classes =
		"Win32_NetworkLoginProfile",
		"win32_ComputerSystem",
		"win32_bios",
		"win32_battery",
		"win32_Physicalmemory",
		"win32_DiskDrive",
		"win32_DesktopMonitor",
		"win32_networkadapter",
		"win32_operatingsystem",
		"win32_logicalDisk",
		"win32_startupCommand",
		"win32_process",
		"Win32_Processor",
		"win32_Service",
		"Win32_Environment",
		"Win32_BaseBoard",
		"Win32_VideoController"
		
		# Loop through each WMI class in array WMI_Classes, creates a variable for each class and retrieves all class data
		Foreach ($Class in $WMI_Classes)
		{
			Update-Progress -Activity "Gathering $Class information"
			New-Variable -Name $Class -Value (Get-WmiObject -Query "SELECT * FROM $Class" -ComputerName $ComputerName)
			$i++
		}
		
		$UserProfiles = @()
		Foreach ($Username in ($Win32_NetworkLoginProfile).Where({ $_.UserType -eq 'Normal Account' }).Caption) { [string]$UserProfiles += "$Username`n" }
		
		# Custom PSTable containing system information
		$Table = $null
		$Table += New-Object -TypeName PSObject -Property @{
			Computername  = $win32_ComputerSystem.PSComputerName
			Domain	      = $win32_ComputerSystem.Domain
			Model		  = $win32_ComputerSystem.Model
			'SOE Build'   = ($Win32_Environment).Where({ $_.Name -eq 'Build' }).VariableValue
			'SOE Version' = ($Win32_Environment).Where({ $_.Name -eq 'SOE' }).VariableValue
			'Operating System' = $win32_operatingsystem.Caption
			'Install Date' = [Management.ManagementDateTimeConverter]::ToDateTime($win32_operatingsystem.InstallDate)
			'OS Architecture' = $win32_operatingsystem.OSArchitecture
			'Last User'   = (Get-LastLoggedOnSamUser -ComputerName $ComputerName)
			'Logged In Now' = $win32_ComputerSystem.Username
			'User Profiles' = $UserProfiles
			BIOS		  = $win32_bios.Manufacturer
			Version	      = $win32_bios.Version + ' ' + $win32_bios.SMBIOSBIOSVersion
			Serial	      = $win32_bios.SerialNumber
			CPU		      = $Win32_Processor.Name
			Manufacturer  = $Win32_Processor.Manufacturer
			'CPU Clock'   = $Win32_Processor.MaxClockSpeed
			Family	      = $Win32_Processor.Caption
			Memory	      = [string]([math]::Round(($win32_ComputerSystem.TotalPhysicalMemory/1GB), 2)) + ' ' + 'GB'
			'RAM Clock'   = [string]$win32_Physicalmemory.speed[0] + ' ' + 'MHz'
			'Num. of Modules' = $win32_Physicalmemory.Count
			'Motherboard Vendor' = $Win32_BaseBoard.Manufacturer
			MB_Model	  = $Win32_BaseBoard.Product
			MB_Serial	  = $Win32_BaseBoard.SerialNumber
			'Monitor(s)'  = $win32_DesktopMonitor.Description
			'Video Controller' = $Win32_VideoController.VideoProcessor
		}
		
		If ($HotFixes -eq $True)
		{
			Foreach ($Item in (Get-HotFix).HotFixID) { [string]$HotFixes += "$Item`n" }
			$Table += New-Object -TypeName PSObject -Property @{ HotFixes = $HotFixes }
		}
		
		# Output system information to either csv, html or to the console depending on which parameter is selected
		If ($ReportType -eq 'csv') # Outputs system information to 
		{
			If ((Test-Path $SavePath) -eq $True) { Remove-Item $SavePath -Confirm }
			Export-Csv -InputObject $Table -Path $SavePath -NoClobber -NoTypeInformation -Force | Out-Null
			Invoke-Item $SavePath
			Write-Host "The system report has been saved to $SavePath" -ForegroundColor Green
		}
		Elseif ($ReportType -eq 'HTML')
		{
			If ((Test-Path $SavePath) -eq $True) { Remove-Item $SavePath -Confirm }
			# Style sheet for HTML report
			$head = @"
        <style>
        h1, h5, th { text-align: center; }
        table { margin: auto; font-family: Segoe UI; box-shadow: 10px 10px 5px #888; border: thin ridge grey; }
        th { background: #0046c3; color: #fff; max-width: 400px; padding: 5px 10px; }
        td { font-size: 15px; padding: 5px 20px; color: #000; }
        tr { background: #b8d1f3; }
        tr:nth-child(even) { background: #dae5f4; }
        tr:nth-child(odd) { background: #b8d1f3; }
        </style>
"@
			[string]$HTML_Table = $Table | ConvertTo-Html -Fragment -As List -PreContent "<center><h2>Generated $(Get-Date)</h2></center>"
			
			ConvertTo-Html -Title "$ComputerName System Report" `
						   -head $head `
						   -PreContent $HTML_Table `
						   -Body "<h1>$ComputerName Report</h1>" |
			Out-File -FilePath $SavePath
			Invoke-Item $SavePath
			Write-Host "The system report has been saved to $SavePath"
		}
		Else
		{
			$Table | Select-Object Computername, Domain, Model, 'SOE Build', 'SOE Version', 'Operating System', 'Install Date', 'OS Architecture', 'Logged In Now', 'Last User', 'User Profiles', `
								   BIOS, Version, Serial, CPU, Manufacturer, 'CPU Clock', Family, Memory, 'RAM Clock', 'Num. of Modules', 'Motherboard Vendor', MB_Model, MB_Serial, 'Monitor(s)', 'Video Controller'
		}
	} # END
	
    # This script was written by Jon Gurgul
    # https://gallery.technet.microsoft.com/scriptcenter/519e1d3a-6318-4e3d-b507-692e962c6666
	Function Get-InstalledSoftware
	{
		Param ([String[]]$Computers)
		If (!$Computers) { $Computers = $ENV:ComputerName }
		$Base = New-Object PSObject;
		$Base | Add-Member Noteproperty ComputerName -Value $Null;
		$Base | Add-Member Noteproperty Name -Value $Null;
		$Base | Add-Member Noteproperty Publisher -Value $Null;
		$Base | Add-Member Noteproperty InstallDate -Value $Null;
		$Base | Add-Member Noteproperty EstimatedSize -Value $Null;
		$Base | Add-Member Noteproperty Version -Value $Null;
		$Base | Add-Member Noteproperty Wow6432Node -Value $Null;
		$Results = New-Object System.Collections.Generic.List[System.Object];
		
		ForEach ($ComputerName in $Computers)
		{
			$Registry = $Null;
			Try { $Registry = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $ComputerName); }
			Catch { Write-Host -ForegroundColor Red "$($_.Exception.Message)"; }
			
			If ($Registry)
			{
				$UninstallKeys = $Null;
				$SubKey = $Null;
				$UninstallKeys = $Registry.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Uninstall", $False);
				$UninstallKeys.GetSubKeyNames() | %{
					$SubKey = $UninstallKeys.OpenSubKey($_, $False);
					$DisplayName = $SubKey.GetValue("DisplayName");
					If ($DisplayName.Length -gt 0)
					{
						$Entry = $Base | Select-Object *
						$Entry.ComputerName = $ComputerName;
						$Entry.Name = $DisplayName.Trim();
						$Entry.Publisher = $SubKey.GetValue("Publisher");
						[ref]$ParsedInstallDate = Get-Date
						If ([DateTime]::TryParseExact($SubKey.GetValue("InstallDate"), "yyyyMMdd", $Null, [System.Globalization.DateTimeStyles]::None, $ParsedInstallDate))
						{
							$Entry.InstallDate = $ParsedInstallDate.Value
						}
						$Entry.EstimatedSize = [Math]::Round($SubKey.GetValue("EstimatedSize")/1KB, 1);
						$Entry.Version = $SubKey.GetValue("DisplayVersion");
						[Void]$Results.Add($Entry);
					}
				}
				
				If ([IntPtr]::Size -eq 8)
				{
					$UninstallKeysWow6432Node = $Null;
					$SubKeyWow6432Node = $Null;
					$UninstallKeysWow6432Node = $Registry.OpenSubKey("Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall", $False);
					If ($UninstallKeysWow6432Node)
					{
						$UninstallKeysWow6432Node.GetSubKeyNames() | %{
							$SubKeyWow6432Node = $UninstallKeysWow6432Node.OpenSubKey($_, $False);
							$DisplayName = $SubKeyWow6432Node.GetValue("DisplayName");
							If ($DisplayName.Length -gt 0)
							{
								$Entry = $Base | Select-Object *
								$Entry.ComputerName = $ComputerName;
								$Entry.Name = $DisplayName.Trim();
								$Entry.Publisher = $SubKeyWow6432Node.GetValue("Publisher");
								[ref]$ParsedInstallDate = Get-Date
								If ([DateTime]::TryParseExact($SubKeyWow6432Node.GetValue("InstallDate"), "yyyyMMdd", $Null, [System.Globalization.DateTimeStyles]::None, $ParsedInstallDate))
								{
									$Entry.InstallDate = $ParsedInstallDate.Value
								}
								$Entry.EstimatedSize = [Math]::Round($SubKeyWow6432Node.GetValue("EstimatedSize")/1KB, 1);
								$Entry.Version = $SubKeyWow6432Node.GetValue("DisplayVersion");
								$Entry.Wow6432Node = $True;
								[Void]$Results.Add($Entry);
							}
						}
					}
				}
			}
		}
		$Results
	}
	
	Function Test-Credential
	{
		
		param
		(
			[parameter(Mandatory = $true)]
			[string]$Username
		)
		
		Try
		{
			[void][System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
		}
		Catch
		{
			Write-Error -Exception 'Not a domain member' -Message "$env:COMPUTERNAME is not currently a domain member"
			return
		}
		
		$password = Read-Host -AsSecureString 'Password'
		$BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($password)
		$PasswdString = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
		
		if ((New-Object System.DirectoryServices.DirectoryEntry('', $Username, $PasswdString)).psbase.Name -ne $null)
		{
			return Write-Host 'Correct' -ForegroundColor Green
		}
		else
		{
			return Write-Host 'Incorrect' -ForegroundColor Red
		}
		
	}
	
	Function Get-BitLockerRecoveryKey
	{
		param
		(
			[string]$Computername
		)
		
		$DomainsAdminsDn = (Get-ADGroup 'Admin Security Group' -Properties Members).DistinguishedName
		$usersAdm = Get-ADUser -Filter { (memberof -eq $DomainsAdminsDn) }
		if ($usersAdm.samAccountName -contains $env:username)
		{
			$Results = (Get-ADObject -Properties msFVE-RecoveryPassword -LdapFilter '(objectcategory=msFVE-RecoveryInformation)' |
				?{ $_.DistinguishedName -like "*$Computername*" } |
				Sort-Object Name -Descending)
			
			if ($Results)
			{
				return $Results[0].'msFVE-RecoveryPassword'
			}
			else
			{
				Write-Host "$Computername does not contain any recovery information. It may not be a Bitlocker managed device." -ForegroundColor Red
			}
		}
		else
		{
			Write-Host "$env:username does not have appropriate domain access." -ForegroundColor Red
		}
		
	}
	
}



