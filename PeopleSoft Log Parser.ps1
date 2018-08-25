#region############### Script parameters, modify as needed #################
        $server = "ServerName"
        $log = "Assist02"
        $LogPath ='\\ServerName\d$\psoft\PT8.52\webserv\ASSIST02\servers\PIA\logs\ASSIST02.log'
        $EventID = 1
        $EventSource = 'AssistStuck'
        $TraceLog = 'C:\Logs\AssistStuck02_005.log'
        $Date = Get-Date
        $EmailRecipients = "","",""
############################################################################
#endregion


# Checks whether the event log exists, if not then it creates an Application event log for AssistStuck
if (![System.Diagnostics.EventLog]::SourceExists("AssistStuck") )
{ New-EventLog -LogName Application -Source "AssistStuck" | Out-Null}

# Checks if the tracelog file exists, if not it will be created
# UPDATED: the script now generates a tracelog file with detailed errors that might occur while processing Assist logs
elseif (![System.IO.File]::Exists($TraceLog))
{ New-Item -Path $TraceLog -Force -ItemType file | Out-Null }
Add-Content -Value "#############################  $Date  #############################" -Path $TraceLog
Add-Content -Value "Starting error check script on $server for $log" -Path $TraceLog

# Function to parse the Assist log for ALL errors and creates a table with relevant data
# UPDATED: this function now supports terminating errors for the purposes of more detailed error logging
function Parse-AssistLog 
{
    [CmdletBinding()]
    param($LogPath)

    [array]$ErrorObj =@()
    [string]$logContent = Get-Content $LogPath
    [array]$Matches = ($logContent.split('####'))|?({$_ -notlike ''})|%{Select-String -InputObject $_ -SimpleMatch '[STUCK]'}
    
    foreach ($Err in $matches)
    {

        $D = ($Err.ToString().Split('<''>')[1]).Split('G')[0].Split('/')
        $NewDate = [datetime]"$($D[1])/$($D[0])/$($D[2])"

        $table =[ordered]@{
    
            Date    = $NewDate
            StuckAt = ($Err.ToString().Split('<[').Split('/')[17].Split(']')[1].Trim())
            Queue   = ($Err.ToString().Split('<''>')[14])
    
        }    

        $ErrorObj += New-Object psobject -Property $table
       
    }

    return $ErrorObj
}

# Calls the Parse-Assist function and stores ONLY [STUCK] errors occurring within last 5 minutes in the RecentErrors array
# UPDATED: if an error occurs while parsing logs the script will stop and log the error in detail then send an email alert to David Nauendorf and Quentin Forrest
Add-Content -Value "Parsing log file" -Path $TraceLog
try
{
    $Logs = Parse-AssistLog -LogPath $LogPath -ErrorAction Stop
    $RecentErrors=@()
    foreach ($item in $Logs)
    { 
        if ($item.Date -gt (Get-Date).AddMinutes(-5) ) 
        { 
            [array]$RecentErrors+=$item 
            Add-Content -Value "Some errors were found in the Assist log file" -Path $TraceLog
            # Write event for SCOM
            Add-Content -Value "An event has been written to the Application log" -Path $TraceLog
        } 
    }
}
catch
{
    Add-Content -Value "An error occurred while parsing the log file" -Path $TraceLog
    Add-Content -Value "$($_.Exception.Message)" -Path $TraceLog
    Send-MailMessage -Body $($_.Exception.Message) -From 'AssistAlert@cpfs.wa.gov.au' -SmtpServer 'webmail.cpfs.wa.gov.au' -Subject 'Error processing Assist logs' -To "David.Nauendorf@cpfs.wa.gov.au"
}


# If at least one [STUCK] error is found write an event to the SCOM monitored server and send the email alert
if ($RecentErrors.Count -gt 0)
{
    Write-EventLog -LogName Application -EntryType Error -Source $EventSource -EventId $EventID -Message "An Assist [STUCK] error occurred at $($RecentErrors.Date) on $server for log file $log"
    $css = @"
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

    $HTMLDetails = @{ 
        Title = 'Assist is [STUCK]'
        Head = $CSS 
        Body = "<h1>Assist is [STUCK]!</h1> $server $logPath"
    } 
     
    $args = @{ 
        To         = $EmailRecipients
        Body       = "$( $RecentErrors | ConvertTo-Html @HTMLDetails)" 
        Subject    = 'NEW - Assist is [STUCK] - ' + $Server + " " + $log
        SmtpServer = 'webmail.cpfs.wa.gov.au'
        From       = 'AssistAlert@cpfs.wa.gov.au'
        BodyAsHtml = $True 
    } 

    Send-MailMessage @args 
    Add-Content -Value "An alert email has been sent" -Path $TraceLog
}

Add-Content -Value "-End Script" -Path $TraceLog