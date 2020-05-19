Function Convert-Base64  {
<#
.Synopsis
   Convert a file to Base64 || Write a string of Base64 to a file

.DESCRIPTION
   This cmdlet will convert the bytes for a given file into Base64 encoding that can be stored in a plain text file or an array. 
   The string of Base64 can then be written back to a file in its original form.
   This can be used for storing images in scripts to style graphical interfaces or even storing small executables in scripts for developing advanced PowerShell tools.

   By default output to the console is suppressed and only copied to the clip board.

.NOTES

    Author  : David Nauendorf
    Created : 2017

.EXAMPLE

   $base64 = Convert-Base64 -ToBase64 -FromFilePath C:\temp\125652.jpg

   $base64 | Convert-Base64 -FromBase64 -ToFilePath c:\temp\newimage.jpg

   Convert-Base64 -FromBase64 -ToFilePath c:\temp\newimage.jpg -Base64String $base64

.EXAMPLE

   Convert-Base64 -ToBase64 -FromFilePath C:\temp\125652.jpg -SuppressOutput:$false 

.LINK

    https://en.wikipedia.org/wiki/Base64

#>
Param
(

    [CmdletBinding()]

    # This parameter switch indicates that a file will be converted to a Base64 string
    [parameter(ParameterSetName='ToBase',   Position=1, Mandatory=$false)]
    [switch]$ToBase64,      
    # Identifies the file that will be converted to Base64
    [parameter(ParameterSetName='ToBase',   Position=2, Mandatory=$true, HelpMessage='I need a file to convert')]
    [ValidateScript({if (!(Test-Path $_ -PathType Leaf) ){throw "The file $_ was not found"} else {return $true} })]
    [string]$FromFilePath,  
    # This parameter switch indicates that a Base64 string will be written to a file
    [parameter(ParameterSetName='FromBase', Position=1, Mandatory=$false)]
    [switch]$FromBase64,    
    # Name of the file being created
    [parameter(ParameterSetName='FromBase', Position=2, Mandatory=$true, HelpMessage='I need a valid folder to write this Base64 to')]
    [ValidateScript({if (!(Test-Path ( Split-Path $_ -Parent) ) ){throw "The folder path $_ was not found"}else {return $true} })]
    [string]$ToFilePath,    
    # This is the string of Base64 that will be written to $ToFilePath
    [parameter(ParameterSetName='FromBase', Position=3, Mandatory=$true, ValueFromPipeline=$true, HelpMessage='I need a valid string of Base64')]
    [string]$Base64String,  
    # Suppresses output of the Base64 string to the console. The content is accessible from the clip board.
    [switch]$SuppressOutput = $true 

)

    Try
    {
        If ($ToBase64)
        {
            Write-Verbose -Message "Generating Base64 from file $FromFilePath"
            $ByteContent = [convert]::ToBase64String((Get-Content $FromFilePath -encoding byte -ErrorAction Stop))
            $ByteContent | clip.exe # This copies your Base64 string directly to the clip board!!!
            if (!$SuppressOutput) {return $ByteContent} else {Write-Verbose -Message "Output to the console has been suppressed"}
        }
        ElseIf ($FromBase64)
        {
            Write-Verbose -Message "Writing Base64 to file $ToFilePath"
            $ByteContent = [Convert]::FromBase64String($Base64String)
            [System.IO.File]::WriteAllBytes($ToFilePath,$ByteContent)
        }
    }

    Catch [System.Management.Automation.MethodInvocationException] {Write-Verbose -Message "Clearly something has gone wrong...." ; $_}

}
