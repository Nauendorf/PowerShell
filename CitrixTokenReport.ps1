Function Get-CitrixTokenReport
{

    param
    (
        [parameter(Mandatory=$false)][string]$Savepath, 
        [parameter(ParameterSetName='Name')][string]$Name, 
        [parameter(ParameterSetName='Token')][string]$Token,
        [switch][parameter(ParameterSetName='FullReport')]$FullReport,
        [switch]$NoOutGrid   
    )

    try
    {
        if ($Name)
        {
            [array]$results=@()
            $searcher = ( ([adsisearcher]"(&(securecomputingCom2000-SafeWord-UserID=*)(Name=$Name))").FindAll() ).Path  
            $results += $searcher|%{[adsi]$_}
        }
        elseif ($Token)
        {
            [array]$results=@()
            $searcher = ( ([adsisearcher]"(securecomputingCom2000-SafeWord-UserID=$Token)").FindAll() ).Path  
            $results += $searcher|%{[adsi]$_}
        }
        elseif ($FullReport)
        {
            [array]$results=@()
            $searcher = ( ([adsisearcher]"(securecomputingCom2000-SafeWord-UserID=*)").FindAll() ).Path  
            $results += $searcher|%{[adsi]$_}
        }
    }
    catch
    {
        Write-Error -Message 'An error occurred while processing your search query' -Exception 'Search error'
    }

    if (!$results){return $false}

    $Table=@()
    foreach ($item in $results)
    {

        $CustomObject = New-Object psobject
        
        [string]$Token = $item.'securecomputingcom2000-safeword-userid'
        [string]$Name = $item.Name -replace [string]'}*{*'
        [string]$TelephoneNumber = $item.TelephoneNumber -replace [string]' *}*{*',''
        [string]$Mobile = $item.Mobile -replace [string]'}*{*',''
        [string]$Address = $item.streetAddress -replace [string]'}*{*',''
        [string]$Department = $item.Department -replace [string]'}*{*',''      
        if (!$item.Manager){$Manager='N/A'}
        else { [string]$Manager = $item.Manager.Split('='',')[1] }
        [string]$Mail = $item.Mail -replace [string]'}*{*',''

        $CustomObject | Add-Member -Name "Token" -Value $Token -MemberType NoteProperty -Force
        $CustomObject | Add-Member -Name "Name" -Value $Name -MemberType NoteProperty -Force
        $CustomObject | Add-Member -Name "TelephoneNumber" -Value $TelephoneNumber -MemberType NoteProperty -Force
        $CustomObject | Add-Member -Name "Mobile" -Value $Mobile -MemberType NoteProperty -Force
        $CustomObject | Add-Member -Name "Address" -Value $Address -MemberType NoteProperty -Force
        $CustomObject | Add-Member -Name "Department" -Value $Department -MemberType NoteProperty -Force
        $CustomObject | Add-Member -Name "Manager" -Value $Manager -MemberType NoteProperty -Force
        $CustomObject | Add-Member -Name "Mail" -Value $Mail -MemberType NoteProperty -Force

        [array]$Table += $CustomObject

    }

    if ($Savepath){ $Table|Export-Csv -Path $Savepath -NoClobber -NoTypeInformation -Force }
    if (!$NoOutGrid){ $Table|Out-GridView -Title 'Assigned token report' } else { return $Table }
    
}
