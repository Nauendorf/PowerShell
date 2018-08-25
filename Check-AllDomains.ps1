function Check-AllDomains ([string]$Identity )
{
    'LDAP://DC=',
    'LDAP://DC=',
    'LDAP://DC='|
    %{

        $LDAPFilter = "(samAccountName=$Identity)"
        $searcher = New-Object DirectoryServices.DirectorySearcher
        $searcher.Filter = $LDAPFilter
        $Searcher.SearchRoot = $SearchRoot
        [adsi]$userObj = ($searcher.FindAll()).Path
        [array]$result += $userObj
    }

    if     (!$result) {return $false}
    elseif ($result)  {return $true}
}
