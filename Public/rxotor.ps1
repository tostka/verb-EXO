#*------v rxotor.ps1 v------
function rxoTOR {
    <#
    .SYNOPSIS
    rxoTOR - Reonnect-EXO to specified Tenant
    .DESCRIPTION
    Reconnect-EXO - Re-establish PSS to https://ps.outlook.com/powershell/
    .EXAMPLE
    rxoTOR
    #>
    Reconnect-EXO -cred $credO365TORSID -Verbose:($VerbosePreference -eq 'Continue') ; 
}

#*------^ rxotor.ps1 ^------