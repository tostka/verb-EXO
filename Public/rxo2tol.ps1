#*------v rxo2tol.ps1 v------
function rxo2TOL {
    <#
    .SYNOPSIS
    rxo2TOL - Reonnect-EXO to specified Tenant
    .DESCRIPTION
    #>
    Reconnect-EXO2 -cred $credO365TOLSID -Verbose:($VerbosePreference -eq 'Continue')
}
#*------^ rxo2tol.ps1 ^------