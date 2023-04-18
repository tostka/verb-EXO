#*------v rxo2tol.ps1 v------
function rxo2TOL {
    <#
    .SYNOPSIS
    rxo2TOL - Reonnect-EXO to specified Tenant
    REVISIONS   :
    * 11:32 AM 4/18/2023 alias into rxoTOL, retiring
    .DESCRIPTION
    #>
    Reconnect-EXO2 -cred $credO365TOLSID -Verbose:($VerbosePreference -eq 'Continue')
}
#*------^ rxo2tol.ps1 ^------
