#*------v rxo2tor.ps1 v------
function rxo2TOR {
    <#
    .SYNOPSIS
    rxo2TOR - Reonnect-EXO to specified Tenant
    REVISIONS   :
    * 11:32 AM 4/18/2023 alias into rxoTOR, retiring
    .DESCRIPTION
    Reconnect-EXO2 - Re-establish PSS to EXO V2 Modern Auth
    .EXAMPLE
    rxo2TOR
    #>
    Reconnect-EXO2 -cred $credO365TORSID -Verbose:($VerbosePreference -eq 'Continue') ; 
}
#*------^ rxo2tor.ps1 ^------
