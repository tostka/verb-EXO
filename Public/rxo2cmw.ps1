#*------v rxo2cmw.ps1 v------
function rxo2CMW {
    <#
    .SYNOPSIS
    rxo2CMW - Reonnect-EXO2 to specified Tenant
    .DESCRIPTION
    Reconnect-EXO2 - Re-establish PSS to EXO V2 Modern Auth
    .EXAMPLE
    rxo2CMW
    #>
    Reconnect-EXO2 -cred $credO365CMWCSID -Verbose:($VerbosePreference -eq 'Continue') ; 
}
#*------^ rxo2cmw.ps1 ^------