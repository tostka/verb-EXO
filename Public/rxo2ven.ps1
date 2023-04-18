#*------v rxo2VEN.ps1 v------
function rxo2VEN {
    <#
    .SYNOPSIS
    rxo2VEN - Reonnect-EXO to specified Tenant
    REVISIONS   :
    * 11:32 AM 4/18/2023 alias into rxoVEN, retiring
    .DESCRIPTION
    Reconnect-EXO2 - Re-establish PSS to EXO V2 Modern Auth
    #>
    Reconnect-EXO2 -cred $credO365VENCSID -Verbose:($VerbosePreference -eq 'Continue') ; 
}
#*------^ rxo2VEN.ps1 ^------
