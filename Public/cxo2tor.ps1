#*------v cxo2tor.ps1 v------
function cxo2TOR {
    <#
    .SYNOPSIS
    cxo2TOR - Connect-EXO to specified Tenant
    .NOTES
    REVISIONS
    * 10:16 AM 7/20/2021 reverted old typo (missing '[exo]2' in call)
    .DESCRIPTION
    Connect-EXO2 - Establish PSS to EXO V2 Modern Auth PS
    .EXAMPLE
    cxo2TOR
    #>
    Connect-EXO2 -cred $credO365TORSID -Verbose:($VerbosePreference -eq 'Continue') ; 
}
#*------^ cxo2tor.ps1 ^------
