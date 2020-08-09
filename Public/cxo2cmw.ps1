#*------v cxo2cmw.ps1 v------
function cxo2cmw {
    <#
    .SYNOPSIS
    cxo2CMW - Connect-EXO to specified Tenant
    .DESCRIPTION
    Connect-EXO2 - Establish PSS to EXO V2 Modern Auth PS
    .EXAMPLE
    cxo2cmw
    #>
    Connect-EXO -cred $credO365CMWCSID-Verbose:($VerbosePreference -eq 'Continue') ; 
}
#*------^ cxo2cmw.ps1 ^------
