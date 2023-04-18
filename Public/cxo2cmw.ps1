#*------v cxo2cmw.ps1 v------
function cxo2cmw {
    <#
    .SYNOPSIS
    cxo2CMW - Connect-EXO to specified Tenant
    .NOTES
    REVISIONS
    * 11:32 AM 4/18/2023 alias into cxocmw, retiring
    * 10:16 AM 7/20/2021 reverted old typo (missing '[exo]2' in call)
    .DESCRIPTION
    Connect-EXO2 - Establish PSS to EXO V2 Modern Auth PS
    .EXAMPLE
    cxo2cmw
    #>
    Connect-EXO2 -cred $credO365CMWCSID -Verbose:($VerbosePreference -eq 'Continue') ; 
}

#*------^ cxo2cmw.ps1 ^------
