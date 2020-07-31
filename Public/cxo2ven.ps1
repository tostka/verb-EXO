#*------v cxo2VEN.ps1 v------
function cxo2VEN {
    <#
    .SYNOPSIS
    cxo2VEN - Connect-EXO to specified Tenant
    .DESCRIPTION
    Connect-EXO2 - Establish PSS to EXO V2 Modern Auth PS
    #>
    Connect-EXO -cred $credO365VENCSID
}
#*------^ cxo2VEN.ps1 ^------