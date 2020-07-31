#*------v cxocmw.ps1 v------
function cxoCMW {
    <#
    .SYNOPSIS
    cxoCMW - Connect-EXO to specified Tenant
    .DESCRIPTION
    Connect-EXO - Establish PSS to https://ps.outlook.com/powershell/
    #>
    Connect-EXO -cred $credO365CMWCSID
}
#*------^ cxocmw.ps1 ^------