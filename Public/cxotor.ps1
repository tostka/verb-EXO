#*------v cxotor.ps1 v------
function cxoTOR {
    <#
    .SYNOPSIS
    cxoTOR - Connect-EXO to specified Tenant
    .DESCRIPTION
    Connect-EXO - Establish PSS to https://ps.outlook.com/powershell/
    #>
    Connect-EXO -cred $credO365TORSID
}
#*------^ cxotor.ps1 ^------