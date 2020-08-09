#*------v cxotol.ps1 v------
function cxoTOL {
    <#
    .SYNOPSIS
    cxoTOL - Connect-EXO to specified Tenant
    .DESCRIPTION
    Connect-EXO - Establish PSS to https://ps.outlook.com/powershell/
    .EXAMPLE
    cxoTOL
    #>
    Connect-EXO -cred $credO365TOLSID -Verbose:($VerbosePreference -eq 'Continue') ; 
}
#*------^ cxotol.ps1 ^------