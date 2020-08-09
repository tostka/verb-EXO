#*------v cxoVEN.ps1 v------
function cxoVEN {
    <#
    .SYNOPSIS
    cxoVEN - Connect-EXO to specified Tenant
    .DESCRIPTION
    Connect-EXO - Establish PSS to https://ps.outlook.com/powershell/
    .EXAMPLE
    cxoVEN
    #>
    Connect-EXO -cred $credO365VENCSID -Verbose:($VerbosePreference -eq 'Continue') ; 
}
#*------^ cxoVEN.ps1 ^------