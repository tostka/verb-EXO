#*------v rxocmw.ps1 v------
function rxoCMW {
    <#
    .SYNOPSIS
    rxoCMW - Reonnect-EXO to specified Tenant
    .DESCRIPTION
    Reconnect-EXO - Re-establish PSS to https://ps.outlook.com/powershell/
    #>
    Reconnect-EXO -cred $credO365CMWCSID
}
#*------^ rxocmw.ps1 ^------