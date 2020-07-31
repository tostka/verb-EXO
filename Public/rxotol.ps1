#*------v rxotol.ps1 v------
function rxoTOL {
    <#
    .SYNOPSIS
    rxoTOL - Reonnect-EXO to specified Tenant
    .DESCRIPTION
    Reconnect-EXO - Re-establish PSS to https://ps.outlook.com/powershell/
    #>
    Reconnect-EXO -cred $credO365TOLSID
}
#*------^ rxotol.ps1 ^------