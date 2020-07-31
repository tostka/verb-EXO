#*------v rxotor.ps1 v------
function rxoTOR {
    <#
    .SYNOPSIS
    rxoTOR - Reonnect-EXO to specified Tenant
    .DESCRIPTION
    Reconnect-EXO - Re-establish PSS to https://ps.outlook.com/powershell/
    #>
    Reconnect-EXO -cred $credO365TORSID
}

#*------^ rxotor.ps1 ^------