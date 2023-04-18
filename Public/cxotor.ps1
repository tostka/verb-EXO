#*------v cxoTOR.ps1 v------
function cxoTOR {
    <#
    .SYNOPSIS
    cxoTOR - wrapper for Connect-EXO to connect to specified Tenant
    .NOTES
    Version     : 1.0.0
    Author      : Todd Kadrie
    Website     :	http://www.toddomation.com
    Twitter     :	@tostka / http://twitter.com/tostka
    CreatedDate : 2020-07-29
    FileName    : cxoTOR.ps1
    License     :
    Copyright   :
    Github      : https://github.com/tostka/verb-exo
    Tags        : Powershell,ExchangeOnline,Exchange,RemotePowershell,Connection,MFA
    REVISIONS   :
    * 11:32 AM 4/18/2023 alias into cxoTOR, retiring
    .DESCRIPTION
    cxoTOR - wrapper for Connect-EXO to connect to specified Tenant
    .EXAMPLE
    cxoTOR
    #>
    [CmdletBinding()]
    [Alias('cxo2cmw' )]
    PARAM()
    Connect-EXO -cred $credO365TORSID -Verbose:($VerbosePreference -eq 'Continue') ; 
}
#*------^ cxoTOR.ps1 ^------