#*------v cxoVEN.ps1 v------
function cxoVEN {
    <#
    .SYNOPSIS
    cxoVEN - wrapper for Connect-EXO to connect to specified Tenant
    .NOTES
    Version     : 1.0.0
    Author      : Todd Kadrie
    Website     :	http://www.toddomation.com
    Twitter     :	@tostka / http://twitter.com/tostka
    CreatedDate : 2020-07-29
    FileName    : cxoVEN.ps1
    License     :
    Copyright   :
    Github      : https://github.com/tostka/verb-exo
    Tags        : Powershell,ExchangeOnline,Exchange,RemotePowershell,Connection,MFA
    REVISIONS   :
    * 11:32 AM 4/18/2023 alias into cxoVEN, retiring
    .DESCRIPTION
    cxoVEN - wrapper for Connect-EXO to connect to specified Tenant
    .EXAMPLE
    cxoVEN
    #>
    [CmdletBinding()]
    [Alias('cxo2cmw' )]
    PARAM()
    Connect-EXO -cred $credO365VENSID -Verbose:($VerbosePreference -eq 'Continue') ; 
}
#*------^ cxoVEN.ps1 ^------