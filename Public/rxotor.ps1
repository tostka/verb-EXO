#*------v rxoTOR.ps1 v------
function rxoTOR {
    <#
    .SYNOPSIS
    rxoTOR - wrapper for Connect-EXO to connect to specified Tenant
    .NOTES
    Version     : 1.0.0
    Author      : Todd Kadrie
    Website     :	http://www.toddomation.com
    Twitter     :	@tostka / http://twitter.com/tostka
    CreatedDate : 2020-07-29
    FileName    : rxoTOR.ps1
    License     :
    Copyright   :
    Github      : https://github.com/tostka/verb-exo
    Tags        : Powershell,ExchangeOnline,Exchange,RemotePowershell,Connection,MFA
    REVISIONS   :
    * 11:32 AM 4/18/2023 alias into rxoTOR, retiring
    .DESCRIPTION
    rxoTOR - wrapper for Connect-EXO to connect to specified Tenant
    .EXAMPLE
    rxoTOR
    #>
    [CmdletBinding()]
    [Alias('cxo2TOR' )]
    PARAM()
    ReConnect-EXO -cred $credO365TORCSID -Verbose:$($VerbosePreference -eq 'Continue') ; 
}
#*------^ rxoTOR.ps1 ^------