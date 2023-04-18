#*------v rxoTOL.ps1 v------
function rxoTOL {
    <#
    .SYNOPSIS
    rxoTOL - wrapper for Connect-EXO to connect to specified Tenant
    .NOTES
    Version     : 1.0.0
    Author      : Todd Kadrie
    Website     :	http://www.toddomation.com
    Twitter     :	@tostka / http://twitter.com/tostka
    CreatedDate : 2020-07-29
    FileName    : rxoTOL.ps1
    License     :
    Copyright   :
    Github      : https://github.com/tostka/verb-exo
    Tags        : Powershell,ExchangeOnline,Exchange,RemotePowershell,Connection,MFA
    REVISIONS   :
    * 11:32 AM 4/18/2023 alias into rxoTOL, retiring
    .DESCRIPTION
    rxoTOL - wrapper for Connect-EXO to connect to specified Tenant
    .EXAMPLE
    rxoTOL
    #>
    [CmdletBinding()]
    [Alias('cxo2TOL' )]
    PARAM()
    ReConnect-EXO -cred $credO365TOLCSID -Verbose:$($VerbosePreference -eq 'Continue') ; 
}
#*------^ rxoTOL.ps1 ^------