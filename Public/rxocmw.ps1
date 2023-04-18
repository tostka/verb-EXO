#*------v rxocmw.ps1 v------
function rxocmw {
    <#
    .SYNOPSIS
    rxocmw - wrapper for Connect-EXO to connect to specified Tenant
    .NOTES
    Version     : 1.0.0
    Author      : Todd Kadrie
    Website     :	http://www.toddomation.com
    Twitter     :	@tostka / http://twitter.com/tostka
    CreatedDate : 2020-07-29
    FileName    : rxocmw.ps1
    License     :
    Copyright   :
    Github      : https://github.com/tostka/verb-exo
    Tags        : Powershell,ExchangeOnline,Exchange,RemotePowershell,Connection,MFA
    REVISIONS   :
    * 11:32 AM 4/18/2023 alias into rxocmw, retiring
    .DESCRIPTION
    rxocmw - wrapper for Connect-EXO to connect to specified Tenant
    .EXAMPLE
    rxocmw
    #>
    [CmdletBinding()]
    [Alias('cxo2cmw' )]
    PARAM()
    ReConnect-EXO -cred $credO365CMWCSID -Verbose:$($VerbosePreference -eq 'Continue') ; 
}
#*------^ rxocmw.ps1 ^------