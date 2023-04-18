#*------v rxoVEN.ps1 v------
function rxoVEN {
    <#
    .SYNOPSIS
    rxoVEN - wrapper for Connect-EXO to connect to specified Tenant
    .NOTES
    Version     : 1.0.0
    Author      : Todd Kadrie
    Website     :	http://www.toddomation.com
    Twitter     :	@tostka / http://twitter.com/tostka
    CreatedDate : 2020-07-29
    FileName    : rxoVEN.ps1
    License     :
    Copyright   :
    Github      : https://github.com/tostka/verb-exo
    Tags        : Powershell,ExchangeOnline,Exchange,RemotePowershell,Connection,MFA
    REVISIONS   :
    * 11:32 AM 4/18/2023 alias into rxoVEN, retiring
    .DESCRIPTION
    rxoVEN - wrapper for Connect-EXO to connect to specified Tenant
    .EXAMPLE
    rxoVEN
    #>
    [CmdletBinding()]
    [Alias('cxo2VEN' )]
    PARAM()
    ReConnect-EXO -cred $credO365VENCSID -Verbose:$($VerbosePreference -eq 'Continue') ; 
}
#*------^ rxoVEN.ps1 ^------