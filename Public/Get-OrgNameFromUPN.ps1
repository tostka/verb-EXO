#*------v Function Get-OrgNameFromUPN v------
function Get-OrgNameFromUPN{
    <#
    .SYNOPSIS
    Get-OrgNameFromUPN.ps1 - Extract organization name from UserPrincipalName ; localized verb-EXO vers of non-'$global:' helper funct from ExchangeOnlineManagement. The globals export fine, these don't and appear to need to be loaded manually
    .NOTES
    Version     : 1.0.0
    Author      : Todd Kadrie
    Website     :	http://www.toddomation.com
    Twitter     :	@tostka / http://twitter.com/tostka
    CreatedDate : 20201109-0833AM
    FileName    : Get-OrgNameFromUPN.ps1
    License     : [none specified]
    Copyright   : [none specified]
    Github      : https://github.com/tostka/verb-exo
    Tags        : Powershell
    AddedCredit : Microsoft (edited version of published commands in the module)
    AddedWebsite:	https://docs.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2
    REVISIONS
    * 8:34 AM 11/9/2020 init
    .DESCRIPTION
    Get-OrgNameFromUPN.ps1 - Extract organization name from UserPrincipalName ; localized verb-EXO vers of non-'$global:' helper funct from ExchangeOnlineManagement. The globals export fine, these don't and appear to need to be loaded manually

    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    None. Returns no objects or output.
    .EXAMPLE
    Get-OrgNameFromUPN
    Stock call
    .LINK
    https://github.com/tostka/verb-EXO
    .LINK
    https://docs.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2
    #>
    [CmdletBinding()]
    param([string] $UPN)
    $fields = $UPN -split '@'
    return $fields[-1]
} 
#*------^ END Function Get-OrgNameFromUPN ^------
