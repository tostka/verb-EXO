#*------v Function Print-Details v------
function Print-Details{
    <#
    .SYNOPSIS
    Print-Details.ps1 - localized verb-EXO vers of non-'$global:' helper funct from ExchangeOnlineManagement. The globals export fine, these don't and appear to need to be loaded manually
    .NOTES
    Version     : 1.0.0
    Author      : Todd Kadrie
    Website     :	http://www.toddomation.com
    Twitter     :	@tostka / http://twitter.com/tostka
    CreatedDate : 20201109-0833AM
    FileName    : Print-Details.ps1
    License     : [none specified]
    Copyright   : [none specified]
    Github      : https://github.com/tostka/verb-exo
    Tags        : Powershell
    AddedCredit : Microsoft (edited version of published commands in the module)
    AddedWebsite:	https://docs.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2
    REVISIONS
    * 8:34 AM 11/9/2020 init
    .DESCRIPTION
    Print-Details.ps1 - localized verb-EXO vers of non-'$global:' helper funct from ExchangeOnlineManagement. The globals export fine, these don't and appear to need to be loaded manually
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    None. Returns no objects or output.
    .EXAMPLE
    Print-Details
    Stock call
    .LINK
    https://github.com/tostka/verb-EXO
    .LINK
    https://docs.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2
    #>
    [CmdletBinding()]
    param()
    Write-Host -ForegroundColor Yellow ""
    Write-Host -ForegroundColor Yellow "----------------------------------------------------------------------------"
    Write-Host -ForegroundColor Yellow "The module allows access to all existing remote PowerShell (V1) cmdlets in addition to the 9 new, faster, and more reliable cmdlets."
    Write-Host -ForegroundColor Yellow ""
    Write-Host -ForegroundColor Yellow "|--------------------------------------------------------------------------|"
    Write-Host -ForegroundColor Yellow "|    Old Cmdlets                    |    New/Reliable/Faster Cmdlets       |"
    Write-Host -ForegroundColor Yellow "|--------------------------------------------------------------------------|"
    Write-Host -ForegroundColor Yellow "|    Get-CASMailbox                 |    Get-EXOCASMailbox                 |"
    Write-Host -ForegroundColor Yellow "|    Get-Mailbox                    |    Get-EXOMailbox                    |"
    Write-Host -ForegroundColor Yellow "|    Get-MailboxFolderPermission    |    Get-EXOMailboxFolderPermission    |"
    Write-Host -ForegroundColor Yellow "|    Get-MailboxFolderStatistics    |    Get-EXOMailboxFolderStatistics    |"
    Write-Host -ForegroundColor Yellow "|    Get-MailboxPermission          |    Get-EXOMailboxPermission          |"
    Write-Host -ForegroundColor Yellow "|    Get-MailboxStatistics          |    Get-EXOMailboxStatistics          |"
    Write-Host -ForegroundColor Yellow "|    Get-MobileDeviceStatistics     |    Get-EXOMobileDeviceStatistics     |"
    Write-Host -ForegroundColor Yellow "|    Get-Recipient                  |    Get-EXORecipient                  |"
    Write-Host -ForegroundColor Yellow "|    Get-RecipientPermission        |    Get-EXORecipientPermission        |"
    Write-Host -ForegroundColor Yellow "|--------------------------------------------------------------------------|"
    Write-Host -ForegroundColor Yellow ""
    Write-Host -ForegroundColor Yellow "To get additional information, run: Get-Help Connect-ExchangeOnline or check https://aka.ms/exops-docs"
    Write-Host -ForegroundColor Yellow ""
    Write-Host -ForegroundColor Yellow "Send your product improvement suggestions and feedback to exocmdletpreview@service.microsoft.com. For issues related to the module, contact Microsoft support. Don't use the feedback alias for problems or support issues."
    Write-Host -ForegroundColor Yellow "----------------------------------------------------------------------------"
    Write-Host -ForegroundColor Yellow ""
} 
#*------^ END Function Print-Details ^------
