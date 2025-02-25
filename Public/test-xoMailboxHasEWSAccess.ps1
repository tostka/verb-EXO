# test-xoMailboxHasEWSAccess.ps1
#*------v Function test-xoMailboxHasEWSAccess v------
function test-xoMailboxHasEWSAccess {
    <#
    .SYNOPSIS
    test-xoMailboxHasEWSAccess .ps1 - Test mailbox against MS's o365 updated Org v CASMailbox EWSEnabled specification, under revised Feb 2025 logic.
    .NOTES
    Version     : 0.0.1
    Author      : Todd Kadrie
    Website     : http://www.toddomation.com
    Twitter     : @tostka / http://twitter.com/tostka
    CreatedDate : 2024-
    FileName    : test-xoMailboxHasEWSAccess .ps1
    License     : MIT License
    Copyright   : (c) 2024 Todd Kadrie
    Github      : https://github.com/tostka/verb-XXX
    Tags        : Powershell
    AddedCredit : REFERENCE
    AddedWebsite: URL
    AddedTwitter: URL
    REVISIONS
    * 9:43 AM 2/25/2025 init
    .DESCRIPTION
    test-xoMailboxHasEWSAccess .ps1 - Test mailbox against MS's o365 updated Org v CASMailbox EWSEnabled specification, under revised Feb 2025 logic.
    Encapsulates following logic table from linked article:

    > The_Exchange_Team
    > Icon for Microsoft rank
    > Microsoft
    > Feb 20, 2025
    > ...
    > Current Behavior
    > Organization Level | User Level      | EWS Requests
    > ------------------ | --------------- | ------------
    > True or <null>     | True or <null>  | Allowed
    > True or <null>     | False           | Not Allowed
    > False              | True            | Allowed
    > False              | False or <null> | Not Allowed
    >
    > New Behavior 
    > To address these issues, we are altering the behavior so that EWS will only be allowed if both the organization-level and user-level EWSEnabled flags are true. Here's a simplified view of the new logic:
    > 
    > Organization Level | User Level     | EWS Requests
    > ------------------ | -------------- | ------------
    > True or <null>     | True or <null> | Allowed
    > True or <null>     | False          | Not Allowed
    > False              | True or <null> | Not Allowed
    > False              | False          | Not Allowed
    > 
    > In short, EWS will be permitted only if both the organization and user-level allow it. This change ensures that administrators have better control over EWS access and can enforce policies more consistently across their entire organization
    > 
    .PARAMETER mailbox
    Mailbox identifier
    .INPUTS
    None. Does not accepted piped input.
    .EXAMPLE 
    test-xoMailboxHasEWSAccess lynctest14@toro.com
    Simple mailbox test
    .LINK
    https://techcommunity.microsoft.com/blog/Exchange/the-way-to-control-ews-usage-in-exchange-online-is-changing/4383083
    #>    
    PARAM([Parameter(Position=0,Mandatory=$True)]$mailbox) ; 
    $orgEWSEnable =  (Get-xoOrganizationConfig).ewsenabled ; 
    $usrEWSEnable = (Get-xoCASMailbox -id $mailbox).ewsenabled ; 
    if(($null -eq $orgEWSEnable -OR $true -eq $orgEWSEnable) -AND ($null -eq $usrEWSEnable -OR $true -eq $usrEWSEnable)){write-host "Org:'$($orgEWSEnable)'&Usr:'$($usrEWSEnable)': Mailbox has EWSEnable function"} ; 
    if(($null -eq $orgEWSEnable -OR $true -eq $orgEWSEnable) -AND ($false -eq $usrEWSEnable)){write-host "Org:'$($orgEWSEnable)':Usr:'$($usrEWSEnable)': Mailbox HAS NO EWSEnable function"} ; 
    if(($false -eq $orgEWSEnable) -AND ($null -eq $usrEWSEnable -OR $true -eq $usrEWSEnable)){write-host "Org:'$($orgEWSEnable)':Usr:'$($usrEWSEnable)': Mailbox HAS NO EWSEnable function"} ; 
    if(($false -eq $orgEWSEnable) -AND ($false -eq $usrEWSEnable)){write-host "Org:'$($orgEWSEnable)':Usr:'$($usrEWSEnable)': Mailbox HAS NO EWSEnable function"} ; 
} ; 
#*------^ END Function test-xoMailboxHasEWSAccess ^------