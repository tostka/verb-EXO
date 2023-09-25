# Reset-xoMailboxFolderPermissionsRecursive.ps1

#*------v Function Reset-xoMailboxFolderPermissionsRecursive v------
function Reset-xoMailboxFolderPermissionsRecursive {
    <#
    .SYNOPSIS
    Reset-xoMailboxFolderPermissionsRecursive - Restores the default permissions for all user-accessible folders for a given mailbox. Can also be used to remove broken recipients. 
   .NOTES
    Version     : 1.0.0
    Author      : Vasil Michev
    Website     : https://www.michev.info/blog/post/2500/how-to-reset-mailbox-folder-permissions
    Twitter     :	
    CreatedDate : 2022-06-15
    FileName    : reset-XOMailboxAllFolderPerms.ps1
    License     : Not Asserted
    Copyright   : Not Asserted
    Github      : https://github.com/michevnew/PowerShell/blob/master/reset-XOMailboxAllFolderPerms.ps1
    Tags        : Powershell,ExchangeOnline,Mailbox,Delegate
    AddedCredit : Todd Kadrie
    AddedWebsite: http://www.toddomation.com
    AddedTwitter: @tostka / http://twitter.com/tostka
    REVISIONS
    * 1:44 PM 9/25/2023 debuged, whatif working; moved includedfolders, excludedfolders into targetable pre-populated params (as overriding code lists in a signed module is a mess; but a param can be done on the fly); 
         strip away outter wrapper script, in favor of descrete (verb-EXO-hosted) reusable functions; added $ThrottleMs fallback; expanded w-v, w-h & w-w into pswlt support
    * 4:43 PM 9/21/2023 works, used on 760151;  add option: We want to leave INTERNAL/EXTERNAL existing, but remove UNKNOWNS: neither is addressed by default below (INT/EXT are set to NONE, and UNKN is *ignored*).
        Add param: -RemoveUnresolveable -> targets usertype:UNKNOWN, including getr-adusere solvable, that lack populated msExchRecipientTypeDetails property
        Add param: -IgnoreInternal - skips reset of existing usertype:Internal to NONE
        Add param: -IgnoreExternal - skips reset of existing usertype:External to NONE
    * 3:35 PM 7/11/2023 works; ADD:$CalendarLimitedDetails param, to drive variant 
    default view (customized in our org), passed in via psboundparameters; 
    completely refactored ExchangeOnlineManagement ineraction to accomodate loss of 
    WinRM/PSSession connections in EOM3+; minor reformatting, added root CBH 
    * 6/15/22 vm posted version
    AddedCredit : REFERENCE
    AddedWebsite: URL
    AddedTwitter: URL
    REVISIONS
    * 12:53 PM 9/22/2023 reflect ren of ReturnFolderList -> get-XOMailboxFolderList; GetFolderPermissions -> get-XOMailboxFolderPermissionsSummary
    * 3:03 PM 7/11/2023 ADD:$CalendarLimitedDetails param, to drive variant default view (customized in our org), passed in via psboundparameters

    .DESCRIPTION
    Reset-xoMailboxFolderPermissionsRecursive - Restores the default permissions for all user-accessible folders for a given mailbox. Can also be used to remove broken recipients. 

    The Reset-xoMailboxFolderPermissionsRecursive cmdlet removes permissions for all user-accessible folders for the given mailbox(es), specified via the -Mailbox parameter. 
    The list of folders is generated via the get-XOMailboxFolderList function. 
    Configure the $includedfolders and $excludedfolders variables to granularly control the folder list.

    Default Actions:
    - wo use of ResetDefaultLevel, Existing Defaults are left unmodified
    - Existing Anonymous grants are also always left unmodified (it uses Continue, and skips any interaction with the grants).
    - If ResetDefaultLevel is used: 
        * This forces Default Calendar grants to AvailabilityOnly (MS Default) (or LimitedDetails, if CalendarLimitedDetails:$true): Set-MailboxPermission
        * and ANY OTHER Default is forced to NONE. 
        -> It strips all but Cal Default access.
    - ALL Internal & External usertype grants are REMOVED completely.

    If -Quiet is *not* used, modifications information is exported to a file in local directory named: 
    yyyy-MM-dd_HH-mm-ss_MailboxFolderPermissionsRemoved.csv (and the same info is returned as an object to pipeline)

    As originally written, this was *selective*, doesn't go after UNKNOWN "orphans" at all. clearly MV, as an MVP knows something that I don't, 
    to choose to leave UNKNOWN's unpurged while purging all existing non-Default/non-Anon Internal|External user-created entries.

    -----------
    ## Unable to find formal docs, taking a stab at _imputing_ the intent of the UserType variants:
    - Default is the generic 'class' grant for users that authenticated with the domain but don't have specific permission.
    - Anonymous is the generic 'class' grant for users that are NOT authenticated with the domain.
    - Internal appears to reflect manually-added (non-Default) grants to objects in the mailbox, that resolve to Recipient Objects
    - External appears to reflect manually-added (non-Default) grants to objects in the mailbox, that resolve to external Recipient Objects
    - Unknown appears (observed):
        - sometimes are named UserName references but consistently lack UserName (resolved RecipientPrincipal.guid.guid). 
        - frequently as broken SID references: NT:S-1-5-21-*
        - both deleted ADUser objects...
             Unable to resolve S-1-5-21-2222296782-158576315-1096482972-3663 to an existing ADUser object (likely deleted TERM)
        ... and ADUser objects that resolve in AD:
            #-=-=-=-=-=-=-=-=
            WARNING: 09:54:03:Resolved
             S-1-5-21-2222296782-158576315-1096482972-39941
             to an existing ADUser object:
            DistinguishedName : CN=FNAME LNAME,OU=Disabled,OU=Users,OU=SITE,DC=SUB,DC=SUB,DC=DOMAIN,DC=com
            enabled           : False
            samaccountname    :  ______x
            sid               : S-1-5-21-2222296782-158576315-1096482972-39941
            UserPrincipalName : FNAME.LNAME.Nelson@DOMAIN.com
            #-=-=-=-=-=-=-=-=
       - opinion: likely resolvable ADU are non-recipients: lacking in msex* attribs on ADU. But are still non-functional: confirm
            [PS]:D:\scripts $ get-aduser -id S-1-5-21-2222296782-158576315-1096482972-39941 -prop * | fl userp*,msex*
            UserPrincipalName            : FNAME.LNAME.Nelson@DOMAIN.com
            msExchALObjectVersion        : 3610
            msExchOmaAdminWirelessEnable : 4
            msExchWhenMailboxCreated     : 2/21/2011 5:24:16 AM
            -> all it's got is the 3 msex's above. Everything else is gone. Prob key test is going to be recipienttypedetails
            #-=-=-=-=-=-=-=-=
            msExchRecipientDisplayType : -2147483642
            msExchRecipientTypeDetails : 2147483648
            msExchRemoteRecipientType  : 4
            msExchSafeRecipientsHash   : {142, 23, 177, 78}
            #-=-=-=-=-=-=-=-=
        => Test get-aduser -prop * | msExchRecipientTypeDetails: if doesn't resolve, purge the NT:S-1-5-21-* entry
        -----------
    And there's a brand new EOM31+ version available today:
    [Managing mailbox folder permissions in bulk in Microsoft 365 - Blog](https://www.michev.info/blog/post/5763/managing-mailbox-folder-permissions-in-bulk-in-microsoft-365)
    # September 20, 2023	Vasil Michev

    --
    [How to reset mailbox folder permissions - Blog](https://www.michev.info/blog/post/2500/how-to-reset-mailbox-folder-permissions)

    what is the best (or at least a proper) way to "reset" folder level 
    permissions, with the added challenge of doing it in bulk

    First of all, if you simply want to "reset" the permissions on a given, 
    "known" folder, the task is easy. Say we have the user JohnSmith and we want to 
    remove any permissions on his Calendar folder

    Next, we need to exclude the "default" permissions entries, as in the 
    ones configured for the Default and Anonymous security principals.  
    There are many additional factors that we need to address, such as the actual 
    folder names, as depending on the localization, the Calendar folder might be 
    renamed to Kalender or whatnot. Then, what if we want to include all folders in 
    the mailbox, not just Calendar? And there are things to consider when removing 
    the permissions as well, such as dealing with orphaned entries, external 
    permissions, published Calendars 

    the building blocks we need to put together: 
    - Account for the type of User, and depending on it handle things accordingly. 
    In other words, for each permission entry, look at the _$entry.User.UserType.Value_.
    Available values will include _Internal_, _External_ and _Unknown_ 
    and all of these will have to be handled differently

    - Utilize the _Get-MailboxFolderStatistics_ cmdlet to get a list of the 
    localized folder names and trim the list to only include folders you care about.
    There's no point in adjusting permissions on Purges folder for example

    - If you are using the above method to get the localized folder names 
    across multiple mailboxes, you need to start to account for throttling! 
    - Decide what you want to do with the Default (and Anonymous) permission level. 
    The regex we used in the above example can be generalized to exclude other 
    entries as well, if needed


    --- Stock setting: 
    Check calendar folder permissions using Get-MailboxFolderPermission user:\calendar 
    and see if Default user has None permissions. Default user should have "AvailabilityOnly"(MS Default) or "LimitedDetails" (TTC)
    --- 
    Relevent discussion on the need/desire to purge broken SID NT:S-1... entries (in public folders in this case: breaks cloud migration).
    https://techcommunity.microsoft.com/legacyfs/online/media/2019/01/FB_Errors.FixesV6.pdf

    ---

    [Correcting Public Folder Permissions before an Office 365 Migration | Practical365](https://practical365.com/correcting-public-folder-permissions-before-an-office-365-migration/)
    Written By Steve Goodman Post published:April 1, 2020
    ...
    > Microsoft's [Source Side Validation 
    script](https://www.microsoft.com/en-us/download/confirmation.aspx?id=100414), 
    described in the blog post above, will generate a log file showing amongst 
    other things, orphaned ACLs you need to remove. However, it doesn't go as far 
    as to assist with the removal itself

    Within the file we'll see lines saying this folder <foldername> permission 
    needs to be removed for NT User:<SID>

    These lines are showing an orphaned ACL. The orphaned ACL occurs when a 
    user is deleted from Active Directory, but the permission is not removed from 
    the Public Folder. This leaves just the security identifier (the SID) showing 
    because it cannot be resolved to an actual user account

    This is a problem because when the folder is migrated to Office 365, the 
    permission cannot be re-applied as the user doesn't exist anymore

    The guidance on the Microsoft blog post doesn't provide you much detail 
    on how to use the log file to remove the permissions, and the guidance it does 
    give doesn't work on Exchange 2010. This is where the 
    Remove-PFPermissionsFromSSV script comes in.   

    The script takes the lines from the log file, and for each line with a 
    permission listed, it uses a command like the one below togGet the folder 
    permissions, find the offending permission entry, and then remove it: 
    Get-PublicFolderClientPermission -Identity <Folder>| Where {$_.User -like 
    <SID>} | Remove-PublicFolderClientPermission -Confirm:$False 

    The story would end here if Microsoft's script was perfect, and unfortunately 
    on a recent migration I encountered a scenario where it didn't pick up all 
    problem ACLs

    The scenario in question was one where my customer, after migration, was 
    converting leaver's mailboxes to shared mailboxes. Quite rightly they were 
    following [this support 
    article](https://support.microsoft.com/en-gb/help/2710029/shared-mailboxes-are-unexpectedly-converted-to-user-mailboxes-after-di) 
    from Microsoft which recommends setting the _msExchRecipientTypeDetails_ to a 
    particular value. The result of that corrupted the way the permissions are 
    evaluated on the Public Folder permissions, so that they do not appear to be 
    "ACLable" (assignable as permissions)

    Upon further investigation, this also applies in another scenario – where 
    you remove Exchange attributes from a Mailbox but keep the underlying Active 
    Directory account (i.e. you run _Disable-Mailbox_). In that scenario the 
    permission also shows in the same way and cannot be applied on the destination 
    – nor can it be used by a user

    When this occurs the permission shows like this when examining it using 
    _Get-PublicFolderClientPermission_ or by using _ExFolders_: 
    ![Correcting Public Folder Permissions before an Office 365 
    Migration](https://www.practical365.com/wp-content/uploads/2020/04/image-5.png) 
    As you can see in the above example, it is prefixed with _NT User:_ and the 
    account name, rather than resolving to a Mailbox, Remote Mailbox or other 
    recipient

    To search for and then resolve this scenario, I've created a simple 
    script called _Remove-NTUSER.ps1_, which you can [download from my GitHub](https://github.com/spgoodman/p365scripts/blob/master/Remove-NTUSER.ps1).


    #-=-=-=-=-=-=-=-=

    .PARAMETER Mailbox
    Use the -Mailbox parameter to designate the mailbox. Any valid Exchange mailbox identifier can be specified. Multiple mailboxes can be specified in a comma-separated list or array, see examples below.
    .PARAMETER ResetDefaultLevel
    Switch to specify reset to *include* default permissions (e.g. coerce Default SecPrin to 'LimitedDetails' & Anonymous:None)
    .PARAMETER CalendarLimitedDetails
    Switch to default Calendar folder view to customized LimitedDetails (vs default AvailabilityOnly)
    .PARAMETER RemoveUnresolveable
    Switch to Remove broken-SID/non-ADUser-resolvable entries targets usertype:UNKNOWN, including getr-adusere solvable, that lack populated msExchRecipientTypeDetails property
    .PARAMETER IgnoreInternal
    Switch to ignore/leave-intact any pre-existing usertype:Internal folder grants
    PARAMETER IgnoreExternal
    Switch to ignore/leave-intact any pre-existing usertype:External folder grants
    .PARAMETER Ticket
    Ticket number
    .PARAMETER Quiet
    Use the -Quiet switch if you want to suppress output to the console.
    .PARAMETER includedfolders
    Configurable string array of folder names to be *included* in processing (generally defaults to these; override to use customize/targed list)[-includedfolders @('Inbox','Calendar')]
    .PARAMETER excludedfolders
    Configurable string array of folder names to be *excluded* from processing (generally defaults to these; override to use customize/targed list)[-excludedfolders @('Inbox','Calendar')]
    .PARAMETER Verbose
    The -Verbose switch provides additional details on the cmdlet progress, it can be useful when troubleshooting issues.
     .INPUTS
    A mailbox identifier.
    .OUTPUTS
    Array of Mailbox address, Folder name and User.
    .EXAMPLE
    PS> Reset-xoMailboxFolderPermissionsRecursive -Mailbox emailaddress@domain.com -ResetDefaultLevel -verbose -whatif:$true
    Typical single user FULL RESET pass , with Whatif & verbose. Includes RESET of all Default role grants to stock UNMODIFIED settings:
    - Effectively WIPES ALL USER-MODIFICATIONS from all user/publicl-accessible folders of the mailbox, 
    - Resets Calendar(s):Default role to 'LimitedDetails' (TTC, MS Default 'AvailabilityOnly' can be set using -CalendarLimitedDetails:$false) & Anonymous role:None. 
    - All other modifications perms are removed, including user configured Internal & External grants. 
    .EXAMPLE
    PS> Reset-xoMailboxFolderPermissionsRecursive -Mailbox 
    Typical single-user pass no ResetDefaultLevel (any user-modifications to the Default & Anonymous roles are left intact), commit updates.
    .EXAMPLE
    PS> Reset-xoMailboxFolderPermissionsRecursive -Mailbox emailaddress@domain.com -ticket 123456 -RemoveUnresolveable -IgnoreInternal -IgnoreExternal -whatif:$false  ; 
    Single user pass, no Default Rest, targets removal of Unresolvable 'broken' grants (NT:S-1-5-21-...) from all user/public-accessible mailbox folders.
    .EXAMPLE
    PS> Reset-xoMailboxFolderPermissionsRecursive -Mailbox emailaddress@domain.com -ticket 123456 -IgnoreInternal -IgnoreExternal -whatif:$false  ; 
    Single user pass, no Default Rest, specifies to ignore both External & Internal user-added grants (left intact). 
    Effectively, this does nothing. Defaults are left unreset. Unknown/broken aren't targeted. and even default Internal/Externals are left untargeted.
    .EXAMPLE
    PS> Reset-xoMailboxFolderPermissionsRecursive -Mailbox @('emailaddress@domain.com','emailaddres2s@domain.com') -ResetDefaultLevel -verbose -whatif:$true
    Typical two-user pass as array, using specifying to include reset of all Default role grants to stock unmoidifed settings, with Whatif & verbose. 
    .EXAMPLE
    PS> Reset-xoMailboxFolderPermissionsRecursive -Mailbox @('emailaddress@domain.com','emailaddres2s@domain.com') -CalendarLimitedDetails:$false ;
    Demo override CalendarLimitedDetails (use the MS default Calendar visibility, 'AvailabilityOnly' (vs this script's default variant 'LimitedDetails').
    .EXAMPLE
    PS> Get-ADPermission -Identity "Christopher Payne" | ?{$_.user -like "S-1-5-21*"} | Remove-ADPermission
    Remove orphaned SID with Exchange Onprem PowerShell
    .EXAMPLE
    Reset-xoMailboxFolderPermissionsRecursive -Mailbox (Get-Mailbox -RecipientTypeDetails RoomMailbox) -Verbose
    This command removes permissions on all user-accessible folders in ALL Room mailboxes in the organization.
    .LINK
    https://github.com/tostka/powershell
    #>
    #Requires -Version 3.0
    [CmdletBinding(SupportsShouldProcess)] #Make sure we can use -WhatIf and -Verbose
    #[CmdletBinding()
    PARAM(
        [Parameter(Mandatory=$False,HelpMessage="Ticket Number [-Ticket '999999']")]
            [string]$Ticket,
        [Parameter(Mandatory=$true,ValueFromPipeline=$false,HelpMessage="Use the -Mailbox parameter to designate the mailbox. Any valid Exchange mailbox identifier can be specified. Multiple mailboxes can be specified in a comma-separated list or array, see examples below.")]
            [ValidateNotNullOrEmpty()]
            [Alias("Identity")]
            [String[]]$Mailbox,
        [Parameter(HelpMessage="Switch to specify reset to *include* default permissions")]
            [switch]$ResetDefaultLevel,
        # "AvailabilityOnly" v LimitedDetails custom
        [Parameter(HelpMessage="Switch to default Calendar folder view to customized LimitedDetails (vs default AvailabilityOnly)")]
            [switch]$CalendarLimitedDetails=$true,
        [Parameter(HelpMessage="Switch to Remove broken-SID/non-ADUser-resolvable entries")]
            [switch]$RemoveUnresolveable,
        [Parameter(HelpMessage="Switch to ignore/leave-intact any pre-existing usertype:Internal folder grants[-IgnoreInternal]")]
            [switch]$IgnoreInternal,
        [Parameter(HelpMessage="Switch to ignore/leave-intact any pre-existing usertype:External folder grants[-IgnoreInternal]")]
            [switch]$IgnoreExternal,
            [switch]$Quiet,
        [Parameter(HelpMessage="Configurable string array of folder names to be *included* in processing (generally defaults to these; override to use customize/targed list)[-includedfolders @('Inbox','Calendar'))")]        
            [string[]]$includedfolders = @("Root","Inbox","Calendar","Contacts","DeletedItems","Drafts","JunkEmail","Journal","Notes","Outbox","SentItems","Tasks","CommunicatorHistory","Clutter","Archive"), 
        [Parameter(HelpMessage="Configurable string array of folder names to be *excluded* from processing (generally defaults to these; override to use customize/targed list)[-excludedfolders @('Inbox','Calendar')]")]        
            [string[]]$excludedfolders = @("News Feed","Quick Step Settings","Social Activity Notifications","Suggested Contacts", "SearchDiscoveryHoldsUnindexedItemFolder", "SearchDiscoveryHoldsFolder","Calendar Logging")
        #[Parameter(HelpMessage="Whatif Flag  [-whatIf]")]
        #    [switch] $whatIf
    ) ; 
    # $CalendarLimitedDetails isn't coming through clean, force it up and move on for now
    BEGIN{
        $PSParameters = New-Object -TypeName PSObject -Property $PSBoundParameters ;
        write-verbose "`$PSBoundParameters:`n$(($PSBoundParameters|out-string).trim())" ;
        $Verbose = ($VerbosePreference -eq 'Continue') ; 
        if($WhatIfPreference){
            $whatif = $true ; 
        } else { $whatif = $false } ; 

        $DefaultRoleUserNames = @("Default","Anonymous","Owner@local","Member@local") ; 
        if(!$ThrottleMs){$ThrottleMs = 500} ; 

        if($CalendarLimitedDetails){$CalPermsDefault = 'LimitedDetails' }
        else {$CalPermsDefault = 'AvailabilityOnly' }

        #$includedfolders = @("Root","Inbox","Calendar", "Contacts", "DeletedItems", "Drafts", "JunkEmail", "Journal", "Notes", "Outbox", "SentItems", "Tasks", "CommunicatorHistory", "Clutter", "Archive") ; 
        #$includedfolders = @("Root","Inbox","Calendar", "Contacts", "DeletedItems", "SentItems", "Tasks") #Trimmed down list of default folders
        #Exclude additional Non-default folders created by Outlook or other mail programs. Folder NAMES, not types! So make sure to include translations too!
        #Exclude SearchDiscoveryHoldsFolder and SearchDiscoveryHoldsUnindexedItemFolder as they're not marked as default folders #Exclude "Calendar Logging" on older Exchange versions
        #$excludedfolders = @("News Feed","Quick Step Settings","Social Activity Notifications","Suggested Contacts", "SearchDiscoveryHoldsUnindexedItemFolder", "SearchDiscoveryHoldsFolder","Calendar Logging") ; 
        $prpADU = 'DistinguishedName','enabled','samaccountname','sid','UserPrincipalName' ; 

        #region ENVIRO_DISCOVER ; #*------v ENVIRO_DISCOVER v------
        #if ($PSScriptRoot -eq "") {
        # 8/29/2023 fix logic break on psv2 ISE (doesn't test PSScriptRoot -eq '' properly, needs $null test).
        #if( -not (get-variable -name PSScriptRoot -ea 0) -OR ($PSScriptRoot -eq '')){
        if( -not (get-variable -name PSScriptRoot -ea 0) -OR ($PSScriptRoot -eq '') -OR ($PSScriptRoot -eq $null)){
            if ($psISE) { $ScriptName = $psISE.CurrentFile.FullPath } 
            elseif($psEditor){
                if ($context = $psEditor.GetEditorContext()) {$ScriptName = $context.CurrentFile.Path } 
            } elseif ($host.version.major -lt 3) {
                $ScriptName = $MyInvocation.MyCommand.Path ;
                $PSScriptRoot = Split-Path $ScriptName -Parent ;
                $PSCommandPath = $ScriptName ;
            } else {
                if ($MyInvocation.MyCommand.Path) {
                    $ScriptName = $MyInvocation.MyCommand.Path ;
                    $PSScriptRoot = Split-Path $MyInvocation.MyCommand.Path -Parent ;
                } else {throw "UNABLE TO POPULATE SCRIPT PATH, EVEN `$MyInvocation IS BLANK!" } ;
            };
            if($ScriptName){
                $ScriptDir = Split-Path -Parent $ScriptName ;
                $ScriptBaseName = split-path -leaf $ScriptName ;
                $ScriptNameNoExt = [system.io.path]::GetFilenameWithoutExtension($ScriptName) ;
            } ; 
        } else {
            if($PSScriptRoot){$ScriptDir = $PSScriptRoot ;}
            else{
                write-warning "Unpopulated `$PSScriptRoot!" ; 
                $ScriptDir=(Split-Path -parent $MyInvocation.MyCommand.Definition) + "\" ;
            }
            if ($PSCommandPath) {$ScriptName = $PSCommandPath } 
            else {
                $ScriptName = $myInvocation.ScriptName
                $PSCommandPath = $ScriptName ;
            } ;
            $ScriptBaseName = (Split-Path -Leaf ((& { $myInvocation }).ScriptName))  ;
            $ScriptNameNoExt = [system.io.path]::GetFilenameWithoutExtension($MyInvocation.InvocationName) ;
        } ;
        if(-not $ScriptDir){
            write-host "Failed `$ScriptDir resolution on PSv$($host.version.major): Falling back to $MyInvocation parsing..." ; 
            $ScriptDir=(Split-Path -parent $MyInvocation.MyCommand.Definition) + "\" ;
            $ScriptBaseName = (Split-Path -Leaf ((&{$myInvocation}).ScriptName))  ; 
            $ScriptNameNoExt = [system.io.path]::GetFilenameWithoutExtension($MyInvocation.InvocationName) ;     
        } else {
            if(-not $PSCommandPath ){
                $PSCommandPath  = $ScriptName ; 
                if($PSCommandPath){ write-host "(Derived missing `$PSCommandPath from `$ScriptName)" ; } ;
            } ; 
            if(-not $PSScriptRoot  ){
                $PSScriptRoot   = $ScriptDir ; 
                if($PSScriptRoot){ write-host "(Derived missing `$PSScriptRoot from `$ScriptDir)" ; } ;
            } ; 
        } ; 
        if(-not ($ScriptDir -AND $ScriptBaseName -AND $ScriptNameNoExt)){ 
            throw "Invalid Invocation. Blank `$ScriptDir/`$ScriptBaseName/`ScriptNameNoExt" ; 
            BREAK ; 
        } ; 

        $smsg = "`$ScriptDir:$($ScriptDir)" ;
        $smsg += "`n`$ScriptBaseName:$($ScriptBaseName)" ;
        $smsg += "`n`$ScriptNameNoExt:$($ScriptNameNoExt)" ;
        $smsg += "`n`$PSScriptRoot:$($PSScriptRoot)" ;
        $smsg += "`n`$PSCommandPath:$($PSCommandPath)" ;  ;
        write-verbose $smsg ; 
        #endregion ENVIRO_DISCOVER ; #*------^ END ENVIRO_DISCOVER ^------        
        write-verbose "checking depednant function availability..." ; 
        $depCmdlets = @('get-XOMailboxFolderList','get-XOMailboxFolderPermissionsSummary') ; 
        $depCmdlets | foreach-object{if(get-command $_ ){write-verbose "gcm'd:dependant function:$($_)"} else { $smsg = "Missing dependant function:$($_)" ; write-warning $smsg ; throw $smsg ; }} ;
        # EOM3+ NO PSS SUPP
        #if (-not ((Get-ConnectionInformation).tokenstatus -eq 'Active')){ Write-Error "No active Exchange connection detected, please connect first. To connect to ExO: https://technet.microsoft.com/en-us/library/jj984289(v=exchg.160).aspx" -ErrorAction Stop ;} ; 
        #Prepare the list of mailboxes
        $smsg = "Parsing the Mailbox parameter..."
        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
        $SMTPAddresses = @{}
        foreach ($mb in $Mailbox) {
            Start-Sleep -Milliseconds 80 #Add some delay to avoid throttling...
            #Make sure a matching mailbox is found and return its Primary SMTP Address
            #$SMTPAddress = (Invoke-Command -Session $session -ScriptBlock { Get-Mailbox $using:mb | Select-Object -ExpandProperty PrimarySmtpAddress } -ErrorAction SilentlyContinue).Address
            # eom3+ direct no pss
            #$SMTPAddress = Get-xoMailbox $mb | Select-Object -ExpandProperty PrimarySmtpAddress -ErrorAction SilentlyContinue;
            #*======v BP Wrapper for running EXO dynamic prefixed-EOM310+ cmdlets (psb-PsXoPrfx.cbp) v======
            # define the splat of all params:
            $pltGMbx = [ordered]@{identity =  $mb ; erroraction = 'STOP'; verbose = $($VerbosePreference -eq "Continue") ;} ;
            $cmdlet = 'get-Mailbox' ; $verb,$noun = $cmdlet.split('-') ;  #Spec cmdletname (VERB-NOUN), & split v/n
            TRY{$xoS = Get-ConnectionInformation -ErrorAction STOP }CATCH{reconnect-exo ; $xoS = Get-ConnectionInformation -ErrorAction STOP }
            TRY{
                if((-not $xos) -OR ($xoS | ?{$_.tokenstatus -notmatch 'Active|Expired' -AND $_.State -ne 'Connected'} )){reconnect-exo ; $xoS = Get-ConnectionInformation -ErrorAction STOP } ; 
                if($xos){
                    $xcmd = "$verb-$($xoS.ModulePrefix)$noun `@pltGMbx" ; # build cmdline w splat, then echo:
                    $smsg = "$($([regex]::match($xcmd,"^(\w+)-(\w+)" ).groups[0].value)) w`n$(($pltGMbx|out-string).trim())" ;
                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                    $SMTPAddress  = invoke-expression $xcmd  | 
                        Select-Object -ExpandProperty PrimarySmtpAddress -ErrorAction SilentlyContinue;
                    if($SMTPAddress){
                        $smsg = "(confirmed valid $([regex]::match($xcmd,"^(\w+)-(\w+)" ).groups[0].value) output)" ; 
                        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                    }
                } else { 
                    $smsg = "Missing `$xos EXO connection!" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                    throw $smsg ; BREAK ; 
                } 
            } CATCH {
                $ErrTrapd=$Error[0] ;
                $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
            } ; 
            <# version 12:43 PM 9/21/2023 moved cixo up to 1st, won't have prefix if not populated, also needs to fail/retry to ensure conn;  
            11:48 AM 9/20/2023 minor tweaks ; 3:01 PM 9/19/2023 initial 
            ## this runs: 1) connection status check, w rxo on demand; 2) splat wrapper with integrated prefix support; 3) try/catch on exec; 
            useful alias: cixo => get-connectioninformation;
            #>
            #*======^ END BP wrapper for running dynamic prefixed-EOM310+ cmdlets (psb-PsXoPrfx.cbp)  ^======
            if (-not $SMTPAddress) { if (-not $Quiet) { 
                $smsg = "Mailbox with identifier $mb not found, skipping..." ;if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; }; 
                continue 
            } elseif (($SMTPAddress.count -gt 1) -or ($SMTPAddresses[$mb]) -or ($SMTPAddresses.ContainsValue($SMTPAddress))) { 
                $smsg = "Multiple mailboxes matching the identifier $mb found, skipping..."; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                continue 
            }else { $SMTPAddresses[$mb] = $SMTPAddress } ; 
        }
        if (-not $SMTPAddresses -or ($SMTPAddresses.Count -eq 0)) { Throw "No matching mailboxes found, check the parameter values." } ; 
        $smsg = "The following list of mailboxes will be used: ""$($SMTPAddresses.Values -join ", ")""" ; 
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
        $smsg = "List of default folder TYPES that will be used: ""$($includedfolders -join ", ")""" ; 
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
        $smsg = "List of folder NAMES that will be excluded: ""$($excludedfolders -join ", ")""" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
    }
    PROCESS{
        $out = @() ; 
        foreach ($smtp in $SMTPAddresses.Values) {
            $smsg = $sBnrS="`n#*------v PROCESSING Mailbox: $($smtp)... v------" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level H2 } 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
            Start-Sleep -Milliseconds 800  ; #Add some delay to avoid throttling...
            $smsg = "Obtaining folder list for mailbox ""$smtp""..." ; 
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            $folders = get-XOMailboxFolderList $smtp ; 
            $smsg = "A total of $($folders.count) folders found for $($smtp)." ; 
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            if (-not $folders) {
                $smsg = "No matching folders found for $($smtp), skipping..." ; 
                continue  ; 
            } ; 
            #Cycle over each folder we are interested in
            foreach ($folder in $folders) {
                $smsg = "`n==PROCESSING:$($folder.name)`n" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level H3 } 
                else{ write-host -foregroundcolor white "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                #"Fix" for folders with "/" characters, treat the Root folder separately
                if ($folder.FolderType -eq "Root") { $foldername = $smtp }
                else { $foldername = $folder.Identity.ToString().Replace([char]63743,"/").Replace($smtp,$smtp + ":") } ; 
                $fPermissions = get-XOMailboxFolderPermissionsSummary $foldername
                if (-not $ResetDefaultLevel) { 
                    $smsg = "no -ResetDefaultLevel: exempting username:$($DefaultRoleUserNames -join '|') from processing" ; 
                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;  ; 
                    $fPermissions = $fPermissions | ? {$_.UserName -notin @("Default","Anonymous","Owner@local","Member@local")}
                    #$fPermissions = $fPermissions | ? {$_.UserName -notin @($($DefaultRoleUserNames)))}
                }  ; #filter out default permissions -> doesn't process defaults, leaves them intact
                if (-not $fPermissions) { 
                    $smsg = "No permission entries found for $($foldername), skipping..." ; 
                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                    continue  ; 
                } ; 

                #Remove the folder permissions for each delegate
                foreach ($u in $fPermissions) {
                    $smsg = "`n" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                    if ($u.UserType -eq "Default") {
                        <# effectively: This forcec Default Cal to AvailOnly (or LtdDetails), and any other Default to NONE. -> It strips all but Cal Default access.
                         Details: 
                            Default perms only get here if $ResetDefaultLevel:$true, this forces xxx:\Calendar Defaults to $CalPermsDefaul (AvailabilityOnly (MS default) or LimitedDetails (local Org), depending on specification)
                            Non Cal folders get set to NONE
                        #>
                        #UserType enumeration https://docs.microsoft.com/en-us/previous-versions/office/developer/exchange-server-2010/ff319704(v%3Dexchg.140) hardcoded solely: Default|Anonymous|Internal|External|Unknown
                        if ($ResetDefaultLevel) {
                            <# https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_scopes?view=powershell-7.3#the-using-scope-modifier
                                scope $using: - Used to access variables defined in another scope while running scripts via cmdlets like Start-Job and Invoke-Command.

                            #>
                            TRY {
                                $smsg = "Resetting permissions on ""$foldername"" for principal ""Default""." ;
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Prompt } 
                                else{ write-host -foregroundcolor yellow "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                                if ($folder.FolderType -eq "Calendar") {
                                    # force any Calendar folder Default grant to AvailOnly (or LtdDetails)
                                    #if (($u.AccessRights -join ",") -ne "AvailabilityOnly") {
                                    # TTC customizes the view as LimitedDetails: $CalPermsDefault
                                    if (($u.AccessRights -join ",") -ne $CalPermsDefault) {
                                        #Invoke-Command -Session $session -ScriptBlock { Set-MailboxFolderPermission -Identity $Using:foldername -User Default -AccessRights AvailabilityOnly -WhatIf:$using:WhatIfPreference -Confirm:$false } -ErrorAction Stop -HideComputerName ;
                                        # can't use -session $session with EOM3+, try direct calls; should work
                                        #Set-xoMailboxFolderPermission -Identity $foldername -User Default -AccessRights AvailabilityOnly -WhatIf:$WhatIfPreference -Confirm:$false -ErrorAction Stop ;
                                        #Set-xoMailboxFolderPermission -Identity $foldername -User Default -AccessRights $CalPermsDefault -WhatIf:$WhatIfPreference -Confirm:$false -ErrorAction Stop ;
                                        #*======v BP Wrapper for running EXO dynamic prefixed-EOM310+ cmdlets (psb-PsXoPrfx.cbp) v======
                                        # define the splat of all params:
                                        $pltSMbxFP = [ordered]@{
                                            Identity =$foldername ;
                                            User ='Default' ;
                                            AccessRights =$CalPermsDefault ;
                                            WhatIf =$WhatIfPreference ;
                                            Confirm =$false ;
                                            ErrorAction = 'Stop' ; 
                                            verbose = $($VerbosePreference -eq "Continue") ;
                                        } ;
                                        $cmdlet = 'Set-MailboxFolderPermission' ; $verb,$noun = $cmdlet.split('-') ;  #Spec cmdletname (VERB-NOUN), & split v/n
                                        TRY{$xoS = Get-ConnectionInformation -ErrorAction STOP }CATCH{reconnect-exo ; $xoS = Get-ConnectionInformation -ErrorAction STOP }
                                        TRY{
                                            if((-not $xos) -OR ($xoS | ?{$_.tokenstatus -notmatch 'Active|Expired' -AND $_.State -ne 'Connected'} )){reconnect-exo ; $xoS = Get-ConnectionInformation -ErrorAction STOP } ; 
                                            if($xos){
                                                $xcmd = "$verb-$($xoS.ModulePrefix)$noun `@pltSMbxFP" ; # build cmdline w splat, then echo:
                                                $smsg =  "$($([regex]::match($xcmd,"^(\w+)-(\w+)" ).groups[0].value)) w`n$(($pltSMbxFP|out-string).trim())" ;
                                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                                                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                                #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                                                $RET = invoke-expression $xcmd  ;
                                                if($RET){
                                                    $smsg = "(confirmed valid $([regex]::match($xcmd,"^(\w+)-(\w+)" ).groups[0].value) output)" ; 
                                                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                                                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                                                }
                                            } else { 
                                                $smsg = "Missing `$xos EXO connection!" ; 
                                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                                                throw $smsg ; BREAK ; 
                                            } 
                                        } CATCH {
                                            $ErrTrapd=$Error[0] ;
                                            $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                                            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                                        } ; 
                                        <# version 12:43 PM 9/21/2023 moved cixo up to 1st, won't have prefix if not populated, also needs to fail/retry to ensure conn;  
                                        11:48 AM 9/20/2023 minor tweaks ; 3:01 PM 9/19/2023 initial 
                                        ## this runs: 1) connection status check, w rxo on demand; 2) splat wrapper with integrated prefix support; 3) try/catch on exec; 
                                        useful alias: cixo => get-connectioninformation;
                                        #>
                                        #*======^ END BP wrapper for running dynamic prefixed-EOM310+ cmdlets (psb-PsXoPrfx.cbp)  ^======
                                    } else { continue } ; 
                                    $outtemp = New-Object psobject -Property ([ordered]@{"Mailbox" = $smtp;"FolderName" = $folder.name;"User" = $u.UserName;"AccessRights" = "AvailabilityOnly"}) ; 
                                } else {
                                    # force any non-calendar folder with a Default grant, to NONE
                                    if (($u.AccessRights -join ",") -ne "None") {
                                        #Invoke-Command -Session $session -ScriptBlock { Set-MailboxFolderPermission -Identity $Using:foldername -User Default -AccessRights None -WhatIf:$using:WhatIfPreference -Confirm:$false } -ErrorAction Stop -HideComputerName 
                                        # eom3+ no pss
                                        #Set-xoMailboxFolderPermission -Identity $Using:foldername -User Default -AccessRights None -WhatIf:$using:WhatIfPreference -Confirm:$false -ErrorAction Stop ;
                                    
                                        #*======v BP Wrapper for running EXO dynamic prefixed-EOM310+ cmdlets (psb-PsXoPrfx.cbp) v======
                                        # define the splat of all params:
                                        $pltSMbxFP = [ordered]@{
                                            Identity =$foldername ;
                                            User ='Default' ;
                                            AccessRights = 'None' ;
                                            WhatIf =$WhatIfPreference ;
                                            Confirm =$false ;
                                            ErrorAction = 'Stop' ; 
                                            verbose = $($VerbosePreference -eq "Continue") ;
                                        } ;
                                        $cmdlet = 'Set-MailboxFolderPermission' ; $verb,$noun = $cmdlet.split('-') ;  #Spec cmdletname (VERB-NOUN), & split v/n
                                        TRY{$xoS = Get-ConnectionInformation -ErrorAction STOP }CATCH{reconnect-exo ; $xoS = Get-ConnectionInformation -ErrorAction STOP }
                                        TRY{
                                            if((-not $xos) -OR ($xoS | ?{$_.tokenstatus -notmatch 'Active|Expired' -AND $_.State -ne 'Connected'} )){reconnect-exo ; $xoS = Get-ConnectionInformation -ErrorAction STOP } ; 
                                            if($xos){
                                                $xcmd = "$verb-$($xoS.ModulePrefix)$noun `@pltSMbxFP" ; # build cmdline w splat, then echo:
                                                $smsg = "$($([regex]::match($xcmd,"^(\w+)-(\w+)" ).groups[0].value)) w`n$(($pltSMbxFP|out-string).trim())" ;
                                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                                                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                                #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                                                $RET = invoke-expression $xcmd  ;
                                                if($RET){$smsg = "(confirmed valid $([regex]::match($xcmd,"^(\w+)-(\w+)" ).groups[0].value) output)" ; if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; }
                                            } else { 
                                                $smsg = "Missing `$xos EXO connection!" ; 
                                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                                                throw $smsg ; BREAK ; 
                                            } 
                                        } CATCH {
                                            $ErrTrapd=$Error[0] ;
                                            $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                                            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                                        } ; 
                                        <# version 12:43 PM 9/21/2023 moved cixo up to 1st, won't have prefix if not populated, also needs to fail/retry to ensure conn;  
                                        11:48 AM 9/20/2023 minor tweaks ; 3:01 PM 9/19/2023 initial 
                                        ## this runs: 1) connection status check, w rxo on demand; 2) splat wrapper with integrated prefix support; 3) try/catch on exec; 
                                        useful alias: cixo => get-connectioninformation;
                                        #>
                                        #*======^ END BP wrapper for running dynamic prefixed-EOM310+ cmdlets (psb-PsXoPrfx.cbp)  ^======
                                    }
                                    else { continue } ; 
                                    $outtemp = New-Object psobject -Property ([ordered]@{"Mailbox" = $smtp;"FolderName" = $folder.name;"User" = $u.UserName;"AccessRights" = "None"}) ; 
                                } ; 
                                $out += $outtemp; if (-not $Quiet -and -not $WhatIfPreference) { 
                                    #$outtemp 
                                    $smsg = "`n$(($outtemp | ft -a |out-string).trim())" ; 
                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                }  ; #Write output to the console unless the -Quiet parameter is used
                            } CATCH {
                                #$_ | fl * -Force; continue
                                $smsg = "`n$(($_ | fl *|out-string).trim())" ; 
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                                continue
                            }  ; #catch-all for any unhandled errors
                        } else { continue } ; 

                    } elseif ($u.UserType -eq "Anonymous") { 
                        $smsg = "$($u.username):UserType:Anonymous, skipping processing" ; 
                        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                        #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                        continue 
                        #Maybe set them all to none when $resetdefault is used?

                    } elseif ($u.UserType -eq "Unknown") { 
                        <# Add param: -RemoveUnresolveable -> targets usertype:UNKNOWN, including getr-adusere solvable, that lack populated msExchRecipientTypeDetails property
                        Add param: -IgnoreInternal - skips reset of existing usertype:Internal to NONE
                        Add param: -IgnoreExternal - skips reset of existing usertype:External to NONE
                        Switch to ignore/leave-intact any pre-existing usertype:Internal folder grants
                        Switch to ignore/leave-intact any pre-existing usertype:External folder grants
                        #>
                        $smsg = "'UNKNOWN entry':" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                        $DoRemove = $false ; 
                        if($RemoveUnresolveable -AND ($u.UserName -match '^NT:S-1-5-21-')){
                            $smsg = "(entry UserName appears to be a BROKEN SID (SECURITY IDENTIFYER == DELETED USER OBJECT/NON-RECIPIENT)" ; 
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                        
                            $smsg = "(attempting: get-aduser -id $($u.UserName.replace('NT:','')) )" ; 
                            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 

                            $DoRemove = $false ; 
                            TRY{
                                if($ADU =  get-aduser -id ($u.UserName.replace('NT:','')) -ErrorAction SilentlyContinue -prop msExchRecipientTypeDetails){
                                    $smsg = "Resolved`n $($u.UserName.replace('NT:',''))`n to an existing ADUser object:`n$(($adu | fl $prpADU |out-string).trim())" ; 
                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                                }else {
                                    $DoRemove = $true ; 
                                    $smsg = "Unable to resolve $($u.UserName.replace('NT:','')) to an existing ADUser object: => REMOVE Grant!" ; 
                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                                    $DoRemove = $true ; 
                                } ; 
                            } CATCH [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException]{
                                $smsg = "Unable to resolve $($u.UserName.replace('NT:','')) to an existing ADUser object: => REMOVE Grant!" ; 
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                                $DoRemove = $true ;
                            } CATCH {
                                $smsg = "$(($_ | fl * |out-string).trim())" ; 
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                                continue 
                            };#catch-all for any unhandled errors
                            if($ADU){
                                $smsg = "Test for EX recipient: populated  msExchRecipientTypeDetail" ; 
                                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                                if($ADU.msExchRecipientTypeDetails){
                                    $smsg = "Found ADUser:$($ADU.userprincipalname) *has* populated msExchRecipientTypeDetail:$($ADU.msExchRecipientTypeDetail)`n*LEAVING EXISTING GRANT IN PLACE!" ; 
                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                                    $DoRemove = $false ; 
                                } else { 
                                    $smsg = "Found ADUser:$($ADU.userprincipalname) has *NO* populated msExchRecipientTypeDetail:$($ADU.msExchRecipientTypeDetail)`n*=> Non-Recipient Security Principal: REMOVE Grant!" ; 
                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                                    $DoRemove = $true ; 
                                } ; 
                             } ;    
                            
                        } elseif($RemoveUnresolveable -AND ($u.User -eq $null)){
                            # non-guid likely still has blank User/user.RecipientPrincipa.guid.guid ($_.user.RecipientPrincipal.value resolve)
                            $smsg = "entry UserName is populated non SID but User is blank (reflects unresolved underlying RecipientPrincipa.guid.guid)" ; 
                            $smsg += "`n$(($u | fl *|out-string).trim())" ; 
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 

                            $smsg += "`n(Setting `$DoRemove:`$true)" ; 
                            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                            $DoRemove = $true ; 
                        } else { 
                            $DoRemove = $false ; 
                            # 12:19 PM 9/25/2023 the DC, entry falls through here, it's got no 
                            $smsg = "Skipping orphaned permissions entry: $($u.UserName)";
                            $smsg += "`n$(($u | fl *|out-string).trim())" ; 
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Prompt } 
                            else{ write-host -foregroundcolor yellow "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                            continue  ; 
                        } ; 

                        # removal handling here on $DoRemove spec
                        if($DoRemove){
                            $smsg = "`nREMOVING NON-FUNCTIONAL GRANT!"
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 

                            # eom3+ no pss
                            #Remove-xoMailboxFolderPermission -Identity $foldername -User $u.User -WhatIf:$WhatIfPreference -Confirm:$false -ErrorAction Stop ;
                            <#
                            # Expand the full name out of the above:
                            Get-exoMailboxFolderPermission -Identity "$($TMBX):\Calendar" | select -expand User | select -expand displayname
                            # out: 
                            NT:S-1-5-21-2222296782-158576315-1096482972-20544
                            #
                            # Target it for removal:(can use the name displayed):
                            Remove-xoMailboxFolderPermission -Identity "$($tmbx):\Calendar” -User "NT:S-1-5-21-2222296782-158576315-1096482972-20544" -whatif ; 
                            => use the populated $u.UserName from this script as -User
                            #>
                            #*======v BP Wrapper for running EXO dynamic prefixed-EOM310+ cmdlets (psb-PsXoPrfx.cbp) v======
                            # define the splat of all params:
                            $pltRMbxFP = [ordered]@{
                                Identity =$foldername ;
                                #User =$u.User ; # user is blank on UNKNOWN's so use username
                                User = $u.UserName ; 
                                #AccessRights = 'None' ;
                                WhatIf =$WhatIfPreference ;
                                Confirm =$false ;
                                ErrorAction = 'Stop' ; 
                                verbose = $($VerbosePreference -eq "Continue") ;
                            } ;
                            $cmdlet = 'Remove-MailboxFolderPermission' ; $verb,$noun = $cmdlet.split('-') ;  #Spec cmdletname (VERB-NOUN), & split v/n
                            TRY{$xoS = Get-ConnectionInformation -ErrorAction STOP }CATCH{reconnect-exo ; $xoS = Get-ConnectionInformation -ErrorAction STOP }
                            TRY{
                                if((-not $xos) -OR ($xoS | ?{$_.tokenstatus -notmatch 'Active|Expired' -AND $_.State -ne 'Connected'} )){reconnect-exo ; $xoS = Get-ConnectionInformation -ErrorAction STOP } ; 
                                if($xos){
                                    $xcmd = "$verb-$($xoS.ModulePrefix)$noun `@pltRMbxFP" ; # build cmdline w splat, then echo:
                                    $smsg = "$($([regex]::match($xcmd,"^(\w+)-(\w+)" ).groups[0].value)) w`n$(($pltRMbxFP|out-string).trim())" ;
                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                    #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                                    $RET = invoke-expression $xcmd  ;
                                    if($RET){$smsg = "(confirmed valid $([regex]::match($xcmd,"^(\w+)-(\w+)" ).groups[0].value) output)" ; if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; }
                                } else { 
                                    $smsg = "Missing `$xos EXO connection!" ; 
                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                                    throw $smsg ; BREAK ; 
                                } 
                            } CATCH {
                                $ErrTrapd=$Error[0] ;
                                $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                                write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" ;
                            } ; 
                            <# version 12:43 PM 9/21/2023 moved cixo up to 1st, won't have prefix if not populated, also needs to fail/retry to ensure conn;  
                            11:48 AM 9/20/2023 minor tweaks ; 3:01 PM 9/19/2023 initial 
                            ## this runs: 1) connection status check, w rxo on demand; 2) splat wrapper with integrated prefix support; 3) try/catch on exec; 
                            useful alias: cixo => get-connectioninformation;
                            #>
                            #*======^ END BP wrapper for running dynamic prefixed-EOM310+ cmdlets (psb-PsXoPrfx.cbp)  ^======
                            $outtemp = New-Object psobject -Property ([ordered]@{"Mailbox" = $smtp;"FolderName" = $folder.name;"User" = $u.UserName}) ;
                            #Write output to the console unless the -Quiet parameter is used 
                            $out += $outtemp; if (-not $Quiet -and -not $WhatIfPreference) { 
                                #$outtemp 
                                $smsg = "`n$(($outtemp | ft -a |out-string).trim())" ; 
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            }  ; #Write output to the console unless the -Quiet parameter is used
                        } else { 
                            $smsg = "`$DoRemove:$($DoRemove): skipping removal of usertype:UNKNOWN grant" ; 
                            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                        } ; 

                    } else {
                        # other: External|Internal wind up here
                        if ($u.UserType -eq "External") { $u.User = $u.UserName }
                        <#
                        Add param: -IgnoreInternal - skips reset of existing usertype:Internal to NONE
                        Add param: -IgnoreExternal - skips reset of existing usertype:External to NONE
                        #>
                        if($u.UserType -eq "External" -AND $IgnoreExternal){
                            $smsg = "UserType:External with -IgnoreExternal specified: *SKIPPING* default purge of EXTERNAL Grant:" ;
                            $smsg += "`n`n$(($u | ft -a identity,user,usertype,username,accessrights|out-string).trim())" ;  
                            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                            #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                            Continue 
                        } 
                        if($u.UserType -eq "Internal" -AND $IgnoreInternal){ 
                            $smsg = "UserType:Internal with -IgnoreInternal specified: *SKIPPING* default purge of INTERNAL Grant:" ; 
                            $smsg += "`n`n$(($u | ft -a identity,user,usertype,username,accessrights|out-string).trim())" ; 
                            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                            #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                            Continue 
                        } 
                        TRY {
                            if (-not $u.User) { continue } ; 
                            $smsg = "Removing permissions on ""$foldername"" for principal ""$($u.UserName)""." ; 
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Prompt } 
                            else{ write-host -foregroundcolor yellow "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                            #Invoke-Command -Session $session -ScriptBlock { Remove-MailboxFolderPermission -Identity $Using:foldername -User $Using:u.User -WhatIf:$using:WhatIfPreference -Confirm:$false } -ErrorAction Stop -HideComputerName ;
                            # eom3+ no pss
                            #Remove-xoMailboxFolderPermission -Identity $foldername -User $u.User -WhatIf:$WhatIfPreference -Confirm:$false -ErrorAction Stop ;
                            #*======v BP Wrapper for running EXO dynamic prefixed-EOM310+ cmdlets (psb-PsXoPrfx.cbp) v======
                            # define the splat of all params:
                            $pltRMbxFP = [ordered]@{
                                Identity =$foldername ;
                                User =$u.User ;
                                #AccessRights = 'None' ;
                                WhatIf =$WhatIfPreference ;
                                Confirm =$false ;
                                ErrorAction = 'Stop' ; 
                                verbose = $($VerbosePreference -eq "Continue") ;
                            } ;
                            $smsg = "Spec cmdletname (VERB-NOUN), then convert cmdlet & splat to `$xcmd string" ; 
                            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                            $cmdlet = 'Remove-MailboxFolderPermission' ; $verb,$noun = $cmdlet.split('-') ;  #Spec cmdletname (VERB-NOUN), & split v/n
                            TRY{$xoS = Get-ConnectionInformation -ErrorAction STOP }CATCH{reconnect-exo ; $xoS = Get-ConnectionInformation -ErrorAction STOP }
                            TRY{
                                if((-not $xos) -OR ($xoS | ?{$_.tokenstatus -notmatch 'Active|Expired' -AND $_.State -ne 'Connected'} )){reconnect-exo ; $xoS = Get-ConnectionInformation -ErrorAction STOP } ; 
                                if($xos){
                                    $xcmd = "$verb-$($xoS.ModulePrefix)$noun `@pltRMbxFP" ; # build cmdline w splat, then echo:
                                    $smsg = "$($([regex]::match($xcmd,"^(\w+)-(\w+)" ).groups[0].value)) w`n$(($pltRMbxFP|out-string).trim())" ;
                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                    #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                                    $RET = invoke-expression $xcmd  ;
                                    if($RET){$smsg = "(confirmed valid $([regex]::match($xcmd,"^(\w+)-(\w+)" ).groups[0].value) output)" ; if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; }
                                } else { 
                                    $smsg = "Missing `$xos EXO connection!" ; 
                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                                    throw $smsg ; BREAK ; 
                                } 
                            } CATCH {
                                $ErrTrapd=$Error[0] ;
                                $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 

                            } ; 
                            <# version 12:43 PM 9/21/2023 moved cixo up to 1st, won't have prefix if not populated, also needs to fail/retry to ensure conn;  
                            11:48 AM 9/20/2023 minor tweaks ; 3:01 PM 9/19/2023 initial 
                            ## this runs: 1) connection status check, w rxo on demand; 2) splat wrapper with integrated prefix support; 3) try/catch on exec; 
                            useful alias: cixo => get-connectioninformation;
                            #>
                            #*======^ END BP wrapper for running dynamic prefixed-EOM310+ cmdlets (psb-PsXoPrfx.cbp)  ^======
                            $outtemp = New-Object psobject -Property ([ordered]@{"Mailbox" = $smtp;"FolderName" = $folder.name;"User" = $u.UserName}) ;
                            #Write output to the console unless the -Quiet parameter is used 
                            $out += $outtemp; if (-not $Quiet -and -not $WhatIfPreference) { 
                                #$outtemp 
                                $smsg = "`n$(($outtemp | ft -a |out-string).trim())" ; 
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            }  ; #Write output to the console unless the -Quiet parameter is used
                        } CATCH [System.Management.Automation.RemoteException] {
                            if (-not $Quiet) {
                                if ($_.CategoryInfo.Reason -eq "UserNotFoundInPermissionEntryException") { 
                                    $smsg = "WARNING: No existing permissions entry found on ""$foldername"" for principal ""$($u.UserName)""" 
                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                                }elseif ($_.CategoryInfo.Reason -eq "CannotChangePermissionsOnFolderException") { 
                                    $smsg = "ERROR: Folder permissions for ""$foldername"" CANNOT be changed!" ;
                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                                }elseif ($_.CategoryInfo.Reason -eq "CannotRemoveSpecialUserException") { 
                                    $smsg = "ERROR: Folder permissions for ""$($u.UserName)"" CANNOT be changed!" 
                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                                }elseif ($_.CategoryInfo.Reason -eq "ManagementObjectNotFoundException") { 
                                    $smsg = "ERROR: Folder ""$foldername"" not found, this should not happen..."
                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                                }elseif ($_.CategoryInfo.Reason -eq "InvalidInternalUserIdException") { 
                                    $smsg = "ERROR: ""$($u.UserName)"" is not a valid security principal for folder-level permissions..."
                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                                }else {
                                    $smsg = "`n$(($_ | fl *|out-string).trim())" ; 
                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 

                                    continue
                                }  ; #catch-all for any unhandled errors
                            } ;  # if-E !quiet
                        } catch {$_ | fl * -Force; continue} ;#catch-all for any unhandled errors
                    } # if-E
                }  ; # ACE loop-E
            } ;  # FOLDERS loop-E
            $smsg = $sBnrS.replace('-v','-^').replace('v-','^-')
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level H2 } 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success

        }  # MBX loop-E
    } ; # PROC-E
    END{
        if ($out) {
            #$out | Export-Csv -Path "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_MailboxFolderPermissionsRemoved.csv" -NoTypeInformation -Encoding UTF8 -UseCulture ;
            #$opath = "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_MailboxFolderPermissionsRemoved.csv" ; 
            #$smsg = "Exporting results to the CSV file...`n$($opath)" ;
            #$out | Export-Csv -Path $opath -NoTypeInformation -Encoding UTF8 -UseCulture ;
            [string]$opath = $null ; 
            if($ticket){$opath += "$($TICKET)-" }
            # $opath += "$($item)_MailboxFolderPermissionsRemoved-$(get-date -format 'yyyyMMdd-HHmmtt').csv" ; 
            if(($SMTPAddresses.Values |  measure | select -expand count ) -gt 3){
                $opath += "$($SMTPAddresses.Values[0]),xxx_MailboxFolderPermissionsRemoved-$(get-date -format 'yyyyMMdd-HHmmtt').csv" ; 
            } else { 
                $opath += "$($SMTPAddresses.Values -join ',')_MailboxFolderPermissionsRemoved-$(get-date -format 'yyyyMMdd-HHmmtt').csv" ; 
            } ; 
            $oPath = join-path -path (join-path -path $ScriptDir -childpath "logs") -ChildPath $opath ; 
            $smsg = "Exporting results to the CSV file...`n$($opath)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
            TRY{
                $out | Export-Csv -Path $opath -NoTypeInformation -Encoding UTF8 -UseCulture ;
            } CATCH {
                $ErrTrapd=$Error[0] ;
                $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } ; 
            #Write output to the console unless the -Quiet parameter is used
            #if (-not $Quiet -and -not $WhatIfPreference) { return $out | Out-Default }  ; 
            if (-not $Quiet -and -not $WhatIfPreference) { return $out  }  ; # above is returning as an array of text with no fields; output the object and aggregate it
        } else { 
            $smsg = "Output is empty, skipping the export to CSV file..." ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Prompt } 
            else{ write-host -foregroundcolor yellow "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
        } ;
        $smsg = "Finish..." ;
        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
    }
} ;
#*------^ END Function Reset-xoMailboxFolderPermissionsRecursive ^------
