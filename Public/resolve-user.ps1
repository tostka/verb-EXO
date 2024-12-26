#*------v resolve-user.ps1 v------
function resolve-user {
    <#
    .SYNOPSIS
    resolve-user.ps1 - Resolve specified array of -users (displayname, emailaddress, samaccountname) to mail asset, lic & ticket descriptors
    .NOTES
    Version     : 1.0.0.
    Author      : Todd Kadrie
    Website     : http://www.toddomation.com
    Twitter     : @tostka / http://twitter.com/tostka
    CreatedDate : 2021-07-13
    FileName    : resolve-user.ps1
    License     : MIT License
    Copyright   : (c) 2021 Todd Kadrie
    Github      : https://github.com/tostka/verb-XXX
    Tags        : Powershell
    AddedCredit : REFERENCE
    AddedWebsite: URL
    AddedTwitter: URL
    REVISIONS
    * 1:43 PM 12/26/2024 bugfix/cmw uses r: as room dname prefix, not recog'd as dname: #updated: $rgxDName CMW uses : in their room names, so went for broader AD dname support, per AI, and web specs, added 1-256char AD restriction         $rgxDName
        also pushed dname in the detect type switch below samaccountname (which is more specific filter) ; added 'RemoteRoomMailbox' &  'RemoteEquipmentMailbox' switch clauses on typedetails handlers; 
        tweaked lic test to exempt shared/room/equip from isUnlicened warnings.
    * 3:44 PM 12/4/2024 updated to support non-hybrid cloud recipients, w ADC sync'd ADU->AADU; updated enviro_discover etc from latest vers
    * 9:04 AM 11/27/2024 add SharedMbx quota support: flipped logic to pull xomailbox to pull any $hSum.xoRcp|?{$_.recipienttype -eq 'UserMailbox'... (any mailbox type), vs orig: recipienttypedetails, which would only stock UserMailbox details type.
    * 4:40 PM 10/16/2024 added code to do above, users I thot were c1 weren't, had rmbxs, so it needs further testing;  cloud first: VEN,INT,AA,HH, may not match ADU properly, but if they have AADU & AADUser.DirSyncEnabled, the .aaduser.ExtensionProperty.onPremisesDistinguishedName will point to the assoicated ADU! Need to re-resolve when missing ADU
    * 12:50 PM 10/11/2024 substantial rewrites in query code to accomodate apostrophe's in names (selective rewrap " vs ' for queries). Still not great, still doesn't necessarily work searching dname on apostrophe'd names, but it gets through the pass wo crashing (as it did previously).
    * 12:06 PM 9/23/2024 added param for regex to detect non-raw text names; ahdd running $usr input through Remove-StringDiacritic & Remove-StringLatinCharacters() ; 
    * 2:16 PM 6/24/2024: rem'd out #Requires -RunasAdministrator; sec chgs in last x mos wrecked RAA detection
    * 4:28 PM 2/27/2024 updated path-detect code (was discovering into the Mods dir);  updated CBH, quota mbx size, LegalHold example; add additional reporting/detecting to LegalHold status; fixed borked/non-dumping $prpMbxHold = ...@{n="InPlaceHolds";e={ ($_.inplaceholds (*KEY* indicator of a hold in place); updated prompts to echo DiscoveryHolds folder & it's newestItem (both indicate LHs, and if not curr, when it was disabled)
    * 2:51 PM 2/26/2024 add | sort version | select -last 1  on gmos, LF installed 3.4.0 parallel to 3.1.0 and broke auth: caused mult versions to come back and conflict with the assignement of [version] type (would require [version[]] to accom both, and then you get to code everything for mult handling)
    * 12:22 PM 9/26/2023 nesting limit loop, pulled vxo & vx2010  reqs
    * 3:59 PM 9/25/2023 working, ready to drop back into vxo finished in-port of get-xoMailboxQuotaStatus support, now functional, also expanded the mbxstat filter to cover room|shared|Equip recipienttypedetails variants; 
        appears I spliced over $getQuotaUsage support from get-xoMailboxQuotaStatus, looks like it needs to be debugged.
    * 12:43 PM 9/18/2023 re-removed the obsolete xow support: EOM31+ doesn't need it.
    * 3:47 PM 12/14/2022 spliced in xow support. Works on initial pass.
    # 3:57 PM 6/29/2022 fundemental retool for exov2 requirements; pulled all 
        aliasing (wasn't functional for exov2, didn't want to rewrite, and with hard 
        mfa req, exov2 is only way forward, there'll never be verb-EXOnoun use again, 
        due to MS fundemental seizure of the prefix for their 13 'special' cmdlets. 
    # 2:49 PM 3/8/2022 pull Requires -modules ...verb-ex2010 ref - it's generating nested errors, when ex2010 requires exo requires ex2010 == loop.
    * 3:55 PM 2/22/2022 extended the cloud federate test code, to include an INT block (though there's no hybrid to arbitrate, the users are onprem in AD at INT)
    * 12:24 PM 2/1/2022 updated CBH, added a crlf on the console echo (headers weren't lining up); added -getMobile & get-exoMobileDeviceStats support, with conditional md output block; added full aliased xo cmds, implementing full -exov2 support.
    * 2:51 PM 12/27/2021 flipped DN & Desc from md tbl to fl (drops a crlf) ; 
         flipped $propsMailx output to md fmt split lines (condensed output vertically) ; 
         added forward props to propsMailx, and test & echo to tag forwarded mbxs; wrapped $prop* vari's for legibility
    * 11:02 AM 12/13/2021 #11111:had $hsum IsADDisabled, typo: to IsAADDisabled
    * 2:40 PM 12/10/2021 more cleanup ; added $hsum.isDirSynced, for further bulk filter/profiling
        flipped $hsum.isUnlicensed -> Islicensed & added msol.Islicensed test to pop ; 
        appears to work in console - output a stack of filterable objects into collection variable.
        further tweaking and nobrain t-shooting outputs ; added 
        output switches: 
        isNoBrain,isSplitBrain,isUnlicensed,IsDisabledOU,IsADDisabled,IsAADDisabled for 
        postfiltering large collections in bulk, to identify patterns ; reformulated 
        nobrain detec, to have an unlic'd block as well as a licensed - with deadwood 
        offboard nobrains, they'll never have a lic. 
    * 4:19 PM 12/9/2021 improved pipeline support; fixed pipeline param mbinding fails ; added supoort for resolving
        baddomain users or op.mailusers where need to resolve aadu.immutableid to
        aduser, to *ensure* we have a hardmatch of problem objects (resolving baddomain
        DDG-DL-AllDOMAIN recipients to internal NoBrain etc. Still doesn't seem to be
        setting $hsum.NoBrain properly in outputs, but is dropping direct to pipe. May
        have borked single-indiceent xml object dumps tho.
    * 10:30 AM 11/8/2021 fixed CBH/HelpMessage tagging on -outobject
    * 3:30 PM 10/12/2021 added new Name:ObjName_guid support (new hires turn up with aduser named this way); added some marginal multi xoRcp & xoMailbox handling (loops outputs on the above, and the mapiTest), but doesn't do full AzureAD,Msoluser,MailUser,Guest lookups for these. It's really about error-suppression, and notifying the issue more than returning the full picture
    * 1:04 PM 9/28/2021 added:$AADUserManager lookup and dump of UPN, OpDN & mail (for correlating what email pol a user should have -> the one their manager does)
    * 1:52 PM 9/17/2021 moved $props to top ; test enabled/acctenabled, licRecon & mapi test results and use ww on issues ; flipped caad's to -silent (match cmsol 1st echo's to confirm tenant, rest silent); ren $xMProps -> $propsMailx, $XMFedProps-> $propsXMFed, $lProps -> $propsLic,$adprops -> $propsADU, $aaduprops -> $propsAADU, $aaduFedProps -> $propsAADUfed, $RcpPropsTbl -> $propsRcpTbl, $pltgM-> $pltGMailObj, $pltgMU -> $pltgMsoUsr
    * 4:33 PM 9/16/2021 fixed typo in get-AzureAdUser call, reworked output (aadu into markdown delimited wide layout), moved user detaiil reporting to below aadu, and output the federated AD remote DN, (proxied through AADU ext prop)
    * 10:56 AM 9/9/2021 force-resolve xoMailbox, added AADUser pop to the msoluser pop block; added test-xxMapiConnectivity as well; expanded ADU outputs - description, when*, Enabled, to look for terms/recent-hires/disabled accts
    * 3:05 PM 9/3/2021 fixed bugs introduced trying to user MaxResults (msol|aad), which come back param not recog'd when actually used - had to implement as postfiltering to assert open set return limits. ; Also implemented $xxxMeta.rgxOPFederatedDom check to resolve obj primarysmtpaddress to federating AD or AAD.
    * 11:20 AM 8/30/2021 added $MaxResults (shutdown return-all recips in addr space, on failure to match oprcp or xorcp ; fixed a couple of typos; minior testing/logic improvements. Still needs genercized 7pswlt support.
    * 1:30 PM 8/27/2021 new sniggle: CMW user that has EXOP mbx, remote: Added xoMailUser support, failed through DName lookups to try '*lname*' for near-missies. Could add trailing 'lnamne[0-=3]* searches, if not rcp/xrcps found...
    * 9:16 AM 8/18/2021 $xMProps: add email-drivers: CustomAttribute5, EmailAddressPolicyEnabled
    * 12:40 PM 8/17/2021 added -outObject, outputs a full descriptive object for each resolved recipient ; added a $hSum hash and shifted all the varis into mountpoints in the hash, with -outObject, the entire hash is conv'd to an obj and appended to $Rpt ; renamed most of the varis/as objects very clearly for what they are, as sub-props of the output objects. Wo -outobject, the usual comma-delim'd string of addresses is output.
    * 3:26 PM 7/29/2021 had sorta bug (AD context was xxxx:, gadu failing throwing undefined error), but debugging added extensive verbose echos, and an AD-specific try/catch to trap AD notfound errors (notorious, they throw terminating fails, unlike other modules; which crashes out processing even when using -EA continue). So it hardens up the fail recovery process.
    * 12:55 PM 7/19/2021 added guest & exo-mailcontact support (resolving missing ext-federated addresses), retolled logic down to grcp & gxrcp to drive balance of tests.
    * 12:05 PM 7/14/2021 rem'd requires: verb-exo  rem'd requires version 5 (gen'ing 'version' is specified more than once.); rem'd the $rgxSamAcctName, gen's parsing errors compiling into mod ;  added alias 'ulu'; added mailcontact excl on init grcp, to force those to exombx qry ; init vers
    .DESCRIPTION
    resolve-user.ps1 - Resolve specified array of -users (displayname, emailaddress, samaccountname) to mail asset, lic & ticket descriptors

    Key parameter options: 

    -getMobile parameter, to return details on xo MobileDevices in use with the EXO mailbox

    -getQuotaUsage parameter, returns details on xo MailboxFolderStatistics and effective Quota, 
        used with users with mailbox size issues (and/or LegalHold symptoms)

        Note: use of -getQuotaUsage also does an extensive check for LegalHold signs in the mailbox. including reporting on:
            - xoMailbox.LitigationHoldEnabled
            - xoMailbox.InPlaceHolds, 
            - xoMailbox.ComplianceTagHoldApplied
            - xoMailbox.DelayHoldApplied 
            - xoMailbox.DelayReleaseHoldApplied 
            - checks if xoMailboxFolderStats 'DiscoveryHolds' folder has ItemsInFolder -gt 0


    .PARAMETER  users
    Array of user descriptors: displayname, emailaddress, UPN, samaccountname (checks clipboard where unspecified)
    .PARAMETER Ticket
    Ticket number[-ticket 123456]
    .PARAMETER getMobile
    switch to return mobiledevice info for target XO Mailbox (not supported for onprem mailboxes)[-getMobile]
    .PARAMETER getQuotaUsage
    switch to return Quota & MailboxFolderStatistics & LegalHold analysis (XO-only)[-getQuotaUsage]
    .PARAMETER rgxAccentedNameChars
    Regular Expression that identifies input 'user' strings that should ahve diacriticals/latin/non-simple english characters replaced, before lookups
    .PARAMETER useEXOv2
    Use EXOv2 (ExchangeOnlineManagement) over basic auth legacy connection [-useEXOv2]
    .PARAMETER outObject
    switch to return a system.object summary to the pipeline[-outObject]
    .INPUTS
    None. Does not accepted piped input.(.NET types, can add description)
    .OUTPUTS
    System.Object - returns summary report to pipeline
    .EXAMPLE
    PS> resolve-user 
    Default, no user identifier specified: attempts to parse a user descriptor from clipboard contents
    .EXAMPLE
    PS> resolve-user -users 'John Public'
    Process user displayname
    .EXAMPLE
    PS> resolve-user -users 'Test@domain.com','User Name','Alias','ExternalContact@emaildomain.com','confroom@tenant.onmicrosoft.com' -verbose ;
    Process an array of descriptors
    .EXAMPLE
    PS> $results = resolve-user -outobject -users 'Test@domain.com','John Public','Alias','ExternalContact@emaildomain.com','confroom@tenant.onmicrosoft.com''  ;
    PS> $feds = $results| group federator | select -expand name ;
    PS> write-verbose "echo filtered subsets" ;
    PS> ($results| ?{$_.federator -eq $feds[1] }).xomailbox
    PS> ($results| ?{$_.federator -eq $feds[1] }).xomailbox.primarysmtpaddress
    PS> write-verbose "profile results:" ; 
    PS> $analysis = foreach ($data in $resolved_objects){
    PS>     $Rpt = [ordered]@{
    PS>         PrimarySmtpAddress = $data.xorcp.primarysmtpaddress ; 
    PS>         ADUser_UPN = $data.aduser.userprincipalname ; 
    PS>         AADUser_UPN = $data.aaduser.UserPrincipalName ; 
    PS>         isDirSynced = $data.isDirSynced ; 
    PS>         IsNoBrain = $data.IsNoBrain ; 
    PS>         isSplitBrain = $data.isSplitBrain;
    PS>         IsLicensed = $data.IsLicensed;
    PS>         IsDisabledOU = $data.IsDisabledOU;
    PS>         IsADDisabled = $data.IsADDisabled; 
    PS>         IsAADDisabled = $data.IsAADDisabled;
    PS>     } ; 
    PS>     [pscustomobject]$Rpt ; 
    PS> } ; 
    PS> write-verbose "output tabular results" ; 
    PS> $analysis | ft -auto ;
    
    Demo does the following: 
        - Process array of users, specify return detailed object (-outobject), for post-processing & filtering,
        - Group results on federation sources,
        - Output summary of EXO mailboxes for the second federator
        - Then output the primary smtpaddress for all EXO mailboxes resolved to that federator
        - Then create a summary object of the is* properties and UPN, primarySmtpAddress, 
        - Finally display the summary as a console table
    .EXAMPLE
    PS> $rptNNNNNN_FName_LName_Domain_com = ulu -o -users 'FName.LName@Domain.com' ;  $rpt655692_FName_LName_Domain_com | xxml .\logs\rpt655692_FName_LName_Domain_com.xml
    Example (from ahk 7uluo! macro parser output) that creates a variable based on ticketnumber & email address (with underscores for alphanums), from the output, and then exports the variable content to xml. 
    Assigns to an immediately parsable inmem variable, along with the canned .xml that can be reloaded in future, or attached to a ticket.
    .EXAMPLE
    PS> resolve-user -users 'John Public' -getmobile
    Example that includes the -getMobile parameter, to return details on xo MobileDevices in use with the EXO mailbox
    .EXAMPLE
    PS> $999999Rpt = resolve-user fname.lname@toro.com -Ticket 99999 -getQuotaUsage -outObject ; 

        10:39:53:===v (1/1):Input: 'FNAME.LNAME@toro.com' | 'FNAME' | 'LNAME' v===(EML)
        get-Rmbx/xMbx: (Rmbx)(TOR USER, fed:ad.toro.com)
        SamAccountName | WindowsEmailAddress
        LNAMEFI         | FNAME.LNAME@toro.com
        Office | RecipientTypeDetails | RemoteRecipientType | IsDirSynced
                | RemoteUserMailbox    | Migrated            |
        ExternalDirectoryObjectId | CustomAttribute5 | EmailAddressPolicyEnabled
                                    |                  | True
        Outlook (xoMAPI) Access Test Result:Success
        xoMailboxStats Count:1
        10:39:56: INFO:  (-getQuotaUsage:running lengthy Get-xoMailboxFolderStatistics...)
        =get-AADuser FNAME.LNAME@toro.com>:
        =get-AADuserManager FNAME.LNAME@toro.com>:
        ===$hSum.ADUser: 
        UPN                 | DName      | FName | LName | Title                             
        FNAME.LNAME@toro.com | FNAME LNAME | FNAME | LNAME  | Supervisor II, Distribution Center
        Company | Dept                            | Ofc          
                | Operations Distribution El Paso | El Paso-D, TX
        Street | City | State | Zip | Phone           | Mobile
                |      |       |     | +1 915 231 7404 |
        Enabled | DN                                                          
        True    | CN=FNAME LNAME,OU=Users,OU=ELP,DC=global,DC=ad,DC=toro,DC=com
        whenCreated          | whenChanged         
        8/18/2017 4:13:54 PM | 2/23/2024 8:23:33 AM
        Desc : 8/21/17 FT for FNAME LNAME 146294 -bk
        LicenseGroup:(direct-assigned E3)
        ===$hSum.AADUserMgr: 
        UserPrincipalName       | Mail                   
        FNAME.LNAME@toro.com | FNAME.LNAME@toro.com
        OpOU : OU=Users,OU=ELP,DC=global,DC=ad,DC=toro,DC=com
        10:40:06: PROMPT:  UserPrincipalName       | Mail                   
        FNAME.LNAME@toro.com | FNAME.LNAME@toro.com
        OpOU : OU=Users,OU=ELP,DC=global,DC=ad,DC=toro,DC=com

        Licenses::
        MCOEV, FLOW_FREE, MCOPSTNC, ENTERPRISEPACK, POWER_BI_STANDARD, EMS, Microsoft_Teams_Audio_Conferencing_select_dial_out

        which specify the following size limits:
        IssueWarningQuotaMB        : 14336
        ProhibitSendQuotaMB        : 15360
        ProhibitSendReceiveQuotaMB : 17408
        (UseDatabaseQuotaDefaults:False)

        Current TotalMailboxSizeMB: 10912.2

        10:40:06: PROMPT:  QuotaStatus: Mailbox is below mandated SendReceiveQuotaMB by 4,447.80 megabytes
        10:40:06: PROMPT:  
        With the following non-zero folder metrics

        Folder                                                               Items SizeMB  OldestItem NewestItem          FolderType               
        ------                                                               ----- ------  ---------- ----------          ----------               
        annnnnnn-nbne-nnnn-anne-necncannbnnn\Inbox                           23774 5764.38 20230111   02/27/2024 16:37:25 Inbox                    
        annnnnnn-nbne-nnnn-anne-necncannbnnn\Deleted Items                   12434 4599.58 20220323   02/27/2024 16:35:34 DeletedItems             
        ...
        annnnnnn-nbne-nnnn-anne-necncannbnnn\Top of Information Store        1     0                                      Root
        10:40:06: INFO:  
        ===output to::
        D:\scripts\logs\823795-FNAME.LNAME@toro.com-folder-sizes-NONHIDDEN-NONZERO-run20240227-1039AM.xml

        10:40:09: WARNING:  
        10:40:09: WARNING:  
        10:40:09: WARNING:  EVIDENCE OF LEGAL HOLD DETECTED!:
        10:40:09: WARNING:  LitigationHoldEnabled    : False
        10:40:09: WARNING:  
        10:40:09: WARNING:  InPlaceHolds             : UniHnbnednbn-bndn-nnnf-nddn-annndnndnnae, UniHnnnneene-ndnd-naae-annn-nnnnnnnnnncn
        10:40:09: WARNING:  
        10:40:09: WARNING:  ComplianceTagHoldApplied : False
        10:40:09: WARNING:  
        10:40:09: WARNING:  DelayHoldApplied         : False
        10:40:09: WARNING:  
        10:40:09: WARNING:  DelayReleaseHoldApplied  : False
        10:40:09: WARNING:  
        10:40:09: WARNING:  
        10:40:09: WARNING:  Folder          Items    SizeMB OldestItem NewestItem          FolderType                    
        10:40:09: WARNING:  
        10:40:09: WARNING:  ------          -----    ------ ---------- ----------          ----------                    
        10:40:09: WARNING:  
        10:40:09: WARNING:  DiscoveryHolds 267225 101967.69            02/21/2024 08:42:57 RecoverableItemsDiscoveryHolds
        10:40:09: WARNING:  
        10:40:09: WARNING:  
        10:40:09: WARNING:  - DiscoveryHolds folder: If In-Place Hold is enabled or if a Microsoft 365 or Office 365 retention policy is assigned to the mailbox, this subfolder contains all items that meet the hold query parameters and are hard deleted.
        10:40:09: WARNING:  - DiscoveryHolds folder.NewestItem: Will reflect *last time LegalHold captured an item* (e.g. if/when LH was disabled and stopped holding traffic, if in the past)
        10:40:09: WARNING:  
    
    Example that includes the -getQuotaUsage parameter, to return details on xo MailboxFolderStatistics and effective Quota, around users with mailbox size issues, and assigns the returned summary to the variable `$999999Rpt
    Note: use of -getQuotaUsage also does an extensive check for LegalHold signs in the mailbox. including reporting on:
        - xoMailbox.LitigationHoldEnabled
        - xoMailbox.InPlaceHolds, 
        - xoMailbox.ComplianceTagHoldApplied
        - xoMailbox.DelayHoldApplied 
        - xoMailbox.DelayReleaseHoldApplied 
        - checks if xoMailboxFolderStats 'DiscoveryHolds' folder has ItemsInFolder -gt 0
    .LINK
    https://github.com/tostka/verb-exo
    #>

    # 2:49 PM 3/8/2022 pull verb-ex2010 ref - I think it's generating nested errors, when ex2010 requires exo requires ex2010 == loop.
    # 12:19 PM 9/26/2023 pull verb-exo ref "
    #Requires -Modules ActiveDirectory, MSOnline, AzureAD, ExchangeOnlineManagement, verb-AAD, verb-ADMS, verb-Auth, verb-IO, verb-logging
    ##Requires -RunasAdministrator
    # VALIDATORS: [ValidateNotNull()][ValidateNotNullOrEmpty()][ValidateLength(24,25)][ValidateLength(5)][ValidatePattern("(lyn|bcc|spb|adl)ms6(4|5)(0|1).(china|global)\.ad\.DOMAIN\.com")][ValidateSet("USEA","GBMK","AUSYD")][ValidateScript({Test-Path $_ -PathType 'Container'})][ValidateScript({Test-Path $_})][ValidateRange(21,65)][ValidateCount(1,3)]
    ## [OutputType('bool')] # optional specified output type
    [CmdletBinding()]
    [Alias('ulu')]
    PARAM(
        #[Parameter(Position=0,Mandatory=$False,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,HelpMessage="Array of user descriptors: displayname, emailaddress, UPN, samaccountname (checks clipboard where unspecified)")]
        # failing to map pipeline to $users, reduce to Value from Pipeline
        [Parameter(Position=0,Mandatory=$False,ValueFromPipeline=$true,HelpMessage="Array of user descriptors: displayname, emailaddress, UPN, samaccountname (checks clipboard where unspecified)")]
            #[ValidateNotNullOrEmpty()]
            #[Alias('ALIAS1', 'ALIAS2')]
            [array]$users,
        [Parameter(Mandatory=$False,HelpMessage="Ticket Number [-Ticket '999999']")]
            [string]$Ticket,
        [Parameter(HelpMessage="switch to return mobiledevice info for target user[-getMobile]")]
            [switch] $getMobile,
        [Parameter(HelpMessage="switch to return Quota & MailboxFolderStatistics & LegalHold analysis (XO-only)[-getQuotaUsage]")]
            [switch]$getQuotaUsage,
        [Parameter(HelpMessage="Regular Expression that identifies input 'user' strings that should ahve diacriticals/latin/non-simple english characters replaced, before lookups (has default value, used to override for future temp exclusion)[-rgxAccentedNameChars `$rgx]")]
            [ValidateNotNullOrEmpty()]
            [regex]$rgxAccentedNameChars = "[^a-zA-Z0-9\s\.\(\)\{\}\/\&\$\#\@\,\`"\'\’\:\–_-]",
        [Parameter(Mandatory=$FALSE,HelpMessage="TenantTag value, indicating Tenants to connect to[-TenOrg 'TOL']")]
        [ValidateNotNullOrEmpty()]
            #[ValidatePattern("^\w{3}$")]
            [string]$TenOrg = $global:o365_TenOrgDefault,
        [Parameter(Mandatory = $false, HelpMessage = "Use specific Credentials (defaults to Tenant-defined SvcAccount)[-Credentials [credential object]]")]
            [System.Management.Automation.PSCredential]$Credential,
        [Parameter(Mandatory = $false, HelpMessage = "Credential User Role spec (SID|CSID|UID|B2BI|CSVC|ESVC|LSVC|ESvcCBA|CSvcCBA|SIDCBA)[-UserRole @('SIDCBA','SID','CSVC')]")]
            # sourced from get-admincred():#182: $targetRoles = 'SID', 'CSID', 'ESVC','CSVC','UID','ESvcCBA','CSvcCBA','SIDCBA' ;
            #[ValidateSet("SID","CSID","UID","B2BI","CSVC","ESVC","LSVC","ESvcCBA","CSvcCBA","SIDCBA")]
            # pulling the pattern from global vari w friendly err
            [ValidateScript({
                if(-not $rgxPermittedUserRoles){$rgxPermittedUserRoles = '(SID|CSID|UID|B2BI|CSVC|ESVC|LSVC|ESvcCBA|CSvcCBA|SIDCBA)'} ;
                if(-not ($_ -match $rgxPermittedUserRoles)){throw "'$($_)' doesn't match `$rgxPermittedUserRoles:`n$($rgxPermittedUserRoles.tostring())" ; } ;
                return $true ;
            })]
            [string[]]$UserRole =  @('ESvcCBA','CSvcCBA','SIDCBA','SID','CSVC'),
            #@('SID','CSVC'),
            # flip to promptless svcAcct use (SID triggers mauth on phn_, includ failthru sid etc trailing, for admins that don't config cba
            # svcAcct use: @('ESvcCBA','CSvcCBA','SIDCBA')
        [Parameter(HelpMessage="Use EXOv2 (ExchangeOnlineManagement) over basic auth legacy connection [-useEXOv2]")]
            [switch] $useEXOv2=$true,
        [Parameter(HelpMessage="Silent output (suppress status echos)[-silent]")]
            [switch] $silent,
        [Parameter(HelpMessage="switch to return a system.object summary to the pipeline[-outObject]")]
            [switch] $outObject
    ) ;
    BEGIN{
        #region CONSTANTS_AND_ENVIRO #*======v CONSTANTS_AND_ENVIRO v======
        #region ENVIRO_DISCOVER ; #*------v ENVIRO_DISCOVER v------
        $Verbose = ($VerbosePreference -eq 'Continue') ; 
        # Debugger:proxy automatic variables that aren't directly accessible when debugging (must be assigned and read back from another vari) ; 
        $rPSCmdlet = $PSCmdlet ; 
        $rPSScriptRoot = $PSScriptRoot ; 
        $rPSCommandPath = $PSCommandPath ; 
        $rMyInvocation = $MyInvocation ; 
        $rPSBoundParameters = $PSBoundParameters ; 
        [array]$score = @() ; 
        if($rPSCmdlet.MyInvocation.InvocationName){
            if($rPSCmdlet.MyInvocation.InvocationName -match '\.ps1$'){
                $score+= 'ExternalScript' 
            }elseif($rPSCmdlet.MyInvocation.InvocationName  -match '^\.'){
                write-warning "dot-sourced invocation detected!:$($rPSCmdlet.MyInvocation.InvocationName)`n(will be unable to leverage script path etc from MyInvocation objects)" ; 
                # dot sourcing is implicit scripot exec
                $score+= 'ExternalScript' ; 
            } else {$score+= 'Function' };
        } ; 
        if($rPSCmdlet.CommandRuntime){
            if($rPSCmdlet.CommandRuntime.tostring() -match '\.ps1$'){$score+= 'ExternalScript' } else {$score+= 'Function' }
        } ; 
        $score+= $rMyInvocation.MyCommand.commandtype.tostring() ; 
        $grpSrc = $score | group-object -NoElement | sort count ;
        if( ($grpSrc |  measure | select -expand count) -gt 1){
            write-warning  "$score mixed results:$(($grpSrc| ft -a count,name | out-string).trim())" ;
            if($grpSrc[-1].count -eq $grpSrc[-2].count){
                write-warning "Deadlocked non-majority results!" ;
            } else {
                $runSource = $grpSrc | select -last 1 | select -expand name ;
            } ;
        } else {
            write-verbose "consistent results" ;
            $runSource = $grpSrc | select -last 1 | select -expand name ;
        };
        write-verbose  "Calculated `$runSource:$($runSource)" ;
        'score','grpSrc' | get-variable | remove-variable ; # cleanup temp varis
        ${CmdletName} = $rPSCmdlet.MyInvocation.MyCommand.Name ; # function self-name (equiv to script's: $MyInvocation.MyCommand.Path) ;
        #region PsParams ; #*------v PsParams v------
        $PSParameters = New-Object -TypeName PSObject -Property $rPSBoundParameters ;
        write-verbose "`$rPSBoundParameters:`n$(($rPSBoundParameters|out-string).trim())" ;
        # pre psv2, no $rPSBoundParameters autovari to check, so back them out:
        if($rPSCmdlet.MyInvocation.InvocationName){
            if($rPSCmdlet.MyInvocation.InvocationName  -match '^\.'){
                $smsg = "detected dot-sourced invocation: Skipping `$PSCmdlet.MyInvocation.InvocationName-tied cmds..." ; 
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
            } else { 
                write-verbose 'Collect all non-default Params (works back to psv2 w CmdletBinding)'
                $ParamsNonDefault = (Get-Command $rPSCmdlet.MyInvocation.InvocationName).parameters | Select-Object -expand keys | Where-Object{$_ -notmatch '(Verbose|Debug|ErrorAction|WarningAction|ErrorVariable|WarningVariable|OutVariable|OutBuffer)'} ;
            } ; 
        } else { 
            $smsg = "(blank `$rPSCmdlet.MyInvocation.InvocationName, skipping Parameters collection)" ; 
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        } ; 
        #endregion PsParams ; #*------^ END PsParams ^------
        <#
        # Debugger:proxy automatic variables that aren't directly accessible when debugging ; 
        $rPSScriptRoot = $PSScriptRoot ; 
        $rPSCommandPath = $PSCommandPath ; 
        $rMyInvocation = $MyInvocation ; 
        $rPSBoundParameters = $PSBoundParameters ; 
        #>
        $ScriptDir = $scriptName = '' ;     
        if($ScriptDir -eq '' -AND ( (get-variable -name rPSScriptRoot -ea 0) -AND (get-variable -name rPSScriptRoot).value.length)){
            $ScriptDir = $rPSScriptRoot
        } ; # populated rPSScriptRoot
        if( (get-variable -name rPSCommandPath -ea 0) -AND (get-variable -name rPSCommandPath).value.length){
            $ScriptName = $rPSCommandPath
        } ; # populated rPSCommandPath
        if($ScriptDir -eq '' -AND $runSource -eq 'ExternalScript'){$ScriptDir = (Split-Path -Path $rMyInvocation.MyCommand.Source -Parent)} # Running from File
        # when $runSource:'Function', $rMyInvocation.MyCommand.Source is empty,but on functions also tends to pre-hit from the rPSCommandPath entFile.FullPath ;
        if( $scriptname -match '\.psm1$' -AND $runSource -eq 'Function'){
            write-host "MODULE-HOMED FUNCTION:Use `$CmdletName to reference the running function name for transcripts etc (under a .psm1 `$ScriptName will reflect the .psm1 file  fullname)"
            if(-not $CmdletName){write-warning "MODULE-HOMED FUNCTION with BLANK `$CmdletNam:$($CmdletNam)" } ;
        } # Running from .psm1 module
        if($ScriptDir -eq '' -AND (Test-Path variable:psEditor)) {
            write-verbose "Running from VSCode|VS" ; 
            $ScriptDir = (Split-Path -Path $psEditor.GetEditorContext().CurrentFile.Path -Parent) ; 
                if($ScriptName -eq ''){$ScriptName = $psEditor.GetEditorContext().CurrentFile.Path }; 
        } ;
        if ($ScriptDir -eq '' -AND $host.version.major -lt 3 -AND $rMyInvocation.MyCommand.Path.length -gt 0){
            $ScriptDir = $rMyInvocation.MyCommand.Path ; 
            write-verbose "(backrev emulating `$rPSScriptRoot, `$rPSCommandPath)"
            $ScriptName = split-path $rMyInvocation.MyCommand.Path -leaf ;
            $rPSScriptRoot = Split-Path $ScriptName -Parent ;
            $rPSCommandPath = $ScriptName ;
        } ;
        if ($ScriptDir -eq '' -AND $rMyInvocation.MyCommand.Path.length){
            if($ScriptName -eq ''){$ScriptName = $rMyInvocation.MyCommand.Path} ;
            $ScriptDir = $rPSScriptRoot = Split-Path $rMyInvocation.MyCommand.Path -Parent ;
        }
        if ($ScriptDir -eq ''){throw "UNABLE TO POPULATE SCRIPT PATH, EVEN `$rMyInvocation IS BLANK!" } ;
        if($ScriptName){
            if(-not $ScriptDir ){$ScriptDir = Split-Path -Parent $ScriptName} ; 
            $ScriptBaseName = split-path -leaf $ScriptName ;
            $ScriptNameNoExt = [system.io.path]::GetFilenameWithoutExtension($ScriptName) ;
        } ; 
        # blank $cmdlet name comming through, patch it for Scripts:
        if(-not $CmdletName -AND $ScriptBaseName){
            $CmdletName = $ScriptBaseName
        }
        # last ditch patch the values in if you've got a $ScriptName
        if($rPSScriptRoot.Length -ne 0){}else{ 
            if($ScriptName){$rPSScriptRoot = Split-Path $ScriptName -Parent }
            else{ throw "Unpopulated, `$rPSScriptRoot, and no populated `$ScriptName from which to emulate the value!" } ; 
        } ; 
        if($rPSCommandPath.Length -ne 0){}else{ 
            if($ScriptName){$rPSCommandPath = $ScriptName }
            else{ throw "Unpopulated, `$rPSCommandPath, and no populated `$ScriptName from which to emulate the value!" } ; 
        } ; 
        if(-not ($ScriptDir -AND $ScriptBaseName -AND $ScriptNameNoExt  -AND $rPSScriptRoot  -AND $rPSCommandPath )){ 
            throw "Invalid Invocation. Blank `$ScriptDir/`$ScriptBaseName/`ScriptNameNoExt" ; 
            BREAK ; 
        } ; 
        # echo results dyn aligned:
        $tv = 'runSource','CmdletName','ScriptName','ScriptBaseName','ScriptNameNoExt','ScriptDir','PSScriptRoot','PSCommandPath','rPSScriptRoot','rPSCommandPath' ; 
        $tvmx = ($tv| Measure-Object -Maximum -Property Length).Maximum * -1 ; 
        if($silent){}else{
            #$tv | get-variable | %{  write-host -fore yellow ("`${0,$tvmx} : {1}" -f $_.name,$_.value) } ; # w-h
            $tv | get-variable | %{  write-verbose ("`${0,$tvmx} : {1}" -f $_.name,$_.value) } ; # w-v
        }
        'tv','tvmx'|get-variable | remove-variable ; # cleanup temp varis        

        #endregion ENVIRO_DISCOVER ; #*------^ END ENVIRO_DISCOVER ^------

        #region COMMON_CONSTANTS ; #*------v COMMON_CONSTANTS v------
        if(-not $DoRetries){$DoRetries = 4 } ;    # # times to repeat retry attempts
        if(-not $RetrySleep){$RetrySleep = 10 } ; # wait time between retries
        if(-not $RetrySleep){$DawdleWait = 30 } ; # wait time (secs) between dawdle checks
        if(-not $DirSyncInterval){$DirSyncInterval = 30 } ; # AADConnect dirsync interval
        if(-not $ThrottleMs){$ThrottleMs = 50 ;}
        if(-not $rgxDriveBanChars){$rgxDriveBanChars = '[;~/\\\.:]' ; } ; # ;~/\.:,
        if(-not $rgxCertThumbprint){$rgxCertThumbprint = '[0-9a-fA-F]{40}' } ; # if it's a 40char hex string -> cert thumbprint  
        if(-not $rgxSmtpAddr){$rgxSmtpAddr = "^([0-9a-zA-Z]+[-._+&'])*[0-9a-zA-Z]+@([-0-9a-zA-Z]+[.])+[a-zA-Z]{2,63}$" ; } ; # email addr/UPN
        if(-not $rgxDomainLogon){$rgxDomainLogon = '^[a-zA-Z][a-zA-Z0-9\-\.]{0,61}[a-zA-Z]\\\w[\w\.\- ]+$' } ; # DOMAIN\samaccountname 
        if(-not $exoMbxGraceDays){$exoMbxGraceDays = 30} ; 
        if(-not $XOConnectionUri ){$XOConnectionUri = 'https://outlook.office365.com'} ; 
        if(-not $SCConnectionUri){$SCConnectionUri = 'https://ps.compliance.protection.outlook.com'} ; 

        write-verbose "Coerce configured but blank Resultsize to Unlimited" ; 
        if(get-variable -name resultsize -ea 0){
            if( ($null -eq $ResultSize) -OR ('' -eq $ResultSize) ){$ResultSize = 'unlimited' }
            elseif($Resultsize -is [int]){} else {throw "Resultsize must be an integer or the string 'unlimited' (or blank)"} ;
        } ; 
        #$ComputerName = $env:COMPUTERNAME ;
        #$NoProf = [bool]([Environment]::GetCommandLineArgs() -like '-noprofile'); # if($NoProf){# do this};
        # XXXMeta derived constants:
        # - AADU Licensing group checks
        # calc the rgxLicGrpName fr the existing $xxxmeta.rgxLicGrpDN: (get-variable tormeta).value.rgxLicGrpDN.split(',')[0].replace('^','').replace('CN=','')
        #$rgxLicGrpName = (get-variable -name "$($tenorg)meta").value.rgxLicGrpDN.split(',')[0].replace('^','').replace('CN=','')
        # use the dn vers LicGrouppDN = $null ; # | ?{$_ -match $tormeta.rgxLicGrpDN}
        #$rgxLicGrpDN = (get-variable -name "$($tenorg)meta").value.rgxLicGrpDN
        # email trigger vari, it will be semi-delimd list of mail-triggering events
        $script:PassStatus = $null ;
        # TenOrg or other looped-specific PassStatus (auto supported by 7pswlt)
        #New-Variable -Name PassStatus_$($tenorg) -scope Script -Value $null ;
        [array]$SmtpAttachment = $null ;
        #write-verbose "start-Timer:Master" ; 
        $swM = [Diagnostics.Stopwatch]::StartNew() ;
        #endregion COMMON_CONSTANTS ; #*------^ END COMMON_CONSTANTS ^------
        #region LOCAL_CONSTANTS ; #*------v LOCAL_CONSTANTS v------
        $rgxEmailAddr = "^([0-9a-zA-Z]+[-._+&'])*[0-9a-zA-Z]+@([-0-9a-zA-Z]+[.])+[a-zA-Z]{2,63}$" ;
        # added support for . fname lname delimiter (supports pasted in dirname of email addresses, as user)
        $rgxDName = "^([a-zA-Z]{2,}(\s|\.)[a-zA-Z]{1,}'?-?[a-zA-Z]{2,}\s?([a-zA-Z]{1,})?)" ;
        #updated: CMW uses : in their room names, so went for broader AD dname support, per AI, and web specs, added 1-256char AD restriction
        $rgxDName ="[a-zA-Z0-9\s$([Regex]::Escape('/\[:;|=,+*?<>') + '\]' + '\"')]{1,256}" ; 
        #"^([a-zA-Z]{2,}\s[a-zA-Z]{1,}'?-?[a-zA-Z]{2,}\s?([a-zA-Z]{1,})?)" ;
        $rgxObjNameNewHires = "^([a-zA-Z]{2,}(\s|\.)[a-zA-Z]{1,}'?-?[a-zA-Z]{2,}\s?([a-zA-Z]{1,})?)_[a-z0-9]{10}"  # Name:Fname LName_f4feebafdb (appending uniqueness guid chunk)
        $rgxSamAcctNameTOR = "^\w{2,20}$" ; # up to 20k, the limit prior to win2k
        #$rgxSamAcctName = "^[^\/\\\[\]:;|=,+?<>@?]+$" # no char limit ;
        $MaxRecips = 25 ; # max number of objects to permit on a return resultsize/,ResultSetSize, to prevent empty set return of everything in the addressspace
        #$rgxADDistNameGAT = ",DC=global,DC=ad,DC=toro,DC=com" ; 
        $rgxADDistNameAT = ",DC=ad,DC=toro,DC=com" ; 
        #$DNDOM = @() ; 
        #$TORMeta.adforestname.split('.') | %{$dndom += "DC=$($_)"} ;
        #$rgxADDistNameAT = [regex]::Escape($DNDOM -join ',') ; 

        # props dyn filtering: write-host "=get-xMbx:>`n$(($hSum.xoMailbox |fl ($xMprops |?{$_ -notmatch '(sam.*|dist.*)'})|out-string).trim())`n-Title:$($hSum.xoUser.Title)";
        # $propsMailx: add email-drivers: CustomAttribute5, EmailAddressPolicyEnabled
        # 11:01 AM 12/27/2021 add forwarding settings (critical to bounce/block tracking for RM)
        #$propsMailx='samaccountname','windowsemailaddress','DistinguishedName','Office','RecipientTypeDetails','RemoteRecipientType','IsDirSynced','ExternalDirectoryObjectId','CustomAttribute5','EmailAddressPolicyEnabled' ;
        $propsMailx='samaccountname','windowsemailaddress','DistinguishedName','Office','RecipientTypeDetails','RemoteRecipientType',
            'IsDirSynced','ExternalDirectoryObjectId','CustomAttribute5','EmailAddressPolicyEnabled',
            'DeliverToMailboxAndForward','DeliverToMailboxAndForward','ForwardingSmtpAddress' ;
        # pulls: 'ImmutableId',
        # 1:41 PM 12/27/2021 add multiline md tbl output
        $propsMailxL1 = 'SamAccountName','WindowsEmailAddress' ; 
        $propsMailxL2 = 'Office','RecipientTypeDetails','RemoteRecipientType', 'IsDirSynced' ;
        $propsMailxL3 = 'ExternalDirectoryObjectId','CustomAttribute5','EmailAddressPolicyEnabled' ; 
        $propsMailxL4 = 'DistinguishedName' ; 
        $propsMailxL5 = 'ForwardingAddress','ForwardingSmtpAddress','DeliverToMailboxAndForward' ;        
        $propsXMFed = 'samaccountname','windowsemailaddress','DistinguishedName','Office','RecipientTypeDetails','RemoteRecipientType',
            'ImmutableId','ExternalDirectoryObjectId','CustomAttribute5','EmailAddressPolicyEnabled' ;
        $propsLic = @{Name='HasLic'; Expression={$_.IsLicensed }},@{Name='LicIssue'; Expression={$_.LicenseReconciliationNeeded }} ;
        $propsADU = 'UserPrincipalName','DisplayName','GivenName','Surname','Title','Company','Department','PhysicalDeliveryOfficeName',
            'StreetAddress','City','State','PostalCode','TelephoneNumber','MobilePhone','Enabled','DistinguishedName',
            'Description','whenCreated','whenChanged'
        #'samaccountname','UserPrincipalName','distinguishedname','Description','title','whenCreated','whenChanged','Enabled','sAMAccountType','userAccountControl' ;
        $propsADUsht = 'Enabled','Description','whenCreated','whenChanged','Title' ;
        $propsAADU = 'UserPrincipalName','DisplayName','GivenName','Surname','Title','Company','Department','PhysicalDeliveryOfficeName',
            'StreetAddress','City','State','PostalCode','TelephoneNumber','MobilePhone','Enabled','DistinguishedName' ;
        #'UserPrincipalName','name','ImmutableId','DirSyncEnabled','LastDirSyncTime','AccountEnabled' ;
        # 3:59 PM 10/9/2024 used for complete miss gadu search results props
        $prpADU = 'DistinguishedName','GivenName','Surname','Name','UserPrincipalName','mailNickname','SamAccountName','physicalDeliveryOfficeName','msExchRecipientDisplayType','msExchRecipientTypeDetails','msExchRemoteRecipientType','msExchWhenMailboxCreated' ; 
        $propsAADUfed = 'UserPrincipalName','name','ImmutableId','DirSyncEnabled','LastDirSyncTime' ;
        $propsRcpTbl = 'Alias','PrimarySmtpAddress','RecipientType','RecipientTypeDetails' ;
        # line1-X AADU outputs
            #$propsMailx='samaccountname','windowsemailaddress','DistinguishedName','Office','RecipientTypeDetails','RemoteRecipientType','IsDirSynced','ImmutableId','ExternalDirectoryObjectId','CustomAttribute5','EmailAddressPolicyEnabled' ;
        <# full size
        $propsADL1 = 'UserPrincipalName','DisplayName','GivenName','Surname','Title' ;
        $propsADL2 = 'Company','Department','PhysicalDeliveryOfficeName' ;
        $propsADL3 = 'StreetAddress','City','State','PostalCode','TelephoneNumber','MobilePhone' ;
        # non-ADU props
        #$propsADL4 = 'DirSyncEnabled','ImmutableId','LastDirSyncTime','UsageLocation' ;
        #$propsADL5 = 'ObjectType','UserType' ;
        #>
        # abbreviated:
        $propsADL1 = @{Name='UPN';Expression={$_.UserPrincipalName }}, @{Name='DName';Expression={$_.DisplayName }}, 
            @{Name='FName';Expression={$_.GivenName }},@{Name='LName';Expression={$_.Surname }},
            @{Name='Title';Expression={$_.Title }};
        $propsADL2 = @{Name='Company';Expression={$_.Company }},@{Name='Dept';Expression={$_.Department }},
            @{Name='Ofc';Expression={$_.PhysicalDeliveryOfficeName }} ;
        $propsADL3 = @{Name='Street';Expression={$_.StreetAddress }}, 'City','State',
            @{Name='Zip';Expression={$_.PostalCode }}, @{Name='Phone';Expression={$_.TelephoneNumber }}, 
            @{Name='Mobile';Expression={$_.MobilePhone }} ;
        $propsADL4 = 'Enabled',@{Name='DN';Expression={$_.DistinguishedName }} ;
        #$propsADL4 = @{Name='Dsync';Expression={$_.DirSyncEnabled }}, @{Name='ImutID';Expression={$_.ImmutableId }}, @{Name='LastDSync';Expression={$_.LastDirSyncTime }}, @{Name='UseLoc';Expression={$_.UsageLocation }};
        #$propsADL5 = 'ObjectType','UserType' ;
        $propsADL5 = 'whenCreated','whenChanged' ; 
        $propsADL6 = @{Name='Desc';Expression={$_.Description }} ;

        # line1-5 AADU outputs
        <# full size
        $propsAADL1 = 'UserPrincipalName','DisplayName','GivenName','Surname','JobTitle' ;
        $propsAADL2 = 'CompanyName','Department','PhysicalDeliveryOfficeName' ;
        $propsAADL3 = 'StreetAddress','City','State','PostalCode','TelephoneNumber','Mobile' ;
        $propsAADL4 = 'DirSyncEnabled','ImmutableId','LastDirSyncTime','UsageLocation' ;
        $propsAADL5 = 'ObjectType','UserType' ;
        #>
        # abbreviated:
        $propsAADL1 = @{Name='UPN';Expression={$_.UserPrincipalName }}, @{Name='DName';Expression={$_.DisplayName }}, 
            @{Name='FName';Expression={$_.GivenName }},@{Name='LName';Expression={$_.Surname }},
            @{Name='Title';Expression={$_.JobTitle }};
        $propsAADL2 = @{Name='Company';Expression={$_.CompanyName }},@{Name='Dept';Expression={$_.Department }},
            @{Name='Ofc';Expression={$_.PhysicalDeliveryOfficeName }} ;
        $propsAADL3 = @{Name='Street';Expression={$_.StreetAddress }}, 'City','State',
            @{Name='Zip';Expression={$_.PostalCode }}, @{Name='Phone';Expression={$_.TelephoneNumber }}, 'Mobile' ;
        $propsAADL4 = @{Name='Dsync';Expression={$_.DirSyncEnabled }}, @{Name='ImutID';Expression={$_.ImmutableId }}, 
            @{Name='LastDSync';Expression={$_.LastDirSyncTime }}, @{Name='UseLoc';Expression={$_.UsageLocation }};
        $propsAADL5 = 'ObjectType','UserType', @{Name='Enabled';Expression={$_.AccountEnabled }} ;

        #$propsAADMgr = 'UserPrincipalName','Mail',@{Name='OpDN';Expression={$_.ExtensionProperty.onPremisesDistinguishedName }} ;
        # get mgr OU, not DN: ExtensionProperty.onPremisesDistinguishedName.split(',') | select -skip 1 ) -join ','
        $propsAADMgr = 'UserPrincipalName','Mail',
            @{Name='OpOU';Expression={($_.ExtensionProperty.onPremisesDistinguishedName.split(',') | select -skip 1) -join ',' }} ;
        $propsAADMgrL1 = 'UserPrincipalName','Mail' ;
        $propsAADMgrL2 = @{Name='OpOU';Expression={($_.ExtensionProperty.onPremisesDistinguishedName.split(',') | select -skip 1) -join ',' }} ;
        $sQot = [char]34 ;
        $sQotS = [char]39 ;

        # 2:42 PM 10/9/2024 add prp for multi-recipient match ft -a dumps
        #$prpFTARcp = 'Name','RecipientTypeDetails','RecipientType','PrimarySmtpAddress','alias' ; 

        if($getMobile){
            # mobile device props
            #$MDtbl=[ordered]@{NoDashRow=$true } ; # out-markdowntable splat
            #$propsMobDevStats = 'DeviceFriendlyName','DeviceType','DeviceOS','ClientType','DeviceID',
            #    'FirstSyncTime','LastSyncAttemptTime','LastSuccessSync','NumberOfFoldersSynced' ; 
            $propsMobL1 = @{Name='FriendlyName';Expression={$_.DeviceFriendlyName }},@{Name='DevType';Expression={$_.DeviceType }},
                @{Name='DevOs';Expression={$_.DeviceOS }},@{Name='ClntType';Expression={$_.ClientType }},
                @{Name='DevID';Expression={$_.DeviceID }} ; 
            # shorten times: (get-date '6/20/2021 1:45:34 AM' -format 'M/d/yy H:mmtt');
            $propsMobL2 = @{Name='1stSyncTime';Expression={(get-date $_.FirstSyncTime -format 'M/d/yy H:mmtt') }},
                @{Name='LastSyncTime';Expression={(get-date $_.LastSyncAttemptTime -format 'M/d/yy H:mmtt') }},
                @{Name='LastSuccSync';Expression={(get-date $_.LastSuccessSync -format 'M/d/yy H:mmtt') }},
                @{Name='#Folders';Expression={$_.NumberOfFoldersSynced }} ; 
        } ; 
        if($getQuotaUsage){

            # 12:54 PM 9/18/2023 adds for MbxFolderStats, Quota & LegalHold eval:
            $prpStat = 'DisplayName',@{n="DBIssueWarningQuotaMB";e={[math]::round($_.DatabaseIssueWarningQuota.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1MB,2)}},
                @{n="DBProhibitSendQuotaMB";e={[math]::round($_.DatabaseProhibitSendQuota.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1MB,2)}},
                @{n="DBProhibitSendReceiveQuotaMB";e={[math]::round($_.DatabaseProhibitSendReceiveQuota.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1MB,2)}},
                @{n="TotalMailboxSizeMB";e={[math]::round($_.TotalItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1MB,2)}},
                'LastLogonTime' ,'LastLogoffTime' ;

            $prpFldr = @{Name='Folder'; Expression={$_.Identity.tostring()}},@{Name="Items"; Expression={$_.ItemsInFolder}}, 
                @{n="SizeMB"; e={[math]::round($_.FolderSize.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1MB,2)}}, 
                @{Name="OldestItem"; Expression={get-date $_.OldestItemReceivedDate -f "yyyyMMdd"}}, 
                @{Name="NewestItem"; Expression={$_.NewestItemReceivedDate -f "yyyyMMdd"}},"FolderType" ;

            # 10:01 AM 2/27/2024 new spec for reporting on LegalHold symptom folders
            $prpFldrLH = @{Name='Folder'; Expression={$_.Name.tostring()}},@{Name="Items"; Expression={$_.ItemsInFolder}}, 
                @{n="SizeMB"; e={[math]::round($_.FolderSize.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1MB,2)}}, 
                @{Name="OldestItem"; Expression={get-date $_.OldestItemReceivedDate -f "yyyyMMdd"}}, 
                @{Name="NewestItem"; Expression={$_.NewestItemReceivedDate -f "yyyyMMdd"}},"FolderType" ;
                
            # 9:41 AM 2/27/2024 fixed borked InPlaceHolds expansion (was empty, and the prop is where JanelS holds actually *appear*)
$prpMbxHold = 'LitigationHoldEnabled',@{n="InPlaceHolds";e={ ($_.inplaceholds ) -join ', '}},
                'ComplianceTagHoldApplied','DelayHoldApplied','DelayReleaseHoldApplied' ; 

            $rgxHiddn = '.*\\(Versions|SubstrateHolds|DiscoveryHolds|Yammer.*|Social\sActivity\sNotifications|Suggested\sContacts|Recipient\sCache|PersonMetadata|Audits|Calendar\sLogging|Purges)$' ; 

        } ; 
        $rgxOPLic = '^CN\=ENT\-APP\-Office365\-(EXOK|F1|MF1)-DL$' ;
        $rgxXLic = '^CN\=ENT\-APP\-Office365\-(EXOK|F1|MF1)-DL$' ;
        #endregion LOCAL_CONSTANTS ; #*------^ END LOCAL_CONSTANTS ^------        
         # ENCODED CONsTANTS & SUPPORT FUNCTIONS:
        #region 2B4 ; #*------v 2B4 v------
        if(-not (get-command 2b4 -ea 0)){function 2b4{[CmdletBinding()][Alias('convertTo-Base64String')] PARAM([Parameter(ValueFromPipeline=$true)][string[]]$str) ; PROCESS{$str|%{[Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($_))}  };} ; } ; 
        #endregion 2B4 ; #*------^ END 2B4 ^------
        #region 2B4C ; #*------v 2B4C v------
        # comma-quoted return
        if(-not (get-command 2b4c -ea 0)){function 2b4c{ [CmdletBinding()][Alias('convertto-Base64StringCommaQuoted')] PARAM([Parameter(ValueFromPipeline=$true)][string[]]$str) ;BEGIN{$outs = @()} PROCESS{[array]$outs += $str | %{[Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($_))} ; } END {'"' + $(($outs) -join '","') + '"' | out-string | set-clipboard } ; } ; } ; 
        #endregion 2B4C ; #*------^ END 2B4C ^------
        #region FB4 ; #*------v FB4 v------
        # DEMO: $SitesNameList = 'THluZGFsZQ==','U3BlbGxicm9vaw==','QWRlbGFpZGU=' | fb4 ;
        if(-not (get-command fb4 -ea 0)){function fb4{[CmdletBinding()][Alias('convertFrom-Base64String')] PARAM([Parameter(ValueFromPipeline=$true)][string[]]$str) ; PROCESS{$str | %{ [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($_)) }; } ; } ; }; 
        #endregion FB4 ; #*------^ END FB4 ^------
        #endregion ENCODED_CONTANTS ; #*------^ END ENCODED_CONTANTS ^------

        #region FUNCTIONS ; #*======v FUNCTIONS v======

        #endregion FUNCTIONS ; #*======^ END FUNCTIONS ^======

        # 3:19 PM 6/23/2022: for exo2, this is going to have to be rearranged, if not shifted into smarter cxo2.
        <#rx10 -Verbose:$false ;
        rxo  -Verbose:$false ;
        cmsol  -Verbose:$false ;
        #>
        <#dx10 ; 
        rxo2 ; 
        rx10 ; 
        caad ;
        #>

        
        #region SERVICE_CONNECTIONS #*======v SERVICE_CONNECTIONS v======
        # PRETUNE STEERING separately *before* pasting in balance of region
        # THIS BLOCK DEPS ON VERB-* FANCY CRED/AUTH HANDLING MODULES THAT *MUST* BE INSTALLED LOCALLY TO FUNCTION
        # NOTE: *DOES* INCLUDE *PARTIAL* DEP-LESS $useExopNoDep=$true OPT THAT LEVERAGES Connect-ExchangeServerTDO, VS connect-ex2010 & CREDS ARE ASSUMED INHERENT TO THE ACCOUNT) 
        # Connect-ExchangeServerTDO HAS SUBSTANTIAL BENEFIT, OF WORKING SEAMLESSLY ON EDGE SERVER AND RANGE OF DOMAIN-=CONNECTED EXOP ROLES
        #*------v STEERING VARIS v------
        $useO365 = $true ;
        $useEXO = $true ; 
        $UseOP=$true ; 
        $UseExOP=$true ;
        $useExopNoDep = $false ; # switch to use Connect-ExchangeServerTDO, vs connect-ex2010 (creds are assumed inherent to the account)
        $useForestWide = $true ; # flag to trigger cross-domain/forest-wide code in AD & EXoP
        $UseOPAD = $true ; 
        $UseMSOL = $false ; # should be hard disabled now in o365
        $UseAAD = $true  ; 
        $useO365 = [boolean]($useO365 -OR $useEXO -OR $UseMSOL -OR $UseAAD)
        $UseOP = [boolean]($UseOP -OR $UseExOP -OR $UseOPAD) ;
        #*------^ END STEERING VARIS ^------
        #*------v EXO V2/3 steering constants v------
        $EOMModName =  'ExchangeOnlineManagement' ;
        $EOMMinNoWinRMVersion = $MinNoWinRMVersion = '3.0.0' ; # support both names
        #*------^ END EXO V2/3 steering constants ^------
        # assert Org from Credential specs (if not param'd)
        # 1:36 PM 7/7/2023 and revised again -  revised the -AND, for both, logic wasn't working
        if($TenOrg){    
            $smsg = "Confirmed populated `$TenOrg" ; 
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        } elseif(-not($tenOrg) -and $Credential){
            $smsg = "(unconfigured `$TenOrg: asserting from credential)" ; 
            if($silent){}elseif ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Prompt } 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            if((get-command get-TenantTag).Parameters.keys -contains 'silent'){
                $TenOrg = get-TenantTag -Credential $Credential -silent ;;
            }else {
                $TenOrg = get-TenantTag -Credential $Credential ;
            }
        } else { 
            # if not using Credentials or a TargetTenants/TenOrg loop, default the $TenOrg on the $env:USERDOMAIN
            $smsg = "(unconfigured `$TenOrg & *NO* `$Credential: fallback asserting from `$env:USERDOMAIN)" ; 
            if($silent){}elseif ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Prompt } 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            switch -regex ($env:USERDOMAIN){
                ([regex]('(' + (( @($TORMeta.legacyDomain,$CMWMeta.legacyDomain)  |foreach-object{[regex]::escape($_)}) -join '|') + ')')).tostring() {$TenOrg = $env:USERDOMAIN.substring(0,3).toupper() } ;
                $TOLMeta.legacyDomain {$TenOrg = 'TOL' }
                default {throw "UNRECOGNIZED `$env:USERDOMAIN!:$($env:USERDOMAIN)" ; exit ; } ;
            } ; 
        } ; 
        #region useO365 ; #*------v useO365 v------
        #$useO365 = $false ; # non-dyn setting, drives variant EXO reconnect & query code
        #if($CloudFirst){ $useO365 = $true } ; # expl: steering on a parameter
        if($useO365){
            #region GENERIC_EXO_CREDS_&_SVC_CONN #*------v GENERIC EXO CREDS & SVC CONN BP v------
            # o365/EXO creds
            <### Usage: Type defaults to SID, if not spec'd - Note: there must be a *logged in & configured *profile*
            $o365Cred = get-TenantCredentials -TenOrg $TenOrg -verbose -userrole SID ;
            Returns a credential set for the $TenOrg Hybrid OnPrem Exchange Org
            .EXAMPLE
            $o365Cred = get-TenantCredentials -TenOrg $TenOrg -verbose -userrole CSVC ;
            Returns the CSVC Userrole credential for the $TenOrg Hybrid OnPrem Exchange Org
            .EXAMPLE
            $o365Cred = get-TenantCredentials -TenOrg $TenOrg -verbose -userrole B2BI ;
            Returns the B2BI Userrole credential for the $TenOrg Hybrid OnPrem Exchange Org
            ###>
            $o365Cred = $null ;
            if($Credential){
                $smsg = "`Credential:Explicit credentials specified, deferring to use..." ; 
                if($silent){}elseif ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                # get-TenantCredentials() return format: (emulating)
                $o365Cred = [ordered]@{
                    Cred=$Credential ; 
                    credType=$null ; 
                } ; 
                $uRoleReturn = resolve-UserNameToUserRole -UserName $Credential.username -verbose:$($VerbosePreference -eq "Continue") ; # Username
                #$uRoleReturn = resolve-UserNameToUserRole -Credential $Credential -verbose = $($VerbosePreference -eq "Continue") ;   # full Credential support
                if($uRoleReturn.UserRole){
                    $o365Cred.credType = $uRoleReturn.UserRole ; 
                } else { 
                    $smsg = "Unable to resolve `$credential.username ($($credential.username))"
                    $smsg += "`nto a usable 'UserRole' spec!" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    throw $smsg ;
                    Break ;
                } ; 
            } else { 
                $pltGTCred=@{TenOrg=$TenOrg ; UserRole=$null; verbose=$($verbose)} ;
                if($UserRole){
                    $smsg = "(`$UserRole specified:$($UserRole -join ','))" ; 
                    if($silent){}elseif ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    $pltGTCred.UserRole = $UserRole; 
                } else { 
                    $smsg = "(No `$UserRole found, defaulting to:'CSVC','SID' " ; 
                    if($silent){}elseif ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                    $pltGTCred.UserRole = 'CSVC','SID' ; 
                } ; 
                $smsg = "get-TenantCredentials w`n$(($pltGTCred|out-string).trim())" ; 
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level verbose } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                $o365Cred = get-TenantCredentials @pltGTCred
            } ; 
            if($o365Cred.credType -AND $o365Cred.Cred -AND $o365Cred.Cred.gettype().fullname -eq 'System.Management.Automation.PSCredential'){
                $smsg = "(validated `$o365Cred contains .credType:$($o365Cred.credType) & `$o365Cred.Cred.username:$($o365Cred.Cred.username)" ;
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                # 9:58 AM 6/13/2024 populate $credential with return, if not populated (may be required for follow-on calls that pass common $Credentials through)
                if((gv Credential) -AND $Credential -eq $null){
                    $credential = $o365Cred.Cred ;
                }elseif($credential.gettype().fullname -eq 'System.Management.Automation.PSCredential'){
                    $smsg = "(`$Credential is properly populated; explicit -Credential was in initial call)" ; 
                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                } else {
                    $smsg = "`$Credential is `$NULL, AND $o365Cred.Cred is unusable to populate!" ;
                    $smsg = "downstream commands will *not* properly pass through usable credentials!" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    throw $smsg ;
                    break ;
                } ;
            } else {
                $smsg = "UNABLE TO RESOLVE FUNCTIONAL CredType/UserRole from specified explicit -Credential:$($Credential.username)!" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                break ;
            } ; 
            if($o365Cred){
                # make it script scope, so we don't have to predetect & purge before using new-variable
                if(get-Variable -Name cred$($tenorg) -scope Script -ea 0 ){ remove-Variable -Name cred$($tenorg) -scope Script } ;
                New-Variable -Name cred$($tenorg) -scope Script -Value $o365Cred.cred ;
                $smsg = "Resolved $($Tenorg) `$o365cred:$($o365Cred.cred.username) (assigned to `$cred$($tenorg))" ;
                if($silent){}elseif ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } else {
                $statusdelta = ";ERROR"; # CHANGE|INCOMPLETE|ERROR|WARN|FAIL ;
                $script:PassStatus += $statusdelta ;
                set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatettus_$($tenorg)).value + $statusdelta) ;
                $smsg = "Unable to resolve $($tenorg) `$o365Cred value!"
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                throw "Unable to resolve $($tenorg) `$o365Cred value!`nEXIT!"
                Break ;
            } ;
            # if we get here, wo a $Credential, w resolved $o365Cred, assign it 
            if(-not $Credential -AND $o365Cred){$Credential = $o365Cred.cred } ; 
            # configure splat for connections: (see above useage)
            # downstream commands
            $pltRXO = [ordered]@{
                Credential = $Credential ;
                verbose = $($VerbosePreference -eq "Continue")  ;
            } ;
            if((get-command Reconnect-EXO).Parameters.keys -contains 'silent'){
                $pltRxo.add('Silent',$silent) ;
            } ;
            # default connectivity cmds - force silent 
            $pltRXOC = [ordered]@{} ; $pltRXO.GetEnumerator() | ?{ $_.Key -notmatch 'silent' }  | ForEach-Object { $pltRXOC.Add($_.Key, $_.Value) } ; $pltRXOC.Add('silent',$silent) ; 
            if((get-command Reconnect-EXO).Parameters.keys -notcontains 'silent'){
                $pltRxo.remove('Silent') ;
            } ; 
            #region EOMREV ; #*------v EOMREV Check v------
            #$EOMmodname = 'ExchangeOnlineManagement' ;
            $pltIMod = @{Name = $EOMmodname ; ErrorAction = 'Stop' ; verbose=$false} ;
            # do a gmo first, faster than gmo -list
            if([version]$EOMMv = (Get-Module @pltIMod).version){}
            elseif([version]$EOMMv = (Get-Module -ListAvailable @pltIMod).version){}
            else {
                $smsg = "$($EOMmodname) PowerShell v$($MinNoWinRMVersion) module is required, do you want to install it?" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Prompt }
                else{ $smsg = "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                $bRet = Read-Host "Enter YYY to continue. Anything else will exit"  ;
                if ($bRet.ToUpper() -eq "YYY") {
                    $smsg = "Installing $($EOMmodname) module..." ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    Install-Module $EOMmodname -Repository PSGallery -AllowClobber -Force ;
                } else {
                    $smsg = "Please install $($EOMmodname) PowerShell v$($MinNoWinRMVersion)  module." ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN }
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    #exit 1
                    break ;
                }  ;
            } ;
            $smsg = "(Checking for WinRM support in this EOM rev...)" ; 
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            if([version]$EOMMv -ge [version]$MinNoWinRMVersion){
                $MinNoWinRMVersion = $EOMMv.tostring() ;
                $IsNoWinRM = $true ;
            }elseif([version]$EOMMv -lt [version]$MinimumVersion){
                $smsg = "Installed $($EOMmodname) is v$($MinNoWinRMVersion): This module is obsolete!" ;
                $smsg += "`nAnd unsupported by this function!" ;
                $smsg += "`nPlease install $($EOMmodname) PowerShell v$($MinNoWinRMVersion)  module!" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                Break ;
            } else {
                $IsNoWinRM = $false ;
            } ;
            [boolean]$UseConnEXO = [boolean]([version]$EOMMv -ge [version]$MinNoWinRMVersion) ;
            #endregion EOMREV ; #*------^ END EOMREV Check  ^------
            #-=-=-=-=-=-=-=-=
            <### CALLS ARE IN FORM: (cred$($tenorg))
            # downstream commands
            $pltRXO = @{
                Credential = (Get-Variable -name cred$($tenorg) ).value ;
                #verbose = $($verbose) ;
                Verbose = $FALSE ; 
            } ;
            if((get-command Reconnect-EXO).Parameters.keys -contains 'silent'){
                $pltRxo.add('Silent',$false) ;
            } ; 
            # default connectivity cmds - force silent false
            $pltRXOC = [ordered]@{} ; $pltRXO.GetEnumerator() | ?{ $_.Key -notmatch 'silent' }  | ForEach-Object { $pltRXOC.Add($_.Key, $_.Value) } ; $pltRXOC.Add('silent',$true) ;
            if((get-command Reconnect-EXO).Parameters.keys -notcontains 'silent'){
                $pltRxo.remove('Silent') ;
            } ; 
            #$pltRXO creds & .username can also be used for AzureAD connections:
            #Connect-AAD @pltRXOC ;
            ###>
            #endregion GENERIC_EXO_CREDS_&_SVC_CONN #*------^ END GENERIC EXO CREDS & SVC CONN BP ^------

        } else {
            $smsg = "(`$useO365:$($useO365))" ; 
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        } ; # if-E if($useO365 ){
        #endregion useO365 ; #*------^ END useO365 ^------

        #region useEXO ; #*------v useEXO v------
        # 1:29 PM 9/15/2022 as of MFA & v205, have to load EXO *before* any EXOP, or gen get-steppablepipeline suffix conflict error
        if($useEXO){
            if ($script:useEXOv2 -OR $useEXOv2) { reconnect-eXO2 @pltRXOC }
            else { reconnect-EXO @pltRXOC } ;
        } else {
            $smsg = "(`$useEXO:$($useEXO))" ; 
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        } ; # if-E 
        #endregion  ; #*------^ END useEXO ^------
  
        #region GENERIC_EXOP_CREDS_&_SRVR_CONN #*------v GENERIC EXOP CREDS & SRVR CONN BP v------
        # steer all onprem code on $XXXMeta.ExOPAccessFromToro & Ex10Server values
        #$UseOP=$true ; 
        #$UseExOP=$true ;
        #$useExopNoDep = $true # switch to use Connect-ExchangeServerTDO, vs connect-ex2010 (creds are assumed inherent to the account) 
        #$useForestWide = $true ; # flag to trigger cross-domain/forest-wide code in AD & EXoP
        <# no onprem dep
        if((Get-Variable  -name "$($TenOrg)Meta").value.ExOPAccessFromToro -AND (Get-Variable  -name "$($TenOrg)Meta").value.Ex10Server){
            $UseOP = $UseExOP = $true ;
            $smsg = "$($TenOrg):Meta.ExOPAccessFromToro($((Get-Variable  -name "$($TenOrg)Meta").value.ExOPAccessFromToro)) -AND/OR Meta.Ex10Server($((Get-Variable  -name "$($TenOrg)Meta").value.Ex10Server)),`ENABLING use of OnPrem Ex system this pass." ;
            if($verbose){ if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
        } else {
            $UseOP = $UseExOP = $false ;
            $smsg = "$($TenOrg):Meta.ExOPAccessFromToro($((Get-Variable  -name "$($TenOrg)Meta").value.ExOPAccessFromToro)) -AND/OR Meta.Ex10Server($((Get-Variable  -name "$($TenOrg)Meta").value.Ex10Server)),`nDISABLING use of OnPrem Ex system this pass." ;
            if($verbose){ if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
        } ;
        #>
        if($UseOP){
            if($useExopNoDep){
                # Connect-ExchangeServerTDO use: creds are implied from the PSSession creds; assumed to have EXOP perms
            } else {
                #*------v GENERIC EXOP CREDS & SRVR CONN BP v------
                # do the OP creds too
                $OPCred=$null ;
                # default to the onprem svc acct
                # userrole='ESVC','SID'
                #$pltGHOpCred=@{TenOrg=$TenOrg ;userrole='ESVC','SID'; verbose=$($verbose)} ;
                # userrole='SID','ESVC'
                $pltGHOpCred=@{TenOrg=$TenOrg ;userrole='SID','ESVC'; verbose=$($verbose)} ;
                $smsg = "get-HybridOPCredentials w`n$(($pltGHOpCred|out-string).trim())" ; 
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level verbose } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                if($OPCred=(get-HybridOPCredentials @pltGHOpCred).cred){
                    # make it script scope, so we don't have to predetect & purge before using new-variable
                    if(get-Variable -Name "cred$($tenorg)OP" -scope Script -ea 0 ){ remove-Variable -Name "cred$($tenorg)OP" -scope Script } ;
                    New-Variable -Name "cred$($tenorg)OP" -scope Script -Value $OPCred ;
                    $smsg = "Resolved $($Tenorg) `$OPCred:$($OPCred.username) (assigned to `$cred$($tenorg)OP)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                } else {
                    $statusdelta = ";ERROR"; # CHANGE|INCOMPLETE|ERROR|WARN|FAIL ;
                    $script:PassStatus += $statusdelta ;
                    set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta) ;
                    $smsg = "Unable to resolve get-HybridOPCredentials -TenOrg $($TenOrg) -userrole 'ESVC' value!"
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    throw "Unable to resolve $($tenorg) `$OPCred value!`nEXIT!"
                    Break ;
                } ;
                $smsg= "Using OnPrem/EXOP cred:`$cred$($tenorg)OP:$((Get-Variable -name "cred$($tenorg)OP" ).value.username)" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                                                            <### CALLS ARE IN FORM: (cred$($tenorg))
                $pltRX10 = @{
                    Credential = (Get-Variable -name "cred$($tenorg)OP" ).value ;
                    #verbose = $($verbose) ;
                    Verbose = $FALSE ; 
                } ;
                $1stConn = $false ; # below uses silent suppr for both x10 & xo!
                if($1stConn){
                    $pltRX10.silent = $pltRXO.silent = $false ;
                } else {
                    $pltRX10.silent = $pltRXO.silent =$true ;
                } ;
                if($pltRX10){ReConnect-Ex2010 @pltRX10 }
                else {ReConnect-Ex2010 }
                #$pltRx10 creds & .username can also be used for local ADMS connections
                ###>
                $pltRX10 = @{
                    Credential = (Get-Variable -name "cred$($tenorg)OP" ).value ;
                    #verbose = $($verbose) ;
                    Verbose = $FALSE ; 
                } ;
                if((get-command Reconnect-Ex2010).Parameters.keys -contains 'silent'){
                    $pltRX10.add('Silent',$false) ;
                } ;
            } ; 
            # defer cx10/rx10, until just before get-recipients qry
            # connect to ExOP X10
            if($useEXOP){
                if($useExopNoDep){ 
                    $smsg = "(Using ExOP:Connect-ExchangeServerTDO(), connect to local ComputerSite)" ; 
                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;           
                    TRY{
                        $Site=[System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::GetComputerSite().Name 
                    }CATCH{$Site=$env:COMPUTERNAME} ;
                    $PSSession = Connect-ExchangeServerTDO -siteName $Site -RoleNames @('HUB','CAS') -verbose ; 
                } else {
                    if($pltRX10){
                        #ReConnect-Ex2010XO @pltRX10 ;
                        ReConnect-Ex2010 @pltRX10 ;
                    } else { Reconnect-Ex2010 ; } ;
                    #Add-PSSnapin -Name 'Microsoft.Exchange.Management.PowerShell.SnapIn'
                    #TK: add: test Exch & AD functional connections
                    TRY{
                        if(get-command -module (get-module |?{$_.name -like 'tmp_*'}).name -name 'get-OrganizationConfig'){} else {
                            $smsg = "(mangled Ex10 conn: dx10,rx10...)" ; 
                            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                            disconnect-ex2010 ; reconnect-ex2010 ; 
                        } ; 
                        if(-not ($OrgName = ((get-OrganizationConfig).DistinguishedName.split(',') |?{$_ -like 'DC=*'}) -join '.' -replace 'DC=','')){
                            $smsg = "Missing Exchange Connection! (no (Get-OrganizationConfig).name returned)" ; 
                            throw $smsg ; 
                            $smsg | write-warning  ; 
                        } ; 
                    } CATCH {
                        $ErrTrapd=$Error[0] ;
                        $smsg = $ErrTrapd ;
                        $smsg += "`n";
                        $smsg += $ErrTrapd.Exception.Message ;
                        if ($logging) { _write-log -LogContent $smsg -Path $logfile -useHost -Level WARN } 
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                        CONTINUE ;
                    } ;
                }
            } ; 
            if($useForestWide){
                #region  ; #*------v OPTIONAL CODE TO ENABLE FOREST-WIDE NATIVE EXCHANGE SUPPORT v------
                $smsg = "(`$useForestWide:$($useForestWide)):Enabling EXoP Forestwide)" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                $smsg = 'Set-AdServerSettings -ViewEntireForest `$True' ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                Set-AdServerSettings -ViewEntireForest $True ;
                #endregion  ; #*------^ END OPTIONAL CODE TO ENABLE FOREST-WIDE NATIVE EXCHANGE SUPPORT ^------
            } ;
        } else {
            $smsg = "(`$useOP:$($UseOP))" ; 
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        }  ;  # if-E $UseOP
        #endregion GENERIC_EXOP_CREDS_&_SRVR_CONN #*------^ END GENERIC EXOP CREDS & SRVR CONN BP ^------
        
        #region UseOPAD #*------v UseOPAD v------
        if($UseOP -OR $UseOPAD){
            #region GENERIC_ADMS_CONN_&_XO #*------v GENERIC ADMS CONN & XO  v------
            $smsg = "(loading ADMS...)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            # always capture load-adms return, it outputs a $true to pipeline on success
            $ADMTLoaded = load-ADMS -Verbose:$FALSE ;
            # 9:32 AM 4/20/2023 trimmed disabled/fw-borked cross-org code
            TRY {
                if(-not(Get-ADDomain  -ea STOP).DNSRoot){
                    $smsg = "Missing AD Connection! (no (Get-ADDomain).DNSRoot returned)" ; 
                    throw $smsg ; 
                    $smsg | write-warning  ; 
                } ; 
                $objforest = get-adforest -ea STOP ; 
                # Default new UPNSuffix to the UPNSuffix that matches last 2 elements of the forestname.
                $forestdom = $UPNSuffixDefault = $objforest.UPNSuffixes | ?{$_ -eq (($objforest.name.split('.'))[-2..-1] -join '.')} ; 
                if($useForestWide){
                    #region  ; #*------v OPTIONAL CODE TO ENABLE FOREST-WIDE AD GC QRY SUPPORT v------
                    $smsg = "(`$useForestWide:$($useForestWide)):Enabling AD Forestwide)" ; 
                    if($silent){}elseif ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    $smsg = 'Set-AdServerSettings -ViewEntireForest `$True' ;
                    if($silent){}elseif ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    #TK 9:44 AM 10/6/2022 need org wide for rolegrps in parent dom (only for onprem RBAC, not EXO)
                    $GcFwide = "$((Get-ADDomainController -Discover -Service GlobalCatalog).hostname):3268" ;        
                    #endregion  ; #*------^ END  OPTIONAL CODE TO ENABLE FOREST-WIDE AD GC QRY SUPPORT  ^------
                } ;    
            } CATCH {
                $ErrTrapd=$Error[0] ;
                $smsg = $ErrTrapd ;
                $smsg += "`n";
                $smsg += $ErrTrapd.Exception.Message ;
                if ($logging) { _write-log -LogContent $smsg -Path $logfile -useHost -Level WARN } 
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                CONTINUE ;
            } ;        
            #endregion GENERIC_ADMS_CONN_&_XO #*------^ END GENERIC ADMS CONN & XO ^------
        } else {
            $smsg = "(`$UseOP:$($UseOP))" ; 
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        }  ;
        #if (!$domaincontroller) { $domaincontroller = get-gcfast } ;
        #if(!$domaincontroller){ if(test-path function:get-gcfast){$domaincontroller = get-gcfast} else { throw "no get-gcfast()!" } ;} else {"(existing `$domaincontroller:$($domaincontroller))"} ;
        # use new get-GCFastXO cross-org dc finde
        # default to Op_ExADRoot forest from $TenOrg Meta
        #if($UseOP -AND -not $domaincontroller){
        if($UseOP -AND -not (get-variable domaincontroller -ea 0)){
            #$domaincontroller = get-GCFastXO -TenOrg $TenOrg -subdomain ((get-variable -name "$($TenOrg)Meta").value['OP_ExADRoot']) -verbose:$($verbose) |?{$_.length};
            # need to debug the above, credential issue?
            # just get it done
            $domaincontroller = get-GCFast
        }  else { 
            # have to defer to get-azuread, or use EXO's native cmds to poll grp members
            # TODO 1/15/2021
            $useEXOforGroups = $true ; 
            $smsg = "$($TenOrg):HAS NO ON-PREM ACTIVEDIRECTORY, DEFERRING ALL GROUP ACCESS & MGMT TO NATIVE EXO CMDS!" ; 
            if($silent){}elseif ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        } ; 
        if($useForestWide -AND -not $GcFwide){
            #region  ; #*------v OPTIONAL CODE TO ENABLE FOREST-WIDE ACTIVEDIRECTORY SUPPORT: v------
            $smsg = "`$GcFwide = Get-ADDomainController -Discover -Service GlobalCatalog" ;
            if($silent){}elseif ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            $GcFwide = "$((Get-ADDomainController -Discover -Service GlobalCatalog).hostname):3268" ;
            $smsg = "Discovered `$GcFwide:$($GcFwide)" ; 
            if($silent){}elseif ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            #endregion  ; #*------^ END OPTIONAL CODE TO ENABLE FOREST-WIDE ACTIVEDIRECTORY SUPPORT ^------
        } ;
        #endregion UseOPAD #*------^ END UseOPAD ^------

        #region MSOL_CONNECTION ; #*------v  MSOL CONNECTION v------
        #$UseMSOL = $false 
        if($UseMSOL){
            #$reqMods += "connect-msol".split(";") ;
            #if ( !(check-ReqMods $reqMods) ) { write-error "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Missing function. EXITING." ; Break ; }  ;
            $smsg = "(loading MSOL...)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } Error|Warn|Debug
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            #connect-msol ;
            connect-msol @pltRXOC ;
        } else {
            $smsg = "(`$UseMSOL:$($UseMSOL))" ; 
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        } ;
        #endregion MSOL_CONNECTION ; #*------^  MSOL CONNECTION ^------

        #region AZUREAD_CONNECTION ; #*------v AZUREAD CONNECTION v------
        #$UseAAD = $false 
        if($UseAAD){
            #$reqMods += "Connect-AAD".split(";") ;
            #if ( !(check-ReqMods $reqMods) ) { write-error "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Missing function. EXITING." ; Break ; }  ;
            $smsg = "(loading AAD...)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            Connect-AAD @pltRXOC ;
        } else {
            $smsg = "(`$UseAAD:$($UseAAD))" ; 
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        } ;
        #endregion AZUREAD_CONNECTION ; #*------^ AZUREAD CONNECTION ^------
  
        <# defined above
        # EXO connection
        $pltRXO = @{
            Credential = (Get-Variable -name cred$($tenorg) ).value ;
            verbose = $($verbose) ; } ;
        #>
        <#
        if($VerbosePreference = "Continue"){
            $VerbosePrefPrior = $VerbosePreference ;
            $VerbosePreference = "SilentlyContinue" ;
            $verbose = ($VerbosePreference -eq "Continue") ;
        } ;
        disconnect-exo ;
        if ($script:useEXOv2) { reconnect-eXO2 @pltRXOC }
        else { reconnect-EXO @pltRXOC } ;
        # reenable VerbosePreference:Continue, if set, during mod loads
        if($VerbosePrefPrior -eq "Continue"){
            $VerbosePreference = $VerbosePrefPrior ;
            $verbose = ($VerbosePreference -eq "Continue") ;
        } ;
        #>
        #-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
        #endregion SERVICE_CONNECTIONS #*======^ END SERVICE_CONNECTIONS ^======
    

        # finally if we're using pipeline, and aggregating, we need to aggreg outside of the process{} block
        if($PSCmdlet.MyInvocation.ExpectingInput){
            # pipeline instantiate an aggregator here
        } ;

        # check if using Pipeline input or explicit params:
        if ($PSCmdlet.MyInvocation.ExpectingInput) {
            $smsg = "Data received from pipeline input: '$($InputObject)'" ;
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        } else {
            # doesn't actually return an obj in the echo
            #$smsg = "Data received from parameter input: '$($InputObject)'" ;
            #if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
            #else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        } ;

    }
    PROCESS{
        $Error.Clear() ; 
       
        if(-not $users){
            $users= (get-clipboard).trim().replace("'",'').replace('"','') ;
            if($users){
                write-verbose "No -users specified, detected value on clipboard:`n$($users)" ;
            } else {
                write-warning "No -users specified, nothing matching dname, emailaddress or samaccountname, found on clipboard. EXITING!" ;
                Break ;
            } ;
        } else {
            write-verbose "($(($users|measure).count)) user(s) specified:`n'$($users -join "','")'" ;
        } ;

        $ttl = ($users|measure).count ; $Procd=0 ;
        [array]$Rpt =@() ;
        # with pipeline input, the pipeline evals as either $_ (if unmapped to a param in binding), or iterating on the mapped value.
        #     the foreach loop below doesn't actually loop. Process{} is the loop with a pipeline-fed param, and the bound - $users - variable once per pipeline bound element - per array item on an array -
        #     is run with the $users value populated with each element in turn. IOW, the foreach is a single-run pass, and the Process{} block is the loop.
        # you need both a bound $users at the top - to handle explicit assigns resolve-user -users $variable.
        # with a process {} block to handle any pipeline passed input. The pipeline still maps to the bound param: $users, but the entire process{} is run per element, rather than iteratign the internal $users foreach.
        #region PIPELINE_PROCESSINGLOOP ; #*------v PIPELINE_PROCESSINGLOOP v------
        foreach ($usr in $users){
            
            #region START-LOG #*======v START-LOG OPTIONS v======
            $useSLogHOl = $true ; # one or 
            $useSLogSimple = $false ; # ... the other
            $useTransName = $false ; # TRANSCRIPTNAME
            $useTransPath = $false ; # TRANSCRIPTPATH
            $useTransRotate = $false ; # TRANSCRIPTPATHROTATE
            $useStartTrans = $false ; # STARTTRANS
            #region START-LOG-HOLISTIC #*------v START-LOG-HOLISTIC v------
            if($useSLogHOl){
                # Single log for script/function example that accomodates detect/redirect from AllUsers scope'd installed code, and hunts a series of drive letters to find an alternate logging dir (defers to profile variables)
                #${CmdletName} = $rPSCmdlet.MyInvocation.MyCommand.Name ;
                if(-not (get-variable LogPathDrives -ea 0)){$LogPathDrives = 'd','c' };
                foreach($budrv in $LogPathDrives){if(test-path -path "$($budrv):\scripts" -ea 0 ){break} } ;
                if(-not (get-variable rgxPSAllUsersScope -ea 0)){
                    $rgxPSAllUsersScope="^$([regex]::escape([environment]::getfolderpath('ProgramFiles')))\\((Windows)*)PowerShell\\(Scripts|Modules)\\.*\.(ps(((d|m))*)1|dll)$" ;
                } ;
                if(-not (get-variable rgxPSCurrUserScope -ea 0)){
                    $rgxPSCurrUserScope="^$([regex]::escape([Environment]::GetFolderPath('MyDocuments')))\\((Windows)*)PowerShell\\(Scripts|Modules)\\.*\.(ps((d|m)*)1|dll)$" ;
                } ;
                $pltSL=[ordered]@{Path=$null ;NoTimeStamp=$false ;Tag=$null ;showdebug=$($showdebug) ; Verbose=$($VerbosePreference -eq 'Continue') ; whatif=$($whatif) ;} ;
                # if using [CmdletBinding(SupportsShouldProcess)] + -WhatIf:$($WhatIfPreference):
                #$pltSL=[ordered]@{Path=$null ;NoTimeStamp=$false ;Tag=$null ;showdebug=$($showdebug) ; Verbose=$($VerbosePreference -eq 'Continue') ; whatif=$($WhatIfPreference) ;} ;
                #$pltSL=[ordered]@{Path=$null ;NoTimeStamp=$false ;Tag="$($ticket)-$($TenOrg)-LASTPASS-" ;showdebug=$($showdebug) ; Verbose=$($VerbosePreference -eq 'Continue') ; whatif=$($WhatIfPreference) ;} ;
                #$pltSL.Tag = $ModuleName ; 
                #$pltSL.Tag = "$($ticket)-$($usr)" ; 
                $pltSL.Tag = $((@($ticket,$usr) |?{$_}) -join '-')
                if($script:rPSCommandPath){ $prxPath = $script:rPSCommandPath }
                elseif($script:PSCommandPath){$prxPath = $script:PSCommandPath}
                if($rMyInvocation.MyCommand.Definition){$prxPath2 = $rMyInvocation.MyCommand.Definition }
                elseif($MyInvocation.MyCommand.Definition){$prxPath2 = $MyInvocation.MyCommand.Definition } ; 
                if($prxPath){
                    if(($prxPath -match $rgxPSAllUsersScope) -OR ($prxPath -match $rgxPSCurrUserScope)){
                        $bDivertLog = $true ; 
                        switch -regex ($prxPath){
                            $rgxPSAllUsersScope{$smsg = "AllUsers"} 
                            $rgxPSCurrUserScope{$smsg = "CurrentUser"}
                        } ;
                        $smsg += " context script/module, divert logging into [$budrv]:\scripts" 
                        write-verbose $smsg  ;
                        if($bDivertLog){
                            if((split-path $prxPath -leaf) -ne $cmdletname){
                                # function in a module/script installed to allusers|cu - defer name to Cmdlet/Function name
                                $pltSL.Path = (join-path -Path "$($budrv):\scripts" -ChildPath "$($cmdletname).ps1") ;
                            } else {
                                # installed allusers|CU script, use the hosting script name
                                $pltSL.Path = (join-path -Path "$($budrv):\scripts" -ChildPath (split-path $prxPath -leaf)) ;
                            }
                        } ;
                    } else {
                        $pltSL.Path = $prxPath ;
                    } ;
               }elseif($prxPath2){
                    if(($prxPath2 -match $rgxPSAllUsersScope) -OR ($prxPath2 -match $rgxPSCurrUserScope) ){
                         $pltSL.Path = (join-path -Path "$($budrv):\scripts" -ChildPath (split-path $prxPath2 -leaf)) ;
                    } elseif(test-path $prxPath2) {
                        $pltSL.Path = $prxPath2 ;
                    } elseif($cmdletname){
                        $pltSL.Path = (join-path -Path "$($budrv):\scripts" -ChildPath "$($cmdletname).ps1") ;
                    } else {
                        $smsg = "UNABLE TO RESOLVE A FUNCTIONAL `$CMDLETNAME, FROM WHICH TO BUILD A START-LOG.PATH!" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Warn } #Error|Warn|Debug 
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        BREAK ;
                    } ; 
                } else{
                    $smsg = "UNABLE TO RESOLVE A FUNCTIONAL `$CMDLETNAME, FROM WHICH TO BUILD A START-LOG.PATH!" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Warn } #Error|Warn|Debug 
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    BREAK ;
                }  ;
                write-verbose "start-Log w`n$(($pltSL|out-string).trim())" ; 
                $logspec = start-Log @pltSL ;
                $error.clear() ;
                TRY {
                    if($logspec){
                        $logging=$logspec.logging ;
                        $logfile=$logspec.logfile ;
                        $transcript=$logspec.transcript ;
                        $stopResults = try {Stop-transcript -ErrorAction stop} catch {} ;
                        if($stopResults){
                            $smsg = "Stop-transcript:$($stopResults)" ; 
                            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                        } ; 
                        $startResults = start-Transcript -path $transcript ;
                        if($startResults){
                            $smsg = "start-transcript:$($startResults)" ; 
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        } ; 
                    } else {throw "Unable to configure logging!" } ;
                } CATCH [System.Management.Automation.PSNotSupportedException]{
                    if($host.name -eq 'Windows PowerShell ISE Host'){
                        $smsg = "This version of $($host.name):$($host.version) does *not* support native (start-)transcription" ; 
                    } else { 
                        $smsg = "This host does *not* support native (start-)transcription" ; 
                    } ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                } CATCH {
                    $ErrTrapd=$Error[0] ;
                    $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                } ;
            } ; 
            #endregion START-LOG-HOLISTIC #*------^ END START-LOG-HOLISTIC ^------
            #region STARTTRANS ; #*------v STARTTRANSCRIPT v------
            if($useStartTrans){
                TRY {
                    if($transcript){
                        $stopResults = try {Stop-transcript -ErrorAction stop} catch {} ;
                        if($stopResults){
                            $smsg = "Stop-transcript:$($stopResults)" ; 
                            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                        } ; 
                        $startResults = start-Transcript -path $transcript ;
                        if($startResults){
                            $smsg = "start-transcript:$($startResults)" ; 
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        } ; 
                    } else {
                        $smsg = "UNPOPULATED `$transcript! - ABORTING!" ; 
                        write-warning $smsg ; 
                        throw $smsg ; 
                        break ; 
                    } ;  
                } CATCH [System.Management.Automation.PSNotSupportedException]{
                    if($host.name -eq 'Windows PowerShell ISE Host'){
                        $smsg = "This version of $($host.name):$($host.version) does *not* support native (start-)transcription" ; 
                    } else { 
                        $smsg = "This host does *not* support native (start-)transcription" ; 
                    } ; 
                    write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" ;
                } CATCH {
                    $ErrTrapd=$Error[0] ;
                    $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                    write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" ;
                } ;
            } ; 
            #endregion STARTTRANS ; #*------^ END STARTTRANSCRIPT ^------
            #endregion START-LOG #*======^ START-LOG OPTIONS ^======

            $useLogBuild = $true ;     
            #region LOGBUILD ; #*------v LOGBUILD v------
            if($useLogBuild){
                # building an outputfile name dynamically using paremeters
                #-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
                # start by rebuilding from base of start-log(): $logfile: 'D:\scripts\logs\get-ExOPSmtpReceiveTLSReport-SERVER50'
                # *first* reset $ofile; it's picking the filename up from the OS
                #$ofile = $null;
                #$ofile += ($logfile -split '-LOG-BATCH-')[0] ; # split existing logfile path out.
                # shift to bottom, build as an array
                [array]$ofile = @() ;
                #$ofile += split-path -leaf ($logfile -split '-LOG-BATCH-')[0] ; 
                if($CmdletName){$ofile += $CmdletName} ; 
                $ofile += $((@($ticket,$usr) |?{$_}) -join '-')
                <# some explicit param adds
                if($Days){
                    $ofile += "$($Days)d" ;
                } ;
                if($TargetDate){
                    $ofile += "T$((get-date $TargetDate -format 'yyyyMMdd-HHmmtt'))"
                } ;
                #>
                <#if($TenantName){
                    $ofile += "To$($TenantName)" ;
                } ;
                #>
                <#
                $smsg = "dyn qry all 'TLS'-named boolean params and append them" ;
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;   
                $ParamsNonDefault | Where-Object{$_ -match 'TLS'} | ForEach-Object{get-variable $_} |Where-Object{$_.value -eq $true} | Select-Object -expand Name |ForEach-Object{
                    $ofile += "$($_)" ;
                } ;
                write-verbose 'append these all by name if $true (e.g. booleans)' ;
                $ParamsAppendTrue = @('Month''NoDetail') ;
                $ParamsNonDefault | Where-Object{$_ -match [regex]"($($ParamsAppendTrue -join '|'))"} | ForEach-Object{get-variable $_} |Where-Object{$_.value -eq $true} | Select-Object -expand Name  |ForEach-Object{
                    $ofile += "$($_)" ;
                } ;
                #>
                <#
                write-verbose 'append these with name & value in a single string (comma prefix the combo)' ;
                $ParamsAppendNameValue = @('Session') ;
                $ParamsNonDefault | Where-Object{$_ -match [regex]"($($ParamsAppendNameValue -join '|'))"} | ForEach-Object{get-variable $_} |Where-Object{$_.value -ne $null} |ForEach-Object{
                    $ofile += ("$($_.Name)-$($_.value)" -join ',' );
                } ;
                #>
                <#
                write-verbose 'append these with name & value where value may be an array of strings (comma prefix the combo)' ;
                $ParamsAppendNameValueArray = @('Session') ;
                $dataColMax = 40 
                $ParamsNonDefault | Where-Object{$_ -match [regex]"($($ParamsAppendNameValue -join '|'))"} | ForEach-Object{get-variable $_} |Where-Object{$_.value -ne $null} |ForEach-Object{
                    $tagitem = ("$($_.Name)-$($_.value -join '~')" -join ',' ) ; 
                    if($tagitem.length -gt $dataColMax){
                        $tagitem = (($tagitem.ToString())).substring(0,[System.Math]::Min($dataColMax, $tagitem.Length)) ; 
                    } else { 
                        $ofile += $tagitem;
                    } ; 
                } ;
                #>
                $ofile += "REPORT" ; 
                $ofile+= "runon$((get-date -format 'yyyyMMdd-HHmmtt'))" ;
                [string]$ofile = $ofile  -join '-' ;
                #[string]$ofile = $ofile  -join ',' ;
                #[string]$ofile += "-log.txt" ;
                [string]$ofile += ".xml" ;
                #[string]$ofile += ".csv" ;
                #[string]$ofile += ".json" ;
                #$ofile = join-path -path (split-path $logfile) -ChildPath $ofile -ea STOP ;
                # clear any os illegal fso chars: chk just the filename
                $pltJP = @{
                    path= (split-path $logfile) ;
                    #childpath = [RegEx]::Replace((split-path $ofile -leaf), "[{0}]" -f ([RegEx]::Escape(-join [System.IO.Path]::GetInvalidFileNameChars())), '')
                    childpath = [RegEx]::Replace(($ofile), "[{0}]" -f ([RegEx]::Escape(-join [System.IO.Path]::GetInvalidFileNameChars())), '') ; 
                    erroraction = 'STOP' ; 
                } ;
                [string]$ofile = join-path @pltJP ;
                $smsg = "`$ofile:$($ofile)" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level PROMPT } 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                #Levels:Error|Warn|Info|H1|H2|H3|Debug|Verbose|Prompt
            } ; 
            #endregion LOGBUILD ; #*------^ END LOGBUILD ^------

            #$fname = $lname = $dname = $OPRcp = $OPMailbox = $OPRemoteMailbox = $ADUser = $xoRcp = $xoMailbox = $xoUser = $xoMemberOf = $MsolUser = $LicenseGroup = $null ;
            $isEml=$isDname=$isSamAcct=$isXORcpMulti  = $false ;

            $hSum = [ordered]@{
                dname = $null ;
                fname = $null ;
                lname = $null ;
                OPRcp = $null ;
                xoRcp = $null ;
                OPMailbox = $null ;
                OPRemoteMailbox = $null ;
                ADUser = $null ;
                Federator = $null ;
                xoMailbox = $null ;
                xoMUser = $null ;
                xoUser = $null ;
                xoMemberOf = $null ;
                txGuest = $null ;
                OPMapiTest = $null ;
                xoMapiTest = $null ;
                MsolUser = $null ;
                AADUser = $null ; # added for MailUser variant
                AADUserMgr = $null ;
                AADUserLics = $null ; 
                LicenseGroup = $null ;
                isDirSynced = $null 
                isNoBrain = $false ;
                isSplitBrain = $false;
                #isUnlicensed = $false ;
                IsLicensed = $false ; 
                IsDisabledOU = $false ; 
                IsADDisabled = $false ; 
                IsAADDisabled = $false ; 
            } ;
            $procd++ ;
            write-verbose "processing:$($usr)" ;
            if($getMobile){
                $smsg = "(-getMobile:retrieving user xo MobileDevices)" ; 
                if($verbose){
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-verbose $smsg } ; 
                } ; 
                $hsum.add('xoMobileDeviceStats',$null) ; 
            } ; 
            if($getQuotaUsage){
                $smsg = "(-getQuotaUsage:retrieving user xo Mailbox*Statistics & Effective Quotas)" ; 
                if($verbose){
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-verbose $smsg } ; 
                } ; 
                $hsum.add('xoMailboxStats',$null) ; 
                $hsum.add('xoMailboxFolderStats',$null) ; 
                $hsum.add('xoEffectiveQuotas',$null) ; 
                $hsum.add('xoNetOfSendReceiveQuotaMB',$null) ; 
                [string]$ofMbxFolderStats = $ofile.replace('REPORT',"folder-sizes-NONHIDDEN-NONZERO") ; 

            } ; 

            if($usr -match $rgxAccentedNameChars){
                # 9:36 AM 9/23/2024 pre remove all diacritics & latin chars 
                #Remove-StringDiacritic -String 'Helen Bräuchle' |Remove-StringLatinCharacters
                $smsg = "Remove-StringDiacritic -String $($usr) (if needed)" ; 
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                $usr = Remove-StringDiacritic -String $usr ; 
            
                $smsg = "Remove-StringLatinCharacters -String $($usr) (if needed)" ; 
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                $usr = Remove-StringLatinCharacters -String $usr ; 
            } ; 

            switch -regex ($usr){
                $rgxEmailAddr {
                # $rgxEmailAddr = "^([0-9a-zA-Z]+[-._+&'])*[0-9a-zA-Z]+@([-0-9a-zA-Z]+[.])+[a-zA-Z]{2,63}$" ;
                    $hSum.fname,$hSum.lname = $usr.split('@')[0].split('.') ;
                    $hSum.dname = $usr ; # temp set eml as dname, re-resolved to proper further on
                    write-verbose "(detected user ($($usr)) as EmailAddr)" ;
                    $isEml = $true ;
                    Break ;
                }
                $rgxObjNameNewHires{
                # $rgxObjNameNewHires = "^([a-zA-Z]{2,}(\s|\.)[a-zA-Z]{1,}'?-?[a-zA-Z]{2,}\s?([a-zA-Z]{1,})?)_[a-z0-9]{10}"  
                # Name:Fname LName_f4feebafdb (appending uniqueness guid chunk)
                    write-verbose "(detected user ($($usr)) as ObjNameNewHires)" ;
                    $hSum.fname,$hSum.lname = $usr.split('_')[0].split(' ');
                    $hSum.dname = $usr.split('_')[0] ;
                    write-verbose "(detected user ($($usr)) as DisplayName)" ;
                    $isObjName = $true ;
                    Break ;
                }
                $rgxSamAcctNameTOR {
                # $rgxSamAcctNameTOR = "^\w{2,20}$" ; # up to 20c, the limit prior to win2k
                    $hSum.lname = $usr ;
                    write-verbose "(detected user ($($usr)) as SamAccountName)" ;
                    $isSamAcct  = $true ;
                    Break ;
                }
                # move dname below samacct, it's a broader spec
                $rgxDName {
                    # $rgxDName = "^([a-zA-Z]{2,}(\s|\.)[a-zA-Z]{1,}'?-?[a-zA-Z]{2,}\s?([a-zA-Z]{1,})?)" ;
                    #updated: CMW uses : in their room names, so went for broader AD dname support, per AI, and web specs, added 1-256char AD restriction
                    #$rgxDName ="[a-zA-Z0-9\s$([Regex]::Escape('/\[:;|=,+*?<>') + '\]' + '\"')]{1,256}" ; 
                    if($usr.contains('.')){
                        write-verbose "(replacing period in DName)" ;
                        $usr = $usr.replace('.',' ') ;
                    };
                    $hSum.fname,$hSum.lname = $usr.split(' ') ;
                    $hSum.dname = $usr ;
                    write-verbose "(detected user ($($usr)) as DisplayName)" ;
                    $isDname = $true ;
                    Break ;
                }
                default {
                    write-warning "$((get-date).ToString('HH:mm:ss')):No -user specified, nothing matching dname, emailaddress or samaccountname, found on clipboard. EXITING!" ;
                    #Break ;
                } ;
            } ;
            
            $sBnr="===v ($($Procd)/$($ttl)):Input: '$($usr)' | '$($hSum.fname)' | '$($hSum.lname)' v===" ;
            if($isEml){$sBnr+="(EML)"}
            elseif($isDname){$sBnr+="(DNAM)"}
            elseif($isObjName){$sBnr+="(ONAM)"}
            elseif($isSamAcct){$sBnr+="(SAM)"}
            write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($sBnr)" ;

            write-host -foreground yellow "get-Rmbx/xMbx: " -nonewline;


            # $isEml=$isDname=$isSamAcct=$false ;
            $MDtbl=[ordered]@{NoDashRow=$true } ; # out-markdowntable splat
            $pltGMailObj=[ordered]@{
                ResultSize = $MaxRecips ;
            } ;
            if($isEml -OR $isSamAcct){
                write-verbose "processing:'identity':$($usr)" ;
                $pltGMailObj.add('identity',$usr) ;
            } ;
            if($isObjName){
                # filter on Name, (not dname)
                $dname = $hSum.dname
                # 11:00 AM 10/11/2024 if dname contains ', have to variant quotewrap
                if($usr -match "[']"){
                    $fltr = "name -like " + $sQot + $usr + $sQot ;
                }else {
                    $fltr = "name -like '$usr'" ;
                } ; 
                write-verbose "processing:'filter':$($fltr)" ;
                $pltGMailObj.add('filter',$fltr) ;
            } ;
            if($isDname){
                # interestinb bug: switched to $hSum.dname: ISE is fine, but ConsoleHost fails to expand the $fltr properly.
                # standard is: Variables: Enclose variables that need to be expanded in single quotation marks (for example, '$User'). Don't use curly-brackets (impedes expansion)
                # workaround: looks like have to proxy the $hsum.Dname, to provide a single non-dotted variable name
                $dname = $hSum.dname
                # 11:00 AM 10/11/2024 if dname contains ', have to variant quotewrap
                if($dname -match "[']"){
                    $fltr = "displayname -like " + $sQot + $dname + $sQot ; 
                }else {
                    $fltr = "displayname -like '$dname'" ;
                } ; 
                # 8:47 AM 10/9/2024 where suffixed 'fname lname (SIT)', need functional wildcard to even hope to hit it, lets see if follow on fname lname filters gap fill, when dname is suffixed arbitrarily
                write-verbose "processing:'filter':$($fltr)" ;
                $pltGMailObj.add('filter',$fltr) ;
            } ;

            $error.clear() ;

            #write-verbose "get-[exo]Recipient w`n$(($pltGMailObj|out-string).trim())" ;
            #write-verbose "get-recipient w`n$(($pltGMailObj|out-string).trim())" ;
            # exclude contacts, they don't represent real onprem mbx assoc, and we need to refer those to EXO mbx qry anyway.
            write-verbose "get-recipient w`n$(($pltGMailObj|out-string).trim())" ;
            rx10 -Verbose:$false -silent ;

            $smsg = "get-recipient w`n$(($pltGMailObj|out-string).trim())`n...| ?{$_.recipienttypedetails -ne 'MailContact'}" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;            

            if($hSum.OPRcp=get-recipient @pltGMailObj -ea 0 | select -first $MaxRecips | ?{$_.recipienttypedetails -ne 'MailContact'}){
                write-verbose "`$hSum.OPRcp found as $($pltGMailObj.filter)" ;
                
            } elseif($isDname -and $hsum.lname) {
                # put in missing *, hits on mis-spellings 'Spark' matches 'Sparks' w wildcard
                if($hsum.lname -match "[']"){
                    $fltr = "recipienttypedetails -ne " + $sQot + "MailContact" + $sQot ; 
                    $fltr += " -AND displayname -like " + $sQot + $($hsum.lname) + $sQot ;
                }else {
                    $fltr = "recipienttypedetails -ne 'MailContact'" ; 
                    $fltr += " -AND displayname -like '$($hsum.lname)'" ;
                } ; 
                if($hsum.fname){
                    # try first 3 of fname first
                    if($hsum.fname -match "[']"){
                        $fltr += " -AND firstName -like " + $sQot + $($hsum.fname.substring(0,3)) + "*" + $sQot ; 
                    }else {
                        $fltr += " -AND firstName -like '$($hsum.fname.substring(0,3))*'" ; 
                    } ; 
                    
                    if($hSum.OPRcp=get-recipient -filter $fltr -ea 0 | select -first $MaxRecips |?{$_.recipienttypedetails -ne 'MailContact'}){
                        write-verbose "`$hSum.OPRcp found as $($pltGMailObj.filter)" ;
                    }elseif($hsum.fname){
                        # retry first initial
                        if($hsum.lname -match "[']"){
                            $fltr = "recipienttypedetails -ne " + $sQot + "MailContact" + $sQot + " -AND lastName -like " + $sQot + $($hsum.lname) + "*" + $sQot ; 
                        }else {
                            $fltr = "recipienttypedetails -ne 'MailContact' -AND lastName -like '$($hsum.lname)*'" ; 
                        } ; 
                        if($hsum.fname -match "[']"){
                            $fltr += " -AND firstName -like " + $sQot + $($hsum.fname.substring(0,1)) + "*" + $sQot ; 
                        }else {
                            $fltr += " -AND firstName -like '$($hsum.fname.substring(0,1))*'" ; 
                        } ; 
                        
                        if($hSum.OPRcp=get-recipient -filter $fltr -ea 0 | select -first $MaxRecips |?{$_.recipienttypedetails -ne 'MailContact'}){
                            write-verbose "`$hSum.OPRcp found as $($pltGMailObj.filter)" ;
                        }elseif($hsum.lname){
                            # do wildcard lname matches
                            if($hsum.lname -match "[']"){
                                    $fltr = "recipienttypedetails -ne " + $sQot + "MailContact" + $sQot + " -AND lastName -like " + $sQot + $($hsum.lname) + "*" + $sQot ; 
                            }else {
                                $fltr = "recipienttypedetails -ne 'MailContact' -AND lastName -like '$($hsum.lname)*'" ; 
                            } ; 
                            
                            if($hSum.OPRcp=get-recipient -filter $fltr -ea 0 | select -first $MaxRecips |?{$_.recipienttypedetails -ne 'MailContact'}){
                                write-verbose "`$hSum.OPRcp found as $($pltGMailObj.filter)" ;
                            }else{
                                $smsg = "(Failed to OP:get-recipient on:$($usr))"
                                if($isDname){$smsg += " or *$($hsum.lname )*"}
                                write-host $smsg ;                            
                            } ;
                        }
                    } ; 
   
                } ; 
            } ; 

            if(-not $hsum.OpRcp){
                $smsg = "(Failed to OP:get-recipient on:$($usr))"
                if($isDname){$smsg += " or *$($hsum.lname )*"}
                write-host $smsg ;
            } else {
                # 8:55 AM 10/9/2024 arrays come through (esp for suffixed duped names): put in an explicit loop
                #$prpFTARcp = 'Name','RecipientTypeDetails','RecipientType','PrimarySmtpAddress','alias'
                $smsg = "`$hSum.OPRcp:`n$(($hSum.OPRcp | ft -a $prpFTARcp |out-string).trim())" ;
                if($hSum.OPRcp -is [array]){
                    $smsg += "`n==> MULTIPLE RECIPIENTS MATCHED!" ; 
                    write-warning $smsg ; 
                } else { 
                    write-verbose $smsg ; 
                } ; 
                $hSum.OPRcp | ForEach-Object{
                    $tmpRcp = $_ ; 
                    #switch ($hSum.OPRcp.recipienttypedetails){
                    switch ($tmpRcp.recipienttypedetails){
                        'RemoteUserMailbox' {write-host "(Rmbx)" -nonewline}
                        # 8:53 AM 10/9/2024 add to cover mbx2shared conversion results
                        'RemoteSharedMailbox' {write-host "(Rmbx *SHARED*)" -nonewline} 
                        # 12:23 PM 12/26/2024 add resource & remote res's
                        'RemoteRoomMailbox' {write-host "(Rmbx *ROOM*)" -nonewline} 
                        'RemoteEquipmentMailbox' {write-host "(Rmbx *EQUIP*)" -nonewline} 
                        'UserMailbox' {write-host "(Mbx)" -nonewline}
                        'SharedMailbox' {write-host "(SMbx)" -nonewline}
                        'RoomMailbox' {write-host "(RoomMbx)" -nonewline}
                        'EquipmentMailbox' {write-host "(EquipMbx)" -nonewline}
                        'MailUser' {
                            $smsg = "MAILUSER WO RMBX DETECTED! - POSSIBLE NOBRAIN?"
                            write-warning $smsg
                            #$hsum.isNoBrain = $true ;    
                        }
                        'MailUniversalDistributionGroup' {write-host "(DG)" -nonewline}
                        'DynamicDistributionGroup'  {write-host "(DDG)" -nonewline}
                        'MailContact' {write-host "(MC)" -nonewline]}
                        default{
                            #$smsg = "Unable to resolve `$hSum.OPRcp.recipienttypedetails:$($hSum.OPRcp.recipienttypedetails)" ; 
                            $smsg = "Unable to resolve `$hSum.OPRcp.recipienttypedetails:$($tmpRcp.OPRcp.recipienttypedetails)" ; 
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                            throw $smsg ; 
                            break ; 
                        }
                    }
                }  # loop-E 
            } ; # if-E

            <# 2:57 PM 10/8/2024 splice over from grcp code above
            #write-verbose "$((get-alias ps1GetxRcp).definition) w`n$(($pltGMailObj|out-string).trim())" ;
            write-verbose "get-xorecipient  w`n$(($pltGMailObj|out-string).trim())" ;
            #rxo  -Verbose:$false -silent ;
            if($hSum.xoRcp=get-xorecipient @pltGMailObj -ea 0 | select -first $MaxRecips ){
                write-verbose "`$hSum.xoRcp found" ;
            } elseif($isDname -and $hsum.lname) {
                $smsg = "Failed:RETRY: detected 'LName':$($hsum.lname) for near matches..." ;
                write-host $smsg ;
                $lname = $hsum.lname ;
                $fltrB = "displayname -like '*$lname*'" ;
                write-verbose "RETRY:get-recipient -filter {$($fltr)}" ;
                if($hSum.xoRcp=get-xorecipient -filter $fltr -ea 0 -ResultSize $MaxRecips |?{$_.recipienttypedetails -ne 'MailContact'}){
                    write-verbose "`$hSum.xoRcp found" ;
                } ;
            }
            if(-not $hsum.xoRcp){
                #$smsg = "Failed to $((get-alias ps1GetxRcp).definition) on:$($usr)"
                $smsg = "get-xorecipient on:$($usr)"
                if($isDname){$smsg += " or *$($hsum.lname )*"} ;
                write-host $smsg ;
            } else {
                $smsg =  "`$hSum.xoRcp:`n$(($hSum.xoRcp|out-string).trim())" ;
                write-verbose $smsg ;
                if($hSum.xoRcp -is [system.array]){
                    write-warning "Multiple matching xoRcps!:$($smsg)`nTHIS WILL NOT RETURN FULL AADUSER ETC FOR BOTH OBJECTS!`nUSE TARGETED UPN ETC TO DUMP VARIANT OBJECTS!" ;
                    $isXORcpMulti = $true ;
                } ;
            } ;
            #>
            #-=-=-=-=-=-=-=-=
            if ($script:useEXOv2) { reconnect-eXO2 @pltRXOC }
            else { reconnect-EXO @pltRXOC } ;
            $smsg = "get-xorecipient w`n$(($pltGMailObj|out-string).trim())`n...| ?{$_.recipienttypedetails -ne 'MailContact'}" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            if($hSum.xoRcp=get-xorecipient @pltGMailObj -ea 0 | select -first $MaxRecips | ?{$_.recipienttypedetails -ne 'MailContact'}){
                write-verbose "`$hSum.xoRcp found as $($pltGMailObj.filter)" ;
            } elseif($isDname -and $hsum.lname) {
                
                # put in missing *, hits on mis-spellings 'Spark' matches 'Sparks' w wildcard
                if($hsum.lname -match "[']"){
                    $fltr = "recipienttypedetails -ne " + $sQot + "MailContact" + $sQot ;
                    $fltr += " -AND lastName -like " + $sQot + $($hsum.lname) + "*" + $sQot ;
                }else{
                    $fltr = "recipienttypedetails -ne 'MailContact'" ;
                    $fltr += " -AND lastName -like '$($hsum.lname)*'" ;
                } ; 
                if($hsum.fname){
                    # try first 3 of fname first
                    if($hsum.fname -match "[']"){
                        $fltr += " -AND firstName -like " + $sQot + $($hsum.fname.substring(0,3)) + "*" + $sQot ;
                    }else{
                        $fltr += " -AND firstName -like '$($hsum.fname.substring(0,3))*'" ;
                    } ; 
                    if($hSum.xoRcp=get-xorecipient -filter $fltr -ea 0 | select -first $MaxRecips |?{$_.recipienttypedetails -ne 'MailContact'}){
                        write-verbose "`$hSum.xoRcp found as $($pltGMailObj.filter)" ;
                    }elseif($hsum.fname){
                        # retry first initial
                        if($hsum.lname -match "[']"){
                             $fltr = "recipienttypedetails -ne " + $sQot + "MailContact" + $sQot + " -AND lastName -like " + $sQot + $($hsum.lname) + "*" + $sQot ;
                        } else { 
                            $fltr = "recipienttypedetails -ne 'MailContact' -AND lastName -like '$($hsum.lname)*'" ;
                        }
                        if($hsum.fname -match "[']"){
                            $fltr += " -AND firstName -like " + $sQot + $($hsum.fname.substring(0,1)) + "*" + $sQot ;
                        } else { 
                            $fltr += " -AND firstName -like '$($hsum.fname.substring(0,1))*'" ;
                        } ; 

                        if($hSum.xoRcp=get-xorecipient -filter $fltr -ea 0 | select -first $MaxRecips |?{$_.recipienttypedetails -ne 'MailContact'}){
                            write-verbose "`$hSum.xoRcp found as $($pltGMailObj.filter)" ;
                        }elseif($hsum.lname){
                            # do wildcard lname matches
                            if($hsum.fname -match "[']"){
                                $fltr = "recipienttypedetails -ne " + $sQot + "MailContact" + $sQot + " -AND lastName -like " + $sQot + $($hsum.lname) + "*" + $sQot ;
                            } else { 
                                $fltr = "recipienttypedetails -ne 'MailContact' -AND lastName -like '$($hsum.lname)*'" ;
                            } ; 
                            if($hSum.xoRcp=get-xorecipient -filter $fltr -ea 0 | select -first $MaxRecips |?{$_.recipienttypedetails -ne 'MailContact'}){
                                write-verbose "`$hSum.xoRcp found as $($pltGMailObj.filter)" ;
                            }else{
                                $smsg = "(Failed to OP:get-xorecipient on:$($usr))"
                                if($isDname){$smsg += " or *$($hsum.lname )*"}
                                write-host $smsg ;
                            } ;
                        }
                    } ;
                } ;
            } ; 
            #-=-=-=-=-=-=-=-=

            # 9:14 AM 10/9/2024 put in the same type of user type reporting as for oprcp
            #-=-=-=-=-=-=-=-=
            if(-not $hSum.xoRcp){
                $smsg = "(Failed to OP:get-recipient on:$($usr))"
                if($isDname){$smsg += " or *$($hsum.lname )*"}
                write-host $smsg ;
            } else {
                # 8:55 AM 10/9/2024 arrays come through (esp for suffixed duped names): put in an explicit loop
                #$prpFTARcp = 'Name','RecipientTypeDetails','RecipientType','PrimarySmtpAddress','alias' ; 
                $smsg = "`$hSum.xoRcp:`n$(($hSum.xoRcp | ft -a $prpFTARcp |out-string).trim())" ;
                if($hSum.xoRcp -is [array]){
                    $smsg += "`n==> MULTIPLE RECIPIENTS MATCHED!" ; 
                    write-warning $smsg ; 
                } else {
                    write-verbose $smsg ;
                } ;
                $hSum.xoRcp | ForEach-Object{
                    $tmpxRcp = $_ ;
                    #switch ($hSum.xoRcp.recipienttypedetails){
                    # patched in xo equiv variants, added SharedMailbox too
                    switch ($tmpxRcp.recipienttypedetails){
                        'RemoteUserMailbox' {write-host "(Rmbx)" -nonewline}
                        # 8:53 AM 10/9/2024 add to cover mbx2shared conversion results
                        'RemoteSharedMailbox' {write-host "(Rmbx *SHARED*)" -nonewline}
                        # 12:23 PM 12/26/2024 add resource & remote res's
                        'RemoteRoomMailbox' {write-host "(Rmbx *ROOM*)" -nonewline}
                        'RemoteEquipmentMailbox' {write-host "(Rmbx *EQUIP*)" -nonewline}
                        'UserMailbox' {write-host "(xMbx)" -nonewline}
                        'SharedMailbox' {write-host "(xSMbx)" -nonewline}
                        'RoomMailbox' {write-host "(xRoomMbx)" -nonewline}
                        'EquipmentMailbox' {write-host "(xEquipMbx)" -nonewline}
                        # no rmbx, but remote obj?
                        'MailUser' {
                            $smsg = "xMAILUSER WO MBX DETECTED! - POSSIBLE NOBRAIN?"
                            write-warning $smsg
                            #$hsum.isNoBrain = $true ;
                        }
                        "GuestMailUser" {
                            $smsg = "xGuestMailUser detected, likely external forest/Inet Guest!"
                            write-warning $smsg
                        } ;
                        'MailUniversalDistributionGroup' {write-host "(xDG)" -nonewline}
                        'DynamicDistributionGroup'  {write-host "(xDDG)" -nonewline}
                        'MailContact' {write-host "(xMC)" -nonewline]}
                        default{
                            #$smsg = "Unable to resolve `$hSum.xoRcp.recipienttypedetails:$($hSum.xoRcp.recipienttypedetails)" ; 
                            $smsg = "Unable to resolve `$hSum.xoRcp.recipienttypedetails:$($tmpxRcp.OPRcp.recipienttypedetails)" ; 
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                            throw $smsg ; 
                            break ; 
                        }
                    }
                }  # loop-E
            } ; # if-E
            #-=-=-=-=-=-=-=-=
            # new rules, with INT/VEN AADU anchored to ADU, but xoMbx anchored solely to AADU (and not OP rcp), it's possible to completely fail onprem get-recipient, and still have a functional mailbox in cloud, that's operating properly.

            $abortReport = $false ; 
            if( ($hSum.OPRcp -OR $hSum.xoRcp) -AND ( ($hSum.OPRcp -is [array]) -AND ($hSum.xoRcp -is [array]) ) ){
                # failed to isolate both op & xo unique recip
                $abortReport = $true ;
            }elseif( ($hSum.OPRcp -OR $hSum.xoRcp) -AND ( ($hSum.xoRcp -isnot [array]) -AND ($hSum.OPRcp -is [array] ) ) ){
                # single cloud, mult onprem -> could be non-hybrid cloud-first recip
                $abortReport = $false ;
            }elseif( ($hSum.OPRcp -OR $hSum.xoRcp) -AND ( ($hSum.OPRcp -isnot [array]) -AND ($hSum.xoRcp -is [array]) ) ){
                # single OP recip, mult cloud; could be legit unonboarded OP rcp
                $abortReport = $false ; 
            } ; 

            if($abortReport){
                $smsg = "`n`n==RecipientArray(s) detected:"
                $smsg += "`nDumping initial OP & XO RecipientLists"
                $smsg += "`nto permit you to winnow down a single targeted user from the returns,"
                $smsg += "`nfor a fresh targeted pass!`n`n" ; 
                #if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                #else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                write-hostCallOutTDO -Object $smsg -Type Warning -Nowrap ;

                $smsg = "`$hSum.OPRcp match(es):`n$(($hSum.OPRcp | ft -a $prpFTARcp |out-string).trim())`n" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                $smsg = "`$hSum.xoRcp match(es):`n$(($hSum.xoRcp | ft -a $prpFTARcp |out-string).trim())`n" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                BREAK ; 
            } ; 

            if($hSum.OPRcp){
                # 9:41 AM 10/9/2024 with array loops we need to accomodate, and aggregate - or it throws errors tying to get-remotemailbox -id [array]
                # also need to += all assigns to acomodate both lookups, not just the last one
                if($hSum.OPRcp -is [array]){
                    $smsg = "ARRAY of OPRcps! Inconcistent results will be returned on attempts following, likely errors!" ; 
                    $smsg += "`n(need to isolate single specific identifier from these outputs, and rerun fresh pass)" ; 
                    write-warning $smsg ; 
                } ; 
                 $hSum.OPRcp | ForEach-Object{
                    $tmpRcp = $_ ; 
                    $error.clear() ;
                    TRY {
                        switch -regex ($tmpRcp.recipienttype){
                            "UserMailbox" {
                                write-verbose "'UserMailbox':get-mailbox $($tmpRcp.identity)"
                                $bufferRcp = $null ; 
                                $bufferRcp = get-mailbox $tmpRcp.identity -resultsize $MaxRecips | select -first $MaxRecips ; 
                                #if($hSum.OPMailbox += get-mailbox $tmpRcp.identity -resultsize $MaxRecips | select -first $MaxRecips ){ ;
                                if($bufferRcp){
                                    $hSum.OPMailbox += $bufferRcp ; 
                                    #write-verbose "`$hSum.OPMailbox:`n$(($hSum.OPMailbox|ft -a |out-string).trim())" ;
                                    if($outObject){

                                    } else {
                                        #$Rpt += $hSum.OPMailbox.primarysmtpaddress ;
                                        $Rpt += $bufferRcp.primarysmtpaddress
                                    } ;
                                    write-verbose "'UserMailbox':Test-MAPIConnectivity -identity $($hSum.OPMailbox.userprincipalname)"
                                    $bufferRcp = $null ; 
                                    $bufferRcp  =  Test-MAPIConnectivity -identity $hSum.OPMailbox.userprincipalname ;
                                    if($bufferRcp){
                                        $hSum.OPMapiTest  += $bufferRcp ;
                                        $smsg = "Outlook (MAPI) Access Test Result:$($bufferRcp.OPMapiTest.result)" ;
                                        if($bufferRcp.OPMapiTest.result -eq 'Success'){
                                            write-host -foregroundcolor green $smsg ;
                                        } else {
                                            write-WARNING $smsg ;
                                        } ;
                                    } else { 
                                        write-warning "Failed to return Test-MAPIConnectivity -identity $($hSum.OPMailbox.userprincipalname) !" ; 
                                    } ; 
                                } ;
                            }
                            "MailUser" {
                                write-verbose "'MailUser':get-remotemailbox $($tmpRcp.identity)"
                                $bufferRcp = $null ; 
                                $bufferRcp  = get-remotemailbox $tmpRcp.identity -resultsize $MaxRecips -ea 0 | select -first $MaxRecips ; 
                                #if($hSum.OPRemoteMailbox += get-remotemailbox $tmpRcp.identity -resultsize $MaxRecips -ea 0 | select -first $MaxRecips){
                                if($bufferRcp){
                                    $hSum.OPRemoteMailbox += $bufferRcp ; 
                                    write-verbose "`$hSum.OPRemoteMailbox:`n$(($hSum.OPRemoteMailbox|ft -a |out-string).trim())" ;
                                }else{
                                    $smsg = "RecipientTypeDetails:MailUser with NO Rmbx! (NoBrain?)" ;
                                    write-warning $smsg ;
                                    if($hsum.xoRcp.ExternalDirectoryObjectId){
                                        # of course has match to AADU  - always does - we're going to need the AADU before we can lookup the ADU
                                        # $pltGadu.identity  +=  $hSum.AADUser.ImmutableId | convert-ImmuntableIDToGUID | select -expand guid ;
                                        caad  -Verbose:$false -silent ;
                                        write-verbose "OPRcp:Mailuser, ensure GET-ADUSER pulls AADUser.matched object for cloud recipient:`nfallback:get-AzureAdUser  -objectid $($hsum.xoRcp.ExternalDirectoryObjectId)" ;
                                        # have to postfilter, if want specific count -maxresults catch's with no $error[0]
                                        $hSum.AADUser   +=  get-AzureAdUser  -objectid $hsum.xoRcp.ExternalDirectoryObjectId | select -first $MaxRecips;  ;
                                    } else {
                                        throw "Unsupported object, blank `$hsum.xoRcp.ExternalDirectoryObjectId!" ;
                                    } ;
                                }
                                if($outObject){

                                } else {
                                    $Rpt += $hSum.OPRemoteMailbox.primarysmtpaddress ;
                                } ;
                            } ;
                            default {
                                write-warning "$((get-date).ToString('HH:mm:ss')):Unsupported RecipientType:($tmpRcp.recipienttype). EXITING!" ;
                                Break ;
                            }
                        }
                        <# get-aduser docs say REsultSetSize is documented,
                        [Get-ADUser (ActiveDirectory) | Microsoft Docs - docs.microsoft.com/](https://docs.microsoft.com/en-us/powershell/module/activedirectory/get-aduser?view=windowsserver2019-ps)
                         but use of it throws: Parameter set cannot be resolved using the specified named parameters.
                         pull it and post filter to 1...
                        #>
                        #ResultSetSize = $MaxRecips
                        #$pltGadu=[ordered]@{Identity = $null ; Properties='*' ;errorAction='STOP' ; } ;
                        $pltGadu=[ordered]@{Identity = $null ; Properties=$propsADU ;errorAction='STOP' ; } ;
                        if($hSum.OPRemoteMailbox ){
                            # get-aduser dox but doesn't really support ResultSetSize, post filter for it.
                            $pltGadu.identity = $hSum.OPRemoteMailbox.samaccountname ;
                        }elseif($hSum.OPMailbox){
                            $pltGadu.identity = $hSum.OPMailbox.samaccountname ;
                        } else {
                            # cloud-first or no brain, neither oprmbx or opmailbox;  should have populated $hSum.AADUser above, use immutable lookup
                            if($hSum.AADUser.DirSyncEnabled){
                                $smsg = "Falling back to AADU Immutable lookup to locate replicated adu source" ;
                                if($pltGadu.identity = $hSum.AADUser.ImmutableId | convert-ImmuntableIDToGUID | select -expand guid){
                                    $smsg = "(Resolved AADU.Immutable ->GUID:$($pltGadu.identity))" ;
                                    write-verbose $smsg ;
                                }else {
                                    $smsg = "UNABLE TO RESOLVE ADU.IMMUTABLEID TO ADU GUID!"
                                    write-warning $smsg ;
                                    throw $smsg ;
                                }
                            } else {
                                $smsg = "$AADUsuer not DirSyncEnabled: CLOUD FIRST!"
                                write-warning $smsg ;
                                #throw $smsg ;
                            } ;
                        };
                        if($pltGadu.identity){
                            <# this is throwing a blank fail
                            WARNING: 15:04:18:Failed processing .
                            Error Message:
                            Error Details:
                            # and dumping balance of processing
                            issue: was in adms drive: :xxxx, gadu was searching root domain only
                            so it was a search fail, throwing an error, but didn't return details. Still good idea to trap not found and echo it
                            #>
                            #$hSum.ADUser  += Get-ADUser @pltGadu ;
                            write-verbose "Get-ADUser w`n$(($pltGadu|out-string).trim())" ;
                            # try a nested local trycatch, against a missing result
                            Try {
                                #Get-ADUser $DN -ErrorAction Stop ;
                                $hSum.ADUser  += Get-ADUser @pltGadu | select -first $MaxRecips ;
                            } Catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
                                write-warning "(no matching ADuser found:$($pltGadu.identity))" ;
                            } catch {
                                $ErrTrapd=$Error[0] ;
                                $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                                else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                Continue ;
                            } ;

                            write-verbose "`$hSum.ADUser:`n$(($hSum.ADUser|fl $propsADU| out-string).trim())" ;
                            $smsg = "(TOR USER, fed:$($TORMeta.adforestname))" ;
                            $hSum.Federator = $TORMeta.adforestname ;
                            write-host -Fore yellow $smsg ;
                        
                            <#
                            if($hSum.OPRemoteMailbox){
                                $smsg = "$(($hSum.OPRemoteMailbox |fl $propsMailx|out-string).trim())"
                            } ;
                            if($hSum.OPMailbox){
                                $smsg =  "$(($hSum.OPMailbox |fl $propsMailx|out-string).trim())" ;
                            } ;
                            #>
                            # swap to md tbl fmt
                            if($hSum.OPRemoteMailbox){$MailRecip = $hSum.OPRemoteMailbox } ; 
                            if($hSum.OPMailbox){$MailRecip = $hSum.OPMailbox } ; 
                            $smsg = "$(($MailRecip| select $propsMailxL1 |out-markdowntable @MDtbl |out-string).trim())" ;
                            $smsg += "`n$(($MailRecip|select $propsMailxL2 |out-markdowntable @MDtbl|out-string).trim())" ;
                            $smsg += "`n$(($MailRecip|select $propsMailxL3 |out-markdowntable @MDtbl|out-string).trim())" ;
                            #$smsg += "`n$(($MailRecip|select $propsMailxL4 |out-markdowntable @MDtbl|out-string).trim())" ;
                            #$smsg += "`n$(($MailRecip|select $propsMailxL4 | fl |out-string).trim())" ;
                            # drop L4 it's DN, which is already in ADU md tbl
                            # flip dn L4 to fl (suppress crlf)

                            write-host $smsg ;
                            #if($MailRecip.ForwardingAddress){
                            #    $smsg += "`n$(($MailRecip|select $propsMailxL5 |out-markdowntable @MDtbl|out-string).trim())" ;
                            #} ; 
                            <#
                            if($hSum.OPRemoteMailbox -AND $hSum.OPRemoteMailbox.ForwardingAddress){
                                write-host $smsg ; # write pending primary (using ww on next)
                                #$smsg = "==FORWARDED rMBX!:`n$(($hSum.OPRemoteMailbox  |ft -a ForwardingAddress,DeliverToMailboxAndForward,ForwardingSmtpAddress|out-string).trim())" ;
                                $smsg = "==FORWARDED rMBX!:" ; 
                                $smsg += "`n$(($MailRecip|select $propsMailxL5 |out-markdowntable @MDtbl|out-string).trim())" ;
                            } ;
                            if($hSum.OPMailbox -AND $hSum.OPMailbox.ForwardingAddress){
                                write-host $smsg ; # write pending primary (using ww on next)
                                $smsg = "==FORWARDED opMBX!:`n$(($hSum.OPMailbox |ft -a ForwardingAddress,DeliverToMailboxAndForward,ForwardingSmtpAddress|out-string).trim())" ;
                            } ;
                            #>
                            if($hSum.OPRemoteMailbox.ForwardingAddress -OR $hSum.OPMailbox.ForwardingAddress){
                                write-host $smsg ; # echo pending, using ww below
                                $smsg = "==FORWARDED rMBX!:" ; 
                                $smsg += "`n$(($MailRecip|select $propsMailxL5 |out-markdowntable @MDtbl|out-string).trim())" ;
                                write-warning $smsg ;
                            } ; 

                            #$smsg += "`n$(($hSum.ADUser |fl $propsADUsht  |out-string).trim())"
                            # these are already in the ADU md tbl dump, drop them
                            #$smsg = "$(($hSum.ADUser |fl $propsADUsht  |out-string).trim())"
                            #write-host $smsg ;
                        } ;
                    } CATCH {
                        $ErrTrapd=$Error[0] ;
                        $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                        else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        Continue ;
                    } ;
                }  # loop-E
            }elseif($hSum.xoRcp){
                foreach($txR in $hSum.xoRcp){
                    TRY {
                        switch -regex ($txR.recipienttypedetails){
                            "UserMailbox" {
                                #write-verbose "$((get-alias ps1GetxMbx).definition) w`n$(($pltGMailObj|out-string).trim())" ;
                                write-verbose "get-exomailbox w`n$(($pltGMailObj|out-string).trim())" ;
                                if($hSum.xoMailbox += get-xomailbox @pltGMailObj -ea 0 | select -first $MaxRecips ){
                                    write-verbose "`$hSum.xoMailbox:`n$(($hSum.xoMailbox|out-string).trim())" ;
                                    if($outObject){

                                    } else {
                                        $Rpt += $hSum.xoMailbox.primarysmtpaddress ;
                                    } ;
                                    if($hSum.xoMailbox -is [system.array]){
                                        write-warning "Multiple mailboxes matched!" ;
                                    } ;
                                    # accomodate array returned (multiple matches):
                                    $ino = 0 ;
                                    foreach($xmbx in $hSum.xoMailbox){
                                        $ino++ ;
                                        if($hSum.xoMailbox -isnot [system.array]){
                                            $smsg = "xmbx$($ino):$($xmbx.userprincipalname)" ;
                                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                        } ;
                                        write-verbose "'xoUserMailbox':Test-exoMAPIConnectivity $($xmbx.userprincipalname)"
                                        $hSum.xoMapiTest  +=  Test-xoMAPIConnectivity -identity $xmbx.userprincipalname ;
                                        $smsg = "Outlook (xoMAPI) Access Test Result:$($hsum.xoMapiTest.result)" ;
                                        if($hsum.xoMapiTest.result -eq 'Success'){
                                            write-host -foregroundcolor green $smsg ;
                                        } else {
                                            write-WARNING $smsg ;
                                        } ;
                                        if($getMobile){
                                            # $devstats = Get-exoMobileDeviceStatistics -Mailbox UPN
                                            #$smsg = "'xoMobileDeviceStats':$((get-alias ps1GetxMobilDevStats).definition) -Mailbox $($xmbx.userprincipalname)"
                                            $smsg = "'xoMobileDeviceStats':Get-xoMobileDeviceStatistics -Mailbox $($xmbx.userprincipalname)"
                                            if($verbose){
                                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                                                else{ write-verbose $smsg } ; 
                                            } ; 
                                            $hsum.xoMobileDeviceStats  +=  Get-xoMobileDeviceStatistics -Mailbox $xmbx.userprincipalname -ea STOP ; 
                                            $smsg = "xoMobileDeviceStats Count:$(($hsum.xoMapiTest|measure).count)" ;
                                            write-host -foregroundcolor green $smsg ;
                                        } ; 
                                        if($getQuotaUsage){
                                            $pltGMbxStatX=[ordered]@{
                                                identity = $hSum.xoMailbox.exchangeguid ;
                                                ErrorAction = 'STOP' ; 
                                            } ;
                                            $smsg = "Get-xoMailboxStatistics  w`n$(($pltGMbxStatX|out-string).trim())"
                                            if($verbose){
                                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                                                else{ write-verbose $smsg } ; 
                                            } ; 
                                            $hSum.xoMailboxStats  +=  Get-xoMailboxStatistics @pltGMbxStatX | select $prpStat;
                                            $smsg = "xoMailboxStats Count:$(($hsum.xoMapiTest|measure).count)" ;
                                            write-host -foregroundcolor green $smsg ;

                                            If($hSum.xoMailbox.UseDatabaseQuotaDefaults){
                                                $hSum.xoEffectiveQuotas  +=  $hSum.xoMailboxStats | select @{N ='IssueWarningQuotaMB'; e={$_.DBIssueWarningQuotaMB}},
                                                @{n='ProhibitSendQuotaMB'; e={$_.DBProhibitSendQuotaMB}},
                                                @{n='ProhibitSendReceiveQuotaMB';e={$_.DBProhibitSendReceiveQuotaMB}}; 
                                            } else {
                                                $hSum.xoEffectiveQuotas  +=  $hSum.xoMailbox | select @{n="IssueWarningQuotaMB";e={[math]::round($_.IssueWarningQuota.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1MB,2)}},
                                                @{n="ProhibitSendQuotaMB";e={[math]::round($_.ProhibitSendQuota.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1MB,2)}},
                                                @{n="ProhibitSendReceiveQuotaMB";e={[math]::round($_.ProhibitSendReceiveQuota.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1MB,2)}} ;
                                            } ;  
                                            $hSum.xoNetOfSendReceiveQuotaMB  +=  $hSum.xoEffectiveQuotas.ProhibitSendQuotaMB - $hSum.xoMailboxStats.TotalMailboxSizeMB ; 

                                            $pltGMbxStatX.add('IncludeOldestAndNewestItems',$true) ; 
                                            $smsg = "Get-xoMailboxFolderStatistics  w`n$(($pltGMbxStatX|out-string).trim())" ;
                                            if($verbose){
                                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                                                else{ write-verbose $smsg } ; 
                                            } ; 
                                            TRY{
                                                $hsum.xoMailboxFolderStats  +=  Get-xoMailboxFolderStatistics @pltGMbxStatX  ;

                                                $smsg = "Export FolderStats to`n$(($ofMbxFolderStats|out-string).trim())" ;
                                                if($verbose){
                                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                                                    else{ write-verbose $smsg } ; 
                                                } ; 
                                                $hsum.xoMailboxFolderStats | ?{$_.ItemsInFolder -gt 0 -AND $_.identity -notmatch $rgxHiddn } | 
                                                    select $prpFldr | sort SizeMB -desc | export-csv  -path $ofMbxFolderStats -notype ;

                                            } CATCH {
                                                $ErrTrapd=$Error[0] ;
                                                $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                                                write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" ;
                                            } ; 
                                        } ; 
                                    } ;
                                    break ;
                                } ;
                            }
                            "MailUser" {
                                # external mail recipient, *not* in TTC - likely in other rgs, and migrated to remote EXOP enviro
                                #$hSum.OPRemoteMailbox += get-remotemailbox $txR.identity  ;
                                caad -silent -verbose:$false ;
                                #write-verbose "`$txR | $((get-alias ps1GetxMUsr).definition)..." ;
                                write-verbose "`$txR | Get-xoMailUser..." ;
                                $hSum.xoMUser  +=  $txR | Get-xoMailUser -ResultSize $MaxRecips | select -first $MaxRecips ;
                                #write-verbose "`$txR | $((get-alias ps1GetxUser).definition)..." ;
                                write-verbose "`$txR | get-xoUser ..." ;
                                $hSum.xoUser  +=  $txR | get-xouser -ResultSize $MaxRecips | select -first $MaxRecips ;
                                write-verbose "`$hSum.xoUser:`n$(($hSum.xoUser|ft -a |out-string).trim())" ;
                                #write-verbose "get-AzureAdUser  -objectid $($hSum.xoUser.userPrincipalName)" ;
                                #$hSum.AADUser   +=  get-AzureAdUser  -objectid $hSum.xoMUser.userPrincipalName -Top $MaxRecips ;
                                write-verbose "`$hSum.xoMUser:`n$(($hSum.xoMUser|ft -a |out-string).trim())" ;
                                #$Rpt += $hSum.OPRemoteMailbox.primarysmtpaddress ;
                                write-host "$($txR.ExternalEmailAddress): matches a MailUser object with UPN:$($hSum.xoMUser.userPrincipalName)" ;
                                if($outObject){

                                } else {
                                    $Rpt += $hSum.xoMUser.primarysmtpaddress ;
                                } ;
                                break ;
                            } ;
                            "GuestMailUser" {
                                #$hSum.OPRemoteMailbox += get-remotemailbox $txR.identity  ;
                                caad -silent -verbose:$false ;
                                #write-verbose "`$txR | $((get-alias ps1GetxUser).definition)..." ;
                                write-verbose "`$txR | get-xoUser..." ; 
                                $hSum.xoUser  +=  $txR | get-xouser -ResultSize $MaxRecips | select -first $MaxRecips ;
                                write-verbose "`$hSum.xoUser:`n$(($hSum.xoUser|ft -a |out-string).trim())" ;
                                write-verbose "get-AzureAdUser  -objectid $($hSum.xoUser.userPrincipalName)" ;
                                $hSum.txGuest  +=  get-AzureAdUser  -objectid $hSum.xoUser.userPrincipalName -Top $MaxRecips | select -first $MaxRecips ;
                                write-verbose "`$hSum.txGuest:`n$(($hSum.txGuest|ft -a |out-string).trim())" ;
                                #$Rpt += $hSum.OPRemoteMailbox.primarysmtpaddress ;
                                write-host "$($txR.ExternalEmailAddress): matches a Guest object with UPN:$($hSum.xoUser.userPrincipalName)" ;
                                if($hSum.txGuest.EmailAddresses -eq $null){
                                    write-warning "Guest appears to have damage from conficting replicated onprem MailContact, as it's EmailAddresses property is *blank*" ;
                                } ;
                                break ;
                            } ;
                            "MailContact" {
                                #$hSum.OPRemoteMailbox += get-remotemailbox $txR.identity  ;
                                #$Rpt += $hSum.OPRemoteMailbox.primarysmtpaddress ;
                                write-host "$($txR.primarysmtpaddress): matches an EXO MailContact with external Email: $($txR.primarysmtpaddress)" ;
                                break ;
                            } ;
                            "MailUniversalSecurityGroup" {
                                #$hSum.OPRemoteMailbox += get-remotemailbox $txR.identity  ;
                                #$Rpt += $hSum.OPRemoteMailbox.primarysmtpaddress ;
                                write-host "$($txR.primarysmtpaddress): matches an EXO MailUniversalSecurityGroup with Dname: $($txR.displayname)" ;
                                break ;
                            } ;
                            default {
                                write-warning "$((get-date).ToString('HH:mm:ss')):Unsupported RecipientType:($hSum.OPRcp.recipienttype). EXITING!" ;
                                Break ;
                            }
                        }
                    } CATCH {
                        $ErrTrapd=$Error[0] ;
                        $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                        else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        Continue ;
                    } ;
                }  # loop-E $txR
                # contacts and guests won't drop with $hSum.OPRemoteMailbox or $hSum.OPMailbox populated
                TRY {
                    $pltGadu=[ordered]@{Identity = $null ; Properties='*' ;errorAction='SilentlyContinue'} ;
                    if($hSum.OPRemoteMailbox ){
                        $pltGadu.identity = $hSum.OPRemoteMailbox.samaccountname;
                    }elseif($hSum.OPMailbox){
                        $pltGadu.identity = $hSum.OPMailbox.samaccountname ;
                    }elseif(-not $hsum.OPRcp -AND $hsum.xorcp -AND $hSum.xomailbox){
                        $smsg = "XOMailbox with NO OPRcp/Rmbx/MUser" ;
                        write-host -foregroundcolor yellow $smsg ;
                        if($hsum.xoRcp.ExternalDirectoryObjectId){
                            # of course has match to AADU  - always does - we're going to need the AADU before we can lookup the ADU
                            if(-not $hSum.AADUser){
                                # $pltGadu.identity  +=  $hSum.AADUser.ImmutableId | convert-ImmuntableIDToGUID | select -expand guid ;
                                Connect-AAD -Verbose:$false -silent ;
                                write-verbose "xoMailbox: ensure GET-ADUSER pulls AADUser.matched object for cloud recipient:`nfallback:get-AzureAdUser  -objectid $($hsum.xoRcp.ExternalDirectoryObjectId)" ;
                                # have to postfilter, if want specific count -maxresults catch's with no $error[0]
                                $hSum.AADUser   +=  get-AzureAdUser  -objectid $hsum.xoRcp.ExternalDirectoryObjectId | select -first $MaxRecips;  ;
                            } ; 
                        } else {
                            throw "Unsupported object, blank `$hsum.xoRcp.ExternalDirectoryObjectId!" ;
                        } ;
                        if($hSum.xomailbox.IsDirSynced){
                            # doesn't mean hybrid exchange obj, means ADU anchored object
                            write-host "XOMailbox.IsDirSynced: anchored to ADUser" ; 
                            if($hSum.AADUser.ExtensionProperty.onPremisesDistinguishedName){
                                switch -regex ($hSum.AADUser.ExtensionProperty.onPremisesDistinguishedName){
                                    $rgxADDistNameAT{
                                        $pltGadu.identity = $hSum.AADUser.ExtensionProperty.onPremisesDistinguishedName ; 
                                        $pltGadu.add('server',(($hSum.AADUser.ExtensionProperty.onPremisesDistinguishedName.split(',') | ?{$_ -match 'DC='} ) -replace 'DC=','') -join '.') ; 
                                    }
                                    default{
                                        $smsg = "Unrecognized AADUser.ExtensionProperty.onPremisesDistinguishedName!" ; 
                                        $smsg += "`n$($hSum.AADUser.ExtensionProperty.onPremisesDistinguishedName)" ; 
                                        throw $smsg ;
                                    }
                                } ; 
                            } else {
                                write-warning "blank AADUser.ExtensionProperty.onPremisesDistinguishedName! (non-ADUser-sync'd object)" ; 
                            } ;  
                        }else{
                            write-warning "xomailbox is *NOT* IsDirSynced!, Cloud-first recipient, wo anchored AzureADUser object!" ; 
                        }
                    } else {
                        write-warning "NO FUNCTIONAL COMBO OF OPRcp xoRcp OR xoMailbox!" ; 
                    };
                    if($pltGadu.identity){
                        write-verbose "Get-ADUser w`n$(($pltGadu|out-string).trim())" ;
                        # try a nested local trycatch, against a missing result
                        Try {
                            #Get-ADUser $DN -ErrorAction Stop ;
                            $hSum.ADUser  += Get-ADUser @pltGadu | select -first $MaxRecips ;
                        } Catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
                            write-warning "(no matching ADuser found:$($pltGadu.identity))" ;
                        } catch {
                            $ErrTrapd=$Error[0] ;
                            $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                            else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            Continue ;
                        } ;

                        write-verbose "`$hSum.ADUser:`n$(($hSum.ADUser|fl $propsADU | out-string).trim())" ;
                        $smsg = "(TOR USER, fed:$($TORMeta.adforestname))" ;
                        $hSum.Federator  +=  $TORMeta.adforestname ;
                        write-host -Fore yellow $smsg ;
                        if($hSum.OPRemoteMailbox){
                            $smsg = "$(($hSum.OPRemoteMailbox |fl $propsMailx|out-string).trim())"
                            #$smsg += "`n-Title:$($hSum.ADUser.Title)"
                            $smsg += "`n$(($hSum.ADUser |fl 'Enabled','Description','whenCreated','whenChanged','Title' |out-string).trim())"
                        } ;
                        if($hSum.OPMailbox){
                            $smsg =  "$(($hSum.OPMailbox |fl $propsMailx|out-string).trim())" ;
                            $smsg += "`n$(($hSum.ADUser |fl 'Enabled','Description','whenCreated','whenChanged','Title' |out-string).trim())"
                        } ; 
                        if( -not $hsum.OPRcp -AND $hsum.xoRcp -AND $hsum.xomailbox){ 
                            $smsg = "Cloud Mailbox is nonDirSync'd NON-HYBRID mail object!" ; 
                            $smsg += "`n$(($hSum.xoMailbox |fl $propsMailx|out-string).trim())" ; 
                            if($hsum.ADUser){
                                if($hsum.Aaduser.DirSyncEnabled){
                                    $smsg += "`nbut ADUser Object IS dirsync'd to AzureADUser object" ; 
                                } else { 
                                    $smsg += "`nADUser Object IS NOT dirsync'd to AzureADUser object" ; 
                                } ; 
                                $smsg += "`n$(($hSum.ADUser |fl 'Enabled','Description','whenCreated','whenChanged','Title' |out-string).trim())"
                            } else {
                                $smsg += "`nNO ADUser Object appears to be cloud-first AADUser object" ;
                            }; 
                        } ;
                        if( -not $hsum.OPRcp -AND -not $hsum.xoRcp -AND $hSum.ADUser -AND $hsum.Aaduser){
                            $smsg = "No detected OnPrem or Cloud Mail Recipient Objects detected" ; 
                            if($hSum.ADUser){
                                $smsg += "`nADUser Object IS NOT dirsync'd to AzureADUser object" ; 
                            } ; 
                            if($hsum.Aaduser.DirSyncEnabled){
                                $smsg += "`nbut ADUser Object IS dirsync'd to AzureADUser object" ; 
                            } ; 
                        } ; 
                        write-host $smsg ;
                    } ;
                } CATCH {
                    $ErrTrapd=$Error[0] ;
                    $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    Continue ;
                } ;

                if($outObject){

                } else {
                    $Rpt += $hSum.xoMailbox.primarysmtpaddress ;
                } ;
                $ino = 0 ;
                foreach($xmbx in $hSum.xoMailbox){
                    $ino++;
                    if($hSum.xoMailbox -isnot [system.array]){
                        $smsg = "xmbx$($ino):$($xmbx.userprincipalname)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ;
                    if($xmbx.isdirsynced){
                        # can be federated to VEN|CMW|Toro
                        switch -regex ($xmbx.primarysmtpaddress.split('@')[1]){
                            $CMWMeta.rgxOPFederatedDom {
                                $smsg="(CMW USER, fed:$($CMWMeta.adforestname))" ;
                                $hSum.Federator  +=  $CMWMeta.adforestname ;
                            }
                            $TORMeta.rgxOPFederatedDom {
                                $smsg="(TOR USER, fed:$($TORMeta.adforestname))" ;
                                $hSum.Federator  +=  $TORMeta.adforestname ;
                            }
                            $VENMeta.rgxOPFederatedDom {
                                $smsg="(VEN USER, fed:$($venmeta.o365_TenantLabel))" ;
                                $hSum.Federator  +=  $VENMETA.o365_TenantLabel ;
                            }
                            $INTMeta.rgxOPFederatedDom {
                                $smsg="(INT USER, fed:$($INTmeta.o365_TenantLabel))" ;
                                $hSum.Federator  +=  $INTMETA.o365_TenantLabel ;
                            }

                        } ;
                    } elseif($hSum.xoMuser.IsDirSynced){
                        switch -regex ($xmbx.primarysmtpaddress.split('@')[1]){
                            $CMWMeta.rgxOPFederatedDom {
                                $smsg="(CMW USER, fed:$($CMWMeta.adforestname))" ;
                                $hSum.Federator  +=  $CMWMeta.adforestname ;
                            }
                            $TORMeta.rgxOPFederatedDom {
                                $smsg="(TOR USER, fed:$($TORMeta.adforestname))" ;
                                $hSum.Federator  +=  $TORMeta.adforestname ;
                            }
                            $VENMeta.rgxOPFederatedDom {
                                $smsg="(VEN USER, fed:$($venmeta.o365_TenantLabel))" ;
                                $hSum.Federator  +=  $VENMETA.o365_TenantLabel ;
                            }
                            $INTMeta.rgxOPFederatedDom {
                                $smsg="(INT USER, fed:$($INTmeta.o365_TenantLabel))" ;
                                $hSum.Federator  +=  $INTMETA.o365_TenantLabel ;
                            }
                        } ;
                    }else{
                        [regex]$rgxTenDom = [regex]::escape("@$($tormeta.o365_TenantDomain)")
                        if($hsum.xoRcp.primarysmtpaddress -match $rgxTenDom){
                                $smsg="(CLOUD-1ST ACCT, unfederated)" ;
                                $hSum.Federator  +=  $TORMeta.o365_TenantDom ;

                        } else {
                            $smsg="(CLOUD-1ST ACCT, unfederated)" ;
                            $hSum.Federator  +=  $TORMeta.o365_TenantDom ;
                        } ;
                    } ;
                } ;  # loop-E
                write-host -Fore yellow $smsg ;
                # skip user lookup if guest already pulled it
                if(-not $hSum.xoUser){
                    $ino = 0 ;
                    foreach($xmbx in $hSum.xoMailbox){
                        #write-verbose "$((get-alias ps1GetxUser).definition) -id $($xmbx.UserPrincipalName)"
                        write-verbose "get-xoUser -id $($xmbx.UserPrincipalName)"
                        $hSum.xoUser += get-xouser -id $xmbx.UserPrincipalName -ResultSize $MaxRecips ;
                        write-verbose "`$hSum.xoUser:`n$(($hSum.xoUser|ft -a |out-string).trim())" ;
                    } ;
                }
                if($hSum.xoMailbox){
                    $ino = 0 ;
                    foreach($xmbx in $hSum.xoMailbox){
                        $ino++ ;
                        if($hSum.xoMailbox -isnot [system.array]){
                            $smsg = "xmbx$($ino):$($xmbx.userprincipalname)" ;
                            write-host $smsg ;
                        } ;
                        write-host -foreground yellow "=get-xMbx:> " -nonewline;
                        write-host "$(($hSum.xoMailbox |fl ($propsMailx |?{$_ -notmatch '(sam.*|dist.*)'})|out-string).trim())`n-Title:$($hSum.xoUser.Title)";
                    } ;

                    if($getMobile){
                        write-host -foreground yellow "===`$hsum.xoMobileDeviceStats: " #-nonewline;
                        $ino = 0 ;
                        foreach($xmob in $hsum.xoMobileDeviceStats){
                            $ino++ ;
                            <#if($hsum.xoMobileDeviceStats -isnot [system.array]){
                                $smsg = "xmob$($ino):$($xmob.userprincipalname)" ;
                                write-host $smsg ;
                            } ;
                            write-host -foreground yellow "=get-xMob:> " -nonewline;
                            write-host "$(($xmob.userprincipalname |fl ($propsMailx |?{$_ -notmatch '(sam.*|dist.*)'})|out-string).trim())`n-Title:$($hSum.xoUser.Title)";
                            #>
                            if($hsum.xoMobileDeviceStats -is [system.array]){
                                 write-host -foreground yellow "=get-xMob$($ino):> " #-nonewline;
                            } else { 
                                write-host -foreground yellow "=get-xMobileDev:> " #-nonewline;
                            } ; 
                            $smsg = "$(($xmob | select $propsMobL1 |out-markdowntable @MDtbl |out-string).trim())" ;
                            $smsg += "`n$(($xmob | select $propsMobL2 |out-markdowntable @MDtbl |out-string).trim())" ;
                            write-host $smsg ;
                        } ;

                    } ; 

                }elseif($hSum.xoMUser){
                    write-host "=get-xMUSR:>`n$(($hSum.xoMUser |fl ($propsMailx |?{$_ -notmatch '(sam.*|dist.*)'})|out-string).trim())`n-Title:$($hSum.xoUser.Title)";
                }elseif($hSum.txGuest){
                    write-host "=get-AADU:>`n$(($hSum.txGuest |fl userp*,PhysicalDeliveryOfficeName,JobTitle|out-string).trim())"
                } ;
                TRY {
                    #write-verbose "$((get-alias ps1GetxRcp).definition) -Filter {Members -eq '$($hSum.xoUser.DistinguishedName)'}`n -RecipientTypeDetails GroupMailbox,MailUniversalDistributionGroup,MailUniversalSecurityGroup"
                    write-verbose "get-xorecipient -Filter {Members -eq '$($hSum.xoUser.DistinguishedName)'}`n -RecipientTypeDetails GroupMailbox,MailUniversalDistributionGroup,MailUniversalSecurityGroup"
                    $hSum.xoMemberOf  +=  get-xorecipient -Filter "Members -eq '$($hSum.xoUser.DistinguishedName)'" -RecipientTypeDetails GroupMailbox,MailUniversalDistributionGroup,MailUniversalSecurityGroup ;
                    write-verbose "`$hSum.xoMemberOf:`n$(($hSum.xoMemberOf|out-string).trim())" ;
                } CATCH {
                    $ErrTrapd=$Error[0] ;
                    $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    Continue ;
                } ;
            } else {
                write-warning "(no matching EXOP or EXO recipient object:$($usr))"
                # do near Lname[0-3]* searches for comparison
                if($hSum.lname){
                    write-warning "Lname ($($hSum.lname) parsed from input),`nattempting similar LName g-rcp:...`n(up to `$MaxRecips:$($MaxRecips))" ;
                    $lname = $hsum.lname ;
                    #$fltrB = "displayname -like '*$lname*'" ;
                    #write-verbose "RETRY:get-recipient -filter {$($fltr)}" ;
                    #get-recipient "$($txusr.lastname.substring(0,3))*"| sort name
                    $substring = "$($hSum.lname.substring(0,3))*"
                    

                    write-host "get-recipient -id $($substring) -ea 0 |?{$_.recipienttypedetails -ne 'MailContact'} :"
                    #==9:21 AM 10/8/2024:  since HR/WD change to SamAcctName as employe#, the above won't match any user created since 2022 or so. , 
                    # need to search on last name first

                    if($hSum.Rcp += get-recipient -id "$($substring)" -ea 0 -ResultSize $MaxRecips | select -first $MaxRecips |?{$_.recipienttypedetails -ne 'MailContact'}){
                        #$hSum.Rcp | write-output ;
                        # $propsRcpTbl
                        write-host -foregroundcolor yellow "`n$(($hSum.Rcp | ft -a $propsRcpTbl |out-string).trim())" ;
                    } ;
                    #write-host "$((get-alias ps1GetxRcp).definition) -id $($substring) -ea 0 |?{$_.recipienttypedetails -ne 'MailContact'} : "
                    write-host "get-xorecipient -id $($substring) -ea 0 |?{$_.recipienttypedetails -ne 'MailContact'} : "
                    if($hSum.xoRcp += get-xorecipient -id "$($substring)" -ea 0 -ResultSize $MaxRecips | select -first $MaxRecips |?{$_.recipienttypedetails -ne 'MailContact'}){
                        #$hSum.xoRcp | write-output ;
                        write-host -foregroundcolor yellow "`n$(($hSum.xoRcp | ft -a $propsRcpTbl |out-string).trim())" ;
                    } ;


                } ;
                # do ADUser search on fname/lname
                if($hSum.lname){
                    # try as surname & givenname
                    if($hSum.lname -match "[']"){
                        $fltr = "surname -eq " + $sQot + $($hSum.lname) + $sQot ; 
                    }else{
                        $fltr = "surname -eq '$($hSum.lname)'"
                    }
                    #$fltr = "givenname -eq '$($hSum.givenname)'" ;
                    if($hSum.fname){
                        if($hSum.fname -match "[']"){
                            $fltr += " -AND givenname -eq " + $sQot + $($hSum.fname) + $sQot ;
                        }else{
                            $fltr += " -AND givenname -eq '$($hSum.fname)'"
                        } ;
                    } ; 
                    if($tmpADo = get-aduser -filter $fltr -ea 0 -Properties *| select -first $MaxRecips){
                        $smsg = "Matched on:get-aduser -filter $($fltr) " ; 
                        write-verbose $smsg ; 
                    }elseif($hSum.lname){
                        # treat as a samaccountname                        
                        if($tmpADo = get-aduser -identity $hSum.lname -ea 0 -Properties *| select -first $MaxRecips){
                            $smsg = "Matched on:get-aduser -identity $($hSum.fname)" ; 
                            write-verbose $smsg ; 
                        } ; 
                    }
                    if($tmpADo){
                        # |?{$_.recipienttypedetails -ne 'MailContact'}){
                        $rno = 0 ; 
                        $tmpADo | foreach-object {
                            $thisADU = $_ ; 
                            $rno++
                            #$hSum.ADUser +=  $thisADU ; 
                            # formatted dump
                            $hsADU=@"

ADUser #$($rno):DN:$(($thisADU.distinguishedname|out-string).trim())
$(($thisADU | ft -a $prpADU[1..3]|out-string).trim())
$(($thisADU | ft -a  $prpADU[4..7]|out-string).trim()) 
$(($thisADU | ft -a  $prpADU[8..11]|out-string).trim())

"@ ;
                            write-host $hsADU ; 
                       } ; 
                    } ; 
                } 
                # do AADUser search on fname/lname
                if($hSum.lname){
                    # try as surname & givenname
                    # Get-AzureADGroup -filter "displayName eq 'ENT-SEC-SslVpn-AU-Administrators-DL'" ; 
                    # works: get-AzureAdUser -Filter "surname eq '$($hSum.surname)' and givenname eq '$($hSum.givenname)'"
                    if($hsum.lname -match "[']"){
                        $fltr = "surname eq " + $sQot + $($hsum.lname) + $sQot ;
                    }else{
                        $fltr = "surname eq '$($hsum.lname)'" ; 
                    }
                    #$fltr = "givenname -eq '$($hSum.givenname)'" ;
                    if($hSum.fname){
                        if($hsum.lname -match "[']"){
                            $fltr += " and givenname eq " + $sQot + $($hsum.fname) + $sQot ;
                        }else{
                            $fltr += " and givenname eq '$($hsum.fname)'"
                        }
                    } ; 
                    if($tmpAADo = get-AzureAdUser  -filter $fltr -ea 0 | select -first $MaxRecips){
                        $smsg = "Matched on:get-AzureAdUser -filter $($fltr) " ; 
                        write-verbose $smsg ; 
                    }elseif($hSum.lname){
                        # treat as a -ObjectId                        
                        if($tmpAADo = get-AzureAdUser -ObjectId $hSum.lname -ea 0 | select -first $MaxRecips){
                            $smsg = "Matched on:get-AzureAdUser -identity $($hSum.fname)" ; 
                            write-verbose $smsg ; 
                        } ; 
                    }
                    if($tmpAADo){
                        # |?{$_.recipienttypedetails -ne 'MailContact'}){
                        $rno = 0 ; 
                        $tmpAADo | foreach-object {
                            $thisADU = $_ ; 
                            $rno++
                            #$hSum.ADUser +=  $thisADU ; 
                            # formatted dump
                            $hsADU=@"

AADUser #$($rno):DN:$(($thisADU.distinguishedname|out-string).trim())
$(($thisADU | ft -a $prpADU[1..3]|out-string).trim())
$(($thisADU | ft -a  $prpADU[4..7]|out-string).trim()) 
$(($thisADU | ft -a  $prpADU[8..11]|out-string).trim())

"@ ;
                            write-host $hsADU ; 
                       } ; 
                    } ; 
                } 


                $abortReport = $true ; 



            } ; # don't break, doesn't continue loop

            if($abortReport ){
                $smsg = "(multiple recipients - or no recipients, but ADUsers, or but AADUsers -  found in OnPrem And/Or Cloud, detailed reporting & output aborted)" ; 
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                BREAK ; 
            } ; 

            # 10:42 AM 9/9/2021 force populate the xoMailbox, ALWAYS - need for xbrain ids
            #if($hSum.xoRcp.recipienttypedetails -eq 'UserMailbox' -AND -not($hSum.xoMailbox)){
            # accomodate array xorcp
            #if(($hSum.xoRcp|?{$_.recipienttypedetails -eq 'UserMailbox'}) -AND -not($hSum.xoMailbox)){
            # issue:quota on Shared: above only keys on recipienttypedetails -eq 'UserMailbox', should be *any* mailbox type, if we want quotas etc for shared/room/equipment  switch to rcptype: $hSum.xoRcp.RecipientType
            if(($hSum.xoRcp|?{$_.recipienttype -eq 'UserMailbox'}) -AND -not($hSum.xoMailbox)){
                #write-verbose "$((get-alias ps1GetxMbx).definition) w`n$(($pltGMailObj|out-string).trim())" ;
                write-verbose "get-xomailbox w`n$(($pltGMailObj|out-string).trim())" ;
                if($hSum.xoMailbox += get-xomailbox @pltGMailObj -ea 0| select -first $MaxRecips ){
                    $ino = 0 ;
                    $mapiResults = @() ;
                    foreach($xmbx in $hSum.xoMailbox){
                        $ino++ ;
                        if($hSum.xoMailbox -is [system.array]){
                            $msgprefix = "xmbx$($ino):" ;
                        } else { $msgprefix = $null } ;
                        $smsg = $msgprefix, "`$hSum.xoMailbox:`n$(($xmbx|ft -a |out-string).trim())" -join ' ' ;
                        write-verbose $smsg ;
                        $smsg = $msgprefix,"'xoUserMailbox':Test-exoMAPIConnectivity $($xmbx.userprincipalname)"  -join ' ' ;
                        write-verbose $smsg ;
                       $mapiResults += Test-xoMAPIConnectivity -identity $xmbx.userprincipalname ;
                        $smsg = "Outlook (xoMAPI) Access Test Result:$($mapiResults[$ino - 1].result)" ;
                        if($mapiResults[$ino - 1].result -eq 'Success'){
                            write-host -foregroundcolor green $smsg ;
                        } else {
                            write-WARNING $smsg ;
                        } ;
                    } ;
                    $hSum.xoMapiTest  +=  $mapiResults ;
                } ;
            } ;
            # 3:42 PM 9/25/2023 bring in new quota support as well - it's not populated in the oprcp first test
            if($getQuotaUsage){
                if(($hSum.xoRcp|?{$_.recipienttypedetails -match 'UserMailbox|SharedMailbox|RoomMailbox|EquipmentMailbox'}) -AND -not($hSum.xoMailboxStats)){
                    $pltGMbxStatX=[ordered]@{
                        identity = $hSum.xoMailbox.exchangeguid ;
                        ErrorAction = 'STOP' ; 
                    } ;
                    $smsg = "Get-xoMailboxStatistics  w`n$(($pltGMbxStatX|out-string).trim())"
                    if($verbose){
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                        else{ write-verbose $smsg } ; 
                    } ; 
                    $hSum.xoMailboxStats  +=  Get-xoMailboxStatistics @pltGMbxStatX | select $prpStat;
                    $smsg = "xoMailboxStats Count:$(($hsum.xoMapiTest|measure).count)" ;
                    write-host -foregroundcolor green $smsg ;

                    If($hSum.xoMailbox.UseDatabaseQuotaDefaults){
                        $hSum.xoEffectiveQuotas = $hSum.xoMailboxStats | select @{N ='IssueWarningQuotaMB'; e={$_.DBIssueWarningQuotaMB}},
                        @{n='ProhibitSendQuotaMB'; e={$_.DBProhibitSendQuotaMB}},
                        @{n='ProhibitSendReceiveQuotaMB';e={$_.DBProhibitSendReceiveQuotaMB}}; 
                    } else {
                        $hSum.xoEffectiveQuotas = $hSum.xoMailbox | select @{n="IssueWarningQuotaMB";e={[math]::round($_.IssueWarningQuota.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1MB,2)}},
                        @{n="ProhibitSendQuotaMB";e={[math]::round($_.ProhibitSendQuota.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1MB,2)}},
                        @{n="ProhibitSendReceiveQuotaMB";e={[math]::round($_.ProhibitSendReceiveQuota.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1MB,2)}} ;
                    } ;  
                    $hSum.xoNetOfSendReceiveQuotaMB  +=  $hSum.xoEffectiveQuotas.ProhibitSendQuotaMB - $hSum.xoMailboxStats.TotalMailboxSizeMB ; 

                    $pltGMbxStatX.add('IncludeOldestAndNewestItems',$true) ; 
                    $smsg = "Get-xoMailboxFolderStatistics  w`n$(($pltGMbxStatX|out-string).trim())" ;
                    if($verbose){
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                        else{ write-verbose $smsg } ; 
                    } ; 
                    $smsg = "(-getQuotaUsage:running lengthy Get-xoMailboxFolderStatistics...)" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                    else{ write-host -foregroundcolor gray "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                    TRY{
                        $hsum.xoMailboxFolderStats  +=  Get-xoMailboxFolderStatistics @pltGMbxStatX  ;

                        $smsg = "Export FolderStats to`n$(($ofMbxFolderStats|out-string).trim())" ;
                        if($verbose){
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                            else{ write-verbose $smsg } ; 
                        } ; 
                        $hsum.xoMailboxFolderStats | ?{$_.ItemsInFolder -gt 0 -AND $_.identity -notmatch $rgxHiddn } | 
                            select $prpFldr | sort SizeMB -desc | export-csv  -path $ofMbxFolderStats -notype ;

                    } CATCH {
                        $ErrTrapd=$Error[0] ;
                        $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                        write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" ;
                    } ; 
                    
                }
            } ; 

            #$pltgMsoUsr=@{UserPrincipalName=$null ; MaxResults= $MaxRecips; ErrorAction= 'STOP' } ;
            # maxresults is documented:
            # but causes a fault with no $error[0], doesn't seem to be functional param, post-filter
            # ren refs of $pltgMsoUsr -> $pltgAADUsr
            $pltgAADUsr=@{UserPrincipalName=$null ; ErrorAction= 'STOP' } ;
            if($hSum.ADUser){$pltgAADUsr.UserPrincipalName  +=  $hSum.ADUser.UserPrincipalName }
            elseif($hSum.xoMailbox){$pltgAADUsr.UserPrincipalName += $hsum.xoMailbox.UserPrincipalName }
            elseif($hSum.xoMUser){$pltgAADUsr.UserPrincipalName  +=  $hSum.xoMUser.UserPrincipalName }
            elseif($hSum.txGuest){$pltgAADUsr.UserPrincipalName  +=  $hSum.txGuest.userprincipalname }
            else{} ;

            if($pltgAADUsr.UserPrincipalName){

                if(-not($hSum.AADUser)){
                    write-host -foregroundcolor yellow "=get-AADuser $($pltgAADUsr.UserPrincipalName)>:" ;
                    TRY{
                        caad  -Verbose:$false -silent ;
                        write-verbose "get-AzureAdUser  -objectid $($pltgAADUsr.UserPrincipalName)" ;
                        # have to postfilter, if want specific count -maxresults catch's with no $error[0]
                        $hSum.AADUser   +=  get-AzureAdUser  -objectid $pltgAADUsr.UserPrincipalName  | select -first $MaxRecips;  ;
                        <# for remote federated, AADU brings in summary of remote ADUser:
                            $hsum.aaduser.ExtensionProperty
                            Key                                                       Value
                            ---                                                       -----
                            odata.metadata                                            https://graph.windows.net/.../$metadata#directoryObjects/@Element
                            odata.type                                                Microsoft.DirectoryServices.User
                            createdDateTime                                           1/13/2021 4:14:48 PM
                            employeeId
                            onPremisesDistinguishedName                               CN=XXX,OU=XXX,...
                            thumbnailPhoto@odata.mediaEditLink                        directoryObjects/.../Microsoft.DirectoryServices.User/thumbnailPhoto
                            thumbnailPhoto@odata.mediaContentType                     image/Jpeg
                            userIdentities                                            []
                            extension_9d88b2c96135413e88afff067058e860_employeeNumber 8621
                             $hsum.aaduser.ExtensionProperty.onPremisesDistinguishedName
                            CN=XXX,OU=XXX,...
                        #>
                        #write-verbose "`$hSum.AADUser:`n$(($hSum.AADUser|out-string).trim())" ;
                        # ObjectId                             DisplayName   UserPrincipalName      UserType

                        #lic pull: $hSum.AADUser | Get-AzureADUserLicenseDetail -ea STOP | select -exp SkuPartNumber
                        write-verbose "`$hsum.AADUserLics = AADU | Get-AzureADUserLicenseDetail -ea STOP | select -exp SkuPartNumber" ;
                        $hsum.AADUserLics  +=   $hSum.AADUser | Get-AzureADUserLicenseDetail -ea STOP | select -exp SkuPartNumber ; 

                    } CATCH {
                        $ErrTrapd=$Error[0] ;
                        $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                        else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        Continue ;
                    } ;

                } ;

                if(-not($hSum.AADUserMgr) -AND $hSum.AADUser ){
                    write-host -foregroundcolor yellow "=get-AADuserManager $($hSum.AADUser.UserPrincipalName)>:" ;
                    TRY{
                        caad  -Verbose:$false -silent ;
                        write-verbose "get-AzureAdUserManager  -objectid $($hSum.AADUser.UserPrincipalName)" ;
                        # have to postfilter, if want specific count -maxresults catch's with no $error[0]
                        $hSum.AADUserMgr   +=  get-AzureAdUserManager  -objectid $hSum.AADUser.UserPrincipalName  | select -first $MaxRecips;  ;
                        #write-verbose "`$hSum.AADUserMgr:`n$(($hSum.AADUserMgr|out-string).trim())" ;
                        # (returns a full AADUser obj for the mgr)
                        # we can output the DN: $hSum.AADUserMgr.ExtensionProperty.onPremisesDistinguishedName
                        # useful for determining what 'org' user should be for email address assigns - they get same addr dom as their mgr
                        # |ft -a  $propsaadmgr
                    } CATCH {
                        $ErrTrapd=$Error[0] ;
                        $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                        else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        Continue ;
                    } ;

                } ;

                # display user info:
                if(-not($hSum.ADUser)){
                    if($hSum.AADUser.DirSyncEnabled -AND $hSum.aaduser.ExtensionProperty.onPremisesDistinguishedName){
                        $pltGadu.Identity = $hSum.aaduser.ExtensionProperty.onPremisesDistinguishedName ; 
                        $hSum.ADUser  += Get-ADUser @pltGadu | select -first $MaxRecips ;
                        if($pltGadu.identity){
                            write-verbose "Get-ADUser w`n$(($pltGadu|out-string).trim())" ;
                            # try a nested local trycatch, against a missing result
                            Try {
                                #Get-ADUser $DN -ErrorAction Stop ;
                                $hSum.ADUser  += Get-ADUser @pltGadu | select -first $MaxRecips ;
                            } Catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
                                write-warning "(no matching ADuser found:$($pltGadu.identity))" ;
                            } catch {
                                $ErrTrapd=$Error[0] ;
                                $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                                else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                Continue ;
                            } ;

                            write-verbose "`$hSum.ADUser:`n$(($hSum.ADUser|fl $propsADU | out-string).trim())" ;
                            $smsg = "(TOR USER, fed:$($TORMeta.adforestname))" ;
                            $hSum.Federator  +=  $TORMeta.adforestname ;
                            write-host -Fore yellow $smsg ;
                            if($hSum.OPRemoteMailbox){
                                $smsg = "$(($hSum.OPRemoteMailbox |fl $propsMailx|out-string).trim())"
                                #$smsg += "`n-Title:$($hSum.ADUser.Title)"
                                $smsg += "`n$(($hSum.ADUser |fl 'Enabled','Description','whenCreated','whenChanged','Title' |out-string).trim())"
                            } ;
                            if($hSum.OPMailbox){
                                $smsg =  "$(($hSum.OPMailbox |fl $propsMailx|out-string).trim())" ;
                                $smsg += "`n$(($hSum.ADUser |fl 'Enabled','Description','whenCreated','whenChanged','Title' |out-string).trim())"
                            } ;
                            write-host $smsg ;
                        } ;
                    } else { 
                        # remote fed, use AADU to proxy remote AD hybrid info:
                        write-host -foreground yellow "===`$hSum.AADUser: " #-nonewline;
                        $smsg = "$(($hSum.AADUser| select $propsAADL1 |out-markdowntable @MDtbl |out-string).trim())" ;
                        $smsg += "`n$(($hSum.AADUser|select $propsAADL2 |out-markdowntable @MDtbl|out-string).trim())" ;
                        $smsg += "`n$(($hSum.AADUser|select $propsAADL3 |out-markdowntable @MDtbl|out-string).trim())" ;
                        $smsg += "`n$(($hSum.AADUser|select $propsAADL4 |out-markdowntable @MDtbl|out-string).trim())" ;
                        $smsg += "`n$(($hSum.AADUser|select $propsAADL5 |out-markdowntable @MDtbl|out-string).trim())" ;
                        #$hsum.aaduser.ExtensionProperty.onPremisesDistinguishedName
                        if($hSum.Federator -ne $TORMeta.adforestname){
                            $smsg += "`n$($hSum.Federator):Remote ADUser.DN:`n$(($hsum.aaduser.ExtensionProperty.onPremisesDistinguishedName|out-string).trim())" ;
                        }  ;
                    }; 

                    write-host $smsg

                    # assert the real names from the user obj
                    $hSum.dname  +=  $hSum.AADUser.DisplayName ;
                    $hSum.fname  +=  $hSum.AADUser.GivenName ;
                    $hSum.lname  +=  $hSum.AADUser.Surname ;

                } else {
                    #write-verbose "`$hSum.AADUser:`n$(($hSum.AADUser| ft -auto ObjectId,DisplayName,UserPrincipalName,UserType |out-string).trim())" ;
                    # defer to ADUser details
                    #"$(($hSum.ADUser |fl $propsMailx |out-markdowntable @MDtbl|out-string).trim())"
                    <#$propsADL1 = 'UserPrincipalName','DisplayName','GivenName','Surname','Title' ;
                    $propsADL2 = 'Company','Department','PhysicalDeliveryOfficeName' ;
                    $propsADL3 = 'StreetAddress','City','State','PostalCode','TelephoneNumber','MobilePhone' ;
                    #>
                    write-host -foreground yellow "===`$hSum.ADUser: " #-nonewline;
                    $smsg = "$(($hSum.ADUser| select $propsADL1 |out-markdowntable @MDtbl |out-string).trim())" ;
                    $smsg += "`n$(($hSum.ADUser|select $propsADL2 |out-markdowntable @MDtbl|out-string).trim())" ;
                    $smsg += "`n$(($hSum.ADUser|select $propsADL3 |out-markdowntable @MDtbl|out-string).trim())" ;
                    $smsg += "`n$(($hSum.ADUser|select $propsADL4 |out-markdowntable @MDtbl|out-string).trim())" ;
                    <# $propsADL5 = 'whenCreated','whenChanged' ; 
                    $propsADL6 = @{Name='Desc';Expression={$_.Description }} ;
                    #>
                    $smsg += "`n$(($hSum.ADUser|select $propsADL5 |out-markdowntable @MDtbl|out-string).trim())" ;
                    # stick desc on trailing line $propsADL5
                    #$smsg += "`n$(($hSum.ADUser|select $propsADL5 |out-markdowntable @MDtbl|out-string).trim())" ;
                    # flip L5 to fl (suppress crlf wrap)
                    $smsg += "`n$(($hSum.ADUser|select $propsADL6 |Format-List|out-string).trim())" ;

                    # moved DN into adl4, w enabled
                    #$smsg += "`n`$ADUser.DN:`n$(($hsum.aduser.DistinguishedName|out-string).trim())" ;
                    #$smsg += "`n$($hSum.ADUser|select Enabled,distinguishedname| convertTo-MarkdownTable -NoDashRow -Border) `$ADUser.DN:`n$(($hsum.aduser.DistinguishedName|out-string).trim())" ;
                    write-host $smsg ;

                    # assert the real names from the user obj
                    $hSum.dname  +=  $hSum.ADUser.DisplayName ;
                    $hSum.fname  +=  $hSum.ADUser.GivenName ;
                    $hSum.lname  +=  $hSum.ADUser.Surname ;
                } ;

                # acct enabled/disabled: .aduser.Enbabled & .aaduser.AccountEnabled
                if($hSum.aduser){
                    if($hSum.aduser.Enabled){} else {
                        $smsg = "ADUser:$($hSum.ADUser.userprincipalname) AD Account is *DISABLED!*"
                        write-warning $smsg ;
                    } ;
                } ;
                # acct enabled/disabled: .aduser.Enbabled & .aaduser.AccountEnabled
                if($hSum.AADUser){
                    if($hSum.aaduser.AccountEnabled){} else {
                        $smsg = "AADUser:$($hSum.AADUser.userprincipalname) AAD Account is *DISABLED!*"
                        write-warning $smsg ;
                    } ;
                } ;
                if($hSum.ADUser){$hSum.LicenseGroup  +=  $hSum.ADUser.memberof |?{$_ -match $rgxOPLic }}

                if($hSum.ADUser){$hSum.LicenseGroup  +=  $hSum.ADUser.memberof |?{$_ -match $rgxOPLic }}
                elseif($hSum.xoMemberOf){$hSum.LicenseGroup  +=  $hSum.xoMemberOf.Name |?{$_ -match $rgxXLic}}
                #if(-not ($hSum.LicenseGroup) -AND ($hSum.MsolUser.licenses.AccountSkuId -contains "$($TORMeta.o365_TenantDom.tolower()):ENTERPRISEPACK")){$hSum.LicenseGroup  +=  '(direct-assigned E3)'} ;
                # $hSum.AADUser ; $aadu | Get-AzureADUserLicenseDetail  | select -exp SkuPartNumber
                #if(-not ($hSum.LicenseGroup) -AND ( $hsum.AADUserLics  -contains "$($TORMeta.o365_TenantDom.tolower()):ENTERPRISEPACK")){$hSum.LicenseGroup  +=  '(direct-assigned E3)'} ;
                # no dom, with aadu licenses
                if(-not ($hSum.LicenseGroup) -AND ( $hsum.AADUserLics  -contains "ENTERPRISEPACK")){$hSum.LicenseGroup  +=  '(direct-assigned E3)'} ;
                if($hSum.LicenseGroup){$smsg = "LicenseGroup:$($hSum.LicenseGroup)"}
                else{$smsg = "LicenseGroup:(unresolved, direct-assigned other?)" } ;
                write-host $smsg ;

                if($hSum.AADUserMgr){
                    #($hSum.AADUserMgr) |ft -a  $propsaadmgr
                    #$smsg += "`nAADUserMgr:`n$(($hSum.AADUserMgr|select $propsAadMgr |out-markdowntable @MDtbl|out-string).trim())" ;
                    # $propsAADMgrL1, $propsAADMgrL2
                    write-host -foreground yellow "===`$hSum.AADUserMgr: " #-nonewline;
                    $smsg = "$(($hSum.AADUserMgr| select $propsAADMgrL1 |out-markdowntable @MDtbl |out-string).trim())" ;
                    #$smsg += "`n$(($hSum.AADUserMgr|select $propsAADMgrL2 |out-markdowntable @MDtbl|out-string).trim())" ;
                    $smsg += "`n$(($hSum.AADUserMgr|Format-List $propsAADMgrL2|out-string).trim())" ;
                    #$smsg += "`n$(($hSum.AADUserMgr|select $propsADL3 |out-markdowntable @MDtbl|out-string).trim())" ;
                } else {
                    $smsg += "(AADUserMgr was blank, or unresolved)" ;
                } ;
                write-host $smsg ;

                if($getQuotaUsage -AND $hSum.xoMailbox){

                    $smsg += "`n`nLicenses::`n$(($hsum.AADUserLics -join ', ' |out-string).trim())`n`n" ; 
                    $smsg += "`nwhich specify the following size limits:`n$(($hSum.xoEffectiveQuotas| fl |out-string).trim())`n(UseDatabaseQuotaDefaults:$($hSum.xoMailbox.UseDatabaseQuotaDefaults))" ; 
                    $smsg += "`n`nCurrent TotalMailboxSizeMB: $($hSum.xoMailboxStats.TotalMailboxSizeMB)`n`n" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level PROMPT } 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success

                    if($hSum.xoNetOfSendReceiveQuotaMB -lt 0){
                        $smsg = "`n`n*** QuotaStatus: Mailbox is *OVER* mandated SendReceiveQuotaMB by $(($hSum.xoNetOfSendReceiveQuotaMB * -1).tostring("N")) megabytes ***`n`n" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 

                    } else { 
                        $smsg = "QuotaStatus: Mailbox is below mandated SendReceiveQuotaMB by $(($hSum.xoNetOfSendReceiveQuotaMB).tostring("N")) megabytes" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Prompt } 
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                    } ;

                    $smsg = "`nWith the following non-zero folder metrics`n`n$((import-csv $ofMbxFolderStats | ft -auto |out-string).trim())" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level PROMPT } 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success

                    $smsg = "`n===`output to::`n$($ofMbxFolderStats)`n" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    
                    # 10:08 AM 2/27/2024: Add: .xoMailboxFolderStats DiscoveryHolds folder only has ItemsInFolder -gt 0 if there's a hold
                    if($hSum.xoMailbox.LitigationHoldEnabled -OR $hSum.xoMailbox.InPlaceHolds -OR $hSum.xoMailbox.ComplianceTagHoldApplied -OR $hSum.xoMailbox.DelayHoldApplied -OR $hSum.xoMailbox.DelayReleaseHoldApplied -OR ($hSum.xoMailboxFolderStats | ?{$_.name -match 'DiscoveryHolds' -AND $_.ItemsInFolder -gt 0})  ){
                        $smsg = "`n`nEVIDENCE OF LEGAL HOLD DETECTED!:`n$(($hSum.xoMailbox | fl $prpMbxHold|out-string).trim())`n`n" ; 
                        if($hSum.xoMailboxFolderStats | ?{$_.name -match 'DiscoveryHolds' -AND $_.ItemsInFolder -gt 0}){
                            $smsg += "`n$(($hSum.xoMailboxFolderStats | ?{$_.name -match '^DiscoveryHolds$'} | ft -a $prpFldrLH|out-string).trim())`n`n" ; 
                            $smsg += "`n- DiscoveryHolds folder: If In-Place Hold is enabled or if a Microsoft 365 or Office 365 retention policy is assigned to the mailbox, this subfolder contains all items that meet the hold query parameters and are hard deleted." ; 
                            $smsg += "`n- DiscoveryHolds folder.NewestItem: Will reflect *last time LegalHold captured an item* (e.g. if/when LH was disabled and stopped holding traffic, if in the past)`n"; 
                        } 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 

                    } else {
                        $smsg = "`n`n*No* evidence Of Legal Hold detected:`n$(($hSum.xoMailbox | fl $prpMbxHold|out-string).trim())`n`n" ; 
                        $smsg = "`n$(($hSum.xoMailboxFolderStats | ?{$_.name -match 'DiscoveryHolds'} | ft -a $prpFldrLH|out-string).trim())`n`n" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level PROMPT } 
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success

                    } ;  

                    $hsInfo = @"

## System Folder Types and purposes:

- Recoverable Items: Items in the Recoverable Items folder aren't calculated toward the user's mailbox quota. In Exchange Online, the Recoverable Items folder has its own quota. In Exchange Online, the quota for the Recoverable Items folder (in the user's primary mailbox) is automatically increased to 100 GB when a mailbox is placed on Litigation Hold or In-Place Hold. 

    ### Subfolders of Recoverable Items:
    
    *  Deletions: This subfolder contains all items deleted from the Deleted Items folder. (In Outlook, a user can soft delete an item by pressing Shift+Delete.) This subfolder is available to users through the Recover Deleted Items feature in Outlook and Outlook on the web.
    
    *  Versions: If In-Place Hold, Litigation Hold, or a Microsoft 365 or Office 365 retention policy is enabled, this subfolder contains the original copy of the item and also if the item is modified multiple times, a copy of the item before modification is saved.
    
    *  Purges: If either Litigation Hold or single item recovery is enabled, this subfolder contains all items that are hard deleted. 
    
    *  Audits: If mailbox audit logging is enabled for a mailbox, this subfolder contains the audit log entries. 
    
    *  DiscoveryHolds: If In-Place Hold is enabled or if a Microsoft 365 or Office 365 retention policy is assigned to the mailbox, this subfolder contains all items that meet the hold query parameters and are hard deleted.

## Deleted item retention
  An item is considered to be soft deleted in the following cases:
    • A user deletes an item or empties all items from the Deleted Items folder.
    • A user presses Shift+Delete to delete an item from any other mailbox folder.
    
  Soft-deleted items are moved to the Deletions subfolder of the Recoverable Items folder. This provides an additional layer of protection so users can recover deleted items without requiring Help desk intervention. Users can use the Recover Deleted Items feature in Outlook or Outlook on the web to recover a deleted item. Users can also use this feature to permanently delete an item. 
  
  Items remain in the Deletions subfolder until the deleted item retention period is reached. The deleted item retention period for Exchange Online is 30 days (Toroco). In addition to a deleted item retention period, the Recoverable Items folder is also subject to quotas. 
  
  When the deleted item retention period expires, the item is completely removed from Exchange Online.

"@ ; 
                    write-host $hsInfo ;   

                } ; 
            } ;

            # do a split-brain/nobrain check
            # switch ($hSum.OPRcp.recipienttypedetails){
            <#
            AD - Users (more effective)
            (sAMAccountType=805306368)
            AD - Users - disabled
            (&(sAMAccountType=805306368)(userAccountControl:1.2.840.113556.1.4.803:=2))
            AD - Users - dont require password
            (&(sAMAccountType=805306368)(userAccountControl:1.2.840.113556.1.4.803:=32))
            AD - Users - mail enabled
            (&(sAMAccountType=805306368)(mailNickname=*))
            AD - Users - password never expires
            (&(sAMAccountType=805306368)(userAccountControl:1.2.840.113556.1.4.803:=65536))

            Select ($hSum.ADUser.sAMAccountType){
                '0'  { $SAType = "SAM_DOMAIN_OBJECT"}
                '268435456' { $SAType = "SAM_GROUP_OBJECT"}
                '268435457' { $SAType = "SAM_NON_SECURITY_GROUP_OBJECT"}
                '536870912' { $SAType = "SAM_ALIAS_OBJECT"}
                '536870913' { $SAType = "SAM_NON_SECURITY_ALIAS_OBJECT"}
                '805306368' { $SAType = "SAM_NORMAL_USER_ACCOUNT"}
                '805306369' { $SAType = "SAM_MACHINE_ACCOUNT"}
                '805306370' { $SAType = "SAM_TRUST_ACCOUNT"}
                '1073741824' { $SAType = "SAM_APP_BASIC_GROUP"}
                '1073741825' { $SAType = "SAM_APP_QUERY_GROUP"}
                '2147483647' { $SAType = "SAM_ACCOUNT_TYPE_MAX"}
                default { $SAType = "UNKNOWN" }
            } ;
            #>
            # ($hSum.ADUser.sAMAccountType -eq '805306368')

            if($hsum.ADUser){
                $hsum.IsADDisabled  +=  [boolean]($hsum.ADUser.Enabled -eq $true) ; 
             } else {
                write-verbose "(no ADUser found)" ;
            } ;
            if($hsum.AADUser){
                $hsum.IsAADDisabled  +=  [boolean]($hsum.AADUser.AccountEnabled -eq $true) ; 
                $hsum.isDirSynced  +=  [boolean]($hsum.AADUser.DirSyncEnabled  -eq $True)
            } else {
                write-verbose "(no AADUser found)" ;
            } ;
            # shift test to aadu
            if($hSum.AADUser){
                $hsum.IsLicensed  +=  [boolean]($hSum.AADUser.assignedlicenses.count -gt 0)
            } else {
                write-verbose "(no AADUser found)" ;
            } ;

            $smsg = "`n"
            if(($hsum.xoRcp.RecipientTypeDetails -match '(UserMailbox|MailUser)') -AND $hsum.IsLicensed -AND $hSum.xomailbox -AND $hSum.OPMailbox){
                <#OPRcp, xorcp, OPMailbox, OPRemoteMailbox, xoMailbox#>
                $smsg += "SPLITBRAIN!:$($hSum.ADUser.userprincipalname).IsLic'd & has *BOTH* xoMbx & opMbx!" ;
                $hsum.IsSplitBrain  +=  $true ;
            }elseif(($hsum.xoRcp.RecipientTypeDetails -match '(UserMailbox|MailUser)') -AND -not($hsum.IsLicensed) -AND $hSum.xomailbox -AND $hSum.OPMailbox){
                <#OPRcp, xorcp, OPMailbox, OPRemoteMailbox, xoMailbox#>
                $smsg += "SPLITBRAIN!:$($hSum.ADUser.userprincipalname).IsLic'd & has *BOTH* xoMbx & opMbx!`nAND is *UNLICENSED!*" ;
                $hsum.IsSplitBrain  +=  $true ;
            } elseif(($hsum.xoRcp.RecipientTypeDetails -match '(UserMailbox|MailUser)') -AND $hsum.IsLicensed -AND -not($hSum.xomailbox) -AND -not($hSum.OPMailbox)){
                $smsg += "NOBRAIN! W LICENSE!:$($hSum.ADUser.userprincipalname).IsLic'd &  has *NEITHER* xoMbx OR opMbx!" ;
                $hsum.IsNoBrain  +=  $true ;
            } elseif (($hsum.xoRcp.RecipientTypeDetails -match '(UserMailbox|MailUser)') -AND -not($hsum.IsLicensed) -AND -not($hSum.xomailbox) -AND -not($hSum.OPMailbox)){
                $smsg += "NOBRAIN! *WO* LICENSE! (TERM?):$($hSum.ADUser.userprincipalname) NOT licensed'd &  has *NEITHER* xoMbx OR opMbx!" ;
                $hsum.IsNoBrain  +=  $true ;
            } elseif($hsum.IsLicensed -eq $false){
                # 12:37 PM 12/26/2024 ACCOMOD UNlic'd non-user mbxs (normal)
                if($hsum.xoRcp.RecipientTypeDetails -match 'SharedMailbox|RoomMailbox|EquipmentMailbox'){
                    $smsg += "$($hSum.ADUser.userprincipalname) Is RecipientTypeDetails:$($hsum.xoRcp.RecipientTypeDetails) _expected unlicensed_" ;
                } ELSE { 
                    $smsg += "$($hSum.ADUser.userprincipalname) Is *UNLICENSED*!" ;
                } ; 
                $hsum.IsLicensed  +=  $false ;
            } elseif($hsum | ?{-not $_.ADUser -AND $_.AADUser -AND $_.xomailbox -AND -not $_.opMailbox -AND -not $_.opRemoteMailbox}){
                # 3:54 PM 10/16/2024 add cloud-first VEN|INT|AA|HH detect
                $smsg += "LICENSED AADUSER CLOUD-FIRST XOMAILBOX  (No ADUser, No OPMailbox, No OPRemoteMailbox)~" ; 
            } ELSE { } ;

            if($hsum.IsSplitBrain -OR $hsum.IsNoBrain -OR (-not $hsum.IsLicensed -AND $hsum.xoRcp.RecipientTypeDetails -NOTmatch 'SharedMailbox|RoomMailbox|EquipmentMailbox') ){
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
            } else { 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
            } ; 

            if($hsum.IsNoBrain){
                switch ($hSum.Federator) {
                    $TORMeta.adforestname {$rgxTermOU = $TORMeta.rgxTermUserOUs }
                    $CMWMeta.adforestname  {$rgxTermOU = $CMWMeta.rgxTermUserOUs }
                    $VENMETA.o365_TenantLabel  {$rgxTermOU = $NULL }
                    $TORMeta.o365_TenantDom   {$rgxTermOU = $NULL }
                    default {
                        write-warning "UNRECOGNIZED `$hsum.FEDERATOR!:$($hSum.Federator)" ;
                    }
                }

                if($rgxTermOU -AND $hsum.ADUser){
                    if($hsum.ADUser.distinguishedname -match $rgxTermOU){
                        $hsum.IsDisabledOU  +=  $true ;
                        $smsg += "`n--ADUser:$($hsum.ADUser.samaccountname) is within a *DISABLED* OU (likely TERM)" ;
                    } else {
                        $hsum.IsDisabledOU  +=  $false ;
                        $smsg += "`n--ADUser:$($hsum.ADUser.samaccountname) is *NOT* in a DISABLED OU (improperly offboarded TERM?)" ;
                    } ;
                } else {
                    $smsg +=  "`n--Cloud-only or other non-AD-resolvable host" ;
                }
                if($hsum.ADUser){
                    $smsg += "`n----$($hsum.ADUser.distinguishedname)" ;
                    $smsg += "`n--ADUser.Description:$($hsum.ADUser.Description)" ;
                    if($hsum.IsADDisabled){
                        $smsg += "`n--ADUser:$($hsum.ADUser.samaccountname) is *DISABLED* for logon (likely TERM)" ;
                    } else {
                        $smsg += "`n--ADUser:$($hsum.ADUser.samaccountname) is *UN-DISABLED* for logon (improperly offboarded TERM?)" ;
                    } ;
                } else {
                    write-verbose "(no ADUser found)" ;
                } ;
                if($hsum.IsAADDisabled){
                    $smsg += "`n--AADUser:$($hsum.AADUser.UserPrincipalName) is *DISABLED* for logon (likely TERM)" ;
                } else {
                    $smsg += "`n--AADUser:$($hsum.AADUser.UserPrincipalName) is *UN-DISABLED* for logon (improperly offboarded TERM?)" ;
                } ;
                $smsg += "`n"
                write-warning $smsg ;
            } ;



            if($outObject){
                if($PSCmdlet.MyInvocation.ExpectingInput){
                    write-verbose "(pipeline input, skipping aggregator, dropping into pipeline)" ;
                    New-Object PSObject -Property $hSum | write-output  ;
                } else {
                    $Rpt += New-Object PSObject -Property $hSum ;
                } ;
            } ELSE {
                # 3:59 PM 9/18/2023 else export to report file 
                $Rpt += New-Object PSObject -Property $hSum ;
                $Rpt | export-clixml -Path $ofile -Depth 100 ;
            } ;
            write-host -foregroundcolor green $sBnr.replace('=v','=^').replace('v=','^=') ;
        } ;
        #endregion PIPELINE_PROCESSINGLOOP ; #*------^ END PIPELINE_PROCESSINGLOOP ^------

    } # PROC-E
    END{
        <## cleanup XO aliases
        get-alias -scope Script |?{$_.name -match '^ps1.*'} | %{Remove-Alias -alias $_.name} ; 
        #>
        if($abortReport){
            $smsg = "(multiple recipients found in OnPrem And/Or Cloud, detailed reporting & output aborted)" ; 
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        }else{
            if($outObject -AND -not ($PSCmdlet.MyInvocation.ExpectingInput)){
                $Rpt | write-output ;
                write-host "(-outObject: Output summary object to pipeline)"
            }elseif($outObject -AND ($PSCmdlet.MyInvocation.ExpectingInput)){
                write-verbose "(pipeline input, individual objects dropped into pipeline)" ;
            } else {
                $oput = ($Rpt | select-object -unique) -join ',' ;
                $oput | out-clipboard ;
                write-host "(output copied to clipboard)"
                $oput |  write-output ;
            } ;
        } ;  # if-E

     } ;
 }

#*------^ resolve-user.ps1 ^------