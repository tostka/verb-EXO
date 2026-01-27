# resolve-user.ps1

#region RESOLVE_USER ; #*------v resolve-user v------
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
    * 10:48 AM 1/19/2026 bugfix: $pltCcOPSvcs.UserRole (postfilter, not match test)
    * 4:31 PM 1/6/2026 fixed brace mismatch; updated in latest CONNECT_O365SERVICES, CALL_CONNECT_O365SERVICES, CALL_CONNECT_OPSERVICES, START_LOG_OPTIONS
    * 10:35 AM 12/22/2025 pasted in latest start_log block for holistic from psparamt; added missing per-user lop stop transcript, right below write_output block
    * 3:04 PM 12/12/2025 was still diverting logging when .psm1 module into modules dir; also outputing massive export-clixml 3gb files, needed testing of the content prior to exports. 
    - found that, since adding Get-MgUser Microsoft.graph to the mix, export-clixml can no longer export importable xml code. Even if you message it and pretest with convertto-xml, and try/catch for export errors (none). 
    still junk output xml. 
    - so coded in json export support; added -useJson & -Depth parameter (defaults to 2), and along the way found that Value/value key use in exports causes json to throw up, so coded up vio\convert-fromJsonSmart() which works past the issue
    (traps specific error, and mmd-conversion replaces instances of the first duplicated key, with the second key name). Necessary to get back to importable exported files from this process. 
    * 4:40 PM 12/11/2025 finished conversion to - marginally functional MG\get-MGUser, get-MGUserLicenseDetails, get-MGUserManager support. 
    rebuilt the property filters to accomodate the mg variant property names; 
    added get-mguserFull(), which force pulls a complete set of usable properties, returns the mguser object to pipeline
    rejiggered properties to populate fields that are there, but empty by default, at least the useful ones. 
    There are some audit log & service tied props, that throw Access Denied.
    Also had to add Group read all scope to get memberof (and still went to the long-functional Exchange-based memberof resolution, the MG is another 💩 show).
    $MGPermissionsScope = @("User.Read.All","Group.Read.All")
    renamed regions AAD -> MG refs 
    Shift of cloud user to MG rEsulted in a new bug - convert-clixml suddenly began never infishing, and genreated a 8gb .xml file
    Attempts to test scale revealed graph.api produces broken unclosed xml elements: 
        `Get-MgUser` piped to `Export-Clixml` is often caused by an issue in how 
        PowerShell handles complex or deeply nested objects during the XML 
        serialization process. This is a known issue within PowerShell itself, 
        especially when dealing with certain object structures returned by the 
        Microsoft Graph cmdlets.  
    Recommendations were to simplfiy the objects:
    Everything but the new MGUser, MGUserManager were long standing normal objs
    - I reselected MGUser with the same $prpMGUser set that spec'd the -propertiees for the Get-MgUser cmd
    - I also converted the MGUserManager hashtable (what comes out of the qry) into a psCustomObject, and remounded it to the exportable object. 
    That ccombo worked past the issues, and resulted in 800k-1.7mb xml files for single accounts. 
    Ridiculous that MS forces this type of half baked garbage on us, when we've had long-standing high-function solutions *for years*


    * 4:35 PM 12/10/2025 extensive fixes: damage from MS fully blocking al AzureAdUser access, and forced march to marginally compat get-MGUser command. Had to do a lot of revising, including in connecto o365 functions. Also fixed breaks in start-log call (was logging into module dir). 
    * 9:39 AM 10/10/2025 add: if -getQuotaUsage, and sharedmailbox recipienttypedetails, output info about Deleted Items & Sent Items OL mgmt regkeys.
    * 1:45 PM 9/23/2025 removed err-source connect-exo2 call (retured) ; added expanded mobile device reporting, testing Microsoft Nativ esync protocol (Outlook|REST clienttype) tests, explciit 'EAS' stigma tagging in outputs (wastes time t-shooting unsupported 3rd-party clients; given Security formally prefers OLM client over otheres).
    * 2:24 PM 8/1/2025 pulled unused whpassfail defs
    * 10:55 AM 4/15/2025 updated added param -ResolveForwards:
        -  to expand MailContacts into object that forwards the contact (net of MsgTracese that show the contact as a leaf recipient, informs *who* forwarded to the contact) ; 
        - new func: resolve-RMbxForwards() pulls all Rmbxs w ForwardingAddress populated, 
            grcps the Forward & builds an indexed hash to look up the primarysmtpAddress of the forwarding target, against  the detail of the forwarding mailbox 
            (for -ResolveForwards lookup speed, run a series of MailContact addresses through, and it only has to build the hash once, recycling the hash for the full set)
        - Also adds extra returned properties: opMailContact,opContactForwards,xoMailContact,xoMailboxForwardingAddress,xoContactForwards
        - made normal MailContact rcp exclusion conditional: exempts when running -ResolveForwards
        - Also expanded rmbx/opmbx/xombx to expand and fully report ForwardingAddress targeted rcp object
        - spliced in new resolve-Enviornment() & Start-Log code to match; working: works
        
    * 3:21 PM 4/12/2025 usable for now ; WIP implemented initial attempt at getting Forwarded MailContacts coded, still throwing odd errors, tho it completes, run against 3  contact addresess.
    * 12:40 PM 1/16/2025 UPDATED cbh WITH DETAILED PARAM DESC & OUTPUT SAMPLES ; 
         fixed missing -getMobile support in the force trailing pass; fixed mis applied $hSum.xoMapiStats for proper metrics
    * 4:41 PM 1/9/2025 rebuffered in latest Server Connections, found that the 
        assumption could use the existing PS session context for REMS, was bogus. So 
        re-enabled the OP cred gather even for useExopNoDep conns. Also reworked 
        connect-exchangeserverTdo() to actually use the credentials passed in, and 
        added the missing import-module $PSS to make the session actually functional 
        for running cmds, wo popping cred prompts. 
    * 8:53 AM 12/31/2024 cbh typo: cleared duped param Tenorg
    * 10:45 AM 12/27/2024 param aliass 'Quota','Perms' ; default -silent $true; updated propsADU to include desc & info ; add: $propsDG &  $propsADL7 ; rework into a loop for perm group summary dump; moved members & managedby into the grp summary; 
        removed nonewlines on the initial OP mbx/rmbx type; tweaked unlic & disabled ww's to only fire on inapprop config (smbx v umbx)
    * 3:43 PM 12/26/2024 add: -getPerms, runs Get-xoMailboxPermission & get-xoRecipientPermissions, outputs/returns non-SELF matches, and expands any group members in user or trustee
         add: aduser.info field, echo into output, if pop'd; 
        bugfix/cmw uses r: as room dname prefix, not recog'd as dname: #updated: $rgxDName CMW uses : in their room names, so went for broader AD dname support, per AI, and web specs, added 1-256char AD restriction         $rgxDName
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

    Typical summary block written to console (write-host, not pipeline):

        10:06:45:===v (1/1):Input: 'lynctest14@DOMAIN.COM' | 'lynctest14' | '' v===(EML)
        get-Rmbx/xMbx: (Rmbx *SHARED*)
        (xSMbx)(TOR USER, fed:SUB.DOMAIN.COM)
        SamAccountName | WindowsEmailAddress
        lynctest14     | lynctest14@DOMAIN.COM
        Office | RecipientTypeDetails | RemoteRecipientType     | IsDirSynced
               | RemoteSharedMailbox  | Migrated, SharedMailbox |
        ExternalDirectoryObjectId | CustomAttribute5 | EmailAddressPolicyEnabled
                                  |                  | True
        Outlook (xoMAPI) Access Test Result:Success
        =get-AADuser lynctest14@DOMAIN.COM>:
        =get-AADuserManager lynctest14@DOMAIN.COM>:
        ===$hSum.ADUser:
        UPN                 | DName       | FName | LName       | Title
        lynctest14@DOMAIN.COM | lync test14 |       | lync test14 |
        Company | Dept | Ofc
                |      |
        Street | City | State | Zip | Phone | Mobile
               |      |       |     |       |
        Enabled | DN
        False   | CN=lync test14,OU=users,OU=SITE,DC=sd,DC=sub,DC=domain,DC=com
        whenCreated           | whenChanged
        5/13/2015 11:32:01 AM | 12/19/2024 3:18:41 PM
        Desc :
        LicenseGroup:(unresolved, direct-assigned other?)
        (AADUserMgr was blank, or unresolved)
        10:06:49: INFO:
        lynctest14@DOMAIN.COM Is RecipientTypeDetails:SharedMailbox _expected unlicensed_
        ===^ (1/1):Input: 'lynctest14@DOMAIN.COM' | 'lynctest14' | '' ^===(EML)


    Key parameter options: 

    -getMobile parameter, to return details on xo MobileDevices in use with the EXO mailbox
        Note: 
            - adds inline output:
                xoMobileDeviceStats Count:2
                Evaluates and reports on Outlook Mobile use, OLM ClientType sync, 
                Stigmatizes and NOTE:'s EAS use with Best Effort support status. 
            - adds outobject property:
            $results:
                xoMobileDeviceStats    (LastSyncTime -LE 30D)
                xoMobileDeviceStatsOLD (LastSyncTime -GT 30D)
                xoMobileOutlookClients (OL Mobile clients)
                xoMobileOtherClients   (Non-OL Mobile clients)
                xoMobileOMSyncTypes    ('Outlook' (MS Native Sync) v 'REST' (legacy GAPI))
                xoMobileOtherSyncTypes ('EAS' et al)
        
        Typical Summary Block - Iphone
        ===$hsum.xoMobileDeviceStats:
        =get-xMob1:(ACTIVE)>
        FriendlyName   | DevType | DevOs             | ClntType | DevID
        iPhone 16 Plus | iPhone  | iOS 18.6.2 22G100 | EAS      | VC6DOHnnnnnnnnnnnnVOUL7KLK
        1stSyncTime  | LastSyncTime  | LastSuccSync  | #Folders
        4/3/25 19:23 | 9/23/25 14:22 | 9/23/25 14:22 | 98
        =get-xMob2:(inactive)>
        FriendlyName  | DevType | DevOs             | ClntType | DevID
        iPhone 7 Plus | iPhone  | iOS 15.8.3 19H386 | EAS      | 15UG7D5nnnnnnnnnnnR14T7EM8
        1stSyncTime  | LastSyncTime | LastSuccSync | #Folders
        8/29/24 8:24 | 4/4/25 6:46  | 4/4/25 6:46  | 83
        ---NON-Outlook Mobile Clients:(device-vendor-supported): 2
        DeviceFriendlyName ClientType LastSyncTime  LastSuccSync
        ------------------ ---------- ------------  ------------
        iPhone 16 Plus     EAS        9/23/25 14:22 9/23/25 14:22
        iPhone 7 Plus      EAS        4/4/25 6:46   4/4/25 6:46

        The following devices use device-vendor-provided/supported 'ExchangeActiveSync/EAS' Mobile clients!
        PLEASE NOTE: By policy EAS clients are *Best Effort* supported:
        Where issues are experienced with legacy EAS/ActiveSync clients,
        Users should be urged to move to _Supported_ Microsoft Outlook Mobile for IOS or Android
        DeviceFriendlyName ClientType LastSyncTime  LastSuccSync
        ------------------ ---------- ------------  ------------
        iPhone 16 Plus     EAS        9/23/25 14:22 9/23/25 14:22
        iPhone 7 Plus      EAS        4/4/25 6:46   4/4/25 6:46   

        Typical Summary Block - Outlook Mobile Android 
        ===$hsum.xoMobileDeviceStats:
        =get-xMob1:(ACTIVE)>
        FriendlyName | DevType | DevOs | ClntType | DevID
                     | Outlook | 15    | Outlook  | D115DF6C8E0nnnnnnnnnnnnn0682152D
        1stSyncTime  | LastSyncTime  | LastSuccSync  | #Folders
        2/24/25 9:03 | 9/23/25 13:52 | 9/23/25 13:52 | 0
        =get-xMob2:(inactive)>
        FriendlyName | DevType | DevOs | ClntType | DevID
                     | Outlook | 14    | Outlook  | 5D9DF50F879nnnnnnnnnnnnnC0B6988D
        1stSyncTime   | LastSyncTime  | LastSuccSync  | #Folders
        8/17/23 17:50 | 2/24/25 22:29 | 2/24/25 22:29 | 0
        =get-xMob3:(inactive)>
        FriendlyName | DevType     | DevOs              | ClntType | DevID
        aaa-8aaa1a2  | WindowsMail | Windows 10.0.17134 | EAS      | BEB93DA5nnnnnnnnnnnn974B6036A907
        1stSyncTime   | LastSyncTime | LastSuccSync | #Folders
        1/24/22 14:58 |              |              | 0
        +++Supported Outlook Mobile Clients: 2

        -----$hsum.xoMobileOMSyncTypes: Outlook
        ++User has has one or more fully compliant 'MS Native Sync'-protocol Outlook Mobile clients
        ---NON-Outlook Mobile Clients:(device-vendor-supported): 1
        DeviceFriendlyName ClientType LastSyncTime LastSuccSync
        ------------------ ---------- ------------ ------------
        aaa-8aaa1a2        EAS

        The following devices use device-vendor-provided/supported 'ExchangeActiveSync/EAS' Mobile clients!
        PLEASE NOTE: By policy EAS clients are *Best Effort* supported:
        Where issues are experienced with legacy Eas/ActiveSync clients,
        Users should be urged to move to _Supported_ Microsoft Outlook Mobile for IOS or Android
        DeviceFriendlyName ClientType LastSyncTime LastSuccSync
        ------------------ ---------- ------------ ------------
        aaa-8aaa1a2        EAS
        14:18:54: INFO:

        PS> $results.xoMobileDeviceStats | ft -a

        FirstSyncTime         LastPolicyUpdateTime  LastSyncAttemptTime  LastSuccessSync      DeviceType  DeviceID                         DeviceUserAgent       DeviceWipeSentTime DeviceWipeRequestTime DeviceWipeAckTime
        -------------         --------------------  -------------------  ---------------      ----------  --------                         ---------------       ------------------ --------------------- -----------------
        1/24/2022 8:58:54 PM                                                                  WindowsMail XXXnnXAnnAAnnnnXnnnnnnnXn0nnAn0n MSFT-WIN-3/10.0.17134
        8/17/2023 10:50:16 PM 1/16/2025 10:08:42 AM 1/16/2025 5:45:14 PM 1/16/2025 5:45:14 PM Outlook     nXnXXn0XnnnXn0nnAAXA0XnnX0XnnnnX Outlook-Android/2.0



    -getQuotaUsage parameter, returns details on xo MailboxFolderStatistics and effective Quota, 
        Used with users with mailbox size issues (and/or LegalHold symptoms)

        Note: use of -getQuotaUsage also does an extensive check for LegalHold signs in the mailbox. including reporting on:
            - xoMailbox.LitigationHoldEnabled
            - xoMailbox.InPlaceHolds, 
            - xoMailbox.ComplianceTagHoldApplied
            - xoMailbox.DelayHoldApplied 
            - xoMailbox.DelayReleaseHoldApplied 
            - checks if xoMailboxFolderStats 'DiscoveryHolds' folder has ItemsInFolder -gt 0

    - getPerms parameter, returns Get-xoMailboxPermission & 
        Get-xoRecipientPermission, non-SELF grants, and membership of any grant 
        groups (XO-only)

        - Adds added inline output (per grant and nested group w membership)

            ## xoMailboxPermission:
            Identity   User                       AccessRights
            --------   ----                       ------------
            XAXXxxxxxx ABC-SEC-Email-XAXXxxxxxx-G {FullAccess}

            ### Expanded Perm Group Summaries:
            -----------
            Identity                   | PrimarySmtpAddress
            ABC-XXX-Xxaxx-XAXXxxxxxx-G | ABC-XXX-Xxaxx-XAXXxxxxxx-G@DOMAIN.COM
            RecipientType              | RecipientTypeDetails       | ManagedBy
            MailUniversalSecurityGroup | MailUniversalSecurityGroup | Xxxaxxx Xaxax
            Description :
            #### Members:
            Alias   PrimarySmtpAddress       RecipientType RecipientTypeDetails
            -----   ------------------       ------------- --------------------
            xaxaxxx Xxxaxxx.Xaxax@DOMAIN.COM   UserMailbox   UserMailbox
            ..


            ## xoRecipientPermission:
            Identity   Trustee                    AccessControlType AccessRights Inherited
            --------   -------                    ----------------- ------------ ---------
            XAXXxxxxxx ABC-XXX-Xxaxx-XAXXxxxxxx-G Allow             {SendAs}


            ### Expanded Perm Group Summaries:
            -----------
            Identity                   | PrimarySmtpAddress
            ABC-XXX-Xxaxx-XAXXxxxxxx-G | ABC-XXX-Xxaxx-XAXXxxxxxx-G@DOMAIN.COM
            RecipientType              | RecipientTypeDetails       | ManagedBy
            MailUniversalSecurityGroup | MailUniversalSecurityGroup | Xxxaxxx Xaxax
            Description :
            #### Members:
            Alias   PrimarySmtpAddress       RecipientType RecipientTypeDetails
            -----   ------------------       ------------- --------------------
            xaxaxxx Xxxaxxx.Xaxax@DOMAIN.COM   UserMailbox   UserMailbox
            ...


    - outObject parameter causes it to return a system.object summary to the pipeline. 
        Can be captured in a variable when calling, for further analysis of the components of the resolved user/mailbox object:

         $results = resolve-user -outObject -users 'USERLOGON@DOMAIN.COM'  ;  

         By default, the returned object includes the following properties & full object copies (if found and resolvable):

            dname           : lynctest14@DOMAIN.COMlync test14
            fname           : lynctest14
            lname           : lync test14
            OPRcp           : SD.SUB.DOMAIN.COM/ABC/USERS/lync test14
            xoRcp           : lync test14_0650dc758f
            OPMailbox       :
            OPRemoteMailbox : lync test14
            ADUser          : CN=lync test14,OU=users,OU=SITE,DC=sd,DC=sub,DC=domain,DC=com
            Federator       : SUB.DOMAIN.COM
            xoMailbox       : lync test14
            xoMUser         :
            xoUser          :
            xoMemberOf      :
            txGuest         :
            OPMapiTest      :
            xoMapiTest      : {Microsoft.Exchange.Monitoring.MapiTransactionOutcome}
            MsolUser        :
            AADUser         : class User {}
            AADUserMgr      :
            AADUserLics     :
            LicenseGroup    :
            isDirSynced     : True
            isNoBrain       : False
            isSplitBrain    : False
            IsLicensed      : 0
            IsDisabledOU    : False
            IsADDisabled    : 0
            IsAADDisabled   : 0

    The following items above are substantial copies of the original cloud or OnPrem objects:

        OPRcp           : OnPrem recipient details
        xoRcp           : Cloud recipient details
        OPMailbox       : OnPrem mailbox details (if present)
        OPRemoteMailbox : OnPrem RemoteMailbox details
        ADUser          : OnPrem ActiveDirectory ADUser object details
        xoMailbox       : Cloud mailbox details 
        xoMUser         : Cloud MailUser object details 
        xoUser          : Cloud Exchange Online 'User' object details
        txGuest         : Cloud Guest details
        OPMapiTest      : Results of OnPrem mailbox access tests
        xoMapiTest      : Results of cloud mailbox access tests
        MsolUser        : Cloud MsolUser object details
        AADUser         : Cloud AzureADUser object details
        AADUserMgr      : Cloud subject user's 'ManagedBy' AzureADUser object details

        Each can be accessed, if -outObject was used and the output assigned to a variable, as a dotted-property of the variable ($variable.property):

            PS> $$results.xomailbox

                Name                   Alias      ServerName    ProhibitSendQuota
                ----                   -----      ----------    -----------------
                lync test14_0650dc758f lynctest14 xxnxx0nxxnnnn 10 GB (10,737,418,240 bytes)


    .PARAMETER users
    Array of user descriptors: displayname, emailaddress, UPN, samaccountname (checks clipboard where unspecified)[-users 'xxx','yyy']
    .PARAMETER Ticket
    Ticket Number [-Ticket '999999']
    .PARAMETER getMobile
    switch to return mobiledevice info for target user[-getMobile]
    .PARAMETER getQuotaUsage
    switch to return Quota & MailboxFolderStatistics & LegalHold analysis (XO-only)[-getQuotaUsage]
    .PARAMETER DeletedItems
    switch to return Quota & MailboxFolderStatistics & LegalHold analysis and return information about DeletedItems and RecoverableItems folders(XO-only)[-DeletedItems]
    .PARAMETER getPerms
    switch to return Get-xoMailboxPermission & Get-xoRecipientPermission, non-SELF grants, and membership of any grant groups (XO-only)[-getPerms]
    .PARAMETER ResolveForwards
    switch to resolve MailContact email addresses against the population of forwarded RemoteMailbox objects(XO-only)[-ResolveForwards]
    .PARAMETER xoMobileDeviceOLDThreshold
    Integer days since LastSyncAttemptTime that classifies a MobileDevice as xoMobileDeviceStatsOLD (defaults to 30)[-xoMobileDeviceOLDThreshold 45]
    .PARAMETER rgxAccentedNameChars
    Regular Expression that identifies input 'user' strings that should have diacriticals/latin/non-simple english characters replaced, before lookups (has default value, used to override for future temp exclusion)[-rgxAccentedNameChars `$rgx]
    .PARAMETER TenOrg
    TenantTag value, indicating Tenants to connect to[-TenOrg 'ABC']
    .PARAMETER Credential
    Use specific Credentials (defaults to Tenant-defined SvcAccount)[-Credentials [credential object]]
    .PARAMETER UserRole
    Credential User Role spec (SID|CSID|UID|B2BI|CSVC|ESVC|LSVC|ESvcCBA|CSvcCBA|SIDCBA)[-UserRole @('SIDCBA','SID','CSVC')]
    .PARAMETER useEXOv2
    Use EXOv2 (ExchangeOnlineManagement) over basic auth legacy connection [-useEXOv2]
    .PARAMETER silent
    Silent output (suppress status echos, defaults true)[-silent]
    .PARAMETER outObject
    switch to return a system.object summary to the pipeline[-outObject]
    .PARAMETER useJson
    switch to format object pipeline & file exports in JSON format (vs default .xml, avoids issues with export-cliXML and MG module data)[-useJson]
    .PARAMETER Depth
    Integer 'depth' to be used for export-clixml (or convertto-json, where -useJson specified) cmdlet for exporting report object to the output format. Directly impacts resulting file size and extent to which full nested objects are captured in the export. Defaults to 2
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
    PS> $999999Rpt = resolve-user fname.lname@DOMAIN.COM -Ticket 99999 -getQuotaUsage -outObject ; 

        10:39:53:===v (1/1):Input: 'FNAME.LNAME@DOMAIN.COM' | 'FNAME' | 'LNAME' v===(EML)
        get-Rmbx/xMbx: (Rmbx)(TOR USER, fed:SUB.DOMAIN.COM)
        SamAccountName | WindowsEmailAddress
        LNAMEFI         | FNAME.LNAME@DOMAIN.COM
        Office | RecipientTypeDetails | RemoteRecipientType | IsDirSynced
                | RemoteUserMailbox    | Migrated            |
        ExternalDirectoryObjectId | CustomAttribute5 | EmailAddressPolicyEnabled
                                    |                  | True
        Outlook (xoMAPI) Access Test Result:Success
        xoMailboxStats Count:1
        10:39:56: INFO:  (-getQuotaUsage:running lengthy Get-xoMailboxFolderStatistics...)
        =get-AADuser FNAME.LNAME@DOMAIN.COM>:
        =get-AADuserManager FNAME.LNAME@DOMAIN.COM>:
        ===$hSum.ADUser: 
        UPN                 | DName      | FName | LName | Title                             
        FNAME.LNAME@DOMAIN.COM | FNAME LNAME | FNAME | LNAME  | Supervisor II, Distribution Center
        Company | Dept                            | Ofc          
                | Operations Distribution El Paso | El Paso-D, TX
        Street | City | State | Zip | Phone           | Mobile
                |      |       |     | +1 915 231 7404 |
        Enabled | DN                                                          
        True    | CN=FNAME LNAME,OU=Users,OU=ELP,DC=SD,DC=sub,DC=domain,DC=com
        whenCreated          | whenChanged         
        8/18/2017 4:13:54 PM | 2/23/2024 8:23:33 AM
        Desc : 8/21/17 FT for FNAME LNAME 146294 -bk
        LicenseGroup:(direct-assigned E3)
        ===$hSum.MGUserMgr: 
        UserPrincipalName       | Mail                   
        FNAME.LNAME@DOMAIN.COM | FNAME.LNAME@DOMAIN.COM
        OpOU : OU=Users,OU=ELP,DC=SD,DC=sub,DC=domain,DC=com
        10:40:06: PROMPT:  UserPrincipalName       | Mail                   
        FNAME.LNAME@DOMAIN.COM | FNAME.LNAME@DOMAIN.COM
        OpOU : OU=Users,OU=ELP,DC=SD,DC=sub,DC=domain,DC=com

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
        D:\scripts\logs\823795-FNAME.LNAME@DOMAIN.COM-folder-sizes-NONHIDDEN-NONZERO-run20240227-1039AM.xml

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
    .EXAMPLE
    PS> $999999Rpt = resolve-user fname.lname@DOMAIN.COM -Ticket 99999 -getPerms -outObject ; 

        # [... additional Permissions output returned]
        10:42:56: PROMPT:
        ## xoMailboxPermission:
        Identity                             User                          AccessRights
        --------                             ----                          ------------
        xx299x9x-x51x-4562-8xx8-x2x45796x2xx ABC-SEC-Email-xxxxxxxxxxxxx-G {FullAccess}

        ### Expanded Perm Group Summaries:
        -----------
        Identity                             | PrimarySmtpAddress
        522x58x1-11x9-4x28-x391-1x8xxx211xxx | ABC-SEC-Email-xxxxxxxxxxxxx-G@DOMAIN.COM
        RecipientType              | RecipientTypeDetails       | ManagedBy
        MailUniversalSecurityGroup | MailUniversalSecurityGroup | Christie Moore
        Description :
        #### Members:
        Alias   PrimarySmtpAddress        RecipientType RecipientTypeDetails
        -----   ------------------        ------------- --------------------
        xxxxxxx xxxxxxxx.xxxxx@DOMAIN.COM UserMailbox   UserMailbox

        ## xoRecipientPermission:
        Identity                             Trustee                              AccessControlType AccessRights Inherited
        --------                             -------                              ----------------- ------------ ---------
        xx299x9x-x51x-4562-8xx8-x2x45796x2xx 522x58x1-11x9-4x28-x391-1x8xxx211xxx Allow             {SendAs}

        ### Expanded Perm Group Summaries:
        -----------
        Identity                             | PrimarySmtpAddress
        522x58x1-11x9-4x28-x391-1x8xxx211xxx | ABC-SEC-Email-xxxxxxxxxxxxx-G@DOMAIN.COM
        RecipientType              | RecipientTypeDetails       | ManagedBy
        MailUniversalSecurityGroup | MailUniversalSecurityGroup | Christie Moore
        Description :
        #### Members:
        Alias   PrimarySmtpAddress        RecipientType RecipientTypeDetails
        -----   ------------------        ------------- --------------------
        xxxxxxx xxxxxxxx.xxxxx@DOMAIN.COM UserMailbox   UserMailbox
    .EXAMPLE
    PS> $999999Rpt = resolve-user fname.lname@DOMAIN.COM -Ticket 99999 -getMobile -outObject ;
        
            .EXAMPLE
    PS> $results = resolve-user -users 'John Public' -getmobile -outobject ; 
        
        ...
        xoMobileDeviceStats Count:2
        ...

        $results.xoMobileDeviceStats: 

        FirstSyncTime         LastPolicyUpdateTime  LastSyncAttemptTime  LastSuccessSync      DeviceType  DeviceID                         DeviceUserAgent       DeviceWipeSentTime DeviceWipeRequestTime DeviceWipeAckTime
        -------------         --------------------  -------------------  ---------------      ----------  --------                         ---------------       ------------------ --------------------- -----------------
        8/17/2023 10:50:16 PM 1/16/2025 10:08:42 AM 1/16/2025 5:45:14 PM 1/16/2025 5:45:14 PM Outlook     nXnXXn0XnnnXn0nnAAXA0XnnX0XnnnnX Outlook-Android/2.0
    
    Demo with the -getMobile parameter, to return details on xo MobileDevices in use with the EXO mailbox. Demos default output 'xoMobileDeviceStats Count' echo, and detailed xoMobileDeviceStats object output
    .EXAMPLE
    PS> resolve-user -users FNAME.LNAME@DOMAIN.com -Ticket 999999 -usejson -depth 2 ; 
    Example using json output, to Depth of 2.
    .LINK
    https://github.com/tostka/verb-exo
    #>

    #Requires -Modules ActiveDirectory, MSOnline, AzureAD, ExchangeOnlineManagement, verb-AAD, verb-ADMS, verb-Auth, verb-IO, verb-logging
    ##Requires -RunasAdministrator
    ## [OutputType('bool')] # optional specified output type
    [CmdletBinding()]
    [Alias('ulu')]
    PARAM(
        [Parameter(Position=0,Mandatory=$False,ValueFromPipeline=$true,HelpMessage="Array of user descriptors: displayname, emailaddress, UPN, samaccountname (checks clipboard where unspecified)[-users 'xxx','yyy']")]
            #[ValidateNotNullOrEmpty()] # pulls string from clipboard if not populated
            [Alias('UserPrincipalName', 'Samaccountname','DisplayName','Name')]
            [array]$users,
        [Parameter(Mandatory=$False,HelpMessage="Ticket Number [-Ticket '999999']")]
            [string]$Ticket,
        [Parameter(HelpMessage="switch to return mobiledevice info for target user[-getMobile]")]
            [Alias('Mobile')]
            [switch] $getMobile,
        [Parameter(HelpMessage="switch to return Quota & MailboxFolderStatistics & LegalHold analysis (XO-only)[-getQuotaUsage]")]
            [Alias('Quota')]
            [switch]$getQuotaUsage,
        [Parameter(HelpMessage="switch to return Quota & MailboxFolderStatistics & LegalHold analysis and return information about DeletedItems and RecoverableItems folders(XO-only)[-DeletedItems]")]
            #[Alias('')]
            [switch]$DeletedItems,
        [Parameter(HelpMessage="switch to return Get-xoMailboxPermission & Get-xoRecipientPermission, non-SELF grants, and membership of any grant groups (XO-only)[-getPerms]")]
            [Alias('Perms','getPermissions')]
            [switch]$getPerms,
        [Parameter(HelpMessage="switch to resolve MailContact email addresses against the population of forwarded RemoteMailbox objects(XO-only)[-ResolveForwards]")]
            [switch]$ResolveForwards,
        [Parameter(HelpMessage="Integer days since LastSyncAttemptTime that classifies a MobileDevice as xoMobileDeviceStatsOLD (defaults to 30)[-xoMobileDeviceOLDThreshold 45]")]
            [int]$xoMobileDeviceOLDThreshold = 30,
        [Parameter(HelpMessage="Regular Expression that identifies input 'user' strings that should have diacriticals/latin/non-simple english characters replaced, before lookups (has default value, used to override for future temp exclusion)[-rgxAccentedNameChars `$rgx]")]
            [ValidateNotNullOrEmpty()]
            [regex]$rgxAccentedNameChars = "[^a-zA-Z0-9\s\.\(\)\{\}\/\&\$\#\@\,\`"\'\’\:\–_-]",
        [Parameter(Mandatory=$FALSE,HelpMessage="TenantTag value, indicating Tenants to connect to[-TenOrg 'ABC']")]
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
        [Parameter(HelpMessage="Silent output (suppress status echos, defaults true)[-silent]")]
            [switch] $silent=$true,
        [Parameter(HelpMessage="switch to return a system.object summary to the pipeline[-outObject]")]
            [switch] $outObject,        
        [Parameter(HelpMessage="switch to format object pipeline & file exports in JSON format (vs default .xml, avoids issues with export-cliXML and MG module data)[-useJson]")]
            [switch]$useJson,
        [Parameter(HelpMessage="Integer 'depth' to be used for export-clixml (or convertto-json, where -useJson specified) cmdlet for exporting report object to the output format. Directly impacts resulting file size and extent to which full nested objects are captured in the export. Defaults to 2")]
            [int]$Depth = 2
    ) ;
    BEGIN{
        #region CONSTANTS_AND_ENVIRO ; #*======v CONSTANTS_AND_ENVIRO v======
        #region CONSTANTS_AND_ENVIRO ; #*======v CONSTANTS_AND_ENVIRO v======
        #region ENVIRO_DISCOVER ; #*------v ENVIRO_DISCOVER v------
        push-TLSLatest
        $Verbose = [boolean]($VerbosePreference -eq 'Continue') ; 
        $rPSCmdlet = $PSCmdlet ; # an object that represents the cmdlet or advanced function that's being run. Available on functions w CmdletBinding (& $args will not be available). (Blank on non-CmdletBinding/Non-Adv funcs).
        $rPSScriptRoot = $PSScriptRoot ; # the full path of the executing script's parent directory., PS2: valid only in script modules (.psm1). PS3+:it's valid in all scripts. (Funcs: ParentDir of the file that hosts the func)
        $rPSCommandPath = $PSCommandPath ; # the full path and filename of the script that's being run, or file hosting the funct. Valid in all scripts.
        $rMyInvocation = $MyInvocation ; # populated only for scripts, function, and script blocks.
        # - $MyInvocation.MyCommand.Name returns name of a function, to identify the current command,  name of the current script (pop'd w func name, on Advfuncs)
        # - Ps3+:$MyInvocation.PSScriptRoot : full path to the script that invoked the current command. The value of this property is populated only when the caller is a script (blank on funcs & Advfuncs)
        # - Ps3+:$MyInvocation.PSCommandPath : full path and filename of the script that invoked the current command. The value of this property is populated only when the caller is a script (blank on funcs & Advfuncs)
        #     ** note: above pair contain information about the _invoker or calling script_, not the current script
        $rPSBoundParameters = $PSBoundParameters ; 
        #region PREF_VARI_DUMP ; #*------v PREF_VARI_DUMP v------
        <#$script:prefVaris = @{
            whatifIsPresent = $whatif.IsPresent
            whatifPSBoundParametersContains = $rPSBoundParameters.ContainsKey('WhatIf') ; 
            whatifPSBoundParameters = $rPSBoundParameters['WhatIf'] ;
            WhatIfPreferenceIsPresent = $WhatIfPreference.IsPresent ; # -eq $true
            WhatIfPreferenceValue = $WhatIfPreference;
            WhatIfPreferenceParentScopeValue = (Get-Variable WhatIfPreference -Scope 1).Value ;
            ConfirmPSBoundParametersContains = $rPSBoundParameters.ContainsKey('Confirm') ; 
            ConfirmPSBoundParameters = $rPSBoundParameters['Confirm'];
            ConfirmPreferenceIsPresent = $ConfirmPreference.IsPresent ; # -eq $true
            ConfirmPreferenceValue = $ConfirmPreference ;
            ConfirmPreferenceParentScopeValue = (Get-Variable ConfirmPreference -Scope 1).Value ; 
            VerbosePSBoundParametersContains = $rPSBoundParameters.ContainsKey('Confirm') ; 
            VerbosePSBoundParameters = $rPSBoundParameters['Verbose'] ;
            VerbosePreferenceIsPresent = $VerbosePreference.IsPresent ; # -eq $true
            VerbosePreferenceValue = $VerbosePreference ;
            VerbosePreferenceParentScopeValue = (Get-Variable VerbosePreference -Scope 1).Value;
            VerboseMyInvContains = '-Verbose' -in $rPSBoundParameters.UnboundArguments ; 
            VerbosePSBoundParametersUnboundArgumentContains = '-Verbose' -in $rPSBoundParameters.UnboundArguments 
        } ;
        write-verbose "`n$(($script:prefVaris.GetEnumerator() | Sort-Object Key | Format-Table Key,Value -AutoSize|out-string).trim())`n" ; 
        #>
        #endregion PREF_VARI_DUMP ; #*------^ END PREF_VARI_DUMP ^------
        #region RV_ENVIRO ; #*------v RV_ENVIRO v------
        $pltRvEnv=[ordered]@{
            PSCmdletproxy = $rPSCmdlet ; 
            PSScriptRootproxy = $rPSScriptRoot ; 
            PSCommandPathproxy = $rPSCommandPath ; 
            MyInvocationproxy = $rMyInvocation ;
            PSBoundParametersproxy = $rPSBoundParameters
            verbose = [boolean]($PSBoundParameters['Verbose'] -eq $true) ; 
        } ;
        write-verbose "(Purge no value keys from splat)" ; 
        $mts = $pltRVEnv.GetEnumerator() |?{$_.value -eq $null} ; $mts |%{$pltRVEnv.remove($_.Name)} ; rv mts -ea 0 -whatif:$false -confirm:$false; 
        $smsg = "resolve-EnvironmentTDO w`n$(($pltRVEnv|out-string).trim())" ; 
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        if(get-command resolve-EnvironmentTDO -ea STOP){}ELSE{
            $smsg = "UNABLE TO gcm resolve-EnvironmentTDO!" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
            BREAK ; 
        } ; 
        $rvEnv = resolve-EnvironmentTDO @pltRVEnv ; 
        $smsg = "`$rvEnv returned:`n$(($rvEnv |out-string).trim())" ; 
        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        #endregion RV_ENVIRO ; #*------^ END RV_ENVIRO ^------
        #region NETWORK_INFO ; #*======v NETWORK_INFO v======
        if(get-command resolve-NetworkLocalTDO  -ea STOP){}ELSE{
            $smsg = "UNABLE TO gcm resolve-NetworkLocalTDO !" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
            BREAK ; 
        } ; 
        $netsettings = resolve-NetworkLocalTDO ; 
        if($env:Userdomain){ 
            switch($env:Userdomain){
                'CMW'{
                    #$logon_SID = $CMW_logon_SID 
                }
                'TORO'{
                    #$o365_SIDUpn = $o365_Toroco_SIDUpn ; 
                    #$logon_SID = $TOR_logon_SID ; 
                }
                $env:COMPUTERNAME{
                    $smsg = "%USERDOMAIN% -EQ %COMPUTERNAME%: $($env:computername) => non-domain-connected, likely edge role Ex server!" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                    if($netsettings.Workgroup){
                        $smsg = "WorkgroupName:$($netsettings.Workgroup)" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;                    
                    } ; 
                } ; 
                default{
                    $smsg = "$($env:userdomain):UNRECOGIZED/UNCONFIGURED USER DOMAIN STRING!" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                    THROW $SMSG 
                    BREAK ; 
                }
            } ; 
        } ;  # $env:Userdomain-E
        #endregion NETWORK_INFO ; #*======^ END NETWORK_INFO ^======
        #region OS_INFO ; #*------v OS_INFO v------
        <# os detect, covers Server 2016, 2008 R2, Windows 10, 11
        if (get-command get-ciminstance -ea 0) {$OS = (Get-ciminstance -class Win32_OperatingSystem)} else {$Os = Get-WMIObject -class Win32_OperatingSystem } ;
        #$isWorkstationOS = $isServerOS = $isW2010 = $isW2011 = $isS2016 = $isS2008R2 = $false ;
        write-host "Detected:`$Os.Name:$($OS.name)`n`$Os.Version:$($Os.Version)" ;
        if ($OS.name -match 'Microsoft\sWindows\sServer') {
            $isServerOS = $true ;
            if ($os.name -match 'Microsoft\sWindows\sServer\s2016'){$isS2016 = $true ;} ;
            if ($os.name -match 'Microsoft\sWindows\sServer\s2008\sR2') { $isS2008R2 = $true ; } ;
        } else { 
            if ($os.name -match '^Microsoft\sWindows\s11') {
                $isWorkstationOS = $true ;
                if ($os.name -match 'Microsoft\sWindows\s11') { $isW2011 = $true ; } ;
            } elseif ($os.name -match '^Microsoft\sWindows\s10') {
                $isWorkstationOS = $true ; $isW2010 = $true
            } else {
                $isWorkstationOS = $true ;
            } ;         
        } ; 
        #>
        #endregion OS_INFO ; #*------^ END OS_INFO ^------
        #region TEST_EXOPLOCAL ; #*------v TEST_EXOPLOCAL v------
        if(get-command test-LocalExchangeInfoTDO -ea STOP){}ELSE{
            $smsg = "UNABLE TO gcm test-LocalExchangeInfoTDO !" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
            BREAK ; 
        } ; 
        $lclExOP = test-LocalExchangeInfoTDO ; 
        write-verbose "Expand returned NoteProperty properties into matching local variables" ; 
        if($host.version.major -gt 2){
            $lclExOP.PsObject.Properties | ?{$_.membertype -eq 'NoteProperty'} | foreach-object{set-variable -name $_.name -value $_.value -verbose -whatif:$false -Confirm:$false ;} ;
        }else{
            write-verbose "Psv2 lacks the above expansion capability; just create simpler variable set" ; 
            $ExVers = $lclExOP.ExVers ; $isLocalExchangeServer = $lclExOP.isLocalExchangeServer ; $IsEdgeTransport = $lclExOP.IsEdgeTransport ;
        } ;
        #
        #endregion TEST_EXOPLOCAL ; #*------^ END TEST_EXOPLOCAL ^------

                                                        <#
    #region PsParams ; #*------v PSPARAMS v------
    $PSParameters = New-Object -TypeName PSObject -Property $rPSBoundParameters ;
    # DIFFERENCES $PSParameters vs $PSBoundParameters:
    # - $PSBoundParameters: System.Management.Automation.PSBoundParametersDictionary (native obj)
    # test/access: ($PSBoundParameters['Verbose'] -eq $true) ; $PSBoundParameters.ContainsKey('Referrer') #hash syntax
    # CAN use as a @PSBoundParameters splat to push through (make sure populated, can fail if wrong type of wrapping code)
    # - $PSParameters: System.Management.Automation.PSCustomObject (created obj)
    # test/access: ($PSParameters.verbose -eq $true) ; $PSParameters.psobject.Properties.name -contains 'SenderAddress' ; # cobj syntax
    # CANNOT use as a @splat to push through (it's a cobj)
    write-verbose "`$rPSBoundParameters:`n$(($rPSBoundParameters|out-string).trim())" ;
    # pre psv2, no $rPSBoundParameters autovari to check, so back them out:
    #>
        <# recycling $rPSBoundParameters into @splat calls: (can't use $psParams, it's a cobj, not a hash!)
        # rgx for filtering $rPSBoundParameters for params to pass on in recursive calls (excludes keys matching below)
        $rgxBoundParamsExcl = '^(Name|RawOutput|Server|Referrer)$' ; 
        if($rPSBoundParameters){
                $pltRvSPFRec = [ordered]@{} ;
                # add the specific Name for this call, and Server spec (which defaults, is generally not 
                $pltRvSPFRec.add('Name',"$RedirectRecord" ) ;
                $pltRvSPFRec.add('Referrer',$Name) ; 
                $pltRvSPFRec.add('Server',$Server ) ;
                $rPSBoundParameters.GetEnumerator() | ?{ $_.key -notmatch $rgxBoundParamsExcl} | foreach-object { $pltRvSPFRec.add($_.key,$_.value)  } ;
                write-host "Resolve-SPFRecord w`n$(($pltRvSPFRec|out-string).trim())" ;
                Resolve-SPFRecord @pltRvSPFRec  | write-output ;
        } else {
            $smsg = "unpopulated `$rPSBoundParameters!" ;
            write-warning $smsg ;
            throw $smsg ;
        };     
        #>
        #endregion PsParams ; #*------^ END PSPARAMS ^------    
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
        if(-not $XODefaultPrefix){$XODefaultPrefix = 'xo' };
        if(-not $SCDefaultPrefix){$SCDefaultPrefix = 'sc' };
        #$rgxADDistNameGAT = ",$(($TORMeta.UnreplicatedOU -split ',' | select -skip 1 ) -join ',')" 
        #$rgxADDistNameAT = ",$(($TORMeta.UnreplicatedOU -split ',' | select -skip 2 ) -join ',')"

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
        # $ByPassLocalExchangeServerTest = $true # rough in, code exists below for exempting service/regkey testing on this variable status. Not yet implemented beyond the exemption code, ported in from orig source.
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
        # interpolate from TORMETA
        $rgxADDistNameGAT = ",$(($TORMeta.UnreplicatedOU -split ',' | select -skip 1 ) -join ',')" 
        $rgxADDistNameAT = ",$(($TORMeta.UnreplicatedOU -split ',' | select -skip 2 ) -join ',')"
        #$DNDOM = @() ; 
        #$TORMeta.adforestname.split('.') | %{$dndom += "DC=$($_)"} ;
        #$rgxADDistNameAT = [regex]::Escape($DNDOM -join ',') ; 

        # props dyn filtering: write-host "=get-xMbx:>`n$(($hSum.xoMailbox |fl ($xMprops |?{$_     -notmatch '(sam.*|dist.*)'})|out-string).trim())`n-Title:$($hSum.xoUser.Title)";
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
            'Description','Info','whenCreated','whenChanged'

        #'samaccountname','UserPrincipalName','distinguishedname','Description','title','whenCreated','whenChanged','Enabled','sAMAccountType','userAccountControl' ;
        $propsADUsht = 'Enabled','Description','whenCreated','whenChanged','Title' ;
        $propsAADU = 'UserPrincipalName','DisplayName','GivenName','Surname','Title','Company','Department','PhysicalDeliveryOfficeName',
            'StreetAddress','City','State','PostalCode','TelephoneNumber','MobilePhone','Enabled','DistinguishedName' ;
        #'UserPrincipalName','name','ImmutableId','DirSyncEnabled','LastDirSyncTime','AccountEnabled' ;
        $propsMGU = 'UserPrincipalName','DisplayName','GivenName','Surname','JobTitle','CompanyName','Department','officeLocation',
            'StreetAddress','City','State','PostalCode','BusinessPhones','MobilePhone','accountEnabled'
        # 12:34 PM 12/11/2025 doesn't exist in MGU proprs (onprem,'onPremisesDistinguishedName' is closest, OnPrem OU) ;
        #'UserPrincipalName','name','ImmutableId','DirSyncEnabled','LastDirSyncTime','AccountEnabled' ;
        # 3:59 PM 10/9/2024 used for complete miss gadu search results props
        $prpADU = 'DistinguishedName','GivenName','Surname','Name','UserPrincipalName','mailNickname','SamAccountName','physicalDeliveryOfficeName','msExchRecipientDisplayType','msExchRecipientTypeDetails','msExchRemoteRecipientType','msExchWhenMailboxCreated' ; 
        $propsAADUfed = 'UserPrincipalName','name','ImmutableId','DirSyncEnabled','LastDirSyncTime' ;
        $propsRcpTbl = 'Alias','PrimarySmtpAddress','RecipientType','RecipientTypeDetails' ;
        $propsDG = 'Identity','PrimarySmtpAddress','Description','RecipientType','RecipientTypeDetails','ManagedBy' ; 
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
        $propsADL7 = 'Info' ;
        
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
        # MGU equivs
        $propsMGUL1 = @{Name='UPN';Expression={$_.UserPrincipalName }}, @{Name='DName';Expression={$_.DisplayName }},
            @{Name='FName';Expression={$_.GivenName }},@{Name='LName';Expression={$_.Surname }},
            @{Name='Title';Expression={$_.JobTitle }};
        $propsMGUL2 = @{Name='Company';Expression={$_.CompanyName }},@{Name='Dept';Expression={$_.Department }},
            @{Name='Ofc';Expression={$_.officeLocation }} ;    
        $propsMGUL3 = @{Name='Street';Expression={$_.StreetAddress }}, 'City','State',
            @{Name='Zip';Expression={$_.PostalCode }}, @{Name='Phone';Expression={$_.BusinessPhones }},
            @{Name='Mobile';Expression={$_.MobilePhone }} ; 
        $propsMGUL4 = @{Name='Dsync';Expression={$_.OnPremisesSyncEnabled }}, @{Name='ImutID';Expression={$_.OnPremisesImmutableId }},
            @{Name='LastDSync';Expression={$_.OnPremisesLastSyncDateTime }}, @{Name='UseLoc';Expression={$_.UsageLocation }};
        $propsMGUL5 = @{Name='ObjectType';Expression={'User'}}, @{Name='UserType';Expression={$_.UserType}}
        # configured ObjectType as a static 'User', as it's no longer a property and implied by the underlying class/object type returned.

        #$propsAADMgr = 'UserPrincipalName','Mail',@{Name='OpDN';Expression={$_.ExtensionProperty.onPremisesDistinguishedName }} ;
        # get mgr OU, not DN: ExtensionProperty.onPremisesDistinguishedName.split(',') | select -skip 1 ) -join ','
        $propsAADMgr = 'UserPrincipalName','Mail',
            @{Name='OpOU';Expression={($_.ExtensionProperty.onPremisesDistinguishedName.split(',') | select -skip 1) -join ',' }} ;
        $propsAADMgrL1 = 'UserPrincipalName','Mail' ;
        $propsAADMgrL2 = @{Name='OpOU';Expression={($_.ExtensionProperty.onPremisesDistinguishedName.split(',') | select -skip 1) -join ',' }} ;
        # MGU vers
        $propsMGUMgr = @{Name='userPrincipalName';Expression={$_.userPrincipalName }},
            @{Name='mail';Expression={$_.mail}} ; 
        #$propsMGUMgrL1 = 'UserPrincipalName','Mail' ;
        #$propsMGUMgrL2 = @{Name='OpOU';Expression={($_.onPremisesDistinguishedName.split(',') | select -skip 1) -join ',' }} ;
        # NOTE w mg, and the mgUserManager, it's a hashtable/dictionary, and to get it to return key/values, YOU HAVE TO EXACTLY MATCH THE CASE OF THE KEY! userPrincipalName <> UserPrincipalName (1st works, 2nd doesn't)
        $propsMGUMgrL1 = @{Name='userPrincipalName';Expression={$_.userPrincipalName }},
            @{Name='mail';Expression={$_.mail}} 
        $propsMGUMgrL2 = @{Name='OpOU';Expression={ ($_.onPremisesDistinguishedName.split(',') | select -skip 1) -join ',' }} ;  
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
            <#
            $propsMobL2 = @{Name='1stSyncTime';Expression={(get-date $_.FirstSyncTime -format 'M/d/yy H:mmtt') }},
                @{Name='LastSyncTime';Expression={(get-date $_.LastSyncAttemptTime -format 'M/d/yy H:mmtt') }},
                @{Name='LastSuccSync';Expression={(get-date $_.LastSuccessSync -format 'M/d/yy H:mmtt') }},
                @{Name='#Folders';Expression={$_.NumberOfFoldersSynced }} ; 
            #>
            # converttimes to local
            $propsMobL2 = @{Name='1stSyncTime';Expression={(get-date $_.FirstSyncTime.ToLocalTime() -format 'M/d/yy H:mm') }},
                @{Name='LastSyncTime';Expression={(get-date $_.LastSyncAttemptTime.ToLocalTime() -format 'M/d/yy H:mm') }},
                @{Name='LastSuccSync';Expression={(get-date $_.LastSuccessSync.ToLocalTime() -format 'M/d/yy H:mm') }},
                @{Name='#Folders';Expression={$_.NumberOfFoldersSynced }} ; 
            # add for tight summaries
            $prpEASDevs = 'DeviceFriendlyName','ClientType',@{Name='LastSyncTime';Expression={(get-date $_.LastSyncAttemptTime.ToLocalTime() -format 'M/d/yy H:mm') }},
                @{Name='LastSuccSync';Expression={(get-date $_.LastSuccessSync.ToLocalTime() -format 'M/d/yy H:mm') }} ; 
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

            $prpFldrDeleted = @{Name='Folder'; Expression={$_.Identity.tostring()}},@{Name="Items"; Expression={$_.ItemsInFolder}}, 
                @{n="SizeMB"; e={[math]::round($_.FolderSize.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1MB,2)}}, 
                @{n="TreeSizeMB"; e={[math]::round($_.FolderAndSubfolderSize.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1MB,2)}}, 
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
            $rgxDelItmsShow = '.*\\(Deleted Items|Recoverable Items|Deletions|DiscoveryHolds|Purges|SubstrateHolds|Versions)$' ; 

        } ; 
        # 2:31 PM 12/26/2024
        # getPerms
        if($getPerms){

            # 12:54 PM 9/18/2023 adds for MbxFolderStats, Quota & LegalHold eval:
            $prpRPerms = 'Identity','Trustee','AccessControlType','AccessRights','Inherited' ;

            $prpMPerms = 'Identity','User','AccessRights'

        } ; 
        $rgxOPLic = '^CN\=ENT\-APP\-Office365\-(EXOK|F1|MF1)-DL$' ;
        $rgxXLic = '^CN\=ENT\-APP\-Office365\-(EXOK|F1|MF1)-DL$' ;
        #endregion LOCAL_CONSTANTS ; #*------^ END LOCAL_CONSTANTS ^------        
    
        #region ENCODED_CONTANTS ; #*------v ENCODED_CONTANTS v------
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
        # FOLLOWING CONSTANTS ARE USED FOR DEPENDANCY-LESS CONNECTIONS
        if(-not $CMW_logon_SID){$CMW_logon_SID = 'Q01XXGQtdG9kZC5rYWRyaWU=' | fb4 } ; 
        if(-not $o365_Toroco_SIDUpn){$o365_Toroco_SIDUpn = 'cy10b2RkLmthZHJpZUB0b3JvLmNvbQ==' | fb4 } ;
        if(-not $TOR_logon_SID){$TOR_logon_SID = 'VE9ST1xrYWRyaXRzcw==' | fb4 } ; 

        #endregion ENCODED_CONTANTS ; #*------^ END ENCODED_CONTANTS ^------
    
        #endregion CONSTANTS_AND_ENVIRO ; #*======^ CONSTANTS_AND_ENVIRO ^======
    
        #region WHPASSFAIL ; #*======v WHPASSFAIL v======
        $whTPad = 72  ; $whTChar = '.' ; # scale $whTPad to longest Testing:xxx line you use in the test array
        if(-not $whPASS){$whPASS = @{ Object = "$([Char]8730) PASS`n" ; ForegroundColor = 'Green' ; NoNewLine = $true  } }
        if(-not $whFAIL){$whFAIL = @{'Object'= if ($env:WT_SESSION) { "$([Char]8730) FAIL`n"} else {" !X! FAIL`n"}; ForegroundColor = 'RED' ; NoNewLine = $true } } ;
        # light diagonal cross: ╳ U+2573 DOESN'T RENDER IN PS, use it if WinTerm
        if(-not $psPASS){$psPASS = "$([Char]8730) PASS`n" } # $smsg = $pspass + " :Tested Drives" ; write-host $smsg ;
        if(-not $psFAIL){$psFAIL = if ($env:WT_SESSION) { "$([Char]8730) FAIL`n"} else {" !X! FAIL`n"} } ; # $smsg = $psfail + " :Tested Drives" ; write-warning $smsg ;    
        <# WHPASSFAIL:SAMPLE TESTS:
        #region WHPASSFAILSimpleTest ; #*------v WHPASSFAILSimpleTest v------
        $tFormat = 'NTFS' ; 
        $smsg = "Testing: Volume.FileSystem against: $($tFormat)" ; #Write-Host "$($smsg)... " -NoNewline ;
        $smsg += " $($whTChar * ($whTPad - $smsg.length))" ; Write-Host "$($smsg) " -NoNewline ;
        if ($VOL.FileSystem -eq $tFormat) {Write-Host @whPASS} else {write-host @whFAIL };
        #endregion WHPASSFAILSimpleTest ; #*------^ END WHPASSFAILSimpleTest ^------    
        #region WHPASSFAILCapacityTest ; #*------v WHPASSFAILCapacityTest v------
        # Test: Capacity match, threshold vs %:
        $tSpaceThresh = 10 * 1GB ; # .9 (for %)
        if($tSpaceThresh -gt 1000){ $smsg = "Testing: Volume.SizeRemainingStatus against: $(RndTo3($tSpaceThresh/1GB))GB" }
        elseif($tSpaceThresh -lt 1){$smsg = "Testing: Volume.SizeRemainingStatus against: $(RndTo3($tSpaceThresh * 100))%" }
        else {$smsg = "Testing: Volume.SizeRemainingStatus against: $($tSpaceThresh)" } ;
        $smsg += " $($whTChar * ($whTPad - $smsg.length))" ; Write-Host "$($smsg) " -NoNewline ;
        if($VerbosePreference -eq 'Continue'){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
        if($tSpaceThresh -lt 1){
            $smsg = "Detected $($tSpaceThresh) is a percentage free test" ;
            if($VerbosePreference -eq 'Continue'){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
            if ($tv.SizeRemaining / $tv.Size -lt $tSpaceThresh) {
                $rptDrive.SizeRemainingStatus = $false ;
                write-host @whFAIL ;
                $smsg = "Insufficient free space on DB drive: $($tv.DriveLetter): $(RndTo2($tv.SizeRemaining/1GB)) GB, needs at least $($tv.Size/1GB * $tSpaceThresh) GB" ;
                $rptDrive.DriveIssues += @($smsg)
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } else {
                $rptDrive.SizeRemainingStatus = $true ;
                Write-Host @whPASS ;   ;
                $smsg = "DB drive: $($tv.DriveLetter): $(RndTo2($tv.SizeRemaining/1GB)) GB free, sufficient for install" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } ;
        }else{
            $smsg = "Detected $($tSpaceThresh) is a free space floor test" ;
            if($VerbosePreference -eq 'Continue'){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
            if ($tv.SizeRemaining -lt $tSpaceThresh){
                $rptDrive.SizeRemainingStatus = $false ;
                write-host @whFAIL ;
                $smsg = "Insufficient free space on $($rptDrive.DriveRole -join ',') drive: $(RndTo2($tv.SizeRemaining/1GB)) GB, needs at least $(RndTo2($tSpaceThresh/1GB)) GB" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } else {
                $rptDrive.SizeRemainingStatus = $true ;
                Write-Host @whPASS ;   ;
                $smsg = "$($rptDrive.DriveRole -join ','): $(RndTo2($tv.SizeRemaining/1GB)) GB free, sufficient for install" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } ;
        } ; 
        #endregion WHPASSFAILCapacityTest ; #*------^ END WHPASSFAILCapacityTest ^------
        #>
        #endregion WHPASSFAIL ; #*======^ END WHPASSFAIL ^======

        #endregion CONSTANTS_AND_ENVIRO ; #*------^ END CONSTANTS_AND_ENVIRO ^------

        #region CHECK_PREREQS ; #*------v CHECK_PREREQS v------
        # critical dependancy Meta variables
        $MetaNames = 'TOR','CMW','TOL' # ,'NOSUCH' ; 
        # critical dependancy Meta variable properties
        $MetaProps = 'legacyDomain','o365_TenantDomain' #,'DOESNTEXIST' ; 
        $isBased = $true ; $gvMiss = @() ; $ppMiss = @() ; 
        foreach($met in $metanames){
            write-verbose "chk:`$$($met)Meta" ; 
            if(-not (gv -name "$($met)Meta" -ea 0)){
                $isBased = $false; $gvMiss += "$($met)Meta" ; 
            } ; 
            foreach($mp in $MetaProps){
                write-verbose "chk:`$$($met)Meta.$($mp)" ; 
                #if(-not (gv -name "$($met)Meta" -ea 0).value[$mp]){ # testing has a value, not is present as a spec!
                if(-not (gv -name "$($met)Meta" -ea 0).value.keys -contains $mp){$isBased = $false; $ppMiss += "$($met)Meta.$($mp)" ; } ; 
            } ; 
        } ; 
        if($gvmiss){write-warning "Missing Dependant Meta variables:`n$(($gvMiss |%{"`$$($_)" }) -join ',')" } ; 
        if($ppMiss){write-warning "Missing Dependant Meta vari properties:`n$(($ppMiss |%{"`$$($_)" }) -join ',')" } ; 
        if(-not $isBased){ write-warning  "missing critical dependancy profile config!" } ; 
        #endregion CHECK_PREREQS ; #*------^ END CHECK_PREREQS ^------

        #region FUNCTIONS ; #*======v FUNCTIONS v======

        # 2b4() 2b4c() & fb4() are located up in the CONSTANTS_AND_ENVIRO\ENCODED_CONTANTS block ( to convert Constant assignement strings)
        
        #region FUNCTIONS_INTERNAL ; #*======v FUNCTIONS_INTERNAL v======

        #*------v Function resolve-RMbxForwards v------
        function resolve-RMbxForwards(){
            <#
            .SYNOPSIS
            Resolves out all RemoteMailboxes (OnPrem) with ForwardingAddress configured; converts the mailboxes into a hashtable keyed on ForwardingAddress. Returns the hash to the pipeline
            .EXAMPLE
            PS> $hshForwards = resolve-RMbxForwards ; 
            PS> $smsg = "Recipient:$($tid) => $($hshForwards[$tid])" ; 
            PS> write-host $smsg ;
            .NOTES
            VERSION:
            * 3:18 PM 4/12/2025 init
            #>
            write-host "get-remotemailbox  -ResultSize unlimited | ?{`$_.ForwardingAddress}..." ; 
            $fwdRmbxs = get-remotemailbox  -ResultSize unlimited | ?{$_.ForwardingAddress} ; 
            $hshForwards = @{} ;  
            write-host "[" ; 
            $forwardedSummary = $fwdRmbxs |%{
                write-host -NoNewline '.'
                $target = $_ ; 
                $smsg = "$(($target | ft -a primarysmtpaddress,forwardingaddress|out-string).trim())" ; 
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                $fwd = $null; 
                if($fwd = get-recipient -id $target.ForwardingAddress -resultsize 1 | select -expand primarysmtpaddress){
                   $hshForwards[$fwd] = $target ; 
                } ; 
            } ; 
            write-host "]" ; 
            $hshForwards | write-output 
        } ; 
        #*------^ END Function resolve-RMbxForwards ^------

        #endregion GET_XOMOBILEDATA ; #*------^ END get-xoMobileData ^------
        function get-xoMobileData {
            <#
            .SYNOPSIS
            Runs EXO get-xoMobildDevice qrys, and parses results into approp $hSum properties (single common function to reduce dupe queries)
            .EXAMPLE
            PS> get-xoMobileData ;             
            .NOTES
            VERSION:
            * 10:52 AM 9/23/2025init
            #>
            # 
            if($xmbx){
                $smsg = "'xoMobileDeviceStats':Get-xoMobileDeviceStatistics -Mailbox $($xmbx.ExchangeGuid.guid)"
                if($verbose){
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else{ write-verbose $smsg } ;
                } ;
                #$hsum.xoMobileDeviceStats  +=  Get-xoMobileDeviceStatistics -Mailbox $xmbx.userprincipalname -ea STOP ;
                # wasn't getting data back: shift to the .xomailbox.ExchangeGuid.guid, it's 100% going to hit and return data
                $xoMobileDeviceStats +=  Get-xoMobileDeviceStatistics -Mailbox $hSum.xoMailbox.exchangeguid.guid -ea STOP | sort LastSuccesssync -Descending ;
            }else{
                $smsg = "'xoMobileDeviceStats':Get-xoMobileDeviceStatistics -Mailbox $($xmbx.ExchangeGuid.guid)"
                if($verbose){
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else{ write-verbose $smsg } ;
                } ;
                #$hsum.xoMobileDeviceStats  +=  Get-xoMobileDeviceStatistics -Mailbox $xmbx.userprincipalname -ea STOP ;
                # wasn't getting data back: shift to the .xomailbox.ExchangeGuid.guid, it's 100% going to hit and return data
                $xoMobileDeviceStats = +=  Get-xoMobileDeviceStatistics -Mailbox $hSum.xoMailbox.exchangeguid.guid -ea STOP | sort LastSuccesssync -Descending ;
            }
            #$hsum.xoMobileDeviceStats  +=  Get-xoMobileDeviceStatistics -Mailbox $hSum.xoMailbox.exchangeguid.guid -ea STOP ;
            $hsum.xoMobileDeviceStats  +=  @($xoMobileDeviceStats | ?{$_.LastSyncAttemptTime -ge (get-date).adddays(-1 * $xoMobileDeviceOLDThreshold)})
            $smsg = "xoMobileDeviceStats Count:$(($hsum.xoMobileDeviceStats|measure).count)" ;
            $hsum.xoMobileDeviceStatsOLD  +=  @($xoMobileDeviceStats | ?{$_.LastSyncAttemptTime -lt (get-date).adddays(-1 * $xoMobileDeviceOLDThreshold)})
            $smsg += "`nxoMobileDeviceStatsOLD Count:$(($hsum.xoMobileDeviceStatsOLD|measure).count)" ;
            $hsum.xoMobileOutlookClients += @($xoMobileDeviceStats | ?{$_.DeviceType -match 'Outlook' -OR $_.DeviceUserAgent -match 'Outlook' -OR $_.DeviceModel  -match 'Outlook'}) ;
            $hsum.xoMobileOtherClients += @($xoMobileDeviceStats | ?{$_.DeviceType -notmatch 'Outlook' -AND $_.DeviceUserAgent -notmatch 'Outlook' -AND $_.DeviceModel  -notmatch 'Outlook'}) ;
            $hsum.xoMobileOMSyncTypes += @(($hsum.xoMobileOutlookClients | group ClientType | select -expand Name ) -join ';')
            if($hsum.xoMobileOMSyncTypes -match 'REST'){
                $smsg += "`n+User has one or more *legacy* 'REST' Outlook Mobile clients" ;
            }elseif($hsum.xoMobileOMSyncTypes -match 'Outlook'){
                $smsg += "`n+++User has has one or more fully compliant 'MS Native Sync'-protocol Outlook Mobile clients" ;
            } ;
            $hsum.xoMobileOtherSyncTypes += @(($hsum.xoMobileOtherClients | group ClientType | select -expand Name ) -join ';')            
            if($hsum.xoMobileOtherClients| ?{$_.ClientType -eq 'EAS'}){ ;
                $smsg += "`n---User has one or more device-vendor-provided 'ExchangeActiveSync' Mobile clients!" ;
                #$smsg += "`nPLEASE NOTE: BY POLICY EAS CLIENTS ARE *BEST EFFORT* supported:"
                #$smsg += "`nWHERE ISSUES ARE EXPERIENCED WITH LEGACY EAS/ACTIVESYNC CLIENTS," ;
                #$smsg += "`nUSERS SHOULD BE URGED TO MOVE TO SUPPORTED MS OUTLOOK MOBILE FOR IOS OR ANDROID CLIENTS" ;
            }
            write-host -foregroundcolor green $smsg ;
        } ; 
        #endregion GET_XOMOBILEDATA ; #*------^ END get-xoMobileData ^------

        #region GET_MGUSERFULL ; #*------v get-MgUserFull v------
        function get-MgUserFull{
            <#
            .SYNOPSIS
            get-MgUserFull.ps1 - Wrapper for get-MGUser that *forces* it to return a full set of user properties, to approx the get-AzureAdUser that they've taken away, wo less f'ing around retrying queries.
            .NOTES
            Version     : 0.0.
            Author      : Todd Kadrie
            Website     : http://www.toddomation.com
            Twitter     : @tostka / http://twitter.com/tostka
            CreatedDate : 2025-
            FileName    : get-MgUserFull.ps1
            License     : MIT License
            Copyright   : (c) 2025 Todd Kadrie
            Github      : https://github.com/tostka/verb-XXX
            Tags        : Powershell
            AddedCredit : REFERENCE
            AddedWebsite: URL
            AddedTwitter: URL
            REVISIONS
            * 10:46 AM 12/11/2025 reworked $prpMGUser list, added items that are unpop'd propoerties, and pushed useful Additionalproperties from OnPrem, into expansion, updated CBH
            * 12:18 PM 12/10/2025 init
            .DESCRIPTION
            get-MgUserFull.ps1 - Wrapper for get-MGUser that *forces* it to return a full set of user properties, to approx the get-AzureAdUser that they've taken away, wo less f'ing around retrying queries.

            MS has lobotomized get-MgUser as compares to the long-standing functional get-AzureAdUser 
            and returning the full suite of user properties now requires a bunch of horse hockey to retrieve - in favor of their cheesball, money grubbing 'lean' property set. 
            fk-em! We're going to force a full property set return, *every time*
            For fancier filter & top use, use those to return an MGUser with a userid, and then recycle the user ID into this, to retrieve a fully populated user object

            .PARAMETER  UserID
            Useridentifier (UPN, GUID etc) [-UserID UPN@DOMAIN.COM]
            .INPUTS
            None. Does not accepted piped input.(.NET types, can add description)
            .OUTPUTS
            Microsoft.Graph.PowerShell.Models.MicrosoftGraphUser
            System.Boolean
            [| get-member the output to see what .NET obj TypeName is returned, to use here]
            .EXAMPLE
            PS> $mgu = get-MgUserFull -userid UPN@DOMAIN.COM ; 
            Typical call
            .LINK
            https://github.com/tostka/verb-MG
            #>
            [CmdletBinding()]
            PARAM(
                [Parameter(HelpMessage="Array of Useridentifiers (UPN, GUID etc) [-UserID UPN@DOMAIN.COM]")]
                    [ValidateNotNullOrEmpty()]
                    [string[]]$UserID
            )
            BEGIN{
                # FORCE fully populated key user properties (overrides default return of subset garbage)
                $prpMGUser = @(
                    # Identity
                    'id','userPrincipalName','mail','mailNickname','proxyAddresses','otherMails','otherMails',
                    # Display/profile
                    'displayName','givenName','surname','jobTitle','department','companyName',
                    'mobilePhone','businessPhones','preferredLanguage',
                    'jobTitle',
                    'department','companyName',
                    'StreetAddress','city','state','PostalCode','country',
                    'officeLocation','UsageLocation',
                    # Account state
                    'accountEnabled','userType',
                    # Licensing
                    'assignedLicenses','assignedPlans',
                    # Hybrid / sync
                    'onPremisesImmutableId','onPremisesDistinguishedName','onPremisesSecurityIdentifier',
                    'OnPremisesUserPrincipalName','onPremisesSamAccountName','onPremisesDomainName',
                    'onPremisesSyncEnabled','OnPremisesLastSyncDateTime','OnPremisesProvisioningErrors','onPremisesExtensionAttributes'
                    # add prev missing sync data props
                    # Misc often used
                    'creationType', 'CreatedDateTime','DeletedDateTime','EmployeeHireDate','EmployeeId','EmployeeType','HireDate',  
                    'Manager',
                    'LicenseAssignmentStates','LicenseDetails','ProvisionedPlans',
                    'MemberOf',
                    # add AdditionalProperties pulls (should move to primary property, also accessible as .additionalproperties['xxx'] property      
                    'mobilePhone','businessPhones',
                    'preferredLanguage'
                ) | select -unique ; 
                if(-not (get-command Get-MgUser)){
                    $smsg = "Missing Get-MgUser!" ; 
                    $smsg += "`nPre-connect to Microsoft.Graph via:" ;
                    $smsg += "`nConnect-MgGraph -Scopes 'User.Read.All', 'Directory.Read.All', 'Group.Read.All'" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                    BREAK ; 
                } ; 
                #region IS_PIPELINE ; #*------v IS_PIPELINE v------
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
                #endregion IS_PIPELINE ; #*------^ END IS_PIPELINE ^------
            } 
            PROCESS{
                foreach($id in $userid){
                    TRY{
                        $smsg = "Get-MgUser -UserId $($id)" ; 
                        if($VerbosePreference -eq "Continue"){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                        $MGUser = Get-MgUser -UserId $id -Property $prpMGUser -erroraction STOP ; 
                    } CATCH {$ErrTrapd=$Error[0] ;
                        write-host -foregroundcolor gray "TargetCatch:} CATCH [$($ErrTrapd.Exception.GetType().FullName)] {"  ;
                        $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                        CONTINUE
                     } ;            
                    if($MGUser){
                        $MGUser | write-output ; 
                    } else{
                        $smsg = "UNABLE TO: Get-MgUser -UserId $($userid)" ; 
                        if(gcm Write-MyWarning -ea 0){Write-MyWarning $smsg } else {
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN} else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        } ;
                    }; 
                } # loop-E
            } ;  # PROC-E
        } ; 
        #endregion GET_MGUSERFULL ; #*------^ END get-MgUserFull ^------

        #endregion FUNCTIONS_INTERNAL ; #*======^ END FUNCTIONS_INTERNAL ^======

        #region FUNCTIONS_FULLYEXTERNAL ; #*======v FUNCTIONS_FULLYEXTERNAL v======

        #region CONNECT_O365SERVICES ; #*======v verb-exo\connect-O365Services v======
        if(-not (get-childitem function:connect-O365Services -ea 0)){
            $smsg = "MISSING DEPENDANT: verb-exo\connect-O365Services!" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
            break ; 
        } ;
        #endregion CONNECT_O365SERVICES ; #*======^ END verb-exo\connect-o365services ^======
        
        #endregion FUNCTIONS_FULLYEXTERNAL ; #*======^ END FUNCTIONS_FULLYEXTERNAL ^======

        #endregion FUNCTIONS ; #*======^ END FUNCTIONS ^======

        #region SERVICE_CONNECTIONS #*======v END SERVICE_CONNECTIONS v======
        $MGPermissionsScope = @("User.Read.All","Group.Read.All") # get-AzureAdUser baseline requirements, memberof requires groupread as well

        #region BROAD_SVC_CONTROL_VARIS ; #*======v BROAD_SVC_CONTROL_VARIS  v======   
        $useO365 = $true ; 
        $useOP = $true ;     
        # (config individual svcs in each block)
        #endregion BROAD_SVC_CONTROL_VARIS ; #*======^ END BROAD_SVC_CONTROL_VARIS ^======

        #region TEST_EXO_CONN ; #*------v TEST_EXO_CONN v------
        # ALT: simplified verify EXO conn: ALT to full CONNECT_O365SERVICES block - USE ONE OR THE OTHER!
        $useEXO = $true ; 
        $useSC = $false ; 
        if(-not $XOConnectionUri ){$XOConnectionUri = 'https://outlook.office365.com'} ;
        if(-not $SCConnectionUri){$SCConnectionUri = 'https://ps.compliance.protection.outlook.com'} ;
        $EXOtestCmdlet = 'Get-xoOrganizationConfig' ; 
        if(gcm $EXOtestCmdlet -ea 0){
            $conns = Get-ConnectionInformation -ea STOP  ; 
            $hasEXO = $hasSC = $false ; 
            #if($conns | %{$_ | ?{$_.ConnectionUri -eq 'https://outlook.office365.com' -AND $_.State -eq 'Connected' -AND $_.TokenStatus -eq 'Active'}}){
            $conns | %{
                if($_ | ?{$_.ConnectionUri -eq $XOConnectionUri}){$hasEXO = $true } ; 
                if($_ | ?{$_.ConnectionUri -eq $SCConnectionUri}){$hasSC = $true } ; 
            }
            if($useEXO -AND $hasEXO){
                write-verbose "EXO ConnectionURI present" ; 
            }elseif(-not $useEXO){}else{
                $smsg = "No Active EXO connection: Run - Connect-ExchangeOnline -Prefix xo -  before running this script!" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                BREAK ; 
            } ; 
            if($useSC -AND $hasSC){
                write-verbose "SCI ConnectionURI present" ; 
            }elseif(-not $useSC){}else{
                $smsg = "No Active SC connection: Run - Connect-IPPSSession -Prefix SC -  before running this script!" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                BREAK ; 
            } ; 
        }else {
            $smsg = "Missing gcm get-xoMailboxFolderStatistics: ExchangeOnlineManagement module *not* loaded!" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
            BREAK ; 
        } ;     
        #endregion TEST_EXO_CONN ; #*------^ END TEST_EXO_CONN ^------
    
        #region CALL_CONNECT_OPSERVICES ; #*======v CALL_CONNECT_OPSERVICES v======
        #$useOP = $false ; 
        if($useOP){
            $pltCcOPSvcs=[ordered]@{
                # environment parameters:
                EnvSummary = $rvEnv ;
                NetSummary = $netsettings ;
                XoPSummary = $lclExOP ;
                # service choices
                UseExOP = $true ;
                useForestWide = $true ;
                useExopNoDep = $false ;
                ExopVers = 'Ex2010' ;
                UseOPAD = $true ;
                useExOPVers = $useExOPVers; # 'Ex2010' ;
                # Service Connection parameters
                TenOrg = $TenOrg ; # $global:o365_TenOrgDefault ;
                Credential = $Credential ;
                #[ValidateSet("SID","ESVC","LSVC")]
                #UserRole = $UserRole ; # @('SID','ESVC') ;
                # if inheriting same $userrole param/default, that was already used for cloud conn, filter out the op unsupported CBA roles
                # exclude csvc as well, go with filter on the supported ValidateSet from get-HybridOPCredentials: ESVC|LSVC|SID
                UserRole = $UserRole |?{$_ -match '(ESVC|LSVC|SID)' -AND $_ -notmatch 'CBA'} ; # @('SID','ESVC') ;
                # svcAcct use: @('ESvcCBA','CSvcCBA','SIDCBA')
                silent = $silent ;
            } ;
    
            write-verbose "(Purge no value keys from splat)" ;
            $mts = $pltCcOPSvcs.GetEnumerator() |?{$_.value -eq $null} ; $mts |%{$pltCcOPSvcs.remove($_.Name)} ; rv mts -ea 0 ;
            if((get-command connect-OPServices -EA STOP).parameters.ContainsKey('whatif')){
                $pltCcOPSvcsnDSR.add('whatif',$($whatif))
            } ;
            $smsg = "connect-OPServices w`n$(($pltCcOPSvcs|out-string).trim())" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            $ret_CcOPSvcs = connect-OPServices @pltCcOPSvcs ; 
    
            # #region CONFIRM_CCOPRETURN ; #*------v CONFIRM_CCOPRETURN v------
            # matches each: $plt.useXXX:$true to matching returned $ret.hasXXX:$true
            $vplt = $pltCcOPSvcs ; $vret = 'ret_CcOPSvcs' ;  ; $ACtionCommand = 'connect-OPServices' ; 
            $vplt.GetEnumerator() |?{$_.key -match '^use' -ANd $_.value -match $true} | foreach-object{
                $pltkey = $_ ;
                $smsg = "$(($pltkey | ft -HideTableHeaders name,value|out-string).trim())" ; 
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                $vtests = @() ;  $vFailMsgs = @()  ; 
                $tprop = $pltkey.name -replace '^use','has';
                if($rProp = (gv $vret).Value.psobject.properties | ?{$_.name -match $tprop}){
                    $smsg = "$(($rprop | ft -HideTableHeaders name,value |out-string).trim())" ; 
                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                    if($rprop.Value -eq $pltkey.value){
                        $vtests += $true ; 
                        $smsg = "Validated: $($pltKey.name):$($pltKey.value) => $($rprop.name):$($rprop.value)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Success } 
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                    } else {
                        $smsg = "NOT VALIDATED: $($pltKey.name):$($pltKey.value) => $($rprop.name):$($rprop.value)" ;
                        $vtests += $false ; 
                        $vFailMsgs += "`n$($smsg)" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                    };
                } else{
                    $smsg = "Unable to locate: $($pltKey.name):$($pltKey.value) to any matching $($rprop.name)!)" ;
                    $smsg = "" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                } ; 
            } ; 
            if($useOP -AND $vtests -notcontains $false){
                $smsg = "==> $($ACtionCommand): confirmed specified connections *all* successful " ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Success } 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
            }elseif($vtests -contains $false -AND (get-variable ret_CcOPSvcs) -AND (gv -name "$($tenorg)meta").value.o365_opdomain.split('.')[0].toupper() -ne $env:userdomain){
                $smsg = "==> $($ACtionCommand): FAILED SOME SPECIFIED CONNECTIONS" ; 
                $smsg += "`nCROSS-ORG ONPREM CONNECTION: ATTEMPTING TO CONNECT TO ONPREM '$((gv -name "$($tenorg)meta").value.o365_Prefix)' $((gv -name "$($tenorg)meta").value.o365_opdomain.split('.')[0].toupper()) domain, FROM $($env:userdomain)!" ;
                $smsg += "`nEXPECTED ERROR, SKIPPING ONPREM ACCESS STEPS (force `$useOP:$false)" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                $useOP = $false ; 
            }elseif(-not $useOP -AND -not (get-variable ret_CcOPSvcs)){
                $smsg = "-useOP: $($useOP), skipped connect-OPServices" ; 
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            } else {
                $smsg = "==> $($ACtionCommand): FAILED SOME SPECIFIED CONNECTIONS" ; 
                $smsg += "`n`$ret_CcOPSvcs:`n$(($ret_CcOPSvcs|out-string).trim())" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                $sdEmail.SMTPSubj = "FAIL Rpt:$($ScriptBaseName):$(get-date -format 'yyyyMMdd-HHmmtt')"
                $sdEmail.SmtpBody = "`n===Processing Summary:" ;
                if($vFailMsgs){
                    $sdEmail.SmtpBody += "`n$(($vFailMsgs|out-string).trim())" ; 
                } ; 
                $sdEmail.SmtpBody += "`n" ;
                if($SmtpAttachment){
                    $sdEmail.SmtpAttachment = $SmtpAttachment
                    $sdEmail.smtpBody +="`n(Logs Attached)" ;
                };
                $sdEmail.SmtpBody += "Pass Completed $([System.DateTime]::Now)" ;
                $smsg = "Send-EmailNotif w`n$(($sdEmail|out-string).trim())" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                Send-EmailNotif @sdEmail ;
                throw $smsg ; 
                BREAK ; 
            } ; 
            #endregion CONFIRM_CCOPRETURN ; #*------^ END CONFIRM_CCOPRETURN ^------
            #region CONFIRM_OPFORESTWIDE ; #*------v CONFIRM_OPFORESTWIDE v------    
            if($useOP -AND $pltCcOPSvcs.useForestWide -AND $ret_CcOPSvcs.hasForestWide -AND $ret_CcOPSvcs.AdGcFwide){
                $smsg = "==> $($ACtionCommand): confirmed has BOTH .hasForestWide & .AdGcFwide ($($ret_CcOPSvcs.AdGcFwide))" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Success } 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success        
            }elseif($pltCcOPSvcs.useForestWide -AND (get-variable ret_CcOPSvcs) -AND (gv -name "$($tenorg)meta").value.o365_opdomain.split('.')[0].toupper() -ne $env:userdomain){
                $smsg = "`nCROSS-ORG ONPREM CONNECTION: ATTEMPTING TO CONNECT TO ONPREM '$((gv -name "$($tenorg)meta").value.o365_Prefix)' $((gv -name "$($tenorg)meta").value.o365_opdomain.split('.')[0].toupper()) domain, FROM $($env:userdomain)!" ;
                $smsg += "`nEXPECTED ERROR, SKIPPING ONPREM FORESTWIDE SPEC" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                $useOP = $false ; 
            }elseif($useOP -AND $pltCcOPSvcs.useForestWide -AND -NOT $ret_CcOPSvcs.hasForestWide){
                $smsg = "==> $($ACtionCommand): MISSING CRITICAL FORESTWIDE SUPPORT COMPONENT:" ; 
                if(-not $ret_CcOPSvcs.hasForestWide){
                    $smsg += "`n----->$($ACtionCommand): MISSING .hasForestWide (Set-AdServerSettings -ViewEntireForest `$True) " ; 
                } ; 
                if(-not $ret_CcOPSvcs.AdGcFwide){
                    $smsg += "`n----->$($ACtionCommand): MISSING .AdGcFwide GC!:`n((Get-ADDomainController -Discover -Service GlobalCatalog).hostname):326) " ; 
                } ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                $smsg = "MISSING SOME KEY CONNECTIONS. DO YOU WANT TO IGNORE, AND CONTINUE WITH CONNECTED SERVICES?" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                $bRet=Read-Host "Enter YYY to continue. Anything else will exit"  ;
                if ($bRet.ToUpper() -eq "YYY") {
                    $smsg = "(Moving on), WITH THE FOLLOW PARTIAL CONNECTION STATUS" ;
                    $smsg += "`n`n$(($ret_CcOPSvcs|out-string).trim())" ; 
                    write-host -foregroundcolor green $smsg  ;
                } else {
                    throw $smsg ; 
                    break ; #exit 1
                } ;         
            }; 
            #endregion CONFIRM_OPFORESTWIDE ; #*------^ END CONFIRM_OPFORESTWIDE ^------
        } ; 
        #endregion CALL_CONNECT_OPSERVICES ; #*======^ END CALL_CONNECT_OPSERVICES ^======
    
        #endregion SERVICE_CONNECTIONS #*======^ END SERVICE_CONNECTIONS ^======
        
        #region IS_PIPELINE ; #*------v IS_PIPELINE v------
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
        #endregion IS_PIPELINE ; #*------^ END IS_PIPELINE ^------
    }#  # BEG-E 
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
            # $pltSL.Tag = $((@($ticket,$usr) |?{$_}) -join '-')
            #region START_LOG_OPTIONS #*======v START_LOG_OPTIONS v======
            $useSLogHOl = $true ; # one or 
            $useTransPath = $false ; # TRANSCRIPTPATH
            $useTransRotate = $false ; # TRANSCRIPTPATHROTATE
            $useStartTrans = $false ; # STARTTRANS
            $useTransNoDep = $false ; # TRANSCRIPT_NODEP
            $useTransBasicScript = $false ; # BASIC_SCRIPT_TRANSCRIPT
            #region START_LOG_HOLISTIC #*------v START_LOG_HOLISTIC v------
            if($useSLogHOl){
                # Single log for script/function example that accomodates detect/redirect from AllUsers scope'd installed code, and hunts a series of drive letters to find an alternate logging dir (defers to profile variables)
                #${CmdletName} = $rPSCmdlet.MyInvocation.MyCommand.Name ;
                if(-not (get-variable LogPathDrives -ea 0)){$LogPathDrives = 'd','c' };
                foreach($budrv in $LogPathDrives){if(test-path -path "$($budrv):\scripts" -ea 0 ){break} } ;
                if(-not (get-variable rgxPSAllUsersScope -ea 0)){$rgxPSAllUsersScope="^$([regex]::escape([environment]::getfolderpath('ProgramFiles')))\\((Windows)*)PowerShell\\(Scripts|Modules)\\.*\.(ps(((d|m))*)1|dll)$" ;} ;
                if(-not (get-variable rgxPSCurrUserScope -ea 0)){$rgxPSCurrUserScope="^$([regex]::escape([Environment]::GetFolderPath('MyDocuments')))\\((Windows)*)PowerShell\\(Scripts|Modules)\\.*\.(ps((d|m)*)1|dll)$" ;} ;
                $pltSL=[ordered]@{Path=$null ;NoTimeStamp=$false ;Tag=$null ;showdebug=$($showdebug) ; Verbose=$($VerbosePreference -eq 'Continue') ;} ;
                if($whatif.ispresent){$pltSL.add('whatif',$($whatif))}
                elseif($WhatIfPreference.ispresent ){$pltSL.add('whatif',$WhatIfPreferenc)} ;         
                # if using [CmdletBinding(SupportsShouldProcess)] + -WhatIf:$($WhatIfPreference):
                #$pltSL=[ordered]@{Path=$null ;NoTimeStamp=$false ;Tag=$null ;showdebug=$($showdebug) ; Verbose=$($VerbosePreference -eq 'Continue') ; whatif=$($WhatIfPreference) ;} ;
                #$pltSL=[ordered]@{Path=$null ;NoTimeStamp=$false ;Tag="$($ticket)-$($TenOrg)-LASTPASS-" ;showdebug=$($showdebug) ; Verbose=$($VerbosePreference -eq 'Continue') ; whatif=$($WhatIfPreference) ;} ;
                #$pltSL.Tag = $((@($ticket,$usr) |?{$_}) -join '-')
                #if($ticket){$pltSL.Tag = $ticket} ;
                #$pltSL.Tag = $env:COMPUTERNAME ; 
                $pltSL.Tag = $((@($ticket,$usr) |?{$_}) -join '-')
                $tagfields = 'ticket','UserPrincipalName','folderscope' ; # DomainName TenOrg ModuleName 
                $tagfields | foreach-object{$fld = $_ ; if(get-variable $fld -ea 0 |?{$_.value} ){$pltSL.Tag += @($((get-variable $fld).value))} } ; 
                if($pltSL.Tag -is [array]){$pltSL.Tag = $pltSL.Tag -join '-' } ; 
                #$transcript = ".\logs\$($Ticket)-$($DomainName)-$(split-path $rMyInvocation.InvocationName -leaf)-$(get-date -format 'yyyyMMdd-HHmmtt')-trans-log.txt" ; 
                #$pltSL.Tag += "-$($DomainName)"
                #
                if($rPSBoundParameters.keys){ # alt: leverage $rPSBoundParameters hash
                    $sTag = @() ; 
                    #$pltSL.TAG = $((@($rPSBoundParameters.keys) |?{$_}) -join ','); # join all params
                    if($rPSBoundParameters['Summary']){ $sTag+= @('Summary') } ; # build elements conditionally, string
                    if($rPSBoundParameters['Number']){ $sTag+= @("Number$($rPSBoundParameters['Number'])") } ; # and keyname,value
                    $pltSL.Tag += "-$($sTag -join ',')" ; # 4:46 PM 7/16/2025 flipped to append, not assign
                } ; 
                #
                if($rvEnv.isScript){
                    write-host "`$script:PSCommandPath:$($script:PSCommandPath)" ;
                    write-host "`$PSCommandPath:$($PSCommandPath)" ;
                    if($rvEnv.PSCommandPathproxy){ $prxPath = $rvEnv.PSCommandPathproxy }
                    elseif($script:PSCommandPath){$prxPath = $script:PSCommandPath}
                    elseif($rPSCommandPath){$prxPath = $rPSCommandPath} ; 
                } ; 
                if($rvEnv.isFunc){
                    if($rvEnv.FuncDir -AND $rvEnv.FuncName){
                            $prxPath = join-path -path $rvEnv.FuncDir -ChildPath $rvEnv.FuncName ; 
                    } else {
                        write-warning "Missing either `$rvEnv.FuncDir -OR `$rvEnv.FuncName!" ; 
                    } ; 
                } ; 
                if(-not $rvEnv.isFunc){
                    # under funcs, this is the scriptblock of the func, not a path
                    if($rvEnv.MyInvocationproxy.MyCommand.Definition){$prxPath2 = $rvEnv.MyInvocationproxy.MyCommand.Definition }
                    elseif($rvEnv.MyInvocationproxy.MyCommand.Definition){$prxPath2 = $rvEnv.MyInvocationproxy.MyCommand.Definition } ; 
                } ; 
                if($prxPath){
                    # 12/12/2025 new code to patch no-ext $prxPath
                    if(-not [System.IO.Path]::GetExtension($prxPath)){
                        write-verbose "no-extension `$prxpath, asserting fake ext (.ps1|.psm1 as approp)" ;                         
                        switch($rvEnv.runSource){
                            'Function'{$prxPath = "$($prxPath).psm1" }
                            'ExternalScript'{$prxPath = "$($prxPath).ps1" }
                            default {
                                $smsg = "NO RECOGNIZED `$rvEnv.runSource: '$($rvEnv.runSource)'`nUNABLE TO SAFELY TEST FOR AllUsers or CU SCOPE!: ABORTING (Could log into module hosting dir!)" ; 
                                write-warning $smsg ; throw $smsg ; 
                                BREAK ; 
                            }
                        } ; 
                    } ; 
                    if(($prxPath -match $rgxPSAllUsersScope) -OR ($prxPath -match $rgxPSCurrUserScope)){
                        $bDivertLog = $true ; 
                        switch -regex ($prxPath){
                            $rgxPSAllUsersScope{$smsg = "AllUsers"} 
                            $rgxPSCurrUserScope{$smsg = "CurrentUser"}
                        } ;
                        $smsg += " context script/module, divert logging into [$budrv]:\scripts" 
                        write-verbose $smsg  ;
                        if($bDivertLog){
                            if((split-path $prxPath -leaf) -ne $rvEnv.CmdletName){
                                # function in a module/script installed to allusers|cu - defer name to Cmdlet/Function name
                                $pltSL.Path = (join-path -Path "$($budrv):\scripts" -ChildPath "$($rvEnv.CmdletName).ps1") ;
                            } else {
                                # installed allusers|CU script, use the hosting script name
                                $pltSL.Path = (join-path -Path "$($budrv):\scripts" -ChildPath (split-path $prxPath -leaf)) ;
                            }
                        } ;
                    } else {
                        $pltSL.Path = $prxPath ;
                    } ;
                }elseif($prxPath2){
                    # 12/12/2025 new code to patch no-ext $prxPath2
                    if(-not [System.IO.Path]::GetExtension($prxPath2)){
                        write-verbose "no-extension `$prxPath2, asserting fake ext (.ps1|.psm1 as approp)" ;                         
                        switch($rvEnv.runSource){
                            'Function'{$prxPath2 = "$($prxPath2).psm1" }
                            'ExternalScript'{$prxPath2 = "$($prxPath2).ps1" }
                            default {
                                $smsg = "NO RECOGNIZED `$rvEnv.runSource: '$($rvEnv.runSource)'`nUNABLE TO SAFELY TEST FOR AllUsers or CU SCOPE!: ABORTING (Could log into module hosting dir!)" ; 
                                write-warning $smsg ; throw $smsg ; 
                                BREAK ; 
                            }
                        } ; 
                    } ; 
                    if(($prxPath2 -match $rgxPSAllUsersScope) -OR ($prxPath2 -match $rgxPSCurrUserScope) ){
                            $pltSL.Path = (join-path -Path "$($budrv):\scripts" -ChildPath (split-path $prxPath2 -leaf)) ;
                    } elseif(test-path $prxPath2) {
                        $pltSL.Path = $prxPath2 ;
                    } elseif($rvEnv.CmdletName){
                        $pltSL.Path = (join-path -Path "$($budrv):\scripts" -ChildPath "$($rvEnv.CmdletName).ps1") ;
                    } else {
                        $smsg = "UNABLE TO RESOLVE A FUNCTIONAL `$rvEnv.CmdletName, FROM WHICH TO BUILD A START-LOG.PATH!" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Warn } #Error|Warn|Debug 
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        BREAK ;
                    } ; 
                } else{
                    $smsg = "UNABLE TO RESOLVE A FUNCTIONAL `$rvEnv.CmdletName, FROM WHICH TO BUILD A START-LOG.PATH!" ; 
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
                        $startResults = start-Transcript -path $transcript -whatif:$false -confirm:$false;
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

                    #region SendMailAlert ; #*------v SendMailAlert v------
                    $SmtpBody += "`n===FAIL Summary:" ;
                    $SmtpBody += "`n$('-'*50)" ;
                    $SmtpBody += "`n$('-'*50)" ;
                    $smsg += "`n$(($smsg |out-string).trim())" ; 
                    $sdEmail = @{
                        smtpFrom = $SMTPFrom ;
                        SMTPTo = $SMTPTo ;
                        SMTPSubj = $SMTPSubj ;
                        #SMTPServer = $SMTPServer ;
                        SmtpBody = $SmtpBody ;
                        SmtpAttachment = $SmtpAttachment ;
                        BodyAsHtml = $false ; # let the htmltag rgx in Send-EmailNotif flip on as needed
                        verbose = $($VerbosePreference -eq "Continue") ;
                    } ;
                    $smsg = "Send-EmailNotif w`n$(($sdEmail|out-string).trim())" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    Send-EmailNotif @sdEmail ;

                    #endregion SendMailAlert ; #*------^ END SendMailAlert ^------
                } ;
            } ; 
            #endregion START_LOG_HOLISTIC #*------^ END START_LOG_HOLISTIC ^------
            # ...
            #endregion START_LOG_OPTIONS #*======^ START_LOG_OPTIONS ^======

            #region VARI_SETUP ; #*------v VARI_SETUP v------
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
                #AADUser = $null ; # added for MailUser variant
                MGUser = $null ; # 12:43 PM 12/10/2025 added with M$ forced block of all azuread access, and forced migration to crappy MG module calls
                #AADUserMgr = $null ;
                MGUserMgr = $null ; # 12:43 PM 12/10/2025 added with M$ forced block of all azuread access, and forced migration to crappy MG module calls
                #AADUserLics = $null ; 
                MGUserLics = $null ; # 12:43 PM 12/10/2025 added with M$ forced block of all azuread access, and forced migration to crappy MG module calls
                LicenseGroup = $null ;
                isDirSynced = $null 
                isNoBrain = $false ;
                isSplitBrain = $false;
                #isUnlicensed = $false ;
                IsLicensed = $false ; 
                IsDisabledOU = $false ; 
                IsADDisabled = $false ; 
                #IsAADDisabled = $false ; 
                IsMGDisabled = $false ; 
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
                $hsum.add('xoMobileDeviceStatsOLD',$null) ; 
                # 9:22 AM 9/23/2025 add xoMobileOutlookClients, xoMobileOMSyncTypes, reflecting supported Outlook Mobile client & the ClientType spec in use for the OLM 'Microsoft's native sync technology'
                # add xoMobileDeviceTypes, xoMobileOtherSyncTypes to make iphone/android types immed vis
                $hsum.add('xoMobileOutlookClients',$null) ; 
                $hsum.add('xoMobileOtherClients',$null) ; 
                $hsum.add('xoMobileOMSyncTypes',$null) ; 
                #$hsum.add('xoMobileDeviceTypes',$null) ; 
                $hsum.add('xoMobileOtherSyncTypes',$null) ; 
                
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
                [string]$ofMbxFolderStats = $logfile.replace('LOG-BATCH-EXEC',"folder-sizes-NONHIDDEN-NONZERO").replace('-log','') ; 
                #$ofile.replace('REPORT',"folder-sizes-NONHIDDEN-NONZERO") ; 

            } ; 
            # 2:35 PM 12/26/2024 getPerms
            if($getPerms){
                $hsum.add('xoMailboxPermission',$null) ; 
                $hsum.add('xoRecipientPermission',$null) ; 
                #$hsum.add('xoMailboxPermissionGroupManagedBy',$null) ; # moved into the group summary
                $hsum.add('xoMailboxPermissionGroups',@()) ; 
                $hsum.add('xoRecipientPermissionGroups',$null) ; 
                #$hsum.add('xoRecipientPermissionGroupManagedBy',@()) ; 
            }
            # 2:44 PM 4/12/2025 add ResolveForwards Mailcontact/ForwardingAddress resolution
            if($ResolveForwards){
                $hsum.add('opMailContact',$null) ;
                $hsum.add('opContactForwards',$null) ; 
                $hsum.add('xoMailContact',$null) ;
                $hsum.add('xoMailboxForwardingAddress',$null) ; 
                $hsum.add('xoContactForwards',$null) ; 
            }
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
            #endregion VARI_SETUP ; #*------^ END VARI_SETUP ^------

            $sBnr="===v ($($Procd)/$($ttl)):Input: '$($usr)' | '$($hSum.fname)' | '$($hSum.lname)' v===" ;
            if($isEml){$sBnr+="(EML)"}
            elseif($isDname){$sBnr+="(DNAM)"}
            elseif($isObjName){$sBnr+="(ONAM)"}
            elseif($isSamAcct){$sBnr+="(SAM)"}
            write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($sBnr)" ;

            write-host -foreground yellow "get-Rmbx/xMbx: " -nonewline;

            #region SPLAT_SETUP ; #*------v SPLAT_SETUP v------
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
            #endregion SPLAT_SETUP ; #*------^ END SPLAT_SETUP ^------

            $error.clear() ;

            #write-verbose "get-[exo]Recipient w`n$(($pltGMailObj|out-string).trim())" ;
            #write-verbose "get-recipient w`n$(($pltGMailObj|out-string).trim())" ;
            # exclude contacts, they don't represent real onprem mbx assoc, and we need to refer those to EXO mbx qry anyway.
            write-verbose "get-recipient w`n$(($pltGMailObj|out-string).trim())" ;
            #rx10 -Verbose:$false -silent ;

            #region OPRCP_DISCOVERY ; #*------v OPRCP_DISCOVERY v------
            if($resolveForwards){
                $smsg = "-resolveForwards: (include MailContacts)`nget-recipient w`n$(($pltGMailObj|out-string).trim())`n..." ; 
            }else{
                $smsg = "get-recipient w`n$(($pltGMailObj|out-string).trim())`n...| ?{$_.recipienttypedetails -ne 'MailContact'}" ; 
            } ; 
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            
            #if($hSum.OPRcp=get-recipient @pltGMailObj -ea 0 | select -first $MaxRecips | ?{$_.recipienttypedetails -ne 'MailContact'}){
            if($hSum.OPRcp=get-recipient @pltGMailObj -ea 0 | select -first $MaxRecips ){
                if($resolveForwards){
                    
                } else { 
                    $hSum.OPRcp | ?{$_.recipienttypedetails -ne 'MailContact'} ; 
                } ; 
                $smsg = "`$hSum.OPRcp found as `n$(($pltGMailObj.GetEnumerator() | ?{$_.key -ne 'ResultSize'}  | ft -a key,value|out-string).trim())" ; 
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                
            } elseif($isDname -and $hsum.lname) {
                # put in missing *, hits on mis-spellings 'Spark' matches 'Sparks' w wildcard
                if($hsum.lname -match "[']"){
                    if(-not $resolveForwards){
                        $fltr = "recipienttypedetails -ne " + $sQot + "MailContact" + $sQot ; 
                        $fltr += " -AND displayname -like " + $sQot + $($hsum.lname) + $sQot ;
                    } else {
                        $fltr = "displayname -like " + $sQot + $($hsum.lname) + $sQot ;
                    };
                    
                }else {
                    if(-not $resolveForwards){
                        $fltr = "recipienttypedetails -ne 'MailContact'" ; 
                        $fltr += " -AND displayname -like '$($hsum.lname)'" ;
                    } else { 
                        $fltr += "displayname -like '$($hsum.lname)'" ;
                    } 
                } ; 
                if($hsum.fname){
                    # try first 3 of fname first
                    if($hsum.fname -match "[']"){
                        $fltr += " -AND firstName -like " + $sQot + $($hsum.fname.substring(0,3)) + "*" + $sQot ; 
                    }else {
                        $fltr += " -AND firstName -like '$($hsum.fname.substring(0,3))*'" ; 
                    } ; 
                    
                    #if($hSum.OPRcp=get-recipient -filter $fltr -ea 0 | select -first $MaxRecips |?{$_.recipienttypedetails -ne 'MailContact'}){
                    if($hSum.OPRcp=get-recipient -filter $fltr -ea 0 | select -first $MaxRecips){
                        if($resolveForwards){
                            
                        } else { 
                            $hSum.OPRcp = $hSum.OPRcp |?{$_.recipienttypedetails -ne 'MailContact'}
                        } ;
                        $smsg = "`$hSum.OPRcp found as `n$(($pltGMailObj.GetEnumerator() | ?{$_.key -ne 'ResultSize'}  | ft -a key,value|out-string).trim())" ; 
                        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                    }elseif($hsum.fname){
                        # retry first initial
                        if($hsum.lname -match "[']"){
                            if($resolveForwards){
                                $fltr = "lastName -like " + $sQot + $($hsum.lname) + "*" + $sQot ; 
                            }else {
                                $fltr = "recipienttypedetails -ne " + $sQot + "MailContact" + $sQot + " -AND lastName -like " + $sQot + $($hsum.lname) + "*" + $sQot ; 
                            };
                        }else {
                            if($resolveForwards){
                                $fltr = "lastName -like '$($hsum.lname)*'" ; 
                            }else {
                                $fltr = "recipienttypedetails -ne 'MailContact' -AND lastName -like '$($hsum.lname)*'" ; 
                            } ;
                        } ; 
                        if($hsum.fname -match "[']"){
                                $fltr += " -AND firstName -like " + $sQot + $($hsum.fname.substring(0,1)) + "*" + $sQot ; 
                        }else {
                            $fltr += " -AND firstName -like '$($hsum.fname.substring(0,1))*'" ; 
                        } ; 
                        
                        #if($hSum.OPRcp=get-recipient -filter $fltr -ea 0 | select -first $MaxRecips |?{$_.recipienttypedetails -ne 'MailContact'}){
                        if($hSum.OPRcp=get-recipient -filter $fltr -ea 0 | select -first $MaxRecips){
                            if($resolveForwards){}else {
                                $hSum.OPRcp=$hSum.OPRcp  |?{$_.recipienttypedetails -ne 'MailContact'} ; 
                            }
                            $smsg = "`$hSum.OPRcp found as `n$(($pltGMailObj.GetEnumerator() | ?{$_.key -ne 'ResultSize'}  | ft -a key,value|out-string).trim())" ; 
                            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                        }elseif($hsum.lname){
                            # do wildcard lname matches
                            if($hsum.lname -match "[']"){
                                if($resolveForwards){
                                    $fltr = "lastName -like " + $sQot + $($hsum.lname) + "*" + $sQot ; 
                                }else {
                                    $fltr = "recipienttypedetails -ne " + $sQot + "MailContact" + $sQot + " -AND lastName -like " + $sQot + $($hsum.lname) + "*" + $sQot ; 
                                }
                            }else {
                                if($resolveForwards){
                                    $fltr = "lastName -like '$($hsum.lname)*'" ; 
                                }else{
                                    $fltr = "recipienttypedetails -ne 'MailContact' -AND lastName -like '$($hsum.lname)*'" ; 
                                }
                            } ; 
                            
                            #if($hSum.OPRcp=get-recipient -filter $fltr -ea 0 | select -first $MaxRecips |?{$_.recipienttypedetails -ne 'MailContact'}){
                            if($hSum.OPRcp=get-recipient -filter $fltr -ea 0 | select -first $MaxRecips){
                                if($resolveForwards){

                                }else{
                                    $hSum.OPRcp=$hSum.OPRcp |?{$_.recipienttypedetails -ne 'MailContact'} ; 
                                }
                                $smsg = "`$hSum.OPRcp found as `n$(($pltGMailObj.GetEnumerator() | ?{$_.key -ne 'ResultSize'}  | ft -a key,value|out-string).trim())" ; 
                                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                            }else{
                                $smsg = "(Failed to OP:get-recipient on:$($usr))"
                                if($isDname){$smsg += " or *$($hsum.lname )*"}
                                write-host $smsg ;                            
                            } ;
                        }
                    } ; 
   
                } ; 
            } ; 
            #endregion OPRCP_DISCOVERY ; #*------^ END OPRCP_DISCOVERY ^------

            #region ECHO_OPRCP ; #*------v ECHO_OPRCP v------
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
                        'RemoteUserMailbox' {write-host "(Rmbx)"}
                        # 8:53 AM 10/9/2024 add to cover mbx2shared conversion results
                        'RemoteSharedMailbox' {write-host "(Rmbx *SHARED*)"} 
                        # 12:23 PM 12/26/2024 add resource & remote res's
                        'RemoteRoomMailbox' {write-host "(Rmbx *ROOM*)"} 
                        'RemoteEquipmentMailbox' {write-host "(Rmbx *EQUIP*)"} 
                        'UserMailbox' {write-host "(Mbx)"}
                        'SharedMailbox' {write-host "(SMbx)"}
                        'RoomMailbox' {write-host "(RoomMbx)"}
                        'EquipmentMailbox' {write-host "(EquipMbx)"}
                        'MailUser' {
                            $smsg = "MAILUSER WO RMBX DETECTED! - POSSIBLE NOBRAIN?"
                            write-warning $smsg
                            #$hsum.isNoBrain = $true ;    
                        }
                        'MailUniversalDistributionGroup' {write-host "(DG)"}
                        'DynamicDistributionGroup'  {write-host "(DDG)"}
                        'MailContact' {write-host "(MC)"]}
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
            #endregion ECHO_OPRCP ; #*------^ END ECHO_OPRCP ^------

            #region XORCP_DISCOVERY ; #*------v XORCP_DISCOVERY v------
            #if ($useEXOv2) { reconnect-eXO2 @pltRXOC }
            #else { reconnect-EXO @pltRXOC } ;
            #write-host -foreground yellow "get-xoMbx/xMbx: " -nonewline;
            if($resolveForwards){
                $smsg = "get-xorecipient w`n$(($pltGMailObj|out-string).trim())`n..." ;
            } else { 
                $smsg = "get-xorecipient w`n$(($pltGMailObj|out-string).trim())`n...| ?{$_.recipienttypedetails -ne 'MailContact'}" ;                
            }
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 

            #if($hSum.xoRcp=get-xorecipient @pltGMailObj -ea 0 | select -first $MaxRecips | ?{$_.recipienttypedetails -ne 'MailContact'}){
            if($hSum.xoRcp=get-xorecipient @pltGMailObj -ea 0 | select -first $MaxRecips){
                if($resolveForwards){

                }else {
                    $hSum.xoRcp=$hSum.xoRcp  | ?{$_.recipienttypedetails -ne 'MailContact'}
                }
                $smsg = "`$hSum.xoRcp found as `n$(($pltGMailObj.GetEnumerator() | ?{$_.key -ne 'ResultSize'}  | ft -a key,value|out-string).trim())" ; 
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                
            } elseif($isDname -and $hsum.lname) {
                
                # put in missing *, hits on mis-spellings 'Spark' matches 'Sparks' w wildcard
                if($hsum.lname -match "[']"){
                    if($resolveForwards){
                        $fltr = "lastName -like " + $sQot + $($hsum.lname) + "*" + $sQot ;
                    }else {
                        $fltr = "recipienttypedetails -ne " + $sQot + "MailContact" + $sQot ;
                        $fltr += " -AND lastName -like " + $sQot + $($hsum.lname) + "*" + $sQot ;
                    }
                }else{
                    if($resolveForwards){
                        $fltr += "lastName -like '$($hsum.lname)*'" ;
                    }else {
                        $fltr = "recipienttypedetails -ne 'MailContact'" ;
                        $fltr += " -AND lastName -like '$($hsum.lname)*'" ;
                    }
                } ; 
                if($hsum.fname){
                    # try first 3 of fname first
                    if($hsum.fname -match "[']"){
                        $fltr += " -AND firstName -like " + $sQot + $($hsum.fname.substring(0,3)) + "*" + $sQot ;
                    }else{
                        $fltr += " -AND firstName -like '$($hsum.fname.substring(0,3))*'" ;
                    } ; 
                    #if($hSum.xoRcp=get-xorecipient -filter $fltr -ea 0 | select -first $MaxRecips |?{$_.recipienttypedetails -ne 'MailContact'}){
                    if($hSum.xoRcp=get-xorecipient -filter $fltr -ea 0 | select -first $MaxRecips){
                        if($resolveForwards){

                        }else {
                            $hSum.xoRcp=$hSum.xoRcp  |?{$_.recipienttypedetails -ne 'MailContact'}
                        }
                        $smsg = "`$hSum.xoRcp found as `n$(($pltGMailObj.GetEnumerator() | ?{$_.key -ne 'ResultSize'}  | ft -a key,value|out-string).trim())" ; 
                        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                    }elseif($hsum.fname){
                        # retry first initial
                        if($hsum.lname -match "[']"){
                            if($resolveForwards){
                                $fltr = "lastName -like " + $sQot + $($hsum.lname) + "*" + $sQot ;
                            }else {
                                $fltr = "recipienttypedetails -ne " + $sQot + "MailContact" + $sQot + " -AND lastName -like " + $sQot + $($hsum.lname) + "*" + $sQot ;
                            }
                        } else { 
                            if($resolveForwards){
                                $fltr = "lastName -like '$($hsum.lname)*'" ;
                            }else{
                                $fltr = "recipienttypedetails -ne 'MailContact' -AND lastName -like '$($hsum.lname)*'" ;
                            }
                        }
                        if($hsum.fname -match "[']"){
                            $fltr += " -AND firstName -like " + $sQot + $($hsum.fname.substring(0,1)) + "*" + $sQot ;
                        } else { 
                            $fltr += " -AND firstName -like '$($hsum.fname.substring(0,1))*'" ;
                        } ; 

                        #if($hSum.xoRcp=get-xorecipient -filter $fltr -ea 0 | select -first $MaxRecips |?{$_.recipienttypedetails -ne 'MailContact'}){
                        if($hSum.xoRcp=get-xorecipient -filter $fltr -ea 0 | select -first $MaxRecips ){
                            if($resolveForwards){

                            }else {
                                $hSum.xoRcp=$hSum.xoRcp |?{$_.recipienttypedetails -ne 'MailContact'} ; 
                            } ; 
                            $smsg = "`$hSum.xoRcp found as `n$(($pltGMailObj.GetEnumerator() | ?{$_.key -ne 'ResultSize'}  | ft -a key,value|out-string).trim())" ; 
                            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                        }elseif($hsum.lname){
                            # do wildcard lname matches
                            if($hsum.fname -match "[']"){
                                if($resolveForwards){
                                    $fltr = "lastName -like " + $sQot + $($hsum.lname) + "*" + $sQot ;
                                }else {
                                    $fltr = "recipienttypedetails -ne " + $sQot + "MailContact" + $sQot + " -AND lastName -like " + $sQot + $($hsum.lname) + "*" + $sQot ;
                                }
                            } else { 
                                if($resolveForwards){
                                       $fltr = "lastName -like '$($hsum.lname)*'" ;
                                }else{
                                    $fltr = "recipienttypedetails -ne 'MailContact' -AND lastName -like '$($hsum.lname)*'" ;
                                }
                            } ; 
                            #if($hSum.xoRcp=get-xorecipient -filter $fltr -ea 0 | select -first $MaxRecips |?{$_.recipienttypedetails -ne 'MailContact'}){
                            if($hSum.xoRcp=get-xorecipient -filter $fltr -ea 0 | select -first $MaxRecips){
                                if($resolveForwards){
                                }else{
                                    $hSum.xoRcp=$hSum.xoRcp |?{$_.recipienttypedetails -ne 'MailContact'}
                                }
                                $smsg = "`$hSum.xoRcp found as `n$(($pltGMailObj.GetEnumerator() | ?{$_.key -ne 'ResultSize'}  | ft -a key,value |out-string).trim())" ; 
                                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                            }else{
                                $smsg = "(Failed to OP:get-xorecipient on:$($usr))"
                                if($isDname){$smsg += " or *$($hsum.lname )*"}
                                write-host $smsg ;
                            } ;
                        }
                    } ;
                } ;
            } ; 
            #endregion XORCP_DISCOVERY ; #*------^ END XORCP_DISCOVERY ^------

            #region ECHO_XORCP ; #*------v ECHO_XORCP v------
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
                        'MailContact' {write-host -nonewline "(xMC)" }
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
            #endregion ECHO_XORCP ; #*------^ END ECHO_XORCP ^------

            # new rules, with INT/VEN AADU anchored to ADU, but xoMbx anchored solely to AADU (and not OP rcp), it's possible to completely fail onprem get-recipient, and still have a functional mailbox in cloud, that's operating properly.

            #region NONUNIQUE_RCPS_ABORT ; #*------v NONUNIQUE_RCPS_ABORT v------
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
            #endregion NONUNIQUE_RCPS_ABORT ; #*------^ END NONUNIQUE_RCPS_ABORT ^------

            #region OP_V_XO_RCPEXPAND ; #*------v OP_V_XO_RCPEXPAND v------
            if($hSum.OPRcp){
                #region OPRCP_EXPAND ; #*------v OPRCP_EXPAND v------
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
                                        #caad  -Verbose:$false -silent ;
                                        #write-verbose "OPRcp:Mailuser, ensure GET-ADUSER pulls AADUser.matched object for cloud recipient:`nfallback:get-AzureAdUser  -objectid $($hsum.xoRcp.ExternalDirectoryObjectId)" ;
                                        write-verbose "OPRcp:Mailuser, ensure Get-MgUser pulls MgUser.matched object for cloud recipient:`nfallback:Get-MgUser -userid $($hsum.xoRcp.ExternalDirectoryObjectId)" ;
                                        # have to postfilter, if want specific count -maxresults catch's with no $error[0]
                                        #$hSum.AADUser   +=  get-AzureAdUser  -objectid $hsum.xoRcp.ExternalDirectoryObjectId | select -first $MaxRecips;  ;
                                        # sub in Get-MgUser get-MgUserFull                                       
                                        $hSum.MGUser   +=  get-MgUserFull -userid $hsum.xoRcp.ExternalDirectoryObjectId | select -first $MaxRecips;  ;
                                    } else {
                                        throw "Unsupported object, blank `$hsum.xoRcp.ExternalDirectoryObjectId!" ;
                                    } ;
                                }
                                if($outObject){

                                } else {
                                    $Rpt += $hSum.OPRemoteMailbox.primarysmtpaddress ;
                                } ;
                            } ;
                            "MailContact" {
                                #$hSum.OPRemoteMailbox += get-remotemailbox $txR.identity  ;
                                #$Rpt += $hSum.OPRemoteMailbox.primarysmtpaddress ;

                                $bufferRcp = $null ; 
                                $bufferRcp  = get-mailcontact $tmpRcp.identity -resultsize $MaxRecips -ea 0 | select -first $MaxRecips ; 
                                #if($hSum.opMailContact += get-mailcontact $tmpRcp.identity -resultsize $MaxRecips -ea 0 | select -first $MaxRecips ; ){
                                if($bufferRcp){
                                    $hSum.opMailContact += $bufferRcp ; 
                                    write-verbose "`$hSum.opMailContact:`n$(($hSum.opMailContact|ft -a |out-string).trim())" ;
                                }else{
                                    $smsg = "RecipientTypeDetails:MailContact with NO Contact!!" ;
                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                                }
                                $smsg = "$($tmpRcp.primarysmtpaddress): matches an EXO MailContact with external Email: $($bufferRcp.primarysmtpaddress)" ;
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                if($ResolveForwards){
                                    if(-not $hshForwards){
                                        $hshForwards = resolve-RMbxForwards ;
                                    } ;
                                    $tid = $bufferRcp.primarysmtpaddress ;
                                    if($hshForwards[$tid]){
                                        write-host "$($bufferRcp.primarysmtpaddress):Forwarding Contact"
                                        $smsg = "Recipient:$($tid) => $($hshForwards[$tid])" ;
                                        write-host $smsg ;
                                        $hsum.opContactForwards = $hshForwards[$tid] ;
                                    } ;
                                } ;
                                break ;
 
                            }
                            default {
                                write-warning "$((get-date).ToString('HH:mm:ss')):Unsupported RecipientType:($tmpRcp.recipienttype). EXITING!" ;
                                Break ;
                            }
                        }
                        #region OP_GADU ; #*------v OP_GADU v------
                        <# get-aduser docs say -REsultSetSize is documented,
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
                            #if($hSum.AADUser.DirSyncEnabled){
                            # forcedmigr aad -> mg: Get-MgUserManager 
                            #  $mgu.OnPremisesSyncEnabled
                             if($hSum.MGUser.OnPremisesSyncEnabled){
                                $smsg = "Falling back to AADU Immutable lookup to locate replicated adu source" ;
                                #if($pltGadu.identity = $hSum.AADUser.ImmutableId | convert-ImmuntableIDToGUID | select -expand guid){
                                # forcedmigr aad -> mg:
                                # ren'd: OnPremisesImmutableId : UC7OjGZAYUK/qsVZatOR6g==
                                if($pltGadu.identity = $hSum.MGUser.OnPremisesImmutableId | convert-ImmuntableIDToGUID | select -expand guid){
                                    #$smsg = "(Resolved AADU.Immutable ->GUID:$($pltGadu.identity))" ;
                                    $smsg = "(Resolved MGU.Immutable ->GUID:$($pltGadu.identity))" ;
                                    write-verbose $smsg ;
                                }else {
                                    $smsg = "UNABLE TO RESOLVE ADU.IMMUTABLEID TO ADU GUID!"
                                    write-warning $smsg ;
                                    throw $smsg ;
                                }
                            } else {
                                #$smsg = "$AADUsuer not DirSyncEnabled: CLOUD FIRST!"
                                $smsg = "$MGUser not OnPremisesSyncEnabled: CLOUD FIRST!"
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

                            if($hSum.OPRemoteMailbox.ForwardingAddress -OR $hSum.OPMailbox.ForwardingAddress){
                                write-host $smsg ; # echo pending, using ww below
                                $smsg = "==FORWARDED rMBX!:" ; 
                                # 10:31 AM 4/15/2025 resolve target of forward
                                $smsg += "`n$(($MailRecip|select $propsMailxL5 |out-markdowntable @MDtbl|out-string).trim())" ; 
                                if($fAddrRcp = $MailRecip.forwardingaddress| get-recipient -ea 0){
                                    $smsg += "`nFORWARDS TO OBJECT:`n$(($fAddrRcp | select name,RecipientType,PrimarySmtpAddress |out-markdowntable @MDtbl|out-string).trim())" ; 
                                } else{
                                     $smsg += "UNABLE TO RESOLVE forwardingaddress TO FUNCTIONAL RECIPIENT!(get-recipient)!" ;
                                }; 
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
                    #endregion OP_GADU ; #*------^ END OP_GADU ^------
                }  # loop-E $hSum.OPRcp
                #endregion OPRCP_EXPAND ; #*------^ END OPRCP_EXPAND ^------
            }elseif($hSum.xoRcp){
                #region XORCP_EXPAND ; #*------v XORCP_EXPAND v------
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
                                        #region xogetMobile ; #*------v xogetMobile v------
                                        if($getMobile){
                                            get-xoMobileData ; 
                                            <#
                                            $smsg = "'xoMobileDeviceStats':Get-xoMobileDeviceStatistics -Mailbox $($xmbx.ExchangeGuid.guid)"
                                            if($verbose){
                                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                                                else{ write-verbose $smsg } ; 
                                            } ; 
                                            #$hsum.xoMobileDeviceStats  +=  Get-xoMobileDeviceStatistics -Mailbox $xmbx.userprincipalname -ea STOP ; 
                                            # wasn't getting data back: shift to the .xomailbox.ExchangeGuid.guid, it's 100% going to hit and return data 
                                            $hsum.xoMobileDeviceStats  +=  Get-xoMobileDeviceStatistics -Mailbox $xmbx.ExchangeGuid.guid -ea STOP ; 
                                            $smsg = "xoMobileDeviceStats Count:$(($hsum.xoMobileDeviceStats|measure).count)" ;
                                            write-host -foregroundcolor green $smsg ;
                                            #>
                                        } ; 
                                        #endregion xogetMobile ; #*------^ END xogetMobile ^------
                                        #region xogetQuotaUsage ; #*------v getQuotaUsage v------
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
                                            $smsg = "xoMailboxStats Count:$(($hsum.xoMailboxStats|measure).count)" ;
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
                                                if($DeletedItems){
                                                    $hsum.xoMailboxFolderStats | 
                                                        select $prpFldrDeleted | sort TreeSizeMB -desc | export-csv  -path $ofMbxFolderStats -notype ;

                                                }else{
                                                    $hsum.xoMailboxFolderStats | ?{$_.ItemsInFolder -gt 0 -AND $_.identity -notmatch $rgxHiddn } | 
                                                        select $prpFldr | sort SizeMB -desc | export-csv  -path $ofMbxFolderStats -notype ;
                                                } ; 

                                            } CATCH {
                                                $ErrTrapd=$Error[0] ;
                                                $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                                                write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" ;
                                            } ; 
                                        } ; 
                                        #endregion xogetQuotaUsage ; #*------^ END xogetQuotaUsage ^------
                                        #region xogetPerms ; #*------v xogetPerms v------
                                        if($getPerms){
                                            $pltGMbxPermX=[ordered]@{
                                                identity = $hSum.xoMailbox.exchangeguid ;
                                                ErrorAction = 'STOP' ;
                                            } ;
                                            $smsg = "Get-xoMailboxPermission  w`n$(($pltGMbxPermX|out-string).trim())"
                                            if($verbose){
                                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                                                else{ write-verbose $smsg } ;
                                            } ;
                                            TRY{
                                                $hSum.xoMailboxPermission  +=  Get-xoMailboxPermission @pltGMbxPermX | ?{$_.user -notmatch 'NT\sAUTHORITY\\SELF'} | select $prpMPerms;
                                                $smsg = "xoMailboxPermission Count:$(($hsum.xoMailboxPermission|measure).count)" ;
                                                write-host -foregroundcolor green $smsg ;
                                                if($hSum.xoMailboxPermission){
                                                    foreach($grp in ($hSum.xoMailboxPermission.user | 
                                                        get-xorecipient  | ?{$_.recipienttype -eq 'MailUniversalSecurityGroup'}) ){
                                                        $hshGrpSumm = [ordered]@{
                                                            Identity = $grp.Identity
                                                            PrimarySmtpAddress = $grp.PrimarySmtpAddress ;
                                                            Description = $grp.Description ;
                                                            RecipientType = $grp.RecipientType ;
                                                            RecipientTypeDetails = $grp.RecipientTypeDetails ;
                                                            ManagedBy = ($grp | get-xodistributiongroup | select -expand managedby | get-xorecipient -ea 0) ;
                                                            Members = ($grp | get-xodistributiongroupmember | get-xorecipient  -ea 0) ;
                                                        } ; 
                                                        $hSum.xoMailboxPermissionGroups += [pscustomobject]$hshGrpSumm ; 
                                                    } ;
                                                } else {
                                                    $smsg = "(no non-SELF Get-xoMailboxPermission returned)" ; 
                                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                                                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                                    #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                                                };
                                            } CATCH {
                                                $ErrTrapd=$Error[0] ;
                                                $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                                                write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" ;
                                            } ;
                                            $smsg = "Get-xoRecipientPermission  w`n$(($pltGMbxPermX|out-string).trim())"
                                            if($verbose){
                                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                                                else{ write-verbose $smsg } ;
                                            } ;
                                            TRY{
                                                $hsum.xoRecipientPermission += Get-xoRecipientPermission @pltGMbxPermX | ?{$_.trustee -notmatch 'NT\sAUTHORITY\\SELF'}  | select $prpRPerms;
                                                $smsg = "xoRecipientPermission Count:$(($hsum.xoRecipientPermission|measure).count)" ;
                                                write-host -foregroundcolor green $smsg ;
                                                if($hsum.xoRecipientPermission){
                                                    foreach($grp in ($hsum.xoRecipientPermission.trustee | 
                                                        get-xorecipient  | ?{$_.recipienttype -eq 'MailUniversalSecurityGroup'}) ){
                                                        $hshGrpSumm = [ordered]@{
                                                            Identity = $grp.Identity
                                                            PrimarySmtpAddress = $grp.PrimarySmtpAddress ;
                                                            Description = $grp.Description ;
                                                            RecipientType = $grp.RecipientType ;
                                                            RecipientTypeDetails = $grp.RecipientTypeDetails ;
                                                            ManagedBy = ($grp | get-xodistributiongroup | select -expand managedby | get-xorecipient -ea 0) ;
                                                            Members = ($grp | get-xodistributiongroupmember | get-xorecipient  -ea 0) ;
                                                        } ; 
                                                        $hSum.xoRecipientPermissionGroups += [pscustomobject]$hshGrpSumm ;
                                                    } ;
                                                } else {
                                                    $smsg = "(no non-SELF Get-xoRecipientPermission returned)" ; 
                                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                                                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                                    #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                                                };
                                            } CATCH {
                                                $ErrTrapd=$Error[0] ;
                                                $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                                                write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" ;
                                            } ;
                                        } ; 
                                        #endregion xogetPerms ; #*------^ END xogetPerms ^------
                                        #region ResolveForwards ; #*------v ResolveForwards v------
                                        # we don't need the hash-Rmbx lookup process, just expand the fwd address to matching recip
                                        if($hSum.xoMailbox.ForwardingAddress){
                                            $smsg = "NOTE:$($hSum.xoMailbox.userprincipalname) has *populated* ForwardingAddress!:" ; 
                                            $smsg += "`nForwardingAddress`n$(($hSum.xoMailbox.ForwardingAddress|out-string).trim())" ;
                                            if($fAddrRcp = $hSum.xoMailbox.ForwardingAddress | get-xorecipient -ea 0){
                                                $smsg += "`n=> which forwards into object`n$(($faddrrcp | ft -a name,RecipientType,PrimarySmtpAddress|out-string).trim())" ;
                                            } else { 
                                                $smsg += "==> UNABLE TO RESOLVE THE ABOVE OBJECT INTO GET-XORECIPIENT (NO RETURN)!" ; 
                                            } ; 
                                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                                            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                                        };
                                        #endregion ResolveForwards ; #*------^ END ResolveForwards ^------
                                    } ; # foreach($xmbx in $hSum.xoMailbox)
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
                                #$hSum.txGuest  +=  get-AzureAdUser  -objectid $hSum.xoUser.userPrincipalName -Top $MaxRecips | select -first $MaxRecips ;
                                # forced upgrade blocked AAD -> MG
                                $hSum.txGuest  +=  get-MgUserFull -userid $hsum.xoRcp.ExternalDirectoryObjectId | select -first $MaxRecips;  ;
                                write-verbose "`$hSum.txGuest:`n$(($hSum.txGuest|ft -a |out-string).trim())" ;
                                #$Rpt += $hSum.OPRemoteMailbox.primarysmtpaddress ;
                                write-host "$($txR.ExternalEmailAddress): matches a Guest object with UPN:$($hSum.xoUser.userPrincipalName)" ;
                                if($null -eq $hSum.txGuest.EmailAddresses){
                                    write-warning "Guest appears to have damage from conficting replicated onprem MailContact, as it's EmailAddresses property is *blank*" ;
                                } ;
                                break ;
                            } ;
                            "MailContact" {
                                $bufferRcp = $null ;
                                $bufferRcp  = get-xomailcontact $txR.identity -resultsize $MaxRecips -ea 0 | select -first $MaxRecips ;
                                if($bufferRcp){
                                    $hSum.xoMailContact += $bufferRcp ;
                                    write-verbose "`$hSum.opMailContact:`n$(($hSum.opMailContact|ft -a |out-string).trim())" ;
                                }else{
                                    $smsg = "RecipientTypeDetails:MailContact with NO Contact!!" ;
                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                }
                                $smsg = "$($txR.primarysmtpaddress): matches an EXO MailContact with external Email: $($bufferRcp.externalemailaddress)" ;
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                if($ResolveForwards){
                                    if(-not $hshForwards){
                                        $hshForwards = resolve-RMbxForwards ;
                                    } ;
                                    $tid = $bufferRcp.primarysmtpaddress ;
                                    if($hshForwards[$tid]){
                                        write-host "$($bufferRcp.primarysmtpaddress):Forwarding Contact"
                                        $smsg = "Recipient:$($tid) => $($hshForwards[$tid])" ;
                                        write-host $smsg ;
                                        $hsum.xoContactForwards = $hshForwards[$tid] ;
                                    } ;
                                } ;
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
                }  # loop-E $hSum.xoRcpx
                # contacts and guests won't drop with $hSum.OPRemoteMailbox or $hSum.OPMailbox populated
                #region XO_GADU ; #*------v XO_GADU v------
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
                            if(-not $hSum.MGUser){
                                # $pltGadu.identity  +=  $hSum.AADUser.ImmutableId | convert-ImmuntableIDToGUID | select -expand guid ;
                                #Connect-AAD -Verbose:$false -silent ;
                                write-verbose "xoMailbox: ensure GET-ADUSER pulls MGUser.matched object for cloud recipient:`nfallback:get-AzureAdUser  -objectid $($hsum.xoRcp.ExternalDirectoryObjectId)" ;
                                # have to postfilter, if want specific count -maxresults catch's with no $error[0]
                                #$hSum.AADUser   +=  get-AzureAdUser  -objectid $hsum.xoRcp.ExternalDirectoryObjectId | select -first $MaxRecips;  ;
                                # forced upgrade blocked AAD -> MG
                                $hSum.MGUser   +=  get-MgUserFull -userid $hsum.xoRcp.ExternalDirectoryObjectId | select -first $MaxRecips; 
                            } ; 
                        } else {
                            throw "Unsupported object, blank `$hsum.xoRcp.ExternalDirectoryObjectId!" ;
                        } ;
                        if($hSum.xomailbox.IsDirSynced){
                            # doesn't mean hybrid exchange obj, means ADU anchored object
                            write-host "XOMailbox.IsDirSynced: anchored to ADUser" ; 
                            #if($hSum.AADUser.ExtensionProperty.onPremisesDistinguishedName){
                            # forcedmigr aad -> mg:
                            if($hsum.MGuser.OnPremisesDistinguishedName){
                                #switch -regex ($hSum.AADUser.ExtensionProperty.onPremisesDistinguishedName){
                                switch -regex ($hsum.MGuser.OnPremisesDistinguishedName){
                                    $rgxADDistNameAT{
                                        #$pltGadu.identity = $hSum.AADUser.ExtensionProperty.onPremisesDistinguishedName ; 
                                        $pltGadu.identity = $hsum.MGuser.OnPremisesDistinguishedName
                                        #$pltGadu.add('server',(($hSum.AADUser.ExtensionProperty.onPremisesDistinguishedName.split(',') | ?{$_ -match 'DC='} ) -replace 'DC=','') -join '.') ; 
                                        $pltGadu.add('server',(($hsum.MGuser.OnPremisesDistinguishedName.split(',') | ?{$_ -match 'DC='} ) -replace 'DC=','') -join '.') ; 
                                    }
                                    default{
                                        $smsg = "Unrecognized MGuser.OnPremisesDistinguishedName!" ; 
                                        $smsg += "`n$($hsum.MGuser.OnPremisesDistinguishedName)" ; 
                                        throw $smsg ;
                                    }
                                } ; 
                            } else {
                                write-warning "blank MGuser.OnPremisesDistinguishedName! (non-ADUser-sync'd object)" ; 
                            } ;  
                        }else{
                            write-warning "xomailbox is *NOT* IsDirSynced!, Cloud-first recipient, wo anchored ADUser object!" ; 
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
                            #$smsg += "`n$(($hSum.ADUser |fl 'Enabled','Description','whenCreated','whenChanged','Title' |out-string).trim())"
                            $smsg += "`n$(($hSum.ADUser |fl 'Enabled','Description','Info','whenCreated','whenChanged','Title' |out-string).trim())"
                        } ;
                        if($hSum.OPMailbox){
                            $smsg =  "$(($hSum.OPMailbox |fl $propsMailx|out-string).trim())" ;
                            #$smsg += "`n$(($hSum.ADUser |fl 'Enabled','Description','whenCreated','whenChanged','Title' |out-string).trim())"
                            $smsg += "`n$(($hSum.ADUser |fl 'Enabled','Description','Info','whenCreated','whenChanged','Title' |out-string).trim())"
                        } ; 
                        if( -not $hsum.OPRcp -AND $hsum.xoRcp -AND $hsum.xomailbox){ 
                            $smsg = "Cloud Mailbox is nonDirSync'd NON-HYBRID mail object!" ; 
                            $smsg += "`n$(($hSum.xoMailbox |fl $propsMailx|out-string).trim())" ; 
                            if($hsum.ADUser){
                                #if($hsum.Aaduser.DirSyncEnabled){
                                # forcedmigr aad -> mg:
                                if($hsum.MgUser.OnPremisesSyncEnabled){
                                    $smsg += "`nbut ADUser Object IS dirsync'd to AzureADUser object" ; 
                                } else { 
                                    $smsg += "`nADUser Object IS NOT dirsync'd to AzureADUser object" ; 
                                } ; 
                                #$smsg += "`n$(($hSum.ADUser |fl 'Enabled','Description','whenCreated','whenChanged','Title' |out-string).trim())"
                                $smsg += "`n$(($hSum.ADUser |fl 'Enabled','Description','Info','whenCreated','whenChanged','Title' |out-string).trim())"
                            } else {
                                $smsg += "`nNO ADUser Object appears to be cloud-first MGUser object" ;
                            }; 
                        } ;
                        if( -not $hsum.OPRcp -AND -not $hsum.xoRcp -AND $hSum.ADUser -AND $hsum.MGUser){
                            $smsg = "No detected OnPrem or Cloud Mail Recipient Objects detected" ; 
                            if($hSum.ADUser){
                                $smsg += "`nADUser Object IS NOT dirsync'd to MgUser object" ; 
                            } ; 
                            if($hsum.MgUser.OnPremisesSyncEnabled){
                                $smsg += "`nbut ADUser Object IS dirsync'd to MgUser object" ; 
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
                #endregion XO_GADU ; #*------^ END XO_GADU ^------
                #endregion XORCP_EXPAND ; #*------^ END XORCP_EXPAND ^------
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
                } ; 

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
                    #region xogetMobile2 ; #*------v xogetMobile2 v------
                    if($getMobile){
                        write-host -foreground yellow "===`$hsum.xoMobileDeviceStats: " #-nonewline;
                        $ino = 0 ;
                        foreach($xmob in $hsum.xoMobileDeviceStats){
                            $ino++ ;
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
                    #endregion xogetMobile2 ; #*------^ END xogetMobile2 ^------

                }elseif($hSum.xoMUser){
                    write-host "=get-xMUSR:>`n$(($hSum.xoMUser |fl ($propsMailx |?{$_ -notmatch '(sam.*|dist.*)'})|out-string).trim())`n-Title:$($hSum.xoUser.Title)";
                }elseif($hSum.txGuest){
                    write-host "=get-AADU:>`n$(($hSum.txGuest |fl userp*,PhysicalDeliveryOfficeName,JobTitle|out-string).trim())"
                } ;

                # populate xoMemberOf
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
                #region XORCP_RETRY ; #*------v XORCP_RETRY v------
                write-warning "(no matching EXOP or EXO recipient object:$($usr))"
                # do near Lname[0-3]* searches for comparison
                if($hSum.lname){
                    write-warning "Lname ($($hSum.lname) parsed from input),`nattempting similar LName g-rcp:...`n(up to `$MaxRecips:$($MaxRecips))" ;
                    $lname = $hsum.lname ;
                    #$fltrB = "displayname -like '*$lname*'" ;
                    #write-verbose "RETRY:get-recipient -filter {$($fltr)}" ;
                    #get-recipient "$($txusr.lastname.substring(0,3))*"| sort name
                    $substring = "$($hSum.lname.substring(0,3))*"
                    
                    if($resolveForwards){
                        write-host "get-recipient -id $($substring) -ea 0 :"
                    }else{
                        write-host "get-recipient -id $($substring) -ea 0 |?{$_.recipienttypedetails -ne 'MailContact'} :"
                    }
                    #==9:21 AM 10/8/2024:  since HR/WD change to SamAcctName as employe#, the above won't match any user created since 2022 or so. , 
                    # need to search on last name first

                    #if($hSum.Rcp += get-recipient -id "$($substring)" -ea 0 -ResultSize $MaxRecips | select -first $MaxRecips |?{$_.recipienttypedetails -ne 'MailContact'}){
                    if($hSum.Rcp += get-recipient -id "$($substring)" -ea 0 -ResultSize $MaxRecips | select -first $MaxRecips){
                        if($resolveForwards){

                        }else{
                            $hSum.Rcp = $hSum.Rcp  |?{$_.recipienttypedetails -ne 'MailContact'}
                        }
                        #$hSum.Rcp | write-output ;
                        # $propsRcpTbl
                        write-host -foregroundcolor yellow "`n$(($hSum.Rcp | ft -a $propsRcpTbl |out-string).trim())" ;
                    } ;
                    #write-host "$((get-alias ps1GetxRcp).definition) -id $($substring) -ea 0 |?{$_.recipienttypedetails -ne 'MailContact'} : "
                    if($resolveForwards){
                        write-host "get-xorecipient -id $($substring) -ea 0 : "
                    }else {
                        write-host "get-xorecipient -id $($substring) -ea 0 |?{$_.recipienttypedetails -ne 'MailContact'} : "
                    }
                    #if($hSum.xoRcp += get-xorecipient -id "$($substring)" -ea 0 -ResultSize $MaxRecips | select -first $MaxRecips |?{$_.recipienttypedetails -ne 'MailContact'}){
                    if($hSum.xoRcp += get-xorecipient -id "$($substring)" -ea 0 -ResultSize $MaxRecips | select -first $MaxRecips){
                        if($resolveForwards){

                        }else {
                            $hSum.xoRcp = $hSum.xoRcp|?{$_.recipienttypedetails -ne 'MailContact'} 
                        }
                        #$hSum.xoRcp | write-output ;
                        write-host -foregroundcolor yellow "`n$(($hSum.xoRcp | ft -a $propsRcpTbl |out-string).trim())" ;
                    } ;


                } ;
                #endregion XORCP_RETRY ; #*------^ END XORCP_RETRY ^------
                #region GADU_NAME ; #*------v GADU_NAME v------
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
                #endregion GADU_NAME ; #*------^ END GADU_NAME ^------
                #region GMGU_NAME ; #*------v GMGU_NAME v------
                # do MGUser search on fname/lname
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
                    #if($tmpAADo = get-AzureAdUser  -filter $fltr -ea 0 | select -first $MaxRecips){
                    # forced M$ migre AAD -> MG
                    # Get-MgUser -Filter "givenName eq 'John' and surname eq 'Doe'" -All
                    if($tmpAADo = Get-MgUser -filter $fltr  -All -ea 0 | select -first $MaxRecips){
                        $smsg = "Matched on:get-AzureAdUser -filter $($fltr) " ; 
                        write-verbose $smsg ; 
                    }elseif($hSum.lname){
                        # treat as a -ObjectId                        
                        #if($tmpAADo = get-AzureAdUser -ObjectId $hSum.lname -ea 0 | select -first $MaxRecips){
                        if($tmpAADo = Get-MgUser -userId $hSum.lname -All -ea 0 | select -first $MaxRecips){
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

MGUser #$($rno):DN:$(($thisADU.distinguishedname|out-string).trim())
$(($thisADU | ft -a $prpADU[1..3]|out-string).trim())
$(($thisADU | ft -a  $prpADU[4..7]|out-string).trim()) 
$(($thisADU | ft -a  $prpADU[8..11]|out-string).trim())

"@ ;
                            write-host $hsADU ; 
                       } ; 
                    } ; 
                } 
                #endregion GMGU_NAME ; #*------^ END GMGU_NAME ^------

                $abortReport = $true ; 

            } ; # don't break, doesn't continue loop
            #endregion OP_V_XO_RCPEXPAND ; #*------^ END OP_V_XO_RCPEXPAND ^------

            if($abortReport ){
                $smsg = "(multiple recipients - or no recipients, but ADUsers, or but MGUsers -  found in OnPrem And/Or Cloud, detailed reporting & output aborted)" ; 
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                BREAK ; 
            } ; 

            #region FORCE_XOMBXINFO ; #*------v FORCE_XOMBXINFO v------
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
            #region xogetMobile ; #*------v xogetMobile v------
            if($getMobile){
                get-xoMobileData ;                 
            } ; 
            #endregion xogetMobile ; #*------^ END xogetMobile ^------
            #region xogetQuotaUsage2 ; #*------v xogetQuotaUsage2 v------
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
                    $smsg = "xoMailboxStats Count:$(($hSum.xoMailboxStats|measure).count)" ;
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
                        if($DeletedItems){
                            $hsum.xoMailboxFolderStats |
                                select $prpFldrDeleted | sort TreeSizeMB -desc | export-csv  -path $ofMbxFolderStats -notype ;
                        }else{
                            $hsum.xoMailboxFolderStats | ?{$_.ItemsInFolder -gt 0 -AND $_.identity -notmatch $rgxHiddn } | 
                                select $prpFldr | sort SizeMB -desc | export-csv  -path $ofMbxFolderStats -notype ;
                        }

                    } CATCH {
                        $ErrTrapd=$Error[0] ;
                        $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                        write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" ;
                    } ; 
                    
                }
            } ; 
            #endregion xogetQuotaUsage2 ; #*------^ END xogetQuotaUsage2 ^------
            #region xogetPerms2 ; #*------v xogetPerms2 v------
            if($getPerms){
                $pltGMbxPermX=[ordered]@{
                    identity = $hSum.xoMailbox.exchangeguid ;
                    ErrorAction = 'STOP' ;
                } ;
                $smsg = "Get-xoMailboxPermission  w`n$(($pltGMbxPermX|out-string).trim())"
                if($verbose){
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else{ write-verbose $smsg } ;
                } ;
                TRY{
                    $hSum.xoMailboxPermission  +=  Get-xoMailboxPermission @pltGMbxPermX | ?{$_.user -notmatch 'NT\sAUTHORITY\\SELF'} | select $prpMPerms;
                    $smsg = "xoMailboxPermission Count:$(($hsum.xoMailboxPermission|measure).count)" ;
                    write-host -foregroundcolor green $smsg ;
                    if($hSum.xoMailboxPermission){
                        foreach($grp in ($hSum.xoMailboxPermission.user |
                            get-xorecipient  | ?{$_.recipienttype -eq 'MailUniversalSecurityGroup'}) ){
                            $hshGrpSumm = [ordered]@{
                                Identity = $grp.Identity
                                PrimarySmtpAddress = $grp.PrimarySmtpAddress ;
                                Description = $grp.Description ;
                                RecipientType = $grp.RecipientType ;
                                RecipientTypeDetails = $grp.RecipientTypeDetails ;
                                ManagedBy = ($grp | get-xodistributiongroup | select -expand managedby | get-xorecipient -ea 0) ;
                                Members = ($grp | get-xodistributiongroupmember | get-xorecipient  -ea 0) ;
                            } ; 
                            $hSum.xoMailboxPermissionGroups += [pscustomobject]$hshGrpSumm ;
                        } ;
                    } else {
                        $smsg = "(no non-SELF Get-xoMailboxPermission returned)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                    };
                } CATCH {
                    $ErrTrapd=$Error[0] ;
                    $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                    write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" ;
                } ;
                $smsg = "Get-xoRecipientPermission  w`n$(($pltGMbxPermX|out-string).trim())"
                if($verbose){
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else{ write-verbose $smsg } ;
                } ;
                TRY{
                    $hsum.xoRecipientPermission += Get-xoRecipientPermission @pltGMbxPermX | ?{$_.trustee -notmatch 'NT\sAUTHORITY\\SELF'}  | select $prpRPerms;
                    $smsg = "xoRecipientPermission Count:$(($hsum.xoRecipientPermission|measure).count)" ;
                    write-host -foregroundcolor green $smsg ;
                    if($hSum.xoRecipientPermission){
                        foreach($grp in ($hsum.xoRecipientPermission.trustee |
                            get-xorecipient  | ?{$_.recipienttype -eq 'MailUniversalSecurityGroup'}) ){
                            $hshGrpSumm = [ordered]@{
                                Identity = $grp.Identity
                                PrimarySmtpAddress = $grp.PrimarySmtpAddress ;
                                Description = $grp.Description ;
                                RecipientType = $grp.RecipientType ;
                                RecipientTypeDetails = $grp.RecipientTypeDetails ;
                                ManagedBy = ($grp | get-xodistributiongroup | select -expand managedby | get-xorecipient -ea 0) ;
                                Members = ($grp | get-xodistributiongroupmember | get-xorecipient  -ea 0) ;
                            } ; 
                            $hSum.xoRecipientPermissionGroups += [pscustomobject]$hshGrpSumm ;
                        } ;
                    } else {
                        $smsg = "(no non-SELF Get-xoRecipientPermission returned)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                    };
                } CATCH {
                    $ErrTrapd=$Error[0] ;
                    $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                    write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" ;
                } ;
            }
            #endregion xogetPerms2 ; #*------^ END xogetPerms2 ^------
            #endregion FORCE_XOMBXINFO ; #*------^ END FORCE_XOMBXINFO ^------

            #region RV_VIA_GMGU ; #*------v RV_VIA_GMGU v------
            #$pltgMsoUsr=@{UserPrincipalName=$null ; MaxResults= $MaxRecips; ErrorAction= 'STOP' } ;
            # maxresults is documented:
            # but causes a fault with no $error[0], doesn't seem to be functional param, post-filter
            # ren refs of $pltgMsoUsr -> $pltgAADUsr
            $pltgMGUsr=@{UserPrincipalName=$null ; ErrorAction= 'STOP' } ;
            if($hSum.ADUser){$pltgMGUsr.UserPrincipalName  +=  $hSum.ADUser.UserPrincipalName }
            elseif($hSum.xoMailbox){$pltgMGUsr.UserPrincipalName += $hsum.xoMailbox.UserPrincipalName }
            elseif($hSum.xoMUser){$pltgMGUsr.UserPrincipalName  +=  $hSum.xoMUser.UserPrincipalName }
            elseif($hSum.txGuest){$pltgMGUsr.UserPrincipalName  +=  $hSum.txGuest.userprincipalname }
            else{} ;

            if($pltgMGUsr.UserPrincipalName){
                #region FORCE_GMGU ; #*------v FORCE_GMGU v------
                if(-not($hSum.MGUser)){
                    write-host -foregroundcolor yellow "=get-MGUser -userid $($pltgMGUsr.UserPrincipalName)>:" ;
                    TRY{
                        #caad  -Verbose:$false -silent ;
                        #write-verbose "get-AzureAdUser  -objectid $($pltgMGUsr.UserPrincipalName)" ;
                        write-verbose "Get-MgUser -userid  $($pltgMGUsr.UserPrincipalName)" ;
                        # have to postfilter, if want specific count -maxresults catch's with no $error[0]
                        #$hSum.AADUser   +=  get-AzureAdUser  -objectid $pltgMGUsr.UserPrincipalName  | select -first $MaxRecips;  ;
                        # forced migration AAD -> MG
                        $hSum.MGUser   +=  get-MgUserFull -userid $pltgMGUsr.UserPrincipalName | select -first $MaxRecips;  ;
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
                            extension_9d88b2c96135413e88afff067058e860_employeeNumber 1234
                             $hsum.aaduser.ExtensionProperty.onPremisesDistinguishedName
                            CN=XXX,OU=XXX,...

                            # MGUser has these: (unpop'd in some cases)
                            OnPremisesDistinguishedName           : CN=Todd Kadrie,OU=IS,OU=Users,OU=LYN,DC=global,DC=ad,DC=toro,DC=com
                            OnPremisesDomainName                  : global.ad.toro.com
                            OnPremisesExtensionAttributes         : Microsoft.Graph.PowerShell.Models.MicrosoftGraphOnPremisesExtensionAttributes
                            OnPremisesImmutableId                 : UC7OjGZAYUK/qsVZatOR6g==
                            OnPremisesLastSyncDateTime            :
                            OnPremisesProvisioningErrors          :
                            OnPremisesSamAccountName              : kadrits
                            OnPremisesSecurityIdentifier          : S-1-5-21-2222296782-158576315-1096482972-75073
                            OnPremisesSyncEnabled                 : True
                            OnPremisesUserPrincipalName           :

                        #>
                        #write-verbose "`$hSum.AADUser:`n$(($hSum.AADUser|out-string).trim())" ;
                        # ObjectId                             DisplayName   UserPrincipalName      UserType

                        #lic pull: $hSum.AADUser | Get-AzureADUserLicenseDetail -ea STOP | select -exp SkuPartNumber
                        write-verbose "`$hSum.MGUserLics = Get-MgUserLicenseDetail -userid $($hSum.MGUser.userprincipalname) | select -exp SkuPartNumber" ;
                        #$hsum.AADUserLics  +=   $hSum.AADUser | Get-AzureADUserLicenseDetail -ea STOP | select -exp SkuPartNumber ; 
                        # forcedmigr aad -> mg: Get-MgUserLicenseDetail -userid $mgu.userprincipalname
                        #$hsum.MGUserLics  +=   $hSum.MGUser | Get-MgUserLicenseDetail -ea STOP | select -exp SkuPartNumber ; 
                        # errors out, doesen't properly support pipeline (garbage)
                        $hsum.MGUserLics  +=  Get-MgUserLicenseDetail -userid $hSum.MGUser.userprincipalname -ea STOP | select -exp SkuPartNumber ; 

                    } CATCH {
                        $ErrTrapd=$Error[0] ;
                        $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                        else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        Continue ;
                    } ;

                } ;
                #endregion FORCE_GMGU ; #*------^ END FORCE_GMGU ^------
                #region FORCE_MGU_MGR ; #*------v FORCE_MGU_MGR v------
                #if(-not($hSum.AADUserMgr) -AND $hSum.AADUser ){
                if(-not($hSum.MGUserMgr) -AND $hSum.MGUser ){
                    write-host -foregroundcolor yellow "=Get-MgUserManager $($hSum.MGUser.UserPrincipalName)>:" ;
                    TRY{
                        #caad  -Verbose:$false -silent ;
                        write-verbose "Get-MgUserManager  -objectid $($hSum.MGUser.UserPrincipalName)" ;
                        # have to postfilter, if want specific count -maxresults catch's with no $error[0]
                        #$hSum.MGUserMgr   +=  get-AzureAdUserManager  -objectid $hSum.MGUser.UserPrincipalName  | select -first $MaxRecips;  ;
                        # forcedmigr aad -> mg: 
                        # they've buried all but the id property in AdditionalProperties [fk that!]
                        #$hSum.MGUserMgr   +=  (Get-MgUserManager  -userid $hSum.MGUser.UserPrincipalName  -Property $prpMGUser | select -first $MaxRecips).AdditionalProperties                        
                        $hSum.MGUserMgr   +=  (Get-MgUserManager  -userid $hSum.MGUser.UserPrincipalName  -Property $prpMGUser | select -first $MaxRecips).AdditionalProperties                        
                        #write-verbose "`$hSum.MGUserMgr:`n$(($hSum.MGUserMgr|out-string).trim())" ;
                        # (returns a full AADUser obj for the mgr)
                        # we can output the DN: $hSum.MGUserMgr.ExtensionProperty.onPremisesDistinguishedName
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
                #endregion FORCE_MGU_MGR ; #*------^ END FORCE_MGU_MGR ^------
                #region FORCE_MGU_MEMBEROF ; #*------v FORCE_MGU_MEMBEROF v------
                if(-not $hSum.xoMemberOf -AND $hsum.xoMailbox.DistinguishedName){
                    TRY {
                        write-verbose "get-xorecipient -Filter {Members -eq '$($hsum.xoMailbox.DistinguishedName)'}`n -RecipientTypeDetails GroupMailbox,MailUniversalDistributionGroup,MailUniversalSecurityGroup"
                        $hSum.xoMemberOf  +=  get-xorecipient -Filter "Members -eq '$($hsum.xoMailbox.DistinguishedName)'" -RecipientTypeDetails GroupMailbox,MailUniversalDistributionGroup,MailUniversalSecurityGroup ;
                        write-verbose "`$hSum.xoMemberOf:`n$(($hSum.xoMemberOf|out-string).trim())" ;
                    } CATCH {
                        $ErrTrapd=$Error[0] ;
                        $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                        else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        Continue ;
                    } ;
                }else{
                    $smsg = "Unable to: `$hSum.xoMemberOf = Get-MgUserMemberOf -UserId $($hSum.MGUser.UserPrincipalName)" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                }
                # as is typical, the MG command is too big a PITA and doesn't return what you want anyway. Use the XO cmdlets. 
                <#
                if(-not $hSum.xoMemberOf -AND $hSum.MGUser.UserPrincipalName){
                    TRY{
                        $hSum.xoMemberOf = (Get-MgUserMemberOf -UserId $hSum.MGUser.UserPrincipalName -Property "DisplayName", "mailNickname" -ConsistencyLevel eventual -ea STOP ).AdditionalProperties.mailNickname
                    } CATCH {
                        $ErrTrapd=$Error[0] ;
                        $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                        else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        Continue ;
                    } ;
                }else{
                    $smsg = "Unable to: `$hSum.xoMemberOf = Get-MgUserMemberOf -UserId $($hSum.MGUser.UserPrincipalName)" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                }
                #>
                #endregion FORCE_MGU_MEMBEROF ; #*------^ END FORCE_MGU_MEMBEROF ^------
                # display user info:
                #region OUTPUT_P1 ; #*------v OUTPUT_P1 v------
                if(-not($hSum.ADUser)){
                    if($hSum.MGUser.OnPremisesSyncEnabled -AND $hSum.aaduser.ExtensionProperty.onPremisesDistinguishedName){
                        #region ADU_FEDERATED ; #*------v ADU_FEDERATED v------
                        $pltGadu.Identity = $hSum.MGUser.ExtensionProperty.onPremisesDistinguishedName ; 
                        $hSum.ADUser  += Get-ADUser @pltGadu | select -first $MaxRecips ;
                        if($pltGadu.identity){
                            write-verbose "Get-ADUser w`n$(($pltGadu|out-string).trim())" ;
                            # try a nested local trycatch, against a missing result
                            Try {
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
                                #$smsg += "`n$(($hSum.ADUser |fl 'Enabled','Description','whenCreated','whenChanged','Title' |out-string).trim())"
                                $smsg += "`n$(($hSum.ADUser |fl 'Enabled','Description','Info','whenCreated','whenChanged','Title' |out-string).trim())"
                            } ;
                            if($hSum.OPMailbox){
                                $smsg =  "$(($hSum.OPMailbox |fl $propsMailx|out-string).trim())" ;
                                #$smsg += "`n$(($hSum.ADUser |fl 'Enabled','Description','whenCreated','whenChanged','Title' |out-string).trim())"
                                $smsg += "`n$(($hSum.ADUser |fl 'Enabled','Description','Info','whenCreated','whenChanged','Title' |out-string).trim())"
                            } ;
                            write-host $smsg ;
                        } ;
                        #endregion ADU_FEDERATED ; #*------^ END ADU_FEDERATED ^------
                    } else { 
                        #region REMOTE_ADU_FEDERATED ; #*------v REMOTE_ADU_FEDERATED v------
                        # remote fed, use AADU to proxy remote AD hybrid info:
                        write-host -foreground yellow "===`$hSum.MGUser: " #-nonewline;
                        <#
                        $smsg = "$(($hSum.AADUser| select $propsAADL1 |out-markdowntable @MDtbl |out-string).trim())" ;
                        $smsg += "`n$(($hSum.AADUser|select $propsAADL2 |out-markdowntable @MDtbl|out-string).trim())" ;
                        $smsg += "`n$(($hSum.AADUser|select $propsAADL3 |out-markdowntable @MDtbl|out-string).trim())" ;
                        $smsg += "`n$(($hSum.AADUser|select $propsAADL4 |out-markdowntable @MDtbl|out-string).trim())" ;
                        $smsg += "`n$(($hSum.AADUser|select $propsAADL5 |out-markdowntable @MDtbl|out-string).trim())" ;
                        #>
                        # retool for MGU
                        $smsg = "$(($hSum.MGUser| select $propsMGUL1 |out-markdowntable @MDtbl |out-string).trim())" ;
                        $smsg += "`n$(($hSum.MGUser|select $propsMGUL2 |out-markdowntable @MDtbl|out-string).trim())" ;
                        $smsg += "`n$(($hSum.MGUser|select $propsMGUL3 |out-markdowntable @MDtbl|out-string).trim())" ;
                        $smsg += "`n$(($hSum.MGUser|select $propsMGUL4 |out-markdowntable @MDtbl|out-string).trim())" ;
                        $smsg += "`n$(($hSum.MGUser|select $propsMGUL5 |out-markdowntable @MDtbl|out-string).trim())" ;
                        #$hsum.aaduser.ExtensionProperty.onPremisesDistinguishedName
                        if($hSum.Federator -ne $TORMeta.adforestname){
                            $smsg += "`n$($hSum.Federator):Remote ADUser.DN:`n$(($hSum.MGUser.ExtensionProperty.onPremisesDistinguishedName|out-string).trim())" ;
                        }  ;
                        #endregion REMOTE_ADU_FEDERATED ; #*------^ END REMOTE_ADU_FEDERATED ^------
                    }; 

                    write-host $smsg

                    # assert the real names from the user obj
                    $hSum.dname  +=  $hSum.MGUser.DisplayName ;
                    $hSum.fname  +=  $hSum.MGUser.GivenName ;
                    $hSum.lname  +=  $hSum.MGUser.Surname ;

                } else {
                    #region OUTPUT_ADU_INFO ; #*------v OUTPUT_ADU_INFO v------
                    # defer to ADUser details
                    write-host -foreground yellow "===`$hSum.ADUser: " #-nonewline;
                    $smsg = "$(($hSum.ADUser| select $propsADL1 |out-markdowntable @MDtbl |out-string).trim())" ;
                    $smsg += "`n$(($hSum.ADUser|select $propsADL2 |out-markdowntable @MDtbl|out-string).trim())" ;
                    $smsg += "`n$(($hSum.ADUser|select $propsADL3 |out-markdowntable @MDtbl|out-string).trim())" ;
                    $smsg += "`n$(($hSum.ADUser|select $propsADL4 |out-markdowntable @MDtbl|out-string).trim())" ;
                    $smsg += "`n$(($hSum.ADUser|select $propsADL5 |out-markdowntable @MDtbl|out-string).trim())" ;
                    # stick desc on trailing line $propsADL5
                    #$smsg += "`n$(($hSum.ADUser|select $propsADL5 |out-markdowntable @MDtbl|out-string).trim())" ;
                    # flip L5 to fl (suppress crlf wrap)
                    $smsg += "`n$(($hSum.ADUser|select $propsADL6 |Format-List|out-string).trim())" ;
                    if($hsum.ADUser.Info){
                        $smsg += "`n$(($hSum.ADUser|select $propsADL7 |Format-List|out-string).trim())" ;
                    } ;
                    # moved DN into adl4, w enabled
                    #$smsg += "`n`$ADUser.DN:`n$(($hsum.aduser.DistinguishedName|out-string).trim())" ;
                    #$smsg += "`n$($hSum.ADUser|select Enabled,distinguishedname| convertTo-MarkdownTable -NoDashRow -Border) `$ADUser.DN:`n$(($hsum.aduser.DistinguishedName|out-string).trim())" ;
                    write-host $smsg ;

                    # assert the real names from the user obj
                    $hSum.dname  +=  $hSum.ADUser.DisplayName ;
                    $hSum.fname  +=  $hSum.ADUser.GivenName ;
                    $hSum.lname  +=  $hSum.ADUser.Surname ;
                    #endregion OUTPUT_ADU_INFO ; #*------^ END OUTPUT_ADU_INFO ^------
                } ;
                #endregion OUTPUT_P1 ; #*------^ END OUTPUT_P1 ^------
                #region ENABLED_STATUS ; #*------v ENABLED_STATUS v------
                # aduser enabled/disabled: .aduser.Enbabled
                if($hSum.aduser){
                    if($hSum.aduser.Enabled){
                        if($hsum.xoRcp.RecipientTypeDetails -match 'SharedMailbox|RoomMailbox|EquipmentMailbox'){
                            $smsg = "ADUser:$($hSum.ADUser.userprincipalname) AD Account w $($hsum.xoRcp.RecipientTypeDetails) mbx is *ENABLED!*"
                            write-warning $smsg ;
                        } ;
                    } else {
                        if($hsum.xoRcp.RecipientTypeDetails -match 'SharedMailbox|RoomMailbox|EquipmentMailbox'){} else { 
                            $smsg = "ADUser:$($hSum.ADUser.userprincipalname) AD Account w $($hsum.xoRcp.RecipientTypeDetails) is *DISABLED!*"
                            write-warning $smsg ;
                        } ; 
                    } ; 
                } ;
                # AADUser enabled/disabled: .aaduser.AccountEnabled
                if($hSum.MGUser){
                    # 2:31 PM 9/23/2025 fixed typo: .Enabled -> .AccountEnabled
                    if($hSum.MGUser.AccountEnabled){
                        if($hsum.xoRcp.RecipientTypeDetails -match 'SharedMailbox|RoomMailbox|EquipmentMailbox'){
                            $smsg = "ADUser:$($hSum.MGUser.userprincipalname) AD Account w $($hsum.xoRcp.RecipientTypeDetails) mbx is *ENABLED!*"
                            write-warning $smsg ;
                        } ;
                    } else {
                        if($hsum.xoRcp.RecipientTypeDetails -match 'SharedMailbox|RoomMailbox|EquipmentMailbox'){} else { 
                            $smsg = "ADUser:$($hSum.MGUser.userprincipalname) AD Account w $($hsum.xoRcp.RecipientTypeDetails) is *DISABLED!*"
                            write-warning $smsg ;
                        } ; 
                    } ; 
                } ;
                #endregion ENABLED_STATUS ; #*------^ END ENABLED_STATUS ^------
                #region LIC_GRP ; #*------v LIC_GRP v------
                if($hSum.ADUser){$hSum.LicenseGroup  +=  $hSum.ADUser.memberof |?{$_ -match $rgxOPLic }}
                elseif($hSum.xoMemberOf){$hSum.LicenseGroup  +=  $hSum.xoMemberOf.Name |?{$_ -match $rgxXLic}}

                #if(-not ($hSum.LicenseGroup) -AND ($hSum.MsolUser.licenses.AccountSkuId -contains "$($TORMeta.o365_TenantDom.tolower()):ENTERPRISEPACK")){$hSum.LicenseGroup  +=  '(direct-assigned E3)'} ;
                # $hSum.AADUser ; $aadu | Get-AzureADUserLicenseDetail  | select -exp SkuPartNumber
                #if(-not ($hSum.LicenseGroup) -AND ( $hsum.AADUserLics  -contains "$($TORMeta.o365_TenantDom.tolower()):ENTERPRISEPACK")){$hSum.LicenseGroup  +=  '(direct-assigned E3)'} ;
                # no dom, with aadu licenses
                if(-not ($hSum.LicenseGroup) -AND ( $hSum.MGUserLics  -contains "ENTERPRISEPACK")){$hSum.LicenseGroup  +=  '(direct-assigned E3)'} ;
                if($hSum.LicenseGroup){$smsg = "LicenseGroup:$($hSum.LicenseGroup)"}
                else{$smsg = "LicenseGroup:(unresolved, direct-assigned other?)" } ;
                write-host $smsg ;
                #endregion LIC_GRP ; #*------^ END LIC_GRP ^------
                #region OUTPUT_MGUserMgr ; #*------v OUTPUT_MGUserMgr v------
                if($hSum.MGUserMgr){
                    #($hSum.AADUserMgr) |ft -a  $propsaadmgr
                    #$smsg += "`nAADUserMgr:`n$(($hSum.AADUserMgr|select $propsAadMgr |out-markdowntable @MDtbl|out-string).trim())" ;
                    # $propsAADMgrL1, $propsAADMgrL2
                    write-host -foreground yellow "===`$hSum.MGUserMgr: " #-nonewline;
                    #$smsg = "$(($hSum.MGUserMgr| select $propsAADMgrL1 |out-markdowntable @MDtbl |out-string).trim())" ;
                    $smsg = "$(($hSum.MGUserMgr| select $propsMGUMgrL1 |out-markdowntable @MDtbl |out-string).trim())" ;
                    #$smsg += "`n$(($hSum.MGUserMgr|select $propsAADMgrL2 |out-markdowntable @MDtbl|out-string).trim())" ;
                    #$smsg += "`n$(($hSum.MGUserMgr|Format-List $propsAADMgrL2|out-string).trim())" ;
                    #$smsg += "`n$(($hSum.MGUserMgr|Format-List $propsMGUMgrL2|out-string).trim())" ;
                    # for mguser, and it's hash, have to change the syntax:
                    $smsg += "`n$(($hSum.MGUserMgr | select $propsMGUMgrL2|out-string).trim())"
                    #$smsg += "`n$(($hSum.MGUserMgr|select $propsADL3 |out-markdowntable @MDtbl|out-string).trim())" ;
                } else {
                    $smsg = "(AADUserMgr was blank, or unresolved)" ;
                } ;
                write-host $smsg ;
                #endregion OUTPUT_MGUserMgr ; #*------^ END OUTPUT_MGUserMgr ^------
                #region OUTPUT_QUOTA_N_SIZE ; #*------v OUTPUT_QUOTA_N_SIZE v------
                if($getQuotaUsage -AND $hSum.xoMailbox){

                    $smsg += "`n`nLicenses:`n$(($hSum.MGUserLics -join ', ' |out-string).trim())`n`n" ; 
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

                    if($DeletedItems){
                        $smsg = "`nWith the following non-zero folder metrics`n`n$((import-csv $ofMbxFolderStats  | ?{$_.ItemsInFolder -gt 0 -AND $_.identity -notmatch $rgxHiddn } |select $prpFldr | ft -auto |out-string).trim())" ; 
                        $smsg += "`n`nAnd the Following Deleted-Items-related folder metrics`n`n$((import-csv $ofMbxFolderStats | ?{$_.identity -match $rgxDelItmsShow } |select $prpFldrDeleted | ft -auto |out-string).trim())`n`n" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level PROMPT } 
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                    }else{
                        $smsg = "`nWith the following non-zero folder metrics`n`n$((import-csv $ofMbxFolderStats | ft -auto |out-string).trim())" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level PROMPT } 
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                    } ; 

                    $smsg = "`n===`output to:`n$($ofMbxFolderStats)`n" ;
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
                    if($hSum.xoMailbox.recipienttypedetails -eq 'SharedMailbox'){
                        $hsInfoSharedMbx = @"

## Shared Mailbox Outlook Handling of Deleted Items & Sent Items. 

*Please note*: The subject mailbox is a _SharedMailbox_-type, which will 
trigger a _Delegate's Outlook client_ to perform special handling of the 
following actions by the Delegate: 

### Deleted Items: 

> When [a Delegated user uses] Microsoft Outlook to delete items from a mailbox folder of another 
> user for whom [the Delegate has] deletion privileges, the deleted items go to *[the Delegate's] own 
> Deleted Items folder* instead of the Deleted Items folder of the mailbox owner. 

Ref: [Items that are deleted from a shared mailbox go to the wrong folder in Outlook - Outlook | Microsoft Learn]
(https://learn.microsoft.com/en-us/troubleshoot/outlook/email-management/deleted-items-go-to-wrong-folder)

The Outlook behavior is controlled through configuring _the Delegate's Legacy Outlook client)_ 
with a custom Registry Key (with Service Desk assistance), that manages the 
Delegate's preference for storage of third-party Sent Items, sent from their Legacy Outlook client.  

The details of workstation registry modification process are covered in the Service Desk kb, and documented by Microsoft at:

[Switch the destination of deleted items - Outlook | Microsoft Learn]
(https://learn.microsoft.com/en-us/troubleshoot/outlook/email-management/deleted-items-go-to-wrong-folder#switch-the-destination-of-deleted-items)

The article above details configuration of the following custom registry key on the 
Delegate's Legacy Outlook workstation: 

    Note: As of October of 2025, Microsoft has not yet delivered an equivelent configurable setting for New Outlook,
    Legacy Outlook is *required* for configuration of Delegate preferences for Outlook mail handling actions.

    HKEY_CURRENT_USER\Software\Microsoft\Office\<x.0>\Outlook\Options\General

    Note: The <x.0> placeholder represents your version of Office (16.0 = Office 2016, Office 2019, Office LTSC 2021, or Microsoft 365, 15.0 = Office 2013).

    DelegateWastebasketStyle, DWORD Value:

    8 = Stores deleted items in [the Delegate's] folder.
    4 = Stores deleted items in the mailbox owner (e.g. the Shared Mailbox) folder 

    Note: Unlike Sent Items behavior (covered below), there is *no* administrator
    configurable setting available, to implement the configuration above 
    directly on a Shared Mailbox.    

### Sent Mail from the Shared Mailbox address:

> [When] using Microsoft Outlook 2016 or a later version, and a user has been delegated 
> permission to send email messages as another user or on behalf of another user from a shared mailbox. 
> ... when [they] send a message as another user or on behalf of the user, the 
> sent message isn't saved to the Sent Items folder of  the shared mailbox. 
> *Instead, _it's saved to the Sent Items folder of [the Delegate's] mailbox*. 

Ref: [Messages sent from a shared mailbox aren't saved to the Sent Items folder - Exchange | Microsoft Learn]
(https://learn.microsoft.com/en-us/troubleshoot/exchange/user-and-shared-mailboxes/sent-mail-is-not-saved?source=recommendations)

The Outlook behavior is controlled through one of two ways:

1. The Delegate's workstation Legacy Outlook client, can be configured 
(with Service Desk assistance), to store 3rd party Sent Items, 
_to the 3rd party mailbox_ by setting the DelegateSentItemsStyle registry key.  

    The details of this process are covered in the Service Desk kb, and documented by Microsoft at:

    [Messages sent from a shared mailbox aren't saved to the Sent Items folder - Exchange | Microsoft Learn]
    (https://learn.microsoft.com/en-us/troubleshoot/exchange/user-and-shared-mailboxes/sent-mail-is-not-saved?source=recommendations)

    The article above details configuration of the following custom registry key on the 
    Delegate's Legacy Outlook workstation: 

        HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Outlook\Preferences

        DelegateSentItemsStyle,  DWORD 32-bit Value.

    DelegateSentItemsStyle | MessageCopyForSentAsEnabled | Expected behavior
    ---------------------- | --------------------------- | -------------------------------------------------------------------------------------------------
    0                      | True                        | A copy of the email will be saved in 
                           |                             | both the primary mailbox and the 
                           |                             | shared mailbox.
    1                      | True                        | Two copies of the email will be saved
                           |                             | in the shared mailbox and no copies 
                           |                             | in the primary mailbox.
    0                      | False                       | A copy of the email will be saved in 
                           |                             | the primary mailbox and no copies 
                           |                             | in the shared mailbox.
    1                      | False                       | A copy of the email will be saved in 
                           |                             | the shared mailbox and no copies 
                           |                             | in the primary mailbox.

2. Or, the Shared mailbox can be configured by an administrator to save 
messages to the Sharted mailbox through a powershell modification on the Shared 
Mailbox itself : 

 - For emails sent as the shared mailbox, run the following command in Exchange PowerShell:

        set-mailbox <mailbox name> -MessageCopyForSentAsEnabled `$True

 - For emails sent on behalf of the shared mailbox, run the following command in Exchange PowerShell:

        set-mailbox <mailbox name> -MessageCopyForSendOnBehalfEnabled `$True

"@ ; 
                        write-host  $hsInfoSharedMbx ; 
                    } ; 

                } ; 
                #endregion OUTPUT_QUOTA_N_SIZE ; #*------^ END OUTPUT_QUOTA_N_SIZE ^------
                #region OUTPUT_PERMS ; #*------v OUTPUT_PERMS v------
                if($getPerms -AND $hSum.xoMailbox){

                    if($hSum.xoMailboxPermission){
                        $smsg = "`n## xoMailboxPermission:`n$(($hsum.xoMailboxPermission | ft -a $prpMPerms |out-string).trim())`n" ; 
                        if($hSum.xoMailboxPermissionGroups){
                            $smsg += "`n### Expanded Perm Group Summaries:" ; 
                            foreach($grp in $hSum.xoMailboxPermissionGroups){
                                $smsg += "`n-----------" ; 
                                $smsg += "`n$(($grp |select $propsDG[0..1] |out-markdowntable @MDtbl|out-string).trim())" ;
                                $smsg += "`n$(($grp |select $propsDG[3..6] |out-markdowntable @MDtbl|out-string).trim())" ;
                                $smsg += "`n$(($grp |select $propsDG[2] |fl |out-string).trim())" ;
                                $smsg += "`n#### Members:`n$(($grp.members | ft -a $propsRcpTbl|out-string).trim())`n`n" ;
                            } ; 
                        } ; 
                    } ; 
                    if($hSum.xoRecipientPermission){
                        $smsg += "`n## xoRecipientPermission:`n$(($hsum.xoRecipientPermission | ft -a $prpRPerms |out-string).trim())`n`n" ; 
                        if($hSum.xoRecipientPermissionGroups){
                            $smsg += "`n### Expanded Perm Group Summaries:" ; 
                            foreach($grp in $hSum.xoRecipientPermissionGroups){
                                $smsg += "`n-----------" ; 
                                $smsg += "`n$(($grp |select $propsDG[0..1] |out-markdowntable @MDtbl|out-string).trim())" ;
                                $smsg += "`n$(($grp |select $propsDG[3..6] |out-markdow ntable @MDtbl|out-string).trim())" ;
                                $smsg += "`n$(($grp |select $propsDG[2] |fl |out-string).trim())" ;
                                $smsg += "`n#### Members:`n$(($grp.members | ft -a $propsRcpTbl|out-string).trim())`n`n" ;
                            } ; 
                        } ; 
                    } ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level PROMPT } 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                } ;
                #endregion OUTPUT_PERMS ; #*------^ END OUTPUT_PERMS ^------
                #region OUTPUT_MOBILE ; #*------v OUTPUT_MOBILE v------
                if($getMobile){
                    write-host -foreground yellow "===`$hsum.xoMobileDeviceStats: " #-nonewline;

                    $ino = 0 ;
                    if($hsum.xoMobileDeviceStats){
                        foreach($xmob in $hsum.xoMobileDeviceStats){
                            $ino++ ;
                            if($hsum.xoMobileDeviceStats -is [system.array]){
                                    write-host -foreground yellow "=get-xMob$($ino):(ACTIVE)> " #-nonewline;
                            } else {
                                write-host -foreground yellow "=get-xMobileDev:(ACTIVE)> " #-nonewline;
                            } ;
                            $smsg = "$(($xmob | select $propsMobL1 |out-markdowntable @MDtbl |out-string).trim())" ;
                            $smsg += "`n$(($xmob | select $propsMobL2 |out-markdowntable @MDtbl |out-string).trim())" ;
                            write-host $smsg ;
                        } ;
                    } ; 
                    if($hsum.xoMobileDeviceStatsOLD){
                        #$smsg = "INACTIVE:(LastSyncAttemptTime -GT $($xoMobileDeviceOLDThreshold)d)" ; 
                        #write-host -foregroundcolor gray $smsg ;
                        foreach($xmob in $hsum.xoMobileDeviceStatsOLD){
                            $ino++ ;
                            if($hsum.xoMobileDeviceStatsOLD -is [system.array]){
                                    write-host -foreground yellow "=get-xMob$($ino):(inactive)> " #-nonewline;
                            } else {
                                write-host -foreground yellow "=get-xMobileDev:(inactive)> " #-nonewline;
                            } ;
                            $smsg = "$(($xmob | select $propsMobL1 |out-markdowntable @MDtbl |out-string).trim())" ;
                            $smsg += "`n$(($xmob | select $propsMobL2 |out-markdowntable @MDtbl |out-string).trim())" ;
                            write-host -foregroundcolor gray $smsg ;
                        } ;
                    } ; 
                    if($hsum.xoMobileOutlookClients){                        
                        $smsg = "+++Supported Outlook Mobile Clients: $($($hsum.xoMobileOutlookClients|measure).count)" ; 
                        <#
                        foreach($xmob in $hsum.xoMobileOutlookClients){
                            $ino++ ;
                            if($hsum.xoMobileDeviceStats -is [system.array]){
                                    write-host -foreground yellow "=get-xMob$($ino):> " #-nonewline;
                            } else {
                                write-host -foreground yellow "=get-xMobileDev:> " #-nonewline;
                            } ;
                            $smsg = "$(($xmob | select $propsMobL1 |out-markdowntable @MDtbl |out-string).trim())" ;
                            $smsg += "`n$(($xmob | select $propsMobL2 |out-markdowntable @MDtbl |out-string).trim())" ;
                            write-host $smsg ;
                        } ;
                        #>
                        $smsg += "`n$(($hsum.xoMobileOutlookClients| ?{$_.ClientType -eq 'EAS'} | ft -a $prpEASDevs|out-string).trim())" ; 
                        if($hsum.xoMobileOMSyncTypes){
                            $smsg += "`n-----`$hsum.xoMobileOMSyncTypes: $($hsum.xoMobileOMSyncTypes)" ; 
                            #write-host -foregroundcolor green $smsg ;
                            if($hsum.xoMobileOMSyncTypes -match 'REST'){
                                $smsg += "`n+User has one or more *legacy* 'REST' Outlook Mobile clients" ;
                            }elseif($hsum.xoMobileOMSyncTypes -match 'Outlook'){
                                $smsg += "`n++User has has one or more fully compliant 'MS Native Sync'-protocol Outlook Mobile clients" ;
                            } ;
                        } ; 
                        write-host -foregroundcolor green $smsg ;
                    }else{
                        write-verbose "(no Outlook Mobile clients returned)" ; 
                    } ; 
                    if($hsum.xoMobileOtherClients){
                        $smsg = "---NON-Outlook Mobile Clients:(device-vendor-supported): $($($hsum.xoMobileOtherClients|measure).count)" ; 
                        <#
                        foreach($xmob in $hsum.xoMobileOtherClients){
                            $ino++ ;
                            if($hsum.xoMobileDeviceStats -is [system.array]){
                                    write-host -foreground yellow "=get-xMob$($ino):> " #-nonewline;
                            } else {
                                write-host -foreground yellow "=get-xMobileDev:> " #-nonewline;
                            } ;
                            $smsg = "$(($xmob | select $propsMobL1 |out-markdowntable @MDtbl |out-string).trim())" ;
                            $smsg += "`n$(($xmob | select $propsMobL2 |out-markdowntable @MDtbl |out-string).trim())" ;
                            write-host $smsg ;
                        } ;
                        #>
                        $smsg += "`n$(($hsum.xoMobileOtherClients| ft -a $prpEASDevs|out-string).trim())" ;                         
                        write-host -foregroundcolor RED $smsg ;
                        if($hsum.xoMobileOMSyncTypes){
                            $smsg += "`n-----`$hsum.xoMobileOtherSyncTypes: $($hsum.xoMobileOtherSyncTypes)" ;
                            write-host -foregroundcolor yellow $smsg ;
                        }
                        if($hsum.xoMobileOtherClients| ?{$_.ClientType -eq 'EAS'}){ ;
                            $smsg = "`nThe following devices use device-vendor-provided/supported 'ExchangeActiveSync/EAS' Mobile clients!" ;
                            $smsg += "`nPLEASE NOTE: By policy EAS clients are *Best Effort* supported:"
                            $smsg += "`nWhere issues are experienced with legacy EAS/ActiveSync clients," ;
                            $smsg += "`nUsers should be urged to move to _Supported_ Microsoft Outlook Mobile for IOS or Android" ;
                            #$prpEASDevs = 'DeviceFriendlyName','ClientType','LastSyncAttemptTime','LastSuccessSync' ; 
                            $smsg += "`n$(($hsum.xoMobileOtherClients| ?{$_.ClientType -eq 'EAS'} | ft -a $prpEASDevs|out-string).trim())" ; 
                            write-host -foregroundcolor yellow $smsg ;
                        }
                    }else{
                        write-verbose "(no 'non'-Outlook Mobile clients returned)" ; 
                    } ; 

                } ;
                #endregion OUTPUT_MOBILE ; #*------^ END OUTPUT_MOBILE ^------
            } ;
            #endregion RV_VIA_GMGU ; #*------^ END RV_VIA_GMGU ^------
            
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
            #region OUTPUT_ACCT_DISABLED ; #*------v OUTPUT_ACCT_DISABLED v------
            if($hsum.ADUser){
                $hsum.IsADDisabled  +=  [boolean]($hsum.ADUser.Enabled -eq $true) ; 
             } else {
                write-verbose "(no ADUser found)" ;
            } ;
            if($hSum.MGUser){
                $hsum.IsMGDisabled  +=  [boolean]($hSum.MGUser.AccountEnabled -eq $true) ; 
                $hsum.isDirSynced  +=  [boolean]($hSum.MGUser.OnPremisesSyncEnabled  -eq $True)
            } else {
                write-verbose "(no AADUser found)" ;
            } ;
            # shift test to aadu
            if($hSum.MGUser){
                $hsum.IsLicensed  +=  [boolean]($hSum.MGUser.assignedlicenses.count -gt 0)
            } else {
                write-verbose "(no AADUser found)" ;
            } ;
            #endregion OUTPUT_ACCT_DISABLED ; #*------^ END OUTPUT_ACCT_DISABLED ^------

            #region ISSUE_DETECT ; #*------v ISSUE_DETECT v------

            #region SPLITBRAIN_NOBRAIN ; #*------v SPLITBRAIN_NOBRAIN v------
            # do a split-brain/nobrain check
            $smsg = "`n"
            if(($hsum.xoRcp.RecipientTypeDetails -match '(UserMailbox|MailUser)') -AND $hsum.IsLicensed -AND $hSum.xomailbox -AND $hSum.OPMailbox){
                #OPRcp, xorcp, OPMailbox, OPRemoteMailbox, xoMailbox
                $smsg += "SPLITBRAIN!:$($hSum.ADUser.userprincipalname).IsLic'd & has *BOTH* xoMbx & opMbx!" ;
                $hsum.IsSplitBrain  +=  $true ;
            }elseif(($hsum.xoRcp.RecipientTypeDetails -match '(UserMailbox|MailUser)') -AND -not($hsum.IsLicensed) -AND $hSum.xomailbox -AND $hSum.OPMailbox){
                #OPRcp, xorcp, OPMailbox, OPRemoteMailbox, xoMailbox#
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

            # conditional w-w, w-h block on status
            #if($hsum.IsSplitBrain -OR $hsum.IsNoBrain -OR (-not $hsum.IsLicensed -AND $hsum.xoRcp.RecipientTypeDetails -NOTmatch 'SharedMailbox|RoomMailbox|EquipmentMailbox') ){
            [boolean[]]$testArray = @(
                $hsum.IsSplitBrain, 
                $hsum.IsNoBrain, 
                (-not $hsum.IsLicensed -AND $hsum.xoRcp.RecipientTypeDetails -NOTmatch 'SharedMailbox|RoomMailbox|EquipmentMailbox')
            ) ; 
            #write-verbose "Test: All `$true" ; 
            #if(($testArray | Where-Object {$_ -eq $true}).Count -eq $testArray.Count){
            #write-verbose "Test: Count `$true meets threshold" ; 
            #$tTrues = $testArray.count -3 ; # test is 3 less than total elem count
            #if(($testArray | Where-Object {$_ -eq $true}).Count -ge $tTrues){write-host "test:$($tTrues)/$($testArray.count) `$true: PASS" } ; 
            write-verbose "Test: Any `$true above" ; 
            # the $smsg is populated further up, this is just the output format on the $smsg text
            if($testArray -contains $true){
                # w-w
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
            } else { 
                # w-h
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
            } ;  
            #endregion SPLITBRAIN_NOBRAIN ; #*------^ END SPLITBRAIN_NOBRAIN ^------

            #region NOBRAIN_DETAILS ; #*------v NOBRAIN_DETAILS v------
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
                    if($hsum.ADUser.Info){
                        $smsg += "`n--ADUser.Info:$($hsum.ADUser.Info)" ;
                    }
                    if($hsum.IsADDisabled){
                        $smsg += "`n--ADUser:$($hsum.ADUser.samaccountname) is *DISABLED* for logon (likely TERM)" ;
                    } else {
                        $smsg += "`n--ADUser:$($hsum.ADUser.samaccountname) is *UN-DISABLED* for logon (improperly offboarded TERM?)" ;
                    } ;
                } else {
                    write-verbose "(no ADUser found)" ;
                } ;
                if($hsum.IsMGDisabled){
                    $smsg += "`n--AADUser:$($hSum.MGUser.UserPrincipalName) is *DISABLED* for logon (likely TERM)" ;
                } else {
                    $smsg += "`n--AADUser:$($hSum.MGUser.UserPrincipalName) is *UN-DISABLED* for logon (improperly offboarded TERM?)" ;
                } ;
                $smsg += "`n"
                write-warning $smsg ;
            } ;
            #endregion NOBRAIN_DETAILS ; #*------^ END NOBRAIN_DETAILS ^------

            #region RMBX_BLOCKED_XOMBX ; #*------v RMBX_BLOCKED_XOMBX v------
            # 2:34 PM 1/9/2025 test for 886258, blocked license-xoMailbox mount issue
            [boolean[]]$testArray = @(
                ($hsum.oprcp.recipienttypedetails -eq 'RemoteUserMailbox'),
                ($hsum.xorcp.recipienttypedetails -eq 'Mailuser'),
                (-not $hsum.xoMailbox),
                $hSum.MGUser,
                $hsum.ADUser,
                $hsum.isDirSynced,
                ($hsum.IsNoBrain -eq 1),
                (-not $hsum.IsLicensed),
                $hsum.opRemoteMailbox.exchangeguid,
                $hsum.opRemoteMailbox.remoteroutingaddress    
            ) ;  
            # test variants: eval patterns of $true/$false
            #write-verbose "Test: Count `$true meets threshold" ;
            #$nTrues = $testArray.count -3 ; # test is 3 less than total elem count
            #if(($testArray | Where-Object {$_ -eq $true}).Count -ge $nTrues){write-verbose "test:$($nTrues)/$($testArray.count) `$true: PASS" } ;
            #write-verbose "Test: Any `$true above" ;
            #if($testArray -contains $true){ # -OR clause
            write-verbose "Test: All `$true" ; # -AND clause
            if(($testArray | Where-Object {$_ -eq $true}).Count -eq $testArray.Count){
                $hsAlertMsg = @"
User has:
- OPRmbx and no XoMailbox!
- Dirsynced AADUser & ADUser
- Detects as NoBrain (neither OP or xo Mailbox)
- is not Licensed
- And Rmbx has populated ExchangeGuid & RemoteRoutingAddress
(against Mailbox that doesn't currently *exist*
with email address that also doesn't currently *exist)
If this matches Incident # 886258:
- if Licensed, the xoMailbox will never mount
    sits 18h+ in status: `"We are preparing a mailbox for the user.`"
- Fix: in that condition is to:
    1. *Remove* the OpRemoteMailbox
    2. Permit ADC replication to replicate, and wait for xoMailbox to mount
    3. Create a new matched OpRmbx with the RemoteMountingAddress and xoMailbox.ExchangeGuid, copied to the OpRmbx.Exchangeguid
    4. Verify if any CA5 setting is missing/required to properly steer primarysmtpaddress
## Detailed status:
### get-RemoteMailbox:
$(($hsum.opRemoteMailbox | fl 'Name','RecipientTypeDetails','RemoteRecipientType','exchangeguid','PrimarySmtpAddress','RemoteRoutingAddress' | fl |out-string).trim())
### Cloud: get-xoRecipient:
$(($hsum.xorcp | fl 'RecipientType','RecipientTypeDetails','PrimarySmtpAddress','Alias' |out-string).trim())
- SMTP EmailAddresses:
$(($hsum.xorcp | select -expand emailaddresses | ?{$_ -match 'smtp:'} | sort |out-string).trim())
### DirSync-settings:
opRemoteMailbox.exchangeguid.guid:`t$($hsum.opRemoteMailbox.exchangeguid.guid)
opRemoteMailbox.RemoteRoutingAddress:`t$($hsum.opRemoteMailbox.RemoteRoutingAddress.guid)
"@ ;        
                $smsg = $hsAlertMsg ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;  
            }
            #endregion RMBX_BLOCKED_XOMBX ; #*------^ END RMBX_BLOCKED_XOMBX ^------
            #endregion ISSUE_DETECT ; #*------^ END ISSUE_DETECT ^------

            #region WRITE_OUTPUT ; #*------v WRITE_OUTPUT v------
            # 4:11 PM 12/11/2025 with move to MG module for user access, I've wound up with massive export .xml files, that were fine pulling from azuread
            <# even running simple export depth tesets, shows broken xml: 
                $xmlString = $Rpt | ConvertTo-Xml -Depth $dpth -NoTypeInformation ; 
                ConvertTo-Xml : Unexpected end of file has occurred. The following elements are not closed: Property, Property, Object, Objects. Line 10, position 47281.
                At line:1 char:21

                A recommendation is to simplify the object
                - for MGUser, force select just the targete props (again)
                $rpt.mguser = $rpt.mguser | select $prpMGUser
                - for the MGUserMgr hashtable/dict, convert it to pscustomobject
                $rpt.MGUserMgr = [pscustomobject]$rpt.MGUserMgr ;
            #>
            write-verbose "MGUser massive xml export impact: pre strip the objects";
            $hSum.mguser = $hSum.mguser | select $prpMGUser
            $hSum.MGUserMgr = [pscustomobject]$hSum.MGUserMgr
            write-verbose 'Test here'
            if($outObject){
                if($PSCmdlet.MyInvocation.ExpectingInput){
                    write-verbose "(pipeline input, skipping aggregator, dropping into pipeline)" ;
                    New-Object PSObject -Property $hSum | write-output  ;
                } else {
                    $Rpt += New-Object PSObject -Property $hSum ;
                } ;
            } ELSE {
                # 3:59 PM 9/18/2023 else export to report file 
                <#
                $Rpt += New-Object PSObject -Property $hSum ;
                #$Rpt | export-clixml -Path $ofile -Depth 100 ;
                write-verbose "exporting to $($logfile.replace('LOG-BATCH-EXEC',"REPORT").replace('-log',''))"
                $Rpt | export-clixml -Path $logfile.replace('LOG-BATCH-EXEC',"REPORT").replace('-log','') -Depth 100 ;
                #>
                $Rpt += New-Object PSObject -Property $hSum ;
                if($useJson){                    
                    TRY{
                        $ofile = $logfile.replace('LOG-BATCH-EXEC',"REPORT").replace('-log','').replace('.txt','.json')
                        $smsg = "exporting to:`n$($ofile)" ; 
                        $smsg += "-useJson: exporting in .json format " ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                        $Rpt | ConvertTo-Json -depth $Depth -ea STOP | 
                            Out-String | Out-File -Encoding utf8 -LiteralPath $ofile -ErrorAction STOP -verbose ; 
                        <# generating:
                            ConvertFrom-Json : Cannot convert the JSON string because a dictionary that was converted from the string contains the duplicated keys 'value' and 'Value'.

                            ...occurs because PowerShell objects and hashtables are **case-insensitive** for property names, while the JSON standard allows keys to be case-sensitive. 
                            The best solutions involve modifying the JSON string _before_ conversion or using an alternative parsing method. 
                        #>
                        # so try 2-staging it, and sub 'value' -> 'Value'?

                    } CATCH {
                        $ErrTrapd=$Error[0] ;
                        write-host -foregroundcolor gray "TargetCatch:} CATCH [$($ErrTrapd.Exception.GetType().FullName)] {"  ;
                        $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                        write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" ;
                     } ;
                    
                }else{
                    $ofile = $logfile.replace('LOG-BATCH-EXEC',"REPORT").replace('-log','').replace('.txt','.xml')
                    $smsg = "exporting to:`n$($ofile)" ;
                    write-host $smsg ; 
                    TRY{
                        # try to pre-detect issues in the source Microsoft.Graph data, by preconverting to xml
                        $xmlString = $Rpt | ConvertTo-Xml -Depth $Depth -NoTypeInformation -ea STOP;
                        write-verbose "`$xmlString.length: $(($xmlString.length|out-string).trim())" ;
                        $xmlString = $NULL ; 
                        #$Rpt | export-clixml -Path $ofile -Depth 100 ;
                        write-verbose "exporting to $($ofile)"
                        $Rpt | export-clixml -Path $ofile -Depth $Depth -ea STOP -verbose ;
                    } CATCH {
                        $ErrTrapd=$Error[0] ;
                        $smsg = "XML conversion issue: failed attempt:`$Rpt | ConvertTo-Xml -Depth 1" ;
                        $smsg += "`nABORTING XML OUTPUT (STATUS RESULTS IN SEMI-HUNG EXPORTS & MULTI-GB XML FILES!)" ;
                        $smsg += "`nPossibly rerun with -outobject; capture return, and attempt to manually convert/export/workwith wo export" ;
                        $smsg += "`nOr try use of -useJson, and play with the -Depth parameter to tune minimum functional export depth" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                        if($ErrTrapd){
                            write-host -foregroundcolor gray "TargetCatch:} CATCH [$($ErrTrapd.Exception.GetType().FullName)] {"  ;
                            $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                            write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)"    
                        } ; 
                    } ;
                }
            } ;
            #endregion WRITE_OUTPUT ; #*------^ END WRITE_OUTPUT ^------
            # 10:34 AM 12/22/2025 missing per-user log stop
            $stopResults = try {Stop-transcript -ErrorAction stop} catch {} ;
            if($stopResults){
                $smsg = "Stop-transcript:$($stopResults)" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } ; 
                        
            write-host -foregroundcolor green $sBnr.replace('=v','=^').replace('v=','^=') ;

        } ; # loop-E $users
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
#endregion RESOLVE_USER ; #*------^ END resolve-user ^------