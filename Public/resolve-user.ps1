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
    .PARAMETER  users
    Array of user descriptors: displayname, emailaddress, UPN, samaccountname (checks clipboard where unspecified)
    .PARAMETER Ticket
    Ticket number[-ticket 123456]
    .PARAMETER getMobile
    switch to return mobiledevice info for target XO Mailbox (not supported for onprem mailboxes)[-getMobile]
    .PARAMETER getQuotaUsage
    switch to return Quota & MailboxFolderStatistics analysis (XO-only)[-getQuotaUsage]
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
    Default, attempts to parse a user descriptor from clipboard
    .EXAMPLE
    PS> resolve-user -users 'John Public'
    Process user displayname
    .EXAMPLE
    PS> resolve-user -users 'Test@domain.com','User Name','Alias','ExternalContact@emaildomain.com','confroom@tenant.onmicrosoft.com' -verbose ;
    Process an array of descriptors
    .EXAMPLE
    PS> $results = resolve-user -outobject -users 'Test@domain.com','John Public','Alias','ExternalContact@emaildomain.com','confroom@tenant.onmicrosoft.com''  ;
    $feds = $results| group federator | select -expand name ;
    # echo filtered subsets
    ($results| ?{$_.federator -eq $feds[1] }).xomailbox
    ($results| ?{$_.federator -eq $feds[1] }).xomailbox.primarysmtpaddress
    # profile results
    $analysis = foreach ($data in $resolved_objects){
        $Rpt = [ordered]@{
            PrimarySmtpAddress = $data.xorcp.primarysmtpaddress ; 
            ADUser_UPN = $data.aduser.userprincipalname ; 
            AADUser_UPN = $data.aaduser.UserPrincipalName ; 
            isDirSynced = $data.isDirSynced ; 
            IsNoBrain = $data.IsNoBrain ; 
            isSplitBrain = $data.isSplitBrain;
            IsLicensed = $data.IsLicensed;
            IsDisabledOU = $data.IsDisabledOU;
            IsADDisabled = $data.IsADDisabled; 
            IsAADDisabled = $data.IsAADDisabled;
        } ; 
        [pscustomobject]$Rpt ; 
    } ; 
    # output tabular results
    $analysis | ft -auto ; 
    - Process array of users, specify return detailed object (-outobject), for post-processing & filtering,
    - Group results on federation sources,
    - Output summary of EXO mailboxes for the second federator
    - Then output the primary smtpaddress for all EXO mailboxes resolved to that federator
    - Then create a summary object of the is* properties and UPN, primarySmtpAddress, 
    - Finally display the summary as a console table
    .EXAMPLE
    $rptNNNNNN_FName_LName_Domain_com = ulu -o -users 'FName.LName@Domain.com' ;  $rpt655692_FName_LName_Domain_com | xxml .\logs\rpt655692_FName_LName_Domain_com.xml
    Example (from ahk 7uluo! macro parser output) that creates a variable based on ticketnumber & email address (with underscores for alphanums), from the output, and then exports the variable content to xml. 
    ves an immediately parsable inmem variable, along with the canned .xml that can be reloaded in future, or attached to a ticket.
    .EXAMPLE
    PS> resolve-user -users 'John Public' -getmobile
    Example that includes the -getMobile parameter, to return details on xo MobileDevices in use with the EXO mailbox
    .EXAMPLE
    PS> resolve-user -users 'John Public' -getQuotaUsage
    Example that includes the -getQuotaUsage parameter, to return details on xo MailboxFolderStatistics and effective Quota
    .LINK
    https://github.com/tostka/verb-exo
    #>
    ###Requires -Version 5
    ##Requires -Modules ActiveDirectory, MSOnline, AzureAD, ExchangeOnlineManagement, verb-AAD, verb-ADMS, verb-Ex2010
    # 2:49 PM 3/8/2022 pull verb-ex2010 ref - I think it's generating nested errors, when ex2010 requires exo requires ex2010 == loop.
    # 12:19 PM 9/26/2023 pull verb-exo ref "
    #Requires -Modules ActiveDirectory, MSOnline, AzureAD, ExchangeOnlineManagement, verb-AAD, verb-ADMS, verb-Auth, verb-IO, verb-logging
    #Requires -RunasAdministrator
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
        [Parameter(HelpMessage="switch to return Quota & MailboxFolderStatistics analysis (XO-only)[-getQuotaUsage]")]
            [switch]$getQuotaUsage,
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
            [string[]]$UserRole = @('SID','CSVC'),
            # svcAcct use: @('ESvcCBA','CSvcCBA','SIDCBA')
        [Parameter(HelpMessage="Use EXOv2 (ExchangeOnlineManagement) over basic auth legacy connection [-useEXOv2]")]
            [switch] $useEXOv2=$true,
        [Parameter(HelpMessage="Silent output (suppress status echos)[-silent]")]
            [switch] $silent,
        [Parameter(HelpMessage="switch to return a system.object summary to the pipeline[-outObject]")]
            [switch] $outObject
    ) ;
    BEGIN{
        # for scripts wo support, can use regions to fake BEGIN;PROCESS;END:
        # ps1 faked:#region BEGIN ; #*------v BEGIN v------
        #region CONSTANTS-AND-ENVIRO #*======v CONSTANTS-AND-ENVIRO v======
        # function self-name (equiv to script's: $MyInvocation.MyCommand.Path) ;
        ${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name ;
        $PSParameters = New-Object -TypeName PSObject -Property $PSBoundParameters ;
        write-verbose "`$PSBoundParameters:`n$(($PSBoundParameters|out-string).trim())" ;
        $Verbose = ($VerbosePreference -eq 'Continue') ; 
        $rgxEmailAddr = "^([0-9a-zA-Z]+[-._+&'])*[0-9a-zA-Z]+@([-0-9a-zA-Z]+[.])+[a-zA-Z]{2,63}$" ;
        # added support for . fname lname delimiter (supports pasted in dirname of email addresses, as user)
        $rgxDName = "^([a-zA-Z]{2,}(\s|\.)[a-zA-Z]{1,}'?-?[a-zA-Z]{2,}\s?([a-zA-Z]{1,})?)" ;
        #"^([a-zA-Z]{2,}\s[a-zA-Z]{1,}'?-?[a-zA-Z]{2,}\s?([a-zA-Z]{1,})?)" ;
        $rgxObjNameNewHires = "^([a-zA-Z]{2,}(\s|\.)[a-zA-Z]{1,}'?-?[a-zA-Z]{2,}\s?([a-zA-Z]{1,})?)_[a-z0-9]{10}"  # Name:Fname LName_f4feebafdb (appending uniqueness guid chunk)
        $rgxSamAcctNameTOR = "^\w{2,20}$" ; # up to 20k, the limit prior to win2k
        #$rgxSamAcctName = "^[^\/\\\[\]:;|=,+?<>@?]+$" # no char limit ;
        $MaxRecips = 25 ; # max number of objects to permit on a return resultsize/,ResultSetSize, to prevent empty set return of everything in the addressspace

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

            $prpMbxHold = 'LitigationHoldEnabled',@{n="InPlaceHolds";e={ ($_.InPlaceHolds | select -expand InPlaceHolds) -join ', '}},
                'ComplianceTagHoldApplied','DelayHoldApplied','DelayReleaseHoldApplied' ; 

            $rgxHiddn = '.*\\(Versions|SubstrateHolds|DiscoveryHolds|Yammer.*|Social\sActivity\sNotifications|Suggested\sContacts|Recipient\sCache|PersonMetadata|Audits|Calendar\sLogging|Purges)$' ; 

        } ; 
        $rgxOPLic = '^CN\=ENT\-APP\-Office365\-(EXOK|F1|MF1)-DL$' ;
        $rgxXLic = '^CN\=ENT\-APP\-Office365\-(EXOK|F1|MF1)-DL$' ;

        <# if we want pipeline to work have to move the clipboard grab out or down into process{}, where pipeline binding will be actually populated
        if(!$users){
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
        #>

        # pre psv2, no $PSBoundParameters autovari to check, so back them out:
        write-verbose 'Collect all non-default Params (works back to psv2 w CmdletBinding)'
        $ParamsNonDefault = (Get-Command $PSCmdlet.MyInvocation.InvocationName).parameters | Select-Object -expand keys | Where-Object{$_ -notmatch '(Verbose|Debug|ErrorAction|WarningAction|ErrorVariable|WarningVariable|OutVariable|OutBuffer)'} ;

        #if ($PSScriptRoot -eq "") {
        if( -not (get-variable -name PSScriptRoot -ea 0) -OR ($PSScriptRoot -eq '')){
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

        # 1:00 PM 9/18/2023 splice in modern svcconns
        #region SERVICE_CONNECTIONS #*======v SERVICE_CONNECTIONS v======
        # PRETUNE STEERING separately *before* pasting in balance of region
        #*------v STEERING VARIS v------
        $useO365 = $true ;
        $useEXO = $true ; 
        $UseOP=$true  ; 
        $UseExOP=$true ;
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
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Prompt } 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            $TenOrg = get-TenantTag -Credential $Credential ;
        } else { 
            # if not using Credentials or a TargetTenants/TenOrg loop, default the $TenOrg on the $env:USERDOMAIN
            $smsg = "(unconfigured `$TenOrg & *NO* `$Credential: fallback asserting from `$env:USERDOMAIN)" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Prompt } 
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
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
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
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    $pltGTCred.UserRole = $UserRole; 
                } else { 
                    $smsg = "(No `$UserRole found, defaulting to:'CSVC','SID' " ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
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
                #$smsg = "(Resolved credentials CredType/UserRole as: $($o365Cred.credType)" ; 
                $smsg = "(validated `$o365Cred contains .credType:$($o365Cred.credType) & `$o365Cred.Cred.username:$($o365Cred.Cred.username)" ; 
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
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
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } else {
                $statusdelta = ";ERROR";
                $script:PassStatus += $statusdelta ;
                set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatettus_$($tenorg)).value + $statusdelta) ;
                $smsg = "Unable to resolve $($tenorg) `$o365Cred value!"
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                throw "Unable to resolve $($tenorg) `$o365Cred value!`nEXIT!"
                Break ;
            } ;
            # configure splat for connections: (see above useage)
            # downstream commands
            $pltRXO = [ordered]@{
                Credential = $Credential ;
                verbose = $($VerbosePreference -eq "Continue")  ;
            } ;
            if((gcm Reconnect-EXO).Parameters.keys -contains 'silent'){
                $pltRxo.add('Silent',$silent) ;
            } ;
            # default connectivity cmds - force silent false
            $pltRXOC = [ordered]@{} ; $pltRXO.GetEnumerator() | ?{ $_.Key -notmatch 'silent' }  | ForEach-Object { $pltRXOC.Add($_.Key, $_.Value) } ; $pltRXOC.Add('silent',$true) ; 
            if((gcm Reconnect-EXO).Parameters.keys -notcontains 'silent'){
                $pltRxo.remove('Silent') ;
            } ; 
            #region EOMREV ; #*------v EOMREV Check v------
            #$EOMmodname = 'ExchangeOnlineManagement' ;
            $pltIMod = @{Name = $EOMmodname ; ErrorAction = 'Stop' ; verbose=$false} ;
            # do a gmo first, faster than gmo -list
            if([version]$EOMMv = (Get-Module @pltIMod| sort version | select -last 1).version){}
            elseif([version]$EOMMv = (Get-Module -ListAvailable @pltIMod| sort version | select -last 1).version){}
            else {
                $smsg = "$($EOMmodname) PowerShell v$($MinNoWinRMVersion) module is required, do you want to install it?" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Prompt }
                else{ write-host -foregroundcolor YELLOW "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
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
            if((gcm Reconnect-EXO).Parameters.keys -contains 'silent'){
                $pltRxo.add('Silent',$false) ;
            } ; 
            # default connectivity cmds - force silent false
            $pltRXOC = [ordered]@{} ; $pltRXO.GetEnumerator() | ?{ $_.Key -notmatch 'silent' }  | ForEach-Object { $pltRXOC.Add($_.Key, $_.Value) } ; $pltRXOC.Add('silent',$true) ;
            if((gcm Reconnect-EXO).Parameters.keys -notcontains 'silent'){
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
                $statusdelta = ";ERROR";
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
            if((gcm Reconnect-Ex2010).Parameters.keys -contains 'silent'){
                $pltRX10.add('Silent',$false) ;
            } ;

            # defer cx10/rx10, until just before get-recipients qry
            #endregion GENERIC_EXOP_CREDS_&_SRVR_CONN #*------^ END GENERIC EXOP CREDS & SRVR CONN BP ^------
            # connect to ExOP X10
            if($useEXOP){
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
            } else { 
            
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
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    $smsg = 'Set-AdServerSettings -ViewEntireForest `$True' ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
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
            #$domaincontroller = get-GCFastXO -TenOrg $TenOrg -subdomain ((gv -name "$($TenOrg)Meta").value['OP_ExADRoot']) -verbose:$($verbose) |?{$_.length};
            # need to debug the above, credential issue?
            # just get it done
            $domaincontroller = get-GCFast
        }  else { 
            # have to defer to get-azuread, or use EXO's native cmds to poll grp members
            # TODO 1/15/2021
            $useEXOforGroups = $true ; 
            $smsg = "$($TenOrg):HAS NO ON-PREM ACTIVEDIRECTORY, DEFERRING ALL GROUP ACCESS & MGMT TO NATIVE EXO CMDS!" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        } ; 
        if($useForestWide -AND -not $GcFwide){
            #region  ; #*------v OPTIONAL CODE TO ENABLE FOREST-WIDE ACTIVEDIRECTORY SUPPORT: v------
            $smsg = "`$GcFwide = Get-ADDomainController -Discover -Service GlobalCatalog" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            $GcFwide = "$((Get-ADDomainController -Discover -Service GlobalCatalog).hostname):3268" ;
            $smsg = "Discovered `$GcFwide:$($GcFwide)" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
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
        # ps1 faked: #region PROCESS ; #*------v PROCESS:$($identity -join ',') v------

        #$dname= 'Todd Kadrie' ;
        #$dname = 'Stacy Sotelo'

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
        # with a process {} block to handle any pipeline passed input. The pipeline still maps to the bound param: $users, but the e3ntire process{} is run per element, rather than iteratign the internal $users foreach.
        foreach ($usr in $users){

            #region TRANSCRIPTNAME ; #*------v TRANSCRIPT FROM SCRIPT/MODULENAME v------
            # ISSUES with use of start-log(), even local, deps on Remove-StringDiacritic() etc, that are in dep modules, so alt: (see 7psOFile for other alts)
            $transcript = "$($ScriptDir)\logs" ; if(!(test-path -path $transcript)){ mkdir $transcript -verbose:$true ; } ;
            # - FOR FUNCTIONS, build from ${CmdletName}
            #$transcript +=  "\$($CmdletName)" ;
            # - OR FOR SCRIPTS, use $ScriptBaseName/$ScriptNameNoExt (which reflect name of hosting .psm1/.ps1 a function was loaded from)
            $transcript +=  "\$($ScriptNameNoExt)" ;
            $logfile = $transcript.replace('-trans','-log') ; $logging = $true ; 
            #endregion TRANSCRIPTNAME ; #*------^ END TRANSCRIPT FROM SCRIPT/MODULENAME ^------
            #region LOGBUILD ; #*------v LOGBUILD v------
            # building an outputfile name dynamically using paremeters
            #-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
            # start by rebuilding from base of start-log(): $logfile: 'D:\scripts\logs\get-ExOPSmtpReceiveTLSReport-SERVER50'
            # *first* reset $ofile; it's picking the filename up from the OS
            $ofile = $null;
            [array]$ofile = $ofile ;
            #$ofile += ("$($_.Name)-$($_.value)" -join ',' );
            #$ofile=".\$($ticket)-$($usr)-folder-sizes-NONHIDDEN-NONZERO-$(get-date -format 'yyyyMMdd-HHmmtt').csv" ;
            $ofile += @("$($ticket)-$($usr)") ; 
            $ofile += "REPORT" ; 
            $ofile += "run$(get-date -format 'yyyyMMdd-HHmmtt')" ; 
            [string]$ofile = $ofile -join '-' ; 
            [string]$ofile += ".xml" ; 
            # split existing logfile path out: e.g. D:\scripts\logs\update-xoRetiringConfRmWindows-(TOR)-LASTPASS
            #[string]$ofile = join-path -path ($logfile -split '-LOG-BATCH-')[0]  -childpath $ofile ; 
            # raw path, no cmdletname prefix
            [string]$ofile = join-path -path (split-path $logfile) -childpath $ofile ; 

            
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

            switch -regex ($usr){
                $rgxEmailAddr {
                    $hSum.fname,$hSum.lname = $usr.split('@')[0].split('.') ;
                    $hSum.dname = $usr ;
                    write-verbose "(detected user ($($usr)) as EmailAddr)" ;
                    $isEml = $true ;
                    Break ;
                }
                $rgxObjNameNewHires{
                    write-verbose "(detected user ($($usr)) as ObjNameNewHires)" ;
                    $hSum.fname,$hSum.lname = $usr.split('_')[0].split(' ');
                    $hSum.dname = $usr.split('_')[0] ;
                    write-verbose "(detected user ($($usr)) as DisplayName)" ;
                    $isObjName = $true ;
                    Break ;
                }
                $rgxDName {
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
                $rgxSamAcctNameTOR {
                    $hSum.lname = $usr ;
                    write-verbose "(detected user ($($usr)) as SamAccountName)" ;
                    $isSamAcct  = $true ;
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
                $fltr = "name -like '$usr'" ;
                write-verbose "processing:'filter':$($fltr)" ;
                $pltGMailObj.add('filter',$fltr) ;
            } ;
            if($isDname){
                # interestinb bug: switched to $hSum.dname: ISE is fine, but ConsoleHost fails to expand the $fltr properly.
                # standard is: Variables: Enclose variables that need to be expanded in single quotation marks (for example, '$User'). Don't use curly-brackets (impedes expansion)
                # workaround: looks like have to proxy the $hsum.Dname, to provide a single non-dotted variable name
                $dname = $hSum.dname
                $fltr = "displayname -like '$dname'" ;
                write-verbose "processing:'filter':$($fltr)" ;
                $pltGMailObj.add('filter',$fltr) ;
            } ;

            $error.clear() ;

            #write-verbose "get-[exo]Recipient w`n$(($pltGMailObj|out-string).trim())" ;
            #write-verbose "get-recipient w`n$(($pltGMailObj|out-string).trim())" ;
            # exclude contacts, they don't represent real onprem mbx assoc, and we need to refer those to EXO mbx qry anyway.
            write-verbose "get-recipient w`n$(($pltGMailObj|out-string).trim())" ;
            rx10 -Verbose:$false -silent ;
            if($hSum.OPRcp=get-recipient @pltGMailObj -ea 0 | select -first $MaxRecips | ?{$_.recipienttypedetails -ne 'MailContact'}){
                write-verbose "`$hSum.OPRcp found" ;
                switch ($hSum.OPRcp.recipienttypedetails){
                    'RemoteUserMailbox' {write-host "(Rmbx)" -nonewline}
                    'UserMailbox' {write-host "(Mbx)" -nonewline}
                    # no rmbx, but remote obj?
                    'MailUser' {
                        $smsg = "MAILUSER WO RMBX DETECTED! - POSSIBLE NOBRAIN?"
                        write-warning $smsg
                        #$hsum.isNoBrain = $true ;
                    }
                    'MailUniversalDistributionGroup' {write-host "(DG)" -nonewline}
                    'DynamicDistributionGroup'  {write-host "(DDG)" -nonewline}
                    'MailContact' {write-host "(MC)" -nonewline]}
                    default{}
                }
            } elseif($isDname -and $hsum.lname) {
                $smsg = "Failed:RETRY: detected 'LName':$($hsum.lname) for near matches..." ;
                write-host $smsg ;
                $lname = $hsum.lname ;
                $fltrB = "displayname -like '*$lname*'" ;
                write-verbose "RETRY:get-recipient -filter {$($fltr)}" ;
                if($hSum.OPRcp=get-recipient -filter $fltr -ea 0 | select -first $MaxRecips |?{$_.recipienttypedetails -ne 'MailContact'}){
                    write-verbose "`$hSum.OPRcp found" ;
                } ;
            };

            if(!$hsum.OpRcp){
                $smsg = "(Failed to OP:get-recipient on:$($usr))"
                if($isDname){$smsg += " or *$($hsum.lname )*"}
                write-host $smsg ;
            } else {
                write-verbose "`$hSum.OPRcp:`n$(($hSum.OPRcp|out-string).trim())" ;
            } ;


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
            if(!$hsum.xoRcp){
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

            if($hSum.OPRcp){
                $error.clear() ;
                TRY {
                    switch -regex ($hSum.OPRcp.recipienttype){
                        "UserMailbox" {
                            write-verbose "'UserMailbox':get-mailbox $($hSum.OPRcp.identity)"
                            if($hSum.OPMailbox=get-mailbox $hSum.OPRcp.identity -resultsize $MaxRecips | select -first $MaxRecips ){ ;
                                #write-verbose "`$hSum.OPMailbox:`n$(($hSum.OPMailbox|out-string).trim())" ;
                                if($outObject){

                                } else {
                                    $Rpt += $hSum.OPMailbox.primarysmtpaddress ;
                                } ;
                                write-verbose "'UserMailbox':Test-MAPIConnectivity -identity $($hSum.OPMailbox.userprincipalname)"
                                $hSum.OPMapiTest = Test-MAPIConnectivity -identity $hSum.OPMailbox.userprincipalname ;
                                $smsg = "Outlook (MAPI) Access Test Result:$($hsum.OPMapiTest.result)" ;
                                if($hsum.OPMapiTest.result -eq 'Success'){
                                    write-host -foregroundcolor green $smsg ;
                                } else {
                                    write-WARNING $smsg ;
                                } ;
                            } ;
                        }
                        "MailUser" {
                            write-verbose "'MailUser':get-remotemailbox $($hSum.OPRcp.identity)"
                            if($hSum.OPRemoteMailbox=get-remotemailbox $hSum.OPRcp.identity -resultsize $MaxRecips -ea 0 | select -first $MaxRecips){
                                write-verbose "`$hSum.OPRemoteMailbox:`n$(($hSum.OPRemoteMailbox|out-string).trim())" ;
                            }else{
                                $smsg = "RecipientTypeDetails:MailUser with NO Rmbx! (NoBrain?)" ;
                                write-warning $smsg ;
                                if($hsum.xoRcp.ExternalDirectoryObjectId){
                                    # of course has match to AADU  - always does - we're going to need the AADU before we can lookup the ADU
                                    # $pltGadu.identity = $hSum.AADUser.ImmutableId | convert-ImmuntableIDToGUID | select -expand guid ;
                                    caad  -Verbose:$false -silent ;
                                    write-verbose "OPRcp:Mailuser, ensure GET-ADUSER pulls AADUser.matched object for cloud recipient:`nfallback:get-AzureAdUser  -objectid $($hsum.xoRcp.ExternalDirectoryObjectId)" ;
                                    # have to postfilter, if want specific count -maxresults catch's with no $error[0]
                                    $hSum.AADUser  = get-AzureAdUser  -objectid $hsum.xoRcp.ExternalDirectoryObjectId | select -first $MaxRecips;  ;
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
                            write-warning "$((get-date).ToString('HH:mm:ss')):Unsupported RecipientType:($hSum.OPRcp.recipienttype). EXITING!" ;
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
                        #$hSum.ADUser =Get-ADUser @pltGadu ;
                        write-verbose "Get-ADUser w`n$(($pltGadu|out-string).trim())" ;
                        # try a nested local trycatch, against a missing result
                        Try {
                            #Get-ADUser $DN -ErrorAction Stop ;
                            $hSum.ADUser =Get-ADUser @pltGadu | select -first $MaxRecips ;
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
            }elseif($hSum.xoRcp){
                foreach($txR in $hSum.xoRcp){
                    TRY {
                        switch -regex ($txR.recipienttypedetails){
                            "UserMailbox" {
                                #write-verbose "$((get-alias ps1GetxMbx).definition) w`n$(($pltGMailObj|out-string).trim())" ;
                                write-verbose "get-exomailbox w`n$(($pltGMailObj|out-string).trim())" ;
                                if($hSum.xoMailbox=get-xomailbox @pltGMailObj -ea 0 | select -first $MaxRecips ){
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
                                        $hSum.xoMapiTest = Test-xoMAPIConnectivity -identity $xmbx.userprincipalname ;
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
                                            $hsum.xoMobileDeviceStats = Get-xoMobileDeviceStatistics -Mailbox $xmbx.userprincipalname -ea STOP ; 
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
                                            $hSum.xoMailboxStats = Get-xoMailboxStatistics @pltGMbxStatX | select $prpStat;
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
                                            $hSum.xoNetOfSendReceiveQuotaMB = $hSum.xoEffectiveQuotas.ProhibitSendQuotaMB - $hSum.xoMailboxStats.TotalMailboxSizeMB ; 

                                            $pltGMbxStatX.add('IncludeOldestAndNewestItems',$true) ; 
                                            $smsg = "Get-xoMailboxFolderStatistics  w`n$(($pltGMbxStatX|out-string).trim())" ;
                                            if($verbose){
                                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                                                else{ write-verbose $smsg } ; 
                                            } ; 
                                            TRY{
                                                $hsum.xoMailboxFolderStats = Get-xoMailboxFolderStatistics @pltGMbxStatX  ;

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
                                #$hSum.OPRemoteMailbox=get-remotemailbox $txR.identity  ;
                                caad -silent -verbose:$false ;
                                #write-verbose "`$txR | $((get-alias ps1GetxMUsr).definition)..." ;
                                write-verbose "`$txR | Get-xoMailUser..." ;
                                $hSum.xoMUser = $txR | Get-xoMailUser -ResultSize $MaxRecips | select -first $MaxRecips ;
                                #write-verbose "`$txR | $((get-alias ps1GetxUser).definition)..." ;
                                write-verbose "`$txR | get-xoUser ..." ;
                                $hSum.xoUser = $txR | get-xouser -ResultSize $MaxRecips | select -first $MaxRecips ;
                                write-verbose "`$hSum.xoUser:`n$(($hSum.xoUser|out-string).trim())" ;
                                #write-verbose "get-AzureAdUser  -objectid $($hSum.xoUser.userPrincipalName)" ;
                                #$hSum.AADUser  = get-AzureAdUser  -objectid $hSum.xoMUser.userPrincipalName -Top $MaxRecips ;
                                write-verbose "`$hSum.xoMUser:`n$(($hSum.xoMUser|out-string).trim())" ;
                                #$Rpt += $hSum.OPRemoteMailbox.primarysmtpaddress ;
                                write-host "$($txR.ExternalEmailAddress): matches a MailUser object with UPN:$($hSum.xoMUser.userPrincipalName)" ;
                                if($outObject){

                                } else {
                                    $Rpt += $hSum.xoMUser.primarysmtpaddress ;
                                } ;
                                break ;
                            } ;
                            "GuestMailUser" {
                                #$hSum.OPRemoteMailbox=get-remotemailbox $txR.identity  ;
                                caad -silent -verbose:$false ;
                                #write-verbose "`$txR | $((get-alias ps1GetxUser).definition)..." ;
                                write-verbose "`$txR | get-xoUser..." ; 
                                $hSum.xoUser = $txR | get-xouser -ResultSize $MaxRecips | select -first $MaxRecips ;
                                write-verbose "`$hSum.xoUser:`n$(($hSum.xoUser|out-string).trim())" ;
                                write-verbose "get-AzureAdUser  -objectid $($hSum.xoUser.userPrincipalName)" ;
                                $hSum.txGuest = get-AzureAdUser  -objectid $hSum.xoUser.userPrincipalName -Top $MaxRecips | select -first $MaxRecips ;
                                write-verbose "`$hSum.txGuest:`n$(($hSum.txGuest|out-string).trim())" ;
                                #$Rpt += $hSum.OPRemoteMailbox.primarysmtpaddress ;
                                write-host "$($txR.ExternalEmailAddress): matches a Guest object with UPN:$($hSum.xoUser.userPrincipalName)" ;
                                if($hSum.txGuest.EmailAddresses -eq $null){
                                    write-warning "Guest appears to have damage from conficting replicated onprem MailContact, as it's EmailAddresses property is *blank*" ;
                                } ;
                                break ;
                            } ;
                            "MailContact" {
                                #$hSum.OPRemoteMailbox=get-remotemailbox $txR.identity  ;
                                #$Rpt += $hSum.OPRemoteMailbox.primarysmtpaddress ;
                                write-host "$($txR.primarysmtpaddress): matches an EXO MailContact with external Email: $($txR.primarysmtpaddress)" ;
                                break ;
                            } ;
                            "MailUniversalSecurityGroup" {
                                #$hSum.OPRemoteMailbox=get-remotemailbox $txR.identity  ;
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
                    } ;
                    if($pltGadu.identity){
                        write-verbose "Get-ADUser w`n$(($pltGadu|out-string).trim())" ;
                        # try a nested local trycatch, against a missing result
                        Try {
                            #Get-ADUser $DN -ErrorAction Stop ;
                            $hSum.ADUser =Get-ADUser @pltGadu | select -first $MaxRecips ;
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
                        $hSum.Federator = $TORMeta.adforestname ;
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
                                $hSum.Federator = $CMWMeta.adforestname ;
                            }
                            $TORMeta.rgxOPFederatedDom {
                                $smsg="(TOR USER, fed:$($TORMeta.adforestname))" ;
                                $hSum.Federator = $TORMeta.adforestname ;
                            }
                            $VENMeta.rgxOPFederatedDom {
                                $smsg="(VEN USER, fed:$($venmeta.o365_TenantLabel))" ;
                                $hSum.Federator = $VENMETA.o365_TenantLabel ;
                            }
                            $INTMeta.rgxOPFederatedDom {
                                $smsg="(INT USER, fed:$($INTmeta.o365_TenantLabel))" ;
                                $hSum.Federator = $INTMETA.o365_TenantLabel ;
                            }

                        } ;
                    } elseif($hSum.xoMuser.IsDirSynced){
                        switch -regex ($xmbx.primarysmtpaddress.split('@')[1]){
                            $CMWMeta.rgxOPFederatedDom {
                                $smsg="(CMW USER, fed:$($CMWMeta.adforestname))" ;
                                $hSum.Federator = $CMWMeta.adforestname ;
                            }
                            $TORMeta.rgxOPFederatedDom {
                                $smsg="(TOR USER, fed:$($TORMeta.adforestname))" ;
                                $hSum.Federator = $TORMeta.adforestname ;
                            }
                            $VENMeta.rgxOPFederatedDom {
                                $smsg="(VEN USER, fed:$($venmeta.o365_TenantLabel))" ;
                                $hSum.Federator = $VENMETA.o365_TenantLabel ;
                            }
                            $INTMeta.rgxOPFederatedDom {
                                $smsg="(INT USER, fed:$($INTmeta.o365_TenantLabel))" ;
                                $hSum.Federator = $INTMETA.o365_TenantLabel ;
                            }
                        } ;
                    }else{
                        [regex]$rgxTenDom = [regex]::escape("@$($tormeta.o365_TenantDomain)")
                        if($hsum.xoRcp.primarysmtpaddress -match $rgxTenDom){
                                $smsg="(CLOUD-1ST ACCT, unfederated)" ;
                                $hSum.Federator = $TORMeta.o365_TenantDom ;

                        } else {
                            $smsg="(CLOUD-1ST ACCT, unfederated)" ;
                            $hSum.Federator = $TORMeta.o365_TenantDom ;
                        } ;
                    } ;
                } ;  # loop-E
                write-host -Fore yellow $smsg ;
                # skip user lookup if guest already pulled it
                if(!$hSum.xoUser){
                    $ino = 0 ;
                    foreach($xmbx in $hSum.xoMailbox){
                        #write-verbose "$((get-alias ps1GetxUser).definition) -id $($xmbx.UserPrincipalName)"
                        write-verbose "get-xoUser -id $($xmbx.UserPrincipalName)"
                        $hSum.xoUser += get-xouser -id $xmbx.UserPrincipalName -ResultSize $MaxRecips ;
                        write-verbose "`$hSum.xoUser:`n$(($hSum.xoUser|out-string).trim())" ;
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
                    $hSum.xoMemberOf = get-xorecipient -Filter "Members -eq '$($hSum.xoUser.DistinguishedName)'" -RecipientTypeDetails GroupMailbox,MailUniversalDistributionGroup,MailUniversalSecurityGroup ;
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
                    if($hSum.Rcp=get-recipient -id $substring -ea 0 -ResultSize $MaxRecips | select -first $MaxRecips |?{$_.recipienttypedetails -ne 'MailContact'}){
                        #$hSum.Rcp | write-output ;
                        # $propsRcpTbl
                        write-host -foregroundcolor yellow "`n$(($hSum.Rcp | ft -a $propsRcpTbl |out-string).trim())" ;
                    } ;
                    #write-host "$((get-alias ps1GetxRcp).definition) -id $($substring) -ea 0 |?{$_.recipienttypedetails -ne 'MailContact'} : "
                    write-host "get-xorecipient -id $($substring) -ea 0 |?{$_.recipienttypedetails -ne 'MailContact'} : "
                    if($hSum.xoRcp=get-xorecipient -id $substring -ea 0 -ResultSize $MaxRecips | select -first $MaxRecips |?{$_.recipienttypedetails -ne 'MailContact'}){
                        #$hSum.xoRcp | write-output ;
                        write-host -foregroundcolor yellow "`n$(($hSum.xoRcp | ft -a $propsRcpTbl |out-string).trim())" ;
                    } ;


                } ;


            } ; # don't break, doesn't continue loop

            # 10:42 AM 9/9/2021 force populate the xoMailbox, ALWAYS - need for xbrain ids
            #if($hSum.xoRcp.recipienttypedetails -eq 'UserMailbox' -AND -not($hSum.xoMailbox)){
            # accomodate array xorcp
            if(($hSum.xoRcp|?{$_.recipienttypedetails -eq 'UserMailbox'}) -AND -not($hSum.xoMailbox)){
                #write-verbose "$((get-alias ps1GetxMbx).definition) w`n$(($pltGMailObj|out-string).trim())" ;
                write-verbose "get-xomailbox w`n$(($pltGMailObj|out-string).trim())" ;
                if($hSum.xoMailbox=get-xomailbox @pltGMailObj -ea 0| select -first $MaxRecips ){
                    $ino = 0 ;
                    $mapiResults = @() ;
                    foreach($xmbx in $hSum.xoMailbox){
                        $ino++ ;
                        if($hSum.xoMailbox -is [system.array]){
                            $msgprefix = "xmbx$($ino):" ;
                        } else { $msgprefix = $null } ;
                        $smsg = $msgprefix, "`$hSum.xoMailbox:`n$(($xmbx|out-string).trim())" -join ' ' ;
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
                    $hSum.xoMapiTest = $mapiResults ;
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
                    $hSum.xoMailboxStats = Get-xoMailboxStatistics @pltGMbxStatX | select $prpStat;
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
                    $hSum.xoNetOfSendReceiveQuotaMB = $hSum.xoEffectiveQuotas.ProhibitSendQuotaMB - $hSum.xoMailboxStats.TotalMailboxSizeMB ; 

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
                        $hsum.xoMailboxFolderStats = Get-xoMailboxFolderStatistics @pltGMbxStatX  ;

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
            if($hSum.ADUser){$pltgAADUsr.UserPrincipalName = $hSum.ADUser.UserPrincipalName }
            elseif($hSum.xoMailbox){$pltgAADUsr.UserPrincipalName += $hsum.xoMailbox.UserPrincipalName }
            elseif($hSum.xoMUser){$pltgAADUsr.UserPrincipalName = $hSum.xoMUser.UserPrincipalName }
            elseif($hSum.txGuest){$pltgAADUsr.UserPrincipalName = $hSum.txGuest.userprincipalname }
            else{} ;

            if($pltgAADUsr.UserPrincipalName){

                if(-not($hSum.AADUser)){
                    write-host -foregroundcolor yellow "=get-AADuser $($pltgAADUsr.UserPrincipalName)>:" ;
                    TRY{
                        caad  -Verbose:$false -silent ;
                        write-verbose "get-AzureAdUser  -objectid $($pltgAADUsr.UserPrincipalName)" ;
                        # have to postfilter, if want specific count -maxresults catch's with no $error[0]
                        $hSum.AADUser  = get-AzureAdUser  -objectid $pltgAADUsr.UserPrincipalName  | select -first $MaxRecips;  ;
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
                        $hsum.AADUserLics =  $hSum.AADUser | Get-AzureADUserLicenseDetail -ea STOP | select -exp SkuPartNumber ; 

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
                        $hSum.AADUserMgr  = get-AzureAdUserManager  -objectid $hSum.AADUser.UserPrincipalName  | select -first $MaxRecips;  ;
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

                    write-host $smsg

                    # assert the real names from the user obj
                    $hSum.dname = $hSum.AADUser.DisplayName ;
                    $hSum.fname = $hSum.AADUser.GivenName ;
                    $hSum.lname = $hSum.AADUser.Surname ;

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
                    $hSum.dname = $hSum.ADUser.DisplayName ;
                    $hSum.fname = $hSum.ADUser.GivenName ;
                    $hSum.lname = $hSum.ADUser.Surname ;
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
                if($hSum.ADUser){$hSum.LicenseGroup = $hSum.ADUser.memberof |?{$_ -match $rgxOPLic }}

                if($hSum.ADUser){$hSum.LicenseGroup = $hSum.ADUser.memberof |?{$_ -match $rgxOPLic }}
                elseif($hSum.xoMemberOf){$hSum.LicenseGroup = $hSum.xoMemberOf.Name |?{$_ -match $rgxXLic}}
                #if(!($hSum.LicenseGroup) -AND ($hSum.MsolUser.licenses.AccountSkuId -contains "$($TORMeta.o365_TenantDom.tolower()):ENTERPRISEPACK")){$hSum.LicenseGroup = '(direct-assigned E3)'} ;
                # $hSum.AADUser ; $aadu | Get-AzureADUserLicenseDetail  | select -exp SkuPartNumber
                #if(!($hSum.LicenseGroup) -AND ( $hsum.AADUserLics  -contains "$($TORMeta.o365_TenantDom.tolower()):ENTERPRISEPACK")){$hSum.LicenseGroup = '(direct-assigned E3)'} ;
                # no dom, with aadu licenses
                if(!($hSum.LicenseGroup) -AND ( $hsum.AADUserLics  -contains "ENTERPRISEPACK")){$hSum.LicenseGroup = '(direct-assigned E3)'} ;
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

                    if($hSum.xoMailbox.LitigationHoldEnabled -OR $hSum.xoMailbox.InPlaceHolds -OR $hSum.xoMailbox.ComplianceTagHoldApplied -OR $hSum.xoMailbox.DelayHoldApplied -OR $hSum.xoMailbox.DelayReleaseHoldApplied){
                        $smsg = "`n`nEVIDENCE OF LEGAL HOLD DETECTED!:`n$(($hSum.xoMailbox | fl $prpMbxHold|out-string).trim())`n`n" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 

                    } else {
                        $smsg = "`n`n*No* evidence Of Legal Hold detected:`n$(($hSum.xoMailbox | fl $prpMbxHold|out-string).trim())`n`n" ; 
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
                $hsum.IsADDisabled = [boolean]($hsum.ADUser.Enabled -eq $true) ; 
             } else {
                write-verbose "(no ADUser found)" ;
            } ;
            if($hsum.AADUser){
                $hsum.IsAADDisabled = [boolean]($hsum.AADUser.AccountEnabled -eq $true) ; 
                $hsum.isDirSynced = [boolean]($hsum.AADUser.DirSyncEnabled  -eq $True)
            } else {
                write-verbose "(no AADUser found)" ;
            } ;
            # shift test to aadu
            if($hSum.AADUser){
                $hsum.IsLicensed = [boolean]($hSum.AADUser.assignedlicenses.count -gt 0)
            } else {
                write-verbose "(no AADUser found)" ;
            } ;

            $smsg = "`n"
            if(($hsum.xoRcp.RecipientTypeDetails -match '(UserMailbox|MailUser)') -AND $hsum.IsLicensed -AND $hSum.xomailbox -AND $hSum.OPMailbox){
                <#OPRcp, xorcp, OPMailbox, OPRemoteMailbox, xoMailbox#>
                $smsg += "SPLITBRAIN!:$($hSum.ADUser.userprincipalname).IsLic'd & has *BOTH* xoMbx & opMbx!" ;
                $hsum.IsSplitBrain = $true ;
            }elseif(($hsum.xoRcp.RecipientTypeDetails -match '(UserMailbox|MailUser)') -AND -not($hsum.IsLicensed) -AND $hSum.xomailbox -AND $hSum.OPMailbox){
                <#OPRcp, xorcp, OPMailbox, OPRemoteMailbox, xoMailbox#>
                $smsg += "SPLITBRAIN!:$($hSum.ADUser.userprincipalname).IsLic'd & has *BOTH* xoMbx & opMbx!`nAND is *UNLICENSED!*" ;
                $hsum.IsSplitBrain = $true ;
            } elseif(($hsum.xoRcp.RecipientTypeDetails -match '(UserMailbox|MailUser)') -AND $hsum.IsLicensed -AND -not($hSum.xomailbox) -AND -not($hSum.OPMailbox)){
                $smsg += "NOBRAIN! W LICENSE!:$($hSum.ADUser.userprincipalname).IsLic'd &  has *NEITHER* xoMbx OR opMbx!" ;
                $hsum.IsNoBrain = $true ;
            } elseif (($hsum.xoRcp.RecipientTypeDetails -match '(UserMailbox|MailUser)') -AND -not($hsum.IsLicensed) -AND -not($hSum.xomailbox) -AND -not($hSum.OPMailbox)){
                $smsg += "NOBRAIN! *WO* LICENSE! (TERM?):$($hSum.ADUser.userprincipalname) NOT licensed'd &  has *NEITHER* xoMbx OR opMbx!" ;
                $hsum.IsNoBrain = $true ;
            } elseif($hsum.IsLicensed -eq $false){
                $smsg += "$($hSum.ADUser.userprincipalname) Is *UNLICENSED*!" ;
                
                $hsum.IsLicensed = $false ;
            } ELSE { } ;

            if($hsum.IsSplitBrain -OR $hsum.IsNoBrain -OR -not $hsum.IsLicensed){
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
                        $hsum.IsDisabledOU = $true ;
                        $smsg += "`n--ADUser:$($hsum.ADUser.samaccountname) is within a *DISABLED* OU (likely TERM)" ;
                    } else {
                        $hsum.IsDisabledOU = $false ;
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

    } # PROC-E
    END{
        <## cleanup XO aliases
        get-alias -scope Script |?{$_.name -match '^ps1.*'} | %{Remove-Alias -alias $_.name} ; 
        #>

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

     } ;
 }

#*------^ resolve-user.ps1 ^------