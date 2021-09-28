﻿#*------v resolve-user.ps1 v------
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
    * 1:04 PM 9/28/2021 added:$AADUserManager lookup and dump of UPN, OpDN & mail (for correlating what email pol a user should have -> the one their manager does)
    * 1:52 PM 9/17/2021 moved $props to top ; test enabled/acctenabled, licRecon & mapi test results and use ww on issues ; flipped caad's to -silent (match cmsol 1st echo's to confirm tenant, rest silent); ren $xMProps -> $propsMailx, $XMFedProps-> $propsXMFed, $lProps -> $propsLic,$adprops -> $propsADU, $aaduprops -> $propsAADU, $aaduFedProps -> $propsAADUfed, $RcpPropsTbl -> $propsRcpTbl, $pltgM-> $pltGMailObj, $pltgMU -> $pltgMsoUsr
    * 4:33 PM 9/16/2021 fixed typo in get-AzureAdUser call, reworked output (aadu into markdown delimited wide layout), moved user detaiil reporting to below aadu, and output the federated AD remote DN, (proxied through AADU ext prop)
    * 10:56 AM 9/9/2021 force-resolve xoMailbox, added AADUser pop to the msoluser pop block; added test-xxMapiConnectivity as well; expanded ADU outputs - description, when*, Enabled, to look for terms/recent-hires/disabled accts
    * 3:05 PM 9/3/2021 fixed bugs introduced trying to user MaxResults (msol|aad), which come back param not recog'd when actually used - had to implement as postfiltering to assert open set return limits. ; Also implemented $xxxMeta.rgxOPFederatedDom check to resolve obj primarysmtpaddress to federating AD or AAD.
    * 11:20 AM 8/30/2021 added $MaxResults (shutdown return-all recips in addr space, on failure to match oprcp or xorcp ; fixed a couple of typos; minior testing/logic improvements. Still needs genercized 7pswlt support.
    * 1:30 PM 8/27/2021 new sniggle: CMW user that has EXOP mbx, remote: Added xoMailUser support, failed through DName lookups to try '*lname*' for near-missies. Could add trailing 'lnamne[0-=3]* searches, if not rcp/xrcps found...
    * 9:16 AM 8/18/2021 $xMProps: add email-drivers: CustomAttribute5, EmailAddressPolicyEnabled
    * 12:40 PM 8/17/2021 added -outObject, outputs a full descriptive object for each resolved recipient ; added a $hSum hash and shifted all the varis into mountpoints in the hash, with -outObject, the entire hash is conv'd to an obj and appended to $Rpt ; renamed most of the varis/as objects very clearly for what they are, as sub-props of the output objects. Wo -outobject, the usual comma-delim'd string of addresses is output.
    * 3:26 PM 7/29/2021 had sorta bug (AD context was adtorocom:, gadu failing throwing undefined error), but debugging added extensive verbose echos, and an AD-specific try/catch to trap AD notfound errors (notorious, they throw terminating fails, unlike other modules; which crashes out processing even when using -EA continue). So it hardens up the fail recovery process.
    * 12:55 PM 7/19/2021 added guest & exo-mailcontact support (resolving missing ext-federated addresses), retolled logic down to grcp & gxrcp to drive balance of tests.
    * 12:05 PM 7/14/2021 rem'd requires: verb-exo  rem'd requires version 5 (gen'ing 'version' is specified more than once.); rem'd the $rgxSamAcctName, gen's parsing errors compiling into mod ;  added alias 'ulu'; added mailcontact excl on init grcp, to force those to exombx qry ; init vers
    .DESCRIPTION
    .PARAMETER  users
    Array of user descriptors: displayname, emailaddress, UPN, samaccountname (checks clipboard where unspecified)
    .PARAMETER useEXOv2
    Use EXOv2 (ExchangeOnlineManagement) over basic auth legacy connection [-useEXOv2]
    .INPUTS
    None. Does not accepted piped input.(.NET types, can add description)
    .OUTPUTS
    Returns report to pipeline
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
    PS> $feds = $results| group federator | select -expand name ;
    PS> ($results| ?{$_.federator -eq $feds[1] }).xomailbox
    PS> ($results| ?{$_.federator -eq $feds[1] }).xomailbox.primarysmtpaddress
    Process array of users, specify return detailed object (-outobject), for post-processing & filtering,
    group results on federation sources,
    output summary of EXO mailboxes for the second federator
    then output the primary smtpaddress for all EXO mailboxes resolved to that federator
    .LINK
    https://github.com/tostka/verb-exo
    .LINK
    #>
    ###Requires -Version 5
    #Requires -Modules ActiveDirectory, MSOnline, AzureAD, ExchangeOnlineManagement, verb-AAD, verb-ADMS, verb-Ex2010
    #Requires -RunasAdministrator
    # VALIDATORS: [ValidateNotNull()][ValidateNotNullOrEmpty()][ValidateLength(24,25)][ValidateLength(5)][ValidatePattern("(lyn|bcc|spb|adl)ms6(4|5)(0|1).(china|global)\.ad\.toro\.com")][ValidateSet("USEA","GBMK","AUSYD")][ValidateScript({Test-Path $_ -PathType 'Container'})][ValidateScript({Test-Path $_})][ValidateRange(21,65)][ValidateCount(1,3)]
    ## [OutputType('bool')] # optional specified output type
    [CmdletBinding()]
    [Alias('ulu')]
    PARAM(
        [Parameter(Position=0,Mandatory=$False,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,HelpMessage="Array of user descriptors: displayname, emailaddress, UPN, samaccountname (checks clipboard where unspecified)")]
        #[ValidateNotNullOrEmpty()]
        #[Alias('ALIAS1', 'ALIAS2')]
        [array]$users,
        [Parameter(HelpMessage="Use EXOv2 (ExchangeOnlineManagement) over basic auth legacy connection [-useEXOv2]")]
        [switch] $useEXOv2,
        [Parameter(HelpMessage="Use EXOv2 (ExchangeOnlineManagement) over basic auth legacy connection [-useEXOv2]")]
        [switch] $outObject

    ) ;
    BEGIN{
        $Verbose = ($VerbosePreference -eq 'Continue') ;
        $rgxEmailAddr = "^([0-9a-zA-Z]+[-._+&'])*[0-9a-zA-Z]+@([-0-9a-zA-Z]+[.])+[a-zA-Z]{2,63}$" ;
        $rgxDName = "^([a-zA-Z]{2,}\s[a-zA-Z]{1,}'?-?[a-zA-Z]{2,}\s?([a-zA-Z]{1,})?)" ;
        $rgxSamAcctNameTOR = "^\w{2,20}$" ; # up to 20k, the limit prior to win2k
        #$rgxSamAcctName = "^[^\/\\\[\]:;|=,+?<>@?]+$" # no char limit ;
        $MaxRecips = 25 ; # max number of objects to permit on a return resultsize/,ResultSetSize, to prevent empty set return of everything in the addressspace

        # props dyn filtering: write-host "=get-xMbx:>`n$(($hSum.xoMailbox |fl ($xMprops |?{$_ -notmatch '(sam.*|dist.*)'})|out-string).trim())`n-Title:$($hSum.xoUser.Title)";
        # $propsMailx: add email-drivers: CustomAttribute5, EmailAddressPolicyEnabled
        $propsMailx='samaccountname','windowsemailaddress','DistinguishedName','Office','RecipientTypeDetails','RemoteRecipientType','IsDirSynced','ExternalDirectoryObjectId','CustomAttribute5','EmailAddressPolicyEnabled' ;
        # pulls: 'ImmutableId',
        $propsXMFed = 'samaccountname','windowsemailaddress','DistinguishedName','Office','RecipientTypeDetails','RemoteRecipientType','ImmutableId','ExternalDirectoryObjectId','CustomAttribute5','EmailAddressPolicyEnabled' ; 
        $propsLic = @{Name='HasLic'; Expression={$_.IsLicensed }},@{Name='LicIssue'; Expression={$_.LicenseReconciliationNeeded }} ;
        $propsADU = 'UserPrincipalName','DisplayName','GivenName','Surname','Title','Company','Department','PhysicalDeliveryOfficeName','StreetAddress','City','State','PostalCode','TelephoneNumber','MobilePhone','Enabled','DistinguishedName','Description','whenCreated','whenChanged'
        #'samaccountname','UserPrincipalName','distinguishedname','Description','title','whenCreated','whenChanged','Enabled','sAMAccountType','userAccountControl' ;
        $propsADUsht = 'Enabled','Description','whenCreated','whenChanged','Title' ;
        $propsAADU = 'UserPrincipalName','DisplayName','GivenName','Surname','Title','Company','Department','PhysicalDeliveryOfficeName','StreetAddress','City','State','PostalCode','TelephoneNumber','MobilePhone','Enabled','DistinguishedName' ;
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
        $propsADL1 = @{Name='UPN';Expression={$_.UserPrincipalName }}, @{Name='DName';Expression={$_.DisplayName }}, @{Name='FName';Expression={$_.GivenName }},@{Name='LName';Expression={$_.Surname }},@{Name='Title';Expression={$_.Title }};
        $propsADL2 = @{Name='Company';Expression={$_.Company }},@{Name='Dept';Expression={$_.Department }},@{Name='Ofc';Expression={$_.PhysicalDeliveryOfficeName }} ;
        $propsADL3 = @{Name='Street';Expression={$_.StreetAddress }}, 'City','State',@{Name='Zip';Expression={$_.PostalCode }}, @{Name='Phone';Expression={$_.TelephoneNumber }}, @{Name='Mobile';Expression={$_.MobilePhone }} ;
        $propsADL4 = 'Enabled',@{Name='DN';Expression={$_.DistinguishedName }} ;
        #$propsADL4 = @{Name='Dsync';Expression={$_.DirSyncEnabled }}, @{Name='ImutID';Expression={$_.ImmutableId }}, @{Name='LastDSync';Expression={$_.LastDirSyncTime }}, @{Name='UseLoc';Expression={$_.UsageLocation }};
        #$propsADL5 = 'ObjectType','UserType' ;

        # line1-5 AADU outputs
        <# full size
        $propsAADL1 = 'UserPrincipalName','DisplayName','GivenName','Surname','JobTitle' ;
        $propsAADL2 = 'CompanyName','Department','PhysicalDeliveryOfficeName' ;
        $propsAADL3 = 'StreetAddress','City','State','PostalCode','TelephoneNumber','Mobile' ;
        $propsAADL4 = 'DirSyncEnabled','ImmutableId','LastDirSyncTime','UsageLocation' ;
        $propsAADL5 = 'ObjectType','UserType' ;
        #>
        # abbreviated:
        $propsAADL1 = @{Name='UPN';Expression={$_.UserPrincipalName }}, @{Name='DName';Expression={$_.DisplayName }}, @{Name='FName';Expression={$_.GivenName }},@{Name='LName';Expression={$_.Surname }},@{Name='Title';Expression={$_.JobTitle }};
        $propsAADL2 = @{Name='Company';Expression={$_.CompanyName }},@{Name='Dept';Expression={$_.Department }},@{Name='Ofc';Expression={$_.PhysicalDeliveryOfficeName }} ;
        $propsAADL3 = @{Name='Street';Expression={$_.StreetAddress }}, 'City','State',@{Name='Zip';Expression={$_.PostalCode }}, @{Name='Phone';Expression={$_.TelephoneNumber }}, 'Mobile' ;
        $propsAADL4 = @{Name='Dsync';Expression={$_.DirSyncEnabled }}, @{Name='ImutID';Expression={$_.ImmutableId }}, @{Name='LastDSync';Expression={$_.LastDirSyncTime }}, @{Name='UseLoc';Expression={$_.UsageLocation }};
        $propsAADL5 = 'ObjectType','UserType', @{Name='Enabled';Expression={$_.AccountEnabled }} ;

        #$propsAADMgr = 'UserPrincipalName','Mail',@{Name='OpDN';Expression={$_.ExtensionProperty.onPremisesDistinguishedName }} ;
        # get mgr OU, not DN: ExtensionProperty.onPremisesDistinguishedName.split(',') | select -skip 1 ) -join ','
        $propsAADMgr = 'UserPrincipalName','Mail',@{Name='OpOU';Expression={($_.ExtensionProperty.onPremisesDistinguishedName.split(',') | select -skip 1) -join ',' }} ;
        $propsAADMgrL1 = 'UserPrincipalName','Mail' ;
        $propsAADMgrL2 = @{Name='OpOU';Expression={($_.ExtensionProperty.onPremisesDistinguishedName.split(',') | select -skip 1) -join ',' }} ;

        $rgxOPLic = '^CN\=ENT\-APP\-Office365\-(EXOK|F1|MF1)-DL$' ;
        $rgxXLic = '^CN\=ENT\-APP\-Office365\-(EXOK|F1|MF1)-DL$' ;

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

        rx10 -Verbose:$false ; rxo  -Verbose:$false ; cmsol  -Verbose:$false ;

    }
    PROCESS{
        #$dname= 'Todd Kadrie' ;
        #$dname = 'Stacy Sotelo'
        $ttl = ($users|measure).count ; $Procd=0 ;
        [array]$Rpt =@() ;
        foreach ($usr in $users){
            #$fname = $lname = $dname = $OPRcp = $OPMailbox = $OPRemoteMailbox = $ADUser = $xoRcp = $xoMailbox = $xoUser = $xoMemberOf = $MsolUser = $LicenseGroup = $null ;
            $isEml=$isDname=$isSamAcct=$false ;

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
                LicenseGroup = $null ;
            } ;
            $procd++ ;
            write-verbose "processing:$($usr)" ;
            switch -regex ($usr){
                $rgxEmailAddr {
                    $hSum.fname,$hSum.lname = $usr.split('@')[0].split('.') ;
                    $hSum.dname = $usr ;
                    write-verbose "(detected user ($($usr)) as EmailAddr)" ;
                    $isEml = $true ;
                }
                $rgxDName {
                    $hSum.fname,$hSum.lname = $usr.split(' ') ;
                    $hSum.dname = $usr ;
                    write-verbose "(detected user ($($usr)) as DisplayName)" ;
                    $isDname = $true ;
                }
                $rgxSamAcctNameTOR {
                    $hSum.lname = $usr ;
                    write-verbose "(detected user ($($usr)) as SamAccountName)" ;
                    $isSamAcct  = $true ;
                }
                default {
                    write-warning "$((get-date).ToString('HH:mm:ss')):No -user specified, nothing matching dname, emailaddress or samaccountname, found on clipboard. EXITING!" ;
                    #Break ;
                } ;
            } ;

            $sBnr="===v ($($Procd)/$($ttl)):Input: '$($usr)' | '$($hSum.fname)' | '$($hSum.lname)' v===" ;
            if($isEml){$sBnr+="(EML)"}
            elseif($isDname){$sBnr+="(DNAM)"}
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


            write-verbose "get-exorecipient w`n$(($pltGMailObj|out-string).trim())" ;
            rxo  -Verbose:$false -silent ;
            if($hSum.xoRcp=get-exorecipient @pltGMailObj -ea 0 | select -first $MaxRecips ){
                write-verbose "`$hSum.xoRcp found" ;
            } elseif($isDname -and $hsum.lname) {
                $smsg = "Failed:RETRY: detected 'LName':$($hsum.lname) for near matches..." ;
                write-host $smsg ;
                $lname = $hsum.lname ;
                $fltrB = "displayname -like '*$lname*'" ;
                write-verbose "RETRY:get-recipient -filter {$($fltr)}" ;
                if($hSum.xoRcp=get-exorecipient -filter $fltr -ea 0 -ResultSize $MaxRecips |?{$_.recipienttypedetails -ne 'MailContact'}){
                    write-verbose "`$hSum.xoRcp found" ;
                } ;
            }
            if(!$hsum.xoRcp){
                $smsg = "Failed to get-exorecipient on:$($usr)"
                if($isDname){$smsg += " or *$($hsum.lname )*"} ;
                write-host $smsg ;
            } else {
                write-verbose "`$hSum.xoRcp:`n$(($hSum.xoRcp|out-string).trim())" ;
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
                            $hSum.OPRemoteMailbox=get-remotemailbox $hSum.OPRcp.identity -resultsize $MaxRecips | select -first $MaxRecips  ;
                            write-verbose "`$hSum.OPRemoteMailbox:`n$(($hSum.OPRemoteMailbox|out-string).trim())" ;
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
                    } ;
                    if($pltGadu.identity){
                        <# this is throwing a blank fail
                        WARNING: 15:04:18:Failed processing .
                        Error Message:
                        Error Details:
                        # and dumping balance of processing
                        issue: was in adms drive: :adtorocom, gadu was searching root domain only
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
                        $smsg = "(TOR USER, fed:ad.toro.com)" ;
                        $hSum.Federator = 'ad.toro.com' ;
                        write-host -Fore yellow $smsg ;
                        if($hSum.OPRemoteMailbox){
                            $smsg = "$(($hSum.OPRemoteMailbox |fl $propsMailx|out-string).trim())"
                        } ;
                        if($hSum.OPMailbox){
                            $smsg =  "$(($hSum.OPMailbox |fl $propsMailx|out-string).trim())" ;
                        } ;
                        $smsg += "`n$(($hSum.ADUser |fl $propsADUsht  |out-string).trim())"
                        write-host $smsg ;
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
                                write-verbose "get-exomailbox w`n$(($pltGMailObj|out-string).trim())" ;
                                if($hSum.xoMailbox=get-exomailbox @pltGMailObj -ea 0 | select -first $MaxRecips ){
                                    write-verbose "`$hSum.xoMailbox:`n$(($hSum.xoMailbox|out-string).trim())" ;
                                    if($outObject){

                                    } else {
                                        $Rpt += $hSum.xoMailbox.primarysmtpaddress ;
                                    } ;
                                    write-verbose "'xoUserMailbox':Test-exoMAPIConnectivity $($hSum.xoMailbox.userprincipalname)"
                                    $hSum.xoMapiTest = Test-exoMAPIConnectivity -identity $hSum.xoMailbox.userprincipalname ;
                                    $smsg = "Outlook (xoMAPI) Access Test Result:$($hsum.xoMapiTest.result)" ;
                                    if($hsum.xoMapiTest.result -eq 'Success'){
                                        write-host -foregroundcolor green $smsg ;
                                    } else {
                                        write-WARNING $smsg ;
                                    } ;
                                    break ;
                                } ;
                            }
                            "MailUser" {
                                # external mail recipient, *not* in TTC - likely in other rgs, and migrated to remote EXOP enviro
                                #$hSum.OPRemoteMailbox=get-remotemailbox $txR.identity  ;
                                caad -silent -verbose:$false ;
                                write-verbose "`$txR | get-exoMailuser..." ;
                                $hSum.xoMUser = $txR | get-exoMailuser -ResultSize $MaxRecips | select -first $MaxRecips ;
                                write-verbose "`$txR | get-exouser..." ;
                                $hSum.xoUser = $txR | get-exouser -ResultSize $MaxRecips | select -first $MaxRecips ;
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
                                write-verbose "`$txR | get-exouser..." ;
                                $hSum.xoUser = $txR | get-exouser -ResultSize $MaxRecips | select -first $MaxRecips ;
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
                        $smsg = "(TOR USER, fed:ad.toro.com)" ;
                        $hSum.Federator = 'ad.toro.com' ;
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
                if($hSum.xoMailbox.isdirsynced){
                    # can be federated to VEN|CMW|Toro
                    switch -regex ($hSum.xoMailbox.primarysmtpaddress.split('@')[1]){
                        $CMWMeta.rgxOPFederatedDom {
                            $smsg="(CMW USER, fed:cmw.internal)" ;
                            $hSum.Federator = 'cmw.internal' ;
                        }
                        $TORMeta.rgxOPFederatedDom {
                            $smsg="(TOR USER, fed:ad.toro.com)" ;
                            $hSum.Federator = 'ad.toro.com' ;
                        }
                        $VENMeta.rgxOPFederatedDom {
                            $smsg="(VEN USER, fed:ventrac)" ;
                            $hSum.Federator = 'ventrac' ;
                        }

                    } ;
                } elseif($hSum.xoMuser.IsDirSynced){
                    switch -regex ($hSum.xoMailbox.primarysmtpaddress.split('@')[1]){
                        $CMWMeta.rgxOPFederatedDom {
                            $smsg="(CMW USER, fed:cmw.internal)" ;
                            $hSum.Federator = 'cmw.internal' ;
                        }
                        $TORMeta.rgxOPFederatedDom {
                            $smsg="(TOR USER, fed:ad.toro.com)" ;
                            $hSum.Federator = 'ad.toro.com' ;
                        }
                        $VENMeta.rgxOPFederatedDom {
                            $smsg="(VEN USER, fed:ventrac)" ;
                            $hSum.Federator = 'ventrac' ;
                        }
                    } ;
                }else{
                    if($hsum.xoRcp.primarysmtpaddress -match "@toroco\.onmicrosoft\.com"){
                            $smsg="(CLOUD-1ST ACCT, unfederated)" ;
                            $hSum.Federator = 'Toroco' ;

                    } else {
                        $smsg="(CLOUD-1ST ACCT, unfederated)" ;
                        $hSum.Federator = 'Toroco' ;
                    } ;
                } ;
                write-host -Fore yellow $smsg ;
                # skip user lookup if guest already pulled it
                if(!$hSum.xoUser){
                    write-verbose "get-exouser -id $($hSum.xoMailbox.UserPrincipalName)"
                    $hSum.xoUser = get-exouser -id $hSum.xoMailbox.UserPrincipalName -ResultSize $MaxRecips ;
                    write-verbose "`$hSum.xoUser:`n$(($hSum.xoUser|out-string).trim())" ;
                }
                if($hSum.xoMailbox){
                    write-host -foreground yellow "=get-xMbx:> " -nonewline;
                    write-host "$(($hSum.xoMailbox |fl ($propsMailx |?{$_ -notmatch '(sam.*|dist.*)'})|out-string).trim())`n-Title:$($hSum.xoUser.Title)";
                }elseif($hSum.xoMUser){
                    write-host "=get-xMUSR:>`n$(($hSum.xoMUser |fl ($propsMailx |?{$_ -notmatch '(sam.*|dist.*)'})|out-string).trim())`n-Title:$($hSum.xoUser.Title)";
                }elseif($hSum.txGuest){
                    write-host "=get-AADU:>`n$(($hSum.txGuest |fl userp*,PhysicalDeliveryOfficeName,JobTitle|out-string).trim())"
                } ;
                TRY {
                    write-verbose "Get-exoRecipient -Filter {Members -eq '$($hSum.xoUser.DistinguishedName)'}`n -RecipientTypeDetails GroupMailbox,MailUniversalDistributionGroup,MailUniversalSecurityGroup"
                    $hSum.xoMemberOf = Get-exoRecipient -Filter "Members -eq '$($hSum.xoUser.DistinguishedName)'" -RecipientTypeDetails GroupMailbox,MailUniversalDistributionGroup,MailUniversalSecurityGroup ;
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
                    write-host "get-exorecipient -id $($substring) -ea 0 |?{$_.recipienttypedetails -ne 'MailContact'} : "
                    if($hSum.xoRcp=get-exorecipient -id $substring -ea 0 -ResultSize $MaxRecips | select -first $MaxRecips |?{$_.recipienttypedetails -ne 'MailContact'}){
                        #$hSum.xoRcp | write-output ;
                        write-host -foregroundcolor yellow "`n$(($hSum.xoRcp | ft -a $propsRcpTbl |out-string).trim())" ;
                    } ;


                } ;


            } ; # don't break, doesn't continue loop

            # 10:42 AM 9/9/2021 force populate the xoMailbox, ALWAYS - need for xbrain ids
            if($hSum.xoRcp.recipienttypedetails -eq 'UserMailbox' -AND -not($hSum.xoMailbox)){
                write-verbose "get-exomailbox w`n$(($pltGMailObj|out-string).trim())" ;
                if($hSum.xoMailbox=get-exomailbox @pltGMailObj -ea 0| select -first $MaxRecips ){
                    write-verbose "`$hSum.xoMailbox:`n$(($hSum.xoMailbox|out-string).trim())" ;
                    write-verbose "'xoUserMailbox':Test-exoMAPIConnectivity $($hSum.xoMailbox.userprincipalname)"
                    $hSum.xoMapiTest = Test-exoMAPIConnectivity -identity $hSum.xoMailbox.userprincipalname ;
                    $smsg = "Outlook (xoMAPI) Access Test Result:$($hsum.xoMapiTest.result)" ;
                    if($hSum.xoMapiTest.result -eq 'Success'){
                        write-host -foregroundcolor green $smsg ;
                    } else {
                        write-WARNING $smsg ;
                    } ;
                } ;
            } ;

            #$pltgMsoUsr=@{UserPrincipalName=$null ; MaxResults= $MaxRecips; ErrorAction= 'STOP' } ;
            # maxresults is documented:
            # but causes a fault with no $error[0], doesn't seem to be functional param, post-filter
            $pltgMsoUsr=@{UserPrincipalName=$null ; ErrorAction= 'STOP' } ;
            if($hSum.ADUser){$pltgMsoUsr.UserPrincipalName = $hSum.ADUser.UserPrincipalName }
            elseif($hSum.xoMailbox){$pltgMsoUsr.UserPrincipalName = $hsum.xoMailbox.UserPrincipalName }
            elseif($hSum.xoMUser){$pltgMsoUsr.UserPrincipalName = $hSum.xoMUser.UserPrincipalName }
            elseif($hSum.txGuest){$pltgMsoUsr.UserPrincipalName = $hSum.txGuest.userprincipalname }
            else{} ;

            if($pltgMsoUsr.UserPrincipalName){
                write-host -foregroundcolor yellow "=get-msoluser $($pltgMsoUsr.UserPrincipalName):(licences)>:" ;
                TRY{
                    cmsol  -Verbose:$false -silent ;
                    write-verbose "get-msoluser w`n$(($pltgMsoUsr|out-string).trim())" ;
                    # have to postfilter, if want specific count -maxresults catch's with no $error[0]
                    $hSum.MsolUser=get-msoluser @pltgMsoUsr | select -first $MaxRecips;  ;
                    write-verbose "`$hSum.MsolUser:`n$(($hSum.MsolUser|out-string).trim())" ;
                } CATCH {
                    $ErrTrapd=$Error[0] ;
                    $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    Continue ;
                } ;

                if(-not($hSum.AADUser)){
                    #write-verbose "get-AzureAdUser  -objectid $($hSum.xoUser.userPrincipalName)" ;
                    #$hSum.AADUser  = get-AzureAdUser  -objectid $hSum.xoMUser.userPrincipalName -Top $MaxRecips ;
                    write-host -foregroundcolor yellow "=get-AADuser $($pltgMsoUsr.UserPrincipalName)>:" ;
                    TRY{
                        caad  -Verbose:$false -silent ;
                        write-verbose "get-AzureAdUser  -objectid $($pltgMsoUsr.UserPrincipalName)" ;
                        # have to postfilter, if want specific count -maxresults catch's with no $error[0]
                        $hSum.AADUser  = get-AzureAdUser  -objectid $pltgMsoUsr.UserPrincipalName  | select -first $MaxRecips;  ;
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
                    if($hSum.Federator -ne 'ad.toro.com'){
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
                    #$smsg += "`n$(($hSum.ADUser|select $propsADL5 |out-markdowntable @MDtbl|out-string).trim())" ;
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

                $smsg = "$(($hSum.MsolUser|fl $propsLic|out-string).trim())`n" ;
                $smsg += "Licenses Assigned:$(($hSum.MsolUser.licenses.AccountSkuId -join '; '|out-string).trim())" ;
                if($hSum.MsolUser.LicenseReconciliationNeeded){
                    write-WARNING $smsg ;
                } else {
                    write-host $smsg ;
                } ;
                if($hSum.ADUser){$hSum.LicenseGroup = $hSum.ADUser.memberof |?{$_ -match $rgxOPLic }}
                elseif($hSum.xoMemberOf){$hSum.LicenseGroup = $hSum.xoMemberOf.Name |?{$_ -match $rgxXLic}}
                if(!($hSum.LicenseGroup) -AND ($hSum.MsolUser.licenses.AccountSkuId -contains 'toroco:ENTERPRISEPACK')){$hSum.LicenseGroup = '(direct-assigned E3)'} ;
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
                    $smsg += "`n$(($hSum.AADUserMgr|fl $propsAADMgrL2|out-string).trim())" ;
                    #$smsg += "`n$(($hSum.AADUserMgr|select $propsADL3 |out-markdowntable @MDtbl|out-string).trim())" ;
                } else {
                    $smsg += "(AADUserMgr was blank, or unresolved)" ; 
                } ;  
                write-host $smsg ;

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
            switch ($hSum.Federator) {
                'ad.toro.com' {
                    if(($hsum.xoRcp -match '(RemoteUserMailbox|UserMailbox|MailUser)') -AND $hSum.MsolUser.IsLicensed -AND $hSum.xomailbox -AND $hSum.OPMailbox){
                        $smsg = "SPLITBRAIN!:$($hSum.ADUser.userprincipalname).IsLic'd & has *BOTH* xoMbx & opMbx!" ;
                        write-warning $smsg ;
                    } elseif(($hsum.xoRcp -match '(RemoteUserMailbox|UserMailbox|MailUser)') -AND $hSum.MsolUser.IsLicensed -AND -not($hSum.xomailbox) -AND -not($hSum.OPMailbox)){
                        $smsg = "NOBRAIN!:$($hSum.ADUser.userprincipalname).IsLic'd &  has *NEITHER* xoMbx OR opMbx!" ;
                        write-warning $smsg ;
                    } elseif($hSum.MsolUser.IsLicensed -eq $false){
                        $smsg = "$($hSum.ADUser.userprincipalname) Is *UNLICENSED*!" ;
                        write-warning $smsg ;
                    } ELSE { } ;

                }
                'cmw.internal' {
                    if($hSum.MsolUser.IsLicensed -eq $false){
                        $smsg = "$($hSum.ADUser.userprincipalname) Is *UNLICENSED*!" ;
                        write-warning $smsg ;
                    } ELSE { } ;
                }
                'ventrac'  {
                    if($hSum.MsolUser.IsLicensed -eq $false){
                        $smsg = "$($hSum.ADUser.userprincipalname) Is *UNLICENSED*!" ;
                        write-warning $smsg ;
                    } ELSE { } ;
                }
                'Toroco'  {
                    if($hSum.MsolUser.IsLicensed -eq $false){
                        $smsg = "$($hSum.ADUser.userprincipalname) Is *UNLICENSED*!" ;
                        write-warning $smsg ;
                    } ELSE { } ;
                }
                default {
                    write-warning "UNRECOGNIZED `$hsum.FEDERATOR!:$($hSum.Federator)" ;
                }
            }

            if($outObject){
                $Rpt += New-Object PSObject -Property $hSum ;
            } ;
            write-host -foregroundcolor green $sBnr.replace('=v','=^').replace('v=','^=') ;
        } ;

    }
    END{
        if($outObject){
            $Rpt | write-output ;
            write-host "(-outObject: Output summary object to pipeline)"
        } else {
            $oput = ($Rpt | select-object -unique) -join ',' ;
            $oput | out-clipboard ;
            write-host "(output copied to clipboard)"
            $oput |  write-output ;
        } ;

     } ;
 }

#*------^ resolve-user.ps1 ^------