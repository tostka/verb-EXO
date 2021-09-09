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
        
            # $xMProps: add email-drivers: CustomAttribute5, EmailAddressPolicyEnabled
            $xMProps='samaccountname','windowsemailaddress','DistinguishedName','Office','RecipientTypeDetails','RemoteRecipientType','IsDirSynced','ImmutableId','ExternalDirectoryObjectId','CustomAttribute5','EmailAddressPolicyEnabled' ;
            $XMFedProps = 'samaccountname','windowsemailaddress','DistinguishedName','Office','RecipientTypeDetails','RemoteRecipientType','ImmutableId','ExternalDirectoryObjectId','CustomAttribute5','EmailAddressPolicyEnabled' ; ;
            $lProps = @{Name='HasLic'; Expression={$_.IsLicensed }},@{Name='LicIssue'; Expression={$_.LicenseReconciliationNeeded }} ;
            $adprops = 'samaccountname','UserPrincipalName','distinguishedname','Description','title','whenCreated','whenChanged' ; 
            $aaduprops = 'UserPrincipalName','name','ImmutableId','DirSyncEnabled','LastDirSyncTime' ;
            $aaduFedProps = 'UserPrincipalName','name','ImmutableId','DirSyncEnabled','LastDirSyncTime' ;
            $RcpPropsTbl = 'Alias','PrimarySmtpAddress','RecipientType','RecipientTypeDetails' ; 

            $rgxOPLic = '^CN\=ENT\-APP\-Office365\-(EXOK|F1|MF1)-DL$' ; 
            $rgxXLic = '^CN\=ENT\-APP\-Office365\-(EXOK|F1|MF1)-DL$' ; 
            write-host -foreground yellow "get-Rmbx/xMbx: " -nonewline;

            # $isEml=$isDname=$isSamAcct=$false ; 
            $pltgM=[ordered]@{
                ResultSize = $MaxRecips ; 
            } ; 
            if($isEml -OR $isSamAcct){
                write-verbose "processing:'identity':$($usr)" ; 
                $pltgM.add('identity',$usr) ;
            } ; 
            if($isDname){
                # interestinb bug: switched to $hSum.dname: ISE is fine, but ConsoleHost fails to expand the $fltr properly. 
                # standard is: Variables: Enclose variables that need to be expanded in single quotation marks (for example, '$User'). Don't use curly-brackets (impedes expansion)
                # workaround: looks like have to proxy the $hsum.Dname, to provide a single non-dotted variable name
                $dname = $hSum.dname
                $fltr = "displayname -like '$dname'" ; 
                write-verbose "processing:'filter':$($fltr)" ; 
                $pltgM.add('filter',$fltr) ;
            } ; 

            $error.clear() ;
            
            #write-verbose "get-[exo]Recipient w`n$(($pltgM|out-string).trim())" ; 
            #write-verbose "get-recipient w`n$(($pltgM|out-string).trim())" ; 
            # exclude contacts, they don't represent real onprem mbx assoc, and we need to refer those to EXO mbx qry anyway.
            write-verbose "get-recipient w`n$(($pltgM|out-string).trim())" ; 
            rx10 -Verbose:$false -silent ;
            if($hSum.OPRcp=get-recipient @pltgM -ea 0 |?{$_.recipienttypedetails -ne 'MailContact'}){
                write-verbose "`$hSum.OPRcp found" ; 
            } elseif($isDname -and $hsum.lname) { 
                $smsg = "Failed:RETRY: detected 'LName':$($hsum.lname) for near matches..." ; 
                write-host $smsg ;
                $lname = $hsum.lname ;
                $fltrB = "displayname -like '*$lname*'" ; 
                write-verbose "RETRY:get-recipient -filter {$($fltr)}" ; 
                if($hSum.OPRcp=get-recipient -filter $fltr -ea 0 |?{$_.recipienttypedetails -ne 'MailContact'}){
                    write-verbose "`$hSum.OPRcp found" ;     
                } ;
            };

            if(!$hsum.OpRcp){
                $smsg = "Failed to get-recipient on:$($usr)"
                if($isDname){$smsg += " or *$($hsum.lname )*"}
                write-host $smsg ;
            } else { 
                write-verbose "`$hSum.OPRcp:`n$(($hSum.OPRcp|out-string).trim())" ; 
            } ; 


            write-verbose "get-exorecipient w`n$(($pltgM|out-string).trim())" ; 
            rxo  -Verbose:$false -silent ; 
            if($hSum.xoRcp=get-exorecipient @pltgM -ea 0 ){
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
                            if($hSum.OPMailbox=get-mailbox $hSum.OPRcp.identity -resultsize $MaxRecips){ ;
                                write-verbose "`$hSum.OPMailbox:`n$(($hSum.OPMailbox|out-string).trim())" ; 
                                if($outObject){

                                } else { 
                                    $Rpt += $hSum.OPMailbox.primarysmtpaddress ; 
                                } ; 
                                write-verbose "'UserMailbox':get-mailbox $($hSum.OPMailbox.userprincipalname)"
                                $hSum.OPMapiTest = Test-MAPIConnectivity -identity $hSum.OPMailbox.userprincipalname ; 
                                write-host -foreground yellow "Outlook (MAPI) Access Test Result:$($hsum.OPMapiTest.result)" ;
                            } ; 
                        } 
                        "MailUser" {
                            write-verbose "'MailUser':get-remotemailbox $($hSum.OPRcp.identity)"
                            $hSum.OPRemoteMailbox=get-remotemailbox $hSum.OPRcp.identity -resultsize $MaxRecips  ;
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
                    $pltGadu=[ordered]@{Identity = $null ; Properties=$adprops ;errorAction='STOP' ; } ;
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
                        
                        write-verbose "`$hSum.ADUser:`n$(($hSum.ADUser|fl $adprops| out-string).trim())" ; 
                        $smsg = "(TOR USER, fed:ad.toro.com)" ; 
                        $hSum.Federator = 'ad.toro.com' ; 
                        write-host -Fore yellow $smsg ; 
                        if($hSum.OPRemoteMailbox){
                            $smsg = "$(($hSum.OPRemoteMailbox |fl $xMProps|out-string).trim())"
                            #$smsg += "`n-Title:$($hSum.ADUser.Title)" 
                            $smsg += "`n$(($hSum.ADUser |fl 'Enabled','Description','whenCreated','whenChanged','Title' |out-string).trim())"
                        } ; 
                        if($hSum.OPMailbox){
                            $smsg =  "$(($hSum.OPMailbox |fl $xMProps|out-string).trim())" ; 
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
            }elseif($hSum.xoRcp){
                foreach($txR in $hSum.xoRcp){
                    TRY {
                        switch -regex ($txR.recipienttypedetails){
                            "UserMailbox" {
                                write-verbose "get-exomailbox w`n$(($pltgM|out-string).trim())" ; 
                                if($hSum.xoMailbox=get-exomailbox @pltgM -ea 0){
                                    write-verbose "`$hSum.xoMailbox:`n$(($hSum.xoMailbox|out-string).trim())" ; 
                                    if($outObject){

                                    } else { 
                                        $Rpt += $hSum.xoMailbox.primarysmtpaddress ; 
                                    } ; 
                                    write-verbose "'xoUserMailbox':Test-exoMAPIConnectivity $($hSum.xoMailbox.userprincipalname)"
                                    $hSum.xoMapiTest = Test-exoMAPIConnectivity -identity $hSum.xoMailbox.userprincipalname ; 
                                    write-host -foreground yellow "Outlook (xoMAPI) Access Test Result:$($hsum.xoMapiTest.result)" ;
                                    break ; 
                                } ; 
                            } 
                            "MailUser" {
                                # external mail recipient, *not* in TTC - likely in other rgs, and migrated to remote EXOP enviro
                                #$hSum.OPRemoteMailbox=get-remotemailbox $txR.identity  ;
                                caad -silent -verbose:$false ; 
                                write-verbose "`$txR | get-exoMailuser..." ;
                                $hSum.xoMUser = $txR | get-exoMailuser -ResultSize $MaxRecips ;
                                write-verbose "`$txR | get-exouser..." ;
                                $hSum.xoUser = $txR | get-exouser -ResultSize $MaxRecips ;
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
                                caad -verbose:$false ; 
                                write-verbose "`$txR | get-exouser..." ;
                                $hSum.xoUser = $txR | get-exouser -ResultSize $MaxRecips ;
                                write-verbose "`$hSum.xoUser:`n$(($hSum.xoUser|out-string).trim())" ;
                                write-verbose "get-AzureAdUser  -objectid $($hSum.xoUser.userPrincipalName)" ; 
                                $hSum.txGuest = get-AzureAdUser  -objectid $hSum.xoUser.userPrincipalName -Top $MaxRecips ;
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
                        #$hSum.ADUser =Get-ADUser @pltGadu ;
                        write-verbose "Get-ADUser w`n$(($pltGadu|out-string).trim())" ; 
                        # try a nested local trycatch, against a missing result
                        Try {
                            #Get-ADUser $DN -ErrorAction Stop ; 
                            $hSum.ADUser =Get-ADUser @pltGadu ;
                        } Catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
                            write-warning "(no matching ADuser found:$($pltGadu.identity))" ; 
                        } catch {
                            $ErrTrapd=$Error[0] ;
                            $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                            else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            Continue ;
                        } ; 

                        write-verbose "`$hSum.ADUser:`n$(($hSum.ADUser|fl $adprops | out-string).trim())" ;
                        $smsg = "(TOR USER, fed:ad.toro.com)" ;
                        $hSum.Federator = 'ad.toro.com' ;  
                        write-host -Fore yellow $smsg ; 
                        if($hSum.OPRemoteMailbox){
                            $smsg = "$(($hSum.OPRemoteMailbox |fl $xMProps|out-string).trim())"
                            #$smsg += "`n-Title:$($hSum.ADUser.Title)" 
                            $smsg += "`n$(($hSum.ADUser |fl 'Enabled','Description','whenCreated','whenChanged','Title' |out-string).trim())"
                        } ; 
                        if($hSum.OPMailbox){
                            $smsg =  "$(($hSum.OPMailbox |fl $xMProps|out-string).trim())" ; 
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
                    write-host "=get-xMbx:>`n$(($hSum.xoMailbox |fl ($xMprops |?{$_ -notmatch '(sam.*|dist.*)'})|out-string).trim())`n-Title:$($hSum.xoUser.Title)";
                }elseif($hSum.xoMUser){
                    write-host "=get-xMUSR:>`n$(($hSum.xoMUser |fl ($xMprops |?{$_ -notmatch '(sam.*|dist.*)'})|out-string).trim())`n-Title:$($hSum.xoUser.Title)";
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
                    if($hSum.Rcp=get-recipient -id $substring -ea 0 -ResultSize $MaxRecips |?{$_.recipienttypedetails -ne 'MailContact'}){
                        #$hSum.Rcp | write-output ;
                        # $RcpPropsTbl 
                        write-host -foregroundcolor yellow "`n$(($hSum.Rcp | ft -a $RcpPropsTbl |out-string).trim())" ; 
                    } ;  
                    write-host "get-exorecipient -id $($substring) -ea 0 |?{$_.recipienttypedetails -ne 'MailContact'} : " 
                    if($hSum.xoRcp=get-exorecipient -id $substring -ea 0 -ResultSize $MaxRecips |?{$_.recipienttypedetails -ne 'MailContact'}){
                        #$hSum.xoRcp | write-output ;
                        write-host -foregroundcolor yellow "`n$(($hSum.xoRcp | ft -a $RcpPropsTbl |out-string).trim())" ; 
                    } ;  


                } ; 


            } ; # don't break, doesn't continue loop

            # 10:42 AM 9/9/2021 force populate the xoMailbox, ALWAYS - need for xbrain ids
            if($hSum.xoRcp.recipienttypedetails -eq 'UserMailbox' -AND -not($hSum.xoMailbox)){
                write-verbose "get-exomailbox w`n$(($pltgM|out-string).trim())" ; 
                if($hSum.xoMailbox=get-exomailbox @pltgM -ea 0){
                    write-verbose "`$hSum.xoMailbox:`n$(($hSum.xoMailbox|out-string).trim())" ; 
                    write-verbose "'xoUserMailbox':Test-exoMAPIConnectivity $($hSum.xoMailbox.userprincipalname)"
                    $hSum.xoMapiTest = Test-exoMAPIConnectivity -identity $hSum.xoMailbox.userprincipalname ; 
                    write-host -foreground yellow "Outlook (xoMAPI) Access Test Result:$($hsum.xoMapiTest.result)" ;
                } ; 
            } ;

            #$pltgMU=@{UserPrincipalName=$null ; MaxResults= $MaxRecips; ErrorAction= 'STOP' } ; 
            # maxresults is documented: 
            # but causes a fault with no $error[0], doesn't seem to be functional param, post-filter
            $pltgMU=@{UserPrincipalName=$null ; ErrorAction= 'STOP' } ; 
            if($hSum.ADUser){$pltgMU.UserPrincipalName = $hSum.ADUser.UserPrincipalName } 
            elseif($hSum.xoMailbox){$pltgMU.UserPrincipalName = $hSum.txMbx.UserPrincipalName }
            elseif($hSum.xoMUser){$pltgMU.UserPrincipalName = $hSum.xoMUser.UserPrincipalName }
            elseif($hSum.txGuest){$pltgMU.UserPrincipalName = $hSum.txGuest.userprincipalname } 
            else{} ; 
            
            if($pltgMU.UserPrincipalName){
                write-host -foregroundcolor yellow "=get-msoluser $($pltgMU.UserPrincipalName):(licences)>:" ;
                TRY{
                    cmsol  -Verbose:$false -silent ;
                    write-verbose "get-msoluser w`n$(($pltgMU|out-string).trim())" ; 
                    # have to postfilter, if want specific count -maxresults catch's with no $error[0]
                    $hSum.MsolUser=get-msoluser @pltgMU | select -first $MaxRecips;  ;
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
                    write-host -foregroundcolor yellow "=get-AADuser $($pltgMU.UserPrincipalName):(licences)>:" ;
                    TRY{
                        caad  -Verbose:$false -silent ;
                        write-verbose "get-AzureAdUser  -objectid $($pltgMU.UserPrincipalName)" ; 
                        # have to postfilter, if want specific count -maxresults catch's with no $error[0]
                        $hSum.AADUser  = get-AzureAdUser  -objectid $pltgMU.UserPrincipalName  | select -first $MaxRecips;  ;
                        #write-verbose "`$hSum.AADUser:`n$(($hSum.AADUser|out-string).trim())" ;
                        # ObjectId                             DisplayName   UserPrincipalName      UserType
                        if(-not($hSum.ADUser)){
                            $smsg = "$(($hSum.OPRemoteMailbox |fl $xMProps|out-string).trim())"
                            "$(($hSum.ADUser |fl $xMProps|out-string).trim())"
                        } else { 
                            write-verbose "`$hSum.AADUser:`n$(($hSum.AADUser| ft -auto ObjectId,DisplayName,UserPrincipalName,UserType |out-string).trim())" ;
                        } ; 
                    } CATCH {
                        $ErrTrapd=$Error[0] ;
                        $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                        else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        Continue ; 
                    } ; 

                } ;

                $smsg = "$(($hSum.MsolUser|fl $lProps|out-string).trim())`n" ;
                $smsg += "Licenses Assigned:$(($hSum.MsolUser.licenses.AccountSkuId -join '; '|out-string).trim())" ; 
                write-host $smsg ; 
                if($hSum.ADUser){$hSum.LicenseGroup = $hSum.ADUser.memberof |?{$_ -match $rgxOPLic }}
                elseif($hSum.xoMemberOf){$hSum.LicenseGroup = $hSum.xoMemberOf.Name |?{$_ -match $rgxXLic}}
                if(!($hSum.LicenseGroup) -AND ($hSum.MsolUser.licenses.AccountSkuId -contains 'toroco:ENTERPRISEPACK')){$hSum.LicenseGroup = '(direct-assigned E3)'} ; 
                if($hSum.LicenseGroup){$smsg = "LicenseGroup:$($hSum.LicenseGroup)"}
                else{$smsg = "LicenseGroup:(unresolved, direct-assigned other?)" } ; 
                write-host $smsg ; 



            } ; 
            
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
        
     }
 }

#*------^ resolve-user.ps1 ^------