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
    PS> .\resolve-user.ps1
    Default, attempts to parse a user descriptor from clipboard
    .EXAMPLE
    PS> .\resolve-user.ps1 -users 'John Public' 
    Process user displayname
    .EXAMPLE
    PS> .\resolve-user.ps1 -users 'Todd.Kadrie@toro.com','Stacy Sotelo','lynctest1','lforsythe@charlesmachine.works','confroom-b3-trenchingrm@toroco.onmicrosoft.com' -verbose ;
    Process an array of descriptors
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
        [switch] $useEXOv2
    ) ;
    BEGIN{
        $Verbose = ($VerbosePreference -eq 'Continue') ; 
        $rgxEmailAddr = "^([0-9a-zA-Z]+[-._+&'])*[0-9a-zA-Z]+@([-0-9a-zA-Z]+[.])+[a-zA-Z]{2,63}$" ; 
        $rgxDName = "^([a-zA-Z]{2,}\s[a-zA-Z]{1,}'?-?[a-zA-Z]{2,}\s?([a-zA-Z]{1,})?)" ; 
        $rgxSamAcctNameTOR = "^\w{2,20}$" ; # up to 20k, the limit prior to win2k
        #$rgxSamAcctName = "^[^\/\\\[\]:;|=,+?<>@�]+$" # no char limit ;

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
            $fname = $lname = $dname = $tRcp = $tMbx = $txRMbx = $tADU = $txMbx = $txU = $xMmbrOf = $mu = $licgrp = $null ; 
            $isEml=$isDname=$isSamAcct=$false ; 
            $procd++ ; 
            write-verbose "processing:$($usr)" ; 
            switch -regex ($usr){
                $rgxEmailAddr {
                    $fname,$lname = $usr.split('@')[0].split('.') ; 
                    $dname = $usr ;
                    write-verbose "(detected user ($($usr)) as EmailAddr)" ; 
                    $isEml = $true ;
                } 
                $rgxDName {
                    $fname,$lname = $usr.split(' ') ;
                    $dname = $usr ; 
                    write-verbose "(detected user ($($usr)) as DisplayName)" ; 
                    $isDname = $true ;
                } 
                $rgxSamAcctNameTOR {
                    $lname = $usr ; 
                    write-verbose "(detected user ($($usr)) as SamAccountName)" ; 
                    $isSamAcct  = $true ;
                } 
                default {
                    write-warning "$((get-date).ToString('HH:mm:ss')):No -user specified, nothing matching dname, emailaddress or samaccountname, found on clipboard. EXITING!" ; 
                    #Break ; 
                } ; 
            } ; 

            $sBnr="===v ($($Procd)/$($ttl)):Input: '$($usr)' | '$($fname)' | '$($lname)' v===" ;
            if($isEml){$sBnr+="(EML)"}
            elseif($isDname){$sBnr+="(DNAM)"}
            elseif($isSamAcct){$sBnr+="(SAM)"}
            write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($sBnr)" ;
        
            $xMProps="samaccountname","windowsemailaddress","DistinguishedName","Office","RecipientTypeDetails" ;
            $lProps = @{Name='HasLic'; Expression={$_.IsLicensed }},@{Name='LicIssue'; Expression={$_.LicenseReconciliationNeeded }} ;
            $adprops = "samaccountname","UserPrincipalName","distinguishedname" ; 
            
            $rgxOPLic = '^CN\=ENT\-APP\-Office365\-(EXOK|F1|MF1)-DL$' ; 
            $rgxXLic = '^CN\=ENT\-APP\-Office365\-(EXOK|F1|MF1)-DL$' ; 
            write-host -foreground yellow "get-Rmbx/xMbx: " -nonewline;

            # $isEml=$isDname=$isSamAcct=$false ; 
            $pltgM=[ordered]@{} ; 
            if($isEml -OR $isSamAcct){
                write-verbose "processing:'identity':$($usr)" ; 
                $pltgM.add('identity',$usr) ;
            } ; 
            if($isDname){
                $fltr = "displayname -like '$dname'" ; 
                write-verbose "processing:'filter':$($fltr)" ; 
                $pltgM.add('filter',$fltr) ;
            } ; 

            $error.clear() ;
            
            #write-verbose "get-[exo]Recipient w`n$(($pltgM|out-string).trim())" ; 
            #write-verbose "get-recipient w`n$(($pltgM|out-string).trim())" ; 
            # exclude contacts, they don't represent real onprem mbx assoc, and we need to refer those to EXO mbx qry anyway.
            write-verbose "get-recipient w`n$(($pltgM|out-string).trim())" ; 
            $tRcp=get-recipient @pltgM -ea 0 |?{$_.recipienttypedetails -ne 'MailContact'}
            write-verbose "get-exorecipient w`n$(($pltgM|out-string).trim())" ; 
            $txRcp=get-exorecipient @pltgM -ea 0 ;

            write-verbose "`$tRcp:`n$(($tRcp|out-string).trim())" ; 
            write-verbose "`$txRcp:`n$(($txRcp|out-string).trim())" ; 

            if($tRcp){
                $error.clear() ;
                TRY {
                    switch -regex ($tRcp.recipienttype){
                        "UserMailbox" {
                            write-verbose "'UserMailbox':get-mailbox $($tRcp.identity)"
                            $tMbx=get-mailbox $tRcp.identity ;
                            write-verbose "`$tMbx:`n$(($tMbx|out-string).trim())" ; 
                            $Rpt += $tMbx.primarysmtpaddress ; 
                        } 
                        "MailUser" {
                            write-verbose "'MailUser':get-remotemailbox $($tRcp.identity)"
                            $txRMbx=get-remotemailbox $tRcp.identity  ;
                            write-verbose "`$txRMbx:`n$(($txRMbx|out-string).trim())" ; 
                            $Rpt += $txRMbx.primarysmtpaddress ;  
                        } ;
                        default {
                            write-warning "$((get-date).ToString('HH:mm:ss')):Unsupported RecipientType:($tRcp.recipienttype). EXITING!" ; 
                            Break ; 
                        }
                    }
                    $pltGadu=[ordered]@{Identity = $null ; Properties='*' ;errorAction='STOP'} ;
                    if($txRMbx ){
                        $pltGadu.identity = $txRmbx.samaccountname;
                    }elseif($tMbx){
                        $pltGadu.identity = $tmbx.samaccountname ;
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
                        #$tADU =Get-ADUser @pltGadu ;
                        write-verbose "Get-ADUser w`n$(($pltGadu|out-string).trim())" ; 
                        # try a nested local trycatch, against a missing result
                        Try {
                            #Get-ADUser $DN -ErrorAction Stop ; 
                            $tADU =Get-ADUser @pltGadu ;
                        } Catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
                            write-warning "(no matching ADuser found:$($pltGadu.identity))" ; 
                        } catch {
                            $ErrTrapd=$Error[0] ;
                            $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                            else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            Continue ;
                        } ; 
                        
                        write-verbose "`$tADU:`n$(($tADU|fl $adprops| out-string).trim())" ; 
                        $smsg = "(TOR USER, fed:ad.toro.com)" ; 
                        write-host -Fore yellow $smsg ; 
                        if($txRMbx){$smsg = "$(($txRMbx |fl $xMProps|out-string).trim())`n-Title:$($tADU.Title)" } ; 
                        if($tMbx){$smsg =  "$(($tMbx |fl $xMProps|out-string).trim())`n-Title:$($tADU.Title)" } ; 
                        write-host $smsg ;
                    } ; 
                } CATCH {
                    $ErrTrapd=$Error[0] ;
                    $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    Continue ; 
                } ; 
            #}elseif($txMbx){
            }elseif($txRcp){
                foreach($txR in $txRcp){
                    TRY {
                        switch -regex ($txR.recipienttypedetails){
                            "UserMailbox" {
                                write-verbose "get-exomailbox w`n$(($pltgM|out-string).trim())" ; 
                                $txMbx=get-exomailbox @pltgM -ea 0 ;
                                write-verbose "`$txMbx:`n$(($txMbx|out-string).trim())" ; 
                                $Rpt += $txMbx.primarysmtpaddress ; 
                                break ; 
                            } 
                            "GuestMailUser" {
                                #$txRMbx=get-remotemailbox $txR.identity  ;
                                caad -verbose:$false ; 
                                get-verbose "`$txR | get-exouser..." ;
                                $txU = $txR | get-exouser ;
                                write-verbose "`$txU:`n$(($txU|out-string).trim())" ;
                                write-verbose "get-AzureAdUser  -objectid $($txU.userPrincipalName)" ; 
                                $txGuest = get-AzureAdUser  -objectid $txU.userPrincipalName ;
                                write-verbose "`$txGuest:`n$(($txGuest|out-string).trim())" ;
                                #$Rpt += $txRMbx.primarysmtpaddress ;  
                                write-host "$($txR.ExternalEmailAddress): matches a Guest object with UPN:$($txU.userPrincipalName)" ; 
                                if($txGuest.EmailAddresses -eq $null){
                                    write-warning "Guest appears to have damage from conficting replicated onprem MailContact, as it's EmailAddresses property is *blank*" ; 
                                } ; 
                                break ; 
                            } ;
                            "MailContact" {
                                #$txRMbx=get-remotemailbox $txR.identity  ;
                                #$Rpt += $txRMbx.primarysmtpaddress ;  
                                write-host "$($txR.primarysmtpaddress): matches an EXO MailContact with external Email: $($txR.primarysmtpaddress)" ; 
                                break ; 
                            } ;
                            "MailUniversalSecurityGroup" {
                                #$txRMbx=get-remotemailbox $txR.identity  ;
                                #$Rpt += $txRMbx.primarysmtpaddress ;  
                                write-host "$($txR.primarysmtpaddress): matches an EXO MailUniversalSecurityGroup with Dname: $($txR.displayname)" ; 
                                break ; 
                            } ;
                            default {
                                write-warning "$((get-date).ToString('HH:mm:ss')):Unsupported RecipientType:($tRcp.recipienttype). EXITING!" ; 
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
                # contacts and guests won't drop with $txRmbx or $tmbx populated
                TRY {
                    $pltGadu=[ordered]@{Identity = $null ; Properties='*' ;errorAction='SilentlyContinue'} ;
                    if($txRMbx ){
                        $pltGadu.identity = $txRmbx.samaccountname;
                    }elseif($tMbx){
                        $pltGadu.identity = $tmbx.samaccountname ;
                    } ; 
                    if($pltGadu.identity){
                        write-verbose "Get-ADUser w`n$(($pltGadu|out-string).trim())" ; 
                        #$tADU =Get-ADUser @pltGadu ;
                        write-verbose "Get-ADUser w`n$(($pltGadu|out-string).trim())" ; 
                        # try a nested local trycatch, against a missing result
                        Try {
                            #Get-ADUser $DN -ErrorAction Stop ; 
                            $tADU =Get-ADUser @pltGadu ;
                        } Catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
                            write-warning "(no matching ADuser found:$($pltGadu.identity))" ; 
                        } catch {
                            $ErrTrapd=$Error[0] ;
                            $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                            else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            Continue ;
                        } ; 

                        write-verbose "`$tADU:`n$(($tADU|fl $adprops | out-string).trim())" ;
                        $smsg = "(TOR USER, fed:ad.toro.com)" ; 
                        write-host -Fore yellow $smsg ; 
                        if($txRMbx){$smsg = "$(($txRMbx |fl $xMProps|out-string).trim())`n-Title:$($tADU.Title)" } ; 
                        if($tMbx){$smsg =  "$(($tMbx |fl $xMProps|out-string).trim())`n-Title:$($tADU.Title)" } ; 
                        write-host $smsg ;
                    } ; 
                } CATCH {
                    $ErrTrapd=$Error[0] ;
                    $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    Continue ; 
                } ; 

                $Rpt += $txMbx.primarysmtpaddress ; 
                if($txMbx.isdirsynced){$smsg="(CMW USER, fed:cmw.internal)"}
                else{$smsg="(CLOUD-1ST ACCT, unfederated)"} ;
                write-host -Fore yellow $smsg ; 
                # skip user lookup if guest already pulled it 
                if(!$txU){
                    write-verbose "get-exouser -id $($txmbx.UserPrincipalName)"
                    $txU = get-exouser -id $txmbx.UserPrincipalName ; 
                    write-verbose "`$txU:`n$(($txU|out-string).trim())" ;
                } 
                if($txMbx){
                    write-host "=get-xMbx:>`n$(($txMbx |fl ($xMprops |?{$_ -notmatch '(sam.*|dist.*)'})|out-string).trim())`n-Title:$($txU.Title)";
                }elseif($txGuest){
                    write-host "=get-AADU:>`n$(($txGuest |fl userp*,PhysicalDeliveryOfficeName,JobTitle|out-string).trim())"
                } ; 
                TRY {
                    write-verbose "Get-exoRecipient -Filter {Members -eq '$($txU.DistinguishedName)'} -RecipientTypeDetails GroupMailbox,MailUniversalDistributionGroup,MailUniversalSecurityGroup"
                    $xMmbrOf = Get-exoRecipient -Filter "Members -eq '$($txU.DistinguishedName)'" -RecipientTypeDetails GroupMailbox,MailUniversalDistributionGroup,MailUniversalSecurityGroup ;
                    write-verbose "`$xMmbrOf:`n$(($xMmbrOf|out-string).trim())" ;
                } CATCH {
                    $ErrTrapd=$Error[0] ;
                    $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    Continue ; 
                } ; 
            } else { write-warning "(no matching EXOP or EXO recipient object:$($usr))"   } ; # don't break, doesn't continue loop

            $pltgMU=@{UserPrincipalName=$null} ; 
            if($tADU){$pltgMU.UserPrincipalName = $tADU.UserPrincipalName } 
            elseif($txMbx){$pltgMU.UserPrincipalName = $txMbx.UserPrincipalName }
            elseif($txGuest){$pltgMU.UserPrincipalName = $txGuest.userprincipalname } 
            else{} ; 
            
            if($pltgMU.UserPrincipalName){
                write-host -foregroundcolor yellow "=get-msoluser $($pltgMU.UserPrincipalName):(licences)>:" ;
                TRY{
                    write-verbose "get-msoluser w`n$(($pltgMU|out-string).trim())" ; 
                    $mu=get-msoluser @pltgMU ;
                    write-verbose "`$mu:`n$(($mu|out-string).trim())" ;
                } CATCH {
                    $ErrTrapd=$Error[0] ;
                    $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    Continue ; 
                } ; 
                $smsg = "$(($mu|fl $lProps|out-string).trim())`n" ;
                $smsg += "Licenses Assigned:$(($mu.licenses.AccountSkuId -join '; '|out-string).trim())" ; 
                write-host $smsg ; 
                if($tadu){$licgrp = $tadu.memberof |?{$_ -match $rgxOPLic }}
                elseif($xMmbrOf){$licgrp = $xMmbrOf.Name |?{$_ -match $rgxXLic}}
                if(!($licgrp) -AND ($mu.licenses.AccountSkuId -contains 'toroco:ENTERPRISEPACK')){$licgrp = '(direct-assigned E3)'} ; 
                if($licgrp){$smsg = "LicGrp:$($licgrp)"}
                else{$smsg = "LicGrp:(unresolved, direct-assigned other?)" } ; 
                write-host $smsg ; 
            } ; 
            write-host -foregroundcolor green $sBnr.replace('=v','=^').replace('v=','^=') ; 
        } ; 
    }
    END{
        $Rpt -join ',' | out-clipboard ; 
        write-host "(output copied to clipboard)"
        $Rpt -join ',' | write-output ;
     }
 }

#*------^ resolve-user.ps1 ^------