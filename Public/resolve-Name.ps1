#*------v resolve-Name.ps1 v------
Function resolve-Name {
    <#
    .SYNOPSIS
    resolve-Name.ps1 - Port 7nlu to a verb-EXO function. Resolves a displayname into Exchange Online/Exchange Onprem mailbox/MsolUser/AzureADUser/ADUser info, and licensing status. Detect's cross-org hybrid AD objects as well. 
    .NOTES
    Version     : 1.0.0
    Author      : Todd Kadrie
    Website     : http://www.toddomation.com
    Twitter     : @tostka / http://twitter.com/tostka
    CreatedDate : 2021-06-09
    FileName    : resolve-Name.ps1
    License     : MIT License
    Copyright   : (c) 2020 Todd Kadrie
    Github      : https://github.com/tostka/verb-EXO
    Tags        : Powershell,ExchangeOnline,Exchange,MsolUser,AzureADUser,ADUser
    REVISIONS
    * 1:17 PM 6/10/2021 added missing $exMProps add lic grp memberof check for aadu, for x-hyb users; add missing $rgxLicGrp, as $rgxLicGrpDN & $rgxLicGrpDName (aduser & aaduser respectively); pulled datestamps on echo's, simplified echo's (removed "$($smsg)")
    * 4:00 PM 6/9/2021 added alias 'nlu' (7nlu is still ahk macro) ; fixed typo; expanded echo for $lic;flipped -displayname to -identifier, and handle smtpaddr|alias|displayname lookups ; init; 
    .DESCRIPTION
    resolve-Name.ps1 - Port 7nlu to a verb-EXO function. Resolves a mailbox user Identifier into Exchange Online/Exchange Onprem mailbox/MsolUser/AzureADUser info, and licensing status. Detect's cross-org hybrid AD objects as well. 
    .PARAMETER TenOrg
    Tenant Org designator (defaults to TOR)
    .PARAMETER Identifier
    User Displayname|UPN|alias to be resolved[-Identifier 'Some Username'
    .PARAMETER Ticket
    Ticket # [-Ticket nnnnn]
    .PARAMETER useEXOv2
    Use EXOv2 (ExchangeOnlineManagement) over basic auth legacy connection [-useEXOv2]
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    .EXAMPLE
    resolve-Name -Identifier 'Some User'
    Command-line resolve displayname to summary details.
    .EXAMPLE
    resolve-Name -Identifier 'Some.User@domain.com'
    Command-line resolve email address to summary details.
    .EXAMPLE
    resolve-Name -Identifier 'alias'
    Command-line resolve mail alias value to summary details.
    .EXAMPLE
    resolve-Name
    Where no -Identifier is specified, defaults to checking clipboard for a Identifier equivelent.
    .LINK
    https://github.com/tostka/verb-EXO
    #>
    #Requires -Modules ActiveDirectory,AzureAD,MSOnline,verb-Auth,verb-IO,verb-Mods,verb-Text,verb-AAD,verb-ADMS,verb-Ex2010,verb-logging
    [CmdletBinding()]
    [Alias('nlu')]
    PARAM(
        [Parameter(Mandatory=$FALSE,HelpMessage="TenantTag value, indicating Tenants to connect to[-TenOrg 'TOR']")]
        $TenOrg = 'TOR',
        [Parameter(Position=0,Mandatory=$False,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,HelpMessage="User Identifier to be resolved[-Identifier 'Some Username'")]        
        $Identifier,
        [Parameter(HelpMessage="Use EXOv2 (ExchangeOnlineManagement) over basic auth legacy connection [-useEXOv2]")]
        [switch] $useEXOv2
    ) ;
    BEGIN {
        $Verbose = ($VerbosePreference -eq 'Continue') ;
        
        #$propsXmbx = 'UserPrincipalName','Alias','ExchangeGuid','Database','ExternalDirectoryObjectId','RemoteRecipientType'
        #$propsOPmbx = 'UserPrincipalName','SamAccountName','RecipientType','RecipientTypeDetails' ; 
        $exMProps='samaccountname','alias','windowsemailaddress','DistinguishedName''RecipientType','RecipientTypeDetails' ;

        #$adprops = "samaccountname", "msExchRemoteRecipientType", "msExchRecipientDisplayType", "msExchRecipientTypeDetails", "userprincipalname" ;
        $adprops = "samaccountname","UserPrincipalName","memberof","msExchMailboxGuid","msexchrecipientdisplaytype","msExchRecipientTypeDetails","msExchRemoteRecipientType"
        
        [regex]$rgxDname = "^[\w'\-,.][^0-9_!?????/\\+=@#$%?&*(){}|~<>;:[\]]{2,}$"
        # below doesn't encode cleanly, mainly black diamonds - better w alt font (non-lucida console)
        #"^[a-zA-Z������acce����ei����ln����������uu��zz��c��������ACCEE��������ILN����������UU��ZZ��ǌ�C��?� ,.'-]+$"
        [regex]$rgxEmailAddr = "^([0-9a-zA-Z]+[-._+&'])*[0-9a-zA-Z]+@([-0-9a-zA-Z]+[.])+[a-zA-Z]{2,63}$"
        [regex]$rgxCMWDomain = 'DC=cmw,DC=internal$' ;
        [regex]$rgxExAlias = "^[0-9a-zA-Z-._+&]{1,64}$" ;
        # used for adu.memberof
        [regex]$rgxLicGrpDN = "^CN=ENT-APP-Office365-.*-DL,((OU=Enterprise\sApplications,)*)OU=ENTERPRISE,DC=global,DC=ad,DC=toro((lab)*),DC=com$" ;  ; 
        # used for taadu memberof
        [regex]$rgxLicGrpDName = "^ENT-APP-Office365-.*-DL((\s)*)$" ;
        #"^ENT-APP-Office365-.*-DL$" ;  
        # cute, we've got cmw AAD grps with trailing spaces: 'ENT-APP-Office365-CMWUsers-E3-DL ', pull trailing $

        if(!$Identifier -AND (gcm get-clipboard) -AND (get-clipboard)){
            $Identifier = get-clipboard ;
            #$cb = get-clipboard ; 
        } elseif($Identifier){


        } else {
            write-warning "No Identifier specified, and clipboard did not match 'Identifier' content" ; 
            Break ;
        } ; 

        <#[regex]$rgxDname = "^[\w'\-,.][^0-9_!?????/\\+=@#$%?&*(){}|~<>;:[\]]{2,}$"
        # below doesn't encode cleanly, mainly black diamonds
        #"^[a-zA-Z??????acce????ei????ln??????????uu??zz??c????????ACCEE????????ILN??????????UU??ZZ????C???? ,.'-]+$"
        [regex]$rgxEmailAddr = "^([0-9a-zA-Z]+[-._+&'])*[0-9a-zA-Z]+@([-0-9a-zA-Z]+[.])+[a-zA-Z]{2,63}$"
        #>
        $IdentifierType = $null ; 
        switch -regex ($Identifier){
            $rgxExAlias {
                write-verbose "(`$Identifier appears to be an Alias)" ;
                $IdentifierType = "Alias" ;
                #$Displayname = $Identifier.split('@')[0].replace('.',' ')
                #$nameparts = $Identifier.split('@')[0].replace('.',' ').split(' ')
                $nameparts = $Identifier.split(' ')
                break;
            } ;
            $rgxDname {
                write-verbose "(`$Identifier appears to be a DisplayName)" ;
                $IdentifierType = "DisplayName" ;
                $nameparts = $Identifier.split(' ')
                break;
            }
            $rgxEmailAddr {
                write-verbose "(`$Identifier appears to be an SmtpAddress)" ;
                $IdentifierType = "SmtpAddress" ;
                #$Displayname = $Identifier.split('@')[0].replace('.',' ')
                $nameparts = $Identifier.split('@')[0].replace('.',' ').split(' ')
                break;
            } ;
            default {
                write-warning "Unable to resolve -Identifier ($($Identifier)) into a proper DisplayName|EmailAddress|Alias string" ;
                $IdentifierType = $null ;
                break ;
            }
        } ;
        #if($Identifier -match $rgxDname){
        #        $nameparts = $Identifier.split(' ')
        switch (($nameparts|measure).count){
            "1" {
                # it's an alias
                #Identifier = vString 
                $fname = "" 
                $lname = $nameparts
            }
            "2" {
                <#/*
                RegExMatch(vString, "^\w*\s\w*$", displayname)
                RegExMatch(vString, "\w*(?=[\s])", fname)
                RegExMatch(vString, "(?<=\s)\w*$", lname)
                */
                #>
                #displayname = vString 
                $fname = $nameparts[0] ;
                $lname = $nameparts[1] ;
            }
            default{
                # assume the last 2/* are the last name ( concat no space for searches).
                #displayname = vString 
                $fname = $nameparts[0] ; 
                $lname = $nameparts[1..[int]($nameparts.getupperbound(0))] -join ' ' ;
            }
        } ;
        #} ; 
        
        $sBnr="===v Input (& splits): '$($Identifier)' | '$($fname)' | '$($lname)' v===" ;
        write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($sBnr)" ;
        #-=-=configure EXO EMS aliases to cover useEXOv2 requirements-=-=-=-=-=-=
        # have to preconnect, as it gcm's the targets
        if ($script:useEXOv2) { reconnect-eXO2 -verbose:$($verbose)}
        else { reconnect-EXO -verbose:$($verbose)} ;
        # in this case, we need an alias for EXO, and non-alias for EXOP
        [array]$cmdletMaps = 'ps1GetxRcp;get-exorecipient;','ps1GetxMbx;get-exomailbox;','ps1GetxUser;get-exoUser;'
        foreach($cmdletMap in $cmdletMaps){
            if($script:useEXOv2){
                if(!($cmdlet= Get-Command $cmdletMap.split(';')[1].replace('-exo','-xo') )){ throw "unable to gcm Alias definition!:$($cmdletMap.split(';')[1])" ; break }
                $nalias = set-alias -name ($cmdletMap.split(';')[0]) -value ($cmdlet.name) -passthru ;
                write-verbose "$($nalias.Name) -> $($nalias.ResolvedCommandName)" ;
            } else {
                if(!($cmdlet= Get-Command $cmdletMap.split(';')[1])){ throw "unable to gcm Alias definition!:$($cmdletMap.split(';')[1])" ; break }
                $nalias = set-alias -name ($cmdletMap.split(';')[0]) -value ($cmdlet.name) -passthru ;                
                write-verbose "$($nalias.Name) -> $($nalias.ResolvedCommandName)" ;
            } ;
        } ;
    
        # shifting from ps1 to a function: need updates self-name:
        ${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name ;

        #$sBnr="#*======v START PASS:$($ScriptBaseName) v======" ; 
        <#$sBnr="#*======v START PASS:$(${CmdletName}) v======" ; 
        $smsg= $sBnr ;   
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
        else{ write-host -foregroundcolor green $smsg } ;
        #>
        
        # steer all onprem code on $XXXMeta.ExOPAccessFromToro & Ex10Server values
        $UseOP=$true ; 

        $useEXO = $true ; # non-dyn setting, drives variant EXO reconnect & query code
        if($useEXO){
            #*------v GENERIC EXO CREDS & SVC CONN BP v------
            # o365/EXO creds
            <### Usage: Type defaults to SID, if not spec'd - Note: there must be a *logged in & configured *profile* 
            $o365Cred=get-TenantCredentials -TenOrg $TenOrg -verbose -userrole SID ;
            Returns a credential set for the $TenOrg Hybrid OnPrem Exchange Org
            .EXAMPLE
            $o365Cred=get-TenantCredentials -TenOrg $TenOrg -verbose -userrole CSVC ;
            Returns the CSVC Userrole credential for the $TenOrg Hybrid OnPrem Exchange Org
            .EXAMPLE
            $o365Cred=get-TenantCredentials -TenOrg $TenOrg -verbose -userrole B2BI ;
            Returns the B2BI Userrole credential for the $TenOrg Hybrid OnPrem Exchange Org
            ###>
            $o365Cred=$null ;
            <# $TenOrg is a mandetory param in this script, skip dyn resolution
            switch -regex ($env:USERDOMAIN){
                "(TORO|CMW)" {$TenOrg = $env:USERDOMAIN.substring(0,3).toupper() } ;
                "TORO-LAB" {$TenOrg = 'TOL' }
                default {
                    throw "UNRECOGNIZED `$env:USERDOMAIN!:$($env:USERDOMAIN)" ; 
                    Break ; 
                } ;
            } ; 
            #>
            if($o365Cred=(get-TenantCredentials -TenOrg $TenOrg -UserRole 'CSVC','SID' -verbose:$($verbose))){
                # make it script scope, so we don't have to predetect & purge before using new-variable
                if(get-Variable -Name cred$($tenorg) -scope Script -ea 0){remove-variable -Name cred$($tenorg) -scope Script} ; 
                New-Variable -Name cred$($tenorg) -scope Script -Value $o365Cred.cred ;
                $smsg = "Resolved $($Tenorg) `$o365cred:$($o365Cred.cred.username) (assigned to `$cred$($tenorg))" ; 
                write-verbose $smsg  ;
            } else {
                $statusdelta = ";ERROR";
                $script:PassStatus += $statusdelta ;
                set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta) ;
                $smsg = "Unable to resolve $($tenorg) `$o365Cred value!"
                write-verbose $smsg  ;
                throw "Unable to resolve $($tenorg) `$o365Cred value!`nEXIT!"
                Break ;
            } ;
            <### CALLS ARE IN FORM: (cred$($tenorg))
            $pltRXO = @{Credential = $Credential ; verbose = $($verbose) ; }
            $pltRXO = @{
                Credential = (Get-Variable -name cred$($tenorg) ).value ;
                verbose = $($verbose) ; } ; 
            if ($script:useEXOv2) { reconnect-eXO2 @pltRXO }
            else { reconnect-EXO @pltRXO } ;
            # or with Tenant-specific cred($Tenorg) lookup
            #$pltRXO creds & .username can also be used for AzureAD connections 
            Connect-AAD @pltRXO ; 
            ###>
            # configure splat for connections: (see above useage)
            $pltRXO = @{
                Credential = (Get-Variable -name cred$($tenorg) ).value ;
                verbose = $($verbose) ; } ; 
            #
            #*------^ END GENERIC EXO CREDS & SVC CONN BP ^------
        } # if-E $useEXO

        if($UseOP){
            #*------v GENERIC EXOP CREDS & SRVR CONN BP v------
            # do the OP creds too
            $OPCred=$null ;
            # default to the onprem svc acct
            $pltGHOpCred=@{TenOrg=$TenOrg ;userrole='ESVC','SID'; verbose=$($verbose)} ;
            if($OPCred=(get-HybridOPCredentials @pltGHOpCred).cred){
                # make it script scope, so we don't have to predetect & purge before using new-variable
                if(get-Variable -Name "cred$($tenorg)OP" -scope Script -ea 0){remove-variable -Name "cred$($tenorg)OP" -scope Script} ; 
                New-Variable -Name "cred$($tenorg)OP" -scope Script -Value $OPCred ;
                $smsg = "Resolved $($Tenorg) `$OPCred:$($OPCred.username) (assigned to `$cred$($tenorg)OP)" ; 
                write-verbose $smsg  ;
            } else {
                $smsg = "Unable to resolve get-HybridOPCredentials -TenOrg $($TenOrg) -userrole 'ESVC' value!"
                write-verbose $smsg  ;
                throw "Unable to resolve $($tenorg) `$OPCred value!`nEXIT!"
                Break ;
            } ;
            $smsg= "Using EXOP cred:`$cred$($tenorg)OP:$((Get-Variable -name "cred$($tenorg)OP" ).value.username)" ;  
            write-verbose $smsg  ; 
            <# CALLS ARE IN FORM: (cred$($tenorg))
            $pltRX10 = @{
                Credential = (Get-Variable -name "cred$($tenorg)OP" ).value ;
                verbose = $($verbose) ; }
            ReConnect-Ex2010XO @pltRX10 ; # cross-prem conns
            Reconnect-Ex2010 @pltRX10 ; # local org conns
            #$pltRx10 creds & .username can also be used for local ADMS connections
            #>
            $pltRX10 = @{
                Credential = (Get-Variable -name "cred$($tenorg)OP" ).value ;
                verbose = $($verbose) ; } ;     
            # TEST
        
            # defer cx10/rx10, until just before get-recipients qry
            #*------^ END GENERIC EXOP CREDS & SRVR CONN BP ^------
            # connect to ExOP X10
            if($pltRX10){
                #ReConnect-Ex2010XO @pltRX10 ;
                ReConnect-Ex2010 @pltRX10 ;
            } else { Reconnect-Ex2010 ; } ; 
        } ;  # if-E $useEXOP

        <# already confirmed in modloads
        # load ADMS
        $reqMods += "load-ADMS".split(";") ;
        if ( !(check-ReqMods $reqMods) ) { write-error "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Missing function. EXITING." ; Break ; }  ;
        #>
        write-verbose "(loading ADMS...)" ;
        # 2:12 PM 6/9/2021 load-ADMS is returning boolean, capture it
        $bRet = load-ADMS -verbose:$($verbose) ;

        if($UseOP){
            # resolve $domaincontroller dynamic, cross-org
            # setup ADMS PSDrives per tenant 
            if(!$global:ADPsDriveNames){
                $smsg = "(connecting X-Org AD PSDrives)" ;
                write-verbose $smsg  ;
                $global:ADPsDriveNames = mount-ADForestDrives -verbose:$($verbose) ;
            } ; 
            if(($global:ADPsDriveNames|measure).count){
                $useEXOforGroups = $false ; 
                $smsg = "Confirming ADMS PSDrives:`n$(($global:ADPsDriveNames.Name|%{get-psdrive -Name $_ -PSProvider ActiveDirectory} | ft -auto Name,Root,Provider|out-string).trim())" ;
                write-verbose $smsg  ;
                # returned object
                #         $ADPsDriveNames
                #         UserName                Status Name        
                #         --------                ------ ----        
                #         DOM\Samacctname   True  [forestname wo punc] 
                #         DOM\Samacctname   True  [forestname wo punc]
                #         DOM\Samacctname   True  [forestname wo punc]
        
            } else { 
                #-=-record a STATUS=-=-=-=-=-=-=
                $statusdelta = ";ERROR";
                $script:PassStatus += $statusdelta ;
                set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta) ;
                #-=-=-=-=-=-=-=-=
                $smsg = "Unable to detect POPULATED `$global:ADPsDriveNames!`n(should have multiple values, resolved to $()"
                write-warning $smsg  ;
                throw "Unable to resolve $($tenorg) `$o365Cred value!`nEXIT!"
                Break ;
            } ; 
        } ; 
        #if (!$domaincontroller) { $domaincontroller = get-gcfast } ;
        #if(!$domaincontroller){ if(test-path function:get-gcfast){$domaincontroller=get-gcfast} else { throw "no get-gcfast()!" } ;} else {"(existing `$domaincontroller:$($domaincontroller))"} ;
        # use new get-GCFastXO cross-org dc finde
        # default to Op_ExADRoot forest from $TenOrg Meta
        $domaincontroller = get-GCFastXO -TenOrg $TenOrg -subdomain ((gv -name "$($TenOrg)Meta").value['OP_ExADRoot']) -verbose:$($verbose) |?{$_.length};


        # MSOL CONNECTION
        $reqMods += "connect-msol".split(";") ;
        if ( !(check-ReqMods $reqMods) ) { write-error "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Missing function. EXITING." ; Break ; }  ;
        write-verbose "(loading AAD...)" ;
        #connect-msol ;
        connect-msol @pltRXO ; 
        #

        # AZUREAD CONNECTION
        $reqMods += "Connect-AAD".split(";") ;
        if ( !(check-ReqMods $reqMods) ) { write-error "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Missing function. EXITING." ; Break ; }  ;
        write-verbose "(loading AAD...)" ;
        #connect-msol ;
        Connect-AAD @pltRXO ; 
        #


        #
        <# EXO connection
        $pltRXO = @{
            Credential = (Get-Variable -name cred$($tenorg) ).value ;
            verbose = $($verbose) ; } ; 
        if($VerbosePreference = "Continue"){
            $VerbosePrefPrior = $VerbosePreference ;
            $VerbosePreference = "SilentlyContinue" ;
            $verbose = ($VerbosePreference -eq "Continue") ;
        } ; 
        #disconnect-exo ;
        if ($script:useEXOv2) { reconnect-eXO2 @pltRXO }
        else { reconnect-EXO @pltRXO } ;
        # reenable VerbosePreference:Continue, if set, during mod loads 
        if($VerbosePrefPrior -eq "Continue"){
            $VerbosePreference = $VerbosePrefPrior ;
            $verbose = ($VerbosePreference -eq "Continue") ;
        } ;
        #>

        
        # 3:00 PM 9/12/2018 shift this to 1x in the script ; - this would need to be customized per tenant, not used (would normally be for forcing UPNs, but CMW uses brand UPN doms)
        #$script:forestdom = ((get-adforest | select -expand upnsuffixes) | ? { $_ -eq 'toro.com' }) ;

        # Clear error variable
        $Error.Clear() ;
        

    } ;  # BEGIN-E
    PROCESS {
        <#$IdentifierType = "DisplayName" ;
        $IdentifierType = "SmtpAddress" ;
        $IdentifierType = "Alias" ;
        #>
        $pltGetxUser=[ordered]@{
            ErrorAction = 'STOP' ;
        } ;
        switch -regex($IdentifierType){
            '(Alias|SmtpAddress)'{
                $pltGetxUser.add('Identity',$Identifier) ;
            }
            'DisplayName'{
                $fltr = "displayname -like '$Identifier'" ; 
                $pltGetxUser.add('filter',$fltr) ;
            }
            default {
                write-warning "Unable to resolve `$IdentifierType ($($IdentifierType)) into a recognized value" ;
                break ;
            }
        } ;

        write-verbose "$((get-alias ps1GetxUser).definition) w`n$(($pltGetxUser|out-string).trim())" ;         
        #rxo ; cmsol ; caad ; rx10 ;
        $error.clear() ;
        TRY {
            $txUser =ps1GetxUser @pltGetxUser ;
            if($msolu = get-msoluser -user $txUser.UserPrincipalName |?{$_.islicensed}){
            #if($msolu = get-msoluser -user $txUser.UserPrincipalName ){
                $tAADu = get-AzureAdUser -objectID $msolu.UserPrincipalName |?{($_.provisionedplans.service -eq 'exchange')} ;
                if($taadu.extensionproperty.onPremisesDistinguishedName -match $rgxCMWDomain){
                    $bCmwAD=$true ;
                    write-host -fo yellow "ADUser is onprem CMW hybrid!:`n$($taadu.extensionproperty.onPremisesDistinguishedName)" ; 
                } elseif($taadu.DirSyncEnabled -AND $taadu.ImmutableId) {
                    #$tadu = get-aduser -filter {UserPrincipalName -eq $txUser.UserPrincipalName }
                    # no use the converted immutableid
                    $guid=New-Object -TypeName guid (,[System.Convert]::FromBase64String($taadu.ImmutableId)) ;
                    $tadu = get-aduser -Identity $guid.guid ; 
                };
            } else { 
                write-warning "No matching licensed MSolu:(get-msoluser -user $txUser.UserPrincipalName)" ; 
            } ; 
        } CATCH {
            $ErrTrapd=$Error[0] ;
            $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
            write-warning $smsg ;
        } ; 
        switch ($txUser.Recipienttype){
            'UserMailbox'{
                $error.clear() ;
                TRY {$xmbx = ps1GetxMbx -id $txUser.UserPrincipalName -ea stop }
                CATCH {
                    $ErrTrapd=$Error[0] ;
                    $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                    write-warning $smsg ;
                } ; 
            }
            'MailUser'{
                $error.clear() ;
                TRY {$opmbx = get-mailbox -id $txUser.UserPrincipalName -ea stop }
                CATCH {
                    $ErrTrapd=$Error[0] ;
                    $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                    write-warning $smsg ;
                } ; 
            } ;
            default {write-warning "non-mailbox/mailuser object"} 
        } ; 
        if($txUser){
            if($tadu){"=get-ADUser:>`n$(($tadu |fl samaccountn*,userpr*,msRTCSIP-PrimaryU*,msRTCSIP-L*,msRTCSIP-Usere*,tit*|out-string).trim())" 
            } else {
                write-host "=get-ADUser:>(Non-local AD user)`n=$((get-alias ps1GetxUser).definition):`n$(($txUser|fl userpr*,tit*,Offi*,Compa*|out-string).trim())" 
            } ;
            if($xmbx){"=get-Xmbx>:`n$(($xmbx| fl ($exMProps |?{$_ -notmatch '(samaccountname|DistinguishedName)'})|out-string).trim())" } ;
            if($opmbx){"=get-OPmbx>:`n$(($opmbx| fl $exMProps |out-string).trim())" };
            if($msolu){
                write-host "$(($msolu|fl @{Name='HasLic';Expression={$_.IsLicensed }},@{Name='LicIssue';Expression={$_.LicenseReconciliationNeeded }}|out-string).trim())" ; 
            "Licenses Assigned:`n$((($msolu.licenses.AccountSkuId) -join ";" | out-string).trim())" ;
                if(!($bCmwAD)){
                    if($LicGrp = $tadu.memberof -match $rgxLicGrpDN){
                        write-host "LicGrp(AD):$(($LicGrp|out-string).trim())" ; 
                    } else { 
                        write-host "LicGrp(AD):(no ADUser.memberof matched pattern:`n$($rgxLicGrpDN.tostring())" ; 
                    } ; 
                } else {
                    write-host -fo yellow  "Unable to expand ADU, user is hybrid AD from CMW.internal domain`nproxying AzureADUser memberof" ; 
                    if($taadu){
                        $mbrof = $taadu | Get-AzureADUserMembership | select DisplayName,DirSyncEnabled,MailEnabled,SecurityEnabled,Mail,objectid ;
                        if($LicGrp = $mbrof.displayname -match $rgxLicGrpDName){
                            write-host "LicGrp(AAD):$(($LicGrp|out-string).trim())" ; 
                        } else { 
                            write-host "LicGrp(AAD):(no ADUser.memberof matched pattern:`n$($rgxLicGrpDName.tostring())" ; 
                        } ; 
                    } else { 
                        write-warning "(unpopulated AzureADUser: skipping memberof)" ; 
                    }
                } ; 
            }else {
                write-warning "Unable to find matching MsolU for $Identifier" ; 
            } ; 
        } ; 
        
    } ;  # PROC-E
    END {
        # =========== wrap up Tenant connections
        <# suppress VerbosePreference:Continue, if set, during mod loads (VERY NOISEY)
        if($VerbosePreference = "Continue"){
            $VerbosePrefPrior = $VerbosePreference ;
            $VerbosePreference = "SilentlyContinue" ;
            $verbose = ($VerbosePreference -eq "Continue") ;
        } ;
        if($script:useEXOv2){
            disconnect-exo2 -verbose:$($verbose) ;
        } else {
            disconnect-exo -verbose:$($verbose) ;
        } ;
        # aad mod *does* support disconnect (msol doesen't!)
        #Disconnect-AzureAD -verbose:$($verbose) ;
        # reenable VerbosePreference:Continue, if set, during mod loads
        if($VerbosePrefPrior -eq "Continue"){
            $VerbosePreference = $VerbosePrefPrior ;
            $verbose = ($VerbosePreference -eq "Continue") ;
        } ;
        #>
        # clear the script aliases
        write-verbose "clearing ps1* aliases in Script scope" ; 
        get-alias -scope Script |Where-Object{$_.name -match '^ps1.*'} | ForEach-Object{Remove-Alias -alias $_.name} ;

        write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($sBnr.replace('=v','=^').replace('v=','^='))" ;
        
        write-verbose "(explicit EXIT...)" ;
        Break ;


    } ;  # END-E
}

#*------^ resolve-Name.ps1 ^------