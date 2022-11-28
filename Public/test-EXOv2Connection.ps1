#*----------v Function test-EXOv2Connection() v----------
function test-EXOv2Connection {
    <#
    .SYNOPSIS
    test-EXOv2Connection.ps1 - Validate EXO connection, and that the proper Tenant is connected (as per provided Credential)
    .NOTES
    Version     : 1.0.0
    Author      : Todd Kadrie
    Website     :	http://www.toddomation.com
    Twitter     :	@tostka / http://twitter.com/tostka
    CreatedDate : 2022-06-24
    FileName    : test-EXOv2Connection.ps1
    License     : MIT License
    Copyright   : (c) 2022 Todd Kadrie
    Github      : https://github.com/tostka/verb-EXO
    Tags        : Powershell
    REVISIONS
    * 3:54 PM 11/28/2022 move into verb-EXO; copied back from months of debugging in ISE on jb
    * 3:59 PM 8/2/2022 got through dbugging EOM v205 SID interactive pass, working ; fully works from mybox w v206p6, cEOM connection, with functional prefix. need to code in divert on cxo2 etc to avoid redundant tests and just do them here.
    * 3:30 PM 7/25/2022 fixed missing else for if #152; works in tests against CBA & SID interactive creds on EOM v205, need to debug now against EOM v206p6, to accomodate PSSession-less connect & test code.
    * 10:18 AM 6/24/2022 init ; ren test-EXOConnection -> test-EXOv2Connection, as this only validates EXOversion2 connections, not basic-auth-based EXOv1
    .DESCRIPTION
    test-EXOv2Connection.ps1 - Validate EXO connection, and that the proper Tenant is connected (as per provided Credential)
    .PARAMETER Credential
    Credential to be used for connection
    .PARAMETER MinNoWinRMVersion
    MinimumVersion required for Non-WinRM connections (of ExchangeOnlineManagement module (defaults to '2.0.6')[-MinimumVersion '2.0.6']
    .OUTPUT
    System.Boolean
    .EXAMPLE
    PS> $oRet = test-EXOv2Connection -Credential $cred -verbose ; 
    PS> if($oRet.Valid){
    PS>     $pssEXOv2 = $oRet.PsSession ; 
    PS>     write-host 'Validated EXOv2 Connected to Tenant aligned with specified Credential'
    PS> } else { 
    PS>     write-warning 'NO EXO USERMAILBOX TYPE LICENSE!'
    PS> } ; 
    Evaluate EXOv2 connection status & Tenant:Credential alignment, with verbose output
    .LINK
    https://github.com/tostka/verb-EXO
    #>
    #Requires -Version 3
    ##Requires -Modules AzureAD, verb-Text
    ##Requires -RunasAdministrator
    # VALIDATORS: [ValidateNotNull()][ValidateNotNullOrEmpty()][ValidateLength(24,25)][ValidateLength(5)][ValidatePattern("some\sregex\sexpr")][ValidateSet("USEA","GBMK","AUSYD")][ValidateScript({Test-Path $_ -PathType 'Container'})][ValidateScript({Test-Path $_})][ValidateRange(21,65)][ValidateCount(1,3)]
    [CmdletBinding()]
     Param(
        [Parameter(Mandatory=$True,HelpMessage="Credentials [-Credentials [credential object]]")]
        [System.Management.Automation.PSCredential]$Credential = $global:credo365TORSID,
        [Parameter(HelpMessage = "MinimumVersion required for Non-WinRM connections (of ExchangeOnlineManagement module (defaults to '2.0.6')[-MinimumVersion '2.0.6']")]
        [version] $MinNoWinRMVersion = '2.0.6'
    )
    BEGIN {
        ${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name ;
        $Verbose = ($VerbosePreference -eq 'Continue') ;
        if(-not $rgxCertFNameSuffix){$rgxCertFNameSuffix = '-([A-Z]{3})$' } ; 
        #*------v PSS & GMO VARIS v------
        # get-pssession session varis
        # select key differentiating properties:
        $pssprops = 'Id','ComputerName','ComputerType','State','ConfigurationName','Availability', 
            'Description','Guid','Name','Path','PrivateData','RootModuleModule', 
            @{name='runspace.ConnectionInfo.ConnectionUri';Expression={$_.runspace.ConnectionInfo.ConnectionUri} },  
            @{name='runspace.ConnectionInfo.ComputerName';Expression={$_.runspace.ConnectionInfo.ComputerName} },  
            @{name='runspace.ConnectionInfo.Port';Expression={$_.runspace.ConnectionInfo.Port} },  
            @{name='runspace.ConnectionInfo.AppName';Expression={$_.runspace.ConnectionInfo.AppName} },  
            @{name='runspace.ConnectionInfo.Credentialusername';Expression={$_.runspace.ConnectionInfo.Credential.username} },  
            @{name='runspace.ConnectionInfo.AuthenticationMechanism';Expression={$_.runspace.ConnectionInfo.AuthenticationMechanism } },  
            @{name='runspace.ExpiresOn';Expression={$_.runspace.ExpiresOn} } ; 
        $EXOv1ConfigurationName = $EXOv2ConfigurationName = $EXoPConfigurationName = "Microsoft.Exchange" ;

        if(-not $EXOv1ConfigurationName){$EXOv1ConfigurationName = "Microsoft.Exchange" };
        if(-not $EXOv2ConfigurationName){$EXOv2ConfigurationName = "Microsoft.Exchange" };
        if(-not $EXoPConfigurationName){$EXoPConfigurationName = "Microsoft.Exchange" };

        if(-not $EXOv1ComputerName){$EXOv1ComputerName = 'ps.outlook.com' };
        if(-not $EXOv1runspaceConnectionInfoAppName){$EXOv1runspaceConnectionInfoAppName = '/PowerShell-LiveID'  };
        if(-not $EXOv1runspaceConnectionInfoPort){$EXOv1runspaceConnectionInfoPort -eq '443' };

        if(-not $EXOv2ComputerName){$EXOv2ComputerName = 'outlook.office365.com' ;}
        if(-not $EXOv2Name){$EXOv2Name = "ExchangeOnlineInternalSession*" ; }
        if(-not $rgxEXoPrunspaceConnectionInfoAppName){$rgxEXoPrunspaceConnectionInfoAppName = '^/(exadmin|powershell)$'}; 
        if(-not $EXoPrunspaceConnectionInfoPort){$EXoPrunspaceConnectionInfoPort = '80' } ; 
        # gmo varis
        if(-not $rgxEXOv1gmoDescription){$rgxEXOv1gmoDescription = "^Implicit\sremoting\sfor\shttps://ps\.outlook\.com/PowerShell" }; 
        if(-not $EXOv1gmoprivatedataImplicitRemoting){$EXOv1gmoprivatedataImplicitRemoting = $true };
        if(-not $rgxEXOv2gmoDescription){$rgxEXOv2gmoDescription = "^Implicit\sremoting\sfor\shttps://outlook\.office365\.com/PowerShell" }; 
        if(-not $EXOv2gmoprivatedataImplicitRemoting){$EXOv2gmoprivatedataImplicitRemoting = $true } ;
        if(-not $rgxExoPsessionstatemoduleDescription){$rgxExoPsessionstatemoduleDescription = '/(exadmin|powershell)$' };
        if(-not $PSSStateOK){$PSSStateOK = 'Opened' };
        if(-not $PSSAvailabilityOK){$PSSAvailabilityOK = 'Available' };
        if(-not $EXOv1GmoFilter){$EXOv1GmoFilter = 'tmp_*' } ; 
        if(-not $EXOv2GmoNoWinRMFilter){$EXOv2GmoNoWinRMFilter = 'tmpEXO_*' };
        $EOMmodname = 'ExchangeOnlineManagement' ;
        $modname = $EOMmodname ;
        # move into a param
        #$MinNoWinRMVersion = '2.0.6' ;        
        #*------^ END PSS & GMO VARIS ^------

        Try { Get-Module $modname -ErrorAction Stop | out-null } Catch {
            $smsg = "Import-Module w`n$(($pltIMod|out-string).trim())" ;
            if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            Import-Module @pltIMod ;
        } ; # IsImported
        [boolean]$IsNoWinRM = [boolean]([version](get-module $modname).version -ge $MinNoWinRMVersion) ; 

        # 12:18 PM 9/17/2022 prestage cert-handling calcs:
        if($credential.username -match $rgxCertThumbprint){
		    $smsg =  "(UserName:Certificate Thumbprint detected)"
		    if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
		    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            # cert CBA non-basic auth
		    <# CertificateThumbPrint = $Credential.UserName ;
			    AppID = $Credential.GetNetworkCredential().Password ;
			    Organization = 'TENANT.onmicrosoft.com' ; # org is on $xxxmeta.o365_TenantDomain
		    #>
            # want the friendlyname to display the cred source in use #$tcert.friendlyname
		    if($tcert = get-childitem -path "Cert:\CurrentUser\My\$($credential.username)"){
			    $certUname = $tcert.friendlyname ; 
			    $certTag = [regex]::match($certUname,$rgxCertFNameSuffix).captures[0].groups[1].value ; 
			    $smsg = "(calc'd CBA values:cred:$($certTag):$([string]$tcert.friendlyname))" ; 
			    if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
			    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            } else { 
			    $smsg = "calc'd CBA values FAIL!: UNABLE TO locate cert matching Cert:\CurrentUser\My\`$credential.username" ;
			    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } 
			    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
			    throw $smsg ; 
			    Break ; 
		    } ;
        } ;
    } ;  # if-E BEGIN    
    PROCESS {
        $oReturn = [ordered]@{
            PSSession = $null ; 
            IsNoWinRM = $false ; 
            Valid = $false ; 
        } ; 
        $isEXOValid = $false ;
        # corrected $EXoPConfigurationName -> $EXOv2ConfigurationName; same value, but vari name should indicate purpose, as well as contents
        if($pssEXOv2 = Get-PSSession | 
                where-object {$_.ConfigurationName -like $EXOv2ConfigurationName -AND (
                    $_.Name -like $EXOv2Name) -AND (
                    $_.ComputerName -eq $EXOv2ComputerName) } ){
                    <# rem'd state/avail tests, run separately below: -AND (
                    $_.State -eq $PSSStateOK)  -AND (
                    $_.Availability -eq $PSSAvailabilityOK)
                    #>
            $smsg = "`n`nEXOv2 PSSessions:`n$(($pssEXOv2 | fl $pssprops|out-string).trim())" ; 
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 

            if($pssEXOv2 | ?{ ($_.State -eq $PSSStateOK)  -AND (
                    $_.Availability -eq $PSSAvailabilityOK)}){

                # verify the exov2 cmdlets actually imported as a tmp_ module w specifid prefix & 1st cmdlet
                # below won't work with updated token support/MFA & loss of test|clear-ActiveToken from EOM (breaking change)
                # but it's needed when using EOM205, which still falls to basicauth! (readded down below)
                #if ( (get-module -name $EXOv1GmoFilter | ForEach-Object { Get-Command -module $_.name -name 'Add-xoAvailabilityAddressSpace' -ea 0 }) -AND (test-EXOToken) ) {

                # 2:47 PM 7/13/2022 revise for exov2 -cred support (where get-msaltoken gets used)
                # test-EXOToken & it's dependancy EOM:test-ActiveToken, *doesn't exist* after EOM v205!, if out the block
                if(-not $IsNoWinRM){
                    # Credential
                    $plttXT=[ordered]@{
                        Credential = $Credential ;
                        verbose = $($VerbosePreference -eq "Continue") ;
                    } ;
                    $smsg = "test-EXOToken w`n$(($plttXT|out-string).trim())" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                } ; 
                if ( (get-module -name $EXOv1GmoFilter | ForEach-Object {
                     Get-Command -module $_.name -name 'Add-xoAvailabilityAddressSpace' -ea 0 }) -AND (
                        test-EXOToken @plttXT)) {
                    $smsg = "(EXOv1Gmo Basic-Auth PSSession module detected)" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    $bExistingEXOGood = $true ;
                    # need as well, to get through if/then ahead of credential/tenant align check
                    $isEXOValid = $true ; 
                }elseif ( (get-module -name $EXOv2GmoNoWinRMFilter | ForEach-Object {
                    Get-Command -module $_.name -name 'Add-xoAvailabilityAddressSpace' -ea 0 }) -AND (
                        test-EXOToken @plttXT)) {
                    $smsg = "(EXOv2GmoNoWinRM module detected)" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    $bExistingEXOGood = $true ;
                    # need as well, to get through if/then ahead of credential/tenant align check
                    $isEXOValid = $true ; 
                } else { $bExistingEXOGood = $false ; }
                
                
            } else{
                # pss but disconnected state
                rxo2 ; 
            } ; 
            
        } elseif($IsNoWinRM -AND ((get-module $EXOv2GmoNoWinRMFilter) -AND (get-module $modName))){
            # no PSS and IsNoWinRM == v206+ PSS-less connection
            # verify the exov2 cmdlets actually imported as a tmp_ module w specifid prefix & 1st cmdlet

            # test-EXOToken() won't work with PSSession-less - it obtains the critical TokenExpireTime from the open PSSession
            # need to recode for these using get-MsalToken:

            # 12:22 PM 8/1/2022 issue with get-msaltoken: it will auth EXO client app (by guid), but it doesn't support the key -prefix param, to make them verb-XOnoun; so you can't use it with hybrid onprem connections.
            # => looks like I'll have to either skip it, or test for cmdlets loaded, to verify. get-msaltoken actually runs an auth session, doesn't just validate one's present. 
            <# [PowerShell Gallery | MSAL.PS.psd1 4.1.0.2 - www.powershellgallery.com/](https://www.powershellgallery.com/packages/MSAL.PS/4.1.0.2/Content/MSAL.PS.psd1)
             nope, it's referring to 'virtual network address prefix'f
            #>

            # it seamlessly reauths, wo prompts, so just validate a core cmdlet is loaded, plust the above
            if([boolean](get-command -name Get-xoOrganizationConfig)){
                $smsg = "(`IsNoWinRM:`$true`nget-module:$($EXOv2GmoNoWinRMFilter)`nget-module:$($modName)`ngcm:Get-xoOrganizationConfig`n=>Appears Valid)" ; 
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 

                $bExistingEXOGood = $isEXOValid = $true ;
                #$IsNoWinRM = $true ; # already tested above
            } else { 
                $bExistingEXOGood = $isEXOValid = $false ;
            } ; 

            # v dead pss-based or msal.net-based code, that can't do prefixes properly. v
            <#
            # splat to run get-MSALToken
            $pltGMT = @{
                TenantId = $null ;
                #ClientId = $credential.GetNetworkCredential().Password ;
                #ClientCertificate = Get-Item "Cert:\CurrentUser\My\$($credential.UserName)" ;
                ErrorAction = 'Stop' ;
                Verbose = $($VerbosePreference -eq "Continue") ;
            } ;

            if($credential.username -match $rgxCertThumbprint){
                $smsg = "(CBA cert auth creds detected)" ; 
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 

                $pltGMT.add('ClientId',$credential.GetNetworkCredential().Password) ;
                $pltGMT.add('ClientCertificate',(Get-Item "Cert:\CurrentUser\My\$($credential.UserName)" -ErrorAction 'STOP') ) ;

            } else { 
                $smsg = "(NON-CBA auth creds detected - trying interactive)" ; 
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                # 1:01 PM 7/8/2022 left off coding here, need interactive options for creds etc

                # try EXOPS clientid
                $pltGMT.add('ClientId','a0c73c16-a7e3-4564-9a95-2bdf47383716') ;
                # this uses the EXOPS clientid with a UPN-based credential
                # $authContext = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext" -ArgumentList https://login.microsoftonline.com/tenantname.onmicrosoft.com/
                # $client_id = "a0c73c16-a7e3-4564-9a95-2bdf47383716" # EXORemPS
                # $Credential = Get-Credential user@tenantname.onmicrosoft.com
                # $AADcredential = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.UserPasswordCredential" -ArgumentList $Credential.UserName,$Credential.Password
                # $authResult = [Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContextIntegratedAuthExtensions]::AcquireTokenAsync($authContext,"https://outlook.office365.com",$client_Id,$AADcredential)
                #
                # force interactive
                $pltGMT.add('Interactive',$true) ; 


            } ; 

            # # leveraging the passed in $Credential (or global in sample below)
            # $tenOrg = 'TOR' ; 
            # ipmo msal.ps -force ; 
            # $pltGMT = @{
            #     TenantId = $null ;
            #     ClientId = $credXXXCBA.GetNetworkCredential().Password ;
            #     ClientCertificate = Get-Item "Cert:\CurrentUser\My\$($credXXXCBA.UserName)" ;
            #     ErrorAction = 'Stop' ;
            # } ;
            # if($TenID= (Get-Variable  -name "$($TenOrg)Meta").value.o365_TenantID ){
            # 	$pltGMT.TenantId = $TenID;
            # } else {
            # 	$smsg = "UNABLE TO RESOLVE `$TENORG:$($TenOrg) TO FUNCTIONAL `$$($TenOrg)meta.o365_TenantID!" ;
            # 	write-WARNING $smsg ;
            # } ;
            # $smsg = "Get-msaltoken w`n$(($pltGMT|out-string).trim())" ;
            # write-host $smsg ;
            # $msalToken = Get-msaltoken @pltGMT -ForceRefresh ; 
            # 
            # #-=-=-=-=-=-=-=-=
            # $msaltoken | fl *
            # AccessToken                  : eyJ0eXAiOi...p6Dx5z9dg
            # IsExtendedLifeTimeToken      : False
            # UniqueId                     :
            # ExpiresOn                    : 7/8/2022 6:15:01 PM +00:00
            # ExtendedExpiresOn            : 7/8/2022 6:15:01 PM +00:00
            # TenantId                     :
            # Account                      :
            # IdToken                      :
            # Scopes                       : {https://graph.microsoft.com/.default}
            # CorrelationId                : 4dec58a7-XXXX-XXXX-b78e-907ee0bc79ee
            # TokenType                    : Bearer
            # ClaimsPrincipal              :
            # AuthenticationResultMetadata : Microsoft.Identity.Client.AuthenticationResultMetadata
            # User  
            # #-=-=-=-=-=-=-=-=
            #             
            #

            #if($TenID= (Get-Variable  -name "$($TenOrg)Meta").value.o365_TenantID ){
	        #    $pltGMT.TenantId = $TenID;
            # use domain, better readable for echos
            if($TenDom= (Get-Variable  -name "$($TenOrg)Meta").value.o365_TenantDomain ){
                $pltGMT.TenantId = $TenDom;
            } else {
	            #$smsg = "UNABLE TO RESOLVE `$TENORG:$($TenOrg) TO FUNCTIONAL `$$($TenOrg)meta.o365_TenantID!" ;
                $smsg = "UNABLE TO RESOLVE `$TENORG:$($TenOrg) TO FUNCTIONAL `$$($TenOrg)meta.o365_TenantDomain!" ;
	            write-WARNING $smsg ;
                throw $smsg ; 
                Break ; 
            } ;

            $smsg = "Get-msaltoken w`n$(($pltGMT|out-string).trim())" ;
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 

            $error.clear() ;
            Try {
                $MsalResponse = Get-MsalToken @pltGMT ; 
                # 3:07 PM 7/13/2022 need to find a way to validate the token status!
                if ($MsalResponse.AccessToken) {
                    # ADD $isEXOValid TOO (needed to get through accdom code)
                    $bExistingEXOGood = $isEXOValid = $true ;
                    $IsNoWinRM = $true ;
                } else { $bExistingEXOGood = $false ; }


            } CATCH {
                $ErrTrapd=$Error[0] ;
                $smsg = "$('*'*5)`nFailed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: `n$(($ErrTrapd|out-string).trim())`n$('-'*5)" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug 
                else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                #-=-record a STATUSWARN=-=-=-=-=-=-=
                $statusdelta = ";WARN"; # CHANGE|INCOMPLETE|ERROR|WARN|FAIL ;
                if(gv passstatus -scope Script -ea 0){$script:PassStatus += $statusdelta } ;
                if(gv -Name PassStatus_$($tenorg) -scope Script -ea 0){set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta)} ; 
                #-=-=-=-=-=-=-=-=
                $smsg = "FULL ERROR TRAPPED (EXPLICIT CATCH BLOCK WOULD LOOK LIKE): } catch[$($ErrTrapd.Exception.GetType().FullName)]{" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                Break #Opts: STOP(debug)|EXIT(close)|CONTINUE(move on in loop cycle)|BREAK(exit loop iteration)|THROW $_/'CustomMsg'(end script with Err output)
            } ; 

            #> # ^ dead pss-based or msal.net-based code, that can't do prefixes properly. ^

            

        } else { 
            $smsg = "Unable to detect EXOv2 PSSession!" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } 
            else{ write-host -ForegroundColor yellow "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
            #throw $smsg ;
            #Break ; 
            $bExistingEXOGood = $isEXOValid = $false ; 
        } ; 

        if($bExistingEXOGood -ANd $isEXOValid){
            # implement caching of accepteddoms into the XXXMeta, in the session (cut back on queries to EXO on acceptedom)
            # swap in non-looping
            if( get-command Get-xoAcceptedDomain) {
                    #$TenOrg = get-TenantTag -Credential $Credential ;
                if(-not (Get-Variable  -name "$($TenOrg)Meta").value.o365_AcceptedDomains){
                    set-Variable  -name "$($TenOrg)Meta" -value ( (Get-Variable  -name "$($TenOrg)Meta").value += @{'o365_AcceptedDomains' = (Get-xoAcceptedDomain).domainname} )
                } ;
            } ;
            
            $smsg = "(validating Tenant:Credential alignment)" ; 
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            
            # spliced in block to precalc $certtag & $certUname up in begin{} (if $credential.username matches a certy thumbprint)
            if( ($credential.username -match $rgxCertThumbprint) -AND ((Get-Variable  -name "$($TenOrg)Meta").value.o365_Prefix -eq $certTag )){
                # 9:59 AM 6/24/2022 need a case for CBA cert (thumbprint username)
                # compare cert fname suffix to $xxxMeta.o365_Prefix
                # validate that the connected EXO is to the CBA Cert tenant
                $smsg = "(EXO Authenticated & Functional CBA cert:$($certTag),($($certUname)))" ;
                if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                $bExistingEXOGood = $isEXOValid = $true ;
            }elseif((Get-Variable  -name "$($TenOrg)Meta").value.o365_AcceptedDomains.contains($Credential.username.split('@')[1].tostring())){
                # validate that the connected EXO is to the $Credential tenant
                $smsg = "(EXO Authenticated & Functional:$($Credential.username.split('@')[1].tostring())),($($Credential.username))" ;
                if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                $bExistingEXOGood = $isEXOValid = $true ;
            # issue: found fresh bug in cxo: svcacct UPN suffix @tenantname.onmicrosoft.com, but testing against AccepteDomain, it's not in there (tho @DOMAIN.mail.onmicrosoft.comis)
            }elseif((Get-Variable  -name "$($TenOrg)Meta").value.o365_TenantDomain -eq ($Credential.username.split('@')[1].tostring())){
                $smsg = "(EXO Authenticated & Functional(TenDom):$($Credential.username.split('@')[1].tostring()))" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                $bExistingEXOGood = $isEXOValid = $true ;
            } else {
                $smsg = "(Credential mismatch:disconnecting from existing EXO:$($eEXO.Identity) tenant)" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } 
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                Disconnect-exo ;
                $bExistingEXOGood = $isEXOValid = $false ;
            } ;

            if($bExistingEXOGood -AND $isEXOValid){
                $oReturn.PSSession = $pssEXOv2 ; 
                if( ($IsNoWinRM = $false) -AND -not $pssEXOv2){
                    $smsg = "IsNoWinRM & no detected EXO PsSession:" ; 
                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                    $oReturn.IsNoWinRM = $IsNoWinRM
                }
                $oReturn.Valid = $isEXOValid ; 
            } else {
                $smsg = "(invalid session `$bExistingEXOGood:$($bExistingEXOGood) -OR `$isEXOValid:$($isEXOValid))" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } 
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                Disconnect-exo ;
            } ;  ; 
        } else { 
            $oReturn.PSSession = $pssEXOv2 ; 
            $oReturn.Valid = $isEXOValid ; 
        } ; 

    }  # PROC-E
    END{
        <# $oReturn = [ordered]@{
            PSSession = $null ; 
            Valid = $false ; 
        } ; 
        #>
        $smsg = "Returning `$oReturn:`n$(($oReturn|out-string).trim())" ; 
        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        New-Object PSObject -Property $oReturn | write-output ; 
    } ;
} ; 
#*------^ END Function test-EXOv2Connection() ^------