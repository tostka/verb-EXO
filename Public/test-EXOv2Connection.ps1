# test-EXOv2Connection.ps1

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
    * 1:55 PM 3/1/2024 added code to repop empty $TenOrg, prior to AcceptedDom caching (came through empty in testing, when no preexisting conn)
    * 2:51 PM 2/26/2024 add | sort version | select -last 1  on gmos, LF installed 3.4.0 parallel to 3.1.0 and broke auth: caused mult versions to come back and conflict with the assignement of [version] type (would require [version[]] to accom both, and then you get to code everything for mult handling)
    * 3:26 PM 5/30/2023 updated CBH, demos ; # reduced the ipmo and vers chk block, removed the lengthy gmo -list; and any autoinstall. Assume EOM is installed, & break if it's not
    * 11:20 AM 4/25/2023 added -CertTag param (passed by connect-exo; used for validating credential alignment w Tenant)
    * 10:28 AM 4/18/2023 #372: added -ea 0 to gv calls (not found error suppress)
    * 2:02 PM 4/17/2023 rev: $MinNoWinRMVersion from 2.0.6 => 3.0.0.
    * 3:58 PM 4/4/2023 reduced the ipmo and vers chk block, removed the lengthy gmo -list; and any autoinstall. Assume EOM is installed, & break if it's not; 
    fixed flipped $IsNoWinRM ; supports EMOv2 v EMOv3 pss/no-pss connections, adds support for get-connectioninformation()
    * 3:14 pm 3/29/2023: REN'D $modname => $EOMModName
    * 3:54 PM 11/29/2022:  force the $MinNoWinRMVersion value to the currnet highest loaded:; 
    * 3:59 PM 8/2/2022 got through dbugging EOM v205 SID interactive pass, working ; fully works from mybox w v206p6, cEOM connection, with functional prefix. need to code in divert on cxo2 etc to avoid redundant tests and just do them here.
    * 3:30 PM 7/25/2022 fixed missing else for if #152; works in tests against CBA & SID interactive creds on EOM v205, need to debug now against EOM v206p6, to accomodate PSSession-less connect & test code.
    * 10:18 AM 6/24/2022 init ; ren test-EXOConnection -> test-EXOv2Connection, as this only validates EXOversion2 connections, not basic-auth-based EXOv1
    .DESCRIPTION
    test-EXOv2Connection.ps1 - Validate EXO connection, and that the proper Tenant is connected (as per provided Credential)
    .PARAMETER Credential
    Credential to be used for connection
    .PARAMETER CertTag
    Cert FriendlyName Suffix to be used for validating credential alignment(Optional but required for CBA calls)[-CertTag `$certtag]
    .PARAMETER MinNoWinRMVersion
    MinimumVersion required for Non-WinRM connections (of ExchangeOnlineManagement module (defaults to '3.0.0')[-MinimumVersion '2.0.6']
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
    .EXAMPLE
    PS> $TenOrg = get-TenantTag -Credential $Credential ;
    PS> if($Credential){
    PS>     $uRoleReturn = resolve-UserNameToUserRole -Credential $Credential
    PS> } elseif($UserPrincipalName){
    PS>     $uRoleReturn = resolve-UserNameToUserRole -UserName $UserPrincipalName
    PS> } ; 
    PS> if($uRoleReturn.TenOrg){
    PS>     $CertTag = $uRoleReturn.TenOrg
    PS> } ; 
    PS> if($CertTag -ne $null){
    PS>     $smsg = "(specifying detected `$CertTag:$($CertTag))" ;
    PS>     if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
    PS>     else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
    PS>     $oRet = test-EXOv2Connection -Credential $credential -CertTag $CertTag -verbose:$($verbose) ;
    PS> } else {
    PS>     $oRet = test-EXOv2Connection -Credential $credential -verbose:$($verbose) ;
    PS> } ;
    PS> if($oRet.Valid){
    PS>     $pssEXOv2 = $oRet.PsSession ;
    PS>     $IsNoWinRM = $oRet.IsNoWinRM ;
    PS>     $smsg = "(Validated EXOv2 Connected to Tenant aligned with specified Credential)" ;
    PS>     if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
    PS>     else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
    PS> } else {
    PS>     $smsg = "NO VALID EXOV2/3 PSSESSION FOUND! (DISCONNECTING...)"
    PS>     if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN }
    PS>     else{ write-host -ForegroundColor yellow "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
    PS>     # capture outlier: shows a session wo the test cmdlet, force reset
    PS>     DisConnect-EXO ;
    PS>     $bExistingEXOGood = $false ;
    PS> } ;    
    Fancier demo using a variety of verb-Auth & verb-xo cmdlets
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
        [Parameter(HelpMessage = "Cert FriendlyName Suffix to be used for validating credential alignment(Optional but required for CBA calls)[-CertTag `$certtag]")]
        [string]$CertTag,
        [Parameter(HelpMessage = "MinimumVersion required for Non-WinRM connections (of ExchangeOnlineManagement module (defaults to '3.0.0')[-MinimumVersion '2.0.6']")]
        [version] $MinNoWinRMVersion = '3.0.0'
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
        if(-not $EXOv1runspaceConnectionInfoPort){$EXOv1runspaceConnectionInfoPort = '443' };

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
        # reduced the ipmo and vers chk block, removed the lengthy gmo -list; and any autoinstall. Assume EOM is installed, & break if it's not
        #region EOMREV ; #*------v EOMREV Check v------
        #$EOMmodname = 'ExchangeOnlineManagement' ;
        $pltIMod = @{Name = $EOMmodname ; ErrorAction = 'Stop' ; verbose=$false} ;
        if($xmod = Get-Module $EOMmodname -ErrorAction Stop| sort version | select -last 1 ){ } else {
            $smsg = "Import-Module w`n$(($pltIMod|out-string).trim())" ;
            if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
            Try {
                Import-Module @pltIMod | out-null ;
                $xmod = Get-Module $EOMmodname -ErrorAction Stop | sort version | select -last 1 ;
            } Catch {
                $ErrTrapd=$Error[0] ;
                $smsg = "$('*'*5)`nFailed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: `n$(($ErrTrapd|out-string).trim())`n$('-'*5)" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                $smsg = $ErrTrapd.Exception.Message ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                Break ;
            } ;
        } ; # IsImported
        if([version]$xmod.version -ge $MinNoWinRMVersion){
            $MinNoWinRMVersion = $xmod.version.tostring() ;
            $IsNoWinRM = $true ; 
        }
        [boolean]$UseConnEXO = [boolean]([version]$xmod.version -ge $MinNoWinRMVersion) ; 
        #endregion EOMREV ; #*------^ END EOMREV Check  ^------

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
                # revise for exov2 -cred support (where get-msaltoken gets used)
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
            
        } elseif($IsNoWinRM -AND ((get-module $EXOv2GmoNoWinRMFilter) -AND (get-module $EOMmodname))){
            # no PSS and IsNoWinRM == v206+ PSS-less connection
            # verify the exov2 cmdlets actually imported as a tmp_ module w specifid prefix & 1st cmdlet

            # test-EXOToken() won't work with PSSession-less - it obtains the critical TokenExpireTime from the open PSSession
            # need to recode for these using get-aadtoken

            # 12:22 PM 8/1/2022 issue with get-msaltoken: it will auth EXO client app (by guid), but it doesn't support the key -prefix param, to make them verb-XOnoun; so you can't use it with hybrid onprem connections.
            # => looks like I'll have to either skip it, or test for cmdlets loaded, to verify. get-msaltoken actually runs an auth session, doesn't just validate one's present. 
            <# [PowerShell Gallery | MSAL.PS.psd1 4.1.0.2 - www.powershellgallery.com/](https://www.powershellgallery.com/packages/MSAL.PS/4.1.0.2/Content/MSAL.PS.psd1)
             nope, it's referring to 'virtual network address prefix'f
            #>
            # EOM v3 adds Get-ConnectionInformation, which has .tokenStatus -eq 'Active'

            if($xmod | Where-Object {$_.version -like "3.*"} ){
                $smsg = "EOM v3+ connection detected" ; 
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                if ((Get-ConnectionInformation).tokenStatus -eq 'Active') {
                    #write-host 'Connecting to Exchange Online' -ForegroundColor Cyan
                    #Connect-ExchangeOnline -UserPrincipalName $adminUPN
                    $bExistingEXOGood = $isEXOValid = $true ;
                }
            } else {  
                $smsg = "EOM v205p6+ connection detected" ; 
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                # it seamlessly reauths, wo prompts, so just validate a core cmdlet is loaded, plust the above
                if([boolean](get-command -name Get-xoOrganizationConfig)){
                    $smsg = "(`IsNoWinRM:`$true`nget-module:$($EXOv2GmoNoWinRMFilter)`nget-module:$($EOMmodname)`ngcm:Get-xoOrganizationConfig`n=>Appears Valid)" ; 
                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 

                    $bExistingEXOGood = $isEXOValid = $true ;
                    #$IsNoWinRM = $true ; # already tested above
                } else { 
                    $bExistingEXOGood = $isEXOValid = $false ;
                } ; 

            } ;
        } else { 
            $smsg = "Unable to detect EXOv2 or EXOv3 PSSession!" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } 
            else{ write-host -ForegroundColor yellow "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
            #throw $smsg ;
            #Break ; 
            $bExistingEXOGood = $isEXOValid = $false ; 
        } ; 

        if($bExistingEXOGood -ANd $isEXOValid){
            # implement caching of accepteddoms into the XXXMeta, in the session (cut back on queries to EXO on acceptedom)
            if( get-command Get-xoAcceptedDomain) {
                    #$TenOrg = get-TenantTag -Credential $Credential ;
                    # 1:10 PM 3/1/2024 Tenorg coming through unpopulated (after pulling legacy code), conditionally re-use the above rem:
                    if( (-not $TenOrg) -AND $Credential){ $TenOrg = get-TenantTag -Credential $Credential } ; 
                if(-not (Get-Variable  -name "$($TenOrg)Meta" -ea 0).value.o365_AcceptedDomains){
                    set-Variable  -name "$($TenOrg)Meta" -value ( (Get-Variable  -name "$($TenOrg)Meta" -ea 0).value += @{'o365_AcceptedDomains' = (Get-xoAcceptedDomain).domainname} )
                } ;
            } ;
            
            $smsg = "(validating Tenant:Credential alignment)" ; 
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            
            if($credential.username -match $rgxCertThumbprint -AND $certTag -eq $null){
                $smsg = "CBA Certificate Thumprint cred uname detected, but -CertTag was *not* pass thru in call!" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                break ; 
            } 
            if( ($credential.username -match $rgxCertThumbprint) -AND ((Get-Variable  -name "$($TenOrg)Meta" -ea 0).value.o365_Prefix -eq $certTag )){
                # 9:59 AM 6/24/2022 need a case for CBA cert (thumbprint username)
                # compare cert fname suffix to $xxxMeta.o365_Prefix
                # validate that the connected EXO is to the CBA Cert tenant
                $smsg = "(EXO Authenticated & Functional CBA cert:$($certTag),($($certUname)))" ;
                if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                $bExistingEXOGood = $isEXOValid = $true ;
            }elseif((Get-Variable  -name "$($TenOrg)Meta" -ea 0).value.o365_AcceptedDomains.contains($Credential.username.split('@')[1].tostring())){
                # validate that the connected EXO is to the $Credential tenant
                $smsg = "(EXO Authenticated & Functional:$($Credential.username.split('@')[1].tostring())),($($Credential.username))" ;
                if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                $bExistingEXOGood = $isEXOValid = $true ;
            # issue: found fresh bug in cxo: svcacct UPN suffix @tenantname.onmicrosoft.com, but testing against AccepteDomain, it's not in there (tho @DOMAIN.mail.onmicrosoft.comis)
            }elseif((Get-Variable  -name "$($TenOrg)Meta" -ea 0).value.o365_TenantDomain -eq ($Credential.username.split('@')[1].tostring())){
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
                if( ($IsNoWinRM -eq $true) -AND -not $pssEXOv2){
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