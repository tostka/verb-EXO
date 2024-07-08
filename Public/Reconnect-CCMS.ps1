# Reconnect-CCMS.ps1

#*------v Reconnect-CCMS.ps1 v------
Function Reconnect-CCMS {
   <#
    .SYNOPSIS
    Reconnect-CCMS - Reestablish connection to Exchange Online (via EXO V2/V3 graph-api module)
    .NOTES
    Version     : 1.0.0
    Author      : Todd Kadrie
    Website     :	http://www.toddomation.com
    Twitter     :	@tostka / http://twitter.com/tostka
    CreatedDate : 2020-07-29
    FileName    :
    License     :
    Copyright   :
    Github      : https://github.com/tostka/verb-exo
    Tags        : Powershell,ExchangeOnline,Exchange,RemotePowershell,Connection,MFA
    REVISIONS   :
    # 4:47 PM 7/8/2024 this is obsoleted; shifted all (re|dis)connect-CCMS functions into connect-exo & reconnect-exo: CCMS Sec & Compl connection mgmt is triggered via the -Prefix cc parameter (any other param is assumed to be native EXO; but -Prefix cc will always generate a connection to Sec & Compliance); 
    * 8:37 AM 4/2/2024 adapt updated reconnect-exo to ccms
    * 2:51 PM 2/26/2024 add | sort version | select -last 1  on gmos, LF installed 3.4.0 parallel to 3.1.0 and broke auth: caused mult versions to come back and conflict with the assignement of [version] type (would require [version[]] to accom both, and then you get to code everything for mult handling)
    * 11:20 AM 9/16/2021 string clean
    * 12:16 PM 5/27/2020 updated cbh, moved alias:rccms win func
    * 4:20 PM 5/14/2020 trimmed redundant func defs from bottom
    * 2:53 PM 5/14/2020 added test & local spec for $rgxCCMSPsHostName, wo it, it can't detect disconnects
    * 1:07 PM 11/25/2019 added *tol/*tor/*cmw alias variants for connect & reconnect
    # 2:42 PM 11/19/2019 started roughing in mfa support
    # 1:24 PM 11/7/2018 switch the test to $EOLSession.state -ne 'Opened' -AND $EOLSession.Availability -ne 'Available'
    # 1:04 PM 6/20/2018 CCMS variant, works
    .DESCRIPTION
    Reconnect-CCMS - Reestablish connection to Exchange Online (via EXO V2/V3 graph-api module)
    .PARAMETER Credential
    Credential to use for this connection [-credential [credential obj variable]
     .PARAMETER MinimumVersion
    MinimumVersion required for ExchangeOnlineManagement module (defaults to '2.0.5')[-MinimumVersion '2.0.6']
    .PARAMETER MinNoWinRMVersion
    MinimumVersion required for Non-WinRM connections (of ExchangeOnlineManagement module (defaults to '3.0.0')[-MinimumVersion '2.0.6']
    .PARAMETER ProxyEnabled
    Use Proxy-Aware SessionOption settings [-ProxyEnabled]
    .PARAMETER UserRole
    Credential Optional User Role spec for credential discovery (wo -Credential)(SID|CSID|UID|B2BI|CSVC|ESVC|LSVC|ESvcCBA|CSvcCBA|SIDCBA)[-UserRole @('SIDCBA','SID','CSVC')]
    .PARAMETER TenOrg
    Optional Tenant Tag (wo -Credential)[-TenOrg 'XYZ']
    .PARAMETER showDebug
    Debugging Flag [-showDebug]
    .PARAMETER silent
    Switch to specify suppression of all but warn/error echos.
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    None. Returns no objects or output.
    .EXAMPLE
    PS>  Reconnect-CCMS;
    Reconnect EXO connection self-locating creds
    .EXAMPLE
    PS>  Reconnect-CCMS -credential $cred ;
    Reconnect EXO connection with explicit [pscredential] object credential specified
    .EXAMPLE
    Reconnect-CCMS -UserRole SIDCBA -TenOrg ABC -verbose  ; 
    Demo use of UserRole (specifying a CBA variant), AND TenOrg spec, to connect (autoresolves against preconfigured credentials in profile)
    .EXAMPLE
    PS> $pltRXOC = [ordered]@{
    PS>     Credential = $Credential ;
    PS>     verbose = $($VerbosePreference -eq "Continue")  ;
    PS>     Silent = $silent ; 
    PS> } ;
    PS> if ($script:useEXOv2 -OR $useEXOv2) { Reconnect-CCMS2 @pltRXOC }
    PS> else { Reconnect-CCMS @pltRXOC } ;    
    Splatted example leveraging prefab $pltRXOC splat, derived from local variables & $VerbosePreference value.
    .EXAMPLE
    PS>  $batchsize = 10 ;
    PS>  $RecordCount=$mr.count #this is the array of whatever you are processing ;
    PS>  $b=0 #this is the initialization of a variable used in the do until loop below ;
    PS>  $mrs = @() ;
    PS>  do {
    PS>      Write-Progress -Activity "Getting move request statistics for all $wave move requests." -Status "Processing Records $b through $($b+$batchsize) of $RecordCount." -PercentComplete ($b/$RecordCount*100) ;
    PS>      $mrs += $mr | Select-Object -skip $b -first $batchsize | foreach-object {Reconnect-CCMS; $_ | Get-OLMoveRequestStatistics} ;
    PS>      $b=$b+$batchsize ;
    PS>      } ;
    PS>  until ($b -gt $RecordCount) ;
    Demo use of a Reconnect-CCMS call mid-looping process to ensure connection stays available for duration of long-running process.    
    .LINK
    Github      : https://github.com/tostka/verb-exo
    #>
    [CmdletBinding()]
    [Alias('rccms')]
    PARAM(
        [Parameter(HelpMessage = "MinimumVersion required for ExchangeOnlineManagement module (defaults to '2.0.5')[-MinimumVersion '2.0.6']")]
            [version] $MinimumVersion = '2.0.5',
        [Parameter(HelpMessage = "MinimumVersion required for Non-WinRM connections (of ExchangeOnlineManagement module (defaults to '3.0.0')[-MinimumVersion '2.0.6']")]
            [version] $MinNoWinRMVersion = '3.0.0',
        [Parameter(HelpMessage="Use Proxy-Aware SessionOption settings [-ProxyEnabled]")]
            [boolean]$ProxyEnabled = $False,
        [Parameter(HelpMessage="Credential to use for this connection [-credential [credential obj variable]")]
            [System.Management.Automation.PSCredential]$Credential,
            # = $global:credo365TORSID, # defer to TenOrg & UserRole resolution
        [Parameter(Mandatory = $false, HelpMessage = "Credential User Role spec (SID|CSID|UID|B2BI|CSVC|ESVC|LSVC|ESvcCBA|CSvcCBA|SIDCBA)[-UserRole @('SIDCBA','SID','CSVC')]")]
            # sourced from get-admincred():#182: $targetRoles = 'SID', 'CSID', 'ESVC','CSVC','UID','ESvcCBA','CSvcCBA','SIDCBA' ; 
            #[ValidateSet("SID","CSID","UID","B2BI","CSVC","ESVC","LSVC","ESvcCBA","CSvcCBA","SIDCBA")]
            # pulling the pattern from global vari w friendly err
            [ValidateScript({
                if(-not $rgxPermittedUserRoles){$rgxPermittedUserRoles = '(SID|CSID|UID|B2BI|CSVC|ESVC|LSVC|ESvcCBA|CSvcCBA|SIDCBA)'} ;
                if(-not ($_ -match $rgxPermittedUserRoles)){throw "'$($_)' doesn't match `$rgxPermittedUserRoles:`n$($rgxPermittedUserRoles.tostring())" ; } ; 
                return $true ; 
            })]
            # cba's don't have perms to s&c: shift to mfa'd sid only
            [string[]]$UserRole = @('SID'),
            # @('SIDCBA','SID','CSVC'),
        [Parameter(Mandatory=$FALSE,HelpMessage="TenantTag value, indicating Tenants to connect to[-TenOrg 'TOL']")]
            [ValidateNotNullOrEmpty()]
            #[ValidatePattern("^\w{3}$")]
            [string]$TenOrg = $global:o365_TenOrgDefault,
        [Parameter(HelpMessage="Debugging Flag [-showDebug]")]
            [switch] $showDebug,
        [Parameter(HelpMessage="Silent output (suppress status echos)[-silent]")]
            [switch]$silent
    ) ;
    BEGIN{
        write-warning "OBSOLETE! shifted all (re|dis)connect-CCMS functions into connect-exo & reconnect-exo: CCMS Sec & Compl connection mgmt is triggered via the -Prefix cc parameter (any other param is assumed to be native EXO; but -Prefix cc will always generate a connection to Sec & Compliance)!" ; 
        BREAK ; 
        $verbose = ($VerbosePreference -eq "Continue") ; 
        if(-not (get-variable rgxExoPsHostName -ea 0)){$rgxExoPsHostName="^(ps\.outlook\.com|outlook\.office365\.com)$" } ;
        if(-not (get-variable rgxCertThumbprint -ea 0)){$rgxCertThumbprint = '[0-9a-fA-F]{40}' ; } ;
        if(-not (get-variable rgxCertFNameSuffix -ea 0)){$rgxCertFNameSuffix = '-([A-Z]{3})$' ; } ; 

        <# 4:45 PM 7/7/2022 workaround msal.ps bug: always ipmo it FIRST: "Get-msaltoken : The property 'Authority' cannot be found on this object. Verify that the property exists."
        # admin/SID module auto-install code (myBoxes UID split-perm CU, all else t AllUsers)
        $modname = 'MSAL.PS' ;
        $error.clear() ;
        $pltIMod = @{Name = $modname ; ErrorAction = 'Stop' ; verbose=$false} ;
        # this forces a specific rev into the ipmo!
        if ($MinimumVersion) { $pltIMod.add('MinimumVersion', $MinimumVersion.tostring()) } ;
        $error.clear() ;
        Try { Get-Module $modname -ErrorAction Stop | out-null } Catch {
            $smsg = "Import-Module w`n$(($pltIMod|out-string).trim())" ;
            if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            Import-Module @pltIMod ;
        } ; # IsImported
        #>
        $EOMmodname = 'ExchangeOnlineManagement' ;
        
        #*------v PSS & GMO VARIS v------
        # get-pssession session varis
        $EXOv1ConfigurationName = $EXOv2ConfigurationName = $EXoPConfigurationName = "Microsoft.Exchange" ;

        if(-not (get-variable EXOv1ConfigurationName -ea 0)){$EXOv1ConfigurationName = "Microsoft.Exchange" };
        if(-not (get-variable EXOv2ConfigurationName -ea 0)){$EXOv2ConfigurationName = "Microsoft.Exchange" };
        if(-not (get-variable EXoPConfigurationName -ea 0)){$EXoPConfigurationName = "Microsoft.Exchange" };

        if(-not (get-variable EXOv1ComputerName -ea 0)){$EXOv1ComputerName = 'ps.outlook.com' };
        if(-not (get-variable EXOv1runspaceConnectionInfoAppName -ea 0)){$EXOv1runspaceConnectionInfoAppName = '/PowerShell-LiveID'  };
        if(-not (get-variable EXOv1runspaceConnectionInfoPort -ea 0)){$EXOv1runspaceConnectionInfoPort = '443' };

        if(-not (get-variable EXOv2ComputerName -ea 0)){$EXOv2ComputerName = 'outlook.office365.com' ;}
        if(-not (get-variable EXOv2Name -ea 0)){$EXOv2Name = "ExchangeOnlineInternalSession*" ; }
        if(-not (get-variable rgxEXoPrunspaceConnectionInfoAppName -ea 0)){$rgxEXoPrunspaceConnectionInfoAppName = '^/(exadmin|powershell)$'}; 
        if(-not (get-variable EXoPrunspaceConnectionInfoPort -ea 0)){$EXoPrunspaceConnectionInfoPort = '80' } ; 
        # gmo varis
        if(-not (get-variable rgxEXOv1gmoDescription -ea 0)){$rgxEXOv1gmoDescription = "^Implicit\sremoting\sfor\shttps://ps\.outlook\.com/PowerShell" }; 
        if(-not (get-variable EXOv1gmoprivatedataImplicitRemoting -ea 0)){$EXOv1gmoprivatedataImplicitRemoting = $true };
        if(-not (get-variable rgxEXOv2gmoDescription -ea 0)){$rgxEXOv2gmoDescription = "^Implicit\sremoting\sfor\shttps://outlook\.office365\.com/PowerShell" }; 
        if(-not (get-variable EXOv2gmoprivatedataImplicitRemoting -ea 0)){$EXOv2gmoprivatedataImplicitRemoting = $true } ;
        if(-not (get-variable rgxExoPsessionstatemoduleDescription -ea 0)){$rgxExoPsessionstatemoduleDescription = '/(exadmin|powershell)$' };
        if(-not (get-variable EXOv1GmoFilter -ea 0)){$EXOv1GmoFilter = 'tmp_*' } ; 
        if(-not (get-variable EXOv2GmoNoWinRMFilter -ea 0)){$EXOv2GmoNoWinRMFilter = 'tmpEXO_*' };
        # add get-connectioninformation.ConnectionURI targeting rgxs for CCMS vs EXO
        if(-not $rgxConnectionUriEXO){$rgxConnectionUriEXO = 'https://outlook\.office365\.com'} ; 
        if(-not $rgxConnectionUriEXO){$rgxConnectionUriCCMS = 'https://ps\.compliance\.protection\.outlook\.com'} ; 
        #*------^ END PSS & GMO VARIS ^------

        # * 11:02 AM 4/4/2023 reduced the ipmo and vers chk block, removed the lengthy gmo -list; and any autoinstall. Assume EOM is installed, & break if it's not
        #region EOMREV ; #*------v EOMREV Check v------
        #$EOMmodname = 'ExchangeOnlineManagement' ;
        $pltIMod = @{Name = $EOMmodname ; ErrorAction = 'Stop' ; verbose=$false} ;
        # do a gmo first, faster than gmo -list
        if([version]$EOMMv = (Get-Module @pltIMod| sort version | select -last 1 ).version){}
        elseif([version]$EOMMv = (Get-Module -ListAvailable @pltIMod| sort version | select -last 1 ).version){} 
        else { 
            $smsg = "$($EOMmodname) PowerShell v$($MinNoWinRMVersion) module is required, do you want to install it?" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Prompt } 
            else{ write-host -foregroundcolor YELLOW "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            $bRet=Read-Host "Enter YYY to continue. Anything else will exit"  ; 
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
        if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
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

        if(-not $Credential){
            if($UserRole){
                $smsg = "Using specified -UserRole:$( $UserRole -join ',' )" ;
                if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            #} else { $UserRole = @('SID','CSVC') } ;
            # S&C doesnt generally have cert or svc support
            } else { $UserRole = @('SID') } ;
            if($TenOrg){
                $smsg = "Using explicit -TenOrg:$($TenOrg)" ; 
            } else { 
                switch -regex ($env:USERDOMAIN){
                    ([regex]('(' + (( @($TORMeta.legacyDomain,$CMWMeta.legacyDomain)  |foreach-object{[regex]::escape($_)}) -join '|') + ')')).tostring() {$TenOrg = $env:USERDOMAIN.substring(0,3).toupper() } ;
                    $TOLMeta.legacyDomain {$TenOrg = 'TOL' }
                    default {throw "UNRECOGNIZED `$env:USERDOMAIN!:$($env:USERDOMAIN)" ; exit ; } ;
                } ;  
                $smsg = "Imputed `$TenOrg from logged on USERDOMAIN:$($TenOrg)" ;             
            } ; 
            if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;             
            
            $o365Cred = $null ;
            $pltGTCred=@{TenOrg=$TenOrg ; UserRole= $UserRole; verbose=$($verbose)} ;
            $smsg = "get-TenantCredentials w`n$(($pltGTCred|out-string).trim())" ; 
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level verbose } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            $o365Cred = get-TenantCredentials @pltGTCred ;

            if($o365Cred.credType -AND $o365Cred.Cred -AND $o365Cred.Cred.gettype().fullname -eq 'System.Management.Automation.PSCredential'){
                $smsg = "(validated `$o365Cred contains .credType:$($o365Cred.credType) & `$o365Cred.Cred.username:$($o365Cred.Cred.username)" ; 
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                $Credential = $o365Cred.Cred ;
            } else { 
                $smsg = "UNABLE TO RESOLVE FUNCTIONAL CredType/UserRole from specified explicit -Credential:$($Credential.username)!" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 

                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                break ; 
            } ; 
            
        } else { 
            # test-exotoken only applies if $UseConnEXO  $false
            $TenOrg = get-TenantTag -Credential $Credential ;
        } ;
        # build the cred etc once, for all below:
        $pltCXO=[ordered]@{
            Credential = $Credential ;
            verbose = $($verbose) ; 
            erroraction = 'STOP' ;
        } ;
        if((gcm connect-CCMS).Parameters.keys -contains 'silent'){
            $pltCXO.add('Silent',$false) ;
        } ;

        # defer to resolve-UserNameToUserRole -Credential $Credential
        <# need certtag further down, for credential align test
        if($credential.username -match $rgxCertThumbprint){
            $smsg =  "(UserName:Certificate Thumbprint detected)"
            if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            if($tcert = get-childitem -path "Cert:\CurrentUser\My\$($credential.username)"){
                $certUname = $tcert.friendlyname ; 
                $certTag = [regex]::match($certUname,$rgxCertFNameSuffix).captures[0].groups[1].value ; 
                $smsg = "(using CBA:cred:$($certTag):$([string]$tcert.friendlyname))" ; 
                if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            } else { 
                $smsg = "UNABLE TO RESOLVE `$TENORG:$($TenOrg) TO FUNCTIONAL `$$($TenOrg)meta.o365_TenantDomain!" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } 
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                throw $smsg ; 
                Break ; 
            } ;
        } ; 
        #>
        $uRoleReturn = resolve-UserNameToUserRole -Credential $Credential ; 
        if($credential.username -match $rgxCertThumbprint){
            $certTag = $uRoleReturn.TenOrg ; 
        } ; 

    } ;  # BEG-E
    PROCESS{
        $bExistingEXOGood = $false ; 
        $exov2Good = $exov3Good = $null ; 
        if( $legXPSSession = Get-PSSession | where-object {$_.ConfigurationName -like "Microsoft.Exchange" -AND $_.Name -match "^(Session|WinRM)\d*" } ){
            # ignore state & Avail, close the conflicting legacy conn's
            $smsg = "(existing legacy-EXO or Broken connections found, closing)" ; 
            if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            # Disconnect-CCMS ; Disconnect-PssBroken ;Start-Sleep -Seconds 3; # 8:36 AM 4/2/2024 deprecated cmd rmv
            Disconnect-CCMS ; Start-Sleep -Seconds 3;
            $bExistingEXOGood = $false ; 
        } ; 
        <# 3:50 PM 4/7/2022 ExchangeOnlinemanagement has a *bug*
        [Issue using ExchangeOnlineManagement v2.0.4 module to connect to Exchange Online remote powershell (EXO) and Exchange On-Prem remote powershell (EXOP) in same powershell window - Microsoft Q&A - docs.microsoft.com/](https://docs.microsoft.com/en-us/answers/questions/451786/issue-using-exchangeonlinemanagement-v204-module-t.html?childToView=804782#answer-804782)
        It *can't* cleanly reconnect the legacy EXO cmdlet dependant implicit-remote sessions, when they time out
        if there's an *existing* Exchange Onprem implicit remote session. 

        No fix, lame workaround: close all implicit remote sessions, before reopening EMO *first*, and then reconnecting EXOnPrem

        #>
        # check for existing implicit remote EXOnPrem sessions we have to kill first (and then post-reconnect)


        #clear invalid existing EXOv2 PSS's
       # fix typo from Name -eq to -like 
       # new token prop expir support:
       #((Get-PSSession | where Name -like ExchangeOnlineInternalSession*).TokenExpiryTime -lt (get-date ))
        #$exov2Broken = Get-PSSession | where-object {$_.ConfigurationName -like "Microsoft.Exchange" -AND (
        #    $_.Name -like "ExchangeOnlineInternalSession*") -AND $_.State -like "*Broken*"}
        # add token support clause: -AND ($_.TokenExpiryTime -lt (get-date )), leaving it out, may be a usecase with Broken but not post expiration
        # just fix the non-wildcard -like's w proper -eq's 
        $exov2Broken = Get-PSSession | where-object {$_.ConfigurationName -eq "Microsoft.Exchange" -AND (
            $_.Name -like "ExchangeOnlineInternalSession*") -AND $_.State -eq "Broken" }
        $exov2Closed = Get-PSSession | where-object {$_.ConfigurationName -like "Microsoft.Exchange" -AND (
            $_.Name -like "ExchangeOnlineInternalSession*") -AND $_.State -like "*Closed*"}
        
        #if ($exov2Broken.count -gt 0){for ($index = 0 ;$index -lt $exov2Broken.count ;$index++){Remove-PSSession -session $exov2Broken[$index]} };
        #if ($exov2Closed.count -gt 0){for ($index = 0 ;$index -lt $exov2Closed.count ; $index++){Remove-PSSession -session $exov2Closed[$index] } } ; 
        # sub full Disconnect-CCMS for Remove-PSSession: dxo2 includes 'Clear-ActiveToken -TokenProvider $session.TokenProvider' in addition to remove-pssession
        $pltDXO2=[ordered]@{
            verbose = $($VerbosePreference -eq "Continue") ;        
            silent = $silent ; 
        } ;
        if ( ($exov2Broken.count -gt 0) -OR ($exov2Closed.count -gt 0)){
            $smsg = "Disconnect-CCMS w`n$(($pltDXO2|out-string).trim())" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
             Disconnect-CCMS @pltDXO2 ;
        };
        
        if($IsNoWinRM){
            # 9:44 AM 4/20/2023 missed $EOMMv ref upgrade
            if($EOMMv.major -ge 3) {
                if ((Get-ConnectionInformation).tokenStatus -eq 'Active') {
                    $exov3Good = $bExistingEXOGood = $true ; 
                } else { 
                    $exov3Good = $bExistingEXOGood = $false ; 
                } ; 
            } else { 

            }
        } else { 
            # appears MFA may not properly support passing back a session vari, so go right to strict hostname matches
            $exov2Good = Get-PSSession | where-object {$_.ConfigurationName -like "Microsoft.Exchange" -AND (
                $_.Name -like "ExchangeOnlineInternalSession*") -AND $_.State -like "*Opened*" -AND (
                $_.Availability -eq 'Available')} ; 
        } ; 
        if($exov2Good -OR $exov3Good ){
            if( get-command Get-xoAcceptedDomain -ea 0) {
                # add accdom caching
                #$TenOrg = get-TenantTag -Credential $Credential ;
                if(!(Get-Variable  -name "$($TenOrg)Meta").value.o365_AcceptedDomains){
                    set-Variable  -name "$($TenOrg)Meta" -value ( (Get-Variable  -name "$($TenOrg)Meta").value += @{'o365_AcceptedDomains' = (Get-xoAcceptedDomain).domainname} )
                } ;
                #if ((Get-xoAcceptedDomain).domainname.contains($Credential.username.split('@')[1].tostring())){
                if( ($credential.username -match $rgxCertThumbprint) -AND ((Get-Variable  -name "$($TenOrg)Meta" -ea 0).value.o365_Prefix -eq $certTag )){
                    # 9:59 AM 6/24/2022 need a case for CBA cert (thumbprint username)
                    # compare cert fname suffix to $xxxMeta.o365_Prefix
                    # validate that the connected EXO is to the CBA Cert tenant
                    $smsg = "(EXO Authenticated & Functional CBA cert:$($certTag),($($certUname)))" ;
                    if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                    $bExistingEXOGood = $isEXOValid = $true ;
                }elseif((Get-Variable  -name "$($TenOrg)Meta").value.o365_AcceptedDomains.contains($Credential.username.split('@')[1].tostring())){
                    # validate that the connected EXO is to the $Credential tenant    
                    $smsg = "(EXO Authenticated & Functional(AccDom):$($Credential.username.split('@')[1].tostring()))" ; 
                    if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                    $bExistingEXOGood = $true ; 
                # issue: found fresh bug in cxo: svcacct UPN suffix @tenantname.onmicrosoft.com, but testing against AccepteDomain, it's not in there (tho @DOMAIN.mail.onmicrosoft.comis)
                }elseif((Get-Variable  -name "$($TenOrg)Meta").value.o365_TenantDomain -eq ($Credential.username.split('@')[1].tostring())){
                    $smsg = "(EXO Authenticated & Functional(TenDom):$($Credential.username.split('@')[1].tostring()))" ; 
                    if($silent){}elseif ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    $bExistingEXOGood = $true ;
                } else { 
                    $smsg = "(Credential mismatch:disconnecting from existing EXO:$($eEXO.Identity) tenant)" ; 
                    if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                    Disconnect-CCMS ; 
                    $bExistingEXOGood = $false ; 
                } ; 
            } else { 
                # capture outlier: shows a session wo the test cmdlet, force reset
                Disconnect-CCMS ; 
                $bExistingEXOGood = $false ; 
            } ; 
        } ; 

        if($bExistingEXOGood -eq $false){
            
            $smsg = "connect-CCMS w $($credential.username):`n$(($pltCXO|out-string).trim())" ; 
            if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            #connect-CCMS -Credential $Credential -verbose:$($verbose) ; 
            connect-CCMS @pltCXO ;    
                       
        } ; 

    } ;  # PROC-E
    END {
        
        <# 1:10 PM 3/1/2024 there are no more pss's in eom, rem it
        $smsg = "Existing PSSessions:`n$((get-pssession|out-string).trim())" ; 
        if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        #>

        if( ($useCCMSConn -AND ($bExistingCCMSGood -eq $false)) -OR (-not($useCCMSConn) -AND $bExistingEXOGood -eq $false) ){
            if($CertTag -ne $null){
                $smsg = "(specifying detected `$CertTag:$($CertTag))" ; 
                if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                $oRet = test-EXOv2Connection -Credential $credential -CertTag $certtag -Prefix $Prefix -verbose:$($verbose) ; 
            } else { 
                $oRet = test-EXOv2Connection -Credential $credential -Prefix $Prefix -verbose:$($verbose) ; 
            } ; 

            if($useCCMSConn){
                $bExistingCCMSGood = $oRet.Valid ;
            } else { 
                $bExistingEXOGood = $oRet.Valid ;
            } ; 

            if($oRet.Valid){
                $pssEXOv2 = $oRet.PsSession ;
                $IsNoWinRM = $oRet.IsNoWinRM ; 
                <#$oRet
                PSSession     :
                IsNoWinRM     : True
                Valid         : True
                Prefix        : xo
                isEXO         : True
                isCCMS        : False
                ConnectionUri : https://outlook.office365.com
                #>        
                switch -regex ($oRet.ConnectionUri){
                    $rgxConnectionUriEXO {
                        if ($oRet.Valid -AND $oRet.isEXO){
                            $smsg = "connected to EXO" ;
                            $bExistingEXOGood = $isEXOValid = $true ;
                        } ;
                    }
                    $rgxConnectionUriCCMS {
                        if ($oRet.Valid -AND $oRet.isCCMS){
                            $smsg = "connected to CCMS" ;
                            $bExistingCCMSGood = $isCCMSValid = $true ;
                        } ;
                    }
                    default {
                        $bExistingEXOGood = $isEXOValid = $bExistingCCMSGood = $isCCMSValid = $FALSE ;
                        $smsg = "unreconized test-EXOv2Connection returned:`$oRet.ConnectionUri:$($oRet.ConnectionUri)!" 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                    }
                } ; 
      
                $smsg += ":Validated Connected to Tenant aligned with specified Credential: `$IsNoWinRM:$($IsNoWinRM)" ;
                if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            } else {
                $smsg = "NO VALID EXOV2/3 SESSION FOUND! (DISCONNECTING...)"
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } 
                else{ write-host -ForegroundColor yellow "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                # capture outlier: shows a session wo the test cmdlet, force reset
                $pltDXO=[ordered]@{
                    confirm = $false ;
                    erroraction = 'STOP' ;
                    whatif = $($whatif) ;
                } ;
                if($Prefix){
                    $pltDXO.add('ModulePrefix',$Prefix) ;
                }
                $smsg = "Disconnect-ExchangeOnline w`n$(($pltDXO|out-string).trim())" ;
                if($silent){}elseif ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                Disconnect-ExchangeOnline @pltDXO ;

                #  DisConnect-CCMS ; # CCMS
                #DisConnect-EXO ;
                $bExistingEXOGood = $false ;
            } ;       
        } else {
          <# 9:08 AM 4/2/2024 unneded post v205p5 - disabled the cod that sets vari, above
          if($bPreExoPPss){
              $smsg = "(EMO bug-workaround: reconnecting prior ExOP PssSession,"
              $smsg += "`nreconnect-Ex2010 -Credential $($pltRX10.Credential.username) -verbose:$($VerbosePreference -eq "Continue"))" ; 
              if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
              else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
              reconnect-Ex2010 -Credential $pltRX10.Credential -verbose:$($VerbosePreference -eq "Continue") ; 
          } ;
          #>
        } ; 

        if($VerbosePreference -eq "Continue"){
            # echo Exchange-tied PSS summary
            if($pssEXOP = Get-PSSession | 
                where-object { ($_.ConfigurationName -eq $EXoPConfigurationName) -AND (
                    $_.runspace.ConnectionInfo.AppName -match '^/(exadmin|powershell)$') -AND (
                    $_.runspace.ConnectionInfo.Port -eq '80') } ){
                $smsg = "`nExOP PSSessions:`n$(($pssEXOP | fl $pssprops|out-string).trim())" ; 
            } ;
            if($pssEXOv1 = Get-PSSession | 
                    where-object {$_.ConfigurationName -like $EXOv1ConfigurationName -AND (
                        $_.ComputerName -eq 'ps.outlook.com') -AND ($_.runspace.ConnectionInfo.AppName -eq '/PowerShell-LiveID') -AND (
                        $_.runspace.ConnectionInfo.Port -eq '443') }){
                $smsg += "`n`nEXOv1 PSSessions:`n$(($pssEXOv1 | fl $pssprops|out-string).trim())" ; 
            } ; 
            if($pssEXOv2 = Get-PSSession | 
                    where-object {$_.ConfigurationName -like $EXOv2ConfigurationName -AND (
                        $_.Name -like $EXOv2Name) -AND ($_.ComputerName -eq $EXOv2ComputerName)} ){
                $smsg += "`n`nEXOv2 PSSessions:`n$(($pssEXOv2 | fl $pssprops|out-string).trim())" ; 
            } ;
            if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            
            if($IsNoWinRM -AND ((get-module $EXOv2GmoNoWinRMFilter) -AND (get-module $EOMModName))){
                $smsg = "(native non-WinRM/Non-PSSession-based EXO connection detected.)" ; 
                if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            } ; 
        } ; 

        # 10:37 AM 4/18/2023: Rem this: Been seldom capturing returns: that's bound to contaiminate pipeline! May have planned to grab and compare, but never really implemented
        #$bExistingEXOGood | write-output ;
        # splice in console color scheming
    }  # END-E
}

#*------^ Reconnect-CCMS.ps1 ^------