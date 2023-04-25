#*------v Reconnect-EXO.ps1 v------
Function Reconnect-EXO {
   <#
    .SYNOPSIS
    Reconnect-EXO - Reestablish connection to Exchange Online (via EXO V2/V3 graph-api module)
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
    # 3:18 PM 4/19/2023 under EOM310: replc $xmod.version refs with $EOMMv...
    * 11:20 AM 4/18/2023 step debugs ;  consolidate reconnect-exo2 into reconnect-exo (alias reconnect-exo2 & rxo2)
    * 2:02 PM 4/17/2023 rev: $MinNoWinRMVersion from 2.0.6 => 3.0.0.
    # * 11:02 AM 4/4/2023 reduced the ipmo and vers chk block, removed the lengthy gmo -list; and any autoinstall. Assume EOM is installed, & break if it's not
    * 3:14 pm 3/29/2023: REN'D $modname => $EOMModName
    * 11:01 AM 12/21/2022 moved $pltCXO2 def out to always occur (was only when !$bExistingEXOGood )
    * 3:59 PM 8/2/2022 got through dbugging EOM v205 SID interactive pass, working ; added -MinimumVersion & - MinNoWinRMVersion ; fully works from mybox w v206p6, cEOM connection, with functional prefix.
    * 3:30 PM 7/25/2022 tests against CBA & SID interactive creds on EOM v205, need to debug now against EOM v206p6, to accomodate PSSession-less connect & test code.
    * 3:54 PM 4/1/2022 add missing $silent param (had support, but no param)
    * 11:04 AM 3/30/2022 recode for ExchangeOnlineManagement v2.0.5, fundemental revisions, since prior v1.0.1
    # below here is orig rxo revs
    * 9:03 AM 12/14/2021 cleaned comments
    * 1:17 PM 8/17/2021 added -silent param
    * 3:20 PM 3/31/2021 fixed pssess typo
    * 8:30 AM 10/22/2020 added $TenOrg, swapped looping meta resolve with 1-liner approach ; added AcceptedDom caching to the middle status test (suppress one more get-exoaccepteddomain call if possible)
    * 1:30 PM 9/21/2020 added caching of AcceptedDomain, dynamically into XXXMeta - checks for .o365_AcceptedDomains, and pops w (Get-exoAcceptedDomain).domainname when blank. 
        As it's added to the $global meta, that means it stays cached cross-session, completely eliminates need to dyn query per rxo, after the first one, that stocks the value
    * 2:39 PM 8/4/2020 fixed -match "^(Session|WinRM)\d*" rgx (lacked ^, mismatched EXOv2 conns)
    * 10:35 AM 7/28/2020 tweaked retry loop to not retry-sleep 1st attempt
    * 3:24 PM 7/24/2020 updated to support tenant-alignment & sub'd out showdebug for verbose
    * 11:48 AM 5/27/2020 added func alias:rxo within the func
    * 2:38 PM 4/20/2020 added local $rgxExoPsHostName
    * 8:45 AM 3/3/2020 public cleanup
    * 9:52 PM 1/16/2020 cleanup
    * 1:07 PM 11/25/2019 added *tol/*tor/*cmw alias variants for connect & reconnect
    * 9:52 AM 11/20/2019 spliced in credential matl
    * 2:55 PM 10/11/2018 connect-exo: added pre sleep skip on tryno 1
    * 8:04 AM 11/20/2017 code in a loop in the reconnect-exo, until it hits or 100% times out
    # 8:47 AM 6/2/2017 cleaned up deadwood, simplified pshelp
    * 7:58 AM 3/15/2017 ren Disconnect/Connect/Reconnect-EXO => Disconnect/Connect/Reconnect-EXO, added pss, doing tweaks to put into prod use
    * 2/10/14 posted version
    .DESCRIPTION
    Reconnect-EXO - Reestablish connection to Exchange Online (via EXO V2/V3 graph-api module)
    .PARAMETER  Credential
    Credential to use for this connection [-credential 'account@domain.com']
     .PARAMETER MinimumVersion
    MinimumVersion required for ExchangeOnlineManagement module (defaults to '2.0.5')[-MinimumVersion '2.0.6']
    .PARAMETER MinNoWinRMVersion
    MinimumVersion required for Non-WinRM connections (of ExchangeOnlineManagement module (defaults to '3.0.0')[-MinimumVersion '2.0.6']
    .PARAMETER ProxyEnabled
    Use Proxy-Aware SessionOption settings [-ProxyEnabled]
    .PARAMETER Credential
    Credential to use for this connection [-credential [credential obj variable]
    .PARAMETER showDebug
    Debugging Flag [-showDebug]
    .PARAMETER silent
    Switch to specify suppression of all but warn/error echos.
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    None. Returns no objects or output.
    .EXAMPLE
    PS>  Reconnect-EXO;
    Reconnect EXO connection
    .EXAMPLE
    PS>  Reconnect-EXO -credential $cred ;
    Reconnect EXO connection with explicit [pscredential] object credential specified
    .EXAMPLE
    PS>  $batchsize = 10 ;
    PS>  $RecordCount=$mr.count #this is the array of whatever you are processing ;
    PS>  $b=0 #this is the initialization of a variable used in the do until loop below ;
    PS>  $mrs = @() ;
    PS>  do {
    PS>      Write-Progress -Activity "Getting move request statistics for all $wave move requests." -Status "Processing Records $b through $($b+$batchsize) of $RecordCount." -PercentComplete ($b/$RecordCount*100) ;
    PS>      $mrs += $mr | Select-Object -skip $b -first $batchsize | foreach-object {Reconnect-EXO; $_ | Get-OLMoveRequestStatistics} ;
    PS>      $b=$b+$batchsize ;
    PS>      } ;
    PS>  until ($b -gt $RecordCount) ;
    Demo use of a Reconnect-EXO call mid-looping process to ensure connection stays available for duration of long-running process.    
    .LINK
    Github      : https://github.com/tostka/verb-exo
    #>
    [CmdletBinding()]
    [Alias('rxo','reconnect-exo2','rxo2')]
    PARAM(
        [Parameter(HelpMessage = "MinimumVersion required for ExchangeOnlineManagement module (defaults to '2.0.5')[-MinimumVersion '2.0.6']")]
        [version] $MinimumVersion = '2.0.5',
        [Parameter(HelpMessage = "MinimumVersion required for Non-WinRM connections (of ExchangeOnlineManagement module (defaults to '3.0.0')[-MinimumVersion '2.0.6']")]
        [version] $MinNoWinRMVersion = '3.0.0',
        [Parameter(HelpMessage="Use Proxy-Aware SessionOption settings [-ProxyEnabled]")]
        [boolean]$ProxyEnabled = $False,
        [Parameter(HelpMessage="Credential to use for this connection [-credential [credential obj variable]")]
        [System.Management.Automation.PSCredential]$Credential = $global:credo365TORSID,
        [Parameter(HelpMessage="Debugging Flag [-showDebug]")]
        [switch] $showDebug,
        [switch]$silent
    ) ;
    BEGIN{
        $verbose = ($VerbosePreference -eq "Continue") ; 
        if(!$rgxExoPsHostName){$rgxExoPsHostName="^(ps\.outlook\.com|outlook\.office365\.com)$" } ;

        # 4:45 PM 7/7/2022 workaround msal.ps bug: always ipmo it FIRST: "Get-msaltoken : The property 'Authority' cannot be found on this object. Verify that the property exists."
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

        $EOMmodname = 'ExchangeOnlineManagement' ;
        
        #*------v PSS & GMO VARIS v------
        # get-pssession session varis
        $EXOv1ConfigurationName = $EXOv2ConfigurationName = $EXoPConfigurationName = "Microsoft.Exchange" ;

        if(-not (gv EXOv1ConfigurationName -ea 0)){$EXOv1ConfigurationName = "Microsoft.Exchange" };
        if(-not (gv EXOv2ConfigurationName -ea 0)){$EXOv2ConfigurationName = "Microsoft.Exchange" };
        if(-not (gv EXoPConfigurationName -ea 0)){$EXoPConfigurationName = "Microsoft.Exchange" };

        if(-not (gv EXOv1ComputerName -ea 0)){$EXOv1ComputerName = 'ps.outlook.com' };
        if(-not (gv EXOv1runspaceConnectionInfoAppName -ea 0)){$EXOv1runspaceConnectionInfoAppName = '/PowerShell-LiveID'  };
        if(-not (gv EXOv1runspaceConnectionInfoPort -ea 0)){$EXOv1runspaceConnectionInfoPort -eq '443' };

        if(-not (gv EXOv2ComputerName -ea 0)){$EXOv2ComputerName = 'outlook.office365.com' ;}
        if(-not (gv EXOv2Name -ea 0)){$EXOv2Name = "ExchangeOnlineInternalSession*" ; }
        if(-not (gv rgxEXoPrunspaceConnectionInfoAppName -ea 0)){$rgxEXoPrunspaceConnectionInfoAppName = '^/(exadmin|powershell)$'}; 
        if(-not (gv EXoPrunspaceConnectionInfoPort -ea 0)){$EXoPrunspaceConnectionInfoPort = '80' } ; 
        # gmo varis
        if(-not (gv rgxEXOv1gmoDescription -ea 0)){$rgxEXOv1gmoDescription = "^Implicit\sremoting\sfor\shttps://ps\.outlook\.com/PowerShell" }; 
        if(-not (gv EXOv1gmoprivatedataImplicitRemoting -ea 0)){$EXOv1gmoprivatedataImplicitRemoting = $true };
        if(-not (gv rgxEXOv2gmoDescription -ea 0)){$rgxEXOv2gmoDescription = "^Implicit\sremoting\sfor\shttps://outlook\.office365\.com/PowerShell" }; 
        if(-not (gv EXOv2gmoprivatedataImplicitRemoting -ea 0)){$EXOv2gmoprivatedataImplicitRemoting = $true } ;
        if(-not (gv rgxExoPsessionstatemoduleDescription -ea 0)){$rgxExoPsessionstatemoduleDescription = '/(exadmin|powershell)$' };
        if(-not (gv EXOv1GmoFilter -ea 0)){$EXOv1GmoFilter = 'tmp_*' } ; 
        if(-not (gv EXOv2GmoNoWinRMFilter -ea 0)){$EXOv2GmoNoWinRMFilter = 'tmpEXO_*' };
        #*------^ END PSS & GMO VARIS ^------

        # * 11:02 AM 4/4/2023 reduced the ipmo and vers chk block, removed the lengthy gmo -list; and any autoinstall. Assume EOM is installed, & break if it's not
        #region EOMREV ; #*------v EOMREV Check v------
        #$EOMmodname = 'ExchangeOnlineManagement' ;
        $pltIMod = @{Name = $EOMmodname ; ErrorAction = 'Stop' ; verbose=$false} ;
        # do a gmo first, faster than gmo -list
        if([version]$EOMMv = (Get-Module @pltIMod).version){}
        elseif([version]$EOMMv = (Get-Module -ListAvailable @pltIMod).version){} 
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
        write-verbose "(Checking for WinRM support in this EOM rev...)" 
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

        # test-exotoken only applies if $UseConnEXO  $false
        $TenOrg = get-TenantTag -Credential $Credential ;

        # build the cred etc once, for all below:
        $pltCXO2=[ordered]@{
            Credential = $Credential ;
            verbose = $($verbose) ; 
            erroraction = 'STOP' ;
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
            Disconnect-Exo ; Disconnect-PssBroken ;Start-Sleep -Seconds 3;
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
        # sub full Disconnect-EXO2 for Remove-PSSession: dxo2 includes 'Clear-ActiveToken -TokenProvider $session.TokenProvider' in addition to remove-pssession
        $pltDXO2=[ordered]@{
            verbose = $($VerbosePreference -eq "Continue") ;        
            silent = $silent ; 
        } ;
        if ( ($exov2Broken.count -gt 0) -OR ($exov2Closed.count -gt 0)){
            $smsg = "Disconnect-EXO2 w`n$(($pltDXO2|out-string).trim())" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
             Disconnect-EXO2 @pltDXO2 ;
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
                if((Get-Variable  -name "$($TenOrg)Meta").value.o365_AcceptedDomains.contains($Credential.username.split('@')[1].tostring())){
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
                    DisConnect-EXO2 ; 
                    $bExistingEXOGood = $false ; 
                } ; 
            } else { 
                # capture outlier: shows a session wo the test cmdlet, force reset
                DisConnect-EXO2 ; 
                $bExistingEXOGood = $false ; 
            } ; 
        } ; 

        if($bExistingEXOGood -eq $false){
            
            $smsg = "connect-exo2 w $($credential.username):`n$(($pltCXO2|out-string).trim())" ; 
            if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            #connect-exo2 -Credential $Credential -verbose:$($verbose) ; 
            connect-exo2 @pltCXO2 ;               
        } ; 

    } ;  # PROC-E
    END {
        # if ( (get-module -name tmp_* | ForEach-Object { Get-Command -module $_.name -name 'Add-xoAvailabilityAddressSpace' -ea 0 }) -AND (test-EXOToken) ) {
        #if( (Get-PSSession | where-object {$_.ConfigurationName -like "Microsoft.Exchange" -AND $_.Name -like "ExchangeOnlineInternalSession*" -AND $_.State -like "*Opened*" -AND ($_.Availability -eq 'Available')}) -AND (test-EXOToken) ){ 
        $validated = $false ;
        if($UseConnEXO -AND $exov3Good  ){
            #connexo should smoothly recycle connections; only v1 requires manual detect & reconnect with basic auth creds
            $validated = $true ; 
        } else { 
            # cred is mandetory - err - in test-exotoken, push it through
            if( (Get-PSSession | where-object {
                    $_.ConfigurationName -like "Microsoft.Exchange" -AND $_.Name -like "ExchangeOnlineInternalSession*" -AND $_.State -like "*Opened*" -AND ($_.Availability -eq 'Available')
                }) -AND (test-EXOToken -Credential $pltCXO2.credential) ){ 
                # implement caching of accepteddoms into the XXXMeta, in the session (cut back on queries to EXO on acceptedom)
                # non-looping
                $validated = $true ; 
            } 
        } ; 
        if($validated){
            if( get-command Get-xoAcceptedDomain -ea 0) {
                #$TenOrg = get-TenantTag -Credential $Credential ;
                if(!(Get-Variable  -name "$($TenOrg)Meta").value.o365_AcceptedDomains){
                    set-Variable  -name "$($TenOrg)Meta" -value ( (Get-Variable  -name "$($TenOrg)Meta").value += @{'o365_AcceptedDomains' = (Get-xoAcceptedDomain).domainname} )
                } ;
            } ; 
            <#
            $credDom = ($Credential.username.split("@"))[1] ;
            $Metas=(get-variable *meta|Where-Object{$_.name -match '^\w{3}Meta$'}) ;
            foreach ($Meta in $Metas){
                if( ($credDom -eq $Meta.value.legacyDomain) -OR ($credDom -eq $Meta.value.o365_TenantDomain) -OR ($credDom -eq $Meta.value.o365_OPDomain)){
                    if(!$Meta.value.o365_AcceptedDomains){
                        set-variable -Name $meta.name -Value ((get-variable -name $meta.name).value  += @{'o365_AcceptedDomains' = (Get-xoAcceptedDomain).domainname} )
                    } ; 
                    break ;
                } ;
            } ;
            #>
            #if ((Get-xoAcceptedDomain).domainname.contains($Credential.username.split('@')[1].tostring())){
            #if($Meta.value.o365_AcceptedDomains.contains($Credential.username.split('@')[1].tostring())){
            if((Get-Variable  -name "$($TenOrg)Meta").value.o365_AcceptedDomains.contains($Credential.username.split('@')[1].tostring())){
                # validate that the connected EXO is to the $Credential tenant    
                $smsg = "(EXOv2 Authenticated & Functional:$($Credential.username.split('@')[1].tostring()))" ; 
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
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } 
                else{ write-ERROR "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                Disconnect-exo2 ; 
                $bExistingEXOGood = $false ; 
            } ;
        } ; 
    } ; # END-E 
}

#*------^ Reconnect-EXO.ps1 ^------