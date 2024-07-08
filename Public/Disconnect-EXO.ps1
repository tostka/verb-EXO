# Disconnect-EXO.ps1

#*------v Disconnect-EXO.ps1 v------
Function Disconnect-EXO {
    <#
    .SYNOPSIS
    Disconnect-EXO - Remove all the existing exchange online connections (incl EMOv1/2 PSSessions & EOM3+ nonWinRM - closes anything ConfigurationName: Microsoft.Exchange -AND Name: ExchangeOnlineInternalSession*)
    .NOTES
    Version     : 1.0.0
    Author      : Todd Kadrie
    Website     :	http://www.toddomation.com
    Twitter     :	@tostka / http://twitter.com/tostka
    CreatedDate : 2020-03-03
    FileName    : 
    License     : 
    Copyright   : 
    Github      : https://github.com/tostka/verb-exo
    Tags        : Powershell,ExchangeOnline,Exchange,RemotePowershell,Connection,MFA
    AddedCredit : ExactMike Perficient
    AddedWebsite:	https://social.technet.microsoft.com/Forums/msonline/en-US/f3292898-9b8c-482a-86f0-3caccc0bd3e5/exchange-powershell-monitoring-remote-sessions?forum=onlineservicesexchange
    AddedTwitter:	
    REVISIONS   :
    * 1:25 PM 7/8/2024 spliced in cxo constatns
    * 9:47 am 4/9/2024:validated updated disconnect-exo -prefix cc -verbose ; 
    * 10:59 AM 4/18/2023 step debugs ; consolidating Disconnect-EXO2 into Disconnect-EXO, aliasing dxo2,Disconnect-EXO2; removing those originals
    * 2:02 PM 4/17/2023 rev: $MinNoWinRMVersion from 2.0.6 => 3.0.0.
    * 12:42 PM 4/17/2023 restored *dxo* 7/26/21 vers; had overwritten on 3/29/22 wiith a copy of dxo2! Needs a verb-exo rebuild to complete.
    * 11:02 AM 4/4/2023 reduced the ipmo and vers chk block, removed the lengthy gmo -list; and any autoinstall. Assume EOM is installed, & break if it's not; 
    add $exov3Good and diff EMOv2 from EMOv3 sessions.
    * 3:14 pm 3/29/2023: REN'D $modname => $EOMModName
    * 3:59 PM 8/2/2022 got through dbugging EOM v205 SID interactive pass, working ; added -MinNoWinRMVersion ; fully works from mybox w v206p6, cEOM connection, with functional prefix.
    * 4:07 PM 7/26/2022 found that MS code doesn't chk for multi vers's installed, when building .dll paths: wrote in code to take highest version.
    * 3:30 PM 7/25/2022 tests against CBA & SID interactive creds on EOM v205, need to debug now against EOM v206p6, to accomodate PSSession-less connect & test code.
    * 10:34 AM 4/4/2022 updated CBH; added -MinimumVersion, defaulted, to support on-the-fly exemption
    * 3:54 PM 4/1/2022 add missing $silent param (had support, but no param)
    * 3:03 PM 3/29/2022 rewrote to reflect current specs in v2.0.5 of ExchangeOnlineManagement:Disconnect-ExchangeOnlineManagement cmds
    # here down is dxo orig revs
    * 11:54 AM 3/31/2021 added verbose suppress on remove-module/session commands
    * 1:14 PM 3/1/2021 added color reset
    * 3:24 PM 7/24/2020 updated to support tenant-alignment & sub'd out showdebug for verbose
    * 11:50 AM 5/27/2020 added alias:dxo win func
    * 2:34 PM 4/20/2020 added local $rgxExoPsHostName
    * 8:45 AM 3/3/2020 public cleanup
    * 9:52 AM 11/20/2019 reviewed for credential matl, no way to see the credential on a given pssession, so there's no way to target and disconnect discretely. It's a shotgun close.
    # 10:27 AM 6/20/2019 switched to common $rgxExoPsHostName
    # 1:12 PM 11/7/2018 added Disconnect-PssBroken
    # 11:23 AM 7/10/2018: made exo-only (was overlapping with CCMS)
    # 8:47 AM 6/2/2017 cleaned up deadwood, simplified pshelp
    * 8:49 AM 3/15/2017 Disconnect-EXO: add Remove-PSTitleBar 'EXO' to clean up on disconnect
    * 2/10/14 posted version
    .DESCRIPTION
    Disconnect-EXO - Remove all the existing exchange online connections (incl EMOv1/2 PSSessions & EOM3+ nonWinRM - closes anything ConfigurationName: Microsoft.Exchange -AND Name: ExchangeOnlineInternalSession*)
    Updated to match v2.0.5 of ExchangeOnlineMangement: Unlike the  v1.0.1 'disconnect', 
    this also implements new Clear-ActiveToken support, to reset the token as well as the session. 
    Doesn't support targeting session id, just wacks all sessions matching the configurationname & name of an EXOv2 pssession.
    .PARAMETER  Prefix
    [verb]-PREFIX[command] PREFIX string for clearly marking cmdlets sourced in this connection [-Prefix tag]
    PARAMETER MinimumVersion
    MinimumVersion required for ExchangeOnlineManagement module (defaults to '2.0.5')[-MinimumVersion '2.0.6']
    .PARAMETER MinNoWinRMVersion
    MinimumVersion required for Non-WinRM connections (of ExchangeOnlineManagement module (defaults to '3.0.0')[-MinimumVersion '2.0.6']
    .PARAMETER silent
    Switch to suppress all non-error echos
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    None. Returns no objects or output.
    .EXAMPLE
    Disconnect-EXO;
    Disconnect all EXOv2 ConfigurationName: Microsoft.Exchange -AND Name: ExchangeOnlineInternalSession* pssession
    .EXAMPLE
    Disconnect-EXO -silent;
    Demos use of the silent parameter to suppress output of details
    .LINK
    Github      : https://github.com/tostka/verb-exo
    #>
    [CmdletBinding()]
    [Alias('dxo','dxo2','Disconnect-EXO2')]
    Param(
        [Parameter(HelpMessage = "[verb]-PREFIX[command] PREFIX string for clearly marking cmdlets sourced in this connection [-Prefix tag]")]
            [string]$Prefix = 'xo',
        [Parameter(HelpMessage = "MinimumVersion required for ExchangeOnlineManagement module (defaults to '2.0.5')[-MinimumVersion '2.0.6']")]
            [version] $MinimumVersion = '2.0.5',
        [Parameter(HelpMessage = "MinimumVersion required for Non-WinRM connections (of ExchangeOnlineManagement module (defaults to '3.0.0')[-MinimumVersion '2.0.6']")]
            [version] $MinNoWinRMVersion = '3.0.0',
        [Parameter(HelpMessage="Silent output (suppress status echos)[-silent]")]
            [switch] $silent

    ) 
    $verbose = ($VerbosePreference -eq "Continue") ; 
    
    #*------v PSS & GMO VARIS v------
    # move into a param
    #$MinNoWinRMVersion = '3.0.0' ; 
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
    $EOMmodname = 'ExchangeOnlineManagement' ;
    $EXOv1ConfigurationName = $EXOv2ConfigurationName = $EXoPConfigurationName = "Microsoft.Exchange" ;
    if(-not (gv EXOv1ComputerName -ea 0 )){$EXOv1ComputerName = 'ps.outlook.com' };
    if(-not (gv EXOv1runspaceConnectionInfoAppName -ea 0 )){$EXOv1runspaceConnectionInfoAppName = '/PowerShell-LiveID'  };
    if(-not (gv EXOv1runspaceConnectionInfoPort -ea 0 )){$EXOv1runspaceConnectionInfoPort = '443' };

    if(-not (gv EXOv2ComputerName -ea 0 )){$EXOv2ComputerName = 'outlook.office365.com' ;}
    if(-not (gv EXOv2Name -ea 0 )){$EXOv2Name = "ExchangeOnlineInternalSession*" ; }
    #if(-not (gv rgxEXoPrunspaceConnectionInfoAppName -ea 0 )){$rgxEXoPrunspaceConnectionInfoAppName = '^/(exadmin|powershell)$'}; 
    #if(-not (gv EXoPrunspaceConnectionInfoPort -ea 0 )){$EXoPrunspaceConnectionInfoPort = '80' } ; 
    # gmo varis
    #if(-not (gv rgxExoPsHostName -ea 0 )){ $rgxExoPsHostName = "^(ps\.outlook\.com|outlook\.office365\.com)$" } ;
    #if(-not (gv rgxEXOv1gmoDescription -ea 0 )){$rgxEXOv1gmoDescription = "^Implicit\sremoting\sfor\shttps://ps\.outlook\.com/PowerShell" }; 
    #if(-not (gv EXOv1gmoprivatedataImplicitRemoting -ea 0 )){$EXOv1gmoprivatedataImplicitRemoting = $true };
    #if(-not (gv rgxEXOv2gmoDescription -ea 0 )){$rgxEXOv2gmoDescription = "^Implicit\sremoting\sfor\shttps://outlook\.office365\.com/PowerShell" }; 
    #if(-not (gv EXOv2gmoprivatedataImplicitRemoting -ea 0 )){$EXOv2gmoprivatedataImplicitRemoting = $true } ;
    #if(-not (gv rgxExoPsessionstatemoduleDescription -ea 0 )){$rgxExoPsessionstatemoduleDescription = '/(exadmin|powershell)$' };
    #if(-not (gv EXOv2StateOK -ea 0 )){$EXOv2StateOK = 'Opened'} ; 
    #if(-not (gv EXOv2AvailabilityOK -ea 0 )){$EXOv2AvailabilityOK = 'Available'} ; 
    #if(-not (gv EXOv2RunStateBad -ea 0 )){ $EXOv2RunStateBad = 'Broken'} ;
    #if(-not (gv EXOv1GmoFilter -ea 0 )){$EXOv1GmoFilter = 'tmp_*' } ; 
    if(-not (gv EXOv2GmoNoWinRMFilter -ea 0 )){$EXOv2GmoNoWinRMFilter = 'tmpEXO_*' };
    # add get-connectioninformation.ConnectionURI targeting rgxs for CCMS vs EXO
    if(-not $rgxConnectionUriEXO){$rgxConnectionUriEXO = 'https://outlook\.office365\.com'} ; 
    if(-not $rgxConnectionUriEXO){$rgxConnectionUriCCMS = 'https://ps\.compliance\.protection\.outlook\.com'} ; 
    $sTitleBarTag = @("EXO2") ;
    #*------^ END PSS & GMO VARIS ^------

    $pssProps = 'Id','Name','ComputerType','ComputerName','ConfigurationName','State','Availability',
        @{name="TokenExpiryTime";expression={get-date $_.TokenExpiryTime.date -format 'yyyyMMdd-HHmmtt'}};
    
    # disable prefix spec, unless actually blanked (e.g. centrally spec'd in profile).
    if (-not $Prefix) {
        $Prefix = 'xo' ; # 4:31 PM 7/29/2020 MS has RESERVED use of the 'exo' prefix [facepalm]
        #$Prefix = 'cc' ; # ccms variant
        $smsg = "(asserting Prefix:$($Prefix)" ;
        if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
    } ;
    if (($Prefix) -and ($Prefix -eq 'EXO')) {
        throw "Prefix 'EXO' is a reserved Prefix, please use a different prefix."
    }
    if($Prefix -eq 'cc'){
        # build in hybrid xo & ccms support, switch on the prefix spec
        $useCCMSConn = $true ; 
    }; 
    if($useCCMSConn){
        # respec userrole
        $UserRole = @('SID') ; 
        $sTitleBarTag = @("CCMS") ;
    } ; 

    # it's pulling the verb-EXO vers of disconnect-exchangeonline, force load the v206:
    #region EOMREV ; #*------v EOMREV Check v------
    # reduced the ipmo and vers chk block, removed the lengthy gmo -list; and any autoinstall. Assume EOM is installed, & break if it's not
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
    if($useConnExo){
        # 2:28 PM 8/1/2022 issue: it sometimes defers to the verb-EXO obsolete disconnect-exchangeonline (which doesn't properly resolve .dll paths, and doesn't exist/conflict in EOMv205), force load it out of the module
        if(-not (get-command -mod $EOMmodname -name Disconnect-ExchangeOnline -ea 0 )){
            $smsg = "(found dxo2, *not* sourced from EOM: ipmo -forcing EOM)" ; 
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            import-module -Name $EOMmodname -force -RequiredVersion $MinNoWinRMVersion ; 
        } ; 

        # just alias disconnect-ExchangeOnline, it retires token etc as well as closing PSS, but biggest reason is it's got a confirm, hard-coded, needs a function to override
        # flip back to the old d-eom call.

        if($xmod | where-object {$_.version -ge $MinNoWinRMVersion} ){
            $smsg = "EOM v3+ connection detected" ;
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
            TRY{
                $conns = Get-ConnectionInformation -ea STOP ;
            } CATCH {
                $ErrTrapd=$Error[0] ;
                $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                throw $smsg ;
                BREAK ;
            } ;
            if($Prefix){
                $conns = $conns | ?{$_.ModulePrefix -eq $Prefix} ;
            } ;
            switch -regex ($conns.ConnectionUri){
                $rgxConnectionUriEXO {
                    if ($conns.tokenStatus -eq 'Active') {
                        $smsg = "(connected to EXO)" ;
                        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                        $bExistingEXOGood = $isEXOValid = $true ;
                    } ;
                }
                $rgxConnectionUriCCMS {
                    if ($conns.tokenStatus -eq 'Active') {
                        $smsg = "(connected to CCMS)" ;
                        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                        $bExistingCCMSGood = $isCCMSValid = $true ;
                    } ;
                }
                default {
                    $bExistingEXOGood = $isEXOValid = $bExistingCCMSGood = $isCCMSValid = $FALSE
                }
            } ;
        }
        # you can use -ConnectionID or -ModulePrefix, but not both, we've already filtered w connid, so use it
        $pltDXO=[ordered]@{
            #ModulePrefix = $Prefix ;
            confirm = $false ;
            erroraction = 'STOP' ;
            whatif = $($whatif) ;
        } ;
        #$prpConnInf = 'ConnectionId','ConnectionUri','State','TokenStatus' ; 
        if($xmod | where-object {$_.version -ge $MinNoWinRMVersion } ){
            if($conns.ConnectionID){
                $pltDXO.add('ConnectionId',$conns.ConnectionID)
                $smsg = "targeting filtered ConnectionID:$($conns.ConnectionID)`n$(($conns | ft -a $prpConnInf |out-string).trim())" ; 
                if($silent){}elseif($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
            } ; 
        } ; 
        if ($conns) {
            $smsg = "Disconnect-ExchangeOnline w`n$(($pltDXO|out-string).trim())" ;
            if($silent){}elseif ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
            Disconnect-ExchangeOnline @pltDXO ;
            #Disconnect-ExchangeOnline -confirm:$false ; 
            # just use the updated RemoveExistingEXOPSSession
            #PRIOR: RemoveExistingEXOPSSession -Verbose:$false ;
            # v2.0.5 3:01 PM 3/29/2022 no longer exists
        } else { 
            $smsg = "(no existing session matched)" ; 
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        } ; 
    } else { 
        $smsg = "(EXOv2 EOM v205 nonWinRM code in use...)" ; 
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        $smsg = "EXOv2 EOM v205 and below are NO LONGER SUPPORTED!" ; 
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
        THROW $SMSG ; 
        BREAK ; 
    } ; 
    # poll session types
    $existingPSSession = Get-PSSession | 
        Where-Object {$_.ConfigurationName -like "Microsoft.Exchange" -and $_.Name -like "ExchangeOnlineInternalSession*"} ;
    if ($existingPSSession.count -gt 0) {
        for ($index = 0; $index -lt $existingPSSession.count; $index++){
            $session = $existingPSSession[$index]
            $smsg = "Remove-PSSession w`n$(($session | format-table -a  $pssprops|out-string).trim())" ; 
            if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            Remove-PSSession -session $session
            $smsg = "Removed the PSSession $($session.Name) connected to $($session.ComputerName)"
            if($silent){}elseif ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            # Remove any active access token from the cache
            $pltCAT=[ordered]@{
                TokenProvider=$session.TokenProvider ; 
            } ;
            if(get-command Clear-ActiveToken -ea 0){
                $smsg = "Clear-ActiveToken w`n$(($pltCAT|out-string).trim())" ; 
                if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                Clear-ActiveToken @pltCAT ;
            } ; 
            # Remove any previous modules loaded because of the current PSSession
            if ($session.PreviousModuleName -ne $null){
                if ((Get-Module $session.PreviousModuleName).Count -ne 0){
                    Remove-Module -Name $session.PreviousModuleName -ErrorAction SilentlyContinue
                }
                $session.PreviousModuleName = $null
            } ; 
            # Remove any leaked module in case of removal of broken session object
            if ($session.CurrentModuleName -ne $null){
                if ((Get-Module $session.CurrentModuleName).Count -ne 0){
                    Remove-Module -Name $session.CurrentModuleName -ErrorAction SilentlyContinue ; 
                } ;  
            }  ; 
        } ;  # loop-E
    } ; # if-E $existingPSSession.count -gt 0
    
    #Disconnect-PssBroken -verbose:$false ;
    Remove-PSTitlebar $sTitleBarTag #-verbose:$($VerbosePreference -eq "Continue");
    #[console]::ResetColor()  # reset console colorscheme
} ; 

#*------^ Disconnect-EXO.ps1 ^------