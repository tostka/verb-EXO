#*------v Disconnect-EXO2.ps1 v------
Function Disconnect-EXO2 {
    <#
    .SYNOPSIS
    Disconnect-EXO2 - Remove all the existing exchange online PSSessions (closes anything ConfigurationName: Microsoft.Exchange -AND Name: ExchangeOnlineInternalSession*)
    .NOTES
    Version     : 1.0.0
    Author      : Todd Kadrie
    Website     :	http://www.toddomation.com
    Twitter     :	@tostka / http://twitter.com/tostka
    CreatedDate : 2020-03-03
    FileName    : 
    License     : 
    Copyright   : 
    Github      : https://github.com/tostka
    Tags        : Powershell,ExchangeOnline,Exchange,RemotePowershell,Connection,MFA
    REVISIONS   :
    * 2:02 PM 4/17/2023 rev: $MinNoWinRMVersion from 2.0.6 => 3.0.0.
    * 11:02 AM 4/4/2023 reduced the ipmo and vers chk block, removed the lengthy gmo -list; and any autoinstall. Assume EOM is installed, & break if it's not; 
    add $exov3Good and diff EMOv2 from EMOv3 sessions.
    * 3:14 pm 3/29/2023: REN'D $modname => $EOMModName
    * 3:59 PM 8/2/2022 got through dbugging EOM v205 SID interactive pass, working ; added -MinNoWinRMVersion ; fully works from mybox w v206p6, cEOM connection, with functional prefix.
    * 4:07 PM 7/26/2022 found that MS code doesn't chk for multi vers's installed, when building .dll paths: wrote in code to take highest version.
    * 3:30 PM 7/25/2022 tests against CBA & SID interactive creds on EOM v205, need to debug now against EOM v206p6, to accomodate PSSession-less connect & test code.
    * 10:34 AM 4/4/2022 updated CBH; added -MinimumVersion, defaulted, to support on-the-fly exemption
    * 3:54 PM 4/1/2022 add missing $silent param (had support, but no param)
    * 3:03 PM 3/29/2022 rewrote to reflect current specs in v2.0.5 of ExchangeOnlineManagement:Disconnect-ExchangeOnlineManagement cmds
    * 11:55 AM 3/31/2021 suppress verbose on module/session cmdlets
    * 1:14 PM 3/1/2021 added color reset
    * 9:55 AM 7/30/2020 EXO v2 version, adapted from Disconnect-EXO, + some content from RemoveExistingPSSession
    .DESCRIPTION
    Disconnect-EXO2 - Remove all the existing exchange online PSSessions (closes anything ConfigurationName: Microsoft.Exchange -AND Name: ExchangeOnlineInternalSession*)
    Updated to match v2.0.5 of ExchangeOnlineMangement: Unlike the  v1.0.1 'disconnect', 
     this also implements new Clear-ActiveToken support, to reset the token as well as the session. 
     Doesn't support targeting session id, just wacks all sessions matching the configurationname & name of an EXOv2 pssession.
    .PARAMETER MinimumVersion
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
    Disconnect-EXO2;
    Disconnect all EXOv2 ConfigurationName: Microsoft.Exchange -AND Name: ExchangeOnlineInternalSession* pssession
    .EXAMPLE
    Disconnect-EXO2 -silent;
    Demos use of the silent parameter to suppress output of details
    .LINK
    #>
    [CmdletBinding()]
    [Alias('dxo2')]
    Param(
        [Parameter(HelpMessage = "MinimumVersion required for ExchangeOnlineManagement module (defaults to '2.0.5')[-MinimumVersion '2.0.6']")]
        [version] $MinimumVersion = '2.0.5',
        [Parameter(HelpMessage = "MinimumVersion required for Non-WinRM connections (of ExchangeOnlineManagement module (defaults to '3.0.0')[-MinimumVersion '2.0.6']")]
        [version] $MinNoWinRMVersion = '3.0.0',
        [Parameter(HelpMessage="Silent output (suppress status echos)[-silent]")]
        [switch] $silent

    ) 
    $verbose = ($VerbosePreference -eq "Continue") ; 

    $EOMmodname = 'ExchangeOnlineManagement' ;
    $ExoPowershellGalleryModule = "Microsoft.Exchange.Management.ExoPowershellGalleryModule.dll" ; # in EOM v205 hosts test|clear-ActiveToken( both nonexist in v206+)
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

    $pssProps = 'Id','Name','ComputerType','ComputerName','ConfigurationName','State','Availability',
        @{name="TokenExpiryTime";expression={get-date $_.TokenExpiryTime.date -format 'yyyyMMdd-HHmmtt'}};
    

    # it's pulling the verb-EXO vers of disconnect-exchangeonline, force load the v206:
    # admin/SID module auto-install code (myBoxes UID split-perm CU, all else t AllUsers)
    #region EOMREV ; #*------v EOMREV Check v------
    $EOMmodname = 'ExchangeOnlineManagement' ;
    $pltIMod = @{Name = $EOMmodname ; ErrorAction = 'Stop' ; verbose=$false} ;
    if($xmod = Get-Module $EOMmodname -ErrorAction Stop){ } else {
        $smsg = "Import-Module w`n$(($pltIMod|out-string).trim())" ;
        if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
        Try {
            Import-Module @pltIMod | out-null ;
            $xmod = Get-Module $EOMmodname -ErrorAction Stop ;
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
    if([version]$xmod.version -ge $MinNoWinRMVersion){$MinNoWinRMVersion = $xmod.version.tostring() ;}
    [boolean]$UseConnEXO = [boolean]([version]$xmod.version -ge $MinNoWinRMVersion) ; 
    #endregion EOMREV ; #*------^ END EOMREV Check  ^------
    if($useConnExo){
        # 2:28 PM 8/1/2022 issue: it sometimes defers to the verb-EXO obsolete disconnect-exchangeonline (which doesn't properly resolve .dll paths, and doesn't exist/conflict in EOMv205), force load it out of the module
        if(-not (get-command -mod 'ExchangeOnlineManagement' -name Disconnect-ExchangeOnline -ea 0 )){
            $smsg = "(found dxo2, *not* sourced from EOM: ipmo -forcing EOM)" ; 
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            import-module -Name 'ExchangeOnlineManagement' -force -RequiredVersion $MinNoWinRMVersion ; 
        } ; 
        # just alias disconnect-ExchangeOnline, it retires token etc as well as closing PSS, but biggest reason is it's got a confirm, hard-coded, needs a function to override
        # flip back to the old d-eom call.
        Disconnect-ExchangeOnline -confirm:$false ; 
        # just use the updated RemoveExistingEXOPSSession
        #PRIOR: RemoveExistingEXOPSSession -Verbose:$false ;
        # v2.0.5 3:01 PM 3/29/2022 no longer exists
    } else { 
        $smsg = "(EXOv2 EOM v205 nonWinRM code in use...)" ; 
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        #$EOMmodname = 'ExchangeOnlineManagement' ;
        #if(-not $EXOv2Name){$EXOv2Name = "ExchangeOnlineInternalSession*" ; } ; 
        #if(-not $EXOv2ConfigurationName){$EXOv2ConfigurationName = "Microsoft.Exchange" };
        $EOMgmtModulePath = split-path (get-module $EOMmodname -list | sort Version | select -last 1).Path ;
        if($IsCoreCLR){
            $EOMgmtModulePath = resolve-path -Path $EOMgmtModulePath\netcore ;
            $smsg = "(.netcore path in use:" ; 
        } else { 
            $EOMgmtModulePath = resolve-path -Path $EOMgmtModulePath\netFramework ; 
            $smsg = "(.netnetFramework path in use:" ;                 
        } ; 
        #$ExoPowershellGalleryModule = "Microsoft.Exchange.Management.ExoPowershellGalleryModule.dll" ;
        $ExoPowershellGalleryModulePath = join-path -path $EOMgmtModulePath -childpath $ExoPowershellGalleryModule ;
        if (-not(get-module $ExoPowershellGalleryModule.replace('.dll','') )) {
            Import-Module $ExoPowershellGalleryModulePath -Verbose:$false -ErrorAction 'STOP';
        } ;    
        if(-not (get-command -module $ExoPowershellGalleryModule.replace('.dll','') | ? Name -match '(clear|test)-ActiveToken')){
            throw "Unable to GCM clear-ActiveToken cmdlet!`n(as provided by:$($ExoPowershellGalleryModulePath))" ; 
        } ; 
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
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            # Remove any active access token from the cache
            $pltCAT=[ordered]@{
                TokenProvider=$session.TokenProvider ; 
            } ;
            $smsg = "Clear-ActiveToken w`n$(($pltCAT|out-string).trim())" ; 
            if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            Clear-ActiveToken @pltCAT ;
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
    
    Disconnect-PssBroken -verbose:$false ;
    Remove-PSTitlebar 'EXO2' #-verbose:$($VerbosePreference -eq "Continue");
    #[console]::ResetColor()  # reset console colorscheme
} ; 

#*------^ Disconnect-EXO2.ps1 ^------