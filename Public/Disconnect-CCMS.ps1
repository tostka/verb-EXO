# Disconnect-CCMS.ps1

#*------v Disconnect-CCMS.ps1 v------
Function Disconnect-CCMS {
    <#
    .SYNOPSIS
    Disconnect-CCMS - Remove all the existing Security & Compliance connections - as identified via specified -Prefix param (default: cc)- (incl EMOv1/2 PSSessions & EOM3+ nonWinRM - closes anything ConfigurationName: Microsoft.Exchange -AND Name: ExchangeOnlineInternalSession*)
    .NOTES
    Version     : 1.0.0
    Author      : Todd Kadrie
    Website     :	http://www.toddomation.com
    Twitter     :	@tostka / http://twitter.com/tostka
    CreatedDate : 2020-03-03
    FileName    : Disconnect-CCMS
    License     : MIT License
    Copyright   : (c) 2020 Todd Kadrie
    Github      : https://github.com/tostka/verb-ccms
    Tags        : Powershell,ExchangeOnline,Exchange,RemotePowershell,Connection,MFA
    AddedCredit : ExactMike Perficient
    AddedWebsite:	https://social.technet.microsoft.com/Forums/msonline/en-US/f3292898-9b8c-482a-86f0-3caccc0bd3e5/exchange-powershell-monitoring-remote-sessions?forum=onlineservicesexchange
    AddedTwitter:	
    REVISIONS   :
    # 4:47 PM 7/8/2024 this is obsoleted; shifted all (re|dis)connect-CCMS functions into connect-exo & reconnect-exo: CCMS Sec & Compl connection mgmt is triggered via the -Prefix cc parameter (any other param is assumed to be native EXO; but -Prefix cc will always generate a connection to Sec & Compliance); 
    8:57 AM 4/2/2024 undbgd update of dxo, added param -prefix to drive targeting: disconnect-exchangeOnline -prefix 'cc' (to get s&c wo impacting exo sessions). add const $sTitleBarTag
    * 12:46 PM 3/5/2024 full step debug pass, with Disconnect-PssBroken -verbose:$false ; rem'd and EOM v3.4.0 in place: ran clean this time
    * 3:15 PM 3/1/2024 rem'd obsolete Disconnect-PssBroken line (old PSS-supporting call from EOM)
    * 2:51 PM 2/26/2024 add | sort version | select -last 1  on gmos, LF installed 3.4.0 parallel to 3.1.0 and broke auth: caused mult versions to come back and conflict with the assignement of [version] type (would require [version[]] to accom both, and then you get to code everything for mult handling)
    * 2:44 PM 3/2/2021 added console TenOrg color support
    # 12:19 PM 5/27/2020 updated cbh, moved alias:dccms win func
    # 1:18 PM 11/7/2018 added Disconnect-PssBroken
    # 12:42 PM 6/20/2018 ported over from disconnect-exo
    .DESCRIPTION
    Disconnect-CCMS - Remove all the existing Security & Compliance connections - as identified via specified -Prefix param (default: cc)- (incl EMOv1/2 PSSessions & EOM3+ nonWinRM - closes anything ConfigurationName: Microsoft.Exchange -AND Name: ExchangeOnlineInternalSession*)

    [Disconnect-ExchangeOnline (ExchangePowerShell) | Microsoft Learn](https://learn.microsoft.com/en-us/powershell/module/exchange/disconnect-exchangeonline?view=exchange-ps)
        Use the Disconnect-ExchangeOnline cmdlet in the Exchange Online PowerShell module to disconnect the connections that you created using the Connect-ExchangeOnline or Connect-IPPSSession cmdlets.

        v3.4.0 examples
        -------------------------- Example 1 --------------------------
        Disconnect-ExchangeOnline
        This example asks for confirmation before disconnecting from Exchange Online PowerShell or Security & Compliance PowerShell.
        -------------------------- Example 2 --------------------------
        Disconnect-ExchangeOnline -Confirm:$false
        This example silently disconnects from Exchange Online PowerShell or Security & Compliance PowerShell without a confirmation prompt or any notification text.
        -------------------------- Example 3 --------------------------
        Disconnect-ExchangeOnline -ConnectionId 1a9e45e8-e7ec-498f-9ac3-0504e987fa85
        This example disconnects the REST-based Exchange Online PowerShell connection with the specified ConnectionId value. Any other remote PowerShell connections to Exchange Online PowerShell or Security & Compliance PowerShell in the same Windows PowerShell
        window are also disconnected.
        -------------------------- Example 4 --------------------------
        Disconnect-ExchangeOnline -ModulePrefix Contoso,Fabrikam
        Updated to match v3.4.0 of ExchangeOnlineMangement: Unlike the  v1.0.1 'disconnect', 
        this also implements new Clear-ActiveToken support, to reset the token as well as the session. 
        Doesn't support targeting session id, just wacks all sessions matching the configurationname & name of an EXOv2 pssession.

        [Get-ConnectionInformation (ExchangePowerShell) | Microsoft Learn](https://learn.microsoft.com/en-us/powershell/module/exchange/get-connectioninformation?view=exchange-ps)

        -------------------------- Example 1 --------------------------
        Get-ConnectionInformation
        This example returns a list of all active REST-based connections with Exchange Online in the current PowerShell instance.
        -------------------------- Example 2 --------------------------
        Get-ConnectionInformation -ConnectionId 1a9e45e8-e7ec-498f-9ac3-0504e987fa85
        This example returns the active REST-based connection with the specified ConnectionId value.
        -------------------------- Example 3 --------------------------
        Get-ConnectionInformation -ModulePrefix Contoso,Fabrikam
        This example returns a list of active REST-based connections that are using the specified prefix values.

    .PARAMETER Prefix
    [verb]-PREFIX[command] PREFIX string for clearly marking cmdlets sourced in this connection [-Prefix tag]
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
    Disconnect-CCMS;
    Disconnect all EXOv2 ConfigurationName: Microsoft.Exchange -AND Name: ExchangeOnlineInternalSession* pssession
    .EXAMPLE
    Disconnect-CCMS -silent;
    Demos use of the silent parameter to suppress output of details
    .LINK
    Github      : https://github.com/tostka/verb-exo
    #>
    [CmdletBinding()]
    [Alias('dccms')]
    Param(
        [Parameter(HelpMessage = "[verb]-PREFIX[command] PREFIX string for clearly marking cmdlets sourced in this connection [-Prefix tag]")]
            [string]$Prefix = 'cc',
        [Parameter(HelpMessage = "MinimumVersion required for ExchangeOnlineManagement module (defaults to '2.0.5')[-MinimumVersion '2.0.6']")]
        [version] $MinimumVersion = '2.0.5',
        [Parameter(HelpMessage = "MinimumVersion required for Non-WinRM connections (of ExchangeOnlineManagement module (defaults to '3.0.0')[-MinimumVersion '2.0.6']")]
        [version] $MinNoWinRMVersion = '3.0.0',
        [Parameter(HelpMessage="Silent output (suppress status echos)[-silent]")]
        [switch] $silent

    ) 
    write-warning "OBSOLETE! shifted all (re|dis)connect-CCMS functions into connect-exo & reconnect-exo: CCMS Sec & Compl connection mgmt is triggered via the -Prefix cc parameter (any other param is assumed to be native EXO; but -Prefix cc will always generate a connection to Sec & Compliance)!" ; 
    BREAK ; 
    $verbose = ($VerbosePreference -eq "Continue") ; 
    $sTitleBarTag = @("CCMS") ;
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
    if(-not (gv EXOv1runspaceConnectionInfoPort -ea 0)){$EXOv1runspaceConnectionInfoPort = '443' };

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
    # add get-connectioninformation.ConnectionURI targeting rgxs for CCMS vs EXO
    if(-not $rgxConnectionUriEXO){$rgxConnectionUriEXO = 'https://outlook\.office365\.com'} ; 
    if(-not $rgxConnectionUriEXO){$rgxConnectionUriCCMS = 'https://ps\.compliance\.protection\.outlook\.com'} ; 
    
    #*------^ END PSS & GMO VARIS ^------

    $pssProps = 'Id','Name','ComputerType','ComputerName','ConfigurationName','State','Availability',
        @{name="TokenExpiryTime";expression={get-date $_.TokenExpiryTime.date -format 'yyyyMMdd-HHmmtt'}};
    

    # it's pulling the verb-EXO vers of disconnect-exchangeonline, force load the v206:
    # admin/SID module auto-install code (myBoxes UID split-perm CU, all else t AllUsers)
    #region EOMREV ; #*------v EOMREV Check v------
    $EOMmodname = 'ExchangeOnlineManagement' ;
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
        #Disconnect-ExchangeOnline -confirm:$false ; 
        # -ModulePrefix
        $pltDXO=[ordered]@{
            ModulePrefix = $Prefix ;
            confirm = $false ; 
            erroraction = 'STOP' ;
            whatif = $($whatif) ;
        } ;
        $smsg = "Disconnect-ExchangeOnline w`n$(($pltDXO|out-string).trim())" ; 
        if($silent){}elseif ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
        Disconnect-ExchangeOnline @pltDXO ; 
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
    
    # here's the 'bug', old legacy pss-based force removal cmd, I'd spliced from v205 into verb-mods
    #Disconnect-PssBroken -verbose:$false ;
    Remove-PSTitlebar $sTitleBarTag #-verbose:$($VerbosePreference -eq "Continue");
    #[console]::ResetColor()  # reset console colorscheme
} ; 

#*------^ Disconnect-CCMS.ps1 ^------