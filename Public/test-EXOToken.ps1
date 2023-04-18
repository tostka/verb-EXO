#*------v test-EXOToken.ps1 v------
function test-EXOToken {
    <#
    .SYNOPSIS
    test-EXOToken - Retrieve and summarize EXOv2 OAuth Active Token (leverages ExchangeOnlineManagement 'Microsoft.Exchange.Management.ExoPowershellGalleryModule.dll', OAuth isn't used for EXO legacy basic-auth connections)
    .NOTES
    Version     : 1.0.0.0
    Author      : Todd Kadrie
    Website     :	http://www.toddomation.com
    Twitter     :	@tostka / http://twitter.com/tostka
    CreatedDate : 2020-08-08
    FileName    : test-EXOToken
    License     : MIT License
    Copyright   : (c) 2020 Todd Kadrie
    Github      : https://github.com/tostka/verb-aad
    Tags        : Powershell,ExchangeOnline,Exchange,RemotePowershell,Connection,MFA
    REVISIONS
    * 2:02 PM 4/17/2023 rev: $MinNoWinRMVersion from 2.0.6 => 3.0.0.
    * 11:02 AM 4/4/2023 reduced the ipmo and vers chk block, removed the lengthy gmo -list; and any autoinstall. Assume EOM is installed, & break if it's not
    * 3:34 PM 3/29/2023 3:14 pm 3/29/2023: REN'D $modname => $EOMModName
    * 10:12 AM 11/28/2022 add: test of get-command -name Test-ActiveToken) before using it (only avail in EOM v205 and less)
    * 3:59 PM 8/2/2022 got through dbugging EOM v205 SID interactive pass, working ; added NoWinRM test and -MinNoWinRMVersion, to bypass attempts with this, post EOM v205 (as v206 completely drops the dependant test|clear-ActiveToken())
    # 10:25 AM 8/2/2022 NOPE! get-msaltoken *authenticates* a fresh connection, like Connect-EOM, 
    - _if_ you spec the PS EXO client guid, as the 
     - so it can sort of 'act' like a token validator, but as it fully authenticates if there's no token, it's *not equive to test-ActiveToken* 
     from ExchangeOnlineManagement v205. and *it completely lacks -Prefix support*! Can't use this for hybrid prefix-tagged coexist with onprem Exch ps sessions.
     fundemental break there. 
    - So I coded out any use of this from connect-exo2/reconnect-exo2, for NoWinRM EOM v206p6+ use.
    * 3:30 PM 7/25/2022 tests against CBA & SID interactive creds on EOM v205, need to debug now against EOM v206p6, to accomodate PSSession-less connect & test code.
    * 8:30 AM 7/11/2022 rewrite for a post-PSS basicauth, EOM v206p6 world: can't use EOM:get-activetoken; add $Credential (needed for msal.ps:get-msaltoken()
    * 4:08 PM 3/29/2022 updated, got test-Activetoken call into gallery mod working again (token time is pulled off of the active setsion and looks like it id's which token /sesion you're closing to remove-Activetoken etc). 
    * 3:41 PM 3/28/2022 update for v2.05, supporting .netcore & .netframework subdirs in the ExchangeOnlineManagement module, failing on trailing test-activetoken code - wants -TokenExpiryTime now.
    * 11:40 AM 5/14/2021 added -ea 0 to the gv tests (suppresses not-found error when called without logging config)
    # 8:34 AM 3/31/2021 added verbose suppress to all import-mods
    * 12:21 PM 8/11/2020 added dependancy mod try/tach, and a catch on the failure error returned by the underlying test-ActiveToken cmd
    * 11:58 AM 8/9/2020 init
    .DESCRIPTION
    test-EXOToken - Retrieve and summarize EXOv2 OAuth Active Token (leverages ExchangeOnlineManagement 'Microsoft.Exchange.Management.ExoPowershellGalleryModule.dll', OAuth isn't used for EXO legacy basic-auth connections)
    Trying to find a way to verify status of token, wo any interactive EXO traffic. Lifted concept from EXOM UpdateImplicitRemotingHandler().
    .PARAMETER Credential
    Credential to be used for connection
     .PARAMETER MinNoWinRMVersion
    MinimumVersion required for Non-WinRM connections (of ExchangeOnlineManagement module (defaults to '3.0.0')[-MinimumVersion '2.0.6']
    Test-ActiveToken doesn't appear to normally be exposed anywhere but with explicit load of the .dll
    .OUTPUT
    System.Boolean
    .EXAMPLE
    $hasActiveToken = test-EXOToken 
    $psss=Get-PSSession | where-object {$_.ConfigurationName -like "Microsoft.Exchange" -AND $_.Name -like "ExchangeOnlineInternalSession*" } ;  
    $sessionIsOpened = $psss.Runspace.RunspaceStateInfo.State -eq 'Opened'
    if (($hasActiveToken -eq $false) -or ($sessionIsOpened -ne $true)){
        #If there is no active user token or opened session then ensure that we remove the old session
        $shouldRemoveCurrentSession = $true;
    } ; 
    Retrieve and evaluate status of EXO user token against PSSessoin status for EXOv2
    .LINK
    https://github.com/tostka/verb-aad
    #>
    #Requires -Modules MSAL.PS,ExchangeOnlineManagement
    [CmdletBinding()] 
    Param(
        [Parameter(Mandatory=$True,HelpMessage="Credentials [-Credentials [credential object]]")]
        [System.Management.Automation.PSCredential]$Credential = $global:credo365TORSID,
        [Parameter(HelpMessage = "MinimumVersion required for Non-WinRM connections (of ExchangeOnlineManagement module (defaults to '3.0.0')[-MinimumVersion '2.0.6']")]
        [version] $MinNoWinRMVersion = '3.0.0'
    ) ;
    BEGIN {
        $verbose = ($VerbosePreference -eq "Continue") ;
        if(-not (gv rgxCertThumbprint -ea 0)){$rgxCertThumbprint = '[0-9a-fA-F]{40}' ; } ;

    } ;
    PROCESS {
        $hasActiveToken = $false ; 
        # Save time and pretest for *any* EXOv2 PSSession, before bothering to test (no session - even closed/broken => no OAuth token)
        # w EOM v206p5, there's no longer even a PSS to detect at all, so this loses function as well.
        # 8:52 AM 7/11/2022 this ^ is equiv to EOM code: $existingPSSession = Get-PSSession | Where-Object {$_.ConfigurationName -like "Microsoft.Exchange" -and $_.Name -like "ExchangeOnlineInternalSession*"}
        
        # * 11:02 AM 4/4/2023 reduced the ipmo and vers chk block, removed the lengthy gmo -list; and any autoinstall. Assume EOM is installed, & break if it's not
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

        if($UseConnEXO){
            $smsg = "$($EOMmodname) v$($MinNoWinRMVersion)+ detected: No dependancy test-ActiveToken() available in later EOM builds" ; 
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 

        } elseif ($exov2 = Get-PSSession | where-object {$_.ConfigurationName -like "Microsoft.Exchange" -AND $_.Name -like "ExchangeOnlineInternalSession*"}){

            $error.clear() ;
            TRY {
                #=load function module (subcomponent of dep module, pathed from same dir)
                #$tmodpath = join-path -path (split-path (get-module $EOMmodname -list).path) -ChildPath 'Microsoft.Exchange.Management.ExoPowershellGalleryModule.dll' ;
                $EOMgmtModulePath = split-path (get-module $EOMmodname -list).Path ; 
                if($IsCoreCLR){
	                $EOMgmtModulePath = resolve-path -Path $EOMgmtModulePath\netcore ;
	                $smsg = "(.netcore path in use:" ; 
                } else {
	                $EOMgmtModulePath = resolve-path -Path $EOMgmtModulePath\netFramework
	                $smsg = "(.netnetFramework path in use:" ;
                } ;
                $smsg += "$($EOMgmtModulePath))" ;
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                $tmodpath = join-path -path $EOMgmtModulePath -ChildPath 'Microsoft.Exchange.Management.ExoPowershellGalleryModule.dll' ;

                if(test-path $tmodpath){ import-module -name $tmodpath -Cmdlet Test-ActiveToken -verbose:$false }
                else { throw "Unable to locate:Microsoft.Exchange.Management.ExoPowershellGalleryModule.dll" } ;  
            } CATCH {
                $ErrTrapd=$Error[0] ;
                $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
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
        
            # targeted load example: # Import the module once more to ensure that Test-ActiveToken is present
            # Import-Module $global:_EXO_ModulePath -Cmdlet Test-ActiveToken;
            # ipmo $EOMgmtModulePath -Cmdlet Test-ActiveToken;
            if(get-command -name Test-ActiveToken){
                $error.clear() ;
                TRY {
                    #$hasActiveToken = Test-ActiveToken ; 
                    # updated from EOM v2.0.5
                    # grab the target session on it's settings:
                    $existingPSSession = Get-PSSession | Where-Object {$_.ConfigurationName -like "Microsoft.Exchange" -and $_.Name -like "ExchangeOnlineInternalSession*"}
                    if ($existingPSSession.count -gt 0) {
                        for ($index = 0; $index -lt $existingPSSession.count; $index++){
                            $session = $existingPSSession[$index]
                            $hasActiveToken = Test-ActiveToken -TokenExpiryTime $session.TokenExpiryTime ; 
                        } ; 
                    } ; 
                } CATCH [System.Management.Automation.RuntimeException] {
                    # reflects: test-activetoken : Object reference not set to an instance of an object.
                    write-verbose "Token not present"
                } CATCH {
                    $ErrTrapd=$Error[0] ;
                    $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
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
            } else { throw "missing:gcm -name Test-ActiveToken" } 

            
        } else {
            $smsg = "Neither NoWinRM (EOM v206+) or existing EXOv2 (v205) PSSession found to confirm!" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } 
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
        } ;  

} # E-PROC
    
    END{ $hasActiveToken | write-output } ;
}

#*------^ test-EXOToken.ps1 ^------