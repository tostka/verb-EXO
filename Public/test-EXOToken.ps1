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
    * 4:08 PM 3/29/2022 updated, got test-Activetoken call into gallery mod working again (token time is pulled off of the active setsion and looks like it id's which token /sesion you're closing to remove-Activetoken etc). 
    * 3:41 PM 3/28/2022 update for v2.05, supporting .netcore & .netframework subdirs in the ExchangeOnlineManagement module, failing on trailing test-activetoken code - wants -TokenExpiryTime now.
    * 11:40 AM 5/14/2021 added -ea 0 to the gv tests (suppresses not-found error when called without logging config)
    # 8:34 AM 3/31/2021 added verbose suppress to all import-mods
    * 12:21 PM 8/11/2020 added dependancy mod try/tach, and a catch on the failure error returned by the underlying test-ActiveToken cmd
    * 11:58 AM 8/9/2020 init
    .DESCRIPTION
    test-EXOToken - Retrieve and summarize EXOv2 OAuth Active Token (leverages ExchangeOnlineManagement 'Microsoft.Exchange.Management.ExoPowershellGalleryModule.dll', OAuth isn't used for EXO legacy basic-auth connections)
    Trying to find a way to verify status of token, wo any interactive EXO traffic. Lifted concept from EXOM UpdateImplicitRemotingHandler().
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
    #Requires -Modules ExchangeOnlineManagement
    [CmdletBinding()] 
    Param() ;
    BEGIN {$verbose = ($VerbosePreference -eq "Continue") } ;
    PROCESS {
        $hasActiveToken = $false ; 
        # Save time and pretest for *any* EXOv2 PSSession, before bothering to test (no session - even closed/broken => no OAuth token)
        $exov2 = Get-PSSession | where-object {$_.ConfigurationName -like "Microsoft.Exchange" -AND $_.Name -like "ExchangeOnlineInternalSession*"} ; 
        if($exov2){
        
            # ==load dependancy module:
            # admin/SID module auto-install code (myBoxes UID split-perm CU, all else t AllUsers)
            $modname = 'ExchangeOnlineManagement' ;
            $minvers = '2.0.5' ; 
            Try {Get-Module -name $modname -listavailable -ErrorAction Stop | out-null } Catch {
                $pltInMod=[ordered]@{Name=$modname} ; 
                if( $env:COMPUTERNAME -match $rgxMyBoxUID ){$pltInMod.add('scope','CurrentUser')} else {$pltInMod.add('scope','AllUsers')} ;
                write-host -foregroundcolor YELLOW "$((get-date).ToString('HH:mm:ss')):Install-Module w scope:$($pltInMod.scope)`n$(($pltInMod|out-string).trim())" ; 
                Install-Module @pltIMod ; 
            } ; # IsInstalled
            $pltIMod = @{Name = $modname ; ErrorAction = 'Stop' ; verbose=$false } ;
            if($minvers){$pltIMod.add('MinimumVersion',$minvers) } ; 
            Try { Get-Module $modname -ErrorAction Stop | out-null } Catch {
                write-verbose "Import-Module w`n$(($pltIMod|out-string).trim())" ; 
                Import-Module @pltIMod ; 
            } ; # IsImported

            $error.clear() ;
            TRY {
                #=load function module (subcomponent of dep module, pathed from same dir)
                #$tmodpath = join-path -path (split-path (get-module $modname -list).path) -ChildPath 'Microsoft.Exchange.Management.ExoPowershellGalleryModule.dll' ;
                $EOMgmtModulePath = split-path (get-module $modname -list).Path ; 
                if($IsCoreCLR){
                    $EOMgmtModulePath = resolve-path -Path $EOMgmtModulePath\netcore ;
                    $smsg = "(.netcore path in use:" ; 
                } else { 
                    $EOMgmtModulePath = resolve-path -Path $EOMgmtModulePath\netFramework
                    $smsg = "(.netnetFramework path in use:" ;                 
                } ; 
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
            if(gcm -name Test-ActiveToken){
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
            write-verbose "No Token: No existing EXOv2 PSSession (ConfigurationName -like 'Microsoft.Exchange' -AND Name -like 'ExchangeOnlineInternalSession*')" ; 
        } ; 
    } ; 
    END{ $hasActiveToken | write-output } ;
}

#*------^ test-EXOToken.ps1 ^------
