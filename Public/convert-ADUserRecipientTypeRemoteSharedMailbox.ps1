# convert-ADUserRecipientTypeRemoteSharedMailbox.ps1

#*------v Function convert-ADUserRecipientTypeRemoteSharedMailbox v------    
function convert-ADUserRecipientTypeRemoteSharedMailbox{
    <#
    .SYNOPSIS
    Convert the passed-in ADUser object RecipientType to RemoteSharedMailbox (sets ADUser.msExchRecipientTypeDetails:'34359738368' ADUUser.msExchRemoteRecipientType:'100').
    1st half of RemoteUserMailbox conversion to ReomteSharedMailbox (followed by convert-exoMailboxTypeSharedMailbox to convert the EXO mailbox type).
    Traditionally this would be applied to RemoteUserMailbox, but no pre-checking is performed, the update is applied as long as the target settings aren't already in place.
    .NOTES
    Author: Todd Kadrie
    Website:	http://www.toddomation.com
    Twitter:	@tostka, http://twitter.com/tostka
    REVISIONS   :
    * 5:22 PM 1/25/2025 dc access issues to BCC, had to add retry code & get-gcfast -exclude to autorecover on ad calls
    * 10:12 AM 5/18/2023 add: adv func beg/proc blocks: full pipeline support ; ren'd mid proc block vari added explicit loop
    * 8:46 AM 5/17/2023 add to vXO; ren'd to convert-ADUserRecipientTypeRemoteSharedMailbox (rmvd _ internal prefix), and aliased orig name(convert-ADUserRecipientType), strongly typed $ADUser as [Microsoft.ActiveDirectory.Management.ADUser]
    * 1:06 PM 8/25/2021 ren convertADUser -> convert-ADUserRecipientType
    # 2:51 PM 12/18/2018 set-adus has functional whatif, moved it into test fire
    # 12:31 PM 10/23/2018 ran full pass live, no unusual errors
    .DESCRIPTION
    Convert the passed-in ADUser object RecipientType to RemoteSharedMailbox (sets ADUser.msExchRecipientTypeDetails:'34359738368' ADUUser.msExchRemoteRecipientType:'100').
    1st half of RemoteUserMailbox conversion to ReomteSharedMailbox (followed by convert-exoMailboxTypeSharedMailbox to convert the EXO mailbox type).
    Traditionally this would be applied to RemoteUserMailbox, but no pre-checking is performed, the update is applied as long as the target settings aren't already in place.

    (does not require passed in Credentials, as all changes are with ActiveDirectory module, which does not support affirmative logon; logon is pickedup from the psdrive AD mapping) 

    .PARAMETER  ADUser
    ADUser object [-ADUser `$ADVariable]
    .PARAMETER domaincontroller
    Option to hardcode a specific DC [-domaincontroller xxxx]
    .PARAMETER Whatif
    Parameter to run a Test no-change pass, and log results [-Whatif switch]
    .PARAMETER ShowDebug
    Parameter to display Debugging messages [-ShowDebug switch]
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    Returns RemoteMailbox object, or $false on failure.
    .EXAMPLE
    $adu=get-aduser -id $rmbx.DistinguishedName -server $domainController -Properties $adprops -ea 0| select $adprops ;
    $bRet=convert-ADUserRecipientType -ADUser $adu -whatif:$($whatif) -showDebug:$($showdebug) ;
    Convert the passed-in ADUser object RecipientType from RemoteUserMailbox to RemoteSharedMailbox.
    (does not require passed in Credentials, as all changes are with ActiveDirectory module, which does not support affirmative logon; logon is pickedup from the psdrive AD mapping automounted on ActieDirectory module load) 
    .LINK
    https://github.com/tostka/verb-EXO
    #>
    [CmdletBinding()]
    [Alias('convert-ADUserRecipientType')]
    PARAM(
        [Parameter(Position=0,Mandatory=$True,ValueFromPipeline=$true,HelpMessage="ADUser object [-ADUser `$ADVariable]")]
            [ValidateNotNullOrEmpty()]
            [Microsoft.ActiveDirectory.Management.ADUser]$ADUser,
        [Parameter(HelpMessage="Option to hardcode a specific DC [-domaincontroller xxxx]")]
            [string]$domaincontroller,
        [Parameter(HelpMessage="Debugging Flag [-showDebug]")]
            [switch] $showDebug,
        [Parameter(HelpMessage="Whatif Flag  [-whatIf]")]
            [switch] $whatIf
    ) # PARAM BLOCK END
    BEGIN {
        if(-not $DoRetries){$DoRetries = 4 } ;    # # times to repeat retry attempts
        if(-not $RetrySleep){$RetrySleep = 10 } ; # wait time between retries
        if(-not $RetrySleep){$DawdleWait = 30 } ; # wait time (secs) between dawdle checks
        if(-not $DirSyncInterval){$DirSyncInterval = 30 } ; # AADConnect dirsync interval

    } ;  # BEGIN-E
    PROCESS {
            foreach($ADU in $ADUser) {
            $error.clear() ;
            $Exit = 0 ;
            Do {
                Try {
                    if(!$domaincontroller){$domaincontroller=get-gcfast} ;
                    $adprops="samaccountname","msExchRemoteRecipientType","msExchRecipientDisplayType","msExchRecipientTypeDetails","UserPrincipalName","DistinguishedName" ;
                    $exprops="SamAccountName","RecipientType","RecipientTypeDetails","UserPrincipalName" ;
                    if($ADU){
                        $smsg= "PRE:ADUser`n$(($ADU| format-list $adprops|out-string ).trim())" ; ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                        $smsg="PRE:Ex Remotemailbox`n$((Get-RemoteMailbox $ADU.userprincipalname -domaincontroller $domaincontroller| format-list $exprops|out-string ).trim())" ;  ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                        if($ADU.msExchRecipientTypeDetails -eq '34359738368' -and $ADU.msExchRemoteRecipientType -eq '100'){
                            $smsg= "PRE:$($ADU.userprincipalname) has already been converted to RemoteSharedMailbox" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Warn } #Error|Warn|Debug
                            else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        } else {
                            $pltSADU=[ordered]@{
                                Identity=$ADU.samaccountname ;
                                Replace=@{msExchRemoteRecipientType=100;msExchRecipientTypeDetails=34359738368}  ;
                                server=$domaincontroller ;
                                whatif=$($whatif) ;
                            } ;
                            $smsg= "Set-aduser with:`n$(($pltSADU|out-string).trim())`n" ;
                            #expand replace values
                            $smsg += "`n$(($pltsadu.replace | fl|out-string).trim())`n" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Warn } #Error|Warn|Debug
                            else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                            $Exit = 0 ;
                            Do {
                                Try {
                                    Set-ADUser @pltSADU ;
                                    $Exit = $DoRetries ;
                                } catch [Microsoft.ActiveDirectory.Management.ADServerDownException] {
                                    $ErrorTrapped=$Error[0] ;
                                    $smsg = "Failed to exec cmd because: $($ErrorTrapped.Exception.Message )" ;
                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                    # re-quire a new DC
                                    $badDC = $domaincontroller ;
                                    $smsg = "PROBLEM CONTACTING $($badDC)!:Resource unavailable: $($ErrorTrapped.Exception.Message)" ;
                                    $smsg += "get-GCFast() an alterate DC (-exclude $($badDC))" ;
                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                    if($domaincontroller = get-gcfast -Exclude $badDC -Verbose){
                                        $script:domaincontroller = $domaincontroller  ;
                                        write-warning "Changing DomainController:($($badDC)->$($domaincontroller)):Waiting 20seconds, for ReplSync..." ;
                                        start-sleep -Seconds 20 ;
                                        Continue ;
                                    } else {
                                        $smsg = "UNABLE TO ACQUIRE A NEW ALTERNATE DC! ABORTING" ;
                                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                        THROW $SMSG ;
                                    } ;  ;
                                } Catch {
                                    $ErrorTrapped=$Error[0] ;
                                    $smsg = "FULL ERROR TRAPPED (EXPLICIT CATCH BLOCK WOULD LOOK LIKE): } CATCH[$($ErrorTrapped.Exception.GetType().FullName)]{" ;
                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug
                                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                    #Break #Opts: STOP(debug)|EXIT(close)|CONTINUE(move on in loop cycle)|BREAK(exit loop iteration)|THROW $_/'CustomMsg'(end script with Err output)
                                    Start-Sleep -Seconds $RetrySleep ;
                                    $Exit ++ ;
                                    $smsg = "Failed to exec cmd because: $($ErrorTrapped)" ;
                                    $smsg += "`nTry #: $Exit" ;
                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                                    else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                    If ($Exit -eq $DoRetries) {
                                        $smsg =  "Unable to exec cmd!" ;
                                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                                        else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                    } ;
                                    Continue ;
                                }  ;
                            } Until ($Exit -eq $DoRetries) ;
                            if(!$whatif){
                                $smsg= "POST:ADUser`n$((Get-adUser -id $ADU.samaccountname -prop $adprops -server $domaincontroller|fl $adprops | out-string).trim())`n" ;;
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                $smsg= "POST:Ex Remotemailbox`n$((Get-RemoteMailbox $ADU.userprincipalname -domaincontroller $domaincontroller| format-list $exprops|out-string ).trim())" ; ;
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            } else {
                                write-host -fore yellow "(whatif detected, skipping update)"
                            } ;
                        } ;
                        $true | write-output ;
                    } else {
                        $smsg="`n:`$tEmlAddr:$($tEmlAddr): not matched against ADUser`n" ;  ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Warn } #Error|Warn|Debug
                        else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ;
                    $Exit = $DoRetries ;
                } Catch {
                    Start-Sleep -Seconds $RetrySleep ;
                    $Exit ++ ;
                    $smsg = "Failed to exec cmd because: $($Error[0])" ;
                    $smsg += "`nTry #: $Exit" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                    else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    If ($Exit -eq $DoRetries) {
                        $smsg =  "Unable to exec cmd!" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                        else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ;
                    Continue ;
                }  ;
            } Until ($Exit -eq $DoRetries) ;
        } ;  # loop-E
    } ;  # PROC-E
} ; 
#*------^ END Function convert-ADUserRecipientTypeRemoteSharedMailbox ^------