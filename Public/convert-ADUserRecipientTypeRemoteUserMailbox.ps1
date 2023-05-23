# convert-ADUserRecipientTypeRemoteUserMailbox.ps1

#*------v Function convert-ADUserRecipientTypeRemoteUserMailbox v------
function convert-ADUserRecipientTypeRemoteUserMailbox{
    <#
    .SYNOPSIS
    Convert the passed-in ADUser object RecipientType to RemoteUserMailbox (Sets ADUser.msExchRecipientTypeDetails:'2147483648' & ADUser.msExchRemoteRecipientType:'4')
    Traditionally this would be applied to RemoteSharedMailbox, but no pre-checking is performed, the update is applied as long as the target settings aren't already in place.
    .NOTES
    Author: Todd Kadrie
    Website:	http://www.toddomation.com
    Twitter:	@tostka, http://twitter.com/tostka
    REVISIONS   :
    * 8:46 AM 5/17/2023 add to vXO; ren'd back to convert-ADUserRecipientTypeRemoteUserMailbox (rmvd _ internal prefix), and aliased orig name(revert-ADuserRecipientType);
            strongly typed $ADUser as [Microsoft.ActiveDirectory.Management.ADUser]; updated CBH ; 
    1:08 PM 8/25/2021 ren revertADuser -> _revert-ADuserRecipientType
    # 2:51 PM 12/18/2018 set-adus has functional whatif, moved it into test fire
    # 1:34 PM 12/17/2018 initi vers
    .DESCRIPTION
    Convert the passed-in ADUser object RecipientType to RemoteUserMailbox (Sets ADUser.msExchRecipientTypeDetails:'2147483648' & ADUser.msExchRemoteRecipientType:'4')
    Traditionally this would be applied to RemoteSharedMailbox, but no pre-checking is performed, the update is applied as long as the target settings aren't already in place.

    (does not require passed in Credentials, as all changes are with ActiveDirectory module, which does not support affirmative logon; logon is pickedup from the psdrive AD mapping) 

    .PARAMETER  ADUser
    ADUser object [-ADUser `$ADVariable]
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
    $bRet=convert-ADUserRecipientTypeRemoteUserMailbox -ADUser $adu -whatif:$($whatif) -showDebug:$($showdebug) ;
    Convert the passed-in ADUser object RecipientType from RemoteUserMailbox to RemoteSharedMailbox.
    (does not require passed in Credentials, as all changes are with ActiveDirectory module, which does not support affirmative logon; logon is pickedup from the psdrive AD mapping automounted on ActieDirectory module load) 
    .LINK
    #>
    [CmdletBinding()]
    [Alias('revert-ADuserRecipientType')]
    PARAM(
        [Parameter(Position=0,Mandatory=$True,ValueFromPipeline=$true,HelpMessage="ADUser object [-ADUser `$ADVariable]")]
        [ValidateNotNullOrEmpty()]
        [Microsoft.ActiveDirectory.Management.ADUser]$ADUser,
        [Parameter(HelpMessage="Debugging Flag [-showDebug]")]
            [switch] $showDebug,
        [Parameter(HelpMessage="Whatif Flag  [-whatIf]")]
            [switch] $whatIf
    ) # PARAM BLOCK END
    BEGIN {} ;  # BEGIN-E
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

                        # convert remoteshared back to remote user:
                        if($ADU.msExchRecipientTypeDetails -eq '2147483648' -and $ADU.msExchRemoteRecipientType -eq '4'){
                            #$smsg= "PRE:$($ADU.userprincipalname) has already been converted to RemoteSharedMailbox" ;
                            $smsg= "PRE:$($ADU.userprincipalname) has already been converted to RemoteUserMailbox" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Warn } #Error|Warn|Debug
                            else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        } else {
                            <# remoteusermbx -> remoteshared
                            $pltSADU=[ordered]@{
                                Identity=$ADU.samaccountname ;
                                Replace=@{msExchRemoteRecipientType=100;msExchRecipientTypeDetails=34359738368}  ;
                                server=$domaincontroller ;
                                whatif=$($whatif) ;
                            } ;
                            #>
                            # remoteshared -> remoteuser
                            $pltSADU=[ordered]@{
                                Identity=$ADU.samaccountname ;
                                Replace=@{msExchRemoteRecipientType=4;msExchRecipientTypeDetails=2147483648}  ;
                                server=$domaincontroller ;
                                #ErrorAction = 'STOP'
                                whatif=$($whatif) ;
                            } ;
                            # whatif=$($whatif) ;

                            $smsg = "Set-aduser with:`n$(($pltSADU|out-string).trim())`n" ;
                            #expand replace values
                            $smsg += "`n$(($pltsadu.replace | fl|out-string).trim())`n" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                            Set-ADUser @pltSADU ;
                            if(!$whatif){
                                $smsg= "POST:ADUser`n$((Get-adUser -id $ADU.samaccountname -prop $adprops -server $domaincontroller|fl $adprops | out-string).trim())`n" ;;
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                $smsg= "POST:Ex Remotemailbox`n$((Get-RemoteMailbox $ADU.userprincipalname -domaincontroller $domaincontroller| format-list $exprops|out-string ).trim())" ; ;
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            } else { write-host -fore yellow "(whatif detected, skipping update)"} ;
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
#*------^ END Function convert-ADUserRecipientTypeRemoteUserMailbox ^------