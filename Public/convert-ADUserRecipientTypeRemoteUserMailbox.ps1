# convert-ADUserRecipientTypeRemoteUserMailbox

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
    .LINK
    #>
    [CmdletBinding()]
    [Alias('revert-ADuserRecipientType')]
    PARAM(
        [Parameter(Position=0,Mandatory=$True,ValueFromPipeline=$true,HelpMessage="ADUser object [-ADUser `$ADVariable]")]
        [ValidateNotNullOrEmpty()]
        [Microsoft.ActiveDirectory.Management.ADUser]$ADUser,
        [switch] $showDebug,
        [Parameter(HelpMessage="Whatif Flag  [-whatIf]")]
        [switch] $whatIf
    ) # PARAM BLOCK END

    $error.clear() ;
    $Exit = 0 ;
    Do {
        Try {
            if(!$domaincontroller){$domaincontroller=get-gcfast} ;
            $adprops="samaccountname","msExchRemoteRecipientType","msExchRecipientDisplayType","msExchRecipientTypeDetails","UserPrincipalName","DistinguishedName" ;
            $exprops="SamAccountName","RecipientType","RecipientTypeDetails","UserPrincipalName" ;
            if($ADUser){
                $smsg= "PRE:ADUser`n$(($ADUser| format-list $adprops|out-string ).trim())" ; ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                $smsg="PRE:Ex Remotemailbox`n$((Get-RemoteMailbox $ADUser.userprincipalname -domaincontroller $domaincontroller| format-list $exprops|out-string ).trim())" ;  ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                <# stock UserMailbox specs:
                14:15:02:PRE:ADUser
                samaccountname             : lynctest9
                msExchRemoteRecipientType  : 4
                msExchRecipientDisplayType : -2147483642
                msExchRecipientTypeDetails : 2147483648
                14:15:02:PRE:Ex Remotemailbox
                SamAccountName       : lynctest9
                RecipientType        : MailUser
                RecipientTypeDetails : RemoteUserMailbox

                14:15:02:Set-aduser with:
                Name                           Value
                ----                           -----
                Identity                       lynctest9
                Replace                        {msExchRecipientTypeDetails, msExchRemoteRecipientType}
                server                         LYNMS812
                whatif                         False
                14:15:02:POST:ADUser
                samaccountname             : lynctest9
                msExchRemoteRecipientType  : 100
                msExchRecipientDisplayType : -2147483642
                msExchRecipientTypeDetails : 34359738368
                14:15:02:POST:Ex Remotemailbox
                SamAccountName       : lynctest9
                RecipientType        : MailUser
                RecipientTypeDetails : RemoteSharedMailbox
                #>
                # 2:39 PM 10/11/2018 add pretest of rtypes
                # if($aduser.msExchRecipientTypeDetails -eq '2147483648' -and $aduser.msExchRemoteRecipientType -eq '4'){"Y"}
                # convert remoteuser -> remoteshared
                #if($aduser.msExchRecipientTypeDetails -eq '34359738368' -and $aduser.msExchRemoteRecipientType -eq '100'){
                # convert remoteshared back to remote user:
                if($aduser.msExchRecipientTypeDetails -eq '2147483648' -and $aduser.msExchRemoteRecipientType -eq '4'){
                    #$smsg= "PRE:$($ADUser.userprincipalname) has already been converted to RemoteSharedMailbox" ;
                    $smsg= "PRE:$($ADUser.userprincipalname) has already been converted to RemoteUserMailbox" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Warn } #Error|Warn|Debug
                    else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                } else {
                    <# remoteusermbx -> remoteshared
                    $splt=[ordered]@{
                        Identity=$ADUser.samaccountname ;
                        Replace=@{msExchRemoteRecipientType=100;msExchRecipientTypeDetails=34359738368}  ;
                        server=$domaincontroller ;
                        whatif=$($whatif) ;
                    } ;
                    #>
                    # remoteshared -> remoteuser
                    $splt=[ordered]@{
                        Identity=$ADUser.samaccountname ;
                        Replace=@{msExchRemoteRecipientType=4;msExchRecipientTypeDetails=2147483648}  ;
                        server=$domaincontroller ;
                        #ErrorAction = 'STOP'
                        whatif=$($whatif) ;
                    } ;
                    # whatif=$($whatif) ;
                    #write-host -fore green "Set-aduser with:`n$(($splt|out-string).trim())`n" ;
                    $smsg = "Set-aduser with:`n$(($splt|out-string).trim())`n" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                    #connect-msol @pltRXO;
                    #Connect-AAD @pltRXO;

                    Set-ADUser @splt ;
                    if(!$whatif){
                        # 7:35 PM 10/11/2018 force up connection
                        $smsg= "POST:ADUser`n$((Get-adUser -id $ADUser.samaccountname -prop $adprops -server $domaincontroller|fl $adprops | out-string).trim())`n" ;;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        $smsg= "POST:Ex Remotemailbox`n$((Get-RemoteMailbox $ADUser.userprincipalname -domaincontroller $domaincontroller| format-list $exprops|out-string ).trim())" ; ;
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

} ; 
#*------^ END Function convert-ADUserRecipientTypeRemoteUserMailbox ^------