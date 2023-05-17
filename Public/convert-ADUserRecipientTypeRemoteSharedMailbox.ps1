# convert-ADUserRecipientTypeRemoteSharedMailbox.ps1

function convert-ADUserRecipientTypeRemoteSharedMailbox{
    <#
    .SYNOPSIS
    Convert the passed-in ADUser object RecipientType to RemoteSharedMailbox (sets ADUser.msExchRecipientTypeDetails:'34359738368' ADUUser.msExchRemoteRecipientType:'100').
    Traditionally this would be applied to RemoteUserMailbox, but no pre-checking is performed, the update is applied as long as the target settings aren't already in place.
    .NOTES
    Author: Todd Kadrie
    Website:	http://www.toddomation.com
    Twitter:	@tostka, http://twitter.com/tostka
    REVISIONS   :
    * 8:46 AM 5/17/2023 add to vXO; ren'd to convert-ADUserRecipientTypeRemoteSharedMailbox (rmvd _ internal prefix), and aliased orig name(convert-ADUserRecipientType), strongly typed $ADUser as [Microsoft.ActiveDirectory.Management.ADUser]
    * 1:06 PM 8/25/2021 ren convertADUser -> convert-ADUserRecipientType
    # 2:51 PM 12/18/2018 set-adus has functional whatif, moved it into test fire
    # 12:31 PM 10/23/2018 ran full pass live, no unusual errors
    .DESCRIPTION
    Convert the passed-in ADUser object RecipientType to RemoteSharedMailbox (sets ADUser.msExchRecipientTypeDetails:'34359738368' ADUUser.msExchRemoteRecipientType:'100').
    Traditionally this would be applied to RemoteUserMailbox, but no pre-checking is performed, the update is applied as long as the target settings aren't already in place.
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
    $bRet=convert-ADUserRecipientType -ADUser $adu -whatif:$($whatif) -showDebug:$($showdebug) ;
    Convert the passed-in ADUser object RecipientType from RemoteUserMailbox to RemoteSharedMailbox.
    .LINK
    https://github.com/tostka/verb-EXO
    #>
    [CmdletBinding()]
    [Alias('convert-ADUserRecipientType')]
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

                # 2:39 PM 10/11/2018 add pretest of rtypes
                # if($aduser.msExchRecipientTypeDetails -eq '2147483648' -and $aduser.msExchRemoteRecipientType -eq '4'){"Y"}
                if($aduser.msExchRecipientTypeDetails -eq '34359738368' -and $aduser.msExchRemoteRecipientType -eq '100'){
                    $smsg= "PRE:$($ADUser.userprincipalname) has already been converted to RemoteSharedMailbox" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Warn } #Error|Warn|Debug
                    else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                } else {
                    $splt=[ordered]@{
                        Identity=$ADUser.samaccountname ;
                        Replace=@{msExchRemoteRecipientType=100;msExchRecipientTypeDetails=34359738368}  ;
                        server=$domaincontroller ;
                        whatif=$($whatif) ;
                    } ;
                    # whatif=$($whatif) ;
                    write-host -fore green "Set-aduser with:`n$(($splt|out-string).trim())`n" ;
                    #2:49 PM 12/18/2018 set-aduser has FUNCTIONAL -whatif!
                    # 7:35 PM 10/11/2018 force up connection
                    #connect-msol @pltRXO;
                    #Connect-AAD @pltRXO;

                    Set-ADUser @splt ;
                    if(!$whatif){
                        $smsg= "POST:ADUser`n$((Get-adUser -id $ADUser.samaccountname -prop $adprops -server $domaincontroller|fl $adprops | out-string).trim())`n" ;;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        $smsg= "POST:Ex Remotemailbox`n$((Get-RemoteMailbox $ADUser.userprincipalname -domaincontroller $domaincontroller| format-list $exprops|out-string ).trim())" ; ;
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

} ; #*------^ END Function convert-ADUserRecipientTypeRemoteSharedMailbox ^------