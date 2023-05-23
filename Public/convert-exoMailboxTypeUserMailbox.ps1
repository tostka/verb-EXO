# convert-exoMailboxTypeUserMailbox.ps1

#*------v Function convert-exoMailboxTypeUserMailbox v------
function convert-exoMailboxTypeUserMailbox{
    <#
    .SYNOPSIS
    convert-exoMailboxTypeUserMailbox() - Set specified EXO mailbox to Regular ('UserMailbox')
    .NOTES
    Author: Todd Kadrie
    Website:	http://www.toddomation.com
    Twitter:	@tostka, http://twitter.com/tostka
    REVISIONS   :
    * 8:46 AM 5/17/2023 add to vXO; ren'd to convert-exoMailboxTypeSharedMailbox (rmvd _ internal prefix), and aliased orig name(convert-xoShared), strongly typed $Mailbox as [System.Object] (get-xomailbox returns that type, not a real 'Mailbox' class).
    #1:10 PM 8/25/2021 ren revertExoUserMbx -> _revert-xoUserMbx
    # 10:00 AM 12/19/2018 _revert-xoUserMbx: added post confirm echo
    # 3:19 PM 12/17/2018 coding revert
    # 12:31 PM 10/23/2018 ran full pass live, no unusual errors
    .DESCRIPTION
    convert-exoMailboxTypeUserMailbox() - Set specified EXO mailbox to Regular (from Shared) type (part of coordinated on-prem ADUser recipienttype hack to make it work without movnig mbxs back onprem to convert).
    .PARAMETER  Mailbox, EXO Mailbox Object
    EXO Mailbox Object
    .PARAMETER  Credential
    Credential to use for this connection [-credential 'account@domain.com']
    .PARAMETER silent
    Switch to specify suppression of all but warn/error echos.
    .PARAMETER Whatif
    Parameter to run a Test no-change pass, and log results [-Whatif switch]
    .PARAMETER ShowDebug
    Parameter to display Debugging messages [-ShowDebug switch]
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    Returns RemoteMailbox object, or $false on failure.
    .EXAMPLE
    $ombx = get-xomailbox -id $targUPN -ea stop ;
    $bRet = convert-exoMailboxTypeUserMailbox -Mailbox $ombx -whatif -showDebug ;
    Pull the target cloud-first EXO mailbox, and pass it as an object in to the convert-exoMailboxTypeUserMailbox(), with whatif & showdebug
    .LINK
    https://github.com/tostka/verb-EXO
    #>
    [CmdletBinding()]
    [Alias('revert-xoUserMbx')]
    PARAM(
        [Parameter(Position=0,Mandatory=$True,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,HelpMessage="UPN [-upn fname.lname@DOMAIN.COM]")]
        [ValidateNotNullOrEmpty()]
            [System.Object]$Mailbox,
        [Parameter(Mandatory = $false, HelpMessage = "Use specific Credentials (defaults to Tenant-defined SvcAccount)[-Credentials [credential object]]")]
            [System.Management.Automation.PSCredential]$Credential,
        [Parameter(HelpMessage="Silent output (suppress status echos)[-silent]")]
            [switch] $silent,
        [Parameter(HelpMessage="Debugging Flag [-showDebug]")]
            [switch] $showDebug,
        [Parameter(HelpMessage="Whatif Flag  [-whatIf]")]
            [switch] $whatIf
    ) # PARAM BLOCK END
    BEGIN {
        $exprops="SamAccountName","RecipientType","RecipientTypeDetails","UserPrincipalName" ;
        # recycling the inbound above into next call in the chain
        $pltRXO = [ordered]@{
            Credential = $Credential ;
            verbose = $($VerbosePreference -eq "Continue")  ;
            silent = $silent ;
        } ;
    } ;  # BEGIN-E
    PROCESS {
        foreach($MBX in $Mailbox){
            $Error.Clear() ;
            Reconnect-EXO @pltRXO;

            # usermailbox -> sharedmailbox
            if($MBX |?{$_.recipienttypedetails -eq 'UserMailbox'}){
                $smsg= "PRE:$($MBX.userprincipalname) has already been converted to RemoteMailbox" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Warn } #Error|Warn|Debug
                else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } else {
                <# convert TO SHARED
                $pltSADU=[ordered]@{
                    Identity=$MBX.userprincipalname ;
                    Type="Shared" ;
                    whatif=$($whatif) ;
                } ;
                #>
                # convert FROM SHARED
                $pltSxM=[ordered]@{
                    Identity=$MBX.userprincipalname ;
                    Type="Regular" ;
                    whatif=$($whatif) ;
                    ErrorAction = 'STOP' ;
                } ;
                $smsg="set-xomailbox with:`n$(($pltSxM|out-string).trim())`n" ; ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                # 9:48 AM 12/19/2018 move out here, get the whatif confirm
                $Exit = 0 ;
                Do {
                    Try {
                        set-xomailbox @pltSxM ;
                        $Exit = $DoRetries ;
                        $true | write-output ;
                    } Catch {
                        $smsg = "Failed to exec cmd because: $($Error[0])" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                        else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        Start-Sleep -Seconds $RetrySleep ;
                        $tryNo = 0 ;
                        Reconnect-EXO @pltRXO;
                        $Exit ++ ;
                        $smsg = "Try #: $Exit" ;
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
                    $smsg = "(waiting for get-xoMailbox to return RecipientTypeDetails -eq 'UserMailbox')" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    $1F=$false ;
                    Do {
                        if($1F){Sleep -s 5} ;
                        write-host  "." -NoNewLine ;
                        $1F=$true ;
                        $Exit = 0 ;
                        Do {
                            Try {
                                $pexombx = get-xomailbox -id $pltSxM.identity -ea stop ;
                                $Exit = $DoRetries ;
                            } Catch {
                                $smsg = "Failed to exec cmd because: $($Error[0])" ;
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                                else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                Start-Sleep -Seconds $RetrySleep ;
                                $tryNo = 0 ;
                                Reconnect-exo @pltRXO;
                                $Exit ++ ;
                                $smsg = "Try #: $Exit" ;
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
                    } Until ($pexombx.RecipientTypeDetails -eq 'UserMailbox') ;

                    $smsg= "POST:EXO Mailbox`n$(($pexombx| format-list $exprops|out-string ).trim())" ; ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                } else { 
                    set-xomailbox @pltSxM ;
                    $smsg = "(whatif detected, skipping post test)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }  else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            
                } ;
            } ;  # if-E Not Converted test
        } ;  # loop-E
    } ;  # PROC-E
} ; 
#*------^ END Function convert-exoMailboxTypeUserMailbox ^------