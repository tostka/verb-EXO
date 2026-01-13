# convert-exoMailboxTypeSharedMailbox 

#*------v Function convert-exoMailboxTypeSharedMailbox v------
function convert-exoMailboxTypeSharedMailbox{
    <#
    .SYNOPSIS
    convert-exoMailboxTypeSharedMailbox() - Set specified EXO mailbox to Shared ('SharedMailbox) type (part of coordinated on-prem ADUser recipienttype hack to make it work without moving mbxs back onprem to convert).
    .NOTES
    Author: Todd Kadrie
    Website:	http://www.toddomation.com
    Twitter:	@tostka, http://twitter.com/tostka
    REVISIONS   :
    * 10:38 AM 1/13/2026 ADD: -ea stop  to try splats
    * 8:46 AM 5/17/2023 add to vXO; ren'd to convert-exoMailboxTypeSharedMailbox (rmvd _ internal prefix), and aliased orig name(convert-xoShared), strongly typed $Mailbox as [System.Object] (get-xomailbox returns that type, not a real 'Mailbox' class).
    # 1:09 PM 8/25/2021 ren convertExoShared -> _convert-xoShared
    # 10:00 AM 12/19/2018 revertExoUserMbx : added post confirm echo
    # 12:31 PM 10/23/2018 ran full pass live, no unusual errors
    .DESCRIPTION
    convert-exoMailboxTypeSharedMailbox() - Set specified EXO mailbox to Shared ('SharedMailbox) type (part of coordinated on-prem ADUser recipienttype hack to make it work without moving mbxs back onprem to convert).
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
    $bRet = convert-exoMailboxTypeSharedMailbox -Mailbox $ombx -whatif -showDebug ;
    Pull the target cloud-first EXO mailbox, and pass it as an object in to the convert-exoMailboxTypeSharedMailbox(), with whatif & showdebug
    .LINK
    #>
    [CmdletBinding()]
    [Alias('convert-xoShared')]
    PARAM(
        [Parameter(Position=0,Mandatory=$True,ValueFromPipeline=$true,HelpMessage="UPN [-upn fname.lname@DOMAIN.COM]")]
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
        $Error.Clear() ;
        foreach($MBX in $Mailbox) {
            Reconnect-EXO @pltRXO;

            # 2:43 PM 10/11/2018 add precheck
            if($MBX |?{$_.recipienttypedetails -eq 'SharedMailbox'}){
                $smsg= "PRE:$($MBX.userprincipalname) has already been converted to RemoteSharedMailbox" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Warn } #Error|Warn|Debug
                else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } else {
                $pltSxM=[ordered]@{
                    Identity=$MBX.userprincipalname ;
                    Type="Shared" ;
                    whatif=$($whatif) ;
                    ErrorAction = 'STOP'
                } ;
                $smsg="set-xomailbox with:`n$(($pltSxM|out-string).trim())`n" ; ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

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

                if(-not $whatif){
                    $smsg = "(waiting for get-xoMailbox to return RecipientTypeDetails -eq 'SharedMailbox')" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    $1F=$false ;
                    Do {
                        if($1F){Sleep -s 5} ;
                        write-host "." -NoNewLine ;
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
                    } Until ($pexombx.RecipientTypeDetails -eq 'SharedMailbox') ;

                    $smsg= "POST:EXO Mailbox`n$(($pexombx| format-list $exprops|out-string ).trim())" ; ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                } else {
                        set-xomailbox @pltSxM ; 
                        $smsg = "(whatif detected, skipping post test)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }  else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                } ;
            } ; # if-E Not Converted test
        } ;  # loop-E
    } ;  # PROC-E
} ; 
#*------^ END Function convert-exoMailboxTypeSharedMailbox ^------