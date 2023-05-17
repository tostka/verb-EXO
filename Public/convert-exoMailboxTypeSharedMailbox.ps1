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
    * 8:46 AM 5/17/2023 add to vXO; ren'd to convert-exoMailboxTypeSharedMailbox (rmvd _ internal prefix), and aliased orig name(convert-xoShared), strongly typed $Mailbox as [System.Object] (get-xomailbox returns that type, not a real 'Mailbox' class).
    # 1:09 PM 8/25/2021 ren convertExoShared -> _convert-xoShared
    # 10:00 AM 12/19/2018 revertExoUserMbx : added post confirm echo
    # 12:31 PM 10/23/2018 ran full pass live, no unusual errors
    .DESCRIPTION
    convert-exoMailboxTypeSharedMailbox() - Set specified EXO mailbox to Shared ('SharedMailbox) type (part of coordinated on-prem ADUser recipienttype hack to make it work without moving mbxs back onprem to convert).
    .PARAMETER  Mailbox, EXO Mailbox Object
    EXO Mailbox Object
    .PARAMETER Whatif
    Parameter to run a Test no-change pass, and log results [-Whatif switch]
    .PARAMETER ShowDebug
    Parameter to display Debugging messages [-ShowDebug switch]
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    Returns RemoteMailbox object, or $false on failure.
    .EXAMPLE
    $ombx=get-xomailbox -id $targUPN -ea stop ;
    $bRet=convert-exoMailboxTypeSharedMailbox -Mailbox $ombx -whatif -showDebug ;
    Pull the target cloud-first EXO mailbox, and pass it as an object in to the convert-exoMailboxTypeSharedMailbox(), with whatif & showdebug
    .LINK
    #>
    [CmdletBinding()]
    [Alias('convert-xoShared')]
    Param(
        [Parameter(Position=0,Mandatory=$True,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,HelpMessage="UPN [-upn fname.lname@DOMAIN.COM]")]
        [ValidateNotNullOrEmpty()]
        [System.Object]$Mailbox,
        [switch] $showDebug,
        [Parameter(HelpMessage="Whatif Flag  [-whatIf]")]
        [switch] $whatIf
    ) # PARAM BLOCK END

    $exprops="SamAccountName","RecipientType","RecipientTypeDetails","UserPrincipalName" ;

    Reconnect-EXO @pltRXO;

    # 2:43 PM 10/11/2018 add precheck
    if($Mailbox |?{$_.recipienttypedetails -eq 'SharedMailbox'}){
        $smsg= "PRE:$($Mailbox.userprincipalname) has already been converted to RemoteSharedMailbox" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Warn } #Error|Warn|Debug
        else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
    } else {
        #$oMSUsr=get-msoluser -UserPrincipalName $adu.UserPrincipalName ;
        #$bRet=set-xomailbox -id lynctest9@DOMAIN.COM  -Type Shared -WhatIf ;
        $splt=[ordered]@{
            Identity=$Mailbox.userprincipalname ;
            Type="Shared" ;
            whatif=$($whatif) ;
        } ;
        $smsg="set-xomailbox with:`n$(($splt|out-string).trim())`n" ; ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

        # 9:51 AM 12/19/2018move exec out here, get the whatif test
        $Exit = 0 ;
        Do {
            Try {
                set-xomailbox @splt ;
                $Exit = $DoRetries ;
                $true | write-output ;
            } Catch {
                $smsg = "Failed to exec cmd because: $($Error[0])" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                Start-Sleep -Seconds $RetrySleep ;
                $tryNo=0 ;
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

        if(!$whatif){
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
                        $pexombx=get-xomailbox -id $splt.identity -ea stop ;
                        $Exit = $DoRetries ;
                    } Catch {
                        $smsg = "Failed to exec cmd because: $($Error[0])" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                        else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        Start-Sleep -Seconds $RetrySleep ;
                        $tryNo=0 ;
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
             set-xomailbox @splt ; 
             $smsg = "(whatif detected, skipping post test)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }  else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        } ;
    } ; # if-E Not Converted test

     <# error ==7:24 pm 10/11/2018: generated here
        19:09:34:#*======v $tmbx:(135/1224):MAILBOX@DOMAIN.COM v======
        ==MAILBOX@DOMAIN.COM:
        19:09:35:PRE:ADUser
        samaccountname             : MAILBOX
        msExchRemoteRecipientType  : 4
        msExchRecipientDisplayType : -2147483642
        msExchRecipientTypeDetails : 2147483648
        userprincipalname          : MAILBOX@DOMAIN.COM
        19:09:35:PRE:Ex Remotemailbox
        SamAccountName       : MAILBOX
        RecipientType        : MailUser
        RecipientTypeDetails : RemoteUserMailbox
        UserPrincipalName    : MAILBOX@DOMAIN.COM
        19:09:35:Set-aduser with:
        Name                           Value
        ----                           -----
        Identity                       MAILBOX
        Replace                        {msExchRecipientTypeDetails, msExchRemoteRecipientType}
        server                         LYNMS812

        19:09:35:POST:ADUser
        samaccountname             : MAILBOX
        msExchRemoteRecipientType  : 100
        msExchRecipientDisplayType : -2147483642
        msExchRecipientTypeDetails : 34359738368
        UserPrincipalName          : MAILBOX@DOMAIN.COM
        DistinguishedName          : CN=MAILBOX,OU=Generic Email Accounts,OU=LYN,DC=global,DC=ad,DC=toro,DC=com

        19:09:35:POST:Ex Remotemailbox
        SamAccountName       : MAILBOX
        RecipientType        : MailUser
        RecipientTypeDetails : RemoteSharedMailbox
        UserPrincipalName    : MAILBOX@DOMAIN.COM
        .19:09:36:set-xomailbox with:
        Name                           Value
        ----                           -----
        Identity                       MAILBOX@DOMAIN.COM
        Type                           Shared
        whatif                         False

        Error on proxy command 'Set-Mailbox -Identity:'MAILBOX@DOMAIN.COM' -Type:'Shared' -WhatIf:$False -Confirm:$False -Force:$True' to server DM6PR04MB4953.namprd04.prod.outlook.com: Server version
        15.20.1228.0000, Proxy method PSWS:
        Cmdlet error with following error message:
        Microsoft.Exchange.Data.Directory.ADServerSettingsChangedException: An error caused a change in the current set of domain controllers..
        [Server=DM5PR0401MB3543,RequestId=b8e8641e-0c35-42a3-b1f4-55042a53e11c,TimeStamp=10/12/2018 12:09:37 AM] .
            + CategoryInfo          : NotSpecified: (:) [Set-Mailbox], CmdletProxyException
            + FullyQualifiedErrorId : Microsoft.Exchange.Configuration.CmdletProxyException,Microsoft.Exchange.Management.RecipientTasks.SetMailbox
            + PSComputerName        : ps.outlook.com

        19:09:37:POST:Confirming RecipientTypeDetails -eq SharedMailbox...
        .........
        #>

} ; 
#*------^ END Function convert-exoMailboxTypeSharedMailbox ^------