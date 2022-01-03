#*------v Function Resolve-xoRcps v------
function Resolve-xoRcps {
    <#
    .SYNOPSIS
    Resolve-xoRcps.ps1 - run a get-exorecipient to re-resolve an array of Recipients into the matching primarysmtpaddress
    .NOTES
    Version     : 1.0.0
    Author      : Todd Kadrie
    Website     :	http://www.toddomation.com
    Twitter     :	@tostka / http://twitter.com/tostka
    CreatedDate : 2021-09-02
    FileName    : Resolve-xoRcps
    License     : MIT License
    Copyright   : (c) 2020 Todd Kadrie
    Github      : https://github.com/tostka/verb-XXX
    Tags        : Powershell
    AddedCredit : REFERENCE
    AddedWebsite:	URL
    AddedTwitter:	URL
    REVISIONS
    *3:20 PM 12/30/2021 expanded, added params: getGroups, getRecipients, getMailboxPrincipals, PreviewThreshold, UpdateInterval, returnObject;
        expanded verbose echos and reporting, the above -get* params shift the complicated regexes internally, where one of the three types is desired. 
    * 9:16 AM 12/3/2021 added pswlt support
    * 8/30/21 init vers
    .DESCRIPTION
    Resolve-xoRcps.ps1 - run a get-exorecipient to re-resolve an array of Recipients into the matching primarysmtpaddress
    
    Backing out the RecipientTypeDetails combos for various niches (to use on the (Match|Block)RecipientTypeDetails param)

    [Get-Recipient (ExchangePowerShell) | Microsoft Docs - docs.microsoft.com/](https://docs.microsoft.com/en-us/powershell/module/exchange/get-recipient?view=exchange-ps)
    -RecipientType
        The RecipientType parameter filters the results by the specified recipient type. Valid values are:
        'DynamicDistributionGroup','MailContact','MailNonUniversalGroup','MailUniversalDistributionGroup',
            'MailUniversalSecurityGroup','MailUser','PublicFolder','UserMailbox'
    -RecipientTypeDetails
        'DiscoveryMailbox','DynamicDistributionGroup','EquipmentMailbox','GroupMailbox','GuestMailUser',
            'LegacyMailbox','LinkedMailbox','LinkedRoomMailbox','MailContact','MailForestContact','MailNonUniversalGroup',
            'MailUniversalDistributionGroup','MailUniversalSecurityGroup','MailUser','PublicFolder','PublicFolderMailbox',
            'RemoteEquipmentMailbox','RemoteRoomMailbox','RemoteSharedMailbox','RemoteTeamMailbox','RemoteUserMailbox',
            'RoomList','RoomMailbox','SchedulingMailbox','SharedMailbox','TeamMailbox','UserMailbox'

    # run the RTD set, pulling one of each type and dumping back the rt|rtd combos, to build rgxs:
    $rtds = 'DiscoveryMailbox','DynamicDistributionGroup','EquipmentMailbox','GroupMailbox','GuestMailUser',
        'LegacyMailbox','LinkedMailbox','LinkedRoomMailbox','MailContact','MailForestContact','MailNonUniversalGroup',
        'MailUniversalDistributionGroup','MailUniversalSecurityGroup','MailUser','PublicFolder','PublicFolderMailbox',
        'RemoteEquipmentMailbox','RemoteRoomMailbox','RemoteSharedMailbox','RemoteTeamMailbox','RemoteUserMailbox',
        'RoomList','RoomMailbox','SchedulingMailbox','SharedMailbox','TeamMailbox','UserMailbox' ; 
    $rtypes = @() ; 
    foreach($rtd in $rtds){
        write-host "==rtd:$($rtd)" ; 
        $rtypes += get-exorecipient -filter "Recipienttypedetails -eq '$rtd'" -ResultSize 1 ; 
    } ; 
    $rtypes | sort RecipientType,RecipientTypeDetails | ft -auto alias,primarys*,recipientt*

    Sanitized Output: (clearly our Tenant did not have quite a few of the RTD types queried)
    ObjType                                                      RecipientType                  RecipientTypeDetails
    -----                                                        -------------                  --------------------
    [DYNAMICDISTRIBUTIONGROUP]                                   DynamicDistributionGroup       DynamicDistributionGroup
    [MAILCONTACT]                                                MailContact                    MailContact
    [UNIFIEDGROUP]                                               MailUniversalDistributionGroup GroupMailbox
    [DISTRIBUTIONGROUP]                                          MailUniversalDistributionGroup MailUniversalDistributionGroup
    [ROOMLIST-DISTRIBUTIONGROUP]                                 MailUniversalDistributionGroup RoomList
    [MAIL-ENABLED SECURITYGROUP]                                 MailUniversalSecurityGroup     MailUniversalSecurityGroup
    [GUEST]                                                      MailUser                       GuestMailUser
    [MAILUSER]                                                   MailUser                       MailUser
    [DISCOVERYSEARCH MAILBOX]                                    UserMailbox                    DiscoveryMailbox
    [EQUIPMENTMAILBOX]                                           UserMailbox                    EquipmentMailbox
    [ROOMMAILBOX]                                                UserMailbox                    RoomMailbox
    [MS BOOKING APP MBX]                                         UserMailbox                    SchedulingMailbox
    [SHAREDMAILBOX]                                              UserMailbox                    SharedMailbox
    [USERMAILBOX]                                                UserMailbox                    UserMailbox

    # all the variant RTDs for 'group' rt's:
    $rtype = $rtypes |?{$_.RecipientType -like '*group*'} | select -expand RecipientTypeDetails | select -Unique
    [regex]$rgx = ('(' + (($rtype |%{[regex]::escape($_)}) -join '|') + ')') ;
    $rgx.tostring() ;
    # 'groups' rtd rgx : (groupmailbox covers UnifiedGrps)
    $_.RecipientTypeDetails -match '(DynamicDistributionGroup|GroupMailbox|MailUniversalDistributionGroup|MailUniversalSecurityGroup|RoomList)'

    # now do secprins: RecipientType: UserMailbox, MailUser
    $rtype = $rtypes |?{$_.RecipientType -like '*user*'} | select -expand RecipientTypeDetails | select -Unique ;
    [regex]$rgx = ('(' + (($rtype |%{[regex]::escape($_)}) -join '|') + ')') ;
    $rgx.tostring() ;
    # 'core' secprin rtd rgx:
    $_.RecipientTypeDetails -match '(DiscoveryMailbox|EquipmentMailbox|GuestMailUser|MailUser|RoomMailbox|SchedulingMailbox|SharedMailbox|UserMailbox)' ; 

    # sender/recipients (approved|blocked targets):  Valid values for this parameter are individual senders in your organization (mailboxes, mail users, and mail contacts) 
    # RecipientType: UserMailbox, MailUser, MailContact
    $rtype = $rtypes |?{$_.RecipientType -match '(User|Contact)'} | select -expand RecipientTypeDetails | select -Unique ;
    [regex]$rgx = ('(' + (($rtype |%{[regex]::escape($_)}) -join '|') + ')') ;
    $rgx.tostring() ;
    # sender/recipients rtd rgx:
    $_.RecipientTypeDetails -match '(DiscoveryMailbox|EquipmentMailbox|GuestMailUser|MailContact|MailUser|RoomMailbox|SchedulingMailbox|SharedMailbox|UserMailbox)'
    - DiscoveryMailbox discovery are for eDisc, not mail delivery
    
    # moderated by:  must be a mailbox, mail user, or mail contact: RecipientType: UserMailbox, MailUser, MailContact (same as above ☝🏻 )

    # mailbox secprins: required to do accessgrant on a mailbox
    [Add-MailboxPermission (ExchangePowerShell) | Microsoft Docs - docs.microsoft.com/](https://docs.microsoft.com/en-us/powershell/module/exchange/add-mailboxpermission?view=exchange-ps)
        You can specify the following types of users or groups (security principals) for this parameter:
            Mailbox users
            Mail users
            Security groups
    
        -- those phrases are RecipientType values, with spaces added - but not sure they really mean "anything of those specific RT's"...?
        -- though you might be able to use a *licensed* sharedmailbox to open another mailbox (?), they won't be able to do it natively, esp with disabled User logon. 
        -- rooms are disabled for logon. like shared, & equipment
        -- prob should exclude non-interactive logon & system in theory: DiscoveryMailbox|SchedulingMailbox|SharedMailbox|EquipmentMailbox|RoomMailbox
        -- CORRECTION: looped through full set of RT:UserMailbox types in the Tenant, *every* one of them added wo complaint using add-mailboxpermission & add-recipientpermission, 
            although many - unlicensed - would likely be unable to actually open another mailbox. 
        -- so technically, it appears should use the entire set, as they *technically* add wo complaint
    $rtype = $rtypes |?{$_.RecipientType -match '(User|MailUniversalSecurityGroup)'} | select -expand RecipientTypeDetails | select -Unique;
    [regex]$rgx = ('(' + (($rtype |%{[regex]::escape($_)}) -join '|') + ')') ;
    $rgx.tostring() ;
    # mailbox secprins (perm grants)
    $_.RecipientTypeDetails -match '(DiscoveryMailbox|EquipmentMailbox|GuestMailUser|MailUniversalSecurityGroup|MailUser|RoomMailbox|SchedulingMailbox|SharedMailbox|UserMailbox)' 

    .PARAMETER Recipients
    Array of Recipients to be resolved against current Exchange environment [-Recipients `$ModeratedBy ]
    .PARAMETER MatchRecipientTypeDetails
    Regex for RecipientTypeDetails value to require for matched Recipients [-MatchRecipientTypeDetails '(UserMailbox|MailUser)']
    .PARAMETER BlockRecipientTypeDetails
    Regex for RecipientTypeDetails value to filter out of matched Recipients [-Block '(MailContact|GuestUser)']
    .PARAMETER getGroups
    Switch that specifies the return of solely 'group' recipients (RecipientTypeDetails matching:(DynamicDistributionGroup|GroupMailbox|MailUniversalDistributionGroup|MailUniversalSecurityGroup|RoomList)) [-getGroup]
    .PARAMETER getRecipients
    Switch that specifies the return of solely 'recipient' objects (RecipientTypeDetails matching:(DiscoveryMailbox|EquipmentMailbox|GuestMailUser|MailContact|MailUser|RoomMailbox|SchedulingMailbox|SharedMailbox|UserMailbox)) [-getRecipients]
    .PARAMETER getMailboxPrincipals
    Switch that specifies the return of solely 'Mailbox Security Principal' recipients (RecipientTypeDetails matching:'(DiscoveryMailbox|EquipmentMailbox|GuestMailUser|MailUniversalSecurityGroup|MailUser|RoomMailbox|SchedulingMailbox|SharedMailbox|UserMailbox)') [-getRecipients]
    .PARAMETER useEXOv2
    Use EXOv2 (ExchangeOnlineManagement) over basic auth legacy connection [-useEXOv2]
    .PARAMETER PreviewThreshold
    Maximum number of preview resolved to display in console (defaults to 25)[-PreviewThreshold 10]
    .PARAMETER UpdateInterval
    Dot crawl update interval (one dot per `$UpdateInterval processed recipients - defaults to 3)[-UpdateInterval 10]
    .PARAMETER returnObject
    Switch to return full Recipient object to pipeline for each resolved recipient (rather than default, PrimarySmtpAddress property) [-returnObject]
    .EXAMPLE
    PS> $pltSDdg.RejectMessagesFrom = (Resolve-xoRcps -Recipients $srcDg.RejectMessagesFrom -MatchRecipientTypeDetails -MatchRecipientTypeDetails '(UserMailbox|MailUser|GuestMailUser|MailContact)' -Verbose:($VerbosePreference -eq 'Continue') -ErrorAction Continue)  ; 
    Resolve mail sender/recipient recip designators on the RejectMessagesFrom varito EXO recipient objects, with -ErrorAction:Continue (echo lookup fails, continue looping), and return the primarysmtpaddresses as an array
    .EXAMPLE
    PS> $pltSDdg.RejectMessagesFrom = (Resolve-xoRcps -Recipients $srcDg.RejectMessagesFrom -MatchRecipientTypeDetails -MatchRecipientTypeDetails '(UserMailbox|MailUser|GuestMailUser)' -Verbose:($VerbosePreference -eq 'Continue') -ErrorAction Continue)  ; 
    Resolve mail 'Security Principal' recip designators on the RejectMessagesFrom varito EXO recipient objects, with -ErrorAction:Continue (echo lookup fails, continue looping), and return the primarysmtpaddresses as an array
    .EXAMPLE
    PS> $pltSDdg.AcceptMessagesOnlyFromDLMembers = (Resolve-xoRcps -Recipients $ApprovedSenderDLs -MatchRecipientTypeDetails '(MailUniversalDistributionGroup|DynamicDistributionGroup|GroupMailbox)' -Verbose:$($VerbosePreference -eq 'Continue') )  ;
    Resolve mail recipient 'group' objects (covers DG| DDG| UnifiedGrp)
    .EXAMPLE
    PS> if($pltSDdg.RejectMessagesFrom){
            $pltSDdg.RejectMessagesFrom = (Resolve-xoRcps -Recipients $srcDg.RejectMessagesFrom -MatchRecipientTypeDetails '(UserMailbox|MailUser|GuestMailUser|MailContact)' -Verbose:($VerbosePreference -eq 'Continue') )  ; 
        } ;
    Resolve recip designators on the RejectMessagesFrom value, to EXO recipient objects, and return the primarysmtpaddress
    .EXAMPLE
    PS> $pltSDdg.AcceptMessagesOnlyFromDLMembers = (Resolve-xoRcps -Recipients $ApprovedSenderDLs -getGroups -Verbose:$($VerbosePreference -eq 'Continue') )  ;
    Resolve mail recipient 'group' objects using the -getGroups parameter (covers DG| DDG| UnifiedGrp)
    .EXAMPLE
    PS> $pltSDdg.AcceptMessagesOnlyFrom = (Resolve-xoRcps -Recipients $ApprovedSenderDLs -getRecipients -Verbose:$($VerbosePreference -eq 'Continue') )  ;
    Resolve mail recipient 'recipient' objects (senders/recipients) using the -getRecipients parameter.
    .EXAMPLE
    PS> $pltSDdg.AcceptMessagesOnlyFrom = (Resolve-xoRcps -Recipients $ApprovedSenderDLs -getRecipients -Verbose:$($VerbosePreference -eq 'Continue') )  ;
    Resolve mail Security Principal recipients (Those that can be used with add-mailboxpermission & add-recipientpermission) using the -getMailboxPrincipals parameter
    (covers DiscoveryMailbox|EquipmentMailbox|GuestMailUser|MailUniversalSecurityGroup|MailUser|RoomMailbox|SchedulingMailbox|SharedMailbox|UserMailbox)
    .EXAMPLE
    PS> $FullRecipientArray = (Resolve-xoRcps -Recipients $ApprovedSenderDLs -getRecipients -returnObject -Verbose:$($VerbosePreference -eq 'Continue') )  ;
    Resolve mail recipient 'recipient' objects (senders/recipients) using the -getRecipients parameter, and return the full Recipient object for each, to the pipeline.                
    .LINK
    https://github.com/tostka/verb-EXO
    #>
    [CmdletBinding()]
    PARAM(
        [Parameter(Mandatory=$True,HelpMessage="Array of Recipients to be resolved against current Exchange environment [-Recipients `$ModeratedBy ]")]
        [array]$Recipients,
        [Parameter(ParameterSetName='MatchRecipients',HelpMessage="Regex for RecipientTypeDetails value to require for matched Recipients [-MatchRecipientTypeDetails '(UserMailbox|MailUser)']")]
        [string]$MatchRecipientTypeDetails,
        [Parameter(HelpMessage="Regex for RecipientTypeDetails value to filter out of matched Recipients [-Block '(MailContact|GuestUser)']")]
        [string]$BlockRecipientTypeDetails,
        [Parameter(ParameterSetName='groups',HelpMessage="Switch that specifies the return of solely 'group' recipients (RecipientTypeDetails matching:(DynamicDistributionGroup|GroupMailbox|MailUniversalDistributionGroup|MailUniversalSecurityGroup|RoomList)) [-getGroup]")]
        [switch] $getGroups,
        [Parameter(ParameterSetName='recipients',HelpMessage="Switch that specifies the return of solely 'recipient' objects (RecipientTypeDetails matching:(DiscoveryMailbox|EquipmentMailbox|GuestMailUser|MailContact|MailUser|RoomMailbox|SchedulingMailbox|SharedMailbox|UserMailbox)) [-getRecipients]")]
        [switch] $getRecipients,
        [Parameter(ParameterSetName='secprincipals',HelpMessage="Switch that specifies the return of solely 'Mailbox Security Principal' recipients (RecipientTypeDetails matching:(DiscoveryMailbox|EquipmentMailbox|GuestMailUser|MailUniversalSecurityGroup|MailUser|RoomMailbox|SchedulingMailbox|SharedMailbox|UserMailbox)) [-getRecipients]")]
        [switch] $getMailboxPrincipals,
        [Parameter(HelpMessage="Use EXOv2 (ExchangeOnlineManagement) over basic auth legacy connection [-useEXOv2]")]
        [switch] $useEXOv2,
        [Parameter(HelpMessage="Maximum number of preview resolved to display in console (defaults to 25)[-PreviewThreshold 10]")]
        [int] $PreviewThreshold = 25,
        [Parameter(HelpMessage="Dot crawl update interval (one dot per `$UpdateInterval processed recipients - defaults to 3)[-UpdateInterval 10]")]
        [int] $UpdateInterval = 3,
        [Parameter(HelpMessage="Switch to return full Recipient object to pipeline for each resolved recipient (rather than default, PrimarySmtpAddress property) [-returnObject]")]
        [switch] $returnObject
    ) 
    <# Can capture the ErrorAction (not necessary, just like -verbose, if call is made with -erroraction specified, it auto-applies to *all* cmds run in the advanced function, that support the -ea param 
    - it's effectively setting $ErrorActionPreference for the func)
    Most useful purp would be if you want to echo status back.
    #>
    #$vErrorAction = $PSBoundParameters["ErrorAction"] ; 
    $verbose = ($VerbosePreference -eq "Continue") ;

    if($getGroups){$MatchRecipientTypeDetails = '(DynamicDistributionGroup|GroupMailbox|MailUniversalDistributionGroup|MailUniversalSecurityGroup|RoomList)'} 
    elseif($getRecipients){$MatchRecipientTypeDetails = '(DiscoveryMailbox|EquipmentMailbox|GuestMailUser|MailContact|MailUser|RoomMailbox|SchedulingMailbox|SharedMailbox|UserMailbox)'} 
    elseif($getMailboxPrincipals){$MatchRecipientTypeDetails = '(DiscoveryMailbox|EquipmentMailbox|GuestMailUser|MailUniversalSecurityGroup|MailUser|RoomMailbox|SchedulingMailbox|SharedMailbox|UserMailbox)'} 
    
    if ($script:useEXOv2) { reconnect-eXO2 }
    [array]$cmdletMaps = 'ps1GetxRcp;get-exorecipient;' ;
    foreach($cmdletMap in $cmdletMaps){
        if($script:useEXOv2){
            if(!($cmdlet= Get-Command $cmdletMap.split(';')[1].replace('-exo','-xo') )){ throw "unable to gcm Alias definition!:$($cmdletMap.split(';')[1])" ; break }
            $nAName = ($cmdletMap.split(';')[0]) ; 
            if(!($nalias = get-alias -name $nAName -ea 0 )){
                $nalias = set-alias -name $nAName -value ($cmdlet.name) -passthru ;
                $smsg = "$($nalias.Name) -> $($nalias.ResolvedCommandName)" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
            } ;
        } else {
            if(!($cmdlet= Get-Command $cmdletMap.split(';')[1])){ throw "unable to gcm Alias definition!:$($cmdletMap.split(';')[1])" ; break }
            $nAName = ($cmdletMap.split(';')[0]);
            if(!($nalias = get-alias -name $nAName -ea 0 )){
                $nalias = set-alias -name ($cmdletMap.split(';')[0]) -value ($cmdlet.name) -passthru ;
                $smsg = "$($nalias.Name) -> $($nalias.ResolvedCommandName)" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 

            } ; 
        } ;
    } ;
    if ($script:useEXOv2) { reconnect-eXO2 }
    else { reconnect-EXO } ;
    if($Recipients){
        $Procd = 0 ; 
        $smsg = "(Resolving recipients...)" ; 
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
        $resolvedRecipients = $Recipients | foreach-object {
            # use the EA if spec'd
            ps1GetxRcp -identity $_ ;
            $Procd ++ ; 
            if(-not($Procd % $UpdateInterval)){
                write-host "." -NoNewline ; 
            } ; 
        } ; 
        write-host "" ; 
        if($MatchRecipientTypeDetails){
            $smsg = "(Resolve-xoRcps:($(($resolvedRecipients|measure).count)) PRE MatchRecipientTypeDetails:"
            if(($resolvedRecipients|measure).count -lt $PreviewThreshold){
                $smsg += "`n$(($resolvedRecipients.primarysmtpaddress|out-string).trim()))" ; 
            } else { 
                $smsg += "`n$(($resolvedRecipients.primarysmtpaddress | select -first $PreviewThreshold |out-string).trim()))`n..." ; 
            } ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
            $resolvedRecipients = $resolvedRecipients |?{$_.RecipientTypeDetails -match $MatchRecipientTypeDetails} ; 
            $smsg = "(Resolve-xoRcps:($(($resolvedRecipients|measure).count)) POST MatchRecipientTypeDetails:`n$(($resolvedRecipients.primarysmtpaddress|out-string).trim()))" ; 
            if(($resolvedRecipients|measure).count -lt $PreviewThreshold){
                $smsg += "`n$(($resolvedRecipients.primarysmtpaddress|out-string).trim()))" ; 
            } else { 
                $smsg += "`n$(($resolvedRecipients.primarysmtpaddress | select -first $PreviewThreshold |out-string).trim()))`n..." ; 
            } ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
        } ; 
        if($BlockRecipientTypeDetails){
            $smsg = "(Resolve-xoRcps:($(($resolvedRecipients|measure).count)) PRE BlockRecipientTypeDetails:`n$(($resolvedRecipients.primarysmtpaddress|out-string).trim()))" ; 
            if(($resolvedRecipients|measure).count -lt $PreviewThreshold){
                $smsg += "`n$(($resolvedRecipients.primarysmtpaddress|out-string).trim()))" ; 
            } else { 
                $smsg += "`n$(($resolvedRecipients.primarysmtpaddress | select -first $PreviewThreshold |out-string).trim()))`n..." ; 
            } ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
            $resolvedRecipients = $resolvedRecipients |?{$_.RecipientTypeDetails -notmatch $BlockRecipientTypeDetails} ; 
            $smsg = "(Resolve-xoRcps:($(($resolvedRecipients|measure).count)) POST BlockRecipientTypeDetails:`n$(($resolvedRecipients.primarysmtpaddress|out-string).trim()))" ; 
            if(($resolvedRecipients|measure).count -lt $PreviewThreshold){
                $smsg += "`n$(($resolvedRecipients.primarysmtpaddress|out-string).trim()))" ; 
            } else { 
                $smsg += "`n$(($resolvedRecipients.primarysmtpaddress | select -first $PreviewThreshold |out-string).trim()))`n..." ; 
            } ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
        } ; 
        if($returnObject){
            $smsg = "(-Returnobject: returning full recipient object array to pipeline)" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
            $resolvedRecipients |write-output ;
        } else { 
            $resolvedRecipients.primarysmtpaddress |write-output ;
        } ; 
        $smsg = "(Resolve-xoRcps:returning:($(($resolvedRecipients|measure).count))`n$(($resolvedRecipients.primarysmtpaddress|out-string).trim()))" ; 
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
    } else { 
        $smsg = "Resolve-xoRcps:No Recipients specified" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        $null | write-output ;
    } ; 
} ; 
#*------^ END Function Resolve-xoRcps ^------