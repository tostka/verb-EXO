#*------v copy-XPermissionGroupToCloudOnly.ps1 v------
function copy-XPermissionGroupToCloudOnly {
    <#
    .SYNOPSIS
    copy-XPermissionGroupToCloudOnly.ps1 - Copy an onprem replicated Mail-Enabled Security Group, used for Mailbox Access grants, to a cloud-only EXO DistributionGroup, to grant EXO perms to foreign-hybrid multi-HCW federated objects in the tenant
    .NOTES
    Version     : 1.0.0.
    Author      : Todd Kadrie
    Website     : http://www.toddomation.com
    Twitter     : @tostka / http://twitter.com/tostka
    CreatedDate : 2021-07-13
    FileName    : copy-XPermissionGroupToCloudOnly.ps1
    License     : MIT License
    Copyright   : (c) 2021 Todd Kadrie
    Github      : https://github.com/tostka/verb-XXX
    Tags        : Powershell
    AddedCredit : REFERENCE
    AddedWebsite: URL
    AddedTwitter: URL
    REVISIONS
    * 2:16 PM 6/24/2024: rem'd out #Requires -RunasAdministrator; sec chgs in last x mos wrecked RAA detection ; pulled msol too
    # 2:49 PM 3/8/2022 pull Requires -modules ...verb-ex2010 ref - it's generating nested errors, when ex2010 requires exo requires ex2010 == loop.
    * 2:40 PM 12/10/2021 more cleanup 
    * 3:51 PM 8/17/2021 added $MembersCloudOnly | select -unique - kept leaking in duplicates in the inputs.
    * 1:40 PM 8/11/2021 ADDED & debugged -Mailbox param (spec target of grants), and code to add-mailboxperm/add-(ad|recipient)permission to OP or EXO target mailbox, and more detailed follow up dump report. Ran against exo-mailbox wio issues. Need to dbug against a still onprem mbx next.
    * 2:19 PM 8/3/2021 step-debugged, looks functional ; init 
    .DESCRIPTION
    copy-XPermissionGroupToCloudOnly.ps1 - Copy an onprem replicated Mail-Enabled Security Group, used for Mailbox Access grants, to a cloud-only EXO DistributionGroup, to grant EXO perms to foreign-hybrid multi-HCW federated objects in the tenant
    This function comes into use when your o365 Tenant/EXO org has hybrid-federated objects. That is, one set of EXO mailboxes federated (and HCW'd) from one on-prem ActiveDirectory/Exchange org, 
    and another set of EXO mailboxes federated (and HCW'd) from *a second separate* on-prem ActiveDirectory/Exchange org. 
    If your Mailbox permission grants are generally performed via OnPrem mail-enabled security groups (which are replicated to cloud), those groups cannot properly accomodate
    Security principals in the second AD org. 
    So this function duplicates a local mail-enabled security group, as a new EXO distributiongroup, with a similar name, and the appended suffixe '_C1' 
    (n.b. in my org, all grant groups end in '-G' by policy, you'll need to tweak the name generation code below if yours lack a '-G' to target for the renames )
    The resulting EXO DG is intended to hold those SecPrincipals that can't be represented in the on-prem Org. 
    In effect you'll have one onprem DG granting permissions for locally federated SecPrins, 
    And this newly duplicated EXO DG granting permissions for externally federated SecPrins.
    .PARAMETER ticket
    ticket number[-ticket nnnnn]
    .PARAMETER SourceGroupName
    Name of on-prem replicated Exchange DistributionGroup to be copied to a cloud-only variant[-SourceGroupName somegroup]
    .PARAMETER Mailbox
    Identifier for the mailbox/mailuser object that the new group should be granted access to (generally matches target of on-prem SourceGroupName permissions grants)[-Mailbox email@domain.com]
    .PARAMETER Owner
    Identifier for the mailbox/mailuser object that will be the Owner of the new group[-Owner email@domain.com]
    .PARAMETER MembersCloudOnly
    Array of cloud-only unreplicated mailbox/mailuser designators to be added as members of the newly copied group[-MembersCloudOnly email@domain.com,email2@domain.com]
    .PARAMETER useEXOv2
    Use EXOv2 (ExchangeOnlineManagement) over basic auth legacy connection [-useEXOv2]
    .PARAMETER Whatif
    Parameter to run a Test no-change pass [-Whatif switch]
    .INPUTS
    None. Does not accepted piped input.(.NET types, can add description)
    .OUTPUTS
    None. Does not accepted piped input.(.NET types, can add description)
    .EXAMPLE
    PS> $whatif = $true ;
        [array]$tgroups = @("627192;LYN-SEC-Email-COMPANYMobilityTeam-G;COMPANYMobilityTeam@COMPANY.com;dccoldiron@charlesmachine.works;member1@domain.com,dccoldiron@charlesmachine.works") ;
        [array]$tgroups += "123457;SIT-SEC-Email-GrantMailbox2-G;GrantMailbox2@domain.com;owner2@domain.com;member1@domain.com,member2@domain.com" ;
        foreach($tgrp in $tgroups){
            $pltCXPermGrp=[ordered]@{
                ticket = $tgrp.split(';')[0] ;
                SourceGroupName = $tgrp.split(';')[1] ;
                Mailbox = $tgrp.split(';')[2] ;
                Owner = $tgrp.split(';')[3] ;
                MembersCloudOnly = $tgrp.split(';')[4].split(',') ;
                verbose=$true ;
                whatif=$($whatif) ;
            } ;
            write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):copy-XPermissionGroupToCloudOnly w`n$(($pltCXPermGrp|out-string).trim())" ;
            copy-XPermissionGroupToCloudOnly @pltCXPermGrp ;
        } ; 
    Example demoing processing of an array of descriptors, as a semicolon-delimited summary of inputs (useful for stacking bulk-creations)
    Schema for the $tgroups input is "[SourceGroupName];[Mailbox];[Owner];[MembersCloudOnly array]"
    .LINK
    https://github.com/tostka/verb-exo
    .LINK
    #>
    ###Requires -Version 5
    ##Requires -Modules ActiveDirectory, MSOnline, AzureAD, ExchangeOnlineManagement, verb-AAD, verb-ADMS, verb-Ex2010, verb-Text
    # 2:49 PM 3/8/2022 pull verb-ex2010 ref - I think it's generating nested errors, when ex2010 requires exo requires ex2010 == loop.
    #Requires -Modules ActiveDirectory, AzureAD, ExchangeOnlineManagement, verb-AAD, verb-ADMS, verb-Text
    ## MSOnline, 
    ##Requires -RunasAdministrator
    # VALIDATORS: [ValidateNotNull()][ValidateNotNullOrEmpty()][ValidateLength(24,25)][ValidateLength(5)][ValidatePattern("(lyn|bcc|spb|adl)ms6(4|5)(0|1).(china|global)\.ad\.COMPANY\.com")][ValidateSet("USEA","GBMK","AUSYD")][ValidateScript({Test-Path $_ -PathType 'Container'})][ValidateScript({Test-Path $_})][ValidateRange(21,65)][ValidateCount(1,3)]
    ## [OutputType('bool')] # optional specified output type
    [CmdletBinding()]
    ##[Alias('somealias')]
    PARAM(
        [Parameter(Mandatory=$true,HelpMessage="ticket number[-ticket nnnnn]")]
        $ticket, 
        [Parameter(Mandatory=$true,HelpMessage="Name of on-prem replicated Exchange DistributionGroup to be copied to a cloud-only variant[-SourceGroupName somegroup]")]
        $SourceGroupName, 
        [Parameter(Mandatory=$true,HelpMessage="Identifier for the mailbox/mailuser object that the new group should be granted access to (generally matches target of on-prem SourceGroupName permissions grants)[-Mailbox email@domain.com]")]
        $Mailbox, 
        [Parameter(Mandatory=$true,HelpMessage="Identifier for the mailbox/mailuser object that will be the Owner of the new group[-Owner email@domain.com]")]
        $Owner, 
        [Parameter(Mandatory=$true,HelpMessage="Array of cloud-only unreplicated mailbox/mailuser designators to be added as members of the newly copied group[-MembersCloudOnly email@domain.com,email2@domain.com]")]
        [array]$MembersCloudOnly, 
        [Parameter(HelpMessage="Use EXOv2 (ExchangeOnlineManagement) over basic auth legacy connection [-useEXOv2]")]
        [switch] $useEXOv2,
        [Parameter(HelpMessage="Whatif Flag  [-whatIf]")]
        [switch] $whatIf
    ) ;
    BEGIN{
        $Verbose = ($VerbosePreference -eq 'Continue') ; 
        $propsdg = 'SamAccountName','ManagedBy','AcceptMessagesOnlyFrom','AcceptMessagesOnlyFromDLMembers','AddressListMembership',
            'Alias','DisplayName','EmailAddresses','ExternalDirectoryObjectId','HiddenFromAddressListsEnabled','EmailAddressPolicyEnabled',
            'PrimarySmtpAddress','RecipientType','RecipientTypeDetails','WindowsEmailAddress','Name','DistinguishedName','WhenChanged','WhenCreated'; 
        $rgxMbxPermLocal = '^(S-\d-\d-\d{2}-\d{10}-\d{9}-\d{10}-\d{5}|NT\sAUTHORITY\\SELF)' ;
        $propsmbxperm = 'User','AccessRights','IsInherited','Deny';
        $propsrcpperm = 'trustee','AccessRights','IsInherited','Deny';
        $propsadperm = 'User','AccessRights','ExtendedRights','IsInherited','Deny';

        connect-AD -Verbose:$false | out-null ; 
        rx10 -Verbose:$false ; rxo  -Verbose:$false ; #cmsol  -Verbose:$false ;
        
    } 
    PROCESS{
        # check ExternalDirectoryObjectId to ensure unfederated
        $sBnr="===v $($SourceGroupName) - $($Owner) v===" ;
        $smsg = $sBnr ; 
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        $smsg = "==Checking for existing:$($SourceGroupName)" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        <# New-exoDistributionGroup -ModeratedBy -RequireSenderAuthenticationEnabled -ModerationEnabled -DisplayName -Confirm -MemberDepartRestriction -IgnoreNamingPolicy -RoomList -HiddenGroupMembershipEnabled -BypassNestedModerationEnabled -CopyOwnerToMember -BccBlocked -Members -MemberJoinRestriction -Type -Alias -ManagedBy -WhatIf -PrimarySmtpAddress -SendModerationNotifications -Notes -OrganizationalUnit -Name -AsJob 
        Set-exoDistributionGroup -HiddenFromAddressListsEnabled
        New-exoDistributionGroup -DisplayName -Name -Members -Type -Alias -PrimarySmtpAddress -ManagedBy -WhatIf -Notes -whatif ; 
        -ManagedBy "Name|Display name|Alias|Distinguished name (DN)|Canonical DN|<domain name>\<account name>|Email address|GUID|LegacyExchangeDN|SamAccountName|User ID or user principal name (UPN)"
        Set-exoDistributionGroup -EmailAddresses -RejectMessagesFromDLMembers -AcceptMessagesOnlyFromSendersOrMembers -AcceptMessagesOnlyFromDLMembers -SimpleDisplayName -MailTip -GrantSendOnBehalfTo -AcceptMessagesOnlyFrom -RejectMessagesFromSendersOrMembers -Alias -DisplayName -ManagedBy -PrimarySmtpAddress -Name -whatif ;
        #>
        if($dg = get-distributiongroup -id $SourceGroupName){
            $tdgName = $dg.Name.replace('-G','-G_C1') ; 
            $nameClean=Remove-StringDiacritic -string $tdgName ;
            $nameClean= Remove-StringLatinCharacters -string $nameClean ;
            $samaccountname=$( ([System.Text.RegularExpressions.Regex]::Replace($nameClean,"[^1-9a-zA-Z_]","").tostring().substring(0,[math]::min([System.Text.RegularExpressions.Regex]::Replace($nameClean,"[^1-9a-zA-Z_]","").tostring().length,20))).toLower() )  ;
            $samaccountname = "$($samaccountname)-$((new-guid).guid.split('-')[0])-C1" ;
            $smsg = "Resolving potential members:`n$(($MembersCloudOnly| select -unique | sort | out-string).trim())" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            $rmbrs = $MembersCloudOnly | select -unique | sort |foreach-object {get-exorecipient -id $_} | select -expand primarysmtpaddress ; 
            $pltNxDG=[ordered]@{
                Notes="$((get-group -id ($dg.alias)).notes),$($ticket) for $($Owner)(Cloud-only replica of on-prem group)" ;
                DisplayName=$tdgName ;
                Name=$tdgName ;
                ManagedBy= $Owner ;
                Members = $rmbrs ; 
                Alias=$samaccountname  ;
                RequireSenderAuthenticationEnabled=$true ; 
                Type = 'Security' ; 
                whatif=$($whatif) ;
                ErrorAction='STOP';
            } ;

            $pltSxDG=[ordered]@{
                identity = $null; 
                HiddenFromAddressListsEnabled=$true;
                whatif=$($whatif) ;
                ErrorAction='STOP';
            } ;
            $smsg = "==Checking for existing:$($tdgName)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            if($xdg = get-exodistributiongroup -id $pltNxDG.DisplayName -ea 0){
                $smsg = "(confirmed existing Dname:'$($xdg.DisplayName)'" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }     else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } else {
                $smsg = "$((get-date).ToString('HH:mm:ss')):xDG:NotFound:$($tgrpName)`nCreating missing SecGrp" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            
                $smsg = "new-exodistributiongroup  w`n$(($pltNxDG|out-string).trim())" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                TRY {
                    $xdg = new-exodistributiongroup  @pltNxDG ;
                    # $xdg captures equiv to get-distibutiongroup 
                    $smsg = "Result:`n$(($xdg|out-string).trim())" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                } CATCH {
                    $smsg = "Failed processing $($_.Exception.ItemName). `nError Message: $($_.Exception.Message)`nError Details: $($_)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } 
                    else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    Continue ;
                } ;
            } ;
            if(!$whatif){
                $pltSxDG.identity = $xdg.primarysmtpaddress ; 
                $smsg = "set-exodistributiongroup w`n$(($pltSxDG|out-string).trim())" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                TRY {
                    set-exodistributiongroup @pltSxDG ;
                    $pxdg = get-exodistributiongroup -id $pltNxDG.DisplayName ;
                    $pxDGm = get-exodistributiongroupmember -id $pltNxDG.DisplayName ;
                } CATCH {
                    $smsg = "Failed processing $($_.Exception.ItemName). `nError Message: $($_.Exception.Message)`nError Details: $($_)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } 
                    else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    Continue ;
                } ;

                if($tmbxr = get-recipient -id $Mailbox -ea 0 ){
                    $smsg = "(-Mailbox:$($tmbxr.PrimarySmtpAddress) specified, adding $($xdg.name) to it's permissions...)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                    TRY {
                        # aliased ExOP|EXO|EXOv2 cmdlets (permits simpler single code block for any of the three variants of targets & syntaxes)
                        # each is '[aliasname];[exOPcmd];[exOcmd] (xOv2cmd is converted from [exocmd])
                        [array]$cmdletMaps= 'ps1GetMbx;get-mailbox;get-exomailbox','ps1SetMbx;Set-Mailbox;Set-exoMailbox','ps1GetMUsr;Get-MailUser;Get-exoMailUser',
                                            'ps1SetMUsr;Set-MailUser;Set-exoMailUser','ps1AddMbxPrm;Add-MailboxPermission;Add-exoMailboxPermission;',
                                            'ps1GetMbxPrm;Get-MailboxPermission;Get-exoMailboxPermission;','ps1RmvMbxPrm;Remove-MailboxPermission;Remove-exoMailboxPermission;',
                                            'ps1AddRcpPrm;Add-ADPermission;Add-exoRecipientPermission;','ps1GetRcpPrm;Get-ADPermission;Get-exoRecipientPermission;',
                                            'ps1RmvRcpPrm;Remove-ADPermission;Remove-exoRecipientPermission;'
                        $OpRcp=$tmbxr ;
                        $pltRXO = [ordered]@{
                            credential =  $credO365TORSID ;
                            Verbose = $($VerbosePreference -eq 'Continue');
                        } ; 
                        reconnect-exo @pltRXO ;
                        foreach($cmdletMap in $cmdletMaps){
                            switch ($OpRcp.recipienttype){
                                "MailUser" {
                                    $iIndex = 2 ;
                                    if($script:useEXOv2){
                                        reconnect-eXO2 @pltRXO ; 
                                        if(!($cmdlet= Get-Command $cmdletMap.split(';')[$iIndex ].replace('-exo','-xo') )){ throw "unable to gcm Alias definition!:$($cmdletMap.split(';')[1])" ; break }
                                        $nalias = set-alias -name ($cmdletMap.split(';')[0]) -value ($cmdlet.name) -passthru ;
                                        write-verbose "$($nalias.Name) -> $($nalias.ResolvedCommandName)" ;
                                    } else {
                                        reconnect-exo @pltRXO ;
                                        if(!($cmdlet= Get-Command $cmdletMap.split(';')[$iIndex ])){ throw "unable to gcm Alias definition!:$($cmdletMap.split(';')[1])" ; break }
                                        $nalias = set-alias -name ($cmdletMap.split(';')[0]) -value ($cmdlet.name) -passthru ;
                                        write-verbose "$($nalias.Name) -> $($nalias.ResolvedCommandName)" ;
                                    } ;
                                }
                                "UserMailbox" { 
                                    $iIndex = 1 ;
                                    reconnect-ex2010 ;
                                    if(!($cmdlet= Get-Command $cmdletMap.split(';')[$iIndex ])){ throw "unable to gcm Alias definition!:$($cmdletMap.split(';')[1])" ; break }
                                    $nalias = set-alias -name ($cmdletMap.split(';')[0]) -value ($cmdlet.name) -passthru ;
                                    write-verbose "$($nalias.Name) -> $($nalias.ResolvedCommandName)" ;
                                }
                                default { throw "Unrecognized recipienttype!:$($OpRcp.recipienttype)" }
                            } ; 
                        } ; 
                        
                        # exo mbx, need to flip to exo rcp, if we're going to get a functional DN for recipientperms cmds: pull the actual mbx instead of rcp (which provided RecipientType to steer balance)
                        $pltGmbx=[ordered]@{
                            Identity=$tmbxr.PrimarySmtpAddress ; 
                            ErrorAction='STOP' ;};

                        $smsg = "$((get-alias ps1GetMbx).definition) w`n$(($pltGmbx|out-string).trim())" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        $tmbxr = ps1GetMbx @pltGmbx ; 
                        
                        $pltAMP=[ordered]@{
                            Identity=$tmbxr.PrimarySmtpAddress ; 
                            User=$pxdg.primarysmtpaddress ; 
                            AccessRights="FullAccess";
                            confirm = $false ; # suppress prompts
                            ErrorAction='STOP' ;
                            whatif=$($whatif);};

                        $pltARP=@{
                            identity=$tmbxr.DistinguishedName ; 
                            trustee=$pxdg.primarysmtpaddress ;
                            AccessRights="SendAs" ;
                            confirm = $false ; # suppress prompts
                            ErrorAction='STOP' ;
                            whatif=$($whatif);}; 
                        # SendAs perms target user onprem, trustee in exo:
                        $smsg = "$((get-alias ps1GetMbxPrm).definition) -Identity $($pltAMP.Identity) | `n?{`$_.user -eq '$($pxdg.name)' -AND `$_.AccessRights -eq '$($pltARP.AccessRights)'}" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        if($mbxperm = ps1GetMbxPrm -Identity $pltAMP.Identity | ?{$_.user -eq $pxdg.name -AND $_.AccessRights -eq $pltAMP.AccessRights}){
                            $smsg = "($($pdxg.name) already granted $($pltAMP.AccessRights) perms on $($pltAMP.identity))" ; 
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        } else {
                            $smsg = "$((get-alias ps1AddMbxPrm).definition) w`n$(($pltAMP|out-string).trim())" ; 
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            $xmp = ps1AddMbxPrm @pltAMP ;
                        } ; 
                        $mbxperm = ps1GetMbxPrm -Identity $pltAMP.Identity -user $pltAMP.user ; 
                        $smsg = "$((get-alias ps1GetMbxPrm).definition):`n$(($mbxperm|ft -wrap $propsmbxperm |out-string).trim())" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                        switch ($OpRcp.recipienttype){
                                "MailUser" {
                                    $pltARP.identity = $tmbxr.distinguishedname ; 
                                    $smsg = "$((get-alias ps1GetRcpPrm).definition) -Identity $($pltARP.Identity) | `n?{`$_.trustee -eq '$($pxdg.name)' -AND `$_.AccessRights -eq '$($pltARP.AccessRights)'}" ; 
                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                    if($rcpperm = ps1GetRcpPrm -Identity $pltARP.Identity | ?{$_.trustee -eq $pxdg.name -AND $_.AccessRights -eq $pltARP.AccessRights}){
                                        $smsg = "(Trustee:$($pxdg.name) already granted AccessRights:$($pltARP.AccessRights) perms on `n$($pltARP.identity))" ; 
                                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                    } else {
                                        $smsg = "$((get-alias ps1AddRcpPrm).definition) w`n$(($pltARP|out-string).trim())" ; 
                                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                        $xmp = ps1AddRcpPrm @pltARP ;
                                    } ; 
                                    $rcpperm= ps1GetRcpPrm -Identity $pltARP.Identity -Trustee $pltARP.trustee -errorAction STOP ; 
                                    $smsg = "$((get-alias ps1GetRcpPrm).definition):`n$(($rcpperm|ft -wrap $propsrcpperm |out-string).trim())" ; 
                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                } ;
                                "UserMailbox" { 
                                    $pltARP.remove('AccessRights') ; 
                                    $pltARP.add('ExtendedRights','Send As') ; 
                                    $pltARP.identity = $tmbxr.distinguishedname ; 
                                    $smsg = "$((get-alias ps1GetRcpPrm).definition) -Identity $($pltARP.Identity) | ?{`$_.user -eq '$($pxdg.name)' -AND `$_.ExtendedRights -eq '$($pltARP.AccessRights)'}" ; 
                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                    if($rcpperm = ps1GetRcpPrm -Identity $pltARP.Identity | ?{$_.user -eq $pxdg.name -AND $_.ExtendedRights -eq $pltARP.AccessRights}){
                                        $smsg = "($($pdxg.name) already granted $($pltARP.AccessRights) perms on $($pltARP.identity))" ; 
                                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                    } else {
                                        $smsg = "$((get-alias ps1AddRcpPrm).definition) w`n$(($pltARP|out-string).trim())" ; 
                                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                        $xmp = ps1AddRcpPrm @pltARP ;
                                    } ; 
                                    $rcpperm= $rcpperm = ps1GetRcpPrm -Identity $pltARP.Identity | ?{$_.user -eq $pxdg.name -AND $_.ExtendedRights -eq $pltARP.AccessRights} ; 
                                    $smsg = "$((get-alias ps1GetRcpPrm).definition) w`n$(($rcpperm|ft -wrap $propsadperm |out-string).trim())" ; 
                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                    # set common props for final report
                                    $propsrcpperm = $propsadperm ; 
                                } ;
                        } ;  # switch-E

                    } CATCH {
                        $smsg = "Failed processing $($_.Exception.ItemName). `nError Message: $($_.Exception.Message)`nError Details: $($_)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } 
                        else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        Continue ;
                    } ;


                } else { 
                    $smsg = "(No -Mailbox specified, slipping $($xdg.name) permissions grant...)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                } ;
                

                $hMsg = @"

*------v REVIEW RESULTS v------

POST:exodistributiongroup
-----------
$(($pxdg|fl $propsdg|out-string).trim())
-----------

Members:
-----------
$(($pxDGm.PrimarySmtpAddress|out-string).trim())
-----------
"@ ; 

            if($Mailbox){
                $hMsg += "Associated Mailbox Permissions:`n$(($mbxperm|ft -wrap $propsmbxperm |out-string).trim())`n`n" ;     

                $hMsg += "Associated Recipient Permissions:`n$(($rcpperm|ft -wrap $propsrcpperm  |out-string).trim())`n`n" ; 
            } ;
            $hMsg += "*------^ REVIEW RESULTS ^------`n" ; 

            $smsg = $hMsg ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

            } else {
                $smsg = "(-whatif detected, skipping:set-exodistributiongroup @pltNxDG" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } ; 
        } else { 
            $smsg = "Unable to get-distributiongroup -id $($SourceGroupName) ; aborting!" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } 
            else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        } ;
        $smsg = $sBnr.replace('=v','=^').replace('v=','^=') ; 
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        
    }
    END{}
 }

#*------^ copy-XPermissionGroupToCloudOnly.ps1 ^------