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
    .PARAMETER  users
    Array of user descriptors: displayname, emailaddress, UPN, samaccountname (checks clipboard where unspecified)
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
        [array]$tgroups = "123456;SIT-SEC-Email-MailboxName-G;owner@domain.com;member1@domain.com,member2@domain.com" ;
        [array]$tgroups += "123457;SIT-SEC-Email-MailboxName2-G;owner2@domain.com;member1@domain.com,member2@domain.com" ;
        foreach($tgrp in $tgroups){
            $ticket = $tgrp.split(';')[0] ;
            $SourceGroupName = $tgrp.split(';')[1] ;
            $Owner = $tgrp.split(';')[2] ; 
            $MembersCloudOnly = $tgrp.split(';')[3].split(',') ;
            $pltCXPermGrp=[ordered]@{
                ticket = $tgrp.split(';')[0] ;
                SourceGroupName = $tgrp.split(';')[1] ;
                Owner = $tgrp.split(';')[2] ; 
                MembersCloudOnly = $tgrp.split(';')[3].split(',') ;
                verbose=$true ;
                whatif=$($whatif) ;
            } ; 
            write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):copy-XPermissionGroupToCloudOnly w`n$(($pltCXPermGrp|out-string).trim())" ; 
            copy-XPermissionGroupToCloudOnly @pltCXPermGrp ;
        } ; 
    Example demoing processing an array of descriptors, as a semicolon-delimited summary of inputs (useful for stacking bulk-creations)
    Schema for the $tgroups input is "[ticket];[source grp identifier];[owner];[members array]"
    .LINK
    https://github.com/tostka/verb-exo
    .LINK
    #>
    ###Requires -Version 5
    #Requires -Modules ActiveDirectory, MSOnline, AzureAD, ExchangeOnlineManagement, verb-AAD, verb-ADMS, verb-Ex2010, verb-Text
    #Requires -RunasAdministrator
    # VALIDATORS: [ValidateNotNull()][ValidateNotNullOrEmpty()][ValidateLength(24,25)][ValidateLength(5)][ValidatePattern("(lyn|bcc|spb|adl)ms6(4|5)(0|1).(china|global)\.ad\.toro\.com")][ValidateSet("USEA","GBMK","AUSYD")][ValidateScript({Test-Path $_ -PathType 'Container'})][ValidateScript({Test-Path $_})][ValidateRange(21,65)][ValidateCount(1,3)]
    ## [OutputType('bool')] # optional specified output type
    [CmdletBinding()]
    ##[Alias('somealias')]
    PARAM(
        [Parameter(Mandatory=$true,HelpMessage="ticket number[-ticket nnnnn]")]
        $ticket, 
        [Parameter(Mandatory=$true,HelpMessage="Name of on-prem replicated Exchange DistributionGroup to be copied to a cloud-only variant[-SourceGroupName somegroup]")]
        $SourceGroupName, 
        [Parameter(Mandatory=$true,HelpMessage="Identifier for the mailbox/mailuser object that will be the Owner of the new group[-Owner email@domain.com]")]
        $Owner, 
        [Parameter(Mandatory=$true,HelpMessage="Array of cloud-only unreplicated mailbox/mailuser designators to be added as members of the newly copied group[-MembersCloudOnly email@domain.com,email2@domain.com]")]
        [array]$MembersCloudOnly, 
        [Parameter(HelpMessage="Use EXOv2 (ExchangeOnlineManagement) over basic auth legacy connection [-useEXOv2]")]
        [switch] $useEXOv2,
        [Parameter(HelpMessage="Whatif Flag  [-whatIf]")]
        [switch] $whatIf=$true
    ) ;
    BEGIN{
        $Verbose = ($VerbosePreference -eq 'Continue') ; 
        $propsdg = 'SamAccountName','ManagedBy','AcceptMessagesOnlyFrom','AcceptMessagesOnlyFromDLMembers','AddressListMembership',
            'Alias','DisplayName','EmailAddresses','ExternalDirectoryObjectId','HiddenFromAddressListsEnabled','EmailAddressPolicyEnabled',
            'PrimarySmtpAddress','RecipientType','RecipientTypeDetails','WindowsEmailAddress','Name','DistinguishedName','WhenChanged','WhenCreated'; 
        connect-AD -Verbose:$false ; rx10 -Verbose:$false ; rxo  -Verbose:$false ; #cmsol  -Verbose:$false ;
    } 
    PROCESS{
        <# 
        For mail-anabled groups, it's a complete waste of time below using new-azureadgroup; set-azureadgroup; add-azureadgroupmember
        michev
            22686Reputation 1218Posts 0Following 38Followers
        answered · Apr 26 2021 at 7:04 AM
        You cannot use the Graph API for that. Mail-enabled security groups are authored in Exchange Online, and Graph currently has no support for Exchange admin operations. Use PowerShell instead.
        E.g. you have to use new-exodistributiongroup  for those. 
        #> 

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
            $smsg = "Resolving potential members:`n$(($MembersCloudOnly|out-string).trim())" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            $rmbrs = $MembersCloudOnly |foreach-object {get-exorecipient -id $_} | select -expand primarysmtpaddress ; 
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
            # set-azureadgroup -ObjectId -MailEnabled -MailNickName -ErrorAction
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
                    $pxDGm = get-exodistributiongroupmembers -id $pltNxDG.DisplayName ;
                } CATCH {
                    $smsg = "Failed processing $($_.Exception.ItemName). `nError Message: $($_.Exception.Message)`nError Details: $($_)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } 
                    else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    Continue ;
                } ;
                    
                $hMsg = @"

POST:exodistributiongroup

$(($pxdg|fl $propsdg|out-string).trim())

Members:
$(($pxDGm.PrimarySmtpAddress|out-string).trim())

Grant this group against EXO mailboxes using:

# EXO PERMS GRANT
`$whatif=`$true ;
`$GrantSplat=@{
    Identity="TARGETMAILBOX@toro.com" ;
    User='$($pxdg.primarysmtpaddress)' ;
    AccessRights="FullAccess"
};
add-exomailboxpermission @Grantsplat -whatif:`$(`$whatif) ;
Get-exoMailboxPermission -Identity `$GrantSplat.Identity |
    ?{(`$_.IsInherited -eq `$false) -AND !(`$_.user -match '^(S-\d-\d-\d{2}-\d{10}-\d{9}-\d{10}-\d{5}|NT\sAUTHORITY\\SELF)')}| 
    select User,AccessRights,IsInherited,Deny | out-string| 
    format-table -wrap ;

#EXO SENDAS GRANT
`$whatif=`$true ;
`$GrantSplat=@{
    Identity="TARGETMAILBOX@toro.com" ;
    trustee='$($pxdg.primarysmtpaddress)' ;
    AccessRights="SendAs" ;
    whatif=`$(`$whatif) ;
};
add-exoRecipientPermission @Grantsplat ;
get-exoRecipientPermission `$Grantsplat.identity|
    ?{(`$_.IsInherited -eq `$false) -AND !(`$_.trustee -match '^(S-\d-\d-\d{2}-\d{10}-\d{9}-\d{10}-\d{5}|NT\sAUTHORITY\\SELF)')} | 
    format-table -wrap trustee,AccessRights,IsInherited,Deny | out-string ;

"@ ; 
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
        #} ;  # loop-E
    }
    END{}
 } ;

#*------^ copy-XPermissionGroupToCloudOnly.ps1 ^------
