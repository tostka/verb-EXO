#*------v Function get-XOMailboxFolderPermissionsSummary v------
    function get-XOMailboxFolderPermissionsSummary {
        <#
	    .SYNOPSIS
	    get-XOMailboxFolderPermissionsSummary - Enumerates all permissions for the given  Exchange Online mailbox folder
	    .NOTES
	    Version     : 1.0.0
	    Author      : Vasil Michev
	    Website     : https://www.michev.info/blog/post/2500/how-to-reset-mailbox-folder-permissions
	    Twitter     :	
	    CreatedDate : 2022-06-15
	    FileName    : get-XOMailboxFolderPermissionsSummary.ps1
	    License     : Not Asserted
	    Copyright   : Not Asserted
	    Github      : https://github.com/michevnew/PowerShell/blob/master/reset-XOMailboxAllFolderPerms.ps1
	    Tags        : Powershell,ExchangeOnline,Mailbox,Delegate
	    AddedCredit : Todd Kadrie
	    AddedWebsite: http://www.toddomation.com
	    AddedTwitter: @tostka / http://twitter.com/tostka
	    REVISIONS
	    * 12:48 PM 9/22/2023 revised (to shift into my verb-exo module for generic use): add/expand CBH; renam GetFolderPermissions -> get-XOMailboxFolderPermissionsSummary (alias orig name)
	    * 6/15/22 vm posted version
	    .DESCRIPTION
	    Enumerates all permissions for the given  Exchange Online mailbox folder
	    .PARAMETER foldername
	    Identifier of the target folder, expressed in 'email@domain.com:\folderpath' format
	    ..INPUTS
        Identifier for the folder.
        .OUTPUTS
        Array with information about the mailbox folder permissions.
	    .EXAMPLE
	    PS> $perms = get-XOMailboxFolderPermissionsSummary user@domain.com:\Calendar ; 
	    This command will return a list of all user-accessible folders for the specified email address
	    .LINK
	    https://github.com/tostka/verb-EXO
	    #>
        #Requires -Modules ExchangeOnlineManagement, verb-Auth
	    [CmdletBinding()]
	    [Alias('GetFolderPermissions')]
        PARAM(
            [Parameter(Mandatory=$true, ValueFromPipeline=$true, HelpMessage="Identifier of the target folder, expressed in 'email@domain.com:\folderpath' format")]
            $foldername
        ) ; 
        $prpFldrPerm = 'Identity','User','AccessRights','SharingPermissionFlags' ; 
        $prpFldrPermLeaf = 'Identity',
            @{n="User";e={$_.User.RecipientPrincipal.Guid.Guid}},
            @{n="UserType";e={$_.User.UserType.ToString()}},
            @{n="UserName";e={$_.User.DisplayName}},
            'AccessRights','SharingPermissionFlags' ; 

        # eom3+ no pssession supp
        #$FolderPerm = Get-xoMailboxFolderPermission $foldername | Select-Object Identity,User,AccessRights,SharingPermissionFlags -ErrorAction Stop |
        #     select Identity,@{n="User";e={$_.User.RecipientPrincipal.Guid.Guid}},@{n="UserType";e={$_.User.UserType.ToString()}},@{n="UserName";e={$_.User.DisplayName}},AccessRights,SharingPermissionFlags ; 
        #*======v BP Wrapper for running EXO dynamic prefixed-EOM310+ cmdlets (psb-PsXoPrfx.cbp) v======
        # define the splat of all params:
        $pltGMFP = [ordered]@{identity = $foldername ;erroraction = 'STOP'; verbose = $($VerbosePreference -eq "Continue") ;} ;
        $cmdlet = 'Get-MailboxFolderPermission' ; $verb,$noun = $cmdlet.split('-') ;  #Spec cmdletname (VERB-NOUN), & split v/n
        TRY{$xoS = Get-ConnectionInformation -ErrorAction STOP }CATCH{reconnect-exo ; $xoS = Get-ConnectionInformation -ErrorAction STOP }
        TRY{
            if((-not $xos) -OR ($xoS | ?{$_.tokenstatus -notmatch 'Active|Expired' -AND $_.State -ne 'Connected'} )){reconnect-exo ; $xoS = Get-ConnectionInformation -ErrorAction STOP } ; 
            if($xos){
                $xcmd = "$verb-$($xoS.ModulePrefix)$noun `@pltGMFP" ; # build cmdline w splat, then echo:
                $smsg =  "$($([regex]::match($xcmd,"^(\w+)-(\w+)" ).groups[0].value)) w`n$(($pltGMFP|out-string).trim())" ;
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                $FolderPerm = invoke-expression $xcmd  | Select-Object $prpFldrPerm -ErrorAction Stop | 
                    Select-Object $prpFldrPermLeaf ; 
                if($FolderPerm){write-verbose "(confirmed valid $([regex]::match($xcmd,"^(\w+)-(\w+)" ).groups[0].value) output)" ; }
            } else { 
                $smsg = "Missing `$xos EXO connection!" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                throw $smsg ; BREAK ; 
            } 
        } CATCH {
            $ErrTrapd=$Error[0] ;
            $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
            write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" ;
        } ; 
        <# version 12:43 PM 9/21/2023 moved cixo up to 1st, won't have prefix if not populated, also needs to fail/retry to ensure conn;  
        11:48 AM 9/20/2023 minor tweaks ; 3:01 PM 9/19/2023 initial 
        ## this runs: 1) connection status check, w rxo on demand; 2) splat wrapper with integrated prefix support; 3) try/catch on exec; 
        useful alias: cixo => get-connectioninformation;
        #>
        #*======^ END BP wrapper for running dynamic prefixed-EOM310+ cmdlets (psb-PsXoPrfx.cbp)  ^======
        if (-not $FolderPerm) { return }
        else { return $FolderPerm }
    } ;
    #*------^ END Function get-XOMailboxFolderPermissionsSummary ^------