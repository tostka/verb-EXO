#*------v Function get-XOMailboxFolderList v------
function get-XOMailboxFolderList {
    <#
	.SYNOPSIS
	get-XOMailboxFolderList - Enumerates all user-accessible folders for the specified Exchange Online mailbox
	.NOTES
	Version     : 1.0.0
	Author      : Vasil Michev
	Website     : https://www.michev.info/blog/post/2500/how-to-reset-mailbox-folder-permissions
	Twitter     :	
	CreatedDate : 2022-06-15
	FileName    : get-XOMailboxFolderList.ps1
	License     : Not Asserted
	Copyright   : Not Asserted
	Github      : https://github.com/michevnew/PowerShell/blob/master/reset-XOMailboxAllFolderPerms.ps1
	Tags        : Powershell,ExchangeOnline,Mailbox,Delegate
	AddedCredit : Todd Kadrie
	AddedWebsite: http://www.toddomation.com
	AddedTwitter: @tostka / http://twitter.com/tostka
	REVISIONS
	* 12:48 PM 9/22/2023 revised (to shift into my verb-exo module for generic use): add/expand CBH; renam ReturnFolderList -> get-XOMailboxFolderList (alias orig name)
	* 6/15/22 vm posted version
	.DESCRIPTION
	get-XOMailboxFolderList - Enumerates all user-accessible folders for the specified Exchange Online mailbox
	PARAMETER SMTPAddress
	Smtp Address of mailbx to be processed
	.INPUTS
    SMTP address of the mailbox.
    .OUTPUTS
    Array with information about the mailbox folders.
	.EXAMPLE
	PS> $folders = get-XOMailboxFolderList -SMTPAddress email@domain.com ; 
	This command will return a list of all user-accessible folders for the specified email address
	.LINK
	https://github.com/tostka/verb-EXO
	#>
    [CmdletBinding()]
    [Alias('ReturnFolderList')]
    PARAM(
        [Parameter(Mandatory=$true, ValueFromPipeline=$true, HelpMessage="Smtp Address of mailbx to be processed")]
            $SMTPAddress
    ) ; 
    #$MBfolders = Invoke-Command -Session $session -ScriptBlock { Get-MailboxFolderStatistics $using:SMTPAddress | Select-Object Name,FolderType,Identity } -HideComputerName -ErrorAction Stop ; 
    # EOM3+ direct, no pssession support
    #$MBfolders = Get-xoMailboxFolderStatistics $SMTPAddress | Select-Object Name,FolderType,Identity -ErrorAction Stop ; 
    #*======v BP Wrapper for running EXO dynamic prefixed-EOM310+ cmdlets (psb-PsXoPrfx.cbp) v======
    # define the splat of all params:
    $pltGMFS = [ordered]@{Identity = $SMTPAddress ; erroraction = 'STOP'; verbose = $($VerbosePreference -eq "Continue") ;} ;
    $cmdlet = 'Get-MailboxFolderStatistics' ; $verb,$noun = $cmdlet.split('-') ;  #Spec cmdletname (VERB-NOUN), & split v/n
    TRY{$xoS = Get-ConnectionInformation -ErrorAction STOP }CATCH{reconnect-exo ; $xoS = Get-ConnectionInformation -ErrorAction STOP }
    TRY{
        if((-not $xos) -OR ($xoS | ?{$_.tokenstatus -notmatch 'Active|Expired' -AND $_.State -ne 'Connected'} )){reconnect-exo ; $xoS = Get-ConnectionInformation -ErrorAction STOP } ; 
        if($xos){
            $xcmd = "$verb-$($xoS.ModulePrefix)$noun `@pltGMFS" ; # build cmdline w splat, then echo:
            $smsg = "$($([regex]::match($xcmd,"^(\w+)-(\w+)" ).groups[0].value)) w`n$(($pltGMFS|out-string).trim())" ;
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            $MBfolders = invoke-expression $xcmd  | Select-Object Name,FolderType,Identity -ErrorAction Stop ; 
            if($MBfolders){write-verbose "(confirmed valid $([regex]::match($xcmd,"^(\w+)-(\w+)" ).groups[0].value) output)" ; }
        } else { 
            $smsg = "Missing `$xos EXO connection!" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
            throw $smsg ; BREAK ; 
        } 
    } CATCH {
        $ErrTrapd=$Error[0] ;
        $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
    } ; 
    <# version 12:43 PM 9/21/2023 moved cixo up to 1st, won't have prefix if not populated, also needs to fail/retry to ensure conn;  
    11:48 AM 9/20/2023 minor tweaks ; 3:01 PM 9/19/2023 initial 
    ## this runs: 1) connection status check, w rxo on demand; 2) splat wrapper with integrated prefix support; 3) try/catch on exec; 
    useful alias: cixo => get-connectioninformation;
    #>
    #*======^ END BP wrapper for running dynamic prefixed-EOM310+ cmdlets (psb-PsXoPrfx.cbp)  ^======
    $MBfolders = $MBfolders | ? {($_.FolderType -eq "User created" -or $_.FolderType -in $includedfolders) -and ($_.Name -notin $excludedfolders)} ; 
    if (-not $MBfolders) { return } 
    else { return ($MBfolders | select Name,FolderType,Identity) } ; 
} ; 
#*------^ END Function get-XOMailboxFolderList ^------