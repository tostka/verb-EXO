#*------v Function Get-xoMailboxFolderPermissionsRecursive v------
    function Get-xoMailboxFolderPermissionsRecursive {
        <#
        .SYNOPSIS
        Gets the current permissions for all user-accessible folders for a given mailbox.
        .NOTES
        Version     : 0.0.
        Author      : Todd Kadrie
        Website     : http://www.toddomation.com
        Twitter     : @tostka / http://twitter.com/tostka
        CreatedDate : 2023-
        FileName    : Get-xoMailboxFolderPermissionsRecursive.ps1
        License     : MIT License
        Copyright   : (c) 2023 Todd Kadrie
        Github      : https://github.com/tostka/verb-XXX
        Tags        : Powershell
        AddedCredit : REFERENCE
        AddedWebsite: URL
        AddedTwitter: URL
        REVISIONS
        * 1:45 PM 9/22/2023 ren Get-MailboxFolderPermissionsRecursive -> Get-xoMailboxFolderPermissionsRecursive (alias orig)
        * 10:47 AM 9/19/2023 rejigger to simply echo what it finds
        .DESCRIPTION
        The Get-xoMailboxFolderPermissionsRecursive cmdlet echoes permissions for all user-accessible folders for the given mailbox(es), specified via the -Mailbox parameter. The list of folders is generated via the get-XOMailboxFolderList function. Configure the $includedfolders and $excludedfolders variables to granularly control the folder list.
        .PARAMETER Mailbox
        Use the -Mailbox parameter to designate the mailbox. Any valid Exchange mailbox identifier can be specified. Multiple mailboxes can be specified in a comma-separated list or array, see examples below.
		.PARAMETER Quiet
		Switch to suppress outputs
        .EXAMPLE
        PS> $mbperms =  Get-xoMailboxFolderPermissionsRecursive -Ticket 999999 -Mailbox user@domain.com.com -OutVariable global:varFolderPermissionsFound ; 
        This command returns all permissions on all user-accessible folders in the user@domain.com mailbox, and tags the output .csv file with the specified ticket number.
        .EXAMPLE
        PS> $return = Get-xoMailboxFolderPermissionsRecursive -Mailbox @('emailaddress@domain.com','emailaddres2s@domain.com') -ResetDefaultLevel -verbose -whatif:$true
        Typical two-user pass as array, using specifying to include reset of all Default levels, with Whatif & verbose, assign output to $return
        .EXAMPLE
        PS> $return = Get-xoMailboxFolderPermissionsRecursive -Mailbox brad.stensrud@toro.com -ticket 760151 ; 
        PS> write-host "`n==Returned permission entries:" ; 
        PS> $return | ft -a ; 
        PS> write-host "==usertype distribution:" ; 
        PS> $return | group usertype |  ft -a count,name ; 
        PS> write-host "==output the subset of UNKNOWN usertype grants:" ; 
        PS> $return |?{$_.usertype -eq 'UNKNOWN'} | ft -a ; 
        PS> write-host "==username distribution:" ; 
        PS> $return |group user |  ft -a count,name ; 
        PS> write-host "==review UserType:Internal & username <> to mailbox owner:" ; 
        PS> write-verbose "derive 'owner' name:Should be the usertype:Internal w AccessRights:OWNER and highest frequency" ; 
        PS> $owner = $return | ?{$_.accessrights -like '*owner*' -AND $_.UserType -eq 'Internal'} | select -expand user | group | sort -desc | select -first 1 name | select -expand name ;
        PS> $return |?{$_.UserType -eq 'Internal' -AND $_.user -ne $owner} | ft -a ; 
        PS> write-host "==Non-Owner Grants:" ; 
        PS> $return |?{$_.user -ne $owner} | ft -a ; 
        PS> write-host "(count:$(($return |?{$_.user -ne $owner} |  measure | select -expand count|out-string).trim()))`n" ; 
        Typical single-user pass with ticket specified, assign output to $return, with range of post analysis examination of returned perm entries
        .EXAMPLE
        PS> write-verbose "Gather folder permissions from target mailbox" ; 
        PS> $mbfp = Get-xoMailboxFolderPermissionsRecursive -Ticket 999999 -Mailbox todd.kadrie@toro.com ; 
        PS> write-host "Echo returned folderperms to console, tabular" ; 
        PS> $mbfp | ft -a ; 
        PS> write-host "echo postfiltered broken-SID perms" ; 
        PS> $mbfp | ?{$_.user -match 'NT:S-1-5-21-'} | ft -a ; 
        PS> write-verbose "Run a removal of each of the broken-SID permissions" ; 
        PS> $mbfp | ?{$_.user -match 'NT:S-1-5-21-'} | %{ remove-xomailboxfolderpermission -id $_.foldername -user $_.user -whatif } ;
        Demo collecting all grants in target mailbox; reviewing return; post-filtering for broken SID entries; and then removing those entries with remove-xoMailboxFolderPermission, whatif pass is specified
        .INPUTS
        A mailbox identifier.
        .OUTPUTS
        Array of Mailbox address, Folder name and User.
        #>
        #Requires -Modules ActiveDirectory, ExchangeOnlineManagement, verb-Auth
        [cmdletbinding()]
        [Alias('Get-MailboxFolderPermissionsRecursive')]
        Param(
            [Parameter(Mandatory=$False,HelpMessage="Ticket Number [-Ticket '999999']")]
                [string]$Ticket,
            [Parameter(Mandatory=$true,ValueFromPipeline=$false,HelpMessage="Use the -Mailbox parameter to designate the mailbox. Any valid Exchange mailbox identifier can be specified. Multiple mailboxes can be specified in a comma-separated list or array, see examples below.")]
                [ValidateNotNullOrEmpty()]
                [Alias("Identity")]
                [String[]]$Mailbox,
            [Parameter(HelpMessage="Switch to suppress outputs")]
                [switch]$Quiet        
        ) ; 
        BEGIN{
            $includedfolders = @("Root","Inbox","Calendar", "Contacts", "DeletedItems", "Drafts", "JunkEmail", "Journal", 
                "Notes", "Outbox", "SentItems", "Tasks", "CommunicatorHistory", "Clutter", "Archive") ; 
            $Defaultfolders = @("Root","Inbox","Calendar", "Contacts", "DeletedItems", "SentItems", "Tasks") #Trimmed down list of default folders
            #Exclude additional Non-default folders created by Outlook or other mail programs. Folder NAMES, not types! So make sure to include translations too!
            #Exclude SearchDiscoveryHoldsFolder and SearchDiscoveryHoldsUnindexedItemFolder as they're not marked as default folders #Exclude "Calendar Logging" on older Exchange versions
            $excludedfolders = @("News Feed","Quick Step Settings","Social Activity Notifications","Suggested Contacts", 
                "SearchDiscoveryHoldsUnindexedItemFolder", "SearchDiscoveryHoldsFolder","Calendar Logging") ; 

            $prpADU = 'DistinguishedName','enabled','samaccountname','sid','UserPrincipalName' ; 
            $DefaultRoleUserNames = @("Default","Anonymous","Owner@local","Member@local") ; 
            if(!$ThrottleMs){$ThrottleMs = 500} ; 
            if( -not (get-variable -name PSScriptRoot -ea 0) -OR ($PSScriptRoot -eq '') -OR ($PSScriptRoot -eq $null)){
            if ($psISE) { $ScriptName = $psISE.CurrentFile.FullPath } 
            elseif($psEditor){
                if ($context = $psEditor.GetEditorContext()) {$ScriptName = $context.CurrentFile.Path } 
            } elseif ($host.version.major -lt 3) {
                $ScriptName = $MyInvocation.MyCommand.Path ;
                $PSScriptRoot = Split-Path $ScriptName -Parent ;
                $PSCommandPath = $ScriptName ;
            } else {
                if ($MyInvocation.MyCommand.Path) {
                    $ScriptName = $MyInvocation.MyCommand.Path ;
                    $PSScriptRoot = Split-Path $MyInvocation.MyCommand.Path -Parent ;
                } else {throw "UNABLE TO POPULATE SCRIPT PATH, EVEN `$MyInvocation IS BLANK!" } ;
            };
            if($ScriptName){
                $ScriptDir = Split-Path -Parent $ScriptName ;
                $ScriptBaseName = split-path -leaf $ScriptName ;
                $ScriptNameNoExt = [system.io.path]::GetFilenameWithoutExtension($ScriptName) ;
            } ; 
        } else {
            if($PSScriptRoot){$ScriptDir = $PSScriptRoot ;}
            else{
                write-warning "Unpopulated `$PSScriptRoot!" ; 
                $ScriptDir=(Split-Path -parent $MyInvocation.MyCommand.Definition) + "\" ;
            }
            if ($PSCommandPath) {$ScriptName = $PSCommandPath } 
            else {
                $ScriptName = $myInvocation.ScriptName
                $PSCommandPath = $ScriptName ;
            } ;
            $ScriptBaseName = (Split-Path -Leaf ((& { $myInvocation }).ScriptName))  ;
            $ScriptNameNoExt = [system.io.path]::GetFilenameWithoutExtension($MyInvocation.InvocationName) ;
        } ;
        if(-not $ScriptDir){
            write-host "Failed `$ScriptDir resolution on PSv$($host.version.major): Falling back to $MyInvocation parsing..." ; 
            $ScriptDir=(Split-Path -parent $MyInvocation.MyCommand.Definition) + "\" ;
            $ScriptBaseName = (Split-Path -Leaf ((&{$myInvocation}).ScriptName))  ; 
            $ScriptNameNoExt = [system.io.path]::GetFilenameWithoutExtension($MyInvocation.InvocationName) ;     
        } else {
            if(-not $PSCommandPath ){
                $PSCommandPath  = $ScriptName ; 
                if($PSCommandPath){ write-host "(Derived missing `$PSCommandPath from `$ScriptName)" ; } ;
            } ; 
            if(-not $PSScriptRoot  ){
                $PSScriptRoot   = $ScriptDir ; 
                if($PSScriptRoot){ write-host "(Derived missing `$PSScriptRoot from `$ScriptDir)" ; } ;
            } ; 
        } ; 
        if(-not ($ScriptDir -AND $ScriptBaseName -AND $ScriptNameNoExt)){ 
            throw "Invalid Invocation. Blank `$ScriptDir/`$ScriptBaseName/`ScriptNameNoExt" ; 
            BREAK ; 
        } ; 

        $smsg = "`$ScriptDir:$($ScriptDir)" ;
        $smsg += "`n`$ScriptBaseName:$($ScriptBaseName)" ;
        $smsg += "`n`$ScriptNameNoExt:$($ScriptNameNoExt)" ;
        $smsg += "`n`$PSScriptRoot:$($PSScriptRoot)" ;
        $smsg += "`n`$PSCommandPath:$($PSCommandPath)" ;  ;
        write-verbose $smsg ; 
        #endregion ENVIRO_DISCOVER ; #*------^ END ENVIRO_DISCOVER ^------

            Write-Verbose "Parsing the Mailbox parameter..."
            $SMTPAddresses = @{}
            foreach ($mb in $Mailbox) {
                Start-Sleep -Milliseconds 80 #Add some delay to avoid throttling...
                #*======v BP Wrapper for running EXO dynamic prefixed-EOM310+ cmdlets (psb-PsXoPrfx.cbp) v======
                # define the splat of all params:
                $pltGMbx = [ordered]@{identity =  $mb ; erroraction = 'STOP'; verbose = $($VerbosePreference -eq "Continue") ;} ;
                $cmdlet = 'get-Mailbox' ; $verb,$noun = $cmdlet.split('-') ;  #Spec cmdletname (VERB-NOUN), & split v/n
                TRY{$xoS = Get-ConnectionInformation -ErrorAction STOP }CATCH{reconnect-exo ; $xoS = Get-ConnectionInformation -ErrorAction STOP }
                TRY{
                    if((-not $xos) -OR ($xoS | ?{$_.tokenstatus -notmatch 'Active|Expired' -AND $_.State -ne 'Connected'} )){reconnect-exo ; $xoS = Get-ConnectionInformation -ErrorAction STOP } ; 
                    if($xos){
                        $xcmd = "$verb-$($xoS.ModulePrefix)$noun `@pltGMbx" ; # build cmdline w splat, then echo:
                        $smsg = "$($([regex]::match($xcmd,"^(\w+)-(\w+)" ).groups[0].value)) w`n$(($pltGMbx|out-string).trim())" ;
                        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                        $SMTPAddress  = invoke-expression $xcmd  | 
                            Select-Object -ExpandProperty PrimarySmtpAddress -ErrorAction SilentlyContinue;
                        if($SMTPAddress){write-verbose "(confirmed valid $([regex]::match($xcmd,"^(\w+)-(\w+)" ).groups[0].value) output)" ; }
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
                if (-not $SMTPAddress) { if (-not $Quiet) { 
                    $smsg = "Mailbox with identifier $mb not found, skipping..." }; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                    continue 
                } elseif (($SMTPAddress.count -gt 1) -or ($SMTPAddresses[$mb]) -or ($SMTPAddresses.ContainsValue($SMTPAddress))) { 
                    $smsg = "Multiple mailboxes matching the identifier $mb found, skipping..."; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                    continue 
                }else { $SMTPAddresses[$mb] = $SMTPAddress } ; 
            }
            if (-not $SMTPAddresses -or ($SMTPAddresses.Count -eq 0)) { 
                Throw "No matching mailboxes found, check the parameter values." 
            } ; 
            $smsg = "The following list of mailboxes will be used:$($SMTPAddresses.values  -join ", ")" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            $smsg = "List of default folder TYPES that will be used:$($includedfolders  -join ", ")" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            $smsg = "List of folder NAMES that will be excluded:$($excludedfolders  -join ", ")" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

        } ; # BEG-E
        PROCESS {
            $out = @() ; 
            foreach ($smtp in $SMTPAddresses.Values) {
                $sBnrS = $smsg ="`n#*------v PROCESSING Mailbox: $($smtp)... v------" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                Start-Sleep -Milliseconds $ThrottleMs  ; #Add some delay to avoid throttling...
                Write-Verbose "Obtaining folder list for mailbox ""$smtp""..." ; 
                $folders = get-XOMailboxFolderList $smtp ; 
                Write-Verbose "A total of $($folders.count) folders found for $($smtp)." ; 
                if (-not $folders) { 
                    $smsg ="No matching folders found for $($smtp), skipping..." ; 
                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                    continue 
                } ; 
                #Cycle over each folder we are interested in
                foreach ($folder in $folders) {
                    $smsg = "`n==PROCESSING:$($folder.name)`n" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Prompt } 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                    #"Fix" for folders with "/" characters, treat the Root folder separately
                    if ($folder.FolderType -eq "Root") { $foldername = $smtp }
                    else { $foldername = $folder.Identity.ToString().Replace([char]63743,"/").Replace($smtp,$smtp + ":") } ; 
                    
                    $fPermissions = get-XOMailboxFolderPermissionsSummary $foldername
                    if (-not $ResetDefaultLevel) { $fPermissions = $fPermissions | ? {$_.UserName -notin @("Default","Anonymous","Owner@local","Member@local")}}  ; #filter out default permissions
                    if (-not $fPermissions) { Write-Verbose "No permission entries found for $($foldername), skipping..." ; continue } ; 
                    #echo the folder permissions for each delegate
                    foreach ($u in $fPermissions) {
                        write-host "`n" ; 
                        if ($u.UserType -eq "Default") {
                            #UserType enumeration https://docs.microsoft.com/en-us/previous-versions/office/developer/exchange-server-2010/ff319704(v%3Dexchg.140)
                            #if ($ResetDefaultLevel) {
                                <# https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_scopes?view=powershell-7.3#the-using-scope-modifier
                                    scope $using: - Used to access variables defined in another scope while running scripts via cmdlets like Start-Job and Invoke-Command.

                                #>
                                TRY {
                                    #write-host -foregroundcolor yellow "Resetting permissions on ""$foldername"" for principal ""Default""." ;
                                    if ($folder.FolderType -eq "Calendar") {
                                        $smsg = "'Default:Calendar entry':" ;
                                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                        #if (($u.AccessRights -join ",") -ne "AvailabilityOnly") {
                                        # TTC customizes the view as LimitedDetails: $CalPermsDefault
                                        if (($u.AccessRights -join ",") -ne $CalPermsDefault) {
                                            #Invoke-Command -Session $session -ScriptBlock { Set-MailboxFolderPermission -Identity $Using:foldername -User Default -AccessRights AvailabilityOnly -WhatIf:$using:WhatIfPreference -Confirm:$false } -ErrorAction Stop -HideComputerName ;
                                            # can't use -session $session with EOM3+, try direct calls; should work
                                            #Set-xoMailboxFolderPermission -Identity $foldername -User Default -AccessRights $CalPermsDefault -WhatIf:$WhatIfPreference -Confirm:$false -ErrorAction Stop ;
                                        } else { continue } ; 
                                        #$outtemp = New-Object psobject -Property ([ordered]@{"Mailbox" = $smtp;"FolderName" = $folder.name;"User" = $u.UserName;"AccessRights" = "AvailabilityOnly"}) ; 
                                        $outtemp = New-Object psobject -Property ([ordered]@{"Mailbox" = $smtp;"FolderName" = $u.identity;"User" = $u.UserName ; 
                                            UserType = $u.UserType ; 
                                            AccessRights = $u.AccessRights ; 
                                            SharingPermissionFlags = $u.SharingPermissionFlags ; 
                                        }) ;
                                        $out += $outtemp; 
                                        if (-not $Quiet ) { 
                                            #$outtemp | ft -a 
                                            $smsg = "`n$(($outtemp | ft -a |out-string).trim())" ; 
                                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                        }  ;
                                    } else {
                                        if (($u.AccessRights -join ",") -ne "None") {
                                            $smsg = "'Default:non-NONE entry':" ;
                                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                            #Invoke-Command -Session $session -ScriptBlock { Set-MailboxFolderPermission -Identity $Using:foldername -User Default -AccessRights None -WhatIf:$using:WhatIfPreference -Confirm:$false } -ErrorAction Stop -HideComputerName 
                                            # eom3+ no pss
                                            #Set-xoMailboxFolderPermission -Identity $Using:foldername -User Default -AccessRights None -WhatIf:$using:WhatIfPreference -Confirm:$false -ErrorAction Stop ;
                                            # dump these
                                            $smsg = "'non-NONE entry':" ;
                                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                            #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                                        } else { continue } ; 
                                        $outtemp = New-Object psobject -Property ([ordered]@{"Mailbox" = $smtp;"FolderName" = $u.identity;"User" = $u.UserName ; 
                                            UserType = $u.UserType ; 
                                            AccessRights = $u.AccessRights ; 
                                            SharingPermissionFlags = $u.SharingPermissionFlags ; 
                                        }) ;
                                        # echo, don't dump, the END is emitting a full obj stack
                                        $out += $outtemp; 
                                        if (-not $Quiet ) { 
                                            #$outtemp | ft -a 
                                            $smsg = "`n$(($outtemp | ft -a |out-string).trim())" ; 
                                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                        }  ;
                                    } ; 
                                    #$out += $outtemp; if (-not $Quiet -and -not $WhatIfPreference) { $outtemp }  ; #Write output to the console unless the -Quiet parameter is used
                                } CATCH {$_ | fl * -Force; continue}  ; #catch-all for any unhandled errors
                            #} else { continue } ; 
                        } elseif ($u.UserType -eq "Anonymous") { 
                            # continue #Maybe set them all to none when $resetdefault is used?
                            $smsg = "'Anonymous entry':" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                            $outtemp = New-Object psobject -Property ([ordered]@{"Mailbox" = $smtp;"FolderName" = $u.identity;"User" = $u.UserName ; 
                                UserType = $u.UserType ; 
                                AccessRights = $u.AccessRights ; 
                                SharingPermissionFlags = $u.SharingPermissionFlags ; 
                            }) ;
                            $out += $outtemp; 
                            if (-not $Quiet ) { 
                                #$outtemp | ft -a 
                                $smsg = "`n$(($outtemp | ft -a |out-string).trim())" ; 
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            }  ;
                        } elseif ($u.UserType -eq "Unknown") { 
                            #write-host -foregroundcolor yellow "Skipping orphaned permissions entry: $($u.UserName)"; continue 
                            # actually on reviews, we *want* to see and dump the orphan/corrupt entries:
                            $smsg = "'UNKNOWN entry':" ; 
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                            # may as well attempt a gadu resolve on the SID
                            if($u.UserName -match '^NT:S-'){
                                $smsg = "(entry UserName appears to be a BROKEN SID (SECURITY IDENTIFYER == DELETED USER OBJECT)" ; 
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                                write-verbose "(attempting: get-aduser -id $($u.UserName.replace('NT:','')) )" ; 
                                TRY{
                                   if($ADU =  get-aduser -id ($u.UserName.replace('NT:','')) -ErrorAction STOP){
                                        $smsg = "Resolved`n $($u.UserName.replace('NT:',''))`n to an existing ADUser object:`n$(($adu | fl $prpADU |out-string).trim())" ; 
                                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                                   }else {
                                        write-warning "Unable to resolve $($u.UserName.replace('NT:','')) to an existing ADUser object (likely deleted TERM)"
                                   } ; 
                           
                                }CATCH{ write-warning "Unable to resolve $($u.UserName.replace('NT:','')) to an existing ADUser object (likely deleted TERM)" }
                            }
                            $outtemp = New-Object psobject -Property ([ordered]@{"Mailbox" = $smtp;"FolderName" = $u.identity;"User" = $u.UserName ; 
                                UserType = $u.UserType ; 
                                AccessRights = $u.AccessRights ; 
                                SharingPermissionFlags = $u.SharingPermissionFlags ; 
                            }) ;
                            #$out += $outtemp; if (-not $Quiet ) { $outtemp | ft -a }  ;
                            # echo, don't dump, the END is emitting a full obj stack
                            $out += $outtemp; 
                            if (-not $Quiet ) { 
                                #$outtemp | ft -a 
                                $smsg = "`n$(($outtemp | ft -a |out-string).trim())" ; 
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            }  ;
                        } elseif ($u.UserType -eq "Internal") { 
                            #write-host -foregroundcolor yellow "Skipping orphaned permissions entry: $($u.UserName)"; continue 
                            # actually on reviews, we *want* to see and dump the entries:
                            $smsg = "'Internal entry':" ; 
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                            $outtemp = New-Object psobject -Property ([ordered]@{"Mailbox" = $smtp;"FolderName" = $u.identity;"User" = $u.UserName ; 
                                UserType = $u.UserType ; 
                                AccessRights = $u.AccessRights ; 
                                SharingPermissionFlags = $u.SharingPermissionFlags ; 
                            }) ;
                            $out += $outtemp; 
                            if (-not $Quiet ) { 
                                #$outtemp | ft -a 
                                $smsg = "`n$(($outtemp | ft -a |out-string).trim())" ; 
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            }  ;
                        } else {
                            if ($u.UserType -eq "External") { $u.User = $u.UserName }
                            TRY {
                                #if (-not $u.User) { continue } ; 
                                #Invoke-Command -Session $session -ScriptBlock { Remove-MailboxFolderPermission -Identity $Using:foldername -User $Using:u.User -WhatIf:$using:WhatIfPreference -Confirm:$false } -ErrorAction Stop -HideComputerName ;
                                # eom3+ no pss
                                #Remove-xoMailboxFolderPermission -Identity $foldername -User $u.User -WhatIf:$WhatIfPreference -Confirm:$false -ErrorAction Stop ;
                                $smsg = "'non-NONE entry':" ;
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                                $outtemp = New-Object psobject -Property ([ordered]@{"Mailbox" = $smtp;"FolderName" = $u.identity;"User" = $u.UserName ; 
                                    UserType = $u.UserType ; 
                                    AccessRights = $u.AccessRights ; 
                                    SharingPermissionFlags = $u.SharingPermissionFlags ; 
                                }) ;
                                $out += $outtemp; 
                                if (-not $Quiet ) { 
                                    #$outtemp | ft -a 
                                    $smsg = "`n$(($outtemp | ft -a |out-string).trim())" ; 
                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                }  ;
                            } CATCH [System.Management.Automation.RemoteException] {
                                if (-not $Quiet) {
                                    if ($_.CategoryInfo.Reason -eq "UserNotFoundInPermissionEntryException") { Write-Host "WARNING: No existing permissions entry found on ""$foldername"" for principal ""$($u.UserName)""" -ForegroundColor Yellow }
                                    elseif ($_.CategoryInfo.Reason -eq "CannotChangePermissionsOnFolderException") { Write-Host "ERROR: Folder permissions for ""$foldername"" CANNOT be changed!" -ForegroundColor Red }
                                    elseif ($_.CategoryInfo.Reason -eq "CannotRemoveSpecialUserException") { Write-Host "ERROR: Folder permissions for ""$($u.UserName)"" CANNOT be changed!" -ForegroundColor Red }
                                    elseif ($_.CategoryInfo.Reason -eq "ManagementObjectNotFoundException") { Write-Host "ERROR: Folder ""$foldername"" not found, this should not happen..." -ForegroundColor Red }
                                    elseif ($_.CategoryInfo.Reason -eq "InvalidInternalUserIdException") { Write-Host "ERROR: ""$($u.UserName)"" is not a valid security principal for folder-level permissions..." -ForegroundColor Red }
                                    else {$_ | fl * -Force; continue}  ; #catch-all for any unhandled errors
                                } ;  # if-E !quiet
                            } catch {$_ | fl * -Force; continue} ;#catch-all for any unhandled errors
                        } # if-E
                    }  ; # ACE loop-E
                } ;  # FOLDERS loop-E
                $smsg = $sBnrS.replace('-v','-^').replace('v-','^-') ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
            }  # MBX loop-E
        } ;  # PROC-E
        END{
            if ($out) {
                #$out | Export-Csv -Path "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_MailboxFolderPermissionsRemoved.csv" -NoTypeInformation -Encoding UTF8 -UseCulture ;
                #$opath = "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_MailboxFolderPermissionsRemoved.csv" ; 
                #write-host "Exporting results to the CSV file...`n$($opath)" ;
                #$out | Export-Csv -Path $opath -NoTypeInformation -Encoding UTF8 -UseCulture ;
                [string]$opath = $null ; 
                if($ticket){$opath += "$($TICKET)-" }
                # $opath += "$($item)_MailboxFolderPermissionsRemoved-$(get-date -format 'yyyyMMdd-HHmmtt').csv" ; 
                if(($SMTPAddresses.Values |  measure | select -expand count ) -gt 3){
                    $opath += "$($SMTPAddresses.Values[0]),xxx_MailboxFolderPermissionsRemoved-$(get-date -format 'yyyyMMdd-HHmmtt').csv" ; 
                } else { 
                    $opath += "$($SMTPAddresses.Values -join ',')_MailboxFolderPermissionsRemoved-$(get-date -format 'yyyyMMdd-HHmmtt').csv" ; 
                } ; 
                $oPath = join-path -path (join-path -path $ScriptDir -childpath "logs") -ChildPath $opath ; 
                write-host "Exporting results to the CSV file...`n$($opath)" ;
                $out | Export-Csv -Path $opath -NoTypeInformation -Encoding UTF8 -UseCulture ;
                #Write output to the console unless the -Quiet parameter is used
                #if (-not $Quiet -and -not $WhatIfPreference) { return $out | Out-Default }  ; 
                if (-not $Quiet -and -not $WhatIfPreference) { return $out  }  ; # above is returning as an array of text with no fields; output the object and aggregate it
            } else { write-host -foregroundcolor yellow "Output is empty, skipping the export to CSV file..." } ;
            Write-Verbose "Finish..." ;
        } ; 
    } ;
    #*------^ END Function Get-xoMailboxFolderPermissionsRecursive ^------