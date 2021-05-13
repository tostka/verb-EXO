#*------v test-xoMailbox.ps1 v------
Function test-xoMailbox {
    <#
    .SYNOPSIS
    test-xoMailbox.ps1 - Run quick mailbox function validation and metrics gathering, to streamline, 'Is it working?' questions. 
    .NOTES
    Version     : 10.0
    Author      : Todd Kadrie
    Website     : http://www.toddomation.com
    Twitter     : @tostka / http://twitter.com/tostka
    CreatedDate : 2021-04-21
    FileName    : test-xoMailbox.ps1
    License     : MIT License
    Copyright   : (c) 2020 Todd Kadrie
    Github      : https://github.com/tostka/verb-XXX
    Tags        : Powershell,ExchangeOnline,Exchange,Resource,MessageTrace
    REVISIONS
    # 12:27 PM 5/11/2021 updated to cover Room|Shared|Equipment mbx types, along with UserMailbox
    # 3:47 PM 5/6/2021 recoded start-log switching, was writing logs into AllUsers profile dir ; swapped out 'Fail' msgtrace expansion, and have it do all non-Delivery statuses, up to the $MsgTraceNonDeliverDetailsLimit = 10; tested, appears functional
    # 3:15 PM 5/4/2021 added trailing |?{$_.length} bug workaround for get-gcfastxo.ps1
    * 4:20 PM 4/29/2021 debugged, looks functional - could benefit from moving the msgtrk summary down into the output block, but [shrug]
    * 7:56 AM 4/28/2021 init
    .DESCRIPTION
    test-xoMailbox.ps1 - Run quick mailbox function validation and metrics gathering, to streamline, 'Is it working?' questions. 
    .PARAMETER Mailboxes
    Array of Mailbox email addresses [-Mailboxes mbx1@domain.com]
    .PARAMETER Ticket
    Ticket # [-Ticket nnnnn]
    .PARAMETER useEXOv2
    Use EXOv2 (ExchangeOnlineManagement) over basic auth legacy connection [-useEXOv2]
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    .EXAMPLE
    test-xoMailbox.ps1 -TenOrg TOR -Mailboxes 'Fname.LName@domain.com','FName2.Lname2@domain.com' -Ticket 610706 -verbose ;
    .EXAMPLE
    .\test-xoMailbox.ps1
    .LINK
    https://github.com/tostka/verb-EXO
    #>
    #Requires -Modules ActiveDirectory,verb-Auth,verb-IO,verb-Mods,verb-Text,verb-Network,verb-AAD,verb-ADMS,verb-Ex2010,verb-logging
    [CmdletBinding()]
    PARAM(
        [Parameter(Mandatory=$FALSE,HelpMessage="TenantTag value, indicating Tenants to connect to[-TenOrg 'TOL']")]
        [ValidateNotNullOrEmpty()]
        $TenOrg = 'TOR',
        [Parameter(Position=0,Mandatory=$True,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,HelpMessage="Array of EXO Mailbox identifiers to be processed [-mailboxes 'xxx','yyy']")]
        [ValidateNotNullOrEmpty()]$Mailboxes,
        [Parameter(Mandatory = $True, HelpMessage = "Ticket # [-Ticket nnnnn]")]
        [array]$Tickets,
        [Parameter(Position=0,Mandatory=$False,HelpMessage="Specific ResourceDelegate address to be confirmed for detailed delivery [emailaddr]")]
        [ValidateNotNullOrEmpty()][string]$TargetDelegate,
        [Parameter(Mandatory=$false,HelpMessage="Days back to search (defaults to 10)[-Days 10]")]
        [int]$Days=10,
        [Parameter(HelpMessage="Use EXOv2 (ExchangeOnlineManagement) over basic auth legacy connection [-useEXOv2]")]
        [switch] $useEXOv2
    ) ;
    BEGIN {
        $Verbose = ($VerbosePreference -eq 'Continue') ;
        #*======v SCRIPT/DOMAIN/MACHINE/INITIALIZATION-DECLARE-BOILERPLATE v======
        # SCRIPT-CONFIG MATERIAL TO SET THE UNDERLYING $DBGPREF:
        if ($ShowDebug) { $DebugPreference = "Continue" ; write-debug "(`$showDebug:$showDebug ;`$DebugPreference:$DebugPreference)" ; };
        if ($Whatif) { write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):`$Whatif is TRUE (`$whatif:$($whatif))" ; };
        # If using WMI calls, push any cred into WMI:
        #if ($Credential -ne $Null) {$WmiParameters.Credential = $Credential }  ;

        <# scriptname with extension
        $ScriptDir = (Split-Path -parent $MyInvocation.MyCommand.Definition) + "\" ;
        $ScriptBaseName = (Split-Path -Leaf ((& { $myInvocation }).ScriptName))  ;
        $ScriptNameNoExt = [system.io.path]::GetFilenameWithoutExtension($MyInvocation.InvocationName) ;
        #>
        # 11:24 AM 6/21/2019 UPDATED HYBRID VERS
        if($showDebug){
            write-host -foregroundcolor green "`SHOWDEBUG: `$PSScriptRoot:$($PSScriptRoot)`n`$PSCommandPath:$($PSCommandPath)" ;
        } ;
        if ($psISE){
                $ScriptDir = Split-Path -Path $psISE.CurrentFile.FullPath ;
                $ScriptBaseName = split-path -leaf $psise.currentfile.fullpath ;
                $ScriptNameNoExt = [system.io.path]::GetFilenameWithoutExtension($psise.currentfile.fullpath) ;
                        $PSScriptRoot = $ScriptDir ;
                if($PSScriptRoot -ne $ScriptDir){ write-warning "UNABLE TO UPDATE BLANK `$PSScriptRoot TO CURRENT `$ScriptDir!"} ;
                $PSCommandPath = $psise.currentfile.fullpath ;
                if($PSCommandPath -ne $psise.currentfile.fullpath){ write-warning "UNABLE TO UPDATE BLANK `$PSCommandPath TO CURRENT `$psise.currentfile.fullpath!"} ;
        } else {
            if($host.version.major -lt 3){
                $PSScriptRoot = Split-Path $MyInvocation.MyCommand.Path -Parent ;
                $PSCommandPath = $myInvocation.ScriptName ;
                $ScriptBaseName = (Split-Path -Leaf ((&{$myInvocation}).ScriptName))  ;
                $ScriptNameNoExt = [system.io.path]::GetFilenameWithoutExtension($MyInvocation.InvocationName) ;
            } elseif($PSScriptRoot) {
                $ScriptDir = $PSScriptRoot ;
                if($PSCommandPath){
                    $ScriptBaseName = split-path -leaf $PSCommandPath ;
                    $ScriptNameNoExt = [system.io.path]::GetFilenameWithoutExtension($PSCommandPath) ;
                } else {
                    $PSCommandPath = $myInvocation.ScriptName ;
                    $ScriptBaseName = (Split-Path -Leaf ((&{$myInvocation}).ScriptName))  ;
                    $ScriptNameNoExt = [system.io.path]::GetFilenameWithoutExtension($MyInvocation.InvocationName) ;
                } ;
            } else {
                if($MyInvocation.MyCommand.Path) {
                    $PSScriptRoot = Split-Path $MyInvocation.MyCommand.Path -Parent ;
                    $ScriptBaseName = (Split-Path -Leaf ((&{$myInvocation}).ScriptName))  ;
                    $ScriptNameNoExt = [system.io.path]::GetFilenameWithoutExtension($MyInvocation.InvocationName) ;
                } else {
                    throw "UNABLE TO POPULATE SCRIPT PATH, EVEN `$MyInvocation IS BLANK!" ;
                } ;
            } ;
        } ;
        if($showDebug){
            write-host -foregroundcolor green "`SHOWDEBUG: `$ScriptDir:$($ScriptDir)`n`$ScriptBaseName:$($ScriptBaseName)`n`$ScriptNameNoExt:$($ScriptNameNoExt)`n`$PSScriptRoot:$($PSScriptRoot)`n`$PSCommandPath:$($PSCommandPath)" ;
        } ;
    
        $ComputerName = $env:COMPUTERNAME ;
        $sQot = [char]34 ; $sQotS = [char]39 ;
        $NoProf = [bool]([Environment]::GetCommandLineArgs() -like '-noprofile'); # if($NoProf){# do this};
        $MyBox = "LYN-3V6KSY1", "TIN-BYTEIII", "TIN-BOX", "TINSTOY", "LYN-8DCZ1G2" ;
        $DomainWork = "TORO";
        $DomHome = "REDBANK";
        $DomLab = "TORO-LAB";
        #$ProgInterval= 500 ; # write-progress wait interval in ms
        # 12:23 PM 2/20/2015 add gui vb prompt support
        #[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null ;
        # 11:00 AM 3/19/2015 should use Windows.Forms where possible, more stable

        switch -regex ($env:COMPUTERNAME) {
            ($rgxMyBox) { $LocalInclDir = "c:\usr\work\exch\scripts" ; }
            ($rgxProdEx2010Servers) { $LocalInclDir = "c:\scripts" ; }
            ($rgxLabEx2010Servers) { $LocalInclDir = "c:\scripts" ; }
            ($rgxProdL13Servers) { $LocalInclDir = "c:\scripts" ; }
            ($rgxLabL13Servers) { $LocalInclDir = "c:\scripts" ; }
            ($rgxAdminJumpBoxes) { $LocalInclDir = (split-path $profile) ; }
        } ;

        $Retries = 4 ;
        $RetrySleep = 5 ;
        $DawdleWait = 30 ; # wait time (secs) between dawdle checks
        $DirSyncInterval = 30 ; # AADConnect dirsync interval 

        # email trigger vari, it will be semi-delimd list of mail-triggering events
        $PassStatus = $null ;


        $bMaintainEquipementLists = $false ; 
        $MsgTraceNonDeliverDetailsLimit = 10 ; # number of non status:Delivered messages to check/dump
    
        $ModMsgWindowMins = 5 ;
        if(!$Retries){$Retries = 4 } ;
        if(!$RetrySleep){$RetrySleep = 5} ;

         #$rgxOnMSAddr='.*.mail\.onmicrosoft\.com' ;
        # cover either mail. or not
        $rgxOnMSAddr = '.*@\w*((\.mail)*)\.onmicrosoft.com'
        $rgxExoSysFolders = '.*\\(Recipient\sCache|PersonMetadata|Audits|Calendar\sLogging|Purges)$'
            #'.*\\(Recipient\sCache|PersonMetadata|Audits|Calendar\sLogging|Purges|GAL\sContacts|Yammer\sRoot|Recoverable\sItems|Deletions|Versions)'
        #$rgxExcl = '.*\\(Recipient\sCache|PersonMetadata|Audits|Calendar\sLogging|Purges)$' ;

        $rgxSID = "^S-\d-\d+-(\d+-){1,14}\d+$" ;
        $rgxEntLicGrps = "CN=ENT-APP-Office365-.*-DL,OU=ENTERPRISE,DC=global,DC=ad,DC=toro((lab)*),DC=com" ;

        $propsXmbx = 'UserPrincipalName','Alias','ExchangeGuid','Database','ExternalDirectoryObjectId','RemoteRecipientType'
        $propsOPmbx = 'UserPrincipalName','SamAccountName','RecipientType','RecipientTypeDetails' ; 
        $propsRmbx = 'UserPrincipalName','ExchangeGuid','RemoteRoutingAddress','RemoteRecipientType' ;
        #$propNames = "SamAccountName","UserPrincipalName","name","mailNickname","msExchHomeServerName","mail","msRTCSIP-UserEnabled","msRTCSIP-PrimaryUserAddress","msRTCSIP-Line","DistinguishedName","Description","info","Enabled","LastLogonDate","userAccountControl","manager","whenChanged","whenCreated","msExchRecipientDisplayType","msExchRecipientTypeDetails","City","Company","Country","countryCode","Office","Department","Division","EmailAddress","employeeType","State","StreetAddress","surname","givenname","telephoneNumber" ;
        #$propNames = "SamAccountName","UserPrincipalName","name","mailNickname","msExchHomeServerName","mail","msRTCSIP-UserEnabled","msRTCSIP-PrimaryUserAddress","msRTCSIP-Line","DistinguishedName","Description","info","Enabled","LastLogonDate","userAccountControl","manager","whenChanged","whenCreated","msExchRecipientDisplayType","msExchRecipientTypeDetails","City","Company","Country","countryCode","Office","Department","Division","EmailAddress","employeeType","State","StreetAddress","surname","givenname","telephoneNumber" ;
        #$adprops = "samaccountname", "msExchRemoteRecipientType", "msExchRecipientDisplayType", "msExchRecipientTypeDetails", "userprincipalname" ;
        $adprops = "samaccountname","UserPrincipalName","memberof","msExchMailboxGuid","msexchrecipientdisplaytype","msExchRecipientTypeDetails","msExchRemoteRecipientType"
        $propsmbxfldrs = @{Name='Folder'; Expression={$_.Identity.tostring()}},@{Name='Items'; Expression={$_.ItemsInFolder}}, @{n='SizeMB'; e={($_.FolderSize.tostring().split('(')[1].split(' ')[0].replace(',','')/1MB).ToString('0.000')}}, @{Name='OldestItem'; Expression={get-date $_.OldestItemReceivedDate -f 'yyyy/MM/dd'}},@{Name='NewestItem'; Expression={get-date $_.NewestItemReceivedDate -f 'yyyy/MM/dd'}} ;
        $propsMsgTrc = @{N='ReceivedLocal';E={$_.Received.ToLocalTime()}},'SenderAddress','RecipientAddress','Subject','Status','ToIP','FromIP','Size','MessageId','MessageTraceId','Index' ; 
        $propsMsgTrcDtl = 'Date','Event','Action','Detail','Data' ;  
        $msgprops = 'Received','SenderAddress','RecipientAddress','Subject','Status' # ,'MessageId','MessageTraceId' ;
        $mtdprops = 'Date','Event','Detail','data' ;
    
        #*======v HELPER FUNCTIONS v======

        #-------v Function _cleanup v-------
        function _cleanup {
            # clear all objects and exit
            # Clear-item doesn't seem to work as a variable release

            # 12:58 PM 7/23/2019 spliced in chunks from same in maintain-restrictedothracctsmbxs.ps1
            # 2:02 PM 9/21/2018 missing $timestampnow, hardcode
            # 8:45 AM 10/13/2015 reset $DebugPreference to default SilentlyContinue, if on
            # # 8:46 AM 3/11/2015 at some time from then to 1:06 PM 3/26/2015 added ISE Transcript
            # 8:39 AM 12/10/2014 shifted to stop-transcriptLog function
            # 7:43 AM 1/24/2014 always stop the running transcript before exiting

            write-verbose "_cleanup" ; 
            <# transcript/log are handled in the mbx loop
            #stop-transcript
            # 11:16 AM 1/14/2015 aha! does this return a value!??
            if ($host.Name -eq "Windows PowerShell ISE Host") {
                # 8:46 AM 3/11/2015 shift the logfilename gen out here, so that we can arch it
                #$Logname= (join-path -path (join-path -path $scriptDir -childpath "logs") -childpath ($scriptNameNoExt + "-" + (get-date -uformat "%Y%m%d-%H%M" ) + "-ISEtrans.log")) ;
                # 2:16 PM 4/27/2015 shift to static timestamp $timeStampNow
                #$Logname= (join-path -path (join-path -path $scriptDir -childpath "logs") -childpath ($scriptNameNoExt + "-" + $timeStampNow + "-ISEtrans.log")) ;
                # 2:02 PM 9/21/2018 missing $timestampnow, hardcode
                $Logname = (join-path -path (join-path -path $scriptDir -childpath "logs") -childpath ($scriptNameNoExt + "-" + (get-date -format 'yyyyMMdd-HHmmtt') + "-ISEtrans.log")) ;
                write-host "`$Logname: $Logname";
                Start-iseTranscript -logname $Logname ;
                #Archive-Log $Logname ;
                # 1:23 PM 4/23/2015 standardize processing file so that we can send a link to open the transcript for review
                $transcript = $Logname
            } else {
                write-verbose "$(get-timestamp):Stop Transcript" ;
                Stop-TranscriptLog ;
                #write-verbose "$(get-timestamp):Archive Transcript" ;
                #Archive-Log $transcript ;
            } # if-E
            
            # also echo the log:
            if ($logging) { 
                # Write-Log -LogContent $smsg -Path $logfile
                $smsg = "`$logging:`$true:written to:`n$($logfile)" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } ; 
            #>

            # need to email transcript before archiving it
            
            #$smtpSubj= "Proc Rpt:$($ScriptBaseName):$(get-date -format 'yyyyMMdd-HHmmtt')"   ;

            #Load as an attachment into the body text:
            #$body = (Get-Content "path-to-file\file.html" ) | converto-html ;
            #$SmtpBody += ("Pass Completed "+ [System.DateTime]::Now + "`nResults Attached: " +$transcript) ;
            # 4:07 PM 10/11/2018 giant transcript, no send
            #$SmtpBody += "Pass Completed $([System.DateTime]::Now)`nResults Attached:($transcript)" ;
            $SmtpBody += "Pass Completed $([System.DateTime]::Now)`nTranscript:($transcript)" ;
            # 12:55 PM 2/13/2019 append the $PassStatus in for reference
            if($PassStatus ){
                $SmtpBody += "`n`$PassStatus triggers:: $($PassStatus)`n" ;
            } ;
            $SmtpBody += "`n$('-'*50)" ;
            #$SmtpBody += (gc $outtransfile | ConvertTo-Html) ;
            # name $attachment for the actual $SmtpAttachment expected by Send-EmailNotif
            #$SmtpAttachment=$transcript ;
            # 1:33 PM 4/28/2017 test for ERROR|CHANGE - actually non-blank, only gets appended to with one or the other
            if($PassStatus ){
                Send-EmailNotif ;
            } else {
                write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):No Email Report: `$Passstatus is `$null ; " ;
            } ;

            
            #$smsg= "#*======^ END PASS:$($ScriptBaseName) ^======" ;
            #this is now a function, use: ${CmdletName}
            $smsg= "#*======^ END PASS:$(${CmdletName}) ^======" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn

            Break
        } #*------^ END Function _cleanup ^------


        #*======^ END HELPER FUNCTIONS ^======

        #*======v SUB MAIN v======

        # email trigger vari, it will be semi-delimd list of mail-triggering events
        $script:PassStatus = $null ;
        $Alltranscripts = @() ;

        # defer transcript until mbx loop
    
        #-=-=configure EXO EMS aliases to cover useEXOv2 requirements-=-=-=-=-=-=
        # have to preconnect, as it gcm's the targets
        if ($script:useEXOv2) { reconnect-eXO2 }
        else { reconnect-EXO } ;
        # in this case, we need an alias for EXO, and non-alias for EXOP
        [array]$cmdletMaps = 'ps1GetxRcp;get-exorecipient;','ps1GetxMbx;get-exomailbox;','ps1SetxMbx;Set-exoMailbox;','ps1GetxUser;get-exoUser;',
            'ps1SetxCalProc;set-exoCalendarprocessing;','ps1GetxCalProc;get-exoCalendarprocessing;','ps1GetxMbxFldrPerm;get-exoMailboxfolderpermission;',
            'ps1GetxAccDom;Get-exoAcceptedDomain;','ps1GetXRetPol;Get-exoRetentionPolicy','ps1GetxDistGrp;get-exoDistributionGroup;',
            'ps1GetxDistGrpMbr;get-exoDistributionGroupmember;','ps1GetxMsgTrc;get-exoMessageTrace;','ps1GetxMsgTrcDtl;get-exoMessageTraceDetail;',
            'ps1GetxMbxFldrStats;get-exoMailboxfolderStatistics','ps1GetxMContact;Get-exomailcontact;','ps1SetxMContact;Set-exomailcontact;',
            'ps1NewxMContact;New-exomailcontact;' ,'ps1TestxMapi;Test-exoMAPIConnectivity' ;
        foreach($cmdletMap in $cmdletMaps){
            if($script:useEXOv2){
                if(!($cmdlet= Get-Command $cmdletMap.split(';')[1].replace('-exo','-xo') )){ throw "unable to gcm Alias definition!:$($cmdletMap.split(';')[1])" ; break }
                $nalias = set-alias -name ($cmdletMap.split(';')[0]) -value ($cmdlet.name) -passthru ;
                write-verbose "$($nalias.Name) -> $($nalias.ResolvedCommandName)" ;
            } else {
                if(!($cmdlet= Get-Command $cmdletMap.split(';')[1])){ throw "unable to gcm Alias definition!:$($cmdletMap.split(';')[1])" ; break }
                $nalias = set-alias -name ($cmdletMap.split(';')[0]) -value ($cmdlet.name) -passthru ;                
                write-verbose "$($nalias.Name) -> $($nalias.ResolvedCommandName)" ;
            } ;
        } ;
    
        # shifting from ps1 to a function: need updates self-name:
        ${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name ;

        #$sBnr="#*======v START PASS:$($ScriptBaseName) v======" ; 
        $sBnr="#*======v START PASS:$(${CmdletName}) v======" ; 
        $smsg= $sBnr ;   
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;


        # steer all onprem code on $XXXMeta.ExOPAccessFromToro & Ex10Server values
        $UseOP=$false ; 
        if((Get-Variable  -name "$($TenOrg)Meta").value.ExOPAccessFromToro -AND (Get-Variable  -name "$($TenOrg)Meta").value.Ex10Server){
            $UseOP = $true ; 
            $smsg = "$($TenOrg):Meta.ExOPAccessFromToro($((Get-Variable  -name "$($TenOrg)Meta").value.ExOPAccessFromToro)) -AND/OR Meta.Ex10Server($((Get-Variable  -name "$($TenOrg)Meta").value.Ex10Server)),`ENABLING use of OnPrem Ex system this pass." ; 
            if($verbose){ if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        } else { 
            $UseOP = $false ; 
            $smsg = "$($TenOrg):Meta.ExOPAccessFromToro($((Get-Variable  -name "$($TenOrg)Meta").value.ExOPAccessFromToro)) -AND/OR Meta.Ex10Server($((Get-Variable  -name "$($TenOrg)Meta").value.Ex10Server)),`nDISABLING use of OnPrem Ex system this pass." ; 
            if($verbose){ if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        } ; 

        $useEXO = $true ; # non-dyn setting, drives variant EXO reconnect & query code
        if($useEXO){
            #*------v GENERIC EXO CREDS & SVC CONN BP v------
            # o365/EXO creds
            <### Usage: Type defaults to SID, if not spec'd - Note: there must be a *logged in & configured *profile* 
            $o365Cred=get-TenantCredentials -TenOrg $TenOrg -verbose -userrole SID ;
            Returns a credential set for the $TenOrg Hybrid OnPrem Exchange Org
            .EXAMPLE
            $o365Cred=get-TenantCredentials -TenOrg $TenOrg -verbose -userrole CSVC ;
            Returns the CSVC Userrole credential for the $TenOrg Hybrid OnPrem Exchange Org
            .EXAMPLE
            $o365Cred=get-TenantCredentials -TenOrg $TenOrg -verbose -userrole B2BI ;
            Returns the B2BI Userrole credential for the $TenOrg Hybrid OnPrem Exchange Org
            ###>
            $o365Cred=$null ;
            <# $TenOrg is a mandetory param in this script, skip dyn resolution
            switch -regex ($env:USERDOMAIN){
                "(TORO|CMW)" {$TenOrg = $env:USERDOMAIN.substring(0,3).toupper() } ;
                "TORO-LAB" {$TenOrg = 'TOL' }
                default {
                    throw "UNRECOGNIZED `$env:USERDOMAIN!:$($env:USERDOMAIN)" ; 
                    Break ; 
                } ;
            } ; 
            #>
            if($o365Cred=(get-TenantCredentials -TenOrg $TenOrg -UserRole 'CSVC','SID' -verbose:$($verbose))){
                # make it script scope, so we don't have to predetect & purge before using new-variable
                New-Variable -Name cred$($tenorg) -scope Script -Value $o365Cred.cred ;
                $smsg = "Resolved $($Tenorg) `$o365cred:$($o365Cred.cred.username) (assigned to `$cred$($tenorg))" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } else {
                $statusdelta = ";ERROR";
                $script:PassStatus += $statusdelta ;
                set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta) ;
                $smsg = "Unable to resolve $($tenorg) `$o365Cred value!"
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                throw "Unable to resolve $($tenorg) `$o365Cred value!`nEXIT!"
                Break ;
            } ;
            <### CALLS ARE IN FORM: (cred$($tenorg))
            $pltRXO = @{Credential = $Credential ; verbose = $($verbose) ; }
            $pltRXO = @{
                Credential = (Get-Variable -name cred$($tenorg) ).value ;
                verbose = $($verbose) ; } ; 
            if ($script:useEXOv2) { reconnect-eXO2 @pltRXO }
            else { reconnect-EXO @pltRXO } ;
            # or with Tenant-specific cred($Tenorg) lookup
            #$pltRXO creds & .username can also be used for AzureAD connections 
            Connect-AAD @pltRXO ; 
            ###>
            # configure splat for connections: (see above useage)
            $pltRXO = @{
                Credential = (Get-Variable -name cred$($tenorg) ).value ;
                verbose = $($verbose) ; } ; 
            #*------^ END GENERIC EXO CREDS & SVC CONN BP ^------
        } # if-E $useEXO

        if($UseOP){
            #*------v GENERIC EXOP CREDS & SRVR CONN BP v------
            # do the OP creds too
            $OPCred=$null ;
            # default to the onprem svc acct
            $pltGHOpCred=@{TenOrg=$TenOrg ;userrole='ESVC' ;verbose=$($verbose)} ;
            if($OPCred=(get-HybridOPCredentials @pltGHOpCred).cred){
                # make it script scope, so we don't have to predetect & purge before using new-variable
                New-Variable -Name "cred$($tenorg)OP" -scope Script -Value $OPCred ;
                $smsg = "Resolved $($Tenorg) `$OPCred:$($OPCred.username) (assigned to `$cred$($tenorg)OP)" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } else {
                $statusdelta = ";ERROR";
                $script:PassStatus += $statusdelta ;
                set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta) ;
                $smsg = "Unable to resolve get-HybridOPCredentials -TenOrg $($TenOrg) -userrole 'ESVC' value!"
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                throw "Unable to resolve $($tenorg) `$OPCred value!`nEXIT!"
                Break ;
            } ;
            $smsg= "Using EXOP cred:`$cred$($tenorg)OP:$((Get-Variable -name "cred$($tenorg)OP" ).value.username)" ;  
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
            <# CALLS ARE IN FORM: (cred$($tenorg))
            $pltRX10 = @{
                Credential = (Get-Variable -name "cred$($tenorg)OP" ).value ;
                verbose = $($verbose) ; }
            ReConnect-Ex2010XO @pltRX10 ; # cross-prem conns
            Reconnect-Ex2010 @pltRX10 ; # local org conns
            #$pltRx10 creds & .username can also be used for local ADMS connections
            #>
            $pltRX10 = @{
                Credential = (Get-Variable -name "cred$($tenorg)OP" ).value ;
                verbose = $($verbose) ; } ;     
            # TEST
        
            # defer cx10/rx10, until just before get-recipients qry
            #*------^ END GENERIC EXOP CREDS & SRVR CONN BP ^------
            # connect to ExOP X10
            if($pltRX10){
                #ReConnect-Ex2010XO @pltRX10 ;
                ReConnect-Ex2010 @pltRX10 ;
            } else { Reconnect-Ex2010 ; } ; 
        } ;  # if-E $useEXOP

        <# already confirmed in modloads
        # load ADMS
        $reqMods += "load-ADMS".split(";") ;
        if ( !(check-ReqMods $reqMods) ) { write-error "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Missing function. EXITING." ; Break ; }  ;
        #>
        write-host -foregroundcolor gray  "(loading ADMS...)" ;
        #write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):MSG" ; 

        load-ADMS ;

        if($UseOP){
            # resolve $domaincontroller dynamic, cross-org
            # setup ADMS PSDrives per tenant 
            if(!$global:ADPsDriveNames){
                $smsg = "(connecting X-Org AD PSDrives)" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                $global:ADPsDriveNames = mount-ADForestDrives -verbose:$($verbose) ;
            } ; 
            if(($global:ADPsDriveNames|measure).count){
                $useEXOforGroups = $false ; 
                $smsg = "Confirming ADMS PSDrives:`n$(($global:ADPsDriveNames.Name|%{get-psdrive -Name $_ -PSProvider ActiveDirectory} | ft -auto Name,Root,Provider|out-string).trim())" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                # returned object
                #         $ADPsDriveNames
                #         UserName                Status Name        
                #         --------                ------ ----        
                #         DOM\Samacctname   True  [forestname wo punc] 
                #         DOM\Samacctname   True  [forestname wo punc]
                #         DOM\Samacctname   True  [forestname wo punc]
        
            } else { 
                #-=-record a STATUS=-=-=-=-=-=-=
                $statusdelta = ";ERROR";
                $script:PassStatus += $statusdelta ;
                set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta) ;
                #-=-=-=-=-=-=-=-=
                $smsg = "Unable to detect POPULATED `$global:ADPsDriveNames!`n(should have multiple values, resolved to $()"
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                throw "Unable to resolve $($tenorg) `$o365Cred value!`nEXIT!"
                Break ;
            } ; 
        } ; 
        #if (!$domaincontroller) { $domaincontroller = get-gcfast } ;
        #if(!$domaincontroller){ if(test-path function:get-gcfast){$domaincontroller=get-gcfast} else { throw "no get-gcfast()!" } ;} else {"(existing `$domaincontroller:$($domaincontroller))"} ;
        # use new get-GCFastXO cross-org dc finde
        # default to Op_ExADRoot forest from $TenOrg Meta
        $domaincontroller = get-GCFastXO -TenOrg $TenOrg -subdomain ((gv -name "$($TenOrg)Meta").value['OP_ExADRoot']) -verbose:$($verbose) |?{$_.length};


        <# MSOL CONNECTION
        $reqMods += "connect-msol".split(";") ;
        if ( !(check-ReqMods $reqMods) ) { write-error "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Missing function. EXITING." ; Break ; }  ;
        write-host -foregroundcolor gray  "(loading AAD...)" ;
        #connect-msol ;
        connect-msol @pltRXO ; 
        #>

        # AZUREAD CONNECTION
        $reqMods += "Connect-AAD".split(";") ;
        if ( !(check-ReqMods $reqMods) ) { write-error "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Missing function. EXITING." ; Break ; }  ;
        write-host -foregroundcolor gray  "(loading AAD...)" ;
        #connect-msol ;
        Connect-AAD @pltRXO ; 
        #


        #
        # EXO connection
        $pltRXO = @{
            Credential = (Get-Variable -name cred$($tenorg) ).value ;
            verbose = $($verbose) ; } ; 
        if($VerbosePreference = "Continue"){
            $VerbosePrefPrior = $VerbosePreference ;
            $VerbosePreference = "SilentlyContinue" ;
            $verbose = ($VerbosePreference -eq "Continue") ;
        } ; 
        disconnect-exo ;
        if ($script:useEXOv2) { reconnect-eXO2 @pltRXO }
        else { reconnect-EXO @pltRXO } ;
        # reenable VerbosePreference:Continue, if set, during mod loads 
        if($VerbosePrefPrior -eq "Continue"){
            $VerbosePreference = $VerbosePrefPrior ;
            $verbose = ($VerbosePreference -eq "Continue") ;
        } ;
        #

        
        # 3:00 PM 9/12/2018 shift this to 1x in the script ; - this would need to be customized per tenant, not used (would normally be for forcing UPNs, but CMW uses brand UPN doms)
        #$script:forestdom = ((get-adforest | select -expand upnsuffixes) | ? { $_ -eq 'toro.com' }) ;

        # Clear error variable
        $Error.Clear() ;
        

    } ;  # BEGIN-E
    PROCESS {
        $rMailboxes = @() ;
        $Error.Clear() ;
        $SearchesRun = 0 ;
        $smsg = "Net:$(($Mailboxes|measure).count) mailboxes" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        $ttl = ($Mailboxes | measure).count ;
        $Procd = 0 ;

        # 9:20 AM 2/25/2019 Tickets will be an array of nnn's to match the mbxs, so use $Procd-1 as the index for tick# in the array
        if((($Mailboxes|measure).count -gt 1) -AND (($Tickets|measure).count -eq 1)){
            # Mult mbxs with single Ticket, expand ticket count to match # of mbxes
            foreach($mbx in $Mailboxes){
                $Tickets+=$Tickets[0] ; 
            } ; 
        } ; 
        foreach($Mailbox in $Mailboxes){
            $Procd++ ;
        
            if ($Ticket = $Tickets[($Procd - 1)]) {
                $sBnr = "#*======v `$Ticket:$($ticket):`$Mailbox:($($Procd)/$($ttl)):$($Mailbox) v======" ;
                $ofileroot = ".\logs\$($ticket)-" ;
                $lTag = "$($ticket)-$($Mailbox)-" ; 
            }else {
                $sBnr = "#*======v `$Ticket:(not spec'd):`$Mailbox:($($Procd)/$($ttl)):$($Mailbox) v======" ;
                $ofileroot = ".\logs\" ;
                $lTag = "$($Mailbox)-" ; 
            } ;

            # detect profile installs (installed mod or script), and redir to stock location
            $dPref = 'd','c' ; foreach($budrv in $dpref){ if(test-path -path "$($budrv):\scripts" -ea 0 ){ break ;  } ;  } ;
            [regex]$rgxScriptsModsAllUsersScope="^$([regex]::escape([environment]::getfolderpath('ProgramFiles')))\\((Windows)*)PowerShell\\(Scripts|Modules)" ;
            [regex]$rgxScriptsModsCurrUserScope="^$([regex]::escape([environment]::getfolderpath('Mydocuments')))\\((Windows)*)PowerShell\\(Scripts|Modules)" ;
            # -Tag "($TenOrg)-LASTPASS" 
            $pltSLog = [ordered]@{ NoTimeStamp=$false ; Tag=$lTag  ; showdebug=$($showdebug) ;whatif=$($whatif) ;} ;
            if($PSCommandPath){
                if(($PSCommandPath -match $rgxScriptsModsAllUsersScope) -OR ($PSCommandPath -match $rgxScriptsModsCurrUserScope) ){
                    # AllUsers or CU installed script, divert into [$budrv]:\scripts (don't write logs into allusers context folder)
                    if($PSCommandPath -match '\.ps(d|m)1$'){
                        # module function: use the ${CmdletName} for childpath
                        $pltSLog.Path= (join-path -Path "$($budrv):\scripts" -ChildPath "$(${CmdletName}).ps1" )  ;
                    } else { 
                        $pltSLog.Path=(join-path -Path "$($budrv):\scripts" -ChildPath (split-path $PSCommandPath -leaf)) ;
                    } ; 
                }else {
                    $pltSLog.Path=$PSCommandPath ;
                } ;
            } else {
                if( ($MyInvocation.MyCommand.Definition -match $rgxScriptsModsAllUsersScope) -OR ($MyInvocation.MyCommand.Definition -match $rgxScriptsModsCurrUserScope) ){
                    $pltSLog.Path=(join-path -Path "$($budrv):\scripts" -ChildPath (split-path $PSCommandPath -leaf)) ;
                } else {
                    $pltSLog.Path=$MyInvocation.MyCommand.Definition ;
                } ;
            } ;
            $smsg = "start-Log w`n$(($pltSLog|out-string).trim())" ;
            if($verbose){ if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            $logspec = start-Log @pltSLog ;

            if($logspec){
                $stopResults = try {Stop-transcript -ErrorAction stop} catch {} ;
                $logging=$logspec.logging ;
                $logfile=$logspec.logfile ;
                $transcript=$logspec.transcript ;
                #Configure default logging from parent script name
                # logfile                        C:\usr\work\o365\scripts\logs\move-MailboxToXo-(TOR)-LASTPASS-LOG-BATCH-WHATIF-log.txt
                # transcript                     C:\usr\work\o365\scripts\logs\move-MailboxToXo-(TOR)-LASTPASS-Transcript-BATCH-WHATIF-trans-log.txt
                #$logfile = $logfile.replace('-LASTPASS','').replace('BATCH',(Remove-InvalidFileNameChars -name $BatchName )) ;
                $logfile = $logfile.replace('-LASTPASS','').replace('BATCH','') ;
                #$transcript = $transcript.replace('-LASTPASS','').replace('BATCH',(Remove-InvalidFileNameChars -name $BatchName )) ;
                $transcript = $transcript.replace('-LASTPASS','').replace('BATCH','') ;
                if(Test-TranscriptionSupported){start-transcript -Path $transcript }
                else { write-warning "$($host.name) v$($host.version.major) does *not* support Transcription!" } ;
            } else {throw "Unable to configure logging!" } ;
    
            $smsg = "$($sBnr)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

            $sBnr="#*======v MAILBOX:$($Mailbox) v======" ;
            $smsg = "$((get-date).ToString('HH:mm:ss')):$($sBnr)" ;
            $pltXcsv = [ordered]@{Path=$null; NoTypeInformation=$true ; } ;
            $error.clear() ;

            $rcp = $null ; $exorcp = $null ;
            $exombx = $null ; $opombx = $null ;  $rmbx = $null ;
            $targetUPN = $NULL ;
            $adu = $null ; $msolu = $null ;
            $bMissingMbx = $false ;
            $exolicdetails = $null ;
            $obj = $null ;
            $grantgroup = $null ;
            $tuser = $null ;
            $UPN = $null ;
            $MailboxFoldersExo = $null ;

            $Exit = 0 ;

            if($pltRXO){
                if ($script:useEXOv2) { reconnect-eXO2 @pltRXO }
                else { reconnect-EXO @pltRXO } ;
            } else {reconnect-exo ;} ; 
            #Reconnect-Ex2010 ;
            if($pltRX10){
                #ReConnect-Ex2010XO @pltRX10 ;
                ReConnect-Ex2010 @pltRX10 ;
            } else { Reconnect-Ex2010 ; } ; 

            $Exit = 0 ; $error.clear() ;
            Do {
                Try {
                    $rcp = get-recipient -id $Mailbox -domaincontroller $domaincontroller -ea stop ;
                
                    $smsg= "(successful get-recipient -id $($Mailbox))" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else { write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    $Exit = $Retries ;
                }
                Catch {
                    $errTrpd=$_ ; 
                    if( $errtrpd -match "\scouldn't\sbe\sfound\son\s" ){
                        $smsg = "(no EXOP recicpient found)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                        else { write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        $Exit = $Retries ;
                    } else { 
                        $smsg=": Error Details: $($errTrpd)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                        Start-Sleep -Seconds $RetrySleep ;
                        if($pltRX10){
                            #ReConnect-Ex2010XO @pltRX10 ;
                            ReConnect-Ex2010 @pltRX10 ;
                        } else { ReConnect-Ex2010  } ; 
                        $Exit ++ ;
                        Write-Verbose "Failed to exec cmd because: $($Error[0])" ;
                        Write-Verbose "Try #: $Exit" ;                    
                    } ; 
                    If ($Exit -eq $Retries) { Write-Warning "Unable to exec cmd!" } ;
                }  ;
            } Until ($Exit -eq $Retries) ;

            $Exit = 0 ; $error.clear() ;
            Do {
                Try {
                    $pltgxrcp=[ordered]@{identity=$Mailbox ; erroraction='stop'; } ; 

                    $smsg= "$((get-alias ps1GetxRcp).definition) w`n$(($pltgxrcp|out-string).trim())" ; 
                    if($verbose){ if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 

                    $exorcp = ps1GetxRcp @pltgxrcp ;
                
                    $smsg= "(successful get-exorecipient -id $($Mailbox))" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                
                    $Exit = $Retries ;
                }
                Catch {
                    # get fails go here with: The operation couldn't be performed because object 'Anthony.Magana@toro.com' couldn't be found on 'CY4PR04A008DC10.NAMPR04A008.PROD.OUTLOOK.COM'.
                    $errTrpd=$_ ; 
                    if( $errtrpd -match "\scouldn't\sbe\sfound\son\s" ){
                        $smsg = "(no EXO recipient found)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                        else { write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        $Exit = $Retries ;
                    } else { 
                        $smsg=": Error Details: $($errTrpd)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                        Start-Sleep -Seconds $RetrySleep ;
                        if($pltRXO){
                            if ($script:useEXOv2) { reconnect-eXO2 @pltRXO }
                            else { reconnect-EXO @pltRXO } ;
                        } else {reconnect-exo ;} ; 
                        $Exit ++ ;
                        Write-Verbose "Failed to exec cmd because: $($Error[0])" ;
                        Write-Verbose "Try #: $Exit" ;                    
                    } ; 
                    If ($Exit -eq $Retries) { Write-Warning "Unable to exec cmd!" } ;
                }  ;
            } Until ($Exit -eq $Retries) ;


            $Exit = 0 ; $error.clear() ;
            Do {
                Try {
                    $pltgxmbx=[ordered]@{identity=$Mailbox ; erroraction='stop'; } ; 
                    $smsg= "$((get-alias ps1GetxMbx).definition) w`n$(($pltgxmbx|out-string).trim())" ; 
                    if($verbose){ if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 

                
                    $exombx = ps1GetxMbx @pltgxmbx ; 

                    $smsg= "(successful get-exomailbox -id $($Mailbox))" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;


                    if ($exombx) {
                        $pltGetxMbxFldrStats=[ordered]@{Identity=$exombx.identity ;IncludeOldestAndNewestItems=$true; erroraction='stop'; } ; 
                        #$smsg= "(collecting get-exomailboxfolderstatistics...)" ;
                        $smsg= "$((get-alias ps1GetxMbxFldrStats).definition) w`n$(($pltGetxMbxFldrStats|out-string).trim())" ; 
                        if($verbose){ if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                        $exofldrs = ps1GetxMbxFldrStats @pltGetxMbxFldrStats | 
                            ?{($_.ItemsInFolder -gt 0 ) -AND ($_.identity -notmatch $rgxExoSysFolders)} ;
                        #$fldrs = get-exomailboxfolderstatistics -id KADRITS -IncludeOldestAndNewestItems ;
                        #$fldrs = $fldrs |?{($_.ItemsInFolder -gt 0 ) -AND ($_.identity -notmatch $rgxExoSysFolders)} ;
                        #$fldrs | ft -auto $propsmbxfldrs ;

                        # do a 7day msgtrc
                        #-=-=-=-=-=-=-=-=
                        $isplt=@{    
                            ticket=$ticket ;
                            uid=$Mailbox;
                            days=$Days ;
                            StartDate='' ;
                            EndDate='' ;
                            Sender="" ;
                            Recipients=$exombx.PrimarySmtpAddress ;
                            MessageSubject="" ;
                            Status='' ;
                            MessageTraceId='' ;
                            MessageId='' ;
                        }  ;
                        $msgtrk=@{ PageSize=1000 ; Page=$null ; StartDate=$null ; EndDate=$null ; } ;

                        if($isplt.days){  
                            $msgtrk.StartDate=(get-date ([datetime]::Now)).adddays(-1*$isplt.days);
                            $msgtrk.EndDate=(get-date) ;
                        } ;
                        if($isplt.StartDate -and !($isplt.days)){$msgtrk.StartDate=$(get-date $isplt.StartDate)} ;
                        if($isplt.EndDate -and !($isplt.days)){$msgtrkEndDate=$(get-date $isplt.EndDate)} 
                        elseif($isplt.StartDate -and !($isplt.EndDate)){
                            $smsg = '(StartDate w *NO* Enddate, asserting currenttime)' ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            $msgtrkEndDate=(get-date) ;
                        } ;

                        TRY{$tendoms=Get-AzureADDomain }CATCH{
                            #write-warning "NOT AAD CONNECTED!" ;BREAK ;
                            $smsg = "Not AAD connected, reconnect..." ; 
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            Connect-AAD @pltRXO ; 
                        } ;

                        $Ten = ($tendoms |?{$_.name -like '*.mail.onmicrosoft.com'}).name.split('.')[0] ;
                        $ofile ="$($isplt.ticket)-$($Ten)-$($isplt.uid)-EXOMsgTrk" ;
                        if($isplt.Sender){
                            if($isplt.Sender -match '\*'){
                                $smsg = "(wild-card Sender detected)" ;
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                $msgtrk.add("SenderAddress",$isplt.Sender) ;
                            } else {
                                $msgtrk.add("SenderAddress",$isplt.Sender) ;
                            } ;
                            $ofile+=",From-$($isplt.Sender.replace("*","ANY"))" ;
                        } ;
                        if($isplt.Recipients){
                            if($isplt.Recipients -match '\*'){
                                $smsg = "(wild-card Recipient detected)" ;
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                $msgtrk.add("RecipientAddress",$isplt.Recipients) ;
                            } else {
                                $msgtrk.add("RecipientAddress",$isplt.Recipients) ;
                            } ;
                            $ofile+=",To-$($isplt.Recipients.replace("*","ANY"))" ;
                        } ;

                        if($isplt.MessageId){
                            $msgtrk.add("MessageId",$isplt.MessageId) ;
                            $ofile+=",MsgId-$($isplt.MessageId.replace('<','').replace('>',''))" ;
                        } ;
                        if($isplt.MessageTraceId){
                            $msgtrk.add("MessageTraceId",$isplt.MessageTraceId) ;
                            $ofile+=",MsgId-$($isplt.MessageTraceId.replace('<','').replace('>',''))" ;
                        } ;
                        if($isplt.MessageSubject){
                            $ofile+=",Subj-$($isplt.MessageSubject.substring(0,[System.Math]::Min(10,$isplt.MessageSubject.Length)))..." ;
                        } ;
                        if($isplt.Status){
                            $msgtrk.add("Status",$isplt.Status)  ;
                            $ofile+=",Status-$($isplt.Status)" ;
                        } ;
                        if($isplt.days){$ofile+= "-$($isplt.days)d-" } ;
                        if($isplt.StartDate){$ofile+= "-$(get-date $isplt.StartDate -format 'yyyyMMdd-HHmmtt')-" } ;
                        if($isplt.EndDate){$ofile+= "$(get-date $isplt.EndDate -format 'yyyyMMdd-HHmmtt')" } ;

                        $smsg = "Running MsgTrk:$($Ten)" ;
                        $smsg += "`n$(($msgtrk|out-string).trim()|out-default)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                        $Page = 1  ;
                        $Msgs=$null ;
                        do {
                            $smsg = "Collecting - Page $($Page)..."  ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                            $msgtrk.Page=$Page ;
                            $PageMsgs = ps1GetxMsgTrc @msgtrk |  ?{$_.SenderAddress -notlike '*micro*' -or $_.SenderAddress -notlike '*root*' }  ;
                            $Page++  ;
                            $Msgs += @($PageMsgs)  ;
                        } until ($PageMsgs -eq $null) ;
                        $Msgs=$Msgs| Sort Received ;
                    
                        $smsg = "==Msgs Returned:$(($Msgs|measure).count)`nRaw matches:$(($Msgs|measure).Count)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                        if($isplt.MessageSubject){
                            $smsg = "Post-Filtering on Subject:$($isplt.MessageSubject)" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                            $Msgs = $Msgs | ?{$_.Subject -like $isplt.MessageSubject} ;
                            $ofile+="-Subj-$($isplt.MessageSubject.replace("*"," ").replace("\"," "))" ;
                            $smsg = "Post Subj filter matches:$(($Msgs|measure).Count)" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                        } ;
                        $ofile+= "-run$(get-date -format 'yyyyMMdd-HHmmtt').csv" ;
                        #$ofile=[RegEx]::Replace($ofile, "[{0}]" -f ([RegEx]::Escape(-join [System.IO.Path]::GetInvalidFileNameChars())), '') ;
                        $ofile = Remove-InvalidFileNameChars -name $ofile ; 
                        $ofile=".\logs\$($ofile)" ;
                     
                        if($Msgs){
                            $Msgs | select  | export-csv -notype -path $ofile  ;
                            "Status Distrib:" ;
                            $smsg = "`n#*------v MOST RECENT MATCH v------" ;
                            $smsg += "`n$(($msgs[-1]| format-list ReceivedLocal,StatusSenderAddress,RecipientAddress,Subject|out-string).trim())";
                            $smsg += "`n#*------^ MOST RECENT MATCH ^------" ;
                            $smsg += "`n#*------v Status DISTRIB v------" ;
                            $smsg += "`n$(($Msgs | select -expand Status | group | sort count,count -desc | select count,name |out-string).trim())";
                            $smsg += "`n#*------^ Status DISTRIB ^------" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            <# Count Name                      Group
                            ----- ----                      -----
                              448 Delivered                 {@{PSComputerName=ps.outlook.com; RunspaceId=25f3aa28-9437-4e30-aa8f-8d83d2d2fc5a; PSShowComputerName=False; Organization=toroco.onmicrosoft.com; MessageId=<BN8PR12MB33158F1D07B3862078D758E6EB499@BN8PR12MB3315.namprd12.prod.o...
                                1 Failed                    {@{PSComputerName=ps.outlook.com; RunspaceId=25f3aa28-9437-4e30-aa8f-8d83d2d2fc5a; PSShowComputerName=False; Organization=toroco.onmicrosoft.com; MessageId=<5bc2055fed4d43d3827cf7f61d37a4c9@CH2PR04MB7062.namprd04.prod.outlook...
                                1 Quarantined               {@{PSComputerName=ps.outlook.com; RunspaceId=25f3aa28-9437-4e30-aa8f-8d83d2d2fc5a; PSShowComputerName=False; Organization=toroco.onmicrosoft.com; MessageId=<threatsim-5f0bc0101d-c200b2590d@app.emaildistro.com>; Received=4/28/...
                                1 FilteredAsSpam            {@{PSComputerName=ps.outlook.com; RunspaceId=25f3aa28-9437-4e30-aa8f-8d83d2d2fc5a; PSShowComputerName=False; Organization=toroco.onmicrosoft.com; MessageId=<SA0PR01MB61858E3C6111672E081373C1E45F9@SA0PR01MB6185.prod.exchangela...
                            #>
                            $nonDelivStats = $msgs | ?{$_.status -ne 'Delivered'} | group status | select-object -expand name ; 
                            
                            foreach ($status in $nonDelivStats){
                                $smsg = "Enumerating Status:$($status) messages" ; 
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                if($StatMsgs  = $msgs|?{$_.status -eq $status}){
                                    if(($StatMsgs |measure).count -gt 10){
                                        $smsg = "(over $($MsgTraceNonDeliverDetailsLimit) $($status) msgs: processing a sample of last $($MsgTraceNonDeliverDetailsLimit)...)" ; 
                                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                        $StatMsgs = $StatMsgs | select-object -Last $MsgTraceNonDeliverDetailsLimit ;
                                    } ; 
                                    $statTtl = ($StatMsgs |measure).count ; $fProcd=0 ; 
                                    $smsg = "$($statTtl) Status:'$($status)' messages returned. Expanding detailed processing..." ; 
                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                        foreach($fmsg in $StatMsgs ){
                                            $fProcd++ ; 
                                            $sBnrS="`n#*------v PROCESSING $($status)#$($fProcd)/$($statTtl): v------" ; 
                                            $smsg= "$($sBnrS)" ;
                                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                                            $pltGetxMsgTrcDtl = [ordered]@{MessageTraceId=$fmsg.MessageTraceId ;RecipientAddress=$fmsg.RecipientAddress} ; 
                                            $smsg= "$((get-alias ps1GetxMsgTrcDtl).definition) w`n$(($pltGetxMsgTrcDtl|out-string).trim())" ; 
                                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                
                                            # this is nested in a try already
                                            $MsgTrkDtl = ps1GetxMsgTrcDtl @pltGetxMsgTrcDtl ; 

                                            if($MsgTrkDtl){
                                                $smsg = "`n-----`n$(( $MsgTrkDtl | fl $propsMsgTrcDtl|out-string).trim())`n-----`n" ; 
                                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                                                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                            } else { 
                                                $smsg = "(no matching MessageTraceDetail was returned by MS for MsgTrcID:$($pltGetxMsgTrcDtl.MessageTraceId), aged out of availability already)" ; 
                                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                                                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                                $smsg = "$($status)#$($fProcd)/$($statTtl):DETAILS`n----`n$(($fmsg|fl | out-string).trim())`n----`n" ; 
                                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                                                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                            } ; 

                                            $smsg= "$($sBnrS.replace('-v','-^').replace('v-','^-'))" ;
                                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                        } ; 
                                
                                }
                            } ; #nonDelivStats
                            if(test-path -path $ofile){
                                $smsg = "(log file confirmed)" ;
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                Resolve-Path -Path $ofile | select -expand Path | out-clipboard ;
                                $smsg = "$($Msgs.count) matches output to:`n'$($ofile)'`n(copied to CB)" ;
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                            } else { 
                                $smsg = "MISSING LOG FILE!"  ; 
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            } ;
                        } else {
                            $smsg = "NO MATCHES FOUND from::`n$(($msgtrk|out-string).trim()|out-default)`n(with any relevant ConnectorId postfilter)" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        } ;

                        #-=-=-=-=-=-=-=-=


                    } else {
                        $smsg = "(no EXO mailbox found, skipping $((get-alias ps1GetxMbxFldrStats).definition) )" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                        else { write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ;
                    $Exit = $Retries ;
                }
                Catch {
                    # get fails go here with: The operation couldn't be performed because object 'Anthony.Magana@toro.com' couldn't be found on 'CY4PR04A008DC10.NAMPR04A008.PROD.OUTLOOK.COM'.
                    $errTrpd=$_ ; 
                    if( $errtrpd -match "\scouldn't\sbe\sfound\son\s" ){
                        $smsg = "(no EXO mailbox found)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                        else { write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        $Exit = $Retries ;
                    } else { 
                        $smsg=": Error Details: $($errTrpd)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                        Start-Sleep -Seconds $RetrySleep ;
                        if($pltRXO){
                            if ($script:useEXOv2) { reconnect-eXO2 @pltRXO }
                            else { reconnect-EXO @pltRXO } ;
                        } else {reconnect-exo ;} ; 
                        $Exit ++ ;
                        Write-Verbose "Failed to exec cmd because: $($Error[0])" ;
                        Write-Verbose "Try #: $Exit" ;                    
                    } ; 
                    If ($Exit -eq $Retries) { Write-Warning "Unable to exec cmd!" } ;
                }  ;
            } Until ($Exit -eq $Retries) ;

            $Exit = 0 ; $error.clear() ;
            Do {
                Try {
                    $pltgRmbx=@{identity=$Mailbox ; domaincontroller=$domaincontroller ; erroraction='stop'; } ; 
                    $smsg= "get-RemoteMailbox w`n$(($pltgRmbx|out-string).trim())" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                    $rmbx = get-remotemailbox @pltgRmbx ;

                    if($rmbx){
                        $smsg= "(successful get-remotemailbox  -id $($Mailbox))" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ;
                    $Exit = $Retries ;
                }
                Catch {
                    # get fails go here with: The operation couldn't be performed because object 'blahblah' couldn't be found on 'BCCMS8100.global.ad.toro.com'.
                    $errTrpd=$_ ; 
                    if( $errtrpd -match "\scouldn't\sbe\sfound\son\s" ){
                        $smsg = "(no EXOP remotemailbox found)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                        else { write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        $Exit = $Retries ;
                    } else { 
                        $smsg=": Error Details: $($errTrpd)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                        Start-Sleep -Seconds $RetrySleep ;
                        if($pltRX10){
                            #ReConnect-Ex2010XO @pltRX10 ;
                            ReConnect-Ex2010 @pltRX10 ;
                        } else { ReConnect-Ex2010  } ; 
                        $Exit ++ ;
                        Write-Verbose "Failed to exec cmd because: $($Error[0])" ;
                        Write-Verbose "Try #: $Exit" ;                    
                    } ; 
                    If ($Exit -eq $Retries) { Write-Warning "Unable to exec cmd!" } ;
                }  ;
            } Until ($Exit -eq $Retries) ;


            if ($exombx ) {
                $targetUPN = $exombx.userprincipalname ;
            }
            elseif ($rmbx) {
                $targetUPN = $rmbx.userprincipalname ;
            }
            else {
                throw "Unable to locate either an EXO mailbox or an on-prem RemoteMailbox. ABORTING!"
                Break ;
            } ;


            $Exit = 0 ; $error.clear() ;
            Do {
                Try {
                    $pltgMbx=[ordered]@{identity=$targetUPN ;domaincontroller=$domaincontroller ;erroraction='SilentlyContinue';  };
                    $smsg = "get-mailbox w`n$(($pltgMbx|out-string).trim())" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                    $opmbx = get-mailbox @pltgMbx ; 
                    if($opmbx){
                        $smsg= "(successful get-mailbox -id $($targetUPN))" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ;
                    $Exit = $Retries ;

                }
                Catch {
                    # get fails go here with: The operation couldn't be performed because object 'blahblah' couldn't be found on 'BCCMS8100.global.ad.toro.com'.
                    $errTrpd=$_ ; 
                    if( $errtrpd -match "\scouldn't\sbe\sfound\son\s" ){
                        $smsg = "(no EXOP mailbox found)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                        else { write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        $Exit = $Retries ;
                    } else { 
                        $smsg=": Error Details: $($errTrpd)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                        Start-Sleep -Seconds $RetrySleep ;
                        if($pltRX10){
                            #ReConnect-Ex2010XO @pltRX10 ;
                            ReConnect-Ex2010 @pltRX10 ;
                        } else { ReConnect-Ex2010  } ; 
                        $Exit ++ ;
                        Write-Verbose "Failed to exec cmd because: $($Error[0])" ;
                        Write-Verbose "Try #: $Exit" ;                    
                    } ; 
                    If ($Exit -eq $Retries) { Write-Warning "Unable to exec cmd!" } ;
                }  ;
            } Until ($Exit -eq $Retries) ;


            $Exit = 0 ; $error.clear() ;
            Do {
                Try {
                    connect-msol @pltRXO ;
                    $pltgMSOLU = [ordered]@{userprincipalname=$targetUPN  ;erroraction='SilentlyContinue';} ; 
                    $smsg = "get-msoluser w`n$(($pltgMSOLU|out-string).trim())" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                    $msolu = get-msoluser @pltgMSOLU ;
                    if($msolu){
                        $smsg= "(successful get-msoluser -userprincipalname $($targetUPN))" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ;
                    $Exit = $Retries ;
                }
                Catch {
                    # get fails go here with: get-msoluser : User Not Found.  User: blah@toro.com.
                    $errTrpd=$_ ; 
                    if( $errtrpd -match "User\sNot\sFound" ){
                        $smsg = "(no EXOP mailbox found)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                        else { write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        $Exit = $Retries ;
                    } else { 
                        $smsg=": Error Details: $($errTrpd)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                        Start-Sleep -Seconds $RetrySleep ;
                        if($pltRXO){
                            if ($script:useEXOv2) { reconnect-eXO2 @pltRXO }
                            else { reconnect-EXO @pltRXO } ;
                        } else {reconnect-exo ;} ; 
                        $Exit ++ ;
                        Write-Verbose "Failed to exec cmd because: $($Error[0])" ;
                        Write-Verbose "Try #: $Exit" ;                    
                    } ; 
                    If ($Exit -eq $Retries) { Write-Warning "Unable to exec cmd!" } ;
                }  ;
            } Until ($Exit -eq $Retries) ;

            if ($msolu) {
                connect-msol @pltRXO;
               # to find group, check user's membership
                # do the full lookup first

                if ($msolu.IsLicensed -AND !($msolu.LicenseReconciliationNeeded)) {
                    $smsg = "USER HAS *NO* LICENSING ISSUES:" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else { write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                }
                else {
                    $smsg = "USER *HAS* LICENSING ISSUES:" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else { write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                } ;

                $smsg = "Get-MsolUser:`n$(($msolu | select userprin*,*Error*,*status*,softdel*,lic*,islic*|out-string).trim())`n" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else { write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                if ($null -eq $msolu.SoftDeleteTimestamp) {
                    $smsg = "$($msol.userprincipalname) has a BLANK SoftDeleteTimestamp`n=>If mailbox missing it would indicate the user wasn't properly de-licensed (or would have fallen into dumpster at >30d).`n That scenario would reflect a replic break (AAD sync loss wo proper update)`n" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else { write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                } ;

                if ($EXOLicDetails = get-MsolUserLicenseDetails -UPNs $targetUPN -showdebug:$($showdebug) ) {
                    $smsg = "Returned License Details`n$(($EXOLicDetails|out-string).trim())" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn
                }
                else {
                    $smsg = "UNABLE TO RETURN License Details FOR `n$(($targetUPN|out-string).trim())" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } ; #Error|Warn
                };

                # 2:12 PM 1/27/2020 rough in Pull-AADSignInReports.ps1 syntax
                $AADSigninSyntax=@"

---------------------------------------------------------------------------------
If you would like to retrieve user AAD Sign On Reports, use the following command:

A. Query the events into json outputs:

.\Pull-AADSignInReports.ps1 -UPNs "$($targetUPN)" -ticket "$($Ticket)" -StartDate (Get-Date).AddDays(-30) -showdebug ;

B. Process & Profile the output .json file from the above:
.\profile-AAD-Signons.ps1 -Files PATH-TO-FILE.json ;

---------------------------------------------------------------------------------

"@ ;

            }
            else {
                $smsg = "$($targetUPN):NO MATCHING MSOLUSER RETURNED!" ; ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } ;

        
            $Exit = 0 ; $error.clear() ;
            Do {
                Try {
                    $aduFilter = { UserPrincipalName -eq $targetUPN }  ;
                    $pltgADU=[ordered]@{filter=$aduFilter;Properties=$adprops ;server=$domaincontroller  ;erroraction='SilentlyContinue'; } ; 
                    $smsg = "get-aduser w`n$(($pltgADU|out-string).trim())" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                    #$adu = get-aduser -filter { UserPrincipalName -eq $targetUPN } -Properties $adprops -server $domaincontroller -ea 0  ; 
                    $adu = get-aduser @pltgADU  ; 
                    if($adu){
                        $smsg= "(successful get-aduser -filter { UserPrincipalName -eq $($targetUPN) }" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ;
                    $Exit = $Retries ;
                }
                Catch {
                    $errTrpd=$_ ;
                    $smsg = "Failed processing $($errTrpd.Exception.ItemName). `nError Message: $($errTrpd.Exception.Message)`nError Details: $($errTrpd)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Error } #Error|Warn|Debug
                    else { write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    Start-Sleep -Seconds $RetrySleep ;
                    $Exit ++ ;
                    $smsg = "Failed to exec cmd because: $($Error[0])`nTry #: $($Exit)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Error } #Error|Warn|Debug
                    else { write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    If ($Exit -eq $Retries) { Write-Warning "Unable to exec cmd!" } ;
                }  ;
            } Until ($Exit -eq $Retries) ;

            # DETERMINE LIC GRANT GROUP
            if ($GrantGroupDN = $adu | select -expand memberof | ? { $_ -match $rgxEntLicGrps } ) {
                Try {
                    # 8:47 AM 3/1/2019 #626:expand the name, (was coming out a hash)
                    $GrantGroup = get-adgroup -id $grantgroupDN | select -expand name ;
                }
                catch {
                    $grantgroup = "(NOT RESOLVABLE)"
                } ;
            }
            else {
                $grantgroup = "(NOT RESOLVABLE)"
            } ;

            # TRIAGE RESULTS
            if ($exombx -AND $opmbx) {
                # SPLIT BRAIN
                $smsg = "USER HAS SPLIT-BRAIN MAILBOX (MBX IN EXO & EXOP)!" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                else { write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

            }elseif (($rcp.RecipientTypeDetails -eq 'RemoteUserMailbox') -AND ($exorcp.RecipientTypeDetails -eq 'MailUser') ) {
                # NO BRAIN
                $smsg = "USER HAS NO-BRAIN MAILBOX (*NO* MBX IN *EITHER* EXO & EXOP)!" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                else { write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

             }
            elseif ( ($rcp.RecipientTypeDetails -match '(User|Shared|Room|Equipment)Mailbox') -AND ($exorcp.RecipientTypeDetails -eq 'MailUser') ) {
                # ON PREM MAILBOX
                #$mapiTest = Test-MAPIConnectivity -id $Mailbox ;
                $pltTMapi=[ordered]@{identity=$Mailbox;domaincontroller=$domaincontroller  ;erroraction='SilentlyContinue'; } ; 
                $smsg = "Test-MAPIConnectivity w`n$(($pltTMapi|out-string).trim())" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                Try {
                    $mapiTest = Test-MAPIConnectivity @pltTMapi  ; 
                    if ($mapiTest.Result -eq 'Success') {
                        $smsg= "(successful Outlook Mbx MAPI validate: Test-MAPIConnectivity -id $($Mailbox) }" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ;
                    $Exit = $Retries ;
                } Catch {
                    $errTrpd=$_ ;
                    $smsg = "Failed processing $($errTrpd.Exception.ItemName). `nError Message: $($errTrpd.Exception.Message)`nError Details: $($errTrpd)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Error } #Error|Warn|Debug
                    else { write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    Start-Sleep -Seconds $RetrySleep ;
                    $Exit ++ ;
                    $smsg = "Failed to exec cmd because: $($Error[0])`nTry #: $($Exit)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Error } #Error|Warn|Debug
                    else { write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    If ($Exit -eq $Retries) { Write-Warning "Unable to exec cmd!" } ;
                }  ;


                $FixText = @"
========================================
Problem User:`t "$($adu.UserPrincipalName)"
Has an *ONPREM* mailbox:Type:$($exombx.recipienttypedetails) hosted in $($opmbx.database)
"@ ;
                if ($mapiTest.Result -eq 'Success') {
                    $FixText2 = @"
*with *NO* detected issues*
Mailbox Outlook connectivity tested and validated functional...
"@ ;
                } else {
                    $FixText2 = @"
*FAILED Test-MAPIConnectivity*
Mailbox Outlook connectivity test failed to connect!...
"@ ;
                }

                $FixText3 = @"

$(($mapiTest|out-string).trim())

The user's o365 LICENSESKUs:                     "$($exolicdetails.LicAccountSkuID)"
With DisplayNames:                               "$($exolicdetails.LicenseFriendlyName)"
The user's o365 Licensing group appears to be:  "$($grantgroup)"

OnPrem RecipientTypeDetails:`t "$($rcp.RecipientTypeDetails)"
OnPrem WhenCreated:`t "$($rcp.WhenCreated)"
EXO RecipientTypeDetails:`t "$($exorcp.RecipientTypeDetails)"
EXO WhenMailboxCreated:`t "$($exombx.WhenMailboxCreated)"
"@

                $smsg = "$($FixText)`n$($FixText2)`n$($FixText3)"  ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else { write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                #output the AAD Signon Profile info: $AADSigninSyntax
                $smsg = "$($AADSigninSyntax)"  ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else { write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

            }elseif ( ($rcp.RecipientTypeDetails -match 'Remote(User|Shared|Room|Equipment)Mailbox' ) -AND ($exorcp.RecipientTypeDetails -match '(User|Shared|Room|Equipment)Mailbox') ) {
                # EXO MAILBOX
                $pltTxmc=@{identity=$Mailbox ;erroraction='SilentlyContinue'; } ;
                Try {
                    $mapiTest = ps1TestXMapi @pltTxmc ; 
                    if ($mapiTest.Result -eq 'Success') {
                        $smsg= "(successful Outlook Mbx MAPI validate: Test-MAPIConnectivity -id $($Mailbox) }" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ;
                    $Exit = $Retries ;
                } Catch {
                    $errTrpd=$_ ;
                    $smsg = "Failed processing $($errTrpd.Exception.ItemName). `nError Message: $($errTrpd.Exception.Message)`nError Details: $($errTrpd)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Error } #Error|Warn|Debug
                    else { write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    Start-Sleep -Seconds $RetrySleep ;
                    $Exit ++ ;
                    $smsg = "Failed to exec cmd because: $($Error[0])`nTry #: $($Exit)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Error } #Error|Warn|Debug
                    else { write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    If ($Exit -eq $Retries) { Write-Warning "Unable to exec cmd!" } ;
                }  ;
                $FixText = @"
========================================
User:`t "$($adu.UserPrincipalName)"
Has an *EXO* mailbox:Type:$($exombx.recipienttypedetails) in db:$($exombx.database)
"@ ;
                if ($mapiTest.Result -eq 'Success') {
                $FixText2 = @"
*with *NO* detected issues*
Mailbox Outlook connectivity tested and validated functional...
"@ ;
                } else {
                    $FixText2 = @"
*FAILED Test-exoMAPIConnectivity*
Mailbox Outlook connectivity test failed to connect!...
"@ ;
                } ; 

                $FixText3 = @"

$(($mapiTest|out-string).trim())

The user's o365 LICENSESKUs:                     "$($exolicdetails.LicAccountSkuID)"
With DisplayNames:                               "$($exolicdetails.LicenseFriendlyName)"
The user's o365 Licensing group appears to be:  "$($grantgroup)"
$(if($rcp.recipienttypedetails -ne 'RemoteUserMailbox'){
    '(non-usermailbox, *non*-licensed is typical status)'
})

OnPrem RecipientTypeDetails:`t "$($rcp.RecipientTypeDetails)"
OnPrem WhenCreated:`t "$($rcp.WhenCreated)"
EXO RecipientTypeDetails:`t "$($exorcp.RecipientTypeDetails)"
EXO WhenMailboxCreated:`t "$($exombx.WhenMailboxCreated)"

Mailbox content:
$(($exofldrs | ft -auto $propsmbxfldrs|out-string).trim())

"@


                $smsg = "$($FixText)`n$($FixText2)`n$($FixText3)"  ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else { write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                #output the AAD Signon Profile info: $AADSigninSyntax
                $smsg = "$($AADSigninSyntax)"  ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else { write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

            } else {
                # UNRECOGNIZED RECIPIENTTYPE COMBO
                # ($rcp.RecipientTypeDetails -eq 'RemoteUserMailbox'), ($exorcp.RecipientTypeDetails -eq 'UserMailbox')
                $smsg= "UNRECOGNIZED RECIPIENTTYPE COMBO:`nOPREM:RecipientTypeDetails:$($rcp.RecipientTypeDetails)`nEXO:RecipientTypeDetails:$($exorcp.RecipientTypeDetails)" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Warn } #Error|Warn|Debug
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

            } ; 

            
            if($host.Name -eq "Windows PowerShell ISE Host" -and $host.version.major -lt 5){
                # 11:51 AM 9/22/2020 isev5 supports transcript, anything prior has to fake it
                # 8:46 AM 3/11/2015 shift the logfilename gen out here, so that we can arch it
                #$Logname= (join-path -path (join-path -path $scriptDir -childpath "logs") -childpath ($scriptNameNoExt + "-" + (get-date -uformat "%Y%m%d-%H%M" ) + "-ISEtrans.log")) ;
                # 2:16 PM 4/27/2015 shift to static timestamp $timeStampNow
                #$Logname= (join-path -path (join-path -path $scriptDir -childpath "logs") -childpath ($scriptNameNoExt + "-" + $timeStampNow + "-ISEtrans.log")) ;
                # 2:02 PM 9/21/2018 missing $timestampnow, hardcode
                #$Logname=(join-path -path (join-path -path $scriptDir -childpath "logs") -childpath ($scriptNameNoExt + "-" + (get-date -format 'yyyyMMdd-HHmmtt') + "-ISEtrans.log")) ;
                # maintain-ExoUsrMbxFreeBusyDetails-TOR-ForceAll-Transcript-BATCH-EXEC-20200921-1539PM-trans-log.txt
                $Logname=$transcript.replace('-trans-log.txt','-ISEtrans-log.txt') ;
                write-host "`$Logname: $Logname";
                Start-iseTranscript -logname $Logname  -Verbose:($VerbosePreference -eq 'Continue') ;
                #Archive-Log $Logname ;
                # 1:23 PM 4/23/2015 standardize processing file so that we can send a link to open the transcript for review
                $transcript = $Logname ;
                if($host.version.Major -ge 5){ stop-transcript  -Verbose:($VerbosePreference -eq 'Continue')} # ISE in psV5 actually supports transcription. If you don't stop it, it just keeps rolling
            } else {
                write-verbose "$((get-date).ToString('HH:mm:ss')):Stop Transcript" ;
                Stop-TranscriptLog -Transcript $transcript -verbose:$($VerbosePreference -eq "Continue") ;
            } # if-E
            
            # also trailing echo the log:
            $smsg = "`$logging:`$true:written to:`n$($logfile)" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            $smsg = "$((get-date).ToString('HH:mm:ss')):$($sBnr.replace('=v','=^').replace('v=','^='))" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            
            $logging = $false ;  # reset logging for next pass

        } ; # loop-E $mailboxes
    } ;  # PROC-E
    END {
        # =========== wrap up Tenant connections
        # suppress VerbosePreference:Continue, if set, during mod loads (VERY NOISEY)
        if($VerbosePreference = "Continue"){
            $VerbosePrefPrior = $VerbosePreference ;
            $VerbosePreference = "SilentlyContinue" ;
            $verbose = ($VerbosePreference -eq "Continue") ;
        } ;
        if($script:useEXOv2){
            disconnect-exo2 -verbose:$($verbose) ;
        } else {
            disconnect-exo -verbose:$($verbose) ;
        } ;
        # aad mod *does* support disconnect (msol doesen't!)
        #Disconnect-AzureAD -verbose:$($verbose) ;
        # reenable VerbosePreference:Continue, if set, during mod loads
        if($VerbosePrefPrior -eq "Continue"){
            $VerbosePreference = $VerbosePrefPrior ;
            $verbose = ($VerbosePreference -eq "Continue") ;
        } ;

        $pltCleanup=@{
            LogPath=$tmpcopy
            summarizeStatus=$false ;
            NoTranscriptStop=$true ;
            showDebug=$($showDebug) ;
            whatif=$($whatif) ;
            Verbose = ($VerbosePreference -eq 'Continue') ;
        } ;
        $smsg = "_cleanup():w`n$(($pltCleanup|out-string).trim())" ;
        if($verbose){ if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
        #_cleanup -LogPath $tmpcopy ;
        _cleanup @pltCleanup ;
        # prod is still showing a running unstopped transcript, kill it again
        $stopResults = try {Stop-transcript -ErrorAction stop} catch {} ;

        $smsg = "$($sBnrS.replace('-v','-^').replace('v-','^-'))" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

        # clear the script aliases
        write-verbose "clearing ps1* aliases in Script scope" ; 
        get-alias -scope Script |Where-Object{$_.name -match '^ps1.*'} | ForEach-Object{Remove-Alias -alias $_.name} ;

        write-verbose "(explicit EXIT...)" ;
        Break ;

    } ;  # END-E
}

#*------^ test-xoMailbox.ps1 ^------