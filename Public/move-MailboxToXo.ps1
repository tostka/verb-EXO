#*------v move-MailboxToXo.ps1 v------
function move-MailboxToXo{
    <#
    .SYNOPSIS
    move-MailboxToXo.ps1 - EX Hybrid Onprem-> EXO mailbox move
    .NOTES
    Version: 1.1.13
    Author: Todd Kadrie
    Website:	http://www.toddomation.com
    Twitter:	@tostka, http://twitter.com/tostka
    REVISIONS   :
    # 12:41 PM 3/14/2022 sync'd to latest mods of move-EXOmailboxNow, largely rem'ing the xo AD material, long-broken by undocumented fw chgs.
    # 2:49 PM 3/8/2022 pull Requires -modules ...verb-ex2010 ref - it's generating nested errors, when ex2010 requires exo requires ex2010 == loop.
    * 2:40 PM 12/10/2021 more cleanup 
    * 11:24 AM 9/16/2021 encoded eml
    # 10:46 AM 6/2/2021 sub'd verb-logging for v-trans
    * 1:54 PM 5/19/2021 expanded get-hybridcred to use both esvc & sid userroles
    * 11:40 AM 5/14/2021 added -ea 0 to the gv tests (suppresses not-found error when called without logging config)
    # 11:11 AM 5/7/2021 replaced verbose & bdebugs ; fixed missing logging function & added trailing echo of path.
    * 3:55 PM 5/6/2021 update logging, with move to module, it's logging into the ALlusers profile dir
    * 11:05 AM 5/5/2021 refactor into a function, add to verb-exo; ren move-ExoMailboxNow -> move-MailboxToXo
    * 9:54 AM 5/5/2021 removed comment dates; rem'd spurious -domainc param on get-exorecipient
    * 3:15 PM 5/4/2021 v1.1.10 added trailing |?{$_.length} bug workaround for get-gcfastxo.ps1
    * 1:36 PM 3/31/2021 updated rgx to incl jumpboxes; repairs to new cross-org material (subseq to maint-RetPol changes/breaks) ; For some reason nowe requires ; rewrote all connect-reconn's to use central def'd splats AcceptLargeDataLoss in the movereq (wasn't there prev, but worked) ; had shifted out of domaincontroller use (due to new use of get-addomainctontroller returning array's instead of single entries, worked around the issue, put dc's back into queries, to cut down on retry/dawdle. ; updated the re-publish script under description, now echo's requiredvers
    * 9:41 AM 3/30/2021 publishing
    * 9:23 AM 11/17/2020 lots of minor revs, made sure all
    (re)connect-(ex2010|exo)'s had creds, and supported Exov2 switch, replaced
    approp write-verbose -> pswlt blocks ; updated catches to full cred & pswlt
    support ; dbg'd whatif succeessfully. Successfully moved TestNewGenericTodd
    * 9:07 AM 10/26/2020 added EXOv2 support (param, and if/then calls) ;
    * 3:17 PM 10/23/2020 exported xo dc code to new verb-adms:get-gcfastXO() ; now
    fully XO-supporting, determins MigEndPt by cycling $xxxMeta.MepArry
    (semi-colon-delimited string of MEP name;fqdn;regex - regex matches regional
    server DB names and steers the mbx MEP to match).
    * 4:13 PM 10/21/2020 midway through multi-tenant update, left off with:still needs ADMS mount-ADForestDrives() and set-location code @ 395 (had to recode mount-admforestdrives and debug cred production code & infra-string inputs before it would work; will need to dupe to suspend variant on final completion
    * 11:48 AM 4/27/2020 updated for jumpbox/published/installed script, genericized
    * 1:46 PM 2/4/2020 #920 Remove-exoMoveRequest: doesn't like $true on force
    * 12:35 PM 11/26/2019 debugged through issues created by genericizing the domain lookups - still dependant on hard MEP choice mappings on db names - added a default mep chooser that just takes index-0 on the list. Better'n nothing but if CMW has multi-meps, it could pick the wrong one.
    * 12:23 PM 11/22/2019 partial port for x-tenant use, need to get the OPCred spec pulled out of globals too...
    * 7:30 AM 8/14/2019 `$BatchName:revise & add SID
    * 9:38 AM 5/8/2019 added -NoTEST param, to skip MEP tests (no point running 2x, after initial whatif pass)
    * 9:39 AM 1/21/2019 final form used to migrate forwarded mbxs to EXO - RetentPolicy code disabled until we have DL's stocked with population of each target policy, for exempting if they turn up Default DRM Retention Policy again.
    * 8:27 PM 1/18/2019 added -BatchName param, to append later items to an existing Batch, without auto-gening the batchname
    * 7:24 PM 1/18/2019 a LOT of updates, used for forarded mbx moves, 97 - would have used MigrationBatch, but needed to use monitor-exoxxx.ps1 to catch moverequest completion and migrate fowd on the fly
    * 3:59 PM 1/18/2019 spliced in the functions & modloads from convert-OPUserGenerics2Shared.ps1 and the write-log() from check-exolicense.ps1.
    * 10:58 AM 11/27/2018 updated the trailing batch status echo
    * 2:07 PM 11/16/2018 fix typo and incomplete Test-exoMigrationServerAvailability  code, now failure causes it to retsta all MEPs, also removed duped credential prompts
    * 2:19 PM 10/9/2018 correct help text for TargetMailboxes - it's an array, not a comma-delim list
    * 2:49 PM 9/19/2018 added -CompleteEveningOf support and self-determining 6pm/5:30($ADSiteCodeAU) Cutover targets by region/MEP
    * 2:50 PM 9/5/2018 added existingmove pre-removal validated
    * 1:57 PM 9/5/2018 added echo'd replicated-monitoring cmd to the mix, updated move test to check that EXOrecipient has the $TORMeta['o365_TenantDomainMail'] address, or abort move
    * 1:12 PM 8/29/2018 added @$TORMeta['o365_TenantDom'] email addr test and auto-fix, echo of monitoring command to console
    * 1:36 PM 8/27/2018 Added UPN identify (was throwing error: The operation couldn't be performed because 'servicedesk' matches multiple entries.). Also looking up explicit $tmbx obj to pull values from, pretest, supporess prompt for predefined $OPcred ;
    * ident code both move-EXOmailboxNow.ps1 move-EXOmailboxSuspend.ps1 & , only diff is Suspend uses: SuspendWhenReadyToComplete=$TRUE in following;
    $MvSplat=@{
        Identity=$tmbx.userprincipalname ;
        remote=$true;
        RemoteHostName="MEPFQDN.DOMAIN.com" ;
        TargetDeliveryDomain="$TORMeta['o365_TenantDomainMail']" ;
        RemoteCredential=$OPcred ;
        BadItemLimit=1000 ;
        Suspend=$false ;
        SuspendWhenReadyToComplete=$false ;
        BatchName=$Batchname ;
        whatif=$($whatif) ;
    } ;
    .DESCRIPTION
    move-MailboxToXo.ps1 - Non-Suspend Onprem-> EXO mailbox move
    .PARAMETER TenOrg
    TenantTag value, indicating Tenants to connect to[-TenOrg 'TOL']
    .PARAMETER TargetMailboxes
    Mailbox identifiers(array)[-Targetmailboxes]
    .PARAMETER BatchFile
    CSV file of mailbox descriptors, including at least PrimarySMTPAddress field [-BatchFile c:\path-to\file.csv]
    .PARAMETER BatchName
    Hard-code MoveRequest BatchName
    .PARAMETER Suspend
    Suspend move on creation Flag [-Suspend]
    .PARAMETER NoTEST
    NoTest Flag [-NoTEST]
    .PARAMETER Credential
    Credential to use for this connection [-credential [credential obj variable]")][System.Management.Automation.PSCredential]
    .PARAMETER UserRole
    Role of account (SID|CSID|UID|B2BI|CSVC|ESvc|LSvc)[-UserRole SID]
    .PARAMETER useEXOv2
    Use EXOv2 (ExchangeOnlineManagement) over basic auth legacy connection [-useEXOv2]
    .PARAMETER ShowDebug
    Parameter to display Debugging messages [-ShowDebug switch]
    .PARAMETER whatif
    Whatif Flag (DEFAULTED TRUE!) [-whatIf]
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    None. Returns no objects or output.
    .EXAMPLE
    move-MailboxToXo.ps1 -TargetMailboxes ACCOUNT@COMPANY.com -showDebug  -whatIf ;
    Perform immediate move of specified mailbox, with debug output & whtif pass
    .EXAMPLE
    move-MailboxToXo.ps1 -TargetMailboxes ACCOUNT@COMPANY.com -showDebug -notest -whatIf ;
    Perform immediate move of specified mailbox, suppress MEP tests (-NoTest), showdebug output & whatif pass
    .LINK
    #>
    ##Requires -Modules ActiveDirectory, ExchangeOnlineManagement, verb-ADMS, verb-Ex2010, verb-IO, verb-logging, verb-Mods, verb-Network, verb-Text, verb-logging
    # 2:49 PM 3/8/2022 pull verb-ex2010 ref - I think it's generating nested errors, when ex2010 requires exo requires ex2010 == loop.
    #Requires -Modules ActiveDirectory, ExchangeOnlineManagement, verb-ADMS,verb-IO, verb-logging, verb-Mods, verb-Network, verb-Text, verb-logging
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$FALSE,HelpMessage="TenantTag value, indicating Tenants to connect to[-TenOrg 'TOL']")]
        [ValidateNotNullOrEmpty()]
        $TenOrg = 'TOR',
        [Parameter(Position=0,Mandatory=$False,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,HelpMessage="Mailbox identifiers(array)[-Targetmailboxes]")]
        [ValidateNotNullOrEmpty()]$TargetMailboxes,
        [Parameter(Mandatory=$false,HelpMessage="CSV file of mailbox descriptors, including at least PrimarySMTPAddress field [-BatchFile c:\path-to\file.csv]")]
        [ValidateScript({Test-Path $_})][string]$BatchFile,
        [Parameter(Position=0,HelpMessage="Hard-code MoveRequest BatchName")]
        [string]$BatchName,
        [Parameter(HelpMessage="Suspend move on creation Flag [-Suspend]")]
        [switch] $Suspend,
        [Parameter(HelpMessage="NoTest Flag [-NoTEST]")]
        [switch] $NoTEST,
        [Parameter(HelpMessage="Credential to use for cloud actions [-credential [credential obj variable]")][System.Management.Automation.PSCredential]
        $Credential,
        # = $global:$credO365TORSID,
        [Parameter(HelpMessage = "Role of account (SID|CSID|UID|B2BI|CSVC|ESvc|LSvc)[-UserRole SID]")]
        [ValidateSet('SID','CSID','UID','B2BI','CSVC','ESVC','LSVC')]
        [string]$UserRole='SID',
        [Parameter(HelpMessage="Use EXOv2 (ExchangeOnlineManagement) over basic auth legacy connection [-useEXOv2]")]
        [switch] $useEXOv2,
        [Parameter(HelpMessage="Unpromtped run Flag [-showDebug]")]
        [switch] $NoPrompt,
        [Parameter(HelpMessage="Debugging Flag [-showDebug]")]
        [switch] $showDebug,
        [Parameter(HelpMessage="Whatif Flag (DEFAULTED TRUE!) [-whatIf]")]
        [switch] $whatIf=$true
    ) # PARAM BLOCK END

    $verbose = ($VerbosePreference -eq "Continue") ;

    #region INIT; # ------
    #*======v SCRIPT/DOMAIN/MACHINE/INITIALIZATION-DECLARE-BOILERPLATE v======
    # SCRIPT-CONFIG MATERIAL TO SET THE UNDERLYING $DBGPREF:
    if ($Whatif) { write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):`$Whatif is TRUE (`$whatif:$($whatif))" ; };
    # If using WMI calls, push any cred into WMI:
    #if ($Credential -ne $Null) {$WmiParameters.Credential = $Credential }  ;

    # scriptname with extension
    $ScriptDir=(Split-Path -parent $MyInvocation.MyCommand.Definition) + "\" ;
    $ScriptBaseName = (Split-Path -Leaf ((&{$myInvocation}).ScriptName))  ;
    $ScriptNameNoExt = [system.io.path]::GetFilenameWithoutExtension($MyInvocation.InvocationName) ;
    $ComputerName = $env:COMPUTERNAME ;

    # build smtpfrom & to fr Meta
    $smtpFrom = (($scriptBaseName.replace(".","-")) + "@$((Get-Variable  -name "$($TenOrg)Meta").value.o365_OPDomain)") ;
    #$smtpSubj= ("Daily Rpt: "+ (Split-Path $transcript -Leaf) + " " + [System.DateTime]::Now) ;
    $smtpSubj= "Proc Rpt:"   ;
    if($whatif) {
        $smtpSubj+="WHATIF:" ;
    } else {
        $smtpSubj+="PROD:" ;
    } ;
    $smtpSubj+= "$($ScriptBaseName):$(get-date -format 'yyyyMMdd-HHmmtt')"   ;
    #$smtpTo=$TORMeta['NotificationAddr1'] ;
    #$smtpTo=$TORMeta['NotificationDlUs'] ;
    $smtpToFailThru="dG9kZC5rYWRyaWVAdG9yby5jb20="| convertFrom-Base64String ; 
    # one bene of looping: no module dependancy, works before modloads occur
    # pull the notifc smtpto from the xxxMeta.NotificationDlUs value
    # non-looping - $TenOrg is an input param, does't need modules to work yet
    if(!$showdebug){
        if((Get-Variable  -name "$($TenOrg)Meta").value.NotificationDlUs){
            $smtpTo = (Get-Variable  -name "$($TenOrg)Meta").value.NotificationDlUs ;
        }elseif((Get-Variable  -name "$($TenOrg)Meta").value.NotificationAddr1){
            $smtpTo = (Get-Variable  -name "$($TenOrg)Meta").value.NotificationAddr1 ;
        } else {
            $smtpTo=$smtpToFailThru;
        } ;
    } else {
        # debug pass, don't send to main dl, use NotificationAddr1    if((Get-Variable  -name "$($TenOrg)Meta").value.NotificationDlUs){
        if((Get-Variable  -name "$($TenOrg)Meta").value.NotificationAddr1){
            #set-variable -Name $meta.name -Value ((get-variable -name $meta.name).value  += @{'o365_AcceptedDomains' = (Get-exoAcceptedDomain).domainname} )
            $smtpTo = (Get-Variable  -name "$($TenOrg)Meta").value.NotificationAddr1 ;
        } else {
            $smtpTo=$smtpToFailThru ;
        } ;
    }

    $sQot = [char]34 ; $sQotS = [char]39 ;
    $NoProf=[bool]([Environment]::GetCommandLineArgs() -like '-noprofile'); # if($NoProf){# do this};

    #$ProgInterval= 500 ; # write-progress wait interval in ms
    # add gui vb prompt support
    #[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null ;
    # should use Windows.Forms where possible, more stable

    switch -regex ($env:COMPUTERNAME) {
        ($rgxMyBox) { $LocalInclDir = "c:\usr\work\exch\scripts" ; }
        ($rgxProdEx2010Servers) { $LocalInclDir = "c:\scripts" ; }
        ($rgxLabEx2010Servers) { $LocalInclDir = "c:\scripts" ; }
        ($rgxProdL13Servers) { $LocalInclDir = "c:\scripts" ; }
        ($rgxLabL13Servers) { $LocalInclDir = "c:\scripts" ; }
        ($rgxAdminJumpBoxes) {
            $LocalInclDir = (split-path $profile) ;
        }
    } ;

    #configure EXO EMS aliases to cover useEXOv2 requirements
    switch ($script:useEXOv2){
        $true {
            write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):using ExoV2 cmdlets" ;
            #reconnect-eXO2 @pltRXO ;
            set-alias ps1GetXRcp get-xorecipient ;
            set-alias ps1GetXMbx get-xomailbox ;
            set-alias ps1SetXMbx Set-xoMailbox ;
            set-alias ps1GetxUser get-xoUser ;
            set-alias ps1GetXCalProc get-xoCalendarprocessing ;
            set-alias ps1GetXMbxFldrPerm get-xoMailboxfolderpermission ;
            set-alias ps1GetXAccDom Get-xoAcceptedDomain ;
            set-alias ps1GGetXRetPol Get-xoRetentionPolicy ;
            set-alias ps1GetXDistGrp get-xoDistributionGroup ;
            set-alias ps1GetXDistGrpMbr get-xoDistributionGroupmember ;
            set-alias ps1TestXMigrSrvrAvail Test-xoMigrationServerAvailability ;
            set-alias ps1GetXMovReq get-xomoverequest ;
            set-alias ps1RmvXMovReq  Remove-xoMoveRequest ;
            set-alias ps1NewXMovReq  New-xoMoveRequest ;
            set-alias ps1GetXMovReqStats Get-xoMoveRequestStatistics
        }
        $false {
            write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):using EXO cmdlets" ;
            #reconnect-exo @pltRXO
            set-alias ps1GetXRcp get-exorecipient ;
            set-alias ps1GetXMbx get-exomailbox ;
            set-alias ps1SetXMbx Set-exoMailbox ;
            set-alias ps1GetxUser get-exoUser ;
            set-alias ps1GetXCalProc get-exoCalendarprocessing  ;
            set-alias ps1GetXMbxFldrPerm get-exoMailboxfolderpermission  ;
            set-alias ps1GetXAccDom Get-exoAcceptedDomain ;
            set-alias ps1GGetXRetPol Get-exoRetentionPolicy
            set-alias ps1GetXDistGrp get-exoDistributionGroup  ;
            set-alias ps1GetXDistGrpMbr get-exoDistributionGroupmember ;
            set-alias ps1TestXMigrSrvrAvail Test-exoMigrationServerAvailability ;
            set-alias ps1GetXMovReq get-exomoverequest ;
            set-alias ps1RmvXMovReq  Remove-exoMoveRequest ;
            set-alias ps1NewXMovReq  New-exoMoveRequest ;
            set-alias ps1GetXMovReqStats Get-exoMoveRequestStatistics
        } ;
    } ;  # SWTCH-E useEXOv2

    $Retries = 4 ;
    $RetrySleep = 5 ;
    $DawdleWait = 30 ; # wait time (secs) between dawdle checks
    $DirSyncInterval = 30 ; # AADConnect dirsync interval
    if(!$ThrottleMs){$ThrottleMs = 50 ;} ;
    $CredRole = 'CSVC' ; # role of svc to be dyn pulled from metaXXX if no -Credential spec'd,

    #$LastDays = -3 ;
    #if($LastDays -gt 0){$LastDays = $LastDays * -1 } ; # flip non-negative to negative integer, updated *-1 in the usage line

    if(!$rgxDriveBanChars){$rgxDriveBanChars = '[;~/\\\.:]' ; } ; # ;~/\.:,

    #*======v FUNCTIONS v======

    #-=-=TEMP SUPPRESS VERBOSE-=-=-=-=-=-=
    # suppress VerbosePreference:Continue, if set, during mod loads (VERY NOISEY)
    if($VerbosePreference = "Continue"){
        $VerbosePrefPrior = $VerbosePreference ;
        $VerbosePreference = "SilentlyContinue" ;
        $verbose = ($VerbosePreference -eq "Continue") ;
    } ;
    <# flip to require specs over tmod loads
    #*------v  MOD LOADS  v------
    # strings are: "[tModName];[tModFile];tModCmdlet"
    $tMods = @() ;
    $tMods+="verb-Auth;C:\sc\verb-Auth\verb-Auth\verb-Auth.psm1;get-password" ;
    $tMods+="verb-logging;C:\sc\verb-logging\verb-logging\verb-logging.psm1;write-log";
    $tMods+="verb-IO;C:\sc\verb-IO\verb-IO\verb-IO.psm1;Add-PSTitleBar" ;
    $tMods+="verb-Mods;C:\sc\verb-Mods\verb-Mods\verb-Mods.psm1;check-ReqMods" ;
    $tMods+="verb-Text;C:\sc\verb-Text\verb-Text\verb-Text.psm1;Remove-StringDiacritic" ;
    #$tMods+="verb-Desktop;C:\sc\verb-Desktop\verb-Desktop\verb-Desktop.psm1;Speak-words" ;
    $tMods+="verb-dev;C:\sc\verb-dev\verb-dev\verb-dev.psm1;Get-CommentBlocks" ;
    $tMods+="verb-Network;C:\sc\verb-Network\verb-Network\verb-Network.psm1;Send-EmailNotif" ;
    $tMods+="verb-Automation.ps1;C:\sc\verb-Automation.ps1\verb-Automation.ps1\verb-Automation.ps1.psm1;Retry-Command" ;
    #$tMods+="verb-AAD;C:\sc\verb-AAD\verb-AAD\verb-AAD.psm1;Build-AADSignErrorsHash";
    $tMods+="verb-ADMS;C:\sc\verb-ADMS\verb-ADMS\verb-ADMS.psm1;load-ADMS";
    $tMods+="verb-Ex2010;C:\sc\verb-Ex2010\verb-Ex2010\verb-Ex2010.psm1;Connect-Ex2010";
    $tMods+="verb-EXO;C:\sc\verb-EXO\verb-EXO\verb-EXO.psm1;Connect-Exo";
    #$tMods+="verb-L13;C:\sc\verb-L13\verb-L13\verb-L13.psm1;Connect-L13";
    #$tMods+="verb-Teams;C:\sc\verb-Teams\verb-Teams\verb-Teams.psm1;Connect-Teams";
    #$tMods+="verb-SOL;C:\sc\verb-SOL\verb-SOL\verb-SOL.psm1;Connect-SOL" ;
    #$tMods+="verb-Azure;C:\sc\verb-Azure\verb-Azure\verb-Azure.psm1;get-AADBearToken" ;
    foreach($tMod in $tMods){
    $tModName = $tMod.split(';')[0] ; $tModFile = $tMod.split(';')[1] ; $tModCmdlet = $tMod.split(';')[2] ;
    $smsg = "( processing `$tModName:$($tModName)`t`$tModFile:$($tModFile)`t`$tModCmdlet:$($tModCmdlet) )" ;
    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
    if($tModName -eq 'verb-Network' -OR $tModName -eq 'verb-Azure'){
        #write-host "GOTCHA!" ;
    } ;
    $lVers = get-module -name $tModName -ListAvailable -ea 0 ;
    if($lVers){   $lVers=($lVers | sort version)[-1];   try {     import-module -name $tModName -RequiredVersion $lVers.Version.tostring() -force -DisableNameChecking -Verbose:$false  } catch {     write-warning "*BROKEN INSTALLED MODULE*:$($tModName)`nBACK-LOADING DCOPY@ $($tModDFile)" ;import-module -name $tModDFile -force -DisableNameChecking -verbose:$false  } ;
    } elseif (test-path $tModFile) {
        write-warning "*NO* INSTALLED MODULE*:$($tModName)`nBACK-LOADING DCOPY@ $($tModDFile)" ;
        try {import-module -name $tModDFile -force -DisableNameChecking -Verbose:$false} # force non-verbose, suppress spam
        catch {   write-error "*FAILED* TO LOAD MODULE*:$($tModName) VIA $(tModFile) !" ;   $tModFile = "$($tModName).ps1" ;   $sLoad = (join-path -path $LocalInclDir -childpath $tModFile) ;   if (Test-Path $sLoad) {       Write-Verbose -verbose ((Get-Date).ToString("HH:mm:ss") + "LOADING:" + $sLoad) ;       . $sLoad ;       if ($showdebug) { Write-Verbose -verbose "Post $sLoad" };   } else {       $sLoad = (join-path -path $backInclDir -childpath $tModFile) ;       if (Test-Path $sLoad) {           Write-Verbose -verbose ((Get-Date).ToString("HH:mm:ss") + "LOADING:" + $sLoad) ;           . $sLoad ;           if ($showdebug) { Write-Verbose -verbose "Post $sLoad" };       } else {           Write-Warning ((Get-Date).ToString("HH:mm:ss") + ":MISSING:" + $sLoad + " EXITING...") ;           exit;       } ;   } ; } ;
    } ;
    if(!(test-path function:$tModCmdlet)){
        write-warning -verbose:$true  "UNABLE TO VALIDATE PRESENCE OF $tModCmdlet`nfailing through to `$backInclDir .ps1 version" ;
        $sLoad = (join-path -path $backInclDir -childpath "$($tModName).ps1") ;
        if (Test-Path $sLoad) {     Write-Verbose -verbose:$true ((Get-Date).ToString("HH:mm:ss") + "LOADING:" + $sLoad) ;     . $sLoad ;     if ($showdebug) { Write-Verbose -verbose "Post $sLoad" };     if(!(test-path function:$tModCmdlet)){         write-warning "$((get-date).ToString('HH:mm:ss')):FAILED TO CONFIRM `$tModCmdlet:$($tModCmdlet) FOR $($tModName)" ;     } else {          write-verbose -verbose:$true  "(confirmed $tModName loaded: $tModCmdlet present)"     }
        } else {     Write-Warning ((Get-Date).ToString("HH:mm:ss") + ":MISSING:" + $sLoad + " EXITING...") ;     exit; } ;
    } else {     write-verbose -verbose:$true  "(confirmed $tModName loaded: $tModCmdlet present)" } ;
    if($tModName -eq 'verb-logging'){

            # if($PSCommandPath){   $logspec = start-Log -Path $PSCommandPath -NoTimeStamp -Tag LASTPASS -showdebug:$($showdebug) -whatif:$($whatif) ;
#             } else {    $logspec = start-Log -Path ($MyInvocation.MyCommand.Definition) -showdebug:$($showdebug) -whatif:$($whatif) ; } ;
#             if($logspec){
#                 $stopResults = try {Stop-transcript -ErrorAction stop} catch {} ;
#                 $logging=$logspec.logging ;
#                 $logfile=$logspec.logfile ;
#                 $transcript=$logspec.transcript ;
#                 #Configure default logging from parent script name
#                 #start-transcript -Path $transcript ;
#             } else {throw "Unable to configure logging!" } ;


    } ;
    } ;  # loop-E
    #*------^ END MOD LOADS ^------
    #>
    #-=-=-=-=RE-ENABLE PRIOR VERBOSE-=-=-=-=
    # reenable VerbosePreference:Continue, if set, during mod loads
    if($VerbosePrefPrior -eq "Continue"){
        $VerbosePreference = $VerbosePrefPrior ;
        $verbose = ($VerbosePreference -eq "Continue") ;
    } ;
    #-=-=-=-=-=-=-=-=

    #*------v Function check-ReqMods  v------
    function check-ReqMods ($reqMods){    $bValidMods=$true ;    $reqMods | foreach-object {        if( !(test-path function:$_ ) ) {          write-error "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Missing $($_) function." ;          $bValidMods=$false ;        }    } ;    write-output $bValidMods ;} ;
    #*------^ END Function check-ReqMods  ^------

    #*======^ END FUNCTIONS ^======

    #*======v SUB MAIN v======

    # email trigger vari, it will be semi-delimd list of mail-triggering events
    $script:PassStatus = $null ;

    # check for $TenOrg & credential alignment
    # with credential un-defaulted, no need to compare $TenOrg & credential
    $tvari = "PassStatus_$($tenorg)" ; if(get-Variable -Name $tvari -scope Script -ea 0){Remove-Variable -Name $tvari -scope Script}
    New-Variable -Name PassStatus_$($tenorg) -scope Script -Value $null ;

    $reqMods+="Add-PSTitleBar;Remove-PSTitleBar".split(";") ;
    #Disconnect-EMSR (variant name in some ps1's for Disconnect-Ex2010)
    #$reqMods+="Reconnect-CCMS;Connect-CCMS;Disconnect-CCMS".split(";") ;
    #$reqMods+="Reconnect-SOL;Connect-SOL;Disconnect-SOL".split(";") ;
    $reqMods+="Test-TranscriptionSupported;Test-Transcribing;Stop-TranscriptLog;Start-IseTranscript;Start-TranscriptLog;get-ArchivePath;Archive-Log;Start-TranscriptLog".split(";") ;
    # add verb-automation content
    $reqMods+="retry-command".split(";") ;
    # lab, fails wo
    $reqMods+="Load-EMSSnap" ;
    # remove dupes
    $reqMods=$reqMods| select -Unique ;

    # detect profile installs (installed mod or script), and redir to stock location
            $dPref = 'd','c' ; foreach($budrv in $dpref){ if(test-path -path "$($budrv):\scripts" -ea 0 ){ break ;  } ;  } ;
            [regex]$rgxScriptsModsAllUsersScope="^$([regex]::escape([environment]::getfolderpath('ProgramFiles')))\\((Windows)*)PowerShell\\(Scripts|Modules)" ;
            [regex]$rgxScriptsModsCurrUserScope="^$([regex]::escape([environment]::getfolderpath('Mydocuments')))\\((Windows)*)PowerShell\\(Scripts|Modules)" ;
            # -Tag "($TenOrg)-LASTPASS" 
            # Tag=$lTag 
            $pltSLog = [ordered]@{ NoTimeStamp=$false ;  showdebug=$($showdebug) ;whatif=$($whatif) ;} ;
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
            $logspec = start-Log @pltSLog 

    # reloc batch calc, to include the substring in the log/transcript
    if(!$BatchName){
        $BatchName = "ExoMoves-$($env:USERNAME)" ;
        # include the 1st TargetMailbox fr the param, use first 12 chars (or less)
        $BatchName += '-' + @($TargetMailboxes)[0].tostring().substring(0,[System.Math]::Min(12,@($TargetMailboxes)[0].length)) ;
        if( (@($TargetMailboxes)|measure).count -gt 1 ){
            # append an ellipses to indicate multiple mbxs moved
            $BatchName += '...' ;
        } ;
        $BatchName += "-$(get-date -format 'yyyyMMdd-HHmmtt')" ;
        $smsg= "Using Dynamic BatchName:$($BatchName)" ;
    } else {
        $smsg= "Using -BatchName:$($BatchName)" ;
    } ;
    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

    if($logspec){
        $stopResults = try {Stop-transcript -ErrorAction stop} catch {} ;
        $logging=$logspec.logging ;
        $logfile=$logspec.logfile ;
        $transcript=$logspec.transcript ;
        #Configure default logging from parent script name
        # logfile                        C:\usr\work\o365\scripts\logs\move-MailboxToXo-(TOR)-LASTPASS-LOG-BATCH-WHATIF-log.txt
        # transcript                     C:\usr\work\o365\scripts\logs\move-MailboxToXo-(TOR)-LASTPASS-Transcript-BATCH-WHATIF-trans-log.txt
        if($Ticket){
            $logfile=$logfile.replace("-BATCH","-$($Ticket)-BATCH") ;
            $transcript=$transcript.replace("-BATCH","-$($Ticket)-BATCH") ;
        } else {
            $logfile=$logfile.replace("-BATCH","-nnnnnn") ;
            $transcript=$transcript.replace("-BATCH","-nnnnnn") ;
        } ;
        $logfile = $logfile.replace('-LASTPASS','').replace('BATCH',(Remove-InvalidFileNameChars -name $BatchName )) ;
        $transcript = $transcript.replace('-LASTPASS','').replace('BATCH',(Remove-InvalidFileNameChars -name $BatchName )) ;
        if(Test-TranscriptionSupported){start-transcript -Path $transcript }
        else { write-warning "$($host.name) v$($host.version.major) does *not* support Transcription!" } ;
    } else {throw "Unable to configure logging!" } ;

    $smsg= "#*======v START PASS:$($ScriptBaseName) v======" ;
    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
    
    <# disable cross-org handling, post tenant migr, no need
    # seeing Curly & Cheech turn up in EX10 queries, pre-purge *any* AD psdrive
    if($existingADPSDrives = get-psdrive -PSProvider ActiveDirectory -ea 0){
        $smsg = "Purging *existing* AD PSDrives found:$(($existingADPSDrives| ft -auto name,provider,root,globalcatalog|out-string).trim())" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        $error.clear() ;
        TRY {
            $existingADPSDrives | remove-psdrive -Verbose:$($verbose) # -WhatIf:$($whatif) ;
        } CATCH {
            Write-Warning "$(get-date -format 'HH:mm:ss'): Failed processing $($_.Exception.ItemName). `nError Message: $($_.Exception.Message)`nError Details: $($_)" ;
            $smsg = "FULL ERROR TRAPPED (EXPLICIT CATCH BLOCK WOULD LOOK LIKE): } catch[$($Error[0].Exception.GetType().FullName)]{" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            BREAK ;
        } ;

    } ;
    # also purge the $global:ADPsDriveNames or $script:ADPsDriveNames
    if(gv -name ADPsDriveNames -scope global -ea 0){
        $error.clear() ;
        TRY {
            Remove-Variable -name ADPsDriveNames -Scope Global ;
        } CATCH {
            Write-Warning "$(get-date -format 'HH:mm:ss'): Failed processing $($_.Exception.ItemName). `nError Message: $($_.Exception.Message)`nError Details: $($_)" ;
            $smsg = "FULL ERROR TRAPPED (EXPLICIT CATCH BLOCK WOULD LOOK LIKE): } catch[$($Error[0].Exception.GetType().FullName)]{" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            BREAK ;
        } ;
    } ;
    if(gv -name ADPsDriveNames -scope script -ea 0){
        $error.clear() ;
        TRY {
            Remove-Variable -name ADPsDriveNames -Scope script ;
        } CATCH {
            Write-Warning "$(get-date -format 'HH:mm:ss'): Failed processing $($_.Exception.ItemName). `nError Message: $($_.Exception.Message)`nError Details: $($_)" ;
            $smsg = "FULL ERROR TRAPPED (EXPLICIT CATCH BLOCK WOULD LOOK LIKE): } catch[$($Error[0].Exception.GetType().FullName)]{" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            BREAK ;
        } ;
    } ;
    #>
    # $XXXMeta.ExOPAccessFromToro & Ex10Server
    # steer all onprem code
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

    # TEST
    if( !(check-ReqMods $reqMods) ) {write-error "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Missing function. EXITING." ; exit ;}  ;

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
        if($o365Cred=(get-TenantCredentials -TenOrg $TenOrg -UserRole 'CSVC','SID' -verbose:$($verbose))){
            # make it script scope, so we don't have to predetect & purge before using new-variable - except now it does [headcratch]
            $tvari = "cred$($tenorg)" ; if(get-Variable -Name $tvari -scope Script){Remove-Variable -Name $tvari -scope Script}
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
            exit ;
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
        $pltGHOpCred=@{TenOrg=$TenOrg ;userrole='ESVC','SID'; verbose=$($verbose)} ;
        if($OPCred=(get-HybridOPCredentials @pltGHOpCred).cred){
            # make it script scope, so we don't have to predetect & purge before using new-variable
            $tvari = "cred$($tenorg)OP" ; if(get-Variable -Name $tvari -scope Script){Remove-Variable -Name $tvari -scope Script} ;
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
            exit ;
        } ;
        $smsg= "Using EXOP cred:`$cred$($tenorg)OP:$((Get-Variable -name "cred$($tenorg)OP" ).value.username)" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        <# CALLS ARE IN FORM: (cred$($tenorg))
        $pltRX10 = @{
            Credential = (Get-Variable -name "cred$($tenorg)OP" ).value ;
            verbose = $($verbose) ; }
        Reconnect-Ex2010 @pltRX10 ; # local org conns
        #$pltRx10 creds & .username can also be used for local ADMS connections
        #>
        $pltRX10 = @{
            Credential = (Get-Variable -name "cred$($tenorg)OP" ).value ;
            verbose = $($verbose) ; } ;
        # TEST
        if( !(check-ReqMods $reqMods) ) {write-error "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Missing function. EXITING." ; exit ;}  ;
        # defer cx10/rx10, until just before get-recipients qry
        #*------^ END GENERIC EXOP CREDS & SRVR CONN BP ^------
        # connect to ExOP X10
        <#
        if($pltRX10){
            ReConnect-Ex2010 @pltRX10 ;
        } else { Reconnect-Ex2010 ; } ;
        #>
    } ;  # if-E $useEXOP


    $smsg= "Using EXOP cred:$($pltRXO.Credential.username)" ;
    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

    $smsg= "Using OPCred cred:$($pltRX10.Credential.username)" ;
    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

    if($VerbosePreference = "Continue"){
        $VerbosePrefPrior = $VerbosePreference ;
        $VerbosePreference = "SilentlyContinue" ;
        $verbose = ($VerbosePreference -eq "Continue") ;
    } ;
    if($pltRX10){ReConnect-Ex2010 @pltRX10 }
    else { Reconnect-Ex2010 ; } ;
    if($VerbosePrefPrior -eq "Continue"){
        $VerbosePreference = $VerbosePrefPrior ;
        $verbose = ($VerbosePreference -eq "Continue") ;
    } ;

    # load ADMS
    $reqMods+="load-ADMS".split(";") ;
    if( !(check-ReqMods $reqMods) ) {write-error "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Missing function. EXITING." ; exit ;}  ;
    $smsg = "(loading ADMS...)" ;
    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

    load-ADMS ;

    # multi-org AD
    <#still needs ADMS mount-ADForestDrives() and set-location code @ 395 (had to recode mount-admforestdrives and debug cred production code & infra-string inputs before it would work; will need to dupe to suspend variant on final completion
    #>
    <# 12:25 PM 3/14/2022 disable no x-org
    if(!$global:ADPsDriveNames){
        $smsg = "(connecting X-Org AD PSDrives)" ;
        if($verbose){ if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
        $global:ADPsDriveNames = mount-ADForestDrives -verbose:$($verbose) ;
    } ;
    #>
    # EXO connection


    #$reqMods+="connect-exo;Reconnect-exo;Disconnect-exo".split(";") ;
    if( !(check-ReqMods $reqMods) ) {write-error "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Missing function. EXITING." ; exit ;}  ;
    $smsg = "(loading EXO...)" ;
    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
    #reconnect-exo -credential $pltRXO.Credential ;
    if ($script:useEXOv2) { reconnect-eXO2 @pltRXO }
    else { reconnect-EXO @pltRXO } ;


    <# RLMS connection
    $reqMods+="Get-LyncServerInSite;load-LMS;Disconnect-LMSR".split(";") ;
    if( !(check-ReqMods $reqMods) ) {write-error "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Missing function. EXITING." ; exit ;}  ;
    write-verbose -verbose:$true  "(loading LMS...)" ;
    Reconnect-L13 ;
    #>

    # forestdom is used to build the expected Tenant onmicrosoft address filter
    #$script:forestdom=((get-adforest | select -expand upnsuffixes) |?{$_ -eq (Get-Variable  -name "$($TenOrg)Meta").value.o365_OPDomain}) ;
    #-=-get-gcfastXO use to pull a root domain (or enable-exforestview for OPEX)=-=-=-=-=-
    if($UseOP){
        # suppress VerbosePreference:Continue, if set, during mod loads (VERY NOISEY)
        if($VerbosePreference -eq "Continue"){
            $VerbosePrefPrior = $VerbosePreference ;
            $VerbosePreference = "SilentlyContinue" ;
            $verbose = ($VerbosePreference -eq "Continue") ;
        } ;
        # we need $domaincontroller pop'd whether useCached or not
        # connect to ExOP X10 4:02 PM 3/23/2021 this SHOULDN'T BE A RECONNECT! should be a cold connect, *with* a pre-disconnect! I just watched it skid from TOR to CMW wo dropping connect!
        if($pltRX10){
            #Disconnect-Ex2010 -verbose:$($verbose) ;
            #get-pssession | remove-pssession ;
            # if it's not a multi-tenant process, don't pre-disconnect
            $smsg = "reconnect-Ex2010 w`n$(($pltRX10|out-string).trim())" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            ReConnect-Ex2010 @pltRX10 ;
        } else { connect-Ex2010 ; } ;
        if($VerbosePrefPrior -eq "Continue"){
            $VerbosePreference = $VerbosePrefPrior ;
            $verbose = ($VerbosePreference -eq "Continue") ;
        } ;
        # pre-clear dc, before querying
        $domaincontroller = $null ;
        <#
        # we don't know which subdoms may be in play
        pushd ; # cache pwd
        if( $tPsd = "$((Get-Variable  -name "$($TenOrg)Meta").value.ADForestName -replace $rgxDriveBanChars):" ){
            if(test-path $tPsd){
                $error.clear() ;
                TRY {
                    set-location -Path $tPsd -ea STOP ;
                    $objForest = get-adforest ;
                    $doms = @($objForest.Domains) ; # ad mod get-adforest vers
                    # do simple detect 2 doms (parent & child), use child (non-parent dom):
                    if(($doms|?{$_ -ne $objforest.name}|measure).count -eq 1){
                        $subdom = $doms|?{$_ -ne $objforest.name} ;
                        $domaincontroller = get-gcfastxo -TenOrg $TenOrg -Subdomain $subdom -verbose:$($verbose) |?{$_.length};
                        $smsg = "get-gcfastxo:returned $($domaincontroller)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } else {
                        # as this is just EX, and not AD search, open the forestview up - all Ex OP qrys will search entire forest
                        enable-forestview
                        $domaincontroller = $null ;
                    } ;
                    $script:forestdom=(($objForest | select -expand upnsuffixes) |?{$_ -eq (Get-Variable  -name "$($TenOrg)Meta").value.o365_OPDomain}) ;
                } CATCH {
                    $ErrTrapd=$Error[0] ;
                    $smsg= "Failed to exec cmd because: $($ErrTrapd)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    $smsg = "FULL ERROR TRAPPED (EXPLICIT CATCH BLOCK WOULD LOOK LIKE): } catch[$($ErrTrapd.Exception.GetType().FullName)]{" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    popd ; # restore dir
                    Exit #Opts: STOP(debug)|EXIT(close)|CONTINUE(move on in loop cycle)|BREAK(exit loop iteration)|THROW $_/'CustomMsg'(end script with Err output)
                } ;
            } else {
                $smsg = "UNABLE TO FIND *MOUNTED* AD PSDRIVE $($Tpsd) FROM `$$($TENorg)Meta!" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug
                else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                Exit #Opts: STOP(debug)|EXIT(close)|CONTINUE(move on in loop cycle)|BREAK(exit loop iteration)|THROW $_/'CustomMsg'(end script with Err output)
            } ;
        } else {
            $smsg = "UNABLE TO RESOLVE PROPER AD PSDRIVE FROM `$$($TENorg)Meta!" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug
            else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            Exit #Opts: STOP(debug)|EXIT(close)|CONTINUE(move on in loop cycle)|BREAK(exit loop iteration)|THROW $_/'CustomMsg'(end script with Err output)
        } ;
        popd ; # cd to prior dir
        #>
        
    } ; # $useOP
    #-=-=-=-=-=-=-

    # move the domaincontroller lookup code down into the per-user, user's subdomain is required to run get-aduser or Exch cmdlets to find the object.

    #-=-NON-SUSPEND, COMPLETE 1ST PASS:SuspendWhenReadyToComplete=$false, launch Suspend=$false=-=-=-=-=-=-=

    # fall back manual prompt creds
    <#
    if($TORMeta['o365_SIDUpn'] -AND $TORMeta['logon_SID']){
        $EXOMoveID=$TORMeta['o365_SIDUpn'] # EXO UPN-based admin sid
        $OPMoveID=$TORMeta['logon_SID'] ; # # ONPrem EX legacy fmt admin sid
    } ;
    #>
    # tenant & hyb creds already gotten above: EXO:$pltRXO.Credential & Ex10/AD:$pltRX10.Credential

    if(!$ThrottleMs){$ThrottleMs = 500 } ;#amt to throttle per pass, to *try* to stay out of throttling

    write-host "`a`n" ;
    write-host -foregroundcolor green "DID YOU PRECONVERT SECURITY GRPS & DUMP ACLS!?" ;
    write-host "`a`n" ;
    if($NoPrompt){
        $smsg = "(-NoPrompt: skipping interactive)" ;
        if($verbose){ if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
        $bRet = 'YYY'
    }else{
        $bRet=Read-Host "Enter YYY to continue. Anything else will exit"
    } ;
    if ($bRet.ToUpper() -eq "YYY") {
        Write-host "Moving on"
    } else {
        Write-Host "Invalid response. Exiting"
        # exit <asserted exit error #>
        exit 1
    } # if-block end

    <# timezone standards:
    $cstzone = [System.TimeZoneInfo]::FindSystemTimeZoneById("Central Standard Time") ;
    $AUCSTzone = [System.TimeZoneInfo]::FindSystemTimeZoneById("Cen. Australia Standard Time") ;
    $GMTzone = [System.TimeZoneInfo]::FindSystemTimeZoneById("GMT Standard Time") ;
    #>

    <#
    if($global:credo365TORSID){$exocred = $global:credo365TORSID}
    else {$exocred = Get-Credential -credential $EXOMoveID } ;

    if($global:credTORSID){$pltRX10.Credential = $global:credTORSID}
    else {$pltRX10.Credential = Get-Credential -credential $OPMoveID} ;
    #>
    if($pltRX10){
        ReConnect-Ex2010 @pltRX10 ;
    } else { Reconnect-Ex2010 ; } ;


    if($BatchFile){
        $smsg= "Using -BatchName:$($BatchName)" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        $TargetMailboxes=import-csv -path $BatchFile | select -expand PrimarySMTPAddress;
    } elseif ($Targetmailboxes){
        # defer to parameter version
    }else {
        $smsg= "MISSING `$BATCHFILE, ABORTING!" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        _cleanup
    } ;  ;

    # moved batchname calc up to logging area (to include in log)

    #if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
    #else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

    $ttl=($TargetMailboxes|measure).count ;
    $Procd=0 ;

    foreach($tMbxId in $TargetMailboxes){
        $Procd++ ;
        $sBnr="#*======v `$tmbx:($($Procd)/$($ttl)):$($tMbxId) v======" ;
        $smsg="$($sBnr)" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

        if ($script:useEXOv2) { reconnect-eXO2 @pltRXO }
        else { reconnect-EXO @pltRXO } ;

        # use new get-GCFastXO cross-org dc finder
        #$domaincontroller = get-GCFastXO -TenOrg $TenOrg -ADObject $tMbxId -verbose:$($verbose) ;
        <# post tenantmigr, no point
        #-=-=use new get-GCFastXO cross-org dc finder against a TenOrg and -ADObject-=-=-=-=-=-=
        $domaincontroller = $null ; # pre-clear, ensure no xo carryover
        if($tMbxId){
            # the get-addomaincontroller is returning an array; use the first item (second is blank)
            $domaincontroller = get-GCFastXO -TenOrg $TenOrg -ADObject $tMbxId -verbose:$($verbose) |?{$_.length} ;
        } else {throw "unpopulated `$TargetMailboxes parameter, unable to resolve a matching OR OP_ExADRoot property" ; } ;
        #-=-=-=-=-=-=-=-=
        #>
        
        # issue is that 2 objects are coming back: first is null, 2nd is the dc spec
        $Exit = 0 ;
        Do {
            Try {
                if(!(get-AdServerSettings).ViewEntireForest){ enable-ForestView } ;
                $ombx=get-mailbox -id $tMbxId -domaincontroller $domaincontroller ;
                # dc issues, drop it, and use dawdles
                #$ombx=get-mailbox -id $tMbxId  ;

                $Exit = $Retries ;
            } Catch {
                $smsg = "Failed to exec cmd because: $($Error[0])" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                Start-Sleep -Seconds $RetrySleep ;
                    if($pltRX10){
                        ReConnect-Ex2010 @pltRX10 ;
                    } else { Reconnect-Ex2010 ; } ;
                $Exit ++ ;
                $smsg = "Try #: $Exit" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" }
                If ($Exit -eq $Retries) {
                    $smsg = "Unable to exec cmd!"
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                } ;
            }  ;
        } Until ($Exit -eq $Retries) ;

        $rgxTenDomAddr = 'smtp:.*@\w*.mail.onmicrosoft.com' ;
        $rgxTenDomEAP = '^smtp:.*(\@\w*\.mail\.onmicrosoft.com)$'
        #$TenantDomainMail=$TORMeta['o365_TenantDomainMail'] ;
        $TenantDomainMail=(Get-Variable  -name "$($TenOrg)Meta").value.o365_TenantDomainMail ;

        # dynamically pull the onmicrosoft address from lowest priority policy
        $error.clear() ;
        TRY {
            ((get-emailaddresspolicy | sort priority -desc)[0] | select -expand EnabledEmailAddressTemplates|?{$_ -match $rgxTenDomEAP }) | out-null ;
            $OnMicrosoftAddrDomain=$matches[1] ;
        } CATCH {
            $smsg = "Failed processing $($_.Exception.ItemName). `nError Message: $($_.Exception.Message)`nError Details: $($_)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            Exit #Opts: STOP(debug)|EXIT(close)|CONTINUE(move on in loop cycle)|BREAK(exit loop iteration)|THROW $_/'CustomMsg'(end script with Err output)
            # -Variable -Name PassStatus_$($tenorg) -scope Script -Value $null ;
        } ;

        if($ombx){

            if($tenantCoAddr=$ombx | select -expand emailaddresses | ?{$_ -match $rgxTenDomAddr}){
                $smsg= "mbx HAS matching address:$($tenantCoAddr)" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } else {
                # moved the $OnMicrosoftAddrDomain lookup outside of the if/then
                if($OnMicrosoftAddrDomain){
                    $smsg= "mbx MISSING @TENANTDOM.mail.onmicrosoft.com:$($addr.PrimarySMTPAddress)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN} #Error|Warn|Debug
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    # precheck for addrs
                    $smsg= "SMTP Addrs:`n$(($ombx | Select -Expand EmailAddresses | ? {$_ -like "smtp:*"}|out-string).trim())" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    $ombx | Select -Expand EmailAddresses | ? {$_ -like "smtp:*"} ;
                    $smsg= "EAP Settings:`n$(($ombx|fl EmailAddressPolicyEnabled,CustomAttribute5|out-string).trim())" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    $newOnMSAddr="$($ombx.alias)$($OnMicrosoftAddrDomain)" ;
                    $spltSetmailbox=@{
                        identity=$ombx.samaccountname ;
                        #domaincontroller= $domaincontroller ;
                        EmailAddresses= @{add="$($newOnMSAddr)"} ;
                        whatif=$($whatif) ;
                    } ;

                    $smsg= "Set-Mailbox w`n$(($spltSetmailbox|out-string).trim())" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                    $smsg= "emailaddresses (expanded):`n$(($spltsetmailbox.emailaddresses |out-string).trim())" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                    $Exit = 0 ;
                    Do {
                        Try {
                            Set-Mailbox @spltSetmailbox ;
                            $Exit = $Retries ;
                        } Catch {
                            $smsg = "Failed to exec cmd because: $($Error[0])" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            Start-Sleep -Seconds $RetrySleep ;
                            if($pltRX10){
                                ReConnect-Ex2010 @pltRX10 ;
                            } else { Reconnect-Ex2010 ; } ;
                            $Exit ++ ;
                            $smsg = "Try #: $Exit" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" }
                            If ($Exit -eq $Retries) {
                                $smsg = "Unable to exec cmd!"
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            } ;
                        }  ;
                    } Until ($Exit -eq $Retries) ;

                    $1F=$false ;Do {
                        if($1F){Sleep -s 5} ;  write-host "." -NoNewLine ; $1F=$true ;
                        Try {
                            $ombx=get-Mailbox $ombx.samaccountname
                        } CATCH {
                            Write-Warning "$(get-date -format 'HH:mm:ss'): Failed processing $($_.Exception.ItemName). `nError Message: $($_.Exception.Message)`nError Details: $($_)" ;
                            $smsg = "FULL ERROR TRAPPED (EXPLICIT CATCH BLOCK WOULD LOOK LIKE): } catch[$($Error[0].Exception.GetType().FullName)]{" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            Break #Opts: STOP(debug)|EXIT(close)|CONTINUE(move on in loop cycle)|BREAK(exit loop iteration)|THROW $_/'CustomMsg'(end script with Err output)
                        } ;
                    } Until ($ombx.emailaddresses -like '*.onmicrosoft.com') ;

                    $smsg= "`nUpdated Addrs:`n$(( $ombx.EmailAddresses |?{$_ -match 'smtp:.*'}|out-string).trim())" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                    write-host -foregroundcolor yellow "$((get-date).ToString('HH:mm:ss')):MONITORING COMMAND FOR ABOVE:" ;
                    "Do {write-host '.' -NoNewLine;Start-Sleep -m (1000 * 60)} Until (Get-exoRecipient $($ombx.userprincipalname) -ea 0| select -expand emailaddresses|?{$_ -like '*@$($TenantDomainMail)'}) ; write-host '``a' "

                } else{
                    $smsg= ":mbx MISSING @$($TenantDomainMail):$($addr.PrimarySMTPAddress)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug
                    else{ write-host -foregroundcolor RED "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                } ;

                if($tenantCoAddr=$ombx | select -expand emailaddresses | ?{$_ -match [regex]'smtp:.*@$($TenantDomainMail)'}){
                    $smsg= "mbx HAS matching address:$($tenantCoAddr)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                } else {
                    $smsg= "mbx MISSING @$($TenantDomainMail):$($addr.PrimarySMTPAddress)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug
                    else{ write-host -foregroundcolor RED "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                } ;

            } ;

            $Exit = 0 ;
            Do {
                Try {
                    #$exorcp=Get-exoRecipient $($ombx.userprincipalname) -ea 0| select -expand emailaddresses|?{$_ -match $rgxTenDomAddr} ;
                    # go back to using static dc
                    $exorcp=Get-exoRecipient $($ombx.userprincipalname) ;
                    $Exit = $Retries ;
                } Catch {
                    $smsg = "Failed to exec cmd because: $($Error[0])" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    Start-Sleep -Seconds $RetrySleep ;
                    if($pltRX10){
                        ReConnect-Ex2010 @pltRX10 ;
                    } else { Reconnect-Ex2010 ; } ;
                    $Exit ++ ;
                    $smsg = "Try #: $Exit" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" }
                    If ($Exit -eq $Retries) {
                        $smsg = "Unable to exec cmd!"
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ;
                }  ;
            } Until ($Exit -eq $Retries) ;

            if($exorcp){
                # use the cached obj
                #$RemoteHostName = ($MigEnd|?{$_.remoteserver -like 'my*'})[0].remoteserver ;
                $MvSplat=[ordered]@{
                    Identity=$ombx.userprincipalname ;
                    remote=$true;
                    #RemoteHostName=$RemoteHostName ;
                    RemoteHostName=$null ;
                    TargetDeliveryDomain= $OnMicrosoftAddrDomain.replace('@','') # @$($TenantDomainMail)
                    RemoteCredential=$pltRX10.Credential ;
                    BadItemLimit=1000 ;
                    AcceptLargeDataLoss=$true ;
                    Suspend=$false ;
                    SuspendWhenReadyToComplete=$false ;
                    BatchName=$Batchname ;
                    whatif=$($whatif) ;
                } ;

                #
                $tMEPID=$null ;
                if((Get-Variable  -name "$($TenOrg)Meta").value.MEPArray){
                        #set-variable -Name $meta.name -Value ((get-variable -name $meta.name).value  += @{'o365_AcceptedDomains' = (Get-exoAcceptedDomain).domainname} )
                        #$smtpTo = (Get-Variable  -name "$($TenOrg)Meta").value.NotificationAddr1 ;
                        # loop the MepArray members, till you find the one the mbx.db matches on, then use it's MEPID
                        foreach($MA in (Get-Variable  -name "$($TenOrg)Meta").value.meparray){
                            #$ombx.Database
                            if($ombx.Database -match $MA.split(';')[2]){
                                $tMEPID = $MA.split(';')[1] ;
                                break ;
                            } ;
                        } ;
                } else {
                    $smsg = "UNABLE TO RESOLVE `$($TenOrg)Meta.MEPArray" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug
                    else{ write-host -foregroundcolor RED "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    throw $smsg ;
                } ;
                if($tMEPID){$MvSplat.RemoteHostName= $tMepID}
                else {
                    $smsg = "`$tMEPID undefined!" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug
                    else{ write-host -foregroundcolor RED "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    throw $smsg ;
                } ;

                if(!$NoTEST){
                    $tMEPID |foreach-object {
                        $error.clear() ;
                        TRY {
                            $smsg= "Testing OnPrem Admin account $($pltRX10.Credential.username) against $($_)" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            if ($script:useEXOv2) { reconnect-eXO2 @pltRXO }
                            else { reconnect-EXO @pltRXO } ;
                            $oTest= ps1TestXMigrSrvrAvail -ExchangeRemoteMove -RemoteServer $_ -Credentials $pltRX10.Credential ;
                        } CATCH {
                            $smsg= "Failed processing $($_.Exception.ItemName). `nError Message: $($_.Exception.Message)`nError Details: $($_)" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug
                            else{ write-host -foregroundcolor RED "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            $bMEPFail=$true ;
                            #Exit #Opts: STOP(debug)|EXIT(close)|Continue(move on in loop cycle)
                            Continue
                        } # try/catch-E ;
                        $smsg= $oTest ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ;
                } else {
                    write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):-NoTEST param detected: Skipping MEP tests" ;
                } ;

                <# version that keys suspend off of explicit script name
                switch -regex ($ScriptBaseName){
                    "^move-EXOmailboxSuspend\.ps1$" {
                        $smsg= "Configuring:SuspendWhenReadyToComplete=`$TRUE" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } ; #Error|Warn
                        $MvSplat.SuspendWhenReadyToComplete=$true ;
                    }
                    "^move-MailboxToXo\.ps1$" {
                        $smsg= "Configuring:SuspendWhenReadyToComplete=`$FALSE" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } ; #Error|Warn
                        $MvSplatSuspendWhenReadyToComplete=$false ;
                    }
                    default {throw "Unrecognized FILENAME:$($ScriptNameNoExt)"}
                } ;
                #>
                # switch to explicit -Suspend param (for move into verb-exo module): $Suspend
                if($Suspend){
                    $smsg= "-Suspend specified: Configuring:SuspendWhenReadyToComplete=`$TRUE" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                    else{ write-host -foregroundcolor YELLOW "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    $MvSplat.SuspendWhenReadyToComplete=$true ;
                } else {
                    $smsg= "Configuring:SuspendWhenReadyToComplete=`$FALSE" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    $MvSplat.SuspendWhenReadyToComplete=$false ;
                } ;

                $Exit = 0 ;
                Do {
                    Try {
                        $existMove=ps1GetXMovReq -Identity $mvsplat.identity -ea 0 ;
                        $Exit = $Retries ;
                    } Catch {
                        $smsg = "Failed to exec cmd because: $($Error[0])" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        Start-Sleep -Seconds $RetrySleep ;
                        if($pltRX10){
                            ReConnect-Ex2010 @pltRX10 ;
                        } else { Reconnect-Ex2010 ; } ;
                        $Exit ++ ;
                        $smsg = "Try #: $Exit" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" }
                        If ($Exit -eq $Retries) {
                            $smsg = "Unable to exec cmd!"
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        } ;
                    }  ;
                } Until ($Exit -eq $Retries) ;

                if($existMove){
                    $smsg= "==Removing ExistMove:$($existMove.alias)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                    $Exit = 0 ;
                    Do {
                        Try {
                            # add force confirm, it prompted to purge some priors
                            ps1RmvXMovReq $existMove.alias -force:$true -confirm:$false -whatif:$($whatif) ;
                            $Exit = $Retries ;
                        } Catch {
                            $smsg = "Failed to exec cmd because: $($Error[0])" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            Start-Sleep -Seconds $RetrySleep ;
                            if ($script:useEXOv2) { reconnect-eXO2 @pltRXO }
                            else { reconnect-EXO @pltRXO } ;
                            $Exit ++ ;
                            $smsg = "Try #: $Exit" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" }
                            If ($Exit -eq $Retries) {
                                $smsg = "Unable to exec cmd!"
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            } ;
                    }  ;

                    } Until ($Exit -eq $Retries) ;


                    "(waiting for move to purge)" ;
                    Do {write-host "." -NoNewLine;Start-Sleep -m (1000 * 10)} Until (!(ps1GetXMovReq -Identity $mvsplat.identity -ea 0))
                } ;

                $smsg= "===$($ombx.UserPrincipalName):$((get-alias ps1NewXMovReq).definition) w`n$(($mvSplat|out-string).trim())" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                $Exit = 0 ;
                Do {
                    Try {
                        $mvResult = ps1NewXMovReq @MvSplat;
                        <# if you don't cap output, it drops into the pipeline:
                        DisplayName           StatusDetail        TotalMailboxSize TotalArchiveSize PercentComplete
                        -----------           ------------        ---------------- ---------------- ---------------
                        Stg-Consumer Warranty WaitingForJobPickup 0 B (0 bytes)                     0
                        Stg-Dlradmin          WaitingForJobPickup 0 B (0 bytes)                     0
                        #>
                        $statusdelta = ";CHANGE";
                        $script:PassStatus += $statusdelta ;
                        set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta) ;
                        $smsg = "Move initiated:$($MvSplat.identity):`n$(($mvResult | ft -auto DisplayName,StatusDetail,TotalMailboxSize,TotalArchiveSize,PercentComplete|out-string).trim())" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                        $Exit = $Retries ;
                    } Catch {
                        $smsg = "Failed to exec cmd because: $($Error[0])" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        Start-Sleep -Seconds $RetrySleep ;
                        if ($script:useEXOv2) { reconnect-eXO2 @pltRXO }
                        else { reconnect-EXO @pltRXO } ;
                        $Exit ++ ;
                        $smsg = "Try #: $Exit" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" }
                        If ($Exit -eq $Retries) {
                            $smsg = "Unable to exec cmd!"
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        } ;
                    }  ;
                } Until ($Exit -eq $Retries) ;

            } else {
                $smsg= "===$($ombx.userprinciplname) missing $((Get-Variable  -name "$($TenOrg)Meta").value.o365_TenantDom) address at EXO END.`nSKIPPING EXO MOVE!" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug
                else{ write-host -foregroundcolor RED "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } ;
        } else {
                $smsg= "===$($tMbxId):NOT FOUND, SKIPPING!:" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug
                else{ write-host -foregroundcolor RED "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        };

        $smsg= "$($sBnr.replace('=v','=^').replace('v=','^='))" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

        # add an EXO delay to avoid issues
        Start-Sleep -Milliseconds $ThrottleMs ;
    } ;  # loop-E Mailboxes

    if(!$whatif){
        $smsg= "CLOUD MIGRATION STATUS:`n$((ps1GetXMovReq -BatchName $BatchName | ps1GetXMovReqStats | FL DisplayName,status,percentcomplete,itemstransferred,BadItemsEncountered|out-string).trim())`n" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        $smsg= "`nContinue monitoring with:`n$((get-alias ps1GetXMovReq).definition) -BatchName $($BatchName) | $((get-alias ps1GetXMovReqStats).definition) | fl DisplayName,status,percentcomplete,itemstransferred,BadItemsEncountered`n" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
    } ;
    

    # return the passstatus to the pipeline
    $script:PassStatus | write-output

    if((get-AdServerSettings).ViewEntireForest){ disable-ForestView } ;

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
    
    # prod is still showing a running unstopped transcript, kill it again
    $stopResults = try {Stop-transcript -ErrorAction stop} catch {} ;
    
    # also trailing echo the log:
    $smsg = "`$logging:`$true:written to:`n$($logfile)" ; 
    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
    $smsg = "$((get-date).ToString('HH:mm:ss')):$($sBnr.replace('=v','=^').replace('v=','^='))" ;
    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
    #*------^ END SUB MAIN ^------
}

#*------^ move-MailboxToXo.ps1 ^------