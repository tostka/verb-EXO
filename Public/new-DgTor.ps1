#*------v new-DgTor.ps1 v------
function new-DgTor {
    <#
    .SYNOPSIS
    new-DgTor.ps1 - Create new DistributionGroup and populate. Notates requestor, ticket# and admin, in Notes field of DL
    .NOTES
    Version     : 1.0.0
    Author      : Todd Kadrie
    Website     :	http://www.toddomation.com
    Twitter     :	@tostka / http://twitter.com/tostka
    CreatedDate : 2021-08-30
    FileName    : new-DgTor.ps1
    License     : MIT License
    Copyright   : (c) 2021 Todd Kadrie
    Github      : https://github.com/tostka/verb-exo
    Tags        : Powershell,ExchangeOnline,Exchange,DistributionGroup,DistributionList,Hybrid
    REVISIONS
    * 2:40 PM 12/10/2021 more cleanup 
    * 4:54 PM 9/30/2021 updated CloudFirst code, and used to create functional 
    exoDG, also added code to dynamically create onprem unreplicated MailContacts 
    for CloudFirst, to represent an onprem AddrBook object ; flipped members & 
    ManagedBy to [string[]]; added UG->DG conversion example,
    with splat; swapped undependable $scriptbasename w $cmdletname (tends to show
    module as name); updated cloudfirst variant code, switched recipient &
    managedby's to raw smtpaddresses ; finally validated hybrid OnPrem DG creation as well. 
    * 4:28 PM 9/14/2021 ren: new-DL-Tor -> new-DgTor ;; debugged -CloudFirst - working as intended ; fixed pipeline bug coming out of Start-IseTranscript , properly returns newdg object to pipeline;  converted to function, add to verb-exo (once it includes proper hybrid, most approp place); -CloudFirst still undebugged
    *2:56 PM 9/13/2021 rewrote for modern modules & template, -outputObject default. Tested ExOP version, haven't done -CloudFirst debugging yet.
    # 11:05 AM 6/13/2019 updated get-admininitials()
    # 8:24 AM 2/13/2018 support empty groups (for infra EXO rule grps, to be stocked later) - looks like it was there, members are only added where specified (foreach). Adding explicit echo's to reflect the status. 
    # 2:54 PM 2/8/2018 updated to support EXO-hosted mailuser owners (members already worked without modification)
    #12:50 PM 11/27/2017 sec wants them defaulted RequireSenderAuthenticationEnabled to $true ; 
    # 1:41 PM 6/13/2017 spliced in latest 3/16/16 get-gcfast()
    # 10:35 AM 4/4/2017 added new -InetReceive:$true param, to configure RequireSenderAuthenticationEnabled=$false
    # 10:32 AM 4/4/2017 had to splice in latest loadmod set - failed if not already in EMS session
    # 12:23 PM 4/3/2017 add default RequireSenderAuthenticationEnabled  $false (allow inet mailing by default)
    # 9:50 AM 3/3/2017 Get-AdminInitials: with the new standard of Fname/name: S-[name] this needs an update to strip prefix S-
    # 9:48 AM 3/2/2017 merged in updated Add-EMSRemote Set
    # 12:44 PM 10/18/2016 update rgx for ticket to accommodate 5-digit (or 6) CW numbers "^\d{6}$"=>^\d{5,6}$
    #* 9:11 AM 9/30/2016 added pretest if(get-command -name set-AdServerSettings -ea 0)
    # 1:55 PM 6/6/2016 debugged, works. 
    # 1:52 PM 6/6/2016 cleanedup typo trailing spaces on some of the dummy hash buils
    # 12:12 PM 6/6/2016 added Execute-WithRetry(), implementing retries via function. - neither worked, just do 
    # 10:50 AM 6/6/2016 : 
        add retry support below params
        * add region tags
        * updated to enable-mbx LoadMod block
        * Move splats below constants - constants should always be in place 1st
    # 1:12 PM 2/11/2016 fixed new bug in get-GCFast, wasn't detecting blank $site
    # 12:20 PM 2/11/2016 updated to standard EMS/AD Call block & Add-EMSRemote()
    #10:49 AM 2/11/2016: updated get-GCFast to current spec, updated any calls for "-site 'lyndale'" to just default to local machine lookup
    # 7:40 AM Add-EMSRemote: 2/5/2016 another damn cls REM IT! I want to see all the connectivity info, switched wh->wv, added explicit echo's of what it's doing.
    # 11:08 AM 1/15/2016 re-updated Add-EMSRemote, using a -eq v -like with a wildcard string. Have to repush copies all over now.
    # 10:43 AM 1/13/2016 updated Add-EMSRemote set
    # 10:02 AM 1/13/2016: fixed cls bug due to spurious ";cls" included in the try/catch boilerplate: Write-Error "$((get-date).ToString('HH:mm:ss')): Command: $($_.InvocationInfo.MyCommand)" ;cls => Write-Error "$((get-date).ToString('HH:mm:ss')): Command: $($_.InvocationInfo.MyCommand)" ;
    # 9:58 AM 10/21/2015 ren $InputSplat.Site => $InputSplat.SiteCode, to sync up name standard across scripts
    # 9:08 AM 10/14/2015 added debugpref maint code to get write-debug to work
    # 11:55 AM 10/7/2015 sorted now; wasn't using -SiteOverRide in the fancy csv-fed param I was using
    #1:56 PM 10/6/2015 fix $SiteOverride, to actually override the ManagedBy's OU (BEA user was forcing to BEA vs override of LYN)
    # 1:09 PM 10/6/2015 updated code to spec - seems to work
    # 10:57 AM 10/6/2015 blanked the @InputSplat{} values
    # 10:43 AM 10/6/2015 ren paras to standard (tix->Ticket)
    # 2:34 PM 10/2/2015 fix break and port in EMSremote etc
    8:53 AM 9/4/2015 sub'd in cleanedup EMSRemote Set
    8:35 AM 9/4/2015 seems to work, from EMSRemote
    2:36 PM 9/3/2015 added -identity $dg.samaccountname to Get-DistributionGroupMember
    2:17 PM 9/3/2015 added -ea 0 to Get-DistributionGroupMember cmds (were erroring if no members)
    1:36 PM 9/3/2015 did alot of tshooing on add-emsr, to get it functional all seems to run right now
    1:12 PM 8/26/2015 removed partial GUI code and validated functions properly.
    1:15 PM 6/24/2015 added whatif test to New-DistributionGroup @DLSplat -ea Stop ;
    1:04 PM 6/24/2015 tests out functional on params as well. Looks pretty good
    12:46 PM 6/24/2015 functional, need to test params
    9:21 AM 6/24/2015 combo/hybrid old ]PSNewDL! with add-MbxAccessGrant.ps1
    1:55 PM 6/15/2015 initial version
    .DESCRIPTION
    new-DgTor.ps1 - Create new DistributionGroup and populate. Notates requestor, ticket# and admin, in Notes field of DL
    Derives DG settings from ManagedBy specified: 
    - Constructs Name using SiteCode (root child OU at 1st level of MangedBy DN): [SITECODE]-DL-[$DNameBase]
    - Find's standard DG OU within the same Site OU tree
    .PARAMETER TenOrg
    TenantTag value, indicating Tenants to connect to[-TenOrg 'TOL']
    .PARAMETER useEXOv2
    Use EXOv2 (ExchangeOnlineManagement) over basic auth legacy connection [-useEXOv2]
    .PARAMETER Ticket
    ITSM Request/Incident Number [nnnnnn]
    .PARAMETER DNameBase
    Base Name string, for DL Name construction. [SIT-DL] will be automatically appended[Base Name String]
    .PARAMETER ManagedBy
    Specify the userid to be responsible for access-grant-approvals[name,emailaddr,alias]
    .PARAMETER HiddenFromAddressLists
    Switch to configure -HiddenFromAddressListsEnabled `$true [-HiddenFromAddressLists]
    .PARAMETER CloudFirst
    Switch to specify EXO Cloud-First DG (vs Federated replicated AD/EXOnPrem DG) [-CloudFirst]
    .PARAMETER SiteOverride
    Specify a 3-letter Site Code. Used to force DL name/placement to vary from ManagedBy's current site[3-letter Site code]
    .PARAMETER Members
    Comma-delimited string of potential users to be granted access[name,emailaddr,alias]
    .PARAMETER InetReceive
    Can receive from external senders [-InetReceive:`$true]
    .PARAMETER HiddenFromAddressLists
    Switch to configure -HiddenFromAddressListsEnabled `$true [-HiddenFromAddressLists]
    .PARAMETER OutObject
    Switch to specify to return the new DG as an object (defaults true)[-OutObject]
    .PARAMETER Whatif
    Parameter to run a Test no-change pass, and log results [-Whatif switch]
    .PARAMETER ShowDebug
    Parameter to display Debugging messages (diverts reports to alt address) [-ShowDebug switch]
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    SystemObject
    [| get-member the output to see what .NET obj TypeName is returned, to use here]
    .EXAMPLE
    PS> new-DgTor -SiteOverride ENT -DNameBase "DL Base DisplayName" -ManagedBy MANAGERID -Members 'MEMBER1@DOMAIN.com','MEMBER2@DOMAIN.com' -HiddenFromAddressLists -showDebug -verbose -Ticket 99999 -whatif;
    Create a DL with a siteoverride (spec'ing as ENT-name, rather than alt site)
    .EXAMPLE
    PS> $ndg = new-DgTor -TenOrg TOR -DNameBase "DL BASE DISPLAYNAME" -ManagedBy MANAGERID -Members 'MEMBER1@DOMAIN.com','MEMBER2@DOMAIN.com' -HiddenFromAddressLists -showDebug -verbose -Ticket 99999 -outobject;
    Create a DG with SiteOverride and return resulting new DG as an object, assigned to $ndg
    .EXAMPLE
    $whatif=$true ; 
    reconnect-exo ; 
    TRY{
        $tugn = 'TeamsUGNamwe_GUID' ;
        $tug = Get-exoUnifiedGroup -Identity $tugn ;
        $tugmbrs = Get-exoUnifiedGroupLinks -Id $tug.name -LinkType Members ;
        $pltNDg=[ordered]@{   TenOrg='TOR' ;
            CloudFirst=$true ;
            DNameBase="IS-$($tug.displayname)" ;
            ManagedBy=($tug.managedby | get-exorecipient -ea STOP | select -expand primarysmtpaddress) ;
            Members=($tugmbrs.primarysmtpaddress| get-exorecipient -ea STOP | select -expand primarysmtpaddress) ;
            HiddenFromAddressLists=$false;
            showDebug=$true;
            verbose=$true;
            Ticket='99999' ;
            outobject=$true;
            whatif=$($whatif);
          } ;
         write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):new-DgTor w`n$(($pltNDg|out-string).trim())" ;
         $ndg = new-DgTor @pltNDg;
         $ndg ;
    } CATCH {
    Write-Warning "$(get-date -format 'HH:mm:ss'): Failed processing $($_.Exception.ItemName). `nError Message: $($_.Exception.Message)`nError Details: $($_)" ;
    break ;
    } ;
    # Traditionally the UG would then be set Hidden from Address Books, in deference to the DG
    Set-exoUnifiedGroup -Id $tug.name -HiddenFromAddressListsEnabled $true
    
    Demo conversion of a Teams Unified Group membership (which permits unsubscribes, and silent loss of mail) into a standard DG. 
    .EXAMPLE
    $whatif=$true ;
    reconnect-exo ;
    reconnect-ex2010 ;
    TRY{
        $tugn = 'TeamsUGNamwe_GUID' ;
        $tug = Get-exoUnifiedGroup -Identity $tugn ;
        $tugmbrs = Get-exoUnifiedGroupLinks -Id $tug.name -LinkType Members ;
        $pltNDg=[ordered]@{   TenOrg='TOR' ;
          CloudFirst=$false ;
          SiteOverride = 'ENT' ;
          DNameBase="IS-$($tug.displayname)" ;
          ManagedBy=($tug.managedby | get-exorecipient -ea continue | select -expand primarysmtpaddress | select -unique | get-recipient -ea continue | select -expand primarysmtpaddress | select -unique) ;
          Members=($tugmbrs.primarysmtpaddress| get-recipient -ea continue | select -expand primarysmtpaddress) ;
          HiddenFromAddressLists=$false;
          showDebug=$true;
          verbose=$true;
          Ticket='99999' ;
          outobject=$true;
          whatif=$($whatif);
         } ;
        write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):new-DgTor w`n$(($pltNDg|out-string).trim())" ;
        $ndg = new-DgTor @pltNDg;
        $ndg ;
    } CATCH {
        Write-Warning "$(get-date -format 'HH:mm:ss'): Failed processing $($_.Exception.ItemName). `nError Message: $($_.Exception.Message)`nError Details: $($_)" ;
        break ;
    } ;
    # Traditionally the UG would then be set Hidden from Address Books, in deference to the DG
    Set-exoUnifiedGroup -Id $tug.name -HiddenFromAddressListsEnabled $true
    
    Demo conversion of a Teams Unified Group membership (which permits unsubscribes, and silent loss of mail) into a standard on-prem hybrid replicated DG, 
    with the prefix/SiteCode overridden to use 'ENT' over the ManagedBy's home SiteCode. 
    
    Note:An onprem DG can have MailContacts in the membership (and in the ManagedBy as well), with the same primarysmtpaddress as remote-hybrid cloud mailboxes. 
    (e.g. EXO mailboxes that aren't Hybrid/AD'd locally to the DG-hosting OnPrem mail org)
    Upon ADC replication of a MailContact member (or ManagedBy) to AzureAD, any MailContact set in either property, will be auto-replaced during replication, 
    with the matching EXO mailbox object. A neat way of maintaining hybrid onPrem-deliverable DG's, containing non-locally-replicated EXO mailboxes.
    .LINK
    https://github.com/tostka/verb-exo        
    #>
    ###Requires -Modules ActiveDirectory, AzureAD, MSOnline, ExchangeOnlineManagement, verb-ADMS, verb-Auth, verb-Ex2010, verb-EXO, verb-IO, verb-logging, verb-Text, verb-logging
    #Requires -Modules ActiveDirectory, AzureAD, MSOnline, ExchangeOnlineManagement, verb-ADMS, verb-Auth, verb-Ex2010, verb-IO, verb-logging, verb-Text, verb-logging
    # VALIDATORS: [ValidateNotNull()][ValidateNotNullOrEmpty()][ValidateLength(24,25)][ValidateLength(5)][ValidatePattern("(lyn|bcc|spb|adl)ms6(4|5)(0|1).(china|global)\.ad\.DOMAIN\.com")][ValidateSet("USEA","GBMK","AUSYD")][ValidateScript({Test-Path $_ -PathType 'Container'})][ValidateScript({Test-Path $_})][ValidateRange(21,65)][ValidateCount(1,3)]
    ## [OutputType('bool')] # optional specified output type
    [CmdletBinding()]
    PARAM(
        [Parameter(HelpMessage="TenantTag value, indicating Tenants to connect to[-TenOrg 'TOL']")]
        [ValidateNotNullOrEmpty()]
        $TenOrg = 'TOR',
        [Parameter(HelpMessage="Use EXOv2 (ExchangeOnlineManagement) over basic auth legacy connection [-useEXOv2]")]
        [switch] $useEXOv2,
        [Parameter(Mandatory=$true,HelpMessage="Base Name string, for DL Name construction. [SIT-DL] will be automatically appended[Base Name String]")]
        [string]$DNameBase,
        [Parameter(Position=0,HelpMessage="ITSM Request/Incident Number [nnnnnn]")]
        #[ValidatePattern("^\d{5,6}$")]
        [string]$Ticket,
        [Parameter(Mandatory=$true,HelpMessage="Specify the userid to be responsible for access-grant-approvals[name,emailaddr,alias]")]
        [string[]]$ManagedBy,
        [Parameter(HelpMessage="Specify a 3-letter Site Code. Used to force DL name/placement to vary from ManageBy's current site[3-letter Site code]")]
        [string]$SiteOverride,
        [Parameter(HelpMessage="Comma-delimited string of potential users to be granted access[name,emailaddr,alias]")]
        [string[]]$Members,
        [Parameter(HelpMessage="Can receive from external senders [-InetReceive:`$true]")]
        [switch]$InetReceive,
        [Parameter(HelpMessage="Switch to configure -HiddenFromAddressListsEnabled `$true [-HiddenFromAddressLists]")]
        [switch]$HiddenFromAddressLists,
        [Parameter(HelpMessage="Switch to specify EXO Cloud-First DG (vs Federated replicated AD/EXOnPrem DG) [-CloudFirst]")]
        [switch]$CloudFirst,
        [Parameter(HelpMessage="Switch to specify to return the new DG as an object (defaults true)[-OutObject]")]
        [switch]$OutObject=$true,
        [Parameter(HelpMessage='Parameter to display Debugging messages (also diverts reports to alt address) [-ShowDebug switch]')]
        [switch] $showDebug=$false,
        [Parameter(HelpMessage="Whatif Flag  [-whatIf]")]
        [switch]$whatIf=$false
    ) ;
    
    ${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name ;         
    $verbose = ($VerbosePreference -eq "Continue") ;

    $progInterval= 500 ; # write-progress wait interval in ms
    # 1:05 PM 4/28/2017 retries
    $DoRetries = 4 ;
    $RetrySleep = 5 ;

    # 10:22 AM 2/24/2016 add an explicit vari $MaxProcessingUserLimit for max-processing users, and bump it > 4000 (causing processing aborts as of 2/17/16)
    $MaxProcessingUserLimit = 10000 ;
    # 2:20 PM 3/25/2015 added optional report inline-in email body
    $bodyAsHtml=$true ;

    # 12:15 PM 2/9/2015 add an SMTP retry limit (per user attempted)
    # 7:20 AM 5/7/2015 leveraging the varis for LineURI non-unique recoveries
    [int]$retryLimit=1; # just one retry to patch lineuri duped users and retry 1x
    [int]$retryDelay=20;    # secs wait time after failure

    # 1:57 PM 2/18/2015
    $abortPassLimit = 4;    # maximum failed users to abort entire pass
    # 9:49 AM 2/17/2015 SMTP Priority level[Normal|High|Low]
    $smtpPriority="Normal";
    # SMTP port (default is 25)
    $smtpPort = 25 ;

   
    
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
    
    
    #*================v FUNCTIONS  v================

    #-------v Function _cleanup v-------
    function _cleanup {
        <#
        .SYNOPSIS
        _cleanup.ps1 - clear all objects, prep close transcript, email report and exit
        .NOTES
        Version     : 1.0.0
        Author      : Todd Kadrie
        Website     :	http://www.toddomation.com
        Twitter     :	@tostka / http://twitter.com/tostka
        CreatedDate : 2020-
        FileName    :
        License     : MIT License
        Copyright   : (c) 2020 Todd Kadrie
        Github      : https://github.com/tostka/verb-XXX
        Tags        : Powershell
        AddedCredit : REFERENCE
        AddedWebsite:	URL
        AddedTwitter:	URL
        REVISIONS
        # 8:47 AM 11/24/2020 cloned over intact from maintain-exousrmbxretentionpolicies
        # 3:15 PM 10/13/2020 added CBH, added params: summarizeStatus,
            NoTranscriptStop, TranscriptItemsLimit, each exempts certain blocks of process
            - trying to genericize for reuse on other scripts ; added html body support
            (using <pre../pre> to preserve text layout, even in outlook display
        # 12:40 PM 10/23/2018 added write-log trainling bnr
        # 2:02 PM 9/21/2018 missing $timestampnow, hardcode
        # 8:45 AM 10/13/2015 reset $DebugPreference to default SilentlyContinue, if on
        # # 8:46 AM 3/11/2015 at some time from then to 1:06 PM 3/26/2015 added ISE Transcript
        # 8:39 AM 12/10/2014 shifted to stop-transcriptLog function
        # 7:43 AM 1/24/2014 always stop the running transcript before exiting
        .DESCRIPTION
        _cleanup.ps1 - clear all objects, prep close transcript, email report and exit
        .PARAMETER  LogPath
        Alt transcript/logfile path for mailing (rather than `$transcript/`$logfile)[-LogPath c:\path-to\log.txt]
        .PARAMETER TranscriptItemsLimit
        Number of transactions to determine Transcript inclusion[-TranscriptItemsLimit]
        .PARAMETER summarizeStatus
        Switch to output a summary of the `$script:PassStatus delimted string[-summarizeStatus]
        .PARAMETER NoTranscriptStop
        Switch to skip transcript stop & exit [-NoTranscriptStop]
        .PARAMETER ShowDebug
        Parameter to display Debugging messages [-ShowDebug switch]
        .PARAMETER Whatif
        Parameter to run a Test no-change pass [-Whatif switch]
        .EXAMPLE
        _cleanup
        Default Call
        .EXAMPLE
        $plt_cleanup=@{LogPath=$tmpcopy summarizeStatus=$true ;  NoTranscriptStop=$true ; showDebug=$($showDebug) ;  whatif=$($whatif) ; } ;
        _cleanup @pltCleanup ;
        Splatted parameter'd call
        #>
        [CmdletBinding()]
        PARAM(
            [Parameter(HelpMessage="Alt transcript/logfile path for mailing (rather than `$transcript/`$logfile)[-LogPath c:\path-to\log.txt]")]
            [ValidateScript({Test-Path $_})]
            $LogPath,
            [Parameter(HelpMessage="Number of transactions to determine Transcript inclusion[-TranscriptItemsLimit]")]
            [int] $TranscriptItemsLimit = 10,
            [Parameter(HelpMessage="Switch to output a summary of the `$script:PassStatus delimted string[-summarizeStatus]")]
            [switch] $summarizeStatus,
            [Parameter(HelpMessage="Switch to skip transcript stop & exit [-NoTranscriptStop]")]
            [switch] $NoTranscriptStop,
            [Parameter(HelpMessage="Debugging Flag [-showDebug]")]
            [switch] $showDebug,
            [Parameter(HelpMessage="Whatif Flag  [-whatIf]")]
            [switch] $whatIf
        ) ;

        # clear all objects, prep close transcript, email report and exit
        # REVISIONS

        $smsg = "_cleanup" ; 
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        if(!$NoTranscriptStop){
            # handle transcript closure in the main script (Tenant loop)
            stop-transcript
            if(($host.Name -eq "Windows PowerShell ISE Host") -AND ($host.version.Major -lt 5)){
                $Logname=(join-path -path (join-path -path $scriptDir -childpath "logs") -childpath ($scriptNameNoExt + "-" + (get-date -format 'yyyyMMdd-HHmmtt') + "-ISEtrans.log")) ;
                $smsg = "`$Logname: $($Logname)";
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                $xRet = Start-iseTranscript -logname $Logname -Verbose:($VerbosePreference -eq 'Continue') ;
                #Archive-Log $Logname -Verbose:($VerbosePreference -eq 'Continue');
                # 1:23 PM 4/23/2015 standardize processing file so that we can send a link to open the transcript for review
                $transcript = $Logname
            } else {
                $smsg = "Stop Transcript" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                $xRet = Stop-TranscriptLog -Verbose:($VerbosePreference -eq 'Continue') ;
                #if($showdebug){ $smsg = "Archive Transcript" };
                #Archive-Log $transcript -Verbose:($VerbosePreference -eq 'Continue') ;
            } # if-E
        } else {
            $smsg = "(_cleanup(): deferring transcript stop to main script)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
        } ; # !$NoTranscriptStop

        # variant options:
        #$smtpSubj= "Proc Rpt:$($ScriptBaseName):$(get-date -format 'yyyyMMdd-HHmmtt')"   ;
        #Load as an attachment into the body text:
        #$body = (Get-Content "path-to-file\file.html" ) | converto-html ;
        #$SmtpBody += ("Pass Completed "+ [System.DateTime]::Now + "`nResults Attached: " +$transcript) ;
        # 4:07 PM 10/11/2018 giant transcript, no send
        #$SmtpBody += "Pass Completed $([System.DateTime]::Now)`nResults Attached:($transcript)" ;
        #$SmtpBody += "Pass Completed $([System.DateTime]::Now)`nTranscript:($transcript)" ;

        # group out the PassStatus_$($tenorg) strings into a report for eml body
        if($script:PassStatus){
            if($summarizeStatus){
                if($script:TargetTenants){
                    # loop the TargetTenants/TenOrgs and summarize each processed
                    foreach($TenOrg in $TargetTenants){
                        $SmtpBody += "`n===Processing Summary: $($TenOrg):" ;
                        # can't split an empty string
                        if((get-Variable -Name PassStatus_$($tenorg)).value){
                            if((get-Variable -Name PassStatus_$($tenorg)).value.split(';') |?{$_ -ne ''}){
                                $SmtpBody += (summarize-PassStatus -PassStatus (get-Variable -Name PassStatus_$($tenorg)).value -verbose:$($VerbosePreference -eq 'Continue') );
                            } ;
                        } else {
                            $SmtpBody += "(no processing of mailboxes in $($TenOrg), this pass)" ;
                        } ;
                        $SmtpBody += "`n" ;

                    } ;
                } ;
            } else {
                # dump PassStatus right into the email
                $SmtpBody += "`n`$script:PassStatus: $($script:PassStatus):" ;
            } ;
            if($SmtpAttachment){
                $smtpBody +="(Logs Attached)"
            };
            $SmtpBody += "`n$('-'*50)" ;
            # include transcript in body, where fewer than limit of processed items logged in PassStatus
            # no, there're 3 transcripts, stored in $Alltranscripts, but skip it#
    #        if( ($script:PassStatus.split(';') |?{$_ -ne ''}|measure).count -lt $TranscriptItemsLimit){
    #            # add full transcript if less than 10 entries processed
    #            $SmtpBody += "`nTranscript:$(gc $transcript)`n" ;
    #        } else {
                if(!$ArchPath ){ $ArchPath = get-ArchivePath } ;
                if($Alltranscripts){
                    $Alltranscripts |%{
                        #$archedTrans = join-path -path $ArchPath -childpath (split-path $transcript -leaf) ;
                        $archedTrans = join-path -path $ArchPath -childpath (split-path $_ -leaf) ;
                        $smtpBody += "`nTranscript accessible at:`n$($archedTrans)`n" ;
                    } ;
                } ;
            #};
        }
        $SmtpBody += "Pass Completed $([System.DateTime]::Now)" ;

        # body rendered in OL loses all wordrwraps
        # force strip out the html
        #$smtpBody = [regex]::Replace($smtpBody, "\<[^\>]*\>", '') ;
        # or do min html format

        $styleCSS = "<style>BODY{font-family: Arial; font-size: 10pt;}" ;
        $styleCSS += "TABLE{border: 1px solid black; border-collapse: collapse;}" ;
        $styleCSS += "TH{border: 1px solid black; background: #dddddd; padding: 5px; }" ;
        $styleCSS += "TD{border: 1px solid black; padding: 5px; }" ;
        $styleCSS += "</style>" ;

    <# simple no css
    $html = @"
<html>
<head><title>$title</title></head>
<body>
<pre>$smtpBody</pre>
</body>
</html>
"@ ;
#>
    # one with style support (goees in the <head../head> block)
    $html = @"
<html>
<head>
$($styleCSS)
<title>$title</title></head>
<body>
<pre>
$($smtpBody)
</pre>
</body>
</html>
"@ ;

        # convertto-html doesn't do raw txt, just objects
        #$smtpBody = $smtpBody | ConvertTo-Html -Head $styleCSS ;
        # use the bp html <pre../pre> version
        $smtpBody = $html ;

        # name $attachment for the actual $SmtpAttachment expected by Send-EmailNotif
        #$SmtpAttachment=$transcript ;
        # test for ERROR|CHANGE - actually non-blank, only gets appended to with one or the other
        # to test for one, (but not a regex)
        <# # always force
        #if($script:passstatus.split(';') -contains 'ERROR'){
        # or run on change/error/passstatus flag
        if([string]::IsNullOrEmpty($script:PassStatus)){
            $smsg = "No Email Report: `$script:PassStatus isNullOrEmpty" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        } else {
        #>
            $Email = @{
                smtpFrom = $SMTPFrom ;
                SMTPTo = $SMTPTo ;
                SMTPSubj = $SMTPSubj ;
                #SMTPServer = $SMTPServer ;
                SmtpBody = $SmtpBody ;
                SmtpAttachment = $SmtpAttachment ;
                BodyAsHtml = $false ; # let the htmltag rgx in Send-EmailNotif flip on as needed
                verbose = $($VerbosePreference -eq "Continue") ;
            } ;
            <# DISABLED add trailing notifc
            $smsg = "Send-EmailNotif w`n$(($Email|out-string).trim())" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            # 9:02 AM 9/2/2021 don't neeed email on these
            Send-EmailNotif @Email ;
            #>
        #} ;

        if(!$NoTranscriptStop){
            #EXIT # trailing tempfile _cleanup in the sub main
            #Break ; 
        } ;

    } #*------^ END Function _cleanup ^------

    #*================^ END FUNCTIONS  ^================

    #*======v SUB MAIN  v======
    
    $sBnr="`n#*======v $(${CmdletName}) : v======" ; 
    $smsg = $sBnr ;
    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

    $Verbose = ($VerbosePreference -eq 'Continue') ; 

    # *** REGION MARKER LOAD
    #region LOAD
    # *** LOADING


    # 1:00 PM 4/28/2017 email trigger vari
    $PassStatus = $null ;

    # 1:01 PM 4/28/2017 add try catch as well - this may be making it zero-tolerance and catching all minor errors, disable it
    #Set-StrictMode -Version 2.0 ;

    #*------v SERVICE CONNECTIONS v------
    #-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
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

    $useEXO = $false ; # non-dyn setting, drives variant EXO reconnect & query code
    if($CloudFirst){ $useEXO = $true } ; 
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
            # make it script scope, so we don't have to predetect & purge before using new-variable
            if(get-Variable -Name cred$($tenorg) -scope Script -ea 0 ){ remove-Variable -Name cred$($tenorg) -scope Script } ;
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
            #verbose = $($verbose) ; 
            Verbose = $FALSE ; Silent = $true ;} ; 
        if ($script:useEXOv2) { reconnect-eXO2 @pltRXO }
        else { reconnect-EXO @pltRXO } ;
        # or with Tenant-specific cred($Tenorg) lookup
        #$pltRXO creds & .username can also be used for AzureAD connections 
        Connect-AAD @pltRXO ; 
        ###>
        # configure splat for connections: (see above useage)
        $pltRXO = @{
            Credential = (Get-Variable -name cred$($tenorg) ).value ;
            #verbose = $($verbose) ; 
            Verbose = $FALSE ; Silent = $true ;} ; 
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
            if(get-Variable -Name "cred$($tenorg)OP" -scope Script -ea 0 ){ remove-Variable -Name "cred$($tenorg)OP" -scope Script } ;
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
            #verbose = $($verbose) ; 
            Verbose = $FALSE ; Silent = $true ; } ;  
        Reconnect-Ex2010 @pltRX10 ; # local org conns
        #$pltRx10 creds & .username can also be used for local ADMS connections
        #>
        $pltRX10 = @{
            Credential = (Get-Variable -name "cred$($tenorg)OP" ).value ;
            #verbose = $($verbose) ; 
            Verbose = $FALSE ; Silent = $true ; } ;     
        
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

    load-ADMS -Verbose:$FALSE ;

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
    if($UseOP){
        $domaincontroller = get-GCFastXO -TenOrg $TenOrg -subdomain ((gv -name "$($TenOrg)Meta").value['OP_ExADRoot']) -verbose:$($verbose) |?{$_.length};
    } ; 


    <# MSOL CONNECTION
    $reqMods += "connect-msol".split(";") ;
    if ( !(check-ReqMods $reqMods) ) { write-error "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Missing function. EXITING." ; Break ; }  ;
    write-host -foregroundcolor gray  "(loading AAD...)" ;
    #connect-msol ;
    connect-msol @pltRXO ; 
    #>

    <#
    # AZUREAD CONNECTION
    $reqMods += "Connect-AAD".split(";") ;
    if ( !(check-ReqMods $reqMods) ) { write-error "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Missing function. EXITING." ; Break ; }  ;
    write-host -foregroundcolor gray  "(loading AAD...)" ;
    #connect-msol ;
    Connect-AAD @pltRXO ; 
    #>


    <# defined above
    # EXO connection
    $pltRXO = @{
        Credential = (Get-Variable -name cred$($tenorg) ).value ;
        verbose = $($verbose) ; } ; 
    #>
    <# 
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
    #>
    #-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
    #*------^ END SERVICE CONNECTIONS ^------


    $error.clear() ;
    TRY {
        #$LMSLoaded = load-LMS ; write-host -foregroundcolor green "`$LMSLoaded: $LMSLoaded" ;
        # 12:55 PM 4/25/2017 add ems
        #$sName="Microsoft.Exchange.Management.PowerShell*"; if (!(Get-PSSnapin | where {$_.Name -eq $sName})) {Add-PSSnapin $sName -ea Stop};
        # 2:04 PM 4/26/2017 use a full func
        <#$EMSLoaded = Load-EMSSnap ; Write-Debug "`$EMSLoaded: $EMSLoaded" ;
        get-exchangeserver | out-null ;
        #$ADMTLoaded = load-ADMS ; write-host -foregroundcolor green "`$ADMTLoaded: $ADMTLoaded" ;
        #>
        <# 2nd gen disabled
        rx10 -Verbose:$false ; 
        rxo  -Verbose:$false ; 
        #cmsol -Verbose:$false ; 
        connect-ad -Verbose:$false | out-null ;;
        if(!$domaincontroller){$domaincontroller=get-gcfast} ;
        #>

        # 11:56 AM 4/24/2015 moved below func defs, in sub main
        $archPath = get-ArchivePath ;

        # 12:44 PM 4/24/2015 fine squash any array coming out (till we get it sorted)
        if($archPath -is [system.array]){
            if($bDebug) {Write-Verbose "Flattening `$archpath array" -verbose:$verbose}
            $archPath = $archPath[0] ;
        }  # if-E;

    } CATCH {
        $ErrTrapd=$Error[0] ;
        $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level warn } #Error|Warn|Debug 
        else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        #-=-record a STATUSWARN=-=-=-=-=-=-=
        $statusdelta = ";WARN"; # CHANGE|INCOMPLETE|ERROR|WARN|FAIL ;
        if(gv passstatus -scope Script -ea 0){$script:PassStatus += $statusdelta } ;
        if(gv -Name PassStatus_$($tenorg) -scope Script -ea 0){set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta)} ; 
        #-=-=-=-=-=-=-=-=
        $smsg = "FULL ERROR TRAPPED (EXPLICIT CATCH BLOCK WOULD LOOK LIKE): } catch[$($ErrTrapd.Exception.GetType().FullName)]{" ; 
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug 
        else{ write-host -foregroundcolor green "$($smsg)" } ;

        set-AdServerSettings -ViewEntireForest $false ;

        Break #Opts: STOP(debug)|EXIT(close)|CONTINUE(move on in loop cycle)|BREAK(exit loop iteration)|THROW $_/'CustomMsg'(end script with Err output)
    } ; 

    #*======V CONFIGURE DEFAULT LOGGING FROM PARENT SCRIPT NAME v======
    <# old code
    #$pltSL=@{ NoTimeStamp=$true ; Tag="($TenOrg)-LASTPASS" ; showdebug=$($showdebug) ; whatif=$($whatif) ; Verbose=$($VerbosePreference -eq 'Continue') ; } ;
    $pltSL=@{ NoTimeStamp=$FALSE ; Tag="($Ticket)" ; showdebug=$($showdebug) ; whatif=$($whatif) ; Verbose=$($VerbosePreference -eq 'Continue') ; } ;
    if($PSCommandPath){   $logspec = start-Log -Path $PSCommandPath @pltSL ;
    } else { $logspec = start-Log -Path ($MyInvocation.MyCommand.Definition) @pltSL ; } ;
    if($logspec){
        $logging=$logspec.logging ;
        $logfile=$logspec.logfile ;
        $transcript=$logspec.transcript ;
        if(Test-TranscriptionSupported){
            $stopResults = try {Stop-transcript -ErrorAction stop} catch {} ;
            start-transcript -Path $transcript ;
        } ;
    } else {throw "Unable to configure logging!" } ;
    #>
    if(!(get-variable LogPathDrives -ea 0)){$LogPathDrives = 'd','c' };
    foreach($budrv in $LogPathDrives){if(test-path -path "$($budrv):\scripts" -ea 0 ){break} } ;
    if(!(get-variable rgxPSAllUsersScope -ea 0)){
        $rgxPSAllUsersScope="^$([regex]::escape([environment]::getfolderpath('ProgramFiles')))\\((Windows)*)PowerShell\\(Scripts|Modules)\\.*\.(ps(((d|m))*)1|dll)$" ;
    } ;
    if(!(get-variable rgxPSCurrUserScope -ea 0)){
        $rgxPSCurrUserScope="^$([regex]::escape([Environment]::GetFolderPath('MyDocuments')))\\((Windows)*)PowerShell\\(Scripts|Modules)\\.*\.(ps((d|m)*)1|dll)$" ;
    } ;
    $pltSL=[ordered]@{Path=$null ;NoTimeStamp=$false ;Tag=$null ;showdebug=$($showdebug) ; Verbose=$($VerbosePreference -eq 'Continue') ; whatif=$($whatif) ;} ;
    $pltSL.Tag = ($ticket,$DNameBase -join '-') ;
    if($script:PSCommandPath){
        if(($script:PSCommandPath -match $rgxPSAllUsersScope) -OR ($script:PSCommandPath -match $rgxPSCurrUserScope)){
            $bDivertLog = $true ;
            switch -regex ($script:PSCommandPath){
                $rgxPSAllUsersScope{$smsg = "AllUsers"}
                $rgxPSCurrUserScope{$smsg = "CurrentUser"}
            } ;
            $smsg += " context script/module, divert logging into [$budrv]:\scripts"
            write-verbose $smsg  ;
            if($bDivertLog){
                if((split-path $script:PSCommandPath -leaf) -ne $cmdletname){
                    # function in a module/script installed to allusers|cu - defer name to Cmdlet/Function name
                    $pltSL.Path = (join-path -Path "$($budrv):\scripts" -ChildPath "$($cmdletname).ps1") ;
                } else {
                    # installed allusers|CU script, use the hosting script name
                    $pltSL.Path = (join-path -Path "$($budrv):\scripts" -ChildPath (split-path $script:PSCommandPath -leaf)) ;
                }
            } ;
        } else {
            $pltSL.Path = $script:PSCommandPath ;
        } ;
    } else {
        if(($MyInvocation.MyCommand.Definition -match $rgxPSAllUsersScope) -OR ($MyInvocation.MyCommand.Definition -match $rgxPSCurrUserScope) ){
             $pltSL.Path = (join-path -Path "$($budrv):\scripts" -ChildPath (split-path $script:PSCommandPath -leaf)) ;
        } elseif(test-path $MyInvocation.MyCommand.Definition) {
            $pltSL.Path = $MyInvocation.MyCommand.Definition ;
        } elseif($cmdletname){
            $pltSL.Path = (join-path -Path "$($budrv):\scripts" -ChildPath "$($cmdletname).ps1") ;
        } else {
            $smsg = "UNABLE TO RESOLVE A FUNCTIONAL `$CMDLETNAME, FROM WHICH TO BUILD A START-LOG.PATH!" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Warn } #Error|Warn|Debug
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            BREAK ;
        } ;
    } ;
    write-verbose "start-Log w`n$(($pltSL|out-string).trim())" ;
    $logspec = start-Log @pltSL ;
    $error.clear() ;
    TRY {
        if($logspec){
            $logging=$logspec.logging ;
            $logfile=$logspec.logfile ;
            $transcript=$logspec.transcript ;
            $stopResults = try {Stop-transcript -ErrorAction stop} catch {} ;
            start-Transcript -path $transcript ;
        } else {throw "Unable to configure logging!" } ;
    } CATCH {
        $ErrTrapd=$Error[0] ;
        $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
        else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
    } ;

    #*======^ CONFIGURE DEFAULT LOGGING FROM PARENT SCRIPT NAME ^======

    # -----------

    $smtpToFailThru=convertFrom-Base64String -string "dG9kZC5rYWRyaWVAdG9yby5jb20=" # simple encoded addr
    # pull the notifc smtpto from the xxxMeta.NotificationDlUs value
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
    <# 
    if($bDivertLog){ # scriptbasename will likely be a module name, use diverted start-log values
        $smtpFrom = ((split-path $pltsl.path -leaf).replace('.','-') + "@$( (Get-Variable  -name "$($TenOrg)Meta").value.o365_OPDomain )")
    } else {
        $smtpFrom = (($scriptBaseName.replace(".","-")) + "@$( (Get-Variable  -name "$($TenOrg)Meta").value.o365_OPDomain )") ;
    } ; 
    #>
    # shift to cmdletname, more dependable
    $smtpFrom = (($CmdletName.replace(".","-")) + "@$( (Get-Variable  -name "$($TenOrg)Meta").value.o365_OPDomain )") ;
    $smtpSubj= "Proc Rpt:"
    if($whatif) {
        $smtpSubj+="WHATIF:" ;
    } else {
        $smtpSubj+="PROD:" ;
    } ;
    <#
    if($bDivertLog){ # scriptbasename will likely be a module name, use diverted start-log values
        $smtpSubj+= "$((split-path $pltsl.path -leaf)):$(get-date -format 'yyyyMMdd-HHmmtt')"   ;
    } else {
        $smtpSubj+= "$($ScriptBaseName):$(get-date -format 'yyyyMMdd-HHmmtt')"   ;
    } ; 
    #>
    # shift to cmdletname, more dependable
    $smtpSubj+= "$($CmdletName):$(get-date -format 'yyyyMMdd-HHmmtt')"   ;
    if(!($bodyAsHtml)){
        # if not inline attachment in body, need to load report as attachment
        $smtpAttachment=$rptfile ;
    } else {
        #9:49 AM 3/26/2015 just blank the attachment if we're not mailing it
        $smtpAttachment=$null;
    };
    # setup body as a hash
    $smtpBody = @() ;
    # (`n = CrLf in body)
    #====== ^ EMAIL HANDLING BOILERPLATE (USE IN SUB MAIN) ^==================================


    #*======v DL PSTED IN CODE  v======
    if($host.version.major -ge 3){
        $InputSplat=[ordered]@{ Dummy = $null ;  } ;
        $pltNewDG=[ordered]@{ Dummy = $null ;  } ;
        $pltSetDG=[ordered]@{ Dummy = $null ;  } ;
    } else {
        $InputSplat=@{ Dummy = $null ;  } ;
        $pltNewDG=@{ Dummy = $null ;  } ;
        $pltSetDG=@{ Dummy = $null ;  } ;

    } ;
    $InputSplat.remove("Dummy") ;
    $InputSplat.Add("DNameBase","") ; 
    $InputSplat.Add("Ticket",$($null)) ; 
    $InputSplat.Add("ManagedBy","") ; 
    $InputSplat.Add("SiteOverride","") ; 
    $InputSplat.Add("Members","") ; 
    $InputSplat.Add("InetReceive","") ; 

    $pltNewDG.remove("Dummy") ;
    $pltNewDG.Add("DisplayName",$("")) ; 
    $pltNewDG.Add("Alias",$("")) ; 
    $pltNewDG.Add("OrganizationalUnit",$("")) ; 
    $pltNewDG.Add("SamAccountName",$($null)) ; # 1:22 PM 6/6/2016 defer SamAccountName too
    $pltNewDG.Add("type",$( "Distribution")) ; 
    $pltNewDG.Add("Notes",$( $null)) ; 
    $pltNewDG.Add("ManagedBy",$($InputSplat.ManagedBy)) ; 
    $pltNewDG.Add("whatif",$($whatif)) ; 
    $pltNewDG.Add("ErrorAction","STOP") ; 

    $pltSetDG.remove("Dummy") ;
    $pltSetDG.Add("Identity","") ; 
    $pltSetDG.Add("whatif",$($whatif)) ; 
    $pltSetDG.Add("ErrorAction","STOP") ; 

    #region SPLATDEFS ; # ------ 

    $smsg = ":===PASS STARTED=== " ;
    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
    if($DNameBase){$InputSplat.DNameBase=$DNameBase};
    if($ManagedBy){$InputSplat.ManagedBy=$ManagedBy};
    #if($Description){$InputSplat.Description=$Description};
    if($Ticket){$InputSplat.Ticket=$Ticket};
    if($SiteOverride){$InputSplat.SiteOverride=$SiteOverride};
    if($Members){$InputSplat.Members=$Members};
    if($InetReceive){$InputSplat.InetReceive=$InetReceive};
    if($HiddenFromAddressLists){$InputSplat.HiddenFromAddressLists=$HiddenFromAddressLists};


    #-=-=-=-=-=-=-=-=
    # alias block switch on rcptype, handles ExOP|Exo|Exov2 variants with single aliased cmd assignements
    # pull onprem recipipent to drive balance of logic
    #rx10 -Verbose:$false ; 
    if($pltRX10){
        ReConnect-Ex2010 @pltRX10 ;
    } else { Reconnect-Ex2010 ; } ; 
    #$OpRcp = $xoRcp = $null ; 
    if($CloudFirst){$isCloud1st = $true } else { $isCloud1st = $false } ; 
    <#
    if(!$Room.identity){ 
        $smsg = "`$Room.Idenity is BLANK! Aborting to avoid returning *entire* recipient base!" ; 
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Warn } #Error|Warn|Debug 
        else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        _cleanup ; 
    } ; 
    if($OpRcp= get-recipient -id $Room.identity -ea 0 ){
        write-verbose "successful:get-recipient -id $($Room.identity) -ea 0" 
    } else {
        write-verbose "failed to get-recipient -id $($Room.identity) -ea 0 ; retry EXO (cloud1st)" 
        $xGRcp = (gcm get-*xo*recipient).name.tolower() | select -unique ; 
        $expr = "$($xGRcp) -id $Room.identity -ea STOP" ; 
        if($xoRcp= invoke-expression $expr ){
            $isCloud1st = $true ;
        } ;
    } ; 
    #>
    # aliased ExOP|EXO|EXOv2 cmdlets (permits simpler single code block for any of the three variants of targets & syntaxes)
    # each is '[aliasname];[exOcmd] (xOv2cmd & exop are converted from [exocmd])
    [array]$cmdletMaps = 'ps1GetxRcp;get-exorecipient;','ps1GetxMbx;get-exomailbox;','ps1SetxMbx;Set-exoMailbox;','ps1GetxUser;get-exoUser;',
        'ps1GetxMUsr;Get-exoMailUser','ps1SetxMUsr;Set-exoMailUser','ps1SetxCalProc;set-exoCalendarprocessing;',
        'ps1GetxCalProc;get-exoCalendarprocessing;','ps1GetxMbxFldrPerm;get-exoMailboxfolderpermission;',
        'ps1GetxMbxFldrPerm;get-exoMailboxfolderpermission','ps1AddxMbxPrm;Add-exoMailboxPermission','ps1GetxMbxPrm;Get-exoMailboxPermission',
        'ps1RmvxMbxPrm;Remove-exoMailboxPermission','ps1AddRcpPrm;Add-exoRecipientPermission','ps1GetRcpPrm;Get-exoRecipientPermission',
        'ps1RmvRcpPrm;Remove-exoRecipientPermission','ps1GetxAccDom;Get-exoAcceptedDomain;','ps1GetxRetPol;Get-exoRetentionPolicy',
        'ps1GetxDistGrp;get-exoDistributionGroup;','ps1GetxDistGrpMbr;get-exoDistributionGroupmember;','ps1GetxMsgTrc;get-exoMessageTrace;',
        'ps1GetxMsgTrcDtl;get-exoMessageTraceDetail;','ps1GetxMbxFldrStats;get-exoMailboxfolderStatistics','ps1GetxMContact;Get-exomailcontact;',
        'ps1SetxMContact;Set-exomailcontact;','ps1NewxMContact;New-exomailcontact','ps1TestxMapi;Test-exoMAPIConnectivity',
        'ps1GetxOrgCfg;Get-exoOrganizationConfig','ps1GetxMbxRegionCfg;Get-exoMailboxRegionalConfiguration',
        'ps1TestxOAuthConn;Test-exoOAuthConnectivity','ps1NewxDistGrp;new-exoDistributionGroup','ps1SetxDistGrp;set-exoDistributionGroup',
        'ps1AddxDistGrpMbr;Add-exoDistributionGroupMember','ps1RmvxDistGrpMbr;remove-exoDistributionGroupMember',
        'ps1GetxDDG;Get-exoDynamicDistributionGroup','ps1NewxDDG;New-exoDynamicDistributionGroup','ps1SetxDDG;Set-exoDynamicDistributionGroup' ;
    [array]$XoOnlyMaps = 'ps1GetxMsgTrcDtl','ps1TestxOAuthConn' ; # cmdlet alias names from above that are skipped for aliasing in EXOP
    # cmdlets from above that have diff names EXO v EXoP: these each have  schema: [alias];[xoCmdlet];[opCmdlet]; op Aliases use the opCmdlet as target
    [array]$XoRenameMaps = 'ps1GetxMsgTrc;get-exoMessageTrace;get-MessageTrackingLog','ps1AddRcpPrm;Add-exoRecipientPermission;Add-AdPermission',
            'ps1GetRcpPrm;Get-exoRecipientPermission;Get-AdPermission','ps1RmvRcpPrm;Remove-exoRecipientPermission;Remove-ADPermission' ;
    # code to summarize & indexed-hash the renamed cmdlets for variant processing
    $XoRenameMapNames = @() ; 
    $oxoRenameMaps = @{} ;
    $XoRenameMaps | foreach {     $XoRenameMapNames += $_.split(';')[0] ;     $name = $_.split(';')[0] ;     $oxoRenameMaps[$name] = $_.split(';')  ;  } ;
    # $isExOP = $isEXO = $false ; 
    # now need to accomodate cloud1st as well
    # filtering the above to subsets:
    $cmdletMapsFltrd = $cmdletmaps|?{$_.split(';')[1] -like '*DistributionGroup*'} ; 
    $cmdletMapsFltrd += $cmdletmaps|?{$_.split(';')[1] -like '*recipient'}
    #$cmdletMapsFltrd = $cmdletmaps # or use full set
    foreach($cmdletMap in $cmdletMapsFltrd){
        <# dbg code
        write-verbose $cmdletMap ;
        if($cmdletMap -eq 'ps1AddRcpPrm Add-exoRecipientPermission'){
            write-host "GOTCHA!" ;
        } ; 
        #>
        <#switch ($OpRcp.recipienttype){
            "MailUser" {
        #>
        #if(($OpRcp.recipienttype -eq 'MailUser') -OR ($xoRcp)){
        if($isCloud1st){
            $isExOP = $false ; $isEXO = $true ; 
            if ($script:useEXOv2) { reconnect-eXO2 @pltRXO }
            else { reconnect-EXO @pltRXO } ;
            # reconnect-exo @pltRXO ;
            if($script:useEXOv2){
                if(!($cmdlet= Get-Command $cmdletMap.split(';')[1].replace('-exo','-xo') )){ throw "unable to gcm Alias definition!:$($cmdletMap.split(';')[1])" ; break }
                $nAName = ($cmdletMap.split(';')[0]) ;
                if(-not(get-alias -name $naname -ea 0 |?{$_.Definition -eq $cmdlet.name})){
                    $nalias = set-alias -name $nAName -value ($cmdlet.name) -passthru ;
                    write-verbose "$($nalias.Name) -> $($nalias.ResolvedCommandName)" ;
                } ;
            } else {
                if(!($cmdlet= Get-Command $cmdletMap.split(';')[1])){ throw "unable to gcm Alias definition!:$($cmdletMap.split(';')[1])" ; break }
                $nAName = ($cmdletMap.split(';')[0]);
                if(-not(get-alias -name $naname -ea 0 |?{$_.Definition -eq $cmdlet.name})){
                    $nalias = set-alias -name ($cmdletMap.split(';')[0]) -value ($cmdlet.name) -passthru ;
                    write-verbose "$($nalias.Name) -> $($nalias.ResolvedCommandName)" ;
                } ;
            } ;
        } else {
            $isExOP = $true ; $isEXO = $false ; 
            if($pltRX10){
                    ReConnect-Ex2010 @pltRX10 ;
            } else { Reconnect-Ex2010 ; } ; 
            if($XoOnlyMaps -contains $cmdletMap.split(';')[0]){
                write-verbose "$($cmdletMap.split(';')[1]) is an XO-Only cmdlet, skipping EXOP alias-creation" ;
            } else {
                if($XoRenameMapNames -contains $cmdletMap.split(';')[0]){
                    write-verbose "$($cmdletMap.split(';')[1]) is an XO-Renamed cmdlet, renaming for EXoP" ;
                    # sub -exoNOUN -> -NOUN using ExOP variant cmdlet
                    if(!($cmdlet= Get-Command $oxoRenameMaps[($cmdletMap.split(';')[0])][2] )){ throw "unable to gcm Alias definition!:$($oxoRenameMaps[($cmdletMap.split(';')[0])][2])" ; break }
                    $nAName = ($cmdletMap.split(';')[0]);
                    if(-not(get-alias -name $naname -ea 0 |?{$_.Definition -eq $cmdlet.name})){
                        $nalias = set-alias -name ($cmdletMap.split(';')[0]) -value ($cmdlet.name) -passthru ;
                        write-verbose "$($nalias.Name) -> $($nalias.ResolvedCommandName)" ;
                    } ;
                } else { 
                    # common cmdlets between all 3 systems
                    # sub -exoNOUN -> -NOUN
                    if(!($cmdlet= Get-Command $cmdletMap.split(';')[1].replace('-exo','-') )){ throw "unable to gcm Alias definition!:$($cmdletMap.split(';')[1])" ; break }
                    $nAName = ($cmdletMap.split(';')[0]);
                    if(-not(get-alias -name $naname -ea 0 |?{$_.Definition -eq $cmdlet.name})){
                        $nalias = set-alias -name ($cmdletMap.split(';')[0]) -value ($cmdlet.name) -passthru ;
                        write-verbose "$($nalias.Name) -> $($nalias.ResolvedCommandName)" ;
                    } ;
                } ; 

            } ; 
        } ;
        <#
            default { throw "Unrecognized recipienttype!:$($OpRcp.recipienttype)" }
        } ; 
        #>
    } ;  # loop-E
    #-=-=-=-=-=-=-=-=

    Write-Host -fore green "`nSpecified Base DL Name: $($InputSplat.DNameBase)" ;
    $error.clear() ;
    TRY {
        #if($ManagedBy){$oManagedBy = $ManagedBy | foreach-object {ps1GetxRcp -id $_ -ResultSize 1 -ea 'STOP' } | select -expand primarysmtpaddress  | select -unique ;} ; 
        if($ManagedBy){
            if($isCloud1st){
                #$oManagedBy = ps1GetxRcp -id $ManagedBy -ResultSize 1 -ea 'Continue' 
                #$oManagedBy = ps1GetxRcp -id $ManagedBy -ResultSize 25 -ea 'Continue' 
                $oManagedBy = $ManagedBy  | ps1GetxRcp -ResultSize 25 -ea 'Continue' 
            } else { 
                # resolving exo smtpaddresss could yield missing recips, pull -ea
                #$oManagedBy = get-recipient -id $ManagedBy -ResultSize 25 #-ea -ea 'Continue' 
                $oManagedBy = $ManagedBy | get-recipient  -ResultSize 25 -ea 'Continue' 
            } ;
        } ; 
    } CATCH {
        $ErrTrapd=$Error[0] ;
        $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level warn } #Error|Warn|Debug 
        else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        #-=-record a STATUSWARN=-=-=-=-=-=-=
        $statusdelta = ";WARN"; # CHANGE|INCOMPLETE|ERROR|WARN|FAIL ;
        if(gv passstatus -scope Script -ea 0){$script:PassStatus += $statusdelta } ;
        if(gv -Name PassStatus_$($tenorg) -scope Script -ea 0){set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta)} ; 
        #-=-=-=-=-=-=-=-=
        $smsg = "FULL ERROR TRAPPED (EXPLICIT CATCH BLOCK WOULD LOOK LIKE): } catch[$($ErrTrapd.Exception.GetType().FullName)]{" ; 
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug 
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        Break #Opts: STOP(debug)|EXIT(close)|CONTINUE(move on in loop cycle)|BREAK(exit loop iteration)|THROW $_/'CustomMsg'(end script with Err output)
    } ; 

    if($isCloud1st){
        $DName=("ENT" + "-DL-" + $InputSplat.DNameBase) ;
    }else {
        $InputSplat.Add("Domain",$($oManagedBy[0].identity.tostring().split("/")[0]) ) ; 
        # force 1st ManagedBy's OU (if it's an array)
        $InputSplat.SiteCode=($oManagedBy[0].identity.tostring().split('/')[1]) ;
        if($domaincontroller){
            $InputSplat.Add("DomainController",$domaincontroller) ; 
        } ; 
        $pltNewDG.Add("DomainController",$domaincontroller) ; 
        $pltSetDG.Add("DomainController",$domaincontroller) ; 

        if($InputSplat.SiteOverride){
            $SiteCode=$InputSplat.SiteOverride;
            $InputSplat.SiteCode=$InputSplat.SiteOverride;
        } else {  
            $SiteCode=$InputSplat.SiteCode.tostring();
        } ;

        if($SiteOverride -eq 'ENT'){
            # ent-named OU, but park it in the ManagedBy's OU - no park it in LYN OU - less confusing if all ENT's are in one place
            $FindOU="^OU=Distribution\sGroups,";
            #$tmpSite = ($oManagedBy[0].identity.tostring().split('/')[1]) ;
            $tmpSite = 'LYN'
            if( ($pltNewDG.OrganizationalUnit = ((Get-ADObject -filter { ObjectClass -eq 'organizationalunit' } -ea continue | ?{($_.distinguishedname -match "$($FindOU).*OU=$($tmpSite),.*") } | select distinguishedname).distinguishedname.tostring()) )) { } else { _cleanup ; Exit ;} 
            $InputSplat.Add("SiteName", $SiteCode) ;
            $DName=($SiteCode + "-DL-" + $InputSplat.DNameBase) ;
        } else { 
            # put the DG obj in the ManagedBy's site
            $FindOU="^OU=Distribution\sGroups,";
            if( ($pltNewDG.OrganizationalUnit = ((Get-ADObject -filter { ObjectClass -eq 'organizationalunit' } -ea continue | ?{($_.distinguishedname -match "$($FindOU).*OU=$($InputSplat.SiteCode),.*") } | select distinguishedname).distinguishedname.tostring()) )) { } else { _cleanup ; Exit ;} 
            $InputSplat.Add("SiteName", $SiteCode) ;
            $DName=($SiteCode + "-DL-" + $InputSplat.DNameBase) ;
        } ;
    
    } ;

    $smsg = "`$Dname:$Dname" ; 
    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

    
    $pltNewDG.Name= $pltNewDG.DisplayName = $($DName);
    $pltNewDG.Alias=$($DName.replace(" ",""));
    # shift to get-admininitials
    $admininitials=get-AdminInitials ; 
    $pltNewDG.Notes="$((get-date -format "MM/dd/yyyy"))" ;
    if($($InputSplat.Ticket)){$pltNewDG.Notes+=" #$($InputSplat.Ticket)" };
    if($isCloud1st){
        $pltNewDG.remove('SamAccountName') ; 
        $pltNewDG.remove('OrganizationalUnit') ; 
    } else {
        $pltNewDG.SamAccountName =$($DName.replace(" ","").replace("-","")) ;
    } ; 
    $pltNewDG.Type = "Distribution";
    $pltNewDG.ManagedBy =$oManagedBy.primarysmtpaddress | select -unique  ;
    $pltNewDG.Notes+=" for $($pltNewDG.ManagedBy -join ',' ) -$($admininitials)" ;

    if($members){
        $pltNewDG.members = $members | ps1GetxRcp -ErrorAction Continue | select -expand primarysmtpaddress  | select -unique ;
    } ; 

    Write-Host -fore yellow "Checking for existing $($pltNewDG.DisplayName)..."  ;
    write-verbose "$((get-date).ToString("HH:mm:ss")):`$SGSrchName:$($SGSrchName)`n`$pltNewDG.DisplayName:$($pltNewDG.DisplayName)";
    $ADGSrchName=$($pltNewDG.DisplayName);

    if($isCloud1st){
        $oDL = (ps1GetxDistGrp -identity $pltNewDG.Alias -ea silentlycontinue)
    } else { 
        $oDL = (ps1GetxDistGrp -identity $pltNewDG.Alias -domaincontroller $($domaincontroller) -ea silentlycontinue)
    } ; 

    if($oDL){
        write-verbose "Existing found: `$oDL:$($oDL.primarysmtpaddress)" ;
        write-verbose "`$oDL.DN:$($oDL.DistinguishedName)" ;
    } else {
    
        $smsg = "$((get-alias ps1NewxDistGrp).definition) w`n$(($pltNewDG|out-string).trim())" ; 
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
       
            $Exit = 0 ; 
            Do {
                Try {
                    $oDL = ps1NewxDistGrp @pltNewDG ;
                    $Exit = $Retries ; 
                } Catch {
                    Start-Sleep -Seconds $RetrySleep ; 
                    $Exit ++ ; 
                    Write-Verbose "Failed to exec cmd because: $($Error[0])" ; 
                    Write-Verbose "Try #: $Exit" ; 
                    If ($Exit -eq $Retries) {Write-Warning "Unable to exec cmd!"} ; 
                } # try-E
            } Until ($Exit -eq $Retries) # loop-E

        if($whatif){
            $smsg = "SKIPPING EXEC: Whatif-only pass";
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        } else {
            # have had sporadic resolution errors, pull it back again to ensure it's fully available
            if($isCloud1st){
                do {Write-Host "." -NoNewLine;Start-Sleep -s 1} until ($odl = ps1GetxDistGrp $oDL.primarysmtpaddress -ea silentlycontinue -resultsize 1)  ;
            } else { 
                do {Write-Host "." -NoNewLine;Start-Sleep -s 1} until ($odl = ps1GetxDistGrp $oDL.primarysmtpaddress -domaincontroller $domaincontroller -resultsize 1) ;
            } ; 
            $smsg = "`$oDL:$($oDL.primarysmtpaddress)" ;
            $smsg += "`n`$oDL.DN:$($oDL.DistinguishedName)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        }  # if-E
            
    } # if-E no DL obj

    if($oDL){
        $pltSetDG.Identity=$($oDL.alias) ;
        if($isCloud1st){
            $ExistMbrs = ps1GetxDistGrpMbr -Identity $oDL.primarysmtpaddress -ErrorAction 'Stop' | select -expand primarysmtpaddress ; 
        } else { 
            $ExistMbrs = ps1GetxDistGrpMbr -Identity $odl.SamAccountName -DomainController $domaincontroller -ErrorAction 'Stop' | select -expand primarysmtpaddress ; 
        } ; 

        $pltAddDGM=[ordered]@{
            identity=$pltSetDG.identity ;
            #Member= $mbr  ; 
            ErrorAction = 'Stop' ; 
            whatif=$($whatif) ; 
            DomainController= $domaincontroller
        } ;
        if($isCloud1st){
            $pltAddDGM.remove('DomainController') ; 
        } 
        $error.clear() ;
        TRY {
            foreach($Mbr in $pltNewDG.members){
                if ($ExistMbrs -notcontains $Mbr) {
                    $smsg = "ADD:$($mbr.samaccountname)"
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                    ps1AddxDistGrpMbr @pltAddDGM -member $mbr ; 
                    
                } else {
                    $smsg = "SKIPPING:$($mbr) is already a member of $($oDL.samaccountname)"
                } ; 
            }; # loop-E

            if($InputSplat.InetReceive){
                $smsg = "-InetReceive:`$true:Updating RequireSenderAuthenticationEnabled to `$false" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                $pltSetDG.add("RequireSenderAuthenticationEnabled",$false) ; 
            
            } ; 

            if($InputSplat.HiddenFromAddressLists){
                $smsg = "-HiddenFromAddressLists:`$true:Setting HiddenFromAddressListsEnabled `$true" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                $pltSetDG.add("HiddenFromAddressListsEnabled",$true) ; 
            } ; 

            $smsg = "$((get-alias ps1SetxDistGrp).definition) w`n$(($pltSetDG|out-string).trim())" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

            ps1SetxDistGrp @pltSetDG 

        } CATCH {
            $ErrTrapd=$Error[0] ;
            $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            #-=-record a STATUSWARN=-=-=-=-=-=-=
            $statusdelta = ";WARN"; # CHANGE|INCOMPLETE|ERROR|WARN|FAIL ;
            if(gv passstatus -scope Script -ea 0){$script:PassStatus += $statusdelta } ;
            if(gv -Name PassStatus_$($tenorg) -scope Script -ea 0){set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta)} ; 
            #-=-=-=-=-=-=-=-=
            $smsg = "FULL ERROR TRAPPED (EXPLICIT CATCH BLOCK WOULD LOOK LIKE): } catch[$($ErrTrapd.Exception.GetType().FullName)]{" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            Break #Opts: STOP(debug)|EXIT(close)|CONTINUE(move on in loop cycle)|BREAK(exit loop iteration)|THROW $_/'CustomMsg'(end script with Err output)
        } ; 

        do{     $smsg =  "REVIEW SETTINGS: " ;
            $smsg = "$("="*6)`n$((Get-Date -Format 'HH:mm:ss')):Results:";
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            if($isCloud1st){
                $oDL=ps1GetxDistGrp -id $odl.Alias -ErrorAction stop ;
            } else { 
                $oDL=ps1GetxDistGrp -id $odl.Alias -domaincontroller $domaincontroller -ErrorAction stop ;
            } ; 
            # 10:19 AM 4/4/2017 add RequireSenderAuthenticationEnabled
            $propsDG = "DisplayName","Alias","WindowsEmailAddress","ManagedBy","RequireSenderAuthenticationEnabled","HiddenFromAddressListsEnabled" ; 
            $smsg = "`n$(($oDL| fl $propsDG |out-string).trim())" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            if($isCloud1st){
                $oDLMbrs = ps1GetxDistGrpMbr -identity $oDL.alias  -ea 0 | select primarysmtpaddress ; 
            } else { 
                $oDLMbrs = ps1GetxDistGrpMbr -identity $oDL.alias -domaincontroller $($domaincontroller) -ea 0 | select distinguishedname;
            } ; 
            $smsg = "`n$(($oDL| fl $propsDG |out-string).trim())" ; 
            $smsg += "`nMembers:`n$(($oDLMbrs|out-string).trim())" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        
            if(-not($OutObject)){
                $bRet=Read-Host "Enter Y to Refresh Review (replication latency)." ;
            } ; 
        } until ($bRet -ne "Y" -OR $OutObject);
        # 1:07 PM 9/30/2021 rem-out the mailcontact creation code, needs debugging. 
        if($isCloud1st -and -not($whatif)){
            # check for onprem recipient on smtpaddr, if none, offer to build a MailContact in unreplicated ($($TenOrg)meta.UnreplicatedOU)
            if($UseOP){
                Reconnect-Ex2010 @pltRX10 ; 
                if($existRcp = get-recipient -id $odl.primarysmtpaddress -domaincontroller $domaincontroller -ErrorAction 0){
                    $smsg = "(existing recipient object for $($odl.primarysmtpaddress) found:$($existRcp.recipienttypedetails) - skipping MContact creation)" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                } else {
                    $smsg = "No conflicting OnPrem recipient found with: $($odl.primarysmtpaddress)" ; 
                    $smsg += "`nDo you want to create an *unreplicated* OnPrem MailContact to point at the EXO object?" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    #*======v PS Simple YYY Confirm Prompt - psb-PSPrompt.cbp v======
                    $bRet=Read-Host "Enter YYY to continue. Anything else will exit" 
                    if ($bRet.ToUpper() -eq "YYY") {
                         Write-host "Moving on"
                            
                            # new-mailcontact -DisplayName -Name -LastName -DomainController -WhatIf -ExternalEmailAddress -Alias -PrimarySmtpAddress -OrganizationalUnit
                            $pltNewMC=[ordered]@{
                                DisplayName = "$($odl.name)-MC" ;
                                Name = "$($odl.name)-MC" ;
                                LastName = "$($odl.name)-MC";
                                DomainController = $domaincontroller;
                                ExternalEmailAddress = $odl.primarysmtpaddress;
                                Alias =  "$($odl.name.replace(' ',''))_$((new-guid).tostring().split('-')[-1])";
                                OrganizationalUnit = "OU=Unreplicated Contacts,$( (Get-Variable  -name "$($TenOrg)Meta").value.UnreplicatedOU )"
                                ErrorAction = 'STOP';
                                WhatIf = $($whatif);
                            } ; 
                            $smsg = "New-MailContact w`n$(($pltNewMC|out-string).trim())" ; 
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                            $error.clear() ;
                            TRY {
                                $nmc = new-mailcontact @pltNewMC ; 
                                $propsMC = 'name','alias','recipienttype','primarysmtpaddress' ;
                                $smsg = "`n$(($nmc|ft -a $propsMC|out-string).trim())" ; 
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            } CATCH {
                                $ErrTrapd=$Error[0] ;
                                $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                                else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                #-=-record a STATUSWARN=-=-=-=-=-=-=
                                $statusdelta = ";WARN"; # CHANGE|INCOMPLETE|ERROR|WARN|FAIL ;
                                if(gv passstatus -scope Script -ea 0){$script:PassStatus += $statusdelta } ;
                                if(gv -Name PassStatus_$($tenorg) -scope Script -ea 0){set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta)} ; 
                                #-=-=-=-=-=-=-=-=
                                $smsg = "FULL ERROR TRAPPED (EXPLICIT CATCH BLOCK WOULD LOOK LIKE): } catch[$($ErrTrapd.Exception.GetType().FullName)]{" ; 
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug 
                                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                Break #Opts: STOP(debug)|EXIT(close)|CONTINUE(move on in loop cycle)|BREAK(exit loop iteration)|THROW $_/'CustomMsg'(end script with Err output)
                            } ; 

                    } else {
                         Write-Host "Invalid response. Skipping Contact creation"
                         # exit <asserted exit error #>
                         ;;exit 1

                    } # if-block end

                }; 
            } ; 
        } ; 
        #

    } else {
        if(!($Whatif)){
            $smsg =   ("FIND/CREATION FAILURE: $($InputSplat.DNameBase) not found.`n") ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug 
            else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        }
        else {
            $smsg = "Whatif-pass completed";
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        }
    };

    _cleanup ; # pipeline, issue is test-transcribing et all are dumping a $true |write-output and trashing the pl, need to refactor the verb-logging content to fix
    # for now move the return below _cleanup

    $smsg += $sBnr.replace('=v','=^').replace('v=','^=') ;
    $smsg += "`n-----------------------"; 
    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
    
    if($OutObject){
        $smsg = "(-OutObject specified: returning DG object to pipeline)" ; 
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
        $oDL | write-output ; 
    } ; 
    
    #write-host "xxx"
}

#*------^ new-DgTor.ps1 ^------