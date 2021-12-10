+#*------v get-xoHistSearch.ps1 v------
function get-xoHistSearch {
    <#
    .SYNOPSIS
    get-xoHistSearch.ps1 - wrapper/automation for EXO's get-HistoricalSearch cmdlet, Assembles ReportTitle & models an export-csv filename, around recipient, sender, reportType etc params specified for get-historicalsearch, also dawdle loops monitoring & alerting the progress of the associated PSJob created by the search submission.
    .NOTES
    Version     : 1.0.0
    Author      : Todd Kadrie
    Website     : http://www.toddomation.com
    Twitter     : @tostka / http://twitter.com/tostka
    CreatedDate : 2021-07-13
    FileName    : get-xoHistSearch.ps1
    License     : MIT License
    Copyright   : (c) 2021 Todd Kadrie
    Github      : https://github.com/tostka/verb-XXX
    Tags        : Powershell
    AddedCredit : REFERENCE
    AddedWebsite: URL
    AddedTwitter: URL
    REVISIONS
    * 2:40 PM 12/10/2021 more cleanup 
    * 12:49 PM 9/28/2021 init; added MsgID support; added MsgID example
    .DESCRIPTION
    get-xoHistSearch.ps1 - wrapper/automation for EXO's get-HistoricalSearch cmdlet, Assembles ReportTitle & models an export-csv filename, around recipient, sender, reportType etc params specified for get-historicalsearch, also dawdle loops monitoring & alerting the progress of the associated PSJob created by the search submission.
    .PARAMETER Requester
    Requester identifier[-Requester user@domain.com]
    .PARAMETER Ticket
    Ticket # [-Ticket nnnnn]
    .PARAMETER Days
    Days of History to retrieve (from current time, defaults to 30) # [-Days nnnnn]
    .PARAMETER StartDate
    Optional StartDate (use of -Days will autocalc from current datetime)[-StartDate mm/dd/yyyy]
    .PARAMETER EndDate
    Optional EndDate (use of -Days will autocalc from current datetime) [-EndDate mm/dd/yyyy]
    .PARAMETER Recipients
    RecipientAddresses [-Recipients 'recip1@domain.com','recip2@domain.com']
    .PARAMETER Sender
    SenderAddress [-Sender 'sender@domain.com']
    .PARAMETER MessageID
    MessageID to be traced [-MessageID '<XXXXX@XXXXX.namprd04.prod.outlook.com>']
    .PARAMETER useEXOv2
    Use EXOv2 (ExchangeOnlineManagement) over basic auth legacy connection [-useEXOv2]
    .INPUTS
    None. Does not accepted piped input.(.NET types, can add description)
    .OUTPUTS
    Returns report to pipeline
    .EXAMPLE
    PS> $pltHS = [ordered]@{ Ticket = 999999;
            Requester = 'uid@domain.com' ;
            Days = 30 ;
            Recipient = $null ;
            Sender = 'sender@domain.com' ;
            NotifyAddress = 'notify@domain.com' ;
            verbose = $true ;
            showdebug = $true ;
         } ;
    write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):get-xoHistSearch w`n$(($pltHS|out-string).trim())" ;
    get-xoHistSearch @pltHS ;
    Demo splatted-params search against Days and Sender
    .EXAMPLE
    PS> $pltHS = [ordered]@{ Ticket = 'toddvac';
            Requester = 'uid@domain.com' ;;
            Days = 25 ;
            Recipient = $null ;
            Sender = sender@domain.com' ;
            NotifyAddress = 'notify@domain.com' ;
            MessageId = '<xxxxxx...@CH0PR04MB8147.namprd04.prod.outlook.com>' ;
            verbose = $true ;
            showdebug = $true ;
        } ;
        write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):get-xoHistSearch w`n$(($pltHS|out-string).trim())" ;
        get-xoHistSearch @pltHS ;
        Demo splatted-params MessageID HistoricalSearch
    .LINK
    https://docs.microsoft.com/en-us/powershell/module/exchange/get-historicalsearch
    .LINK
    https://github.com/tostka/verb-exo
    #>
    ###Requires -Version 5
    #Requires -Modules ExchangeOnlineManagement,verb-Auth, verb-IO, verb-logging, verb-Text
    ###Requires -Modules ActiveDirectory, AzureAD, MSOnline, ExchangeOnlineManagement, verb-ADMS, verb-Auth, verb-Ex2010, verb-EXO, verb-IO, verb-logging, verb-Text
    #Requires -RunasAdministrator
    # VALIDATORS: [ValidateNotNull()][ValidateNotNullOrEmpty()][ValidateLength(24,25)][ValidateLength(5)][ValidatePattern("(lyn|bcc|spb|adl)ms6(4|5)(0|1).(china|global)\.ad\.COMPANY\.com")][ValidateSet("USEA","GBMK","AUSYD")][ValidateScript({Test-Path $_ -PathType 'Container'})][ValidateScript({Test-Path $_})][ValidateRange(21,65)][ValidateCount(1,3)]
    ## [OutputType('bool')] # optional specified output type
    [CmdletBinding(DefaultParameterSetName='Days')]
    #[CmdletBinding()]
    #[Alias('gxhs')]
    PARAM(
        [Parameter(Mandatory=$False,HelpMessage="TenantTag value, indicating Tenants to connect to[-TenOrg 'TOL']")]
        $TenOrg = 'TOR',
        [Parameter(Mandatory=$False,HelpMessage="Requester identifier[-Requester user@domain.com]")]
        $Requester,
        [Parameter(Mandatory=$False,HelpMessage="Ticket # [-Ticket nnnnn]")]
        $Ticket,
        [Parameter(ParameterSetName='Days',Mandatory=$False,HelpMessage="Days of History to retrieve (from current time, defaults to 30) # [-Days nnnnn]")]
        [int]$Days=30,
        [Parameter(ParameterSetName='Date',Mandatory=$False,HelpMessage="Optional StartDate (use of -Days will autocalc from current datetime)[-StartDate mm/dd/yyyy]")]
        [DateTime]$StartDate,
        [Parameter(ParameterSetName='Date',Mandatory=$False,HelpMessage="Optional EndDate (use of -Days will autocalc from current datetime) [-EndDate mm/dd/yyyy]")]
        [DateTime]$EndDate,
        [Parameter(Mandatory=$False,HelpMessage="RecipientAddresses [-Recipients 'recip1@domain.com','recip2@domain.com']")]
        [string]$Recipients,
        [Parameter(Mandatory=$False,HelpMessage="SenderAddress [-Sender 'sender@domain.com']")]
        [string]$Sender,
        [Parameter(Mandatory=$False,HelpMessage="MessageID to be traced [-MessageID '<XXXXX@XXXXX.namprd04.prod.outlook.com>']")]
        [string]$MessageID,
        [Parameter(Mandatory=$False,HelpMessage="Result status Notification Address [-NotifyAddress 'recipx@domain.com']")]
        [string]$NotifyAddress,
        [Parameter(HelpMessage="Use EXOv2 (ExchangeOnlineManagement) over basic auth legacy connection [-useEXOv2]")]
        [switch] $useEXOv2,
        [Parameter(HelpMessage="Debugging Flag [-showDebug]")]
        [switch] $showDebug
    ) ;
    BEGIN{
        ${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name ;
        $Verbose = ($VerbosePreference -eq 'Continue') ;
        #$rgxEmailAddr = "^([0-9a-zA-Z]+[-f._+&'])*[0-9a-zA-Z]+@([-0-9a-zA-Z]+[.])+[a-zA-Z]{2,63}$" ;
        #$rgxDName = "^([a-zA-Z]{2,}\s[a-zA-Z]{1,}'?-?[a-zA-Z]{2,}\s?([a-zA-Z]{1,})?)" ;
        #$rgxSamAcctNameTOR = "^\w{2,20}$" ; # up to 20k, the limit prior to win2k
        #$rgxSamAcctName = "^[^\/\\\[\]:;|=,+?<>@?]+$" # no char limit ;

        $propsJob = "Status",@{name="Stat"; expression={$_.ReportStatusDescription}},@{name="From"; expression={$_.SenderAddress}},@{name="To"; expression={$_.RecipientAddress}},@{name="Prog"; expression={$_.JobProgress}},@{name="ETA"; expression={(get-date ($_.EstimatedCompletionTime.ToLocalTime()) -f 'MM/dd HH:mmtt')}} ;
        $propsJobResults = 'JobId','FileRows','ErrorCode','ErrorDescription','Status','ReportStatusDescription','SenderAddress','RecipientAddress','MessageID','CompletionDate','JobProgress','EstimatedCompletionTime','FileUrl' ;

        <#
        $progInterval= 500 ; # write-progress wait interval in ms
        $DoRetries = 4 ;
        $RetrySleep = 5 ;
        [int]$retryLimit=1; # just one retry to patch lineuri duped users and retry 1x
        [int]$retryDelay=20;    # secs wait time after failure
        #>

        #$ComputerName = $env:COMPUTERNAME ;
        #$sQot = [char]34 ; $sQotS = [char]39 ;

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

        #====== v EMAIL HANDLING BOILERPLATE (USE IN SUB MAIN) v==================================
        $bodyAsHtml=$true ;
        $smtpPriority="Normal";
        # SMTP port (default is 25)
        $smtpPort = 25 ;
        $smtpToFailThru="dG9kZC5rYWRyaWVAdG9yby5jb20="| convertfrom-Base64String
        # pull the notifc smtpto from the xxxMeta.NotificationDlUs value
        if(!$showdebug){
            if((Get-Variable  -name "$($TenOrg)Meta").value.NotificationDlUs){
                $smtpTo = (Get-Variable  -name "$($TenOrg)Meta").value.NotificationDlUs ;
            }elseif((Get-Variable  -name "$($TenOrg)Meta").value.NotificationAddr1){
                $smtpTo = (Get-Variable  -name "$($TenOrg)Meta").value.NotificationAddr1 ;
            } else {$smtpTo=$smtpToFailThru} ;
        } else {
            # debug pass, don't send to main dl, use NotificationAddr1    if((Get-Variable  -name "$($TenOrg)Meta").value.NotificationDlUs){
            if((Get-Variable  -name "$($TenOrg)Meta").value.NotificationAddr1){
                #set-variable -Name $meta.name -Value ((get-variable -name $meta.name).value  += @{'o365_AcceptedDomains' = (Get-exoAcceptedDomain).domainname} )
                $smtpTo = (Get-Variable  -name "$($TenOrg)Meta").value.NotificationAddr1 ;
            } else {$smtpTo=$smtpToFailThru } ;
        } ;
        $smtpFrom = (($scriptBaseName.replace(".","-")) + "@$( (Get-Variable  -name "$($TenOrg)Meta").value.o365_OPDomain )") ;
        $smtpSubj= "Proc Rpt:"
        if($whatif) {$smtpSubj+="WHATIF:" }
        else {$smtpSubj+="PROD:" } ;
        $smtpSubj+= "$($ScriptBaseName):$(get-date -format 'yyyyMMdd-HHmmtt')"   ;
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
            # 10:32 AM 9/14/2021: _cleanup(): # only mail on PassStatus
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
            $pltCleanup=@{LogPath=$tmpcopy summarizeStatus=$true ;  NoTranscriptStop=$true ; showDebug=$($showDebug) ;  whatif=$($whatif) ; } ;
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
                    # 8:46 AM 3/11/2015 shift the logfilename gen out here, so that we can arch it
                    #$Logname= (join-path -path (join-path -path $scriptDir -childpath "logs") -childpath ($scriptNameNoExt + "-" + (get-date -uformat "%Y%m%d-%H%M" ) + "-ISEtrans.log")) ;
                    # 2:16 PM 4/27/2015 shift to static timestamp $timeStampNow
                    #$Logname= (join-path -path (join-path -path $scriptDir -childpath "logs") -childpath ($scriptNameNoExt + "-" + $timeStampNow + "-ISEtrans.log")) ;
                    # 2:02 PM 9/21/2018 missing $timestampnow, hardcode
                    $Logname=(join-path -path (join-path -path $scriptDir -childpath "logs") -childpath ($scriptNameNoExt + "-" + (get-date -format 'yyyyMMdd-HHmmtt') + "-ISEtrans.log")) ;
                    $smsg = "`$Logname: $($Logname)";
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    Start-iseTranscript -logname $Logname -Verbose:($VerbosePreference -eq 'Continue') ;
                    #Archive-Log $Logname -Verbose:($VerbosePreference -eq 'Continue');
                    # 1:23 PM 4/23/2015 standardize processing file so that we can send a link to open the transcript for review
                    $transcript = $Logname
                } else {
                    $smsg = "Stop Transcript" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Warn } #Error|Warn|Debug
                    else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    Stop-TranscriptLog -Verbose:($VerbosePreference -eq 'Continue') ;
                    #if($showdebug){ $smsg = "Archive Transcript" };
                    #Archive-Log $transcript -Verbose:($VerbosePreference -eq 'Continue') ;
                } # if-E
            } else {
                $smsg = "(_cleanup(): deferring transcript stop to main script)" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } ; # !$NoTranscriptStop
            # add trailing notifc
            $smsg = "Mailing Report" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
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
                                if((get-Variable -Name PassStatus_$($tenorg)).value.split(';') |Where-Object{$_ -ne ''}){
                                    $SmtpBody += (summarize-PassStatus -PassStatus (get-Variable -Name PassStatus_$($tenorg)).value -verbose:$($VerbosePreference -eq 'Continue') );
                                } ;
                            } else {
                                $SmtpBody += "(no processing of mailboxes in $($TenOrg), this pass)" ;
                            } ;
                            $SmtpBody += "`n" ;

                        } ;
                    } ;

                    if($PassStatus){
                        if($PassStatus.split(';') |Where-Object{$_ -ne ''}){
                            $SmtpBody += (summarize-PassStatus -PassStatus $PassStatus -verbose:$($VerbosePreference -eq 'Continue') );
                        } ;
                    } else {
                        $SmtpBody += "(no `$PassStatus updates, this pass)" ;
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
                        $Alltranscripts |ForEach-Object{
                            #$archedTrans = join-path -path $ArchPath -childpath (split-path $transcript -leaf) ;
                            $archedTrans = join-path -path $ArchPath -childpath (split-path $_ -leaf) ;
                            $smtpBody += "`nTranscript accessible at:`n$($archedTrans)`n" ;
                        } ;
                    } ;
                #};
            }
            $SmtpBody += "Pass Completed $([System.DateTime]::Now)" ;
            $SmtpBody += "`n" + $MailBody
            # body rendered in OL loses all wordrwraps
            # force strip out the html
            #$smtpBody = [regex]::Replace($smtpBody, "\<[^\>]*\>", '') ;
            $styleCSS = "<style>BODY{font-family: Arial; font-size: 10pt;}" ;
            $styleCSS += "TABLE{border: 1px solid black; border-collapse: collapse;}" ;
            $styleCSS += "TH{border: 1px solid black; background: #dddddd; padding: 5px; }" ;
            $styleCSS += "TD{border: 1px solid black; padding: 5px; }" ;
            $styleCSS += "</style>" ;
            <#
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
            # # always force
            #if($script:passstatus.split(';') -contains 'ERROR'){
            # only mail on PassStatus
            if([string]::IsNullOrEmpty($script:PassStatus)){
                $smsg = "No Email Report: `$script:PassStatus isNullOrEmpty" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } else {

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
                $smsg = "Send-EmailNotif w`n$(($Email|out-string).trim())" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                Send-EmailNotif @Email ;
            } ;
            #add an exit comment
            $smsg = "END $BARSD4 $scriptBaseName $BARSD4"
            $smsg += "`n$BARSD40"
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            # finally restore the DebugPref if set
            if ($ShowDebug -OR ($DebugPreference = "Continue")) {
                $smsg = "Resetting `$DebugPreference from 'Continue' back to default 'SilentlyContinue'" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                $showdebug=$false
                # 8:41 AM 10/13/2015 also need to enable write-debug output (and turn this off at end of script, it's a global, normally SilentlyContinue)
                $DebugPreference = "SilentlyContinue" ;
            } # if-E
            $smsg= "#*======^ END PASS:$($ScriptBaseName) ^======" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            if(!$NoTranscriptStop){
                EXIT # trailing tempfile cleanup in the sub main
            } ;
        } #*------^ END Function _cleanup ^------

        #*======v SUB MAIN  v====== (not really, but it's a landmark for post-functions exec)

        #rx10 -Verbose:$false ;
        #rxo  -Verbose:$false ; cmsol  -Verbose:$false ;

        $sBnr="`n#*======v $(${CmdletName}) : v======" ;
        $smsg = $sBnr ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

        # email trigger vari, and email body aggretating log
        $PassStatus = $MailBody = $null ;

        # add try catch as well - this may be making it zero-tolerance and catching all minor errors, disable it
        #Set-StrictMode -Version 2.0 ;

        #*------v SERVICE CONNECTIONS v------
        #-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
        # steer all onprem code on $XXXMeta.ExOPAccessFromToro & Ex10Server values
        $UseOP=$false ;
        <#
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
        #>
        $UseOP=$false ;

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
                # make it script scope, so we don't have to predetect & purge before using new-variable
                if(get-Variable -Name cred$($tenorg) -scope Script){
                    Remove-Variable -Name cred$($tenorg) -scope Script
                } ;
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
                if(get-Variable -Variable -Name "cred$($tenorg)OP" -scope Script){
                    Remove-Variable -Name "cred$($tenorg)OP" -scope Script ;
                } ;
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
            if(($global:ADPsDriveNames|Measure-Object).count){
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
        $domaincontroller = get-GCFastXO -TenOrg $TenOrg -subdomain ((Get-Variable -name "$($TenOrg)Meta").value['OP_ExADRoot']) -verbose:$($verbose) |Where-Object{$_.length};


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
        #-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
        #*------^ END SERVICE CONNECTIONS ^------


        $error.clear() ;
        TRY {
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
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            #-=-record a STATUSWARN=-=-=-=-=-=-=
            $statusdelta = ";WARN"; # CHANGE|INCOMPLETE|ERROR|WARN|FAIL ;
            if(Get-Variable passstatus -scope Script -ea 0){$script:PassStatus += $statusdelta } ;
            if(Get-Variable -Name PassStatus_$($tenorg) -scope Script -ea 0){set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta)} ;
            #-=-=-=-=-=-=-=-=
            $smsg = "FULL ERROR TRAPPED (EXPLICIT CATCH BLOCK WOULD LOOK LIKE): } catch[$($ErrTrapd.Exception.GetType().FullName)]{" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug
            else{ write-host -foregroundcolor green "$($smsg)" } ;

            #set-AdServerSettings -ViewEntireForest $false ;

            Break #Opts: STOP(debug)|EXIT(close)|CONTINUE(move on in loop cycle)|BREAK(exit loop iteration)|THROW $_/'CustomMsg'(end script with Err output)
        } ;

        #*======V CONFIGURE DEFAULT LOGGING FROM PARENT SCRIPT NAME v======
        if(!(get-variable LogPathDrives -ea 0)){$LogPathDrives = 'd','c' };
        foreach($budrv in $LogPathDrives){if(test-path -path "$($budrv):\scripts" -ea 0 ){break} } ;
        if(!(get-variable rgxPSAllUsersScope -ea 0)){
            $rgxPSAllUsersScope="^$([regex]::escape([environment]::getfolderpath('ProgramFiles')))\\((Windows)*)PowerShell\\(Scripts|Modules)\\.*\.(ps(((d|m))*)1|dll)$" ;
        } ;
        if(!(get-variable rgxPSCurrUserScope -ea 0)){
            $rgxPSCurrUserScope="^$([regex]::escape([Environment]::GetFolderPath('MyDocuments')))\\((Windows)*)PowerShell\\(Scripts|Modules)\\.*\.(ps((d|m)*)1|dll)$" ;
        } ;
        $pltSL=[ordered]@{Path=$null ;NoTimeStamp=$false ;Tag=$null ;showdebug=$($showdebug) ; Verbose=$($VerbosePreference -eq 'Continue') ; whatif=$($whatif) ;} ;
        $pltSL.Tag = $ticket -join ',' ;
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



    } # BEGIN-E
    PROCESS{

        $pltHS=@{
            ReportTitle=$null ;
            StartDate=(get-date ([datetime]::Today)).adddays($days * -1) ;
            EndDate=(get-date) ;
            ReportType="MessageTrace" ;
            NotifyAddress=$NotifyAddress ;
        } ;
        if($StartDate){
          $smsg = "(setting `$pltHS.StartDate to `$StartDate)" ;
          if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
          else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
          $pltHS.StartDate = $StartDate
        } ;
        if($EndDate){
            $smsg = "(setting `$pltHS.EndDate to `$EndDate)" ;
          if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
          else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            $pltHS.EndDate = $EndDate
        } ;
        if($MessageID){
            $pltHs.add('MessageID',$MessageID) ;
        } ;
        if($ticket){$pltHS.ReportTitle = "$($ticket)-" };
        if($Requester){$pltHS.ReportTitle += "$($Requester) " } ;
        $pltHS.ReportTitle = "$($pltHS.ReportType) " ;
        if($Recipients){
            $pltHS.add("RecipientAddress","$($Recipients)" ) ;
            $pltHS.ReportTitle = "TO-$($recip) " ;
        } ;
        if($Sender){
          $pltHS.add("SenderAddress","$($Sender)" ) ;
          $pltHS.ReportTitle += "FROM-$($Sender) " ;
        } ;
        if($days){
            $pltHS.ReportTitle += "$($days)D-History"    } else {        $pltHS.ReportTitle += "$(get-date $pltHS.StartDate -format 'yyyyMMdd-HHmmtt')" ;
            $pltHS.ReportTitle += "-$(get-date $pltHS.EndDate -format 'yyyyMMdd-HHmmtt')" ;
        } ;

        $smsg = "===Start-ExoHistoricalSearch w `$pltHS:$($pltHS|out-string)" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

        TRY {
            $HSo=Start-ExoHistoricalSearch @pltHS ;
        } CATCH {
            $ErrTrapd=$Error[0] ;
            $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            #-=-record a STATUSWARN=-=-=-=-=-=-=
            $statusdelta = ";WARN"; # CHANGE|INCOMPLETE|ERROR|WARN|FAIL ;
            if(Get-Variable passstatus -scope Script -ea 0){$script:PassStatus += $statusdelta } ;
            if(Get-Variable -Name PassStatus_$($tenorg) -scope Script -ea 0){set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta)} ;
            #-=-=-=-=-=-=-=-=
            $smsg = "FULL ERROR TRAPPED (EXPLICIT CATCH BLOCK WOULD LOOK LIKE): } catch[$($ErrTrapd.Exception.GetType().FullName)]{" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            Break #Opts: STOP(debug)|EXIT(close)|CONTINUE(move on in loop cycle)|BREAK(exit loop iteration)|THROW $_/'CustomMsg'(end script with Err output)
        } ;

        $smsg = "===Confirming new HS Job:$($Hso.JobID)" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

        $error.clear() ;
        TRY {
            $oHSJob = Get-ExoHistoricalSearch -JobID $hSO.jobID ;
        } CATCH {
            $ErrTrapd=$Error[0] ;
            $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            #-=-record a STATUSWARN=-=-=-=-=-=-=
            $statusdelta = ";WARN"; # CHANGE|INCOMPLETE|ERROR|WARN|FAIL ;
            if(Get-Variable passstatus -scope Script -ea 0){$script:PassStatus += $statusdelta } ;
            if(Get-Variable -Name PassStatus_$($tenorg) -scope Script -ea 0){set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta)} ;
            #-=-=-=-=-=-=-=-=
            $smsg = "FULL ERROR TRAPPED (EXPLICIT CATCH BLOCK WOULD LOOK LIKE): } catch[$($ErrTrapd.Exception.GetType().FullName)]{" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            Break #Opts: STOP(debug)|EXIT(close)|CONTINUE(move on in loop cycle)|BREAK(exit loop iteration)|THROW $_/'CustomMsg'(end script with Err output)
        } ;

        $1F=$false ;
        Do {
            if($1F){Start-Sleep -s 60} ;
            write-host "." -NoNewLine ;
            $1F=$true ;
            $oHSJob = Get-ExoHistoricalSearch -JobID $hSO.jobID ;

            $smsg = "`n$(($oHSJob | ft -auto $propsJob|out-string).trim())" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

        } Until ($oHSJob.status -eq 'Done') ;

        $ofile ="$($ticket)-$($Requester)-EXOHistSrch" ;
        #$ofile+=",$($pltHS.ReportTitle)" ;
        $ofile+=",$(create-AcronymFromCaps $pltHS.ReportType)" ;
        if($pltHS.SenderAddress){$ofile+=",From-$($pltHS.SenderAddress.replace("*","ANY"))" } ;
        if($pltHS.RecipientAddress){$ofile+=",To-$($pltHS.RecipientAddress.replace("*","ANY"))" } ;
        if($pltHS.MessageID){$ofile+=",MsgID-$($pltHS.replace('<','').replace('>','').substring(0,8))" } ;
        $ofile+=",$(get-date $pltHS.StartDate -format 'yyyyMMdd-HHmm')-$(get-date $pltHS.EndDate -format 'yyyyMMdd-HHmm')" ;
        #$ofile+= ",run-$(get-date -format 'yyyyMMdd-HHmmtt').csv" ;
        $ofile+= ".csv" ;
        $ofile=[RegEx]::Replace($ofile, "[{0}]" -f ([RegEx]::Escape(-join [System.IO.Path]::GetInvalidFileNameChars())), '') ;

        if($env:computername -match $rgxMyWorkstationsW){
            $ofile="c:\usr\work\incid\$($ofile)"
        } else {
            $drvhunt = 'd','c' ;
            foreach($dh in $drvhunt){
                $dhpath = "$($dh):\scripts\logs" ;
                if(test-path $dhpath ){ break } else {$dhpath = $null} ;
            } ;
            if($dhpath){$ofile="$($dhpath)\$($ofile)"}
        } ;

        write-host "`a" ;
        write-host "`a" ;
        write-host "`a" ;
        if($oHSJob.CompletionDate){
            $ts = New-TimeSpan -Start $oHSJob.SubmitDate -End $oHSJob.CompletionDate
        } else {
            $ts = New-TimeSpan -Start $oHSJob.SubmitDate -End (get-date) ;
        } ;
        $msg = "DONE!`n(Report gen took: {0:g} (h:m:s:ms)`n" -f $ts ;
        $msg += "`n$(($oHSJob |fl $propsJobResults |out-string).trim())`n`nUse CSV filename:`n$($ofile)`n" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

        $smsg = "(also avail at https://admin.exchange.microsoft.com/#/messagetrace under 'Downloadable reports')" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

        $msg = "Post-download, to convert the download .csv to [name]-expanded.csv (MessageTrace csv equivelent) `nconvert-HistoricalSearchCSV.ps1 -ToCSV -Files $($ofile)"  ;
        $smsg += "`nNOTE:MS ENCODING BREAK: Force the encoding when direct import-csv'ing the raw HistSearch .csv! `nimport-csv -encoding unicode -path '$($ofile)'" ;
        #write-host $msg ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;


    }
    END{
        _Cleanup

    } ;
}

#*------^ get-xoHistSearch.ps1 ^------