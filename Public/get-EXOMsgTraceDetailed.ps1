#*------v get-EXOMsgTraceDetailed.ps1 v------
function get-EXOMsgTraceDetailed {
    <#
    .SYNOPSIS
    get-EXOMsgTraceDetailed.ps1 - Run a MessageTrace with output summarizing, export to csv, and optional followup with MessageTraceDetail, summarize (expand TransportRules opt), and export to csv
    .NOTES
    Version     : 1.0.0
    Author      : Todd Kadrie
    Website     : http://www.toddomation.com
    Twitter     : @tostka / http://twitter.com/tostka
    CreatedDate : 2021-11-05
    FileName    : get-EXOMsgTraceDetailed.ps1
    License     : MIT License
    Copyright   : (c) 2021 Todd Kadrie
    Github      : https://github.com/tostka/verb-mod
    Tags        : Powershell,Exchange,ExchangeOnline,Tracking,Delivery
    AddedCredit : REFERENCE
    AddedWebsite: URL
    AddedTwitter: URL
    REVISIONS
    * 4:38 PM 1/24/2022 retooled start/enddateto convert 'local tz' inputs, to t 
        GMT/UTC; and track/convert content back to local time ; added testing of 
        msgtrace splat params, only when populated ;  updated CBH & Examples (covering 
        variant formats of booking response msgs); converted hash summary output to 
        psobject ; ren -doMTD -> -Detailed ; validate $days is a positive integer;
        swapped strings with new central constants: $sFulltimeStamp, $sFiletimestamp
    * 4:04 PM 11/19/2021 flipped wh,wv,ww to wlt - added -days ; updated logic testing for dates/days against MS 10d limit (stored as new constant) ; checks out functional; needs 7pswlt rplcments of write-*
    * 12:40 PM 11/15/2021 - expanded subject -match/-like to post test and use the opposing option where the detected failed to yield filtered msgs. 
    * 3:46 pm 11/12/2021 - added -Subject test-IsRegexPattern() and autoflip tween -match & -like post filtering. 
    * 2:37 PM 11/5/2021 init
    .DESCRIPTION
    get-EXOMsgTraceDetailed.ps1 - Run a MessageTrace with output summarizing, export to csv, and optional followup with MessageTraceDetail, summarize (expand TransportRules opt), and export to csv
    
    This function wraps the EXO get-MessageTrace & get-MessageTraceDetail, to run structured message traces with export to CSV, optional follow-on Get-MessageTraceDetail, post-filtering on specified Subject, and outputs a summary hashtable object with the following:

    Returns summary object to pipeline, with following properties:
    [obj].MTMessages: MessageTracking messages matched
    [obj].MTMessagesCSVFile full path to exported MTMessages as csv file
    [obj].MTDetails: MessageTrackingDetail refactored summary of MTD as transactions
    [obj].MTDCSVFile: full path to exported MTDs as csv file 
    [obj].MTDReport: expanded Detail summary output
    [obj].MTDRptCSVFile: full path to exported MTDReport as csv file 
    > Note: As of 4/2021, MS wrecked utility of get-MessageTrace, dropping range from 30 days to 10 days, with silent failure to return -gt 10d (not even a range error). 
    > So there's not a lot of utility to supporting -Enddate (date) -Days 11, to pull historical 11day windows: If it's more than 10d old, you've got to use HistSearch regardless. 

    .PARAMETER ticket
    Ticket [-ticket 999999]
    .PARAMETER SenderAddress
    SenderAddress[-SenderAddress addr@domain.com]
    .PARAMETER RecipientAddress
    RecipientAddress [-RecipientAddress addr@domain.com]
    .PARAMETER StartDate
    Start of range to be searched[-StartDate '11/5/2021 2:16 PM
    .PARAMETER EndDate
    End of range to be searched (defaults to current time if unspecified)[-EndDate '11/5/2021 5:16 PM']
    .PARAMETER subject
    Subject of target message [-Subject 'Some subject']
    .PARAMETER MessageId
    MessageId of target message [-MessageId '[messageid string]']
    .PARAMETER MessageTraceId
    MessageTraceId of target message [-MessageTraceId '[MessageTraceId string]']
    .PARAMETER MessageTraceDetailLimit
    Integer number of maximum messages to be follow-up MessageTraceDetail'd [-MessageTraceDetailLimit 20]
    .PARAMETER doMTDReportRuleHits
    switch to do Summarize & Expansion of any MTD TransportRule events (defaults true) [-doMTDReportRuleHits]
    .PARAMETER Detailed
    switch to perform MessageTrackingDetail pass, after intial MessageTrace (up to limti specified in -MessageTraceDetailLimit (defaults true) [-Detailed]
    .INPUTS
    None. Does not accepted piped input.(.NET types, can add description)
    .OUTPUTS
    None. Returns no objects or output (.NET types)
    .EXAMPLE
    PS> $results = get-EXOMsgTraceDetailed.ps1 -ticket 651268 -SenderAddress='SENDER@DOMAIN.COM' -RecipientAddress='RECIPIENT@DOMAIN.COM' -StartDate='11/1/2021  4:35:39 PM' -Subject 'Accepted: Exmark/RLC Bring Up' -verbose ;
    # dump messages table and group Status
    $results.MTMessages | ft -a ReceivedLocal,Sender*,Recipient*,subject,*status,*ip ;
    $results.MTMessages | group status | ft -auto count,name ;
    # dump MessageTraceDetail table
    $results.MTDetails | sort Date | ft -a date,event,action,detail,sender*,recipient* ;
    # dump MessageTraceDetail Summary Report as table
    $results.MTDReport| sort date | ft -a DateLocal,Event,Action,Detail ;
    # echo csv output files
    $results.MTMessagesCSVFile ;
    $results.MTDRptCSVFile ;
    Run a typical MessageTrace on sender & recipient, specified start/end dates, and subject, with default 100-message MessageTraceDetail report, with verbose output.
    And then demo of working with the data returned
    .EXAMPLE
    PS> $results = get-EXOMsgTraceDetailed.ps1 -ticket 651268 -SenderAddress='ATTENDEE@DOMAIN.COM' -RecipientAddress='ORGANIZER@DOMAIN.COM' -StartDate='11/1/2021  4:35:39 PM' -Subject 'Accepted: [MEETINGSUBJ]' -verbose ;
    Run a Meeting ACCEPTED MessageTrace - 
        no booking conflict, 
        From: Attendee To: Originator
        Subject: 'Accepted: [MEETINGSUBJ]'
    - with default 100-message MessageTraceDetail report, with verbose output.
    .EXAMPLE
    PS> $results = get-EXOMsgTraceDetailed.ps1 -ticket 651268 -SenderAddress='ROOM@DOMAIN.COM' -RecipientAddress='ORGANIZER@DOMAIN.COM' -StartDate='11/1/2021  4:35:39 PM' -Subject 'Declined: [MEETINGSUBJ]' -verbose ;
    Run a Meeting DECLINED MessageTrace - 
         Booking conflict, 
         From: Room, To: Originator (and copy to any SendOnBehalf delegate that actually created the meeting)
         Subject is: 'Declined: [MEETINGSUBJ]'
    - with default 100-message MessageTraceDetail report, with verbose output.
    .EXAMPLE
    PS> $results = get-EXOMsgTraceDetailed.ps1 -ticket 651268 -SenderAddress='ROOM@DOMAIN.COM' -RecipientAddress='ORGANIZER@DOMAIN.COM' -StartDate='11/1/2021  4:35:39 PM' -Subject 'Tentative: [MEETINGSUBJ]' -verbose ;
    Run a Meeting TENTATIVE response (Moderated resource), MessageTrace, - 
        reflects a AllRequestinPolicy:`$true resource ;
        w ResourceDelegates; 
        no booking conflict;
        but pending ResDelegate approval
        From: Room, To: Originator (and copy to any SendOnBehalf delegate that actually created the meeting)
        Subject is: 'Tentative: [MEETINGSUBJ]'
     -  with default 100-message MessageTraceDetail report, with verbose output. 
    .EXAMPLE
    PS> $results = get-EXOMsgTraceDetailed.ps1 -ticket 651268 -SenderAddress='ORGANIZER@DOMAIN.COM' -RecipientAddress='RESDELEGATE@DOMAIN.COM' -StartDate='11/1/2021  4:35:39 PM' -Subject 'FW: [MEETINGSUBJ]' -verbose ;
    Run a Meeting 'FW: [MEETINGSUBJ]' MODERATION REQUEST MessageTrace - 
        TO: ResourceDelegates (redirected Forward) FROM: ORGANIZER
        reflects a Resource with: AllRequestinPolicy:`$true; 
        ResourceDelegates configured; 
        no booking conflict, but pending ResDelegate approval 
    - MessageTrace (which will come from Meeting Originator email address), to the ResDelegate addresses, with default 100-message MessageTraceDetail report, with verbose output.
    .LINK
    https://docs.microsoft.com/en-us/powershell/module/exchange/get-messagetrace
    .LINK
    https://docs.microsoft.com/en-us/powershell/module/exchange/get-messagetracedetail
    .LINK
    https://github.com/tostka/verb-exo
    #>
    # VALIDATORS: [ValidateNotNull()][ValidateNotNullOrEmpty()][ValidateLength(24,25)][ValidateLength(5)][ValidatePattern("some\sregex\sexpr")][ValidateSet("US","GB","AU")][ValidateScript({Test-Path $_ -PathType 'Container'})][ValidateScript({Test-Path $_})][ValidateRange(21,65)]#positiveInt:[ValidateRange(0,[int]::MaxValue)]#negativeInt:[ValidateRange([int]::MinValue,0)][ValidateCount(1,3)]
    [CmdletBinding(DefaultParameterSetName='Days')]
    PARAM(
        [Parameter(Mandatory=$True,HelpMessage="Ticket [-ticket 999999]")]
        [ValidateNotNullOrEmpty()]    
        [string]$ticket,
        [Parameter(HelpMessage="SenderAddress[-SenderAddress addr@domain.com]")]
        [string]$SenderAddress,
        [Parameter(HelpMessage="RecipientAddress [-RecipientAddress addr@domain.com]")]
        [string]$RecipientAddress,
        [Parameter(ParameterSetName='Dates',HelpMessage="Start of range to be searched[-StartDate '11/5/2021 2:16 PM']")]
        [string]$StartDate,
        [Parameter(ParameterSetName='Dates',HelpMessage="End of range to be searched (defaults to current time if unspecified)[-EndDate '11/5/2021 5:16 PM']")]
        [string]$EndDate=(get-date),
        [Parameter(ParameterSetName='Days',HelpMessage="Days to be searched, back from current time(Alt to use of StartDate & EndDate)[-Days 7]")]
        [ValidateRange(0,[int]::MaxValue)]
        [int]$Days,
        [Parameter(HelpMessage="Subject of target message [-Subject 'Some subject']")]
        [string]$subject,
        [Parameter(HelpMessage="MessageId of target message [-MessageId '[messageid string]']")]
        [string]$MessageId,
        [Parameter(HelpMessage="MessageTraceId of target message [-MessageTraceId '[MessageTraceId string]']")] 
        [string]$MessageTraceId,
        [Parameter(HelpMessage="Integer number of maximum messages to be follow-up MessageTraceDetail'd [-MessageTraceDetailLimit 20]")]
        [int]$MessageTraceDetailLimit = 100,
        [Parameter(HelpMessage="switch to do Summarize & Expansion of any MTD TransportRule events (defaults true) [-DetailedReportRuleHits]")]
        [switch]$DetailedReportRuleHits= $true,
        [Parameter(HelpMessage="switch to perform MessageTrackingDetail pass, after intial MessageTrace (up to limti specified in -MessageTraceDetailLimit (defaults true) [-Detailed]")]
        [switch]$Detailed=$true
    ) ;
    <# #-=-=-=MUTUALLY EXCLUSIVE PARAMS OPTIONS:-=-=-=-=-=
# designate a default paramset, up in cmdletbinding line
[CmdletBinding(DefaultParameterSetName='SETNAME')]
  # * set blank, if none of the sets are to be forced (eg optional mut-excl params)
  # * force exclusion by setting ParameterSetName to a diff value per exclusive param

# example:single $Computername param with *multiple* ParameterSetName's, and varying Mandatory status per set
    [Parameter(ParameterSetName='LocalOnly', Mandatory=$false)]
    $LocalAction,
    [Parameter(ParameterSetName='Credential', Mandatory=$true)]
    [Parameter(ParameterSetName='NonCredential', Mandatory=$false)]
    $ComputerName,
    # $Credential as tied exclusive parameter
    [Parameter(ParameterSetName='Credential', Mandatory=$false)]
    $Credential ;    
    # effect: 
    -computername is mandetory when credential is in use
    -when $localAction param (w localOnly set) is in use, neither $Computername or $Credential is permitted
    write-verbose -verbose:$verbose "ParameterSetName:$($PSCmdlet.ParameterSetName)"
    Can also steer processing around which ParameterSetName is in force:
    if ($PSCmdlet.ParameterSetName -eq 'LocalOnly') {
        return "some localonly stuff" ; 
    } ;     
#-=-=-=-=-=-=-=-=
#>
    BEGIN{
        # get-EXOMsgTraceDetailed.ps1 -ticket 651268 -SenderAddress='SENDER@exmark.com' -RecipientAddress='RECIPIENT@domain.com' -StartDate='11/1/2021  4:35:39 PM' -Subject 'Accepted: [MTGSUBJECT]';
        <#$ticket = '651268' ;
        $subject = 'Accepted: Exmark/RLC Bring Up' ;
        $MessageId=$null ; 
        $MessageTraceId=$null ; 
        $Detailed=$true ;
        $MessageTraceDetailLimit = 100 ; 
        $DetailedReportRuleHits= $true ;
        #>

        $propsMT = @{N='ReceivedLocal';E={$_.Received.ToLocalTime()}},'SenderAddress','RecipientAddress','Subject','Status','ToIP','FromIP','Size','MessageId','MessageTraceId','Index' ;
        # setup a refactor of Receivedlocal on Received, but return *all* properties
        $propsMTAll = 'RunspaceId','Organization','MessageId','Received', @{N='ReceivedLocal';E={$_.Received.ToLocalTime()}},'SenderAddress','RecipientAddress','Subject','Status','ToIP','FromIP','Size','MessageTraceId','StartDate','EndDate','Index'
        #$propsMTD = 'Date','Event','Action','Detail','Data' ;
        # add a locatltime variant
        $propsMTD = @{N='DateLocal';E={$_.Date.ToLocalTime()}},'Date','Event','Action','Detail','Data' ;

        $propsMsgDump = @{N='ReceivedLocal';E={$_.Received.ToLocalTime()}},'Status','SenderAddress','RecipientAddress','Subject' ;
        $DaysLimit = 10 # reflect the current MS get-messagetrace window limit
        $sFulltimeStamp = 'MM/dd/yyyy-HH:mm:ss.fff' ;
        $sFiletimestamp = 'yyyyMMdd-HHmmtt' ;

        ${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name ;
        # Get parameters this function was invoked with
        $PSParameters = New-Object -TypeName PSObject -Property $PSBoundParameters ;
        $Verbose = ($VerbosePreference -eq 'Continue') ; 
        if ($PSScriptRoot -eq "") {
            if ($psISE) { $ScriptName = $psISE.CurrentFile.FullPath } 
            elseif ($context = $psEditor.GetEditorContext()) {$ScriptName = $context.CurrentFile.Path } 
            elseif ($host.version.major -lt 3) {
                $ScriptName = $MyInvocation.MyCommand.Path ;
                $PSScriptRoot = Split-Path $ScriptName -Parent ;
                $PSCommandPath = $ScriptName ;
            } else {
                if ($MyInvocation.MyCommand.Path) {
                    $ScriptName = $MyInvocation.MyCommand.Path ;
                    $PSScriptRoot = Split-Path $MyInvocation.MyCommand.Path -Parent ;
                } else {throw "UNABLE TO POPULATE SCRIPT PATH, EVEN `$MyInvocation IS BLANK!" } ;
            };
            $ScriptDir = Split-Path -Parent $ScriptName ;
            $ScriptBaseName = split-path -leaf $ScriptName ;
            $ScriptNameNoExt = [system.io.path]::GetFilenameWithoutExtension($ScriptName) ;
        } else {
            $ScriptDir = $PSScriptRoot ;
            if ($PSCommandPath) {$ScriptName = $PSCommandPath } 
            else {
                $ScriptName = $myInvocation.ScriptName
                $PSCommandPath = $ScriptName ;
            } ;
            $ScriptBaseName = (Split-Path -Leaf ((& { $myInvocation }).ScriptName))  ;
            $ScriptNameNoExt = [system.io.path]::GetFilenameWithoutExtension($MyInvocation.InvocationName) ;
        } ;
        if ($showDebug) { 
            $smsg = "`$ScriptDir:$($ScriptDir)`n`$ScriptBaseName:$($ScriptBaseName)`n`$ScriptNameNoExt:$($ScriptNameNoExt)`n`$PSScriptRoot:$($PSScriptRoot)`n`$PSCommandPath:$($PSCommandPath)" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
        } ;
        $ComputerName = $env:COMPUTERNAME ;
        $NoProf = [bool]([Environment]::GetCommandLineArgs() -like '-noprofile'); # if($NoProf){# do this};
        # silently stop any running transcripts
        $stopResults = try {Stop-transcript -ErrorAction stop} catch {} ; 

        # #*------v STANDARD START-LOG BP v------
       if(!(get-variable LogPathDrives -ea 0)){$LogPathDrives = 'd','c' };
        foreach($budrv in $LogPathDrives){if(test-path -path "$($budrv):\scripts" -ea 0 ){break} } ;
        if(!(get-variable rgxPSAllUsersScope -ea 0)){
            $rgxPSAllUsersScope="^$([regex]::escape([environment]::getfolderpath('ProgramFiles')))\\((Windows)*)PowerShell\\(Scripts|Modules)\\
        .*\.(ps(((d|m))*)1|dll)$" ;
        } ;
        if(!(get-variable rgxPSCurrUserScope -ea 0)){
            $rgxPSCurrUserScope="^$([regex]::escape([Environment]::GetFolderPath('MyDocuments')))\\((Windows)*)PowerShell\\(Scripts|Modules)\\.
        *\.(ps((d|m)*)1|dll)$" ;
        } ;
        $pltSL=[ordered]@{Path=$null ;NoTimeStamp=$false ;Tag=$null ;showdebug=$($showdebug) ; Verbose=$($VerbosePreference -eq 'Continue') ;
        whatif=$($whatif) ;} ;
        $pltSL.Tag = $Ticket ;
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
            if(($MyInvocation.MyCommand.Definition -match $rgxPSAllUsersScope) -OR ($MyInvocation.MyCommand.Definition -match
        $rgxPSCurrUserScope) ){
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
                # capture the start or it dumps into pipe
                $startResults = start-Transcript -path $transcript ;
            } else {throw "Unable to configure logging!" } ;
        } CATCH {
            $ErrTrapd=$Error[0] ;
            $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details:
        $($ErrTrapd)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        } ;
        #*------^ END STANDARD START-LOG BP ^------

        if ($PSCmdlet.ParameterSetName -eq 'Dates') {
            if($EndDate -and -not $StartDate){
                $StartDate = (get-date $EndDate).addDays(-1 * $DaysLimit) ; 
            } ; 
            
        } else {
            if (-not $Days) {
                $StartDate = (get-date $EndDate).addDays(-1 * $DaysLimit) ; 
                $smsg = "No Days, StartDate or EndDate specified. Defaulting to $($DaysLimit)day Search window:$((get-date).adddays(-1 * $DaysLimit))" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } else {
                $EndDate = (get-date) ;
                $StartDate = (get-date $EndDate).addDays(-1 * $Days) ; 
                $smsg = "-Days:$($Days) specified: "
                #$smsg += "calculated StartDate:$((get-date $StartDate -format $sFulltimeStamp ))" ; 
                #$smsg += ", calculated EndDate:$((get-date $EndDate -format $sFulltimeStamp ))" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                #(get-date -format $sFiletimestamp);
            } ; 
        } ;

        $smsg = "(converting `$StartDate & `$EndDate to UTC, using input as `$StartLocal & `$EndLocal)" ; 
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 

        # convert dates to GMT .ToUniversalTime(
        $StartDate = ([datetime]$StartDate).ToUniversalTime() ; 
        $EndDate = ([datetime]$EndDate).ToUniversalTime() ; 
        $StartLocal = ([datetime]$StartDate).ToLocalTime() ; 
        $EndLocal = ([datetime]$EndDate).ToLocalTime() ; 
        
        # sanity test the start/end dates, just in case (won't throw an error in gxmt)
        if($StartDate -gt $EndDate){
            $smsg = "`-StartDate:$($StartDate) is GREATER THAN -EndDate:($EndDate)!" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
            throw $smsg ; 
            break ; 
        } ; 

        $smsg = "`$StartDate:$(get-date -Date $StartLocal -format $sFulltimeStamp )" ;
        $smsg += "`n`$EndDate:$(get-date -Date $EndLocal -format $sFulltimeStamp )" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 

        if((New-TimeSpan -Start $StartDate -End (get-date)).days -gt $DaysLimit){
            $smsg = "Search span (between -StartDate & -EndDate, or- Days in use) *exceeds* MS supported days history limit!`nReduce the window below a historical 10d, or use get-HistoricalSearch instead!" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug 
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
            Break ; 
        } ; 

    }  # BEG-E
    PROCESS {

        # default StartDate to -10 can't do more
        $pltMsgT=[ordered]@{
          #SenderAddress=$SenderAddress;
          #RecipientAddress=$RecipientAddress;
          #StartDate=(get-date $StartDate);
          #StartDate= $StartDate;
          #EndDate=(get-date $EndDate);
          #EndDate=$EndDate;
          Page= 1 ; # default it to 1 vs $null as we'll be purging empties further down
          ErrorAction = 'STOP' ;
          verbose = $($VerbosePreference -eq "Continue") ;
        } ;
        # throwing errors using unpopulated, so add them conditionally 

        $ofile ="$($ticket)-MsgTrc" ;
        if($SenderAddress){
            $pltMsgT.add('SenderAddress',$SenderAddress) ;
            $ofile+=",From-$($pltMsgT.SenderAddress.replace("*","ANY"))" 
        } ;
        if($RecipientAddress){
            $pltMsgT.add('RecipientAddress',$RecipientAddress) ;
            $ofile+=",To-$($pltMsgT.RecipientAddress.replace("*","ANY"))" ;
        } ;
        if($MessageId){
            $pltMsgT.add('MessageId',$MessageId) ; 
            $ofile+=",MsgId-$($pltMsgT.MessageId.replace('<','').replace('>',''))" ;
        } ;
        if($MessageTraceId){
            $pltMsgT.add('MessageTraceId',$MessageTraceId) ; 
            $ofile+=",MsgId-$($pltMsgT.MessageTraceId.replace('<','').replace('>',''))"  ;
        } ;
        if($subject){
            $ofile+=",Subj-$($subject.substring(0,[System.Math]::Min(15,$subject.Length)))..." 
        } ;
        if($StartDate){
            $pltMsgT.add('StartDate',$StartDate) ; 
            $ofile+= "-$(get-date $pltMsgT.StartDate -format $sFiletimestamp)-"  ;
        } ;
        if($EndDate){
            $pltMsgT.add('EndDate',$EndDate) ; 
            $ofile+= "$(get-date $pltMsgT.EndDate -format $sFiletimestamp)" ;
        } ;
        $ofile=[RegEx]::Replace($ofile, "[{0}]" -f ([RegEx]::Escape(-join [System.IO.Path]::GetInvalidFileNameChars())), '') ;
        # use the tested redirected $logfile path
        #$ofile = join-path -path (split-path $logfile) -ChildPath $ofile ; 
        $hReports = [ordered]@{} ; 
        rxo ;
        $error.clear() ;
        TRY {
            # prepurge empty hash value keys:
            #$pltMsgT=$pltMsgT.GetEnumerator()|? value ;
            # remove null keyed objects
            #$pltMsgT | Foreach {$p = $_ ;@($p.GetEnumerator()) | ?{ ($_.Value | Out-String).length -eq 0 } | Foreach-Object {$p.Remove($_.Key)} ;} ;
            # skip it, we're only adding populated items now
            #write-verbose "hashtype:$($pltmsgt.GetType().FullName)" ; 
            # and issue was first untested negative integer -Days; and 2nd GMT window for start/enddate, so the 'local' input needs to be converted to/from gmt to get the targeted content.

            $smsg = "Get-exoMessageTrace  w`n$(($pltMsgT|out-string).trim())" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-host "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
            $Page = 1  ;
            $Msgs=$null ;
            do {
                $smsg = "Collecting - Page $($Page)..."  ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                $pltMsgT.Page=$Page ;
                $PageMsgs = Get-exoMessageTrace @pltMsgT |  ?{$_.SenderAddress -notlike '*micro*' -or $_.SenderAddress -notlike '*root*' }  ;
                $Page++  ;
                $Msgs += @($PageMsgs)  ;
            } until ($PageMsgs -eq $null) ;
            $Msgs=$Msgs| Sort Received ;
            $smsg = "Raw sender/recipient Msgs:$(($Msgs|measure).Count)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-host "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
            if($subject){
                $smsg = "Post-Filtering on Subject:$($subject)" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                # detect whether to filter on -match (regex) or -like (asterisk, or default non-regex)
                if(test-IsRegexPattern -string $subject -verbose:$($VerbosePreference -eq "Continue")){
                    $smsg = "(detected -subject as regex - using -match comparisonn)" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    $MsgsFltrd = $Msgs | ?{$_.Subject -match $subject} ;
                    if(-not $MsgsFltrd){
                        $smsg = "Subject: regex -match comparison *FAILED* to return matches`nretrying Subject filter as -Like..." ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug 
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        $MsgsFltrd = $Msgs | ?{$_.Subject -like $subject} ;
                    } ; 
                } else { 
                    $smsg = "(detected -subject as NON-regex - using -like comparison)" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    $MsgsFltrd = $Msgs | ?{$_.Subject -like $subject} ;
                    if(-not $MsgsFltrd){
                        $smsg = "Subject: -like comparison *FAILED* to return matches`nretrying Subject filter as -match..." ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug 
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        $MsgsFltrd = $Msgs | ?{$_.Subject -match $subject} 
                    } ; 
                } ; 
                $smsg = "Post Subj filter matches:$(($MsgsFltrd|measure).Count)" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                $msgs = $MsgsFltrd ; 
            } ;
            $ofile+= "-r$(get-date -format $sFiletimestamp).csv" ;
            $ofile = (split-path $ofile -leaf) # can't fix a full path, just the leaf name, then re-path
            $ofile=[RegEx]::Replace($ofile, "[{0}]" -f ([RegEx]::Escape(-join [System.IO.Path]::GetInvalidFileNameChars())), '') ;
            #$ofile=".\logs\$($ofile)" ;
            #$ofile=(join-path -path (split-path $logfile) -childpath "logs\$($ofile)") ;
            $ofile=(join-path -path (split-path $logfile) -childpath $ofile) ;
            if($Msgs){
                # reselect with local time variant
                $Msgs = $Msgs | select $propsMTAll ; 
                $smsg = "($(($Msgs|measure).count)msgs | export-csv $($ofile))" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                $Msgs | select $propsMT | export-csv -notype -path $ofile  ;
                $smsg = "(adding `$hReports.MTMessages)" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                $hReports.add('MTMessages',$msgs) ; 

                # add the csvfilename
                $smsg = "(adding `$hReports.MTMessagesCSVFile:$($ofile))" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                $hReports.add('MTMessagesCSVFile',$ofile) ; 

                $smsg = "Status Distrib:" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host -foregroundcolor yellow "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                $smsg = "`n#*------v MOST RECENT MATCH v------" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host -foregroundcolor white "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                $smsg = "$(($msgs[-1]| fl $propsMsgDump |out-string).trim())";
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                $smsg = "`n#*------^ MOST RECENT MATCH ^------" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host -foregroundcolor white "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host -foregroundcolor white "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                $smsg = "`n#*------v Status DISTRIB v------" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host -foregroundcolor blue "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                $smsg = "$(($Msgs | select -expand Status | group | sort count,count -desc | select count,name|out-string).trim())" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                $smsg = "`n#*------^ Status DISTRIB ^------" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host -foregroundcolor blue "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                if(test-path -path $ofile){
                    $smsg = "(log file confirmed)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-host "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                    Resolve-Path -Path $ofile | select -expand Path | out-clipboard ;
                    $smsg = "$($Msgs.count) matches output to:`n'$($ofile)'`n(copied to CB)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-host "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                } else { "MISSING MsgTrc LOG FILE!" } ;
                if($Detailed){
                    if($msgs.count -gt $MessageTraceDetailLimit){
                        $smsg = "$($msgs.count) EXCEEDS `$MessageTraceDetailLimit:$($MessageTraceDetailLimit)!.`nget-MTD'ing only most recent $($MessageTraceDetailLimit) msgs...!"
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug 
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                        $mtdmsgs = $msgs | select -last $MessageTraceDetailLimit ; 
                    } else { $mtdmsgs = $msgs }  ; 
                    $smsg = "`n[$(($msgs|measure).count)msgs]|=>Get-exoMessageTraceDetail:" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-host "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 

                    $mtds = $mtdmsgs | Get-exoMessageTraceDetail ;

                    $mtdRpt = @() ; 
                    if($DetailedReportRuleHits){
                        $TRules = get-exotransportrule  ; 
                        $smsg = "Checking for `$mtds|`?{$_.Event -eq 'Transport rule'}:" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                        else{ write-host "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                    } ; 
                    foreach($mtd in $mtds){
                        $mtdsummary = [ordered]@{
                            Date = $mtd.Date ; 
                            DateLocal = ([datetime]$mtd.Date).ToLocalTime() ; 
                            Event = $mtd.Event ;
                            Action = $mtd.Action ;
                            Detail = $mtd.Detail ;
                            TRuleName = $null ; 
                            TRuleDetails = $null ; 
                        } ; 
                        if($DetailedReportRuleHits){
                            if ($mtd| ?{$_.Event -eq 'Transport rule'}){
                                # $smsg = "`n$(($mtd | fl Date,Event,Action,Detail |out-string).trim())" ; 
                                if($mtd.detail -match "Transport\srule:\s'',\sID:\s\('(\w{8}-\w{4}-\w{4}-\w{4}-\w{12})'\)"){
                                    #$smsg = "$(($trules|?{$_.guid -eq $matches[1]}  | format-list Name,State,Priority|out-string).trim())" ; 
                                    $ruledetail = $trules|?{$_.guid -eq $matches[1]}  | select Name,Guid,State,Priority ; 
                                    $mtdsummary.TRuleName = $ruledetail.Name ; 
                                    $mtdsummary.TRuleDetails = $ruledetail ; 
                                } ; 
                                $smsg = "`n$(($mtdsummary| fl Date,Event,Action,Detail,TRuleName |out-string).trim())" ; 
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                                else{ write-host "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                            } ; 
                        } else {
                            $smsg = "`n$(($mtdsummary| fl Date,Event,Action,Detail|out-string).trim())" ; 
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                            else{ write-host "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                        }  ;
   
                        $mtdRpt += [pscustomobject]$mtdsummary ; 
                    } ; 
                
                    if($mtds){
                        $ofileMTD = $ofile.replace('-MsgTrc','-MTD') ;
                        $smsg = "($(($mtds|measure).count)mtds | export-csv $($ofileMTD))" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                        $mtds | select $propsMTD | export-csv -notype -path $ofileMTD  ;

                        if(test-path -path $ofileMTD){
                            $smsg = "(log file confirmed)" ;
                            $smsg += "`n$($mtds.count) MTD matches output to:`n'$($ofileMTD)" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                            else{ write-host "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 

                        } else { write-warning "MISSING MTD LOG FILE!" } ;
                        $hReports.add('MTDetails',$mtds) ; 

                        <# add the csvfilename
                        $smsg = "(adding `$hReports.MTDCSVFile:$($ofileMTD))" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                        $hReports.add('MTDCSVFile',$MTDCSVFile) ; 
                        #>
                        #$hReports.add('MTDReport',$ofileMTD) ; 
                        # mtdreport
                        $hReports.add('MTDReport', $mtdRpt) ; 
                        $ofileMTDRpt = $ofile.replace('-MsgTrc','-MTDRpt') ;
                        $mtdRpt | export-csv -notype -path $ofileMTDRpt  ;
                        if(test-path -path $ofileMTD){
                            $smsg = "(log file confirmed)" ;
                            $smsg += "`n$($mtdRpt.count) MTDReport matches output to:`n'$($ofileMTDRpt)" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                            else{ write-host "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                            # add the csvfilename
                            $smsg = "(adding `$hReports.MTDRptCSVFile:$($ofileMTDRpt))" ; 
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                            $hReports.add('MTDRptCSVFile',$ofileMTDRpt) ; 

                        } else { 
                            $smsg = "MISSING MTD LOG FILE!" 
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug 
                            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                        } ;
                   } ;
                } ;
            } else {
                $smsg = "NO MATCHES FOUND from::`n$(($pltMsgT|out-string).trim()|out-default)`n(with any relevant ConnectorId postfilter)" 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
            } ;
        } CATCH {
            $ErrTrapd=$Error[0] ;
            $smsg = "$('*'*5)`nFailed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: `n$(($ErrTrapd|out-string).trim())`n$('-'*5)" ;
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

        
    } ;  # PROC-E
    END {
        # convert the hashtable to object for output to pipeline
        $Rpt += New-Object PSObject -Property $hReports ;
        $smsg = "(Returning summary object to pipeline)" ; 
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
        else{ write-host "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
        $Rpt | Write-Output ; 
    } ; 
}

#*------^ get-EXOMsgTraceDetailed.ps1 ^------