#*------v Get-EXOMessageTraceExportedTDO v------
function Get-EXOMessageTraceExportedTDO {
    <#
    .SYNOPSIS
    Get-EXOMessageTraceExportedTDO - Run a MessageTrace with output summarizing Fails, expanding Qurantines, (expand TransportRules opt), and export to csv, with optional followup with Get-xoMessageTraceDetail, 
    .NOTES
    Version     : 2.0.0
    Author      : Todd Kadrie
    Website     : http://www.toddomation.com
    Twitter     : @tostka / http://twitter.com/tostka
    CreatedDate : 2021-11-05
    FileName    : Get-EXOMessageTraceExportedTDO.ps1
    License     : MIT License
    Copyright   : (c) 2024 Todd Kadrie
    Github      : https://github.com/tostka/verb-EXO
    Tags        : Powershell,Exchange,ExchangeOnline,Tracking,Delivery
    AddedCredit : REFERENCE
    AddedWebsite: URL
    AddedTwitter: URL
    REVISIONS
    * 4:39 PM 12/3/2024 add: updated CBH demos; FailReason, to cover other fails with a Detail: Reason:\s string, and echo out some of the Get-xoMessageTraceDetail detail (though it should be stored in the export as well).
    * 1:45 PM 11/27/2024 minor updates, appears functional;  updated Fail echos for OtherAccount block, citing DDG exclusion setting under CA4 of UserMailbox types.
    * 4:20 PM 11/25/2024 updated from get-exomessagetraceexportedtdo(), more silent suppression, integrated dep-less ExOP conn support
        add: constants for rgxFailSecBlock, $rgxFailOOO, $rgxFailRecallSubj, $rgxFailOtherAcctBlock, $FailOtherAcctBlockExemptionGroup, $rgxFailConfRmExtBlock
    * 5:34 PM 11/22/2024 fundmental rework of the output, looping single & multip failcode entries, and outputing summary for types; removed inline processing outputs, in favor of condenced explanations in the $hsFailxxx outputs herestring reports ; also added the recipienttypedetails, aduser.enabled & TermOU status etc to the output on each Fail message exported as MsgsFail;
        added ConfRm block explicit, and expanded Security* transportrule echo's to cite the rule (if can be parsed from Detail); added recipienttypedetails support for shared|room|equpiment mailboxes
    * 2:59 PM 11/21/2024 working, added code to target 'otherfails', non-OOO, non-Recall, expaanded into Get-xoMessageTraceDetail, then id mail loops, and transport rule blocks (OtherAccounts currently); 
        added considerable expansion and profiling to the fails, also added new sub objects to the return object: 'MsgLast','MsgsFail','MsgsFailOther','MsgsFailOtherDetail','MTMessages','MTMessagesCSVFile','StatusHisto'
    * 5:12 PM 11/20/2024 added -DetailedOtherFails, to force Get-xoMessageTraceDetail on any 'Other' fails, only;
    * 12:41 PM 11/20/2024: update: #1352, 1357, fixed typo in senderaddress/recipaddr $ofile construction
    * 12:54 PM 10/24/2024 confirmed still func;  rename to a more variant of the stock Get-xoMessageTrace : get-EXOMsgTraceDetailed -> Get-EXOMessageTraceExportedTDO; alias Get-EXOMessageTraceExported & prior name: get-EXOMsgTraceDetailed
    * 12:54 PM 10/14/2024 added fully enumerated splat demo    
    * 2:00 PM 10/7/2024 
    * 3:22 PM 9/27/2024 substantial retool, to make it a single goto middleware func for msgtracks, appears working; added params to better approx both Get-xoMessageTrace and existing 7psmsgtrkexo BP calls;
        Added aliases for all Get-MessageTrackingLog & new-xoHistoricalSearch equiv params as well (to cut down on confusion, it takes any synonym for the field)
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
    Get-EXOMessageTraceExportedTDO - Run a MessageTrace with output summarizing Fails, expanding Qurantines, (expand TransportRules opt), and export to csv, with optional followup with Get-xoMessageTraceDetail, 

    This function wraps the EXO get-MessageTrace & get-MessageTraceDetail, to run structured message traces with export to CSV, optional follow-on Get-MessageTraceDetail, post-filtering on specified Subject, and outputs a summary hashtable object with the following:

        Returns summary object to pipeline, with following properties:
        
        [obj].MTMessagesCSVFile full path to exported MTMessages as csv file
        [obj].MTMessages: MessageTracking messages matched
        [obj].StatusHisto: Histogram of Status entries for MTMessages array
        [obj].MsgLast: Last Message returned on track
        [obj].MsgsFail: Status:Fail messages returned on track
        [obj].MsgsFailOOO: Status:Fail messages returned on track that are a product of sender OutOfOffice external Sec Pol Blocks
        [obj].MsgsFailRecall: Status:Fail messages returned on track that are a product of sender Recall attempts
        [obj].MsgsFailOther: Status:Fail messages returned on track that are not OutOfOffice SecPol blocks, or Recalls
        [obj].MsgsFailOtherDetail: Get-xoMessageTraceDetail on .MsgsFailOther messages
        [obj].MTDetails: MessageTrackingDetail refactored summary of MTD as transactions
        [obj].MTDReport: expanded Detail summary output
        [obj].MTDCSVFile: full path to exported MTDs as csv file 

        Exports the object to .xml file as well (named for the main $ofile, renamed ext to .xml)

        For MsgsFailOther, that trace to Mail Loops, runs get-xorecipient, get-recipient & get-aduser on problem Recipient and profiles for incomplete offboard issues.

        > Note: As of 4/2021, MS wrecked utility of get-MessageTrace, dropping range from 30 days to 10 days, with silent failure to return -gt 10d (not even a range error). 
        > So there's not a lot of utility to supporting -Enddate (date) -Days 11, to pull historical 11day windows: If it's more than 10d old, you've got to use HistSearch regardless. 

    .PARAMETER ticket
    Ticket [-ticket 999999]
    .PARAMETER Requestor
    Ticket Customer email identifier. [-Requestor 'fname.lname@domain.com']
    .PARAMETER Tag
    Tag string (Variable Name compatible: no spaces A-Za-z0-9_ only) that is used for Variables and export file name construction. [-Tag 'LastDDGSend']
    .PARAMETER SenderAddress
    SenderAddress (an array runs search on each)[-SenderAddress addr@domain.com]
    .PARAMETER RecipientAddress
    RecipientAddress (an array runs search on each)[-RecipientAddress addr@domain.com]
    .PARAMETER StartDate
    Start of range to be searched[-StartDate '11/5/2021 2:16 PM']
    .PARAMETER EndDate
    End of range to be searched (defaults to current time if unspecified)[-EndDate '11/5/2021 5:16 PM']
    .PARAMETER Days
    Days to be searched, back from current time(Alt to use of StartDate & EndDate; Note:MS won't search -gt 10 days)[-Days 7]
    .PARAMETER Subject
    Subject of target message (emulated via post filtering, not supported param of Get-xoMessageTrace) [-Subject 'Some subject']
    .PARAMETER Status
    The Status parameter filters the results by the delivery status of the message (None|GettingStatus|Failed|Pending|Delivered|Expanded|Quarantined|FilteredAsSpam),an array runs search on each). [-Status 'Failed']
    .PARAMETER MessageId
    MessageId of target message(s) (include any <> and enclose in quotes; an array runs search on each)[-MessageId '<nnnn-nn.xxx....outlook.com>']
    .PARAMETER MessageTraceId
    The MessageTraceId parameter can be used with the recipient address to uniquely identify a message trace and obtain more details. A message trace ID is generated for every message that's processed by the system. [-MessageTraceId 'nnnneacn-ccnn-ndnb-annn-nednfncnnnna']
    .PARAMETER FromIP
    The FromIP parameter filters the results by the source IP address. For incoming messages, the value of FromIP is the public IP address of the SMTP email server that sent the message. For outgoing messages from Exchange Online, the value is blank. [-FromIP '123.456.789.012']
    .PARAMETER ToIP
    The ToIP parameter filters the results by the destination IP address. For outgoing messages, the value of ToIP is the public IP address in the resolved MX record for the destination domain. For incoming messages to Exchange Online, the value is blank. [-ToIP '123.456.789.012']
    .PARAMETER SimpleTrack
    switch to just return the net messages on the initial track (no Fail/Quarantine, MTDetail or other post-processing summaries) [-simpletrack]
    .PARAMETER Detailed
    switch to perform MessageTrackingDetail pass, after intial MessageTrace (up to limit specified in -MessageTraceDetailLimit (defaults true) [-Detailed]
    .PARAMETER DetailedOtherFails
    switch to perform MessageTrackingDetail pass, for any 'Other' Fails (up to limit specified in -MessageTraceDetailLimit (defaults true) [-DetailedOtherFails]
    .PARAMETER DetailedReportRuleHits
    switch to do Summarize & Expansion of any MTD TransportRule events (defaults true) [-DetailedReportRuleHits]
    .PARAMETER MessageTraceDetailLimit
    Integer number of maximum messages to be follow-up MessageTraceDetail'd [-MessageTraceDetailLimit 20]
    .PARAMETER NoQuarCheck
    switch to do Summarize & Expansion of any MTD TransportRule events (defaults true) [-NoQuarCheck]
    .PARAMETER DoExports
    switch to perform configured csv exports of results (defaults true) [-DoExports]
    .PARAMETER TenOrg
    Tenant Tag (3-letter abbrebiation)[-TenOrg 'XYZ']
    .PARAMETER Credential
    Use specific Credentials (defaults to Tenant-defined SvcAccount)[-Credentials [credential object]]
    .PARAMETER UserRole
    Credential User Role spec (SID|CSID|UID|B2BI|CSVC|ESVC|LSVC|ESvcCBA|CSvcCBA|SIDCBA)[-UserRole @('SIDCBA','SID','CSVC')]
    .PARAMETER useEXOv2
    Use EXOv2 (ExchangeOnlineManagement) over basic auth legacy connection [-useEXOv2]
    .PARAMETER Silent
    .INPUTS
    None. Does not accepted piped input.(.NET types, can add description)
    .OUTPUTS
    None. Returns no objects or output (.NET types)
    System.Boolean
    [| get-member the output to see what .NET obj TypeName is returned, to use here]
    .EXAMPLE
        PS> $results = Get-EXOMessageTraceExportedTDO -ticket 651268 -SenderAddress='SENDER@DOMAIN.COM' -RecipientAddress='RECIPIENT@DOMAIN.COM' -StartDate='11/1/2021  4:35:39 PM' -Subject 'Accepted: Exmark/RLC Bring Up' -verbose ;
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
        PS> $results = Get-EXOMessageTraceExportedTDO -ticket 651268 -SenderAddress='ATTENDEE@DOMAIN.COM' -RecipientAddress='ORGANIZER@DOMAIN.COM' -StartDate='11/1/2021  4:35:39 PM' -Subject 'Accepted: [MEETINGSUBJ]' -verbose ;
        Run a Meeting ACCEPTED MessageTrace - 
            no booking conflict, 
            From: Attendee To: Originator
            Subject: 'Accepted: [MEETINGSUBJ]'
        - with default 100-message MessageTraceDetail report, with verbose output.
        .EXAMPLE
        PS> $results = Get-EXOMessageTraceExportedTDO -ticket 651268 -SenderAddress='ROOM@DOMAIN.COM' -RecipientAddress='ORGANIZER@DOMAIN.COM' -StartDate='11/1/2021  4:35:39 PM' -Subject 'Declined: [MEETINGSUBJ]' -verbose ;
        Run a Meeting DECLINED MessageTrace - 
             Booking conflict, 
             From: Room, To: Originator (and copy to any SendOnBehalf delegate that actually created the meeting)
             Subject is: 'Declined: [MEETINGSUBJ]'
        - with default 100-message MessageTraceDetail report, with verbose output.
        .EXAMPLE
        PS> $results = Get-EXOMessageTraceExportedTDO -ticket 651268 -SenderAddress='ROOM@DOMAIN.COM' -RecipientAddress='ORGANIZER@DOMAIN.COM' -StartDate='11/1/2021  4:35:39 PM' -Subject 'Tentative: [MEETINGSUBJ]' -verbose ;
        Run a Meeting TENTATIVE response (Moderated resource), MessageTrace, - 
            reflects a AllRequestinPolicy:`$true resource ;
            w ResourceDelegates; 
            no booking conflict;
            but pending ResDelegate approval
            From: Room, To: Originator (and copy to any SendOnBehalf delegate that actually created the meeting)
            Subject is: 'Tentative: [MEETINGSUBJ]'
         -  with default 100-message MessageTraceDetail report, with verbose output. 
        .EXAMPLE
        PS> $results = Get-EXOMessageTraceExportedTDO -ticket 651268 -SenderAddress='ORGANIZER@DOMAIN.COM' -RecipientAddress='RESDELEGATE@DOMAIN.COM' -StartDate='11/1/2021  4:35:39 PM' -Subject 'FW: [MEETINGSUBJ]' -verbose ;
        Run a Meeting 'FW: [MEETINGSUBJ]' MODERATION REQUEST MessageTrace - 
            TO: ResourceDelegates (redirected Forward) FROM: ORGANIZER
            reflects a Resource with: AllRequestinPolicy:`$true; 
            ResourceDelegates configured; 
            no booking conflict, but pending ResDelegate approval 
        - MessageTrace (which will come from Meeting Originator email address), to the ResDelegate addresses, with default 100-message MessageTraceDetail report, with verbose output.
        .EXAMPLE
        PS> $pltGxMT=[ordered]@{
        PS>    Ticket = '999999' ; 
        PS>    Requestor = 'fname.lname@domain.tld' ; 
        PS>    Tag = 'TestGxmtD' ;
        PS>    RecipientAddress  = 'fname.lname@domain.tld','fname.lname@domain2.TLD' ;
        PS>    senderaddress = 'fname.lname@domain.tld','fname.lname@domain2.TLD' ;
        PS>    StartDate = (get-date ).AddDays(-1) ;
        PS>    EndDate = (get-date ) ;
        PS>    Subject="" ;
        PS>    Status='' ;
        PS>    MessageTraceId='' ;
        PS>    MessageId='' ;
        PS>    FromIP='' ;
        PS>    NoQuarCheck='';
        PS>    Tag='' ;
        PS>     verbose = $true ; 
        PS> } ;
        PS> $pltGxMT = [ordered]@{} ;
        PS> $pltI.GetEnumerator() | ?{ $_.value}  | ForEach-Object { $pltGxMT.Add($_.Key, $_.Value) } ;
        PS> $vn = (@("xoMsgs$($pltI.ticket)",$pltI.Tag) | ?{$_}) -join '_' ;
        PS> write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):Get-EXOMessageTraceExportedTDO w`n$(($pltGxMT|out-string).trim())`n(assign to `$$($vn))" ;
        PS> if(gv $vn -ea 0){rv $vn} ;
        PS> if($tmsgs = Get-EXOMessageTraceExportedTDO @pltGxMT){sv -na $vn -va $tmsgs ;
        PS> write-host "(assigned to `$$vn)"} ;
        Splatted demo
        .EXAMPLE
        PS> $pltGxMT=[ordered]@{
        PS>     Ticket = '99999' ;
        PS>     Requestor = 'fname.lname@domain.tld' ; 
        PS>     Tag = 'AnyTraffic' ;
        PS>     senderaddress = '*@DOMAIN.COM' ;
        PS>     StartDate = (get-date ).AddDays(-10) ;
        PS>     EndDate = (get-date ) ;
        PS>     erroraction = 'STOP' ;
        PS>     verbose = $true ;
        PS> } ;
        PS> $pltGxMT = [ordered]@{} ;
        PS> $pltI.GetEnumerator() | ?{ $_.value}  | ForEach-Object { $pltGxMT.Add($_.Key, $_.Value) } ;
        PS> $vn = (@("xoMsgs$($pltI.ticket)",$pltI.Tag) | ?{$_}) -join '_' ;
        PS> write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):Get-EXOMessageTraceExportedTDO w`n$(($pltGxMT|out-string).trim())`n(assign to `$$($vn))" ;
        PS> if(gv $vn -ea 0){rv $vn} ;
        PS> if($tmsgs = Get-EXOMessageTraceExportedTDO @pltGxMT){sv -na $vn -va $tmsgs ;
        PS> write-host "(assigned to `$$vn)"} ;
        Demo search on wildcard sender address (using * wildcard character)
        .EXAMPLE
        PS> $pltGxMT=[ordered]@{
        PS>     Ticket = '999999' ; 
        PS>     Requestor = 'fname.lname@domain.tld' ; 
        PS>     Tag = 'SEARCHTAG' ;
        PS>     senderaddress = 'fname.lname@domain.tld','fname.lname@domain2.TLD' ;
        PS>     StartDate = (get-date ).AddDays(-1) ;
        PS>     EndDate = (get-date ) ;
        PS>     RecipientAddress = 'fname.lname@domain.tld'  ; 
        PS>     Days = 10 ; 
        PS>     subject = 'MSGSUBJECT' ; 
        PS>     Status = $null ; # 'None|GettingStatus|Failed|Pending|Delivered|Expanded|Quarantined|FilteredAsSpam'
        PS>     MessageId = '<NNNN.NA.NNNNNNNNNNNN@SUB.DOMAIN.TLD>'; 
        PS>     MessageTraceId = 'nnnnccdn-nnen-nfnn-nnan-nnnendnebnce' ; 
        PS>     FromIP = $null ; 
        PS>     ToIP = $null ; 
        PS>     SimpleTrack = $false ;
        PS>     DetailedReportRuleHits= $true ; 
        PS>     NoQuarCheck = $false ; 
        PS>     DoExports=$TRUE ; 
        PS>     Detailed = $false ; 
        PS>     #TenOrg = global:o365_TenOrgDefault ; 
        PS>     #Credential = $null ;
        PS>     #UserRole = @('SIDCBA','SID','CSVC') ; 
        PS>     #useEXOv2 = $true
        PS>     #silent = $false ;
        PS>     verbose = $true ; 
        PS>     Tag='' ;
        PS> } ;
        PS> $pltGxMT = [ordered]@{} ;
        PS> $pltI.GetEnumerator() | ?{ $_.value}  | ForEach-Object { $pltGxMT.Add($_.Key, $_.Value) } ;
        PS> $vn = (@("xoMsgs$($pltI.ticket)",$pltI.Tag) | ?{$_}) -join '_' ;
        PS> write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):Get-EXOMessageTraceExportedTDO w`n$(($pltGxMT|out-string).trim())`n(assign to `$$($vn))" ;
        PS> if(gv $vn -ea 0){rv $vn} ;
        PS> if($tmsgs = Get-EXOMessageTraceExportedTDO @pltGxMT){sv -na $vn -va $tmsgs ;
        PS> write-host "(assigned to `$$vn)"} ;
        Fully eunmerated splat parameters demo, with constructed variable output (uses $pltI.ticket & $pltI.tag)
        .LINK
        https://docs.microsoft.com/en-us/powershell/module/exchange/get-messagetrace
        .LINK
        https://docs.microsoft.com/en-us/powershell/module/exchange/get-messagetracedetail
        .LINK
        https://github.com/tostka/verb-exo
    #>
    #Requires -Modules AzureAD, ExchangeOnlineManagement 
    # VALIDATORS: [ValidateNotNull()][ValidateNotNullOrEmpty()][ValidateLength(24,25)][ValidateLength(5)][ValidatePattern("some\sregex\sexpr")][ValidateSet("US","GB","AU")][ValidateScript({Test-Path $_ -PathType 'Container'})][ValidateScript({Test-Path $_})][ValidateRange(21,65)]#positiveInt:[ValidateRange(0,[int]::MaxValue)]#negativeInt:[ValidateRange([int]::MinValue,0)][ValidateCount(1,3)]
    [CmdletBinding(DefaultParameterSetName='Days')]
    [Alias('get-EXOMsgTraceDetailed','Get-EXOMessageTraceExported')]
    PARAM(
        [Parameter(Mandatory=$false,HelpMessage="Ticket [-ticket 999999]")]
            [ValidateNotNullOrEmpty()]    
            [string]$ticket,
         [Parameter(HelpMessage="Ticket Customer email identifier. [-Requestor 'fname.lname@domain.com']")] 
            [Alias('UID')]
            [string]$Requestor,
        [Parameter(HelpMessage="Tag string (Variable Name compatible: no spaces A-Za-z0-9_ only) that is used for Variables and export file name construction. [-Tag 'LastDDGSend']")] 
            [ValidatePattern('^[A-Za-z0-9_]*$')]
            [string]$Tag,
        [Parameter(HelpMessage="SenderAddress (an array runs search on each)[-SenderAddress addr@domain.com]")]
            [Alias('Sender')]
            [string[]]$SenderAddress, # MultiValuedProperty
        [Parameter(HelpMessage="RecipientAddress (an array runs search on each)[-RecipientAddress addr@domain.com]")]
            [Alias('Recipients')]
            [string[]]$RecipientAddress, # MultiValuedProperty
        [Parameter(ParameterSetName='Dates',HelpMessage="Start of range to be searched[-StartDate '11/5/2021 2:16 PM']")]
            [Alias('Start')]
            [DateTime]$StartDate,
        [Parameter(ParameterSetName='Dates',HelpMessage="End of range to be searched (defaults to current time if unspecified)[-EndDate '11/5/2021 5:16 PM']")]
            [Alias('End')]
            [DateTime]$EndDate=(get-date),
        [Parameter(ParameterSetName='Days',HelpMessage="Days to be searched, back from current time(Alt to use of StartDate & EndDate; Note:MS won't search -gt 10 days)[-Days 7]")]
            #[ValidateRange(0,[int]::MaxValue)]
            [ValidateRange(0,10)] # MS won't search beyond 10, and silently returns incomplete results
            [int]$Days,
        [Parameter(HelpMessage="Subject of target message (emulated via post filtering, not supported param of Get-xoMessageTrace) [-Subject 'Some subject']")]
            [Alias('MessageSubject')]
            [string]$subject,
        [Parameter(HelpMessage="The Status parameter filters the results by the delivery status of the message (None|GettingStatus|Failed|Pending|Delivered|Expanded|Quarantined|FilteredAsSpam),an array runs search on each). [-Status 'Failed']")]
            [Alias('DeliveryStatus','EventId')]
            [ValidateSet('None','GettingStatus','Failed','Pending','Delivered','Expanded','Quarantined','FilteredAsSpam')]
            [string[]]$Status, # MultiValuedProperty
        [Parameter(HelpMessage="MessageId of target message(s) (include any <> and enclose in quotes; an array runs search on each)[-MessageId '<nnnn-nn.xxx....outlook.com>']")]
            # Get-xoMessageTrace specs <MultiValuedProperty>: "just means that you can provide multiple values (i.e. an array) as the argument to the parameter. If your users input something like alice@example.com,bob@example.com,charlie@example.com, you need to split the delims"
            [string[]]$MessageId, # MultiValuedProperty
        [Parameter(HelpMessage="The MessageTraceId parameter can be used with the recipient address to uniquely identify a message trace and obtain more details. A message trace ID is generated for every message that's processed by the system. [-MessageTraceId 'nnnneacn-ccnn-ndnb-annn-nednfncnnnna']")] 
            [Guid]$MessageTraceId,
        [Parameter(HelpMessage="The FromIP parameter filters the results by the source IP address. For incoming messages, the value of FromIP is the public IP address of the SMTP email server that sent the message. For outgoing messages from Exchange Online, the value is blank. [-FromIP '123.456.789.012']")] 
            [string]$FromIP, 
        [Parameter(HelpMessage="The ToIP parameter filters the results by the destination IP address. For outgoing messages, the value of ToIP is the public IP address in the resolved MX record for the destination domain. For incoming messages to Exchange Online, the value is blank. [-ToIP '123.456.789.012']")] 
            [string]$ToIP,
        [Parameter(HelpMessage="switch to just return the net messages on the initial track (no Fail/Quarantine, MTDetail or other post-processing summaries) [-simpletrack]")]
            [switch]$SimpleTrack,
        [Parameter(HelpMessage="switch to do Summarize & Expansion of any MTD TransportRule events (defaults true) [-DetailedReportRuleHits]")]
            [switch]$DetailedReportRuleHits= $true,
        [Parameter(HelpMessage="Integer number of maximum messages to be follow-up MessageTraceDetail'd [-MessageTraceDetailLimit 20]")]
            [int]$MessageTraceDetailLimit = 100,
        [Parameter(HelpMessage="switch to do Summarize & Expansion of any MTD TransportRule events (defaults true) [-NoQuarCheck]")]
            [switch]$NoQuarCheck,
        [Parameter(HelpMessage="switch to perform configured csv exports of results (defaults true) [-DoExports]")]
            [switch]$DoExports=$TRUE,
        [Parameter(HelpMessage="switch to perform Get-xoMessageTraceDetail pass, after intial MessageTrace (up to limit specified in -MessageTraceDetailLimit (defaults false) [-Detailed]")]
            [switch]$Detailed,
        [Parameter(HelpMessage="switch to perform Get-xoMessageTraceDetail pass, for any 'Other' Fails (up to limit specified in -MessageTraceDetailLimit (defaults true) [-DetailedOtherFails]")]
            [switch]$DetailedOtherFails = $true,
        # Service Connection Supporting Varis (AAD, EXO, EXOP)
        [Parameter(Mandatory=$FALSE,HelpMessage="TenantTag value, indicating Tenants to connect to[-TenOrg 'TOL']")]
            [ValidateNotNullOrEmpty()]
            #[ValidatePattern("^\w{3}$")]
            [string]$TenOrg = $global:o365_TenOrgDefault,
        [Parameter(Mandatory = $false, HelpMessage = "Use specific Credentials (defaults to Tenant-defined SvcAccount)[-Credentials [credential object]]")]
            [System.Management.Automation.PSCredential]$Credential,
        [Parameter(Mandatory = $false, HelpMessage = "Credential User Role spec (SID|CSID|UID|B2BI|CSVC|ESVC|LSVC|ESvcCBA|CSvcCBA|SIDCBA)[-UserRole @('SIDCBA','SID','CSVC')]")]
            # sourced from get-admincred():#182: $targetRoles = 'SID', 'CSID', 'ESVC','CSVC','UID','ESvcCBA','CSvcCBA','SIDCBA' ;
            #[ValidateSet("SID","CSID","UID","B2BI","CSVC","ESVC","LSVC","ESvcCBA","CSvcCBA","SIDCBA")]
            # pulling the pattern from global vari w friendly err
            [ValidateScript({
                if(-not $rgxPermittedUserRoles){$rgxPermittedUserRoles = '(SID|CSID|UID|B2BI|CSVC|ESVC|LSVC|ESvcCBA|CSvcCBA|SIDCBA)'} ;
                if(-not ($_ -match $rgxPermittedUserRoles)){throw "'$($_)' doesn't match `$rgxPermittedUserRoles:`n$($rgxPermittedUserRoles.tostring())" ; } ;
                return $true ;
            })]
            [string[]]$UserRole = @('SIDCBA','SID','CSVC'),
            #@('SID','CSVC'),
            # svcAcct use: @('ESvcCBA','CSvcCBA','SIDCBA')
        [Parameter(HelpMessage="Use EXOv2 (ExchangeOnlineManagement) over basic auth legacy connection [-useEXOv2]")]
            [switch] $useEXOv2=$true,
        [Parameter(HelpMessage="Silent output (suppress status echos)[-silent]")]
            [switch] $silent
    ) ;
    BEGIN{
        #region CONSTANTS_AND_ENVIRO #*======v CONSTANTS_AND_ENVIRO v======
        # Debugger:proxy automatic variables that aren't directly accessible when debugging (must be assigned and read back from another vari) ; 
        $rPSCmdlet = $PSCmdlet ; 
        $rPSScriptRoot = $PSScriptRoot ; 
        $rPSCommandPath = $PSCommandPath ; 
        $rMyInvocation = $MyInvocation ; 
        $rPSBoundParameters = $PSBoundParameters ; 
        [array]$score = @() ; 
        if($rPSCmdlet.MyInvocation.InvocationName){
            if($rPSCmdlet.MyInvocation.InvocationName -match '\.ps1$'){
                $score+= 'ExternalScript' 
            }elseif($rPSCmdlet.MyInvocation.InvocationName  -match '^\.'){
                write-warning "dot-sourced invocation detected!:$($rPSCmdlet.MyInvocation.InvocationName)`n(will be unable to leverage script path etc from MyInvocation objects)" ; 
                # dot sourcing is implicit scripot exec
                $score+= 'ExternalScript' ; 
            } else {$score+= 'Function' };
        } ; 
        if($rPSCmdlet.CommandRuntime){
            if($rPSCmdlet.CommandRuntime.tostring() -match '\.ps1$'){$score+= 'ExternalScript' } else {$score+= 'Function' }
        } ; 
        $score+= $rMyInvocation.MyCommand.commandtype.tostring() ; 
        $grpSrc = $score | group-object -NoElement | sort count ;
        if( ($grpSrc |  measure | select -expand count) -gt 1){
            write-warning  "$score mixed results:$(($grpSrc| ft -a count,name | out-string).trim())" ;
            if($grpSrc[-1].count -eq $grpSrc[-2].count){
                write-warning "Deadlocked non-majority results!" ;
            } else {
                $runSource = $grpSrc | select -last 1 | select -expand name ;
            } ;
        } else {
            write-verbose "consistent results" ;
            $runSource = $grpSrc | select -last 1 | select -expand name ;
        };
        write-verbose  "Calculated `$runSource:$($runSource)" ;
        'score','grpSrc' | get-variable | remove-variable ; # cleanup temp varis

        # function self-name (equiv to script's: $MyInvocation.MyCommand.Path) ;
        ${CmdletName} = $rPSCmdlet.MyInvocation.MyCommand.Name ;
        $PSParameters = New-Object -TypeName PSObject -Property $rPSBoundParameters ;
        write-verbose "`$rPSBoundParameters:`n$(($rPSBoundParameters|out-string).trim())" ;
        $Verbose = ($VerbosePreference -eq 'Continue') ; 
        # pre psv2, no $rPSBoundParameters autovari to check, so back them out:
        if($rPSCmdlet.MyInvocation.InvocationName){
            if($rPSCmdlet.MyInvocation.InvocationName  -match '^\.'){
                $smsg = "detected dot-sourced invocation: Skipping `$PSCmdlet.MyInvocation.InvocationName-tied cmds..." ; 
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
            } else { 
                write-verbose 'Collect all non-default Params (works back to psv2 w CmdletBinding)'
                $ParamsNonDefault = (Get-Command $rPSCmdlet.MyInvocation.InvocationName).parameters | Select-Object -expand keys | Where-Object{$_ -notmatch '(Verbose|Debug|ErrorAction|WarningAction|ErrorVariable|WarningVariable|OutVariable|OutBuffer)'} ;
            } ; 
        } else { 
            $smsg = "(blank `$rPSCmdlet.MyInvocation.InvocationName, skipping Parameters collection)" ; 
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        } ; 
        #region ENVIRO_DISCOVER ; #*------v ENVIRO_DISCOVER v------
        <#
        # Debugger:proxy automatic variables that aren't directly accessible when debugging ; 
        $rPSScriptRoot = $PSScriptRoot ; 
        $rPSCommandPath = $PSCommandPath ; 
        $rMyInvocation = $MyInvocation ; 
        $rPSBoundParameters = $PSBoundParameters ; 
        #>
        $ScriptDir = $scriptName = '' ;     
        if($ScriptDir -eq '' -AND ( (get-variable -name rPSScriptRoot -ea 0) -AND (get-variable -name rPSScriptRoot).value.length)){
            $ScriptDir = $rPSScriptRoot
        } ; # populated rPSScriptRoot
        if( (get-variable -name rPSCommandPath -ea 0) -AND (get-variable -name rPSCommandPath).value.length){
            $ScriptName = $rPSCommandPath
        } ; # populated rPSCommandPath
        if($ScriptDir -eq '' -AND $runSource -eq 'ExternalScript'){$ScriptDir = (Split-Path -Path $rMyInvocation.MyCommand.Source -Parent)} # Running from File
        # when $runSource:'Function', $rMyInvocation.MyCommand.Source is empty,but on functions also tends to pre-hit from the rPSCommandPath entFile.FullPath ;
        if( $scriptname -match '\.psm1$' -AND $runSource -eq 'Function'){
            write-host "MODULE-HOMED FUNCTION:Use `$CmdletName to reference the running function name for transcripts etc (under a .psm1 `$ScriptName will reflect the .psm1 file  fullname)"
            if(-not $CmdletName){write-warning "MODULE-HOMED FUNCTION with BLANK `$CmdletNam:$($CmdletNam)" } ;
        } # Running from .psm1 module
        if($ScriptDir -eq '' -AND (Test-Path variable:psEditor)) {
            write-verbose "Running from VSCode|VS" ; 
            $ScriptDir = (Split-Path -Path $psEditor.GetEditorContext().CurrentFile.Path -Parent) ; 
                if($ScriptName -eq ''){$ScriptName = $psEditor.GetEditorContext().CurrentFile.Path }; 
        } ;
        if ($ScriptDir -eq '' -AND $host.version.major -lt 3 -AND $rMyInvocation.MyCommand.Path.length -gt 0){
            $ScriptDir = $rMyInvocation.MyCommand.Path ; 
            write-verbose "(backrev emulating `$rPSScriptRoot, `$rPSCommandPath)"
            $ScriptName = split-path $rMyInvocation.MyCommand.Path -leaf ;
            $rPSScriptRoot = Split-Path $ScriptName -Parent ;
            $rPSCommandPath = $ScriptName ;
        } ;
        if ($ScriptDir -eq '' -AND $rMyInvocation.MyCommand.Path.length){
            if($ScriptName -eq ''){$ScriptName = $rMyInvocation.MyCommand.Path} ;
            $ScriptDir = $rPSScriptRoot = Split-Path $rMyInvocation.MyCommand.Path -Parent ;
        }
        if ($ScriptDir -eq ''){throw "UNABLE TO POPULATE SCRIPT PATH, EVEN `$rMyInvocation IS BLANK!" } ;
        if($ScriptName){
            if(-not $ScriptDir ){$ScriptDir = Split-Path -Parent $ScriptName} ; 
            $ScriptBaseName = split-path -leaf $ScriptName ;
            $ScriptNameNoExt = [system.io.path]::GetFilenameWithoutExtension($ScriptName) ;
        } ; 
        # blank $cmdlet name comming through, patch it for Scripts:
        if(-not $CmdletName -AND $ScriptBaseName){
            $CmdletName = $ScriptBaseName
        }
        # last ditch patch the values in if you've got a $ScriptName
        if($rPSScriptRoot.Length -ne 0){}else{ 
            if($ScriptName){$rPSScriptRoot = Split-Path $ScriptName -Parent }
            else{ throw "Unpopulated, `$rPSScriptRoot, and no populated `$ScriptName from which to emulate the value!" } ; 
        } ; 
        if($rPSCommandPath.Length -ne 0){}else{ 
            if($ScriptName){$rPSCommandPath = $ScriptName }
            else{ throw "Unpopulated, `$rPSCommandPath, and no populated `$ScriptName from which to emulate the value!" } ; 
        } ; 
        if(-not ($ScriptDir -AND $ScriptBaseName -AND $ScriptNameNoExt  -AND $rPSScriptRoot  -AND $rPSCommandPath )){ 
            throw "Invalid Invocation. Blank `$ScriptDir/`$ScriptBaseName/`ScriptNameNoExt" ; 
            BREAK ; 
        } ; 
        # echo results dyn aligned:
        $tv = 'runSource','CmdletName','ScriptName','ScriptBaseName','ScriptNameNoExt','ScriptDir','PSScriptRoot','PSCommandPath','rPSScriptRoot','rPSCommandPath' ; 
        $tvmx = ($tv| Measure-Object -Maximum -Property Length).Maximum * -1 ; 
        if($silent){}else{
            #$tv | get-variable | %{  write-host -fore yellow ("`${0,$tvmx} : {1}" -f $_.name,$_.value) } ; # w-h
            $tv | get-variable | %{  write-verbose ("`${0,$tvmx} : {1}" -f $_.name,$_.value) } ; # w-v
        }
        'tv','tvmx'|get-variable | remove-variable ; # cleanup temp varis        

        #endregion ENVIRO_DISCOVER ; #*------^ END ENVIRO_DISCOVER ^------

        # Get-EXOMessageTraceExportedTDO -ticket 651268 -SenderAddress='SENDER@exmark.com' -RecipientAddress='RECIPIENT@domain.com' -StartDate='11/1/2021  4:35:39 PM' -Subject 'Accepted: [MTGSUBJECT]';
        <#$ticket = '651268' ;
        $subject = 'Accepted: Exmark/RLC Bring Up' ;
        $MessageId=$null ; 
        $MessageTraceId=$null ; 
        $Detailed=$true ;
        $MessageTraceDetailLimit = 100 ; 
        $DetailedReportRuleHits= $true ;
        #>
        if(-not (gcm Remove-InvalidVariableNameChars -ea 0)){
            Function Remove-InvalidVariableNameChars ([string]$Name) {
                ($Name.tochararray() -match '[A-Za-z0-9_]') -join '' | write-output ;
            };
        } ;

        #$prpGXMTfta = 'ReceivedLocal','Status','SenderAddress','RecipientAddress','Subject','MessageId' ;
        #$prpGXQMfta = 'ReceivedTime','Type','Direction','SenderAddress','RecipientAddress','Subject','MessageId','Size','ReleaseStatus','Expires','ReleasedBy' ;
        [regex]$rgxHdrSenderIDKeys = ('(?i:' + (('spf','dkim','dmarc','d=','smtp.mailfrom','smtp.rcpttodomain','header.from=','helo','Return-Path:','From:','Subject:','Sender:','Submitter:','Reply-To:','To:','Message-ID:','client-ip','X-Mailer:','X-Received:','Received: from','ARC-Authentication-Results:','arc=','oda=','compauth=','reason=' |%{[regex]::escape($_)}) -join '|') + ')') ;
        [regex]$rgxReturnPath = "Return-Path:((\n|\r|\s)*)([0-9a-zA-Z]+[-._+&='])*[0-9a-zA-Z]+@([-0-9a-zA-Z]+[.])+[a-zA-Z]{2,63}" ;

        $propsMT = 'Received',@{N='ReceivedLocal';E={[datetime]$_.Received.ToLocalTime()}},'SenderAddress','RecipientAddress','Subject','Status','ToIP','FromIP','Size','MessageId','MessageTraceId','Index' ;
        # setup a refactor of Receivedlocal on Received, but return *all* properties
        $propsMTAll = 'RunspaceId','Organization','MessageId','Received', @{N='ReceivedLocal';E={[datetime]$_.Received.ToLocalTime()}},'SenderAddress','RecipientAddress','Subject','Status','ToIP','FromIP','Size','MessageTraceId','StartDate','EndDate','Index'
        #$propsMTD = 'Date','Event','Action','Detail','Data' ;
        # add a locatltime variant
        $propsMTD = @{N='DateLocal';E={$_.Date.ToLocalTime()}},'Date','Event','Action','Detail','Data' ;

        $propsMsgDump = @{N='ReceivedLocal';E={$_.Received.ToLocalTime()}},'Status','SenderAddress','RecipientAddress','Subject' ;
        $DaysLimit = 10 # reflect the current MS get-messagetrace window limit
        #$sFulltimeStamp = 'MM/dd/yyyy-HH:mm:ss.fff' ;
        #$sFiletimestamp = 'yyyyMMdd-HHmm' ;
        $s24HTimestamp = 'yyyyMMdd-HHmm'
        $sFiletimestamp =  $s24HTimestamp

        # block identifying filters
        $rgxFailOOOSubj = '^Automatic\sreply:\s' ; 
        $rgxFailRecallSubj = '^Recall:\s' ; 
        $rgxFailOtherAcctBlock = 'OtherAccts-External-Mail-Rejection' ; 
        $FailOtherAcctBlockExemptionGroup = 'LYN-DL-OPExch-OtherAcctMbxs-ExternalMailOK@toro.com' ; 
        $rgxFailConfRmExtBlock = 'ConfRm-External-Mail-Rejection' ; 
        $rgxFailSecBlock = '^Security(\s-\s|-)' ; 

        #endregion CONSTANTS_AND_ENVIRO ; #*------^ END CONSTANTS_AND_ENVIRO ^------
        
        #region FUNCTIONS ; #*======v FUNCTIONS v======
        # Pull the CUser mod dir out of psmodpaths:
        #$CUModPath = $env:psmodulepath.split(';')|?{$_ -like '*\Users\*'} ;
    
        # 2b4() 2b4c() & fb4() are located up in the CONSTANTS_AND_ENVIRO\ENCODED_CONTANTS block ( to convert Constant assignement strings)

        #region SWRITELOG ; #*------v SIMPLIFIED WRITE-LOG v------
        if(-not(get-command write-log -ea 0)){
            function write-log  {
                [Alias('write-log')]
                Param (
                    [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)][Alias('LogContent')][ValidateNotNullOrEmpty()][string]$Message,
                    [Parameter(Mandatory = $false)][string]$Path = 'C:\Logs\PowerShellLog.log',
                    [Parameter(Mandatory = $false)][ValidateSet('Error','Warn','Info','H1','H2','H3','Debug','Verbose','Prompt')][string]$Level = "Info",
                    [switch] $useHost
                )  ;
                if($host.Name -eq 'Windows PowerShell ISE Host' -AND $host.version.major -lt 3){
                    write-verbose "(low-contrast/visibility ISE 2 detected: using alt colors)" ;
                    $pltWH = @{foregroundcolor = 'yellow' ; backgroundcolor = 'black'} ;
                    $pltErr=@{foregroundcolor='yellow';backgroundcolor='red'};
                    $pltWarn=@{foregroundcolor='black';backgroundcolor='yellow'};
                    $pltInfo=@{foregroundcolor='green';backgroundcolor='black'};
                    $pltH1=@{foregroundcolor='black';backgroundcolor='darkyellow'};
                    $pltH2=@{foregroundcolor='black';backgroundcolor='gray'};
                    $pltH3=@{foregroundcolor='black';backgroundcolor='darkgray'};
                    $pltDbg=@{foregroundcolor='red';backgroundcolor='black'};
                    $pltVerb=@{foregroundcolor='Gray';backgroundcolor='black'};
                    $pltPrmpt=@{foregroundcolor='Blue';backgroundcolor='White'};
                } else {
                    $pltWH = @{} ;
                    $pltErr=@{foregroundcolor='yellow';backgroundcolor='red'};
                    $pltWarn=@{foregroundcolor='yellow';backgroundcolor='red'};
                    $pltInfo=@{foregroundcolor='yellow';backgroundcolor='red'};
                    $pltH1=@{foregroundcolor='black';backgroundcolor='darkyellow'};
                    $pltH2=@{foregroundcolor='black';backgroundcolor='gray'};
                    $pltH3=@{foregroundcolor='black';backgroundcolor='darkgray'};
                    $pltDbg=@{foregroundcolor='red';backgroundcolor='black'};
                    $pltVerb=@{foregroundcolor='yellow';backgroundcolor='red'};
                    $pltPrmpt=@{foregroundcolor='Blue';backgroundcolor='White'};
                } ; 
                if (-not (Test-Path $Path)) {
                        Write-Verbose "Creating $Path."  ;
                        $NewLogFile = New-Item $Path -Force -ItemType File  ;
                }  ; 
                $FormattedDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"  ;
                $EchoTime = "$((get-date).ToString('HH:mm:ss')): " ;
                switch ($Level) {
                    'Error' {
                        $LevelText = 'ERROR: ' ; $smsg = $EchoTime ;
                        if ($useHost) {
                            $smsg += $LevelText + $Message ;
                            write-host @pltErr $smsg ; 
                        } else {if (-not $NoEcho) { Write-Error ($smsg + $Message) } } ;
                    }
                    'Warn' {
                        $LevelText = 'WARNING: ' ; $smsg = $EchoTime ;
                        if ($useHost) {
                            $smsg += $LevelText + $Message ; 
                            write-host @pltWarn $smsg ; 
                        } else {if (-not $NoEcho) { Write-Warning ($smsg + $Message) } } ;
                    }
                    'Info' {
                        $LevelText = 'INFO: ' ; $smsg = $EchoTime ;
                            $smsg += $LevelText + $Message ; 
                            if (-not $NoEcho) { write-host @pltInfo $smsg ;} ;
                    }
                    'H1' {
                        $LevelText = '# ' ; $smsg = $EchoTime ;
                        $smsg += $LevelText + $Message ;  
                        if (-not $NoEcho) { write-host @pltH1 $smsg ; };             
                    }
                    'H2' {
                        $LevelText = '## ' ; $smsg = $EchoTime ;
                        $smsg += $LevelText + $Message ; 
                        if (-not $NoEcho) { write-host @pltH2 $smsg ;};
                    }
                    'H3' {
                        $LevelText = '### ' ; $smsg = $EchoTime ;
                        $smsg += $LevelText + $Message ; 
                        if (-not $NoEcho) { write-host @pltH3 $smsg };
                    }
                    'Debug' {
                        $LevelText = 'DEBUG: ' ; $smsg = ($EchoTime + $LevelText + '(' + $Message + ')') ;
                        write-host @pltDbg $smsg ;
                        if (-not $NoEcho) { Write-Host $smsg }  ;                
                    }
                    'Verbose' {
                        $LevelText = 'VERBOSE: ' ; $smsg = ($EchoTime + '(' + $Message + ')') ;
                        if ($useHost) {                    
                            $smsg = ($EchoTime + $LevelText + '(' + $Message + ')') ;
                            $smsg += $LevelText + $Message ; 
                            if (-not $NoEcho) {write-host @pltVerb $smsg ;} ; 
                        }else {if (-not $NoEcho) { Write-Verbose ($smsg) } } ;          
                    }
                    'Prompt' {
                        $LevelText = 'PROMPT: ' ; $smsg = $EchoTime ;
                        $smsg += $LevelText + $Message ; 
                        if (-not $NoEcho) { write-host @pltPrmpt $smsg ; } ; 
                    }
                } ;
                "$FormattedDate $LevelText : $Message" | Out-File -FilePath $Path -Append  ;
            } ;
        } ; 
        #endregion SWRITELOG ; #*------^ END SIMPLIFIED write-log  ^------

        #region SSTARTLOG ; #*------v SIMPLIFIED start-log v------
        #*------v Start-Log.ps1 v------
        if(-not(get-command start-log -ea 0)){
            function Start-Log {
                [CmdletBinding()]
                PARAM(
                    [Parameter(Position=0,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,HelpMessage="Path to target script (defaults to `$PSCommandPath) [-Path .\path-to\script.ps1]")]
                    # rem out validation, for module installed in AllUsers etc, we don't want to have to spec a real existing file. No bene to testing
                    #[ValidateScript({Test-Path (split-path $_)})] 
                    $Path,
                    [Parameter(HelpMessage="Tag string to be used with -Path filename spec, to construct log file name [-tag 'ticket-123456]")]
                    [string]$Tag,
                    [Parameter(HelpMessage="Flag that suppresses the trailing timestamp value from the generated filenames[-NoTimestamp]")]
                    [switch] $NoTimeStamp,
                    [Parameter(HelpMessage="Flag that leads the returned filename with the Tag parameter value[-TagFirst]")]
                    [switch] $TagFirst,
                    [Parameter(HelpMessage="Debugging Flag [-showDebug]")]
                    [switch] $showDebug,
                    [Parameter(HelpMessage="Whatif Flag  [-whatIf]")]
                    [switch] $whatIf=$true
                ) ;
                #${CmdletName} = $rPSCmdlet.MyInvocation.MyCommand.Name ;
                #$PSParameters = New-Object -TypeName PSObject -Property $rPSBoundParameters ;
                $Verbose = ($VerbosePreference -eq 'Continue') ; 
                $transcript = join-path -path (Split-Path -parent $Path) -ChildPath "logs" ;
                if (-not (test-path -path $transcript)) { write-host "Creating missing log dir $($transcript)..." ; mkdir $transcript  ; } ;
                #$transcript = join-path -path $transcript -childpath "$([system.io.path]::GetFilenameWithoutExtension($Path))" ; 
                if($Tag){
                    # clean for fso use
                    $Tag = Remove-StringDiacritic -String $Tag ; # verb-text
                    $Tag = Remove-StringLatinCharacters -String $Tag ; # verb-text
                    $Tag = Remove-InvalidFileNameChars -Name $Tag ; # verb-io, (inbound Path is assumed to be filesystem safe)
                    if($TagFirst){
                        $smsg = "(-TagFirst:Building filenames with leading -Tag value)" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                        $transcript = join-path -path $transcript -childpath "$($Tag)-$([system.io.path]::GetFilenameWithoutExtension($Path))" ; 
                        #$transcript = "$($Tag)-$($transcript)" ; 
                    } else { 
                        $transcript = join-path -path $transcript -childpath "$([system.io.path]::GetFilenameWithoutExtension($Path))" ; 
                        $transcript += "-$($Tag)" ; 
                    } ;
                } else {
                    $transcript = join-path -path $transcript -childpath "$([system.io.path]::GetFilenameWithoutExtension($Path))" ; 
                }; 
                $transcript += "-Transcript-BATCH"
                if(-not $NoTimeStamp){ $transcript += "-$(get-date -format 'yyyyMMdd-HHmmtt')" } ; 
                $transcript += "-trans-log.txt"  ;
                # add log file variant as target of Write-Log:
                $logfile = $transcript.replace("-Transcript", "-LOG").replace("-trans-log", "-log") ;
                if(get-variable whatif -ea 0){
                    if ($whatif) {
                        $logfile = $logfile.replace("-BATCH", "-BATCH-WHATIF") ;
                        $transcript = $transcript.replace("-BATCH", "-BATCH-WHATIF") ;
                    } else {
                        $logfile = $logfile.replace("-BATCH", "-BATCH-EXEC") ;
                        $transcript = $transcript.replace("-BATCH", "-BATCH-EXEC") ;
                    } ;
                } ; 
                $logging = $True ;

                $hshRet= [ordered]@{
                    logging=$logging ;
                    logfile=$logfile ;
                    transcript=$transcript ;
                } ;
                if($showdebug -OR $verbose){
                    write-verbose -verbose:$true "$(($hshRet|out-string).trim())" ;  ;
                } ;
                Write-Output $hshRet ;
            }
        } ; 
        #*------^ END Start-Log.ps1 ^------
        #endregion SSTARTLOG ; #*------^ END SIMPLIFIED start-log ^------

        #region CONNEXOPTDO ; #*------v CONNEXOPTDO v------
        #*------v Function Connect-ExchangeServerTDO v------
        if(-not(get-command Connect-ExchangeServerTDO -ea 0)){
            Function Connect-ExchangeServerTDO {
                <#
                .SYNOPSIS
                Connect-ExchangeServerTDO.ps1 - Dependancy-less Function that, fed an Exchange server name, or AD SiteName, and optional RoleNames array, 
                will obtain a list of Exchange servers from AD (in the specified scope), and then run the list attempting to PowershellREmote (REMS) connect to each server, 
                stopping at the first successful connection.
                .NOTES
                Version     : 3.0.3
                Author      : Todd Kadrie
                Website     : http://www.toddomation.com
                Twitter     : @tostka / http://twitter.com/tostka
                CreatedDate : 2024-05-30
                FileName    : Connect-ExchangeServerTDO.ps1
                License     : (none-asserted)
                Copyright   : (none-asserted)
                Github      : https://github.com/tostka/verb-Ex2010
                Tags        : Powershell, ActiveDirectory, Exchange, Discovery
                AddedCredit : Brian Farnsworth
                AddedWebsite: https://codeandkeep.com/
                AddedTwitter: URL
                AddedCredit : David Paulson
                AddedWebsite: https://techcommunity.microsoft.com/t5/exchange-team-blog/exchange-health-checker-has-a-new-home/ba-p/2306671
                AddedTwitter: URL
                REVISIONS
                * 3:54 PM 11/26/2024 integrated back TLS fixes, and ExVersNum flip from June; syncd dbg & vx10 copies.
                * 12:57 PM 6/11/2024 Validated, Ex2010 & Ex2019, hub, mail & edge roles: tested ☑️ on CMW mail role (Curly); and Jumpbox; 
                    copied in CBH from repo copy, which has been updated/debugged compat on CMW Edge 
                    includes local snapin detect & load for edge role (simplest EMS load option for Edge role, from David Paulson's original code; no longer published with Ex2010 compat)
                * 1:30 PM 9/5/2024 added  update-SecurityProtocolTDO() SB to begin
                * 12:49 PM 6/21/2024 flipped PSS Name to Exchange$($ExchVers[dd])
                * 11:28 AM 5/30/2024 fixed failure to recognize existing functional PSSession; Made substantial update in logic, validate works fine with other orgs, and in our local orgs.
                * 4:02 PM 8/28/2023 debuged, updated CBH, renamed connect-ExchangeSErver -> Connect-ExchangeServerTDO (avoid name clashes, pretty common verb-noun combo).
                * 12:36 PM 8/24/2023 init

                .DESCRIPTION
                Connect-ExchangeServerTDO.ps1 - Dependancy-less Function that, fed an Exchange server name, or AD SiteName, and optional RoleNames array, 
                will obtain a list of Exchange servers from AD (in the specified scope), and then run the list attempting to PowershellRemote (REMS) connect to each server, 
                stopping at the first successful connection.

                Relies upon/requires get-ADExchangeServerTDO(), to return a descriptive summary of the Exchange server(s) revision etc, for connectivity logic.
                Supports Exchange 2010 through 2019, as implemented.
        
                Intent, as contrasted with verb-EXOP/Ex2010 is to have no local module dependancies, when running EXOP into other connected orgs, where syncing profile & supporting modules code can be problematic. 
                This uses native ADSI calls, which are supported by Windows itself, without need for external ActiveDirectory module etc.

                The particular approach inspired by BF's demo func that accompanied his take on get-adExchangeServer(), which I hybrided with my own existing code for cred-less connectivity. 
                I added get-OrganizationConfig testing, for connection pre/post confirmation, along with Exchange Server revision code for continutional handling of new-pssession remote powershell EMS connections.
                Also shifted connection code into _connect-EXOP() internal func.
                As this doesn't rely on local module presence, it doesn't have to do the usual local remote/local invocation detection you'd do for non-dehydrated on-server EMS (more consistent this way, anyway; 
                there are only a few cmdlet outputs I'm aware of, that have fundementally broken returns dehydrated, and require local non-remote EMS use to function.

                My core usage would be to paste the function into the BEGIN{} block for a given remote org process, to function as a stricly local ad-hoc function.
                .PARAMETER name
                FQDN of a specific Exchange server[-Name EXSERVER.DOMAIN.COM]
                .PARAMETER discover
                Boolean paraameter that drives auto-discovery of target Exchange servers for connection (defaults `$true)[-discover:`$false]
                .PARAMETER credential
                Use specific Credentials[-Credentials [credential object]
                    .PARAMETER Site
                Name of specific AD site to be searched for ExchangeServers (defaults to global variable `$TenOrg_ADSiteDefaultName if present)[-Site 'SITENAME']
                .PARAMETER RoleNames
                Array of Server 'Role' name strings to be filtered against (MBX|CAS|HUB|UM|MBX|EDGE)[-RoleNames 'HUB','CAS']
                .PARAMETER TenOrg
                Tenant Tag (3-letter abbrebiation - defaults to variable `$global:o365_TenOrgDefault if present)[-TenOrg 'XYZ']
                .INPUTS
                None. Does not accepted piped input.(.NET types, can add description)
                .OUTPUTS
                [system.object] Returns a system object containing a successful PSSession
                System.Boolean
                [| get-member the output to see what .NET obj TypeName is returned, to use here]
                System.Array of System.Object's
                .EXAMPLE
                PS> $PSSession = Connect-ExchangeServerTDO -siteName SITENAME -RoleNames @('HUB','CAS') -verbose 
                Demo's connecting to a functional Hub or CAS server in the SITENAME site with verbose outputs, the `PSSession variable will contain information about the successful connection. Makes automatic Exchangeserver discovery calls into AD (using ADSI) leveraging the separate get-ADExchangeServerTDO()
                .EXAMPLE
                PS> TRY{$Site=[System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::GetComputerSite().Name}CATCH{$Site=$env:COMPUTERNAME} ;
                PS> $PSSession = Connect-ExchangeServerTDO -siteName $Site -RoleNames @('HUB','CAS') -verbose ; 
                Demo including support for EdgeRole, which is detected on it's lack of AD Site specification (which gets fed through to call, by setting the Site to the machine itself).
                .LINK
                https://codeandkeep.com/PowerShell-ActiveDirectory-Exchange-Part1/
                .LINK
                https://github.com/Lucifer1993/PLtools/blob/main/HealthChecker.ps1
                .LINK
                https://microsoft.github.io/CSS-Exchange/Diagnostics/HealthChecker/
                .LINK
                https://bitbucket.org/tostka/powershell/
                .LINK
                https://github.com/tostka/verb-Ex2010
                #>        
                [CmdletBinding(DefaultParameterSetName='discover')]
                PARAM(
                    [Parameter(Position=0,Mandatory=$true,ParameterSetName='name',HelpMessage="FQDN of a specific Exchange server[-Name EXSERVER.DOMAIN.COM]")]
                        [String]$name,
                    [Parameter(Position=0,ParameterSetName='discover',HelpMessage="Boolean paraameter that drives auto-discovery of target Exchange servers for connection (defaults `$true)[-discover:`$false]")]
                        [bool]$discover=$true,
                    [Parameter(Position=1,HelpMessage = "Use specific Credentials[-Credentials [credential object]")]
                        [Management.Automation.PSCredential]$credential,
                    [Parameter(Position=1,HelpMessage="Name of specific AD site to be searched for ExchangeServers (defaults to global variable `$TenOrg_ADSiteDefaultName if present)[-Site 'SITENAME']")]
                        [Alias('Site')]
                        [string]$SiteName = (gv -name "$($TenOrg)_ADSiteDefaultName" -ea 0).Value,
                    [Parameter(Position=2,HelpMessage="Array of Server 'Role' name strings to be filtered against (MBX|CAS|HUB|UM|MBX|EDGE)[-RoleNames 'HUB','CAS']")]
                        [ValidateSet('MBX','CAS','HUB','UM','MBX','EDGE')]
                        [string[]]$RoleNames = @('HUB','CAS'),
                    [Parameter(Mandatory=$FALSE,HelpMessage="Tenant Tag (3-letter abbrebiation - defaults to global:o365_TenOrgDefault if present)[-TenOrg 'XYZ']")]
                        [ValidateNotNullOrEmpty()]
                        [string]$TenOrg = $global:o365_TenOrgDefault
                ) ;
                BEGIN{
                    $Verbose = ($VerbosePreference -eq 'Continue') ;
                    $CurrentVersionTlsLabel = [Net.ServicePointManager]::SecurityProtocol ; # Tls, Tls11, Tls12 ('Tls' == TLS1.0)  ;
			        write-verbose "PRE: `$CurrentVersionTlsLabel : $($CurrentVersionTlsLabel )" ;
			        # psv6+ already covers, test via the SslProtocol parameter presense
			        if ('SslProtocol' -notin (Get-Command Invoke-RestMethod).Parameters.Keys) {
				        $currentMaxTlsValue = [Math]::Max([Net.ServicePointManager]::SecurityProtocol.value__,[Net.SecurityProtocolType]::Tls.value__) ;
				        write-verbose "`$currentMaxTlsValue : $($currentMaxTlsValue )" ;
				        $newerTlsTypeEnums = [enum]::GetValues('Net.SecurityProtocolType') | Where-Object { $_ -gt $currentMaxTlsValue }
				        if($newerTlsTypeEnums){
					        write-verbose "Appending upgraded/missing TLS `$enums:`n$(($newerTlsTypeEnums -join ','|out-string).trim())" ;
				        } else {
					        write-verbose "Current TLS `$enums are up to date with max rev available on this machine" ;
				        };
				        $newerTlsTypeEnums | ForEach-Object {
					        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor $_
				        } ;
			        } ;
                    $smsg = "#*------v Function _connect-ExOP v------" ;
                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                    function _connect-ExOP{
                        [CmdletBinding()]
                        PARAM(
                            [Parameter(Position=0,Mandatory=$true,HelpMessage="Exchange server AD Summary system object[-Server EXSERVER.DOMAIN.COM]")]
                                [system.object]$Server,
                            [Parameter(Position=1,HelpMessage = "Use specific Credentials[-Credentials [credential object]")]
                                [Management.Automation.PSCredential]$credential
                        );
                        $verbose = $($VerbosePreference -eq "Continue") ;
                        if([double]$ExVersNum = [regex]::match($Server.version,"Version\s(\d+\.\d+)\s\(Build\s(\d+\.\d+)\)").groups[1].value){
                            switch -regex ([string]$ExVersNum) {
                                '15.2' { $isEx2019 = $true ; $ExVers = 'Ex2019' }
                                '15.1' { $isEx2016 = $true ; $ExVers = 'Ex2016'}
                                '15.0' { $isEx2013 = $true ; $ExVers = 'Ex2013'}
                                '14.*' { $isEx2010 = $true ; $ExVers = 'Ex2010'}
                                '8.*' { $isEx2007 = $true ; $ExVers = 'Ex2007'}
                                '6.5' { $isEx2003 = $true ; $ExVers = 'Ex2003'}
                                '6' {$isEx2000 = $true ; $ExVers = 'Ex2000'} ;
                                default {
                                    $smsg = "UNRECOGNIZED ExVersNum.Major.Minor string:$($ExVersNum)! ABORTING!" ;
                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                    THROW $SMSG ;
                                    BREAK ;
                                }
                            } ;
                        }else {
                            $smsg = "UNABLE TO RESOLVE `$ExVersNum from `$Server.version:$($Server.version)!" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            throw $smsg ;
                            break ;
                        } ;
                        if($Server.RoleNames -eq 'EDGE'){
                            if(($isLocalExchangeServer = (Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\Setup')) -or
                                ($isLocalExchangeServer = (Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup')) -or
                                $ByPassLocalExchangeServerTest)
                            {
                                if((Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\EdgeTransportRole') -or
                                        (Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\EdgeTransportRole'))
                                {
                                    $smsg = "We are on Exchange Edge Transport Server"
                                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                                    $IsEdgeTransport = $true
                                }
                                TRY {
                                    Get-ExchangeServer -ErrorAction Stop | Out-Null
                                    $smsg = "Exchange PowerShell Module already loaded."
                                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                                    $passed = $true 
                                }CATCH {
                                    $smsg = "Failed to run Get-ExchangeServer"
                                    if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                                    if($isLocalExchangeServer){
                                        write-host  "Loading Exchange PowerShell Module..."
                                        TRY{
                                            if($IsEdgeTransport){
                                                # implement local snapins access on edge role: Only way to get access to EMS commands.
                                                [xml]$PSSnapIns = Get-Content -Path "$env:ExchangeInstallPath\Bin\exshell.psc1" -ErrorAction Stop
                                                ForEach($PSSnapIn in $PSSnapIns.PSConsoleFile.PSSnapIns.PSSnapIn){
                                                    write-verbose ("Trying to add PSSnapIn: {0}" -f $PSSnapIn.Name)
                                                    Add-PSSnapin -Name $PSSnapIn.Name -ErrorAction Stop
                                                } ; 
                                                Import-Module $env:ExchangeInstallPath\bin\Exchange.ps1 -ErrorAction Stop ; 
                                                $passed = $true #We are just going to assume this passed.
                                            }else{
                                                Import-Module $env:ExchangeInstallPath\bin\RemoteExchange.ps1 -ErrorAction Stop
                                                Connect-ExchangeServer -Auto -ClientApplication:ManagementShell
                                                $passed = $true #We are just going to assume this passed.
                                            } 
                                        }CATCH {
                                            $smsg = "Failed to Load Exchange PowerShell Module..." ; 
                                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                                            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                                        }                               
                                    } ;
                                } FINALLY {
                                    if($LoadExchangeVariables -and $passed -and $isLocalExchangeServer){
                                        if($ExInstall -eq $null -or $ExBin -eq $null){
                                            if(Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\Setup'){
                                                $Global:ExInstall = (Get-ItemProperty HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\Setup).MsiInstallPath
                                            }else{
                                                $Global:ExInstall = (Get-ItemProperty HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup).MsiInstallPath
                                            }
        
                                            $Global:ExBin = $Global:ExInstall + "\Bin"
        
                                            $smsg = ("Set ExInstall: {0}" -f $Global:ExInstall)
                                            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                                            $smsg = ("Set ExBin: {0}" -f $Global:ExBin)
                                            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                                        } ; 
                                    } ; 
                                } ; 
                            } else  {
                                $smsg = "Does not appear to be an Exchange 2010 or newer server." ; 
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 

                            }
                            if(get-command -Name Get-OrganizationConfig -ea 0){
                                $smsg = "Running in connected/Native EMS" ; 
                                if($silent){}elseif ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                                Return $true ; 
                            } else { 
                                TRY{
                                    $smsg = "Initiating Edge EMS local session (exshell.psc1 & exchange.ps1)" ; 
                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                    #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                                    # 5;36 PM 5/30/2024 didn't work, went off to nowhere for a long time, and exited the script
                                    #& (gcm powershell.exe).path -PSConsoleFile "$($env:ExchangeInstallPath)bin\exshell.psc1" -noexit -command ". '$($env:ExchangeInstallPath)bin\Exchange.ps1'"
                                    <# [Adding the Transport Server to Exchange - Mark Lewis Blog](https://marklewis.blog/2020/11/19/adding-the-transport-server-to-exchange/)
                                    To access the management console on the transport server, I opened PowerShell then ran
                                    exshell.psc1
                                    Followed by
                                    exchange.ps1
                                    At this point, I was able to create a new subscription using he following PowerShel
                                    #>
                                    invoke-command exshell.psc1 ; 
                                    invoke-command exchange.ps1
                                    if(get-command -Name Get-OrganizationConfig -ea 0){
                                        $smsg = "Running in connected/Native EMS" ; 
                                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                        #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                                        Return $true ;
                                    } else { return $false };  
                                }CATCH{
                                    Write-Error $_ ;
                                } ;
                            } ; 
                        } else {
                            $pltNPSS=@{ConnectionURI="http://$($Server.FQDN)/powershell"; ConfigurationName='Microsoft.Exchange' ; name="Exchange$($ExVersNum.tostring())"} ;
                            # use ExVersUnm dd instead of hardcoded (Exchange2010)
                            if($ExVersNum -ge 15){
                                $smsg = "EXOP.15+:Adding -Authentication Kerberos" ;
                                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                                $pltNPSS.add('Authentication',"Kerberos") ;
                                $pltNPSS.name = $ExVers ;
                            } ;
                            $smsg = "Adding EMS (connecting to $($Server.FQDN))..." ;
                            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                            $smsg = "New-PSSession w`n$(($pltNPSS|out-string).trim())" ;
                            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                            $ExPSS = New-PSSession @pltNPSS  ;
                            $ExIPSS = Import-PSSession $ExPSS -allowclobber ;
                            $ExPSS | write-output ;
                            $ExPSS= $ExIPSS = $null ;
                        } ; 
                    } ;
                    $smsg = "#*------^ END Function _connect-ExOP ^------" ;
                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                    $pltGADX=@{
                        ErrorAction='Stop';
                    } ;
                } ;
                PROCESS{
                    if($PSBoundParameters.ContainsKey('credential')){
                        $pltGADX.Add('credential',$credential) ;
                    }
                    if($SiteName){
                        $pltGADX.Add('siteName',$siteName) ;
                    } ;
                    if($RoleNames){
                        $pltGADX.Add('RoleNames',$RoleNames) ;
                    } ;
                    TRY{
                        if($discover){
                            $smsg = "Getting list of Exchange Servers" ;
                            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                            $exchServers=get-ADExchangeServerTDO @pltGADX | sort responsetime ;
                        }else{
                            $exchServers=get-ADExchangeServerTDO @pltGADX | sort responsetime ;
                        } ;
                        $pltTW=@{
                            'ErrorAction'='Stop';
                        } ;
                        $pltCXOP = @{
                            verbose = $($VerbosePreference -eq "Continue") ;
                        } ;
                        if($pltGADX.credential){
                            $pltCXOP.Add('Credential',$pltCXOP.Credential) ;
                        } ;
                        $prpPSS = 'Id','Name','ComputerName','ComputerType','State','ConfigurationName','Availability' ; 
                        foreach($exServer in $exchServers){
                            $smsg = "testing conn to:$($exServer.name.tostring())..." ; 
                            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                            if(get-command -module (get-module |?{$_.name -like 'tmp_*'}).name -name 'get-OrganizationConfig' -ea SilentlyContinue){
                                if($pssEXOP = Get-PSSession |  where-object { ($_.ConfigurationName -eq 'Microsoft.Exchange') -AND ( $_.runspace.ConnectionInfo.AppName -match '^/(exadmin|powershell)$') -AND ( $_.runspace.ConnectionInfo.Port -eq '80') }){
                                    if($pssEXOP.State -ne "Opened" -OR $pssEXOP.Availability -ne "Available"){
                                        $pssEXOP | remove-pssession ; $pssEXOP = $null ;
                                    } ;
                                } ; 
                            } else {
                                $smsg = "(mangled ExOP conn: disconnect/reconnect...)" ;
                                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                                if($pssEXOP = Get-PSSession |  where-object { ($_.ConfigurationName -eq 'Microsoft.Exchange') -AND ( $_.runspace.ConnectionInfo.AppName -match '^/(exadmin|powershell)$') -AND ( $_.runspace.ConnectionInfo.Port -eq '80') }){
                                    if($pssEXOP.State -ne "Opened" -OR $pssEXOP.Availability -ne "Available"){
                                        $pssEXOP | remove-pssession ; $pssEXOP = $null ;
                                    } ;
                                } ; 
                            } ;
                            if(-not $pssEXOP){
                                $smsg = "Connecting to: $($exServer.FQDN)" ;
                                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                                if($NoTest){
                                    $ExPSS =$ExPSS = _connect-ExOP @pltCXOP -Server $exServer
                                } else {
                                    TRY{
                                        $smsg = "Testing Connection: $($exServer.FQDN)" ;
                                        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                                        If(test-connection $exServer.FQDN -count 1 -ea 0) {
                                            $smsg = "confirmed pingable..." ;
                                            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                                        } else {
                                            $smsg = "Unable to Ping $($exServer.FQDN)" ; ;
                                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                                            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                        } ;
                                        $smsg = "Testing WinRm: $($exServer.FQDN)" ;
                                        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                                        $winrm=Test-WSMan @pltTW -ComputerName $exServer.FQDN ;
                                        if($winrm){
                                            $ExPSS = _connect-ExOP @pltCXOP -Server $exServer;
                                        } else {
                                            $smsg = "Unable to Test-WSMan $($exServer.FQDN) (skipping)" ; ;
                                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                        } ;
                                    }CATCH{
                                        $errMsg="Server: $($exServer.FQDN)] $($_.Exception.Message)" ;
                                        Write-Error -Message $errMsg ;
                                        continue ;
                                    } ;
                                };
                            } else {
                                $smsg = "$((get-date).ToString('HH:mm:ss')):Accepting first valid connection w`n$(($pssEXOP | ft -a $prpPSS|out-string).trim())" ; 
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Success }
                                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                $ExPSS = $pssEXOP ; 
                                break ; 
                            }  ;
                        } ;
                    }CATCH{
                        Write-Error $_ ;
                    } ;
                } ;
                END{
                    if(-not $ExPSS){
                        $smsg = "NO SUCCESSFUL CONNECTION WAS MADE, WITH THE SPECIFIED INPUTS!" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        $smsg = "(returning `$false to the pipeline...)" ;
                        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                        return $false
                    } else{
                        if($ExPSS.State -eq "Opened" -AND $ExPSS.Availability -eq "Available"){
                            if(-not ($OrgName = ((get-OrganizationConfig).DistinguishedName.split(',') |?{$_ -like 'DC=*'}) -join '.' -replace 'DC=','')){
                                $smsg = "Missing Exchange Connection! (no (Get-OrganizationConfig).name returned)" ;
                                throw $smsg ;
                                $smsg | write-warning  ;
                            } else {
                                $smsg = "(connected to EXOP.Org:$($orgName))" ;
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Success }
                                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            } ;
                            return $ExPSS
                        } ;
                    } ; 
                } ;
            } ;
        } ; 
        #*------^ END Function Connect-ExchangeServerTDO ^------
        #endregion CONNEXOPTDO ; #*------^ END CONNEXOPTDO ^------
    
        #region GADEXSERVERTDO ; #*------v  v------
        #*------v Function get-ADExchangeServerTDO v------
        if(-not(get-command get-ADExchangeServerTDO -ea 0)){
            Function get-ADExchangeServerTDO {
                <#
                .SYNOPSIS
                get-ADExchangeServerTDO.ps1 - Returns Exchangeserver summary(s) from AD records
                .NOTES
                Version     : 3.0.1
                Author      : Todd Kadrie
                Website     : http://www.toddomation.com
                Twitter     : @tostka / http://twitter.com/tostka
                CreatedDate : 2015-09-03
                FileName    : get-ADExchangeServerTDO.ps1
                License     : (none-asserted)
                Copyright   : (none-asserted)
                Github      : https://github.com/tostka/verb-Ex2010
                Tags        : Powershell, ActiveDirectory, Exchange, Discovery
                AddedCredit : Mike Pfeiffer
                AddedWebsite: mikepfeiffer.net
                AddedTwitter: URL
                AddedCredit : Sammy Krosoft 
                AddedWebsite: http://aka.ms/sammy
                AddedTwitter: URL
                AddedCredit : Brian Farnsworth
                AddedWebsite: https://codeandkeep.com/
                AddedTwitter: URL
                REVISIONS
                * 3:57 PM 11/26/2024 updated simple write-host,write-verbose with full pswlt support;  syncd dbg & vx10 copies.
                * 12:57 PM 6/11/2024 Validated, Ex2010 & Ex2019, hub, mail & edge roles: tested ☑️ on CMW mail role (Curly); and Jumpbox; copied in CBH from repo copy, which has been updated/debugged compat on CMW Edge 
                * 2:05 PM 8/28/2023 REN -> Get-ExchangeServerInSite -> get-ADExchangeServerTDO (aliased orig); to better steer profile-level options - including in cmw org, added -TenOrg, and default Site to constructed vari, targeting new profile $XXX_ADSiteDefault vari; Defaulted -Roles to HUB,CAS as well.
                * 3:42 PM 8/24/2023 spliced together combo of my long-standing, and some of the interesting ideas BF's version had. Functional prod:
                    - completely removed ActiveDirectory module dependancies from BF's code, and reimplemented in raw ADSI calls. Makes it fully portable, even into areas like Edge DMZ roles, where ADMS would never be installed.

                * 3:17 PM 8/23/2023 post Edge testing: some logic fixes; add: -Names param to filter on server names; -Site & supporting code, to permit lookup against sites *not* local to the local machine (and bypass lookup on the local machine) ; 
                    ren $Ex10siteDN -> $ExOPsiteDN; ren $Ex10configNC -> $ExopconfigNC
                * 1:03 PM 8/22/2023 minor cleanup
                * 10:31 AM 4/7/2023 added CBH expl of postfilter/sorting to draw predictable pattern 
                * 4:36 PM 4/6/2023 validated Psv51 & Psv20 and Ex10 & 16; added -Roles & -RoleNames params, to perform role filtering within the function (rather than as an external post-filter step). 
                For backward-compat retain historical output field 'Roles' as the msexchcurrentserverroles summary integer; 
                use RoleNames as the text role array; 
                    updated for psv2 compat: flipped hash key lookups into properties, found capizliation differences, (psv2 2was all lower case, wouldn't match); 
                flipped the [pscustomobject] with new... psobj, still psv2 doesn't index the hash keys ; updated for Ex13+: Added  16  "UM"; 20  "CAS, UM"; 54  "MBX" Ex13+ ; 16385 "CAS" Ex13+ ; 16439 "CAS, HUB, MBX" Ex13+
                Also hybrided in some good ideas from SammyKrosoft's Get-SKExchangeServers.psm1 
                (emits Version, Site, low lvl Roles # array, and an array of Roles, for post-filtering); 
                # 11:20 AM 4/21/2021 fixed/suppressed noisy verbose calls
                * 12:08 PM 5/15/2020 fixed vpn issue: Try/Catch'd around recently failing $ADSite::GetComputerSite().GetDirectoryEntry().distinguishedName qry
                * 11:22 AM 3/13/2020 Get-ExchangeServerInSite added a ping-test, to only return matches that are pingable, added -NoPing param, to permit (faster) untested bypass
                * 6:59 PM 1/15/2020 cleanup
                # 10:03 AM 11/16/2018 Get-ExchangeServerInSite:can't do AD-related functions when not AD authentictaed (home, pre-vpn connect). Added if/then test on status and abort balance when false.
                * 11/18/18 BF's posted rev
                # 12:10 PM 8/1/2017 updated example code at bottom, to accommodate variant sites
                # 11:28 AM 3/31/2016 validated that latest round of updates are still functional
                #1:58 PM 9/3/2015 - added pshelp and some docs
                #April 12, 2010 - web version
                .DESCRIPTION
                get-ADExchangeServerTDO.ps1 - Returns Exchangeserver summary(s) from AD records

                Hybrided together ideas from Brian Farnsworth's blog post
                [PowerShell - ActiveDirectory and Exchange Servers – CodeAndKeep.Com – Code and keep calm...](https://codeandkeep.com/PowerShell-ActiveDirectory-Exchange-Part1/)
                ... with much older concepts from  Sammy Krosoft, and much earlier Mike Pfeiffer. 

                - Subbed in MP's use of ADSI for ActiveDirectory Ps mod cmds - it's much more dependancy-free; doesn't require explicit install of the AD ps module
                ADSI support is built into windows.
                - spliced over my addition of Roles, RoleNames, Name & NoTest params, for prefiltering and suppressing testing.


                [briansworth · GitHub](https://github.com/briansworth)

                Uses an ADSI DirectorySearcher to search the current Active Directory site for Exchange on-prem servers.
                        Intent is to discover connection points for Powershell, wo the need to preload/pre-connect to Exchange.

                        But, as a non-Exchange-Management-Shell-dependant info source on Exchange Server configs, it can be used before connection, with solely AD-available data, to check configuration spes on the subject server(s). 

                        For example, this query will return sufficient data under Version to indicate which revision of Exchange is in use:


                        Returned object (in array):
                        Site      : {ADSITENAME}
                        Roles     : {64}
                        Version   : {Version 15.1 (Build 32375.7)}
                        Name      : SERVERNAME
                        RoleNames : EDGE
                        FQDN      : SERVERNAME.DOMAIN.TLD

                        ... includes the post-filterable Role property ($_.Role -contains 'CAS') which reflects the following
                        installed-roles ('msExchCurrentServerRoles') on the discovered servers
                            2   {"MBX"} # Ex10
                            4   {"CAS"}
                            16  {"UM"}
                            20  {"CAS, UM" -split ","} # 
                            32  {"HUB"}
                            36  {"CAS, HUB" -split ","}
                            38  {"CAS, HUB, MBX" -split ","}
                            54  {"MBX"} # Ex13+
                            64  {"EDGE"}
                            16385   {"CAS"} # Ex13+
                            16439   {"CAS, HUB, MBX"  -split ","} # Ex13+
                .PARAMETER Roles
                Array of msExchCurrentServerRoles 'role' integers to be filtered against (2|4|16|20|32|36|38|54|64|16385|16439)[-Roles @(38,16385)]
                .PARAMETER RoleNames
                Array of Server 'Role' name strings to be filtered against (MBX|CAS|HUB|UM|MBX|EDGE)[-RoleNames 'HUB','CAS']
                .PARAMETER Server
                Array of Server name strings to be filtered against[-Server 'SERVER1','SERVER2']
                .PARAMETER SiteName
                Name of specific AD SiteName to be searched for ExchangeServers (defaults to global variable `$TenOrg_ADSiteDefaultName if present)[-SiteName 'SITENAME']
                .PARAMETER TenOrg
                Tenant Tag (3-letter abbrebiation - defaults to variable `$global:o365_TenOrgDefault if present)[-TenOrg 'XYZ']
                .PARAMETER NoPing
                Switch to suppress default 'pingable' test (e.g. returns all matches, no testing)[-NoPing]
                .INPUTS
                None. Does not accepted piped input.(.NET types, can add description)
                .OUTPUTS
                None. Returns no objects or output (.NET types)
                System.Boolean
                [| get-member the output to see what .NET obj TypeName is returned, to use here]
                System.Array of System.Object's
                .EXAMPLE
                PS> If(!($ExchangeServer)){$ExchangeServer = (get-ADExchangeServerTDO| ?{$_.RoleNames -contains 'CAS' -OR $_.RoleNames -contains 'HUB' -AND ($_.FQDN -match "^SITECODE") } | Get-Random ).FQDN
                Return a random Hub Cas Role server in the local Site with a fqdn beginning SITECODE
                .EXAMPLE
                PS> $localADExchserver = get-ADExchangeServerTDO -Names $env:computername -SiteName ([System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::GetComputerSite().name)
                Demo, if run from an Exchange server, return summary details about the local server (-SiteName isn't required, is default imputed from local server's Site, but demos explicit spec for remote sites)
                .EXAMPLE
                PS> $regex = '(' + [regex]($ADSiteCodeUK,$ADSiteCodeAU -join "|") + ')'
                PS> switch -regex ($($env:computername).substring(0,3)){
                PS>    "$($ADSiteCodeUS)" {$tExRole=36 } ;
                PS>    "$($regex)" {$tExRole= 38 }  default { write-error "$((get-date).ToString('HH:mm:ss')):UNRECOGNIZED SERVER PREFIX!."; } ;
                PS> } ;
                PS> $exhubcas = (get-ADExchangeServerTDO |?{($_.roles -eq $tExRole) -AND ($_.FQDN -match "$($env:computername.substring(0,3)).*")} | Get-Random ).FQDN ;
                Use a switch block to select different role combo targets for a given server fqdn prefix string.
                .EXAMPLE
                PS> $ExchangeServer = get-ADExchangeServerTDO | ?{$_.Roles -match '(4|20|32|36|38|16385|16439)'} | select -expand fqdn | get-random ; 
                Another/Older approach filtering on the Roles integer (targeting combos with Hub or CAS in the mix)
                .EXAMPLE
                PS> $ret = get-ADExchangeServerTDO -Roles @(4,20,32,36,38,16385,16439) -verbose 
                Demo use of the -Roles param, feeding it an array of Role integer values to be filtered against. In this case, the Role integers that include a CAS or HUB role.
                .EXAMPLE
                PS> $ret = get-ADExchangeServerTDO -RoleNames 'HUB','CAS' -verbose ;
                Demo use of the -RoleNames param, feeding it the array 'HUB','CAS' Role name strings to be filtered against
                PS> $ret = get-ADExchangeServerTDO -Names 'SERVERName' -verbose ;
                Demo use of the -RoleNames param, feeding it the array 'HUB','CAS' Role name strings to be filtered against
                .EXAMPLE
                PS> $ExchangeServer = get-ADExchangeServerTDO | sort version,roles,name | ?{$_.rolenames -contains 'CAS'}  | select -last 1 | select -expand fqdn ;
                Demo post sorting & filtering, to deliver a rule-based predictable pattern for server selection: 
                Above will always pick the highest Version, 'CAS' RoleName containing, alphabetically last server name (that is pingable). 
                And should stick to that pattern, until the servers installed change, when it will shift to the next predictable box.
                .EXAMPLE
                PS> $ExOPServer = get-ADExchangeServerTDO -Name LYNMS650 -SiteName Lyndale
                PS> if([double]$ExVersNum = [regex]::match($ExOPServer.version,"Version\s(\d+\.\d+)\s\(Build\s(\d+\.\d+)\)").groups[1].value){
                PS>     switch -regex ([string]$ExVersNum) {
                PS>         '15\.2' { $isEx2019 = $true ; $ExVers = 'Ex2019' }
                PS>         '15\.1' { $isEx2016 = $true ; $ExVers = 'Ex2016'}
                PS>         '15\.0' { $isEx2013 = $true ; $ExVers = 'Ex2013'}
                PS>         '14\..*' { $isEx2010 = $true ; $ExVers = 'Ex2010'}
                PS>         '8\..*' { $isEx2007 = $true ; $ExVers = 'Ex2007'}
                PS>         '6\.5' { $isEx2003 = $true ; $ExVers = 'Ex2003'}
                PS>         '6|6\.0' {$isEx2000 = $true ; $ExVers = 'Ex2000'} ;
                PS>         default {
                PS>             $smsg = "UNRECOGNIZED ExchangeServer.AdminDisplayVersion.Major.Minor string:$($ExOPServer.version)! ABORTING!" ;
                PS>             write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" ;
                PS>         }
                PS>     } ; 
                PS> }else {
                PS>     $smsg = "UNABLE TO RESOLVE `$ExVersNum from `$ExOPServer.version:$($ExOPServer.version)!" ; 
                PS>     write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)"  ; 
                PS>     throw $smsg ; 
                PS>     break ; 
                PS> } ; 
                Demo of parsing the returned Version property, into the proper Exchange Server revision.      
                .LINK
                https://github.com/tostka/verb-XXX
                .LINK
                https://bitbucket.org/tostka/powershell/
                .LINK
                http://mikepfeiffer.net/2010/04/find-exchange-servers-in-the-local-active-directory-site-using-powershell/
                .LINK
                https://github.com/SammyKrosoft/Search-AD-Using-Plain-PowerShell/blob/master/Get-SKExchangeServers.psm1
                .LINK
                https://github.com/tostka/verb-Ex2010
                .LINK
                https://codeandkeep.com/PowerShell-ActiveDirectory-Exchange-Part1/
                #>
                [CmdletBinding()]
                [Alias('Get-ExchangeServerInSite')]
                PARAM(
                    [Parameter(Position=0,HelpMessage="Array of Server name strings to be filtered against[-Server 'SERVER1','SERVER2']")]
                        [string[]]$Server,
                    [Parameter(Position=1,HelpMessage="Name of specific AD site to be searched for ExchangeServers (defaults to global variable `$TenOrg_ADSiteDefaultName if present)[-SiteName 'SITENAME']")]
                        [Alias('Site')]
                        [string]$SiteName = (gv -name "$($TenOrg)_ADSiteDefaultName" -ea 0).Value,
                    [Parameter(Position=2,HelpMessage="Array of Server 'Role' name strings to be filtered against (MBX|CAS|HUB|UM|MBX|EDGE)[-RoleNames 'HUB','CAS']")]
                        [ValidateSet('MBX','CAS','HUB','UM','MBX','EDGE')]
                        [string[]]$RoleNames = @('HUB','CAS'),
                    [Parameter(HelpMessage="Array of msExchCurrentServerRoles 'role' integers to be filtered against (2|4|16|20|32|36|38|54|64|16385|16439)[-Roles @(38,16385)]")]
                        [ValidateSet(2,4,16,20,32,36,38,54,64,16385,16439)]
                        [int[]]$Roles,
                    [Parameter(HelpMessage="Switch to suppress default 'pingable' test (e.g. returns all matches, no testing)[-NoTest]")]
                        [Alias('NoPing')]
                        [switch]$NoTest,
                    [Parameter(HelpMessage="Milliseconds of max timeout to wait during port 80 test (defaults 100)[-SpeedThreshold 500]")]
                        [int]$SpeedThreshold=100,
                    [Parameter(Mandatory=$FALSE,HelpMessage="Tenant Tag (3-letter abbrebiation - defaults to global:o365_TenOrgDefault if present)[-TenOrg 'XYZ']")]
                        [ValidateNotNullOrEmpty()]
                        [string]$TenOrg = $global:o365_TenOrgDefault,
                    [Parameter(Mandatory = $false, HelpMessage = "Use specific Credentials[-Credentials [credential object]]")]
                        [System.Management.Automation.PSCredential]$Credential
                ) ;
                BEGIN{
                    ${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name ;
                    $Verbose = ($VerbosePreference -eq 'Continue') ;
                    $_sBnr="#*======v $(${CmdletName}): v======" ;
                    $smsg = $_sBnr ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level H1 }
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                } ;
                PROCESS{
                    TRY{
                        $configNC = ([ADSI]"LDAP://RootDse").configurationNamingContext ;
                        $search = new-object DirectoryServices.DirectorySearcher([ADSI]"LDAP://$configNC") ;
                        $bLocalEdge = $false ; 
                        if($Sitename -eq $env:COMPUTERNAME){
                            $smsg = "`$SiteName -eq `$env:COMPUTERNAME:$($SiteName):$($env:COMPUTERNAME)" ; 
                            $smsg += "`nThis computer appears to be an EdgeRole system (non-ADConnected)" ; 
                            $smsg += "`n(Blanking `$sitename and continuing discovery)" ; 
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                            #$bLocalEdge = $true ; 
                            $SiteName = $null ; 
                    
                        } ; 
                        If($siteName){
                            $smsg = "Getting Site: $siteName" ;
                            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                            $objectClass = "objectClass=site" ;
                            $objectName = "name=$siteName" ;
                            $search.Filter = "(&($objectClass)($objectName))" ;
                            $site = ($search.Findall()) ;
                            $siteDN = ($site | select -expand properties).distinguishedname  ;
                        } else {
                            $smsg = "(No -Site specified, resolving site from local machine domain-connection...)" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Prompt }
                            else{ write-host -foregroundcolor green "$($smsg)" } ;
                            TRY{$siteDN = [System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::GetComputerSite().GetDirectoryEntry().distinguishedName}
                            CATCH [System.Management.Automation.MethodInvocationException]{
                                $ErrTrapd=$Error[0] ;
                                if(($ErrTrapd.Exception -match 'The computer is not in a site.') -AND $env:ExchangeInstallPath){
                                    $smsg = "$($env:computername) is non-ADdomain-connected" ;
                                    $smsg += "`nand has `$env:ExchangeInstalled populated: Likely Edge Server" ;
                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Prompt }
                                    else{ write-host -foregroundcolor YELLOW "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                    $vers = (get-item "$($env:ExchangeInstallPath)\Bin\Setup.exe").VersionInfo.FileVersionRaw ; 
                                    $props = @{
                                        Name=$env:computername;
                                        FQDN = ([System.Net.Dns]::gethostentry($env:computername)).hostname;
                                        Version = "Version $($vers.major).$($vers.minor) (Build $($vers.Build).$($vers.Revision))" ; 
                                        #"$($vers.major).$($vers.minor)" ; 
                                        #$exServer.serialNumber[0];
                                        Roles = [System.Object[]]64 ;
                                        RoleNames = @('EDGE');
                                        DistinguishedName =  "CN=$($env:computername),CN=Servers,CN=Exchange Administrative Group (FYDIBOHF23SPDLT),CN=Administrative Groups,CN=First Organization,CN=Microsoft Exchange,CN=Services,CN=Configuration,CN={nnnnnnnn-FAKE-GUID-nnnn-nnnnnnnnnnnn}" ;
                                        Site = [System.Object[]]'NOSITE'
                                        ResponseTime = if($rsp){$rsp.ResponseTime} else { 0} ;
                                        NOTE = "This summary object, returned for a non-AD-connected EDGE server, *approximates* what would be returned on an AD-connected server" ;
                                    } ;
                            
                                    $smsg = "(-NoTest:Defaulting Fast:`$true)" ;
                                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                                    $props.add('Fast',$true) ;
                            
                                    return (New-Object -TypeName PsObject -Property $props) ;
                                }elseif(-not $env:ExchangeInstallPath){
                                    $smsg = "Non-Domain Joined machine, with NO ExchangeInstallPath e-vari: `nExchange is not installed locally: local computer resolution fails:`nPlease specify an explicit -Server, or -SiteName" ;
                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                    $false | write-output ;
                                } else {
                                    $smsg = "$($env:computername) is both NON-Domain-joined -AND lacks an Exchange install (NO ExchangeInstallPath e-vari)`nPlease specify an explicit -Server, or -SiteName" ;
                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                    $false | write-output ;
                                };
                            } CATCH {
                                $siteDN =$ExOPsiteDN ;
                                write-warning "`$siteDN lookup FAILED, deferring to hardcoded `$ExOPsiteDN string in infra file!" ;
                            } ;
                        } ;
                        $smsg = "Getting Exservers in Site:$($siteDN)" ;
                        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                        $search = new-object DirectoryServices.DirectorySearcher([ADSI]"LDAP://$configNC") ;
                        $objectClass = "objectClass=msExchExchangeServer" ;
                        $version = "versionNumber>=1937801568" ;
                        $site = "msExchServerSite=$siteDN" ;
                        $search.Filter = "(&($objectClass)($version)($site))" ;
                        $search.PageSize = 1000 ;
                        [void] $search.PropertiesToLoad.Add("name") ;
                        [void] $search.PropertiesToLoad.Add("msexchcurrentserverroles") ;
                        [void] $search.PropertiesToLoad.Add("networkaddress") ;
                        [void] $search.PropertiesToLoad.Add("msExchServerSite") ;
                        [void] $search.PropertiesToLoad.Add("serialNumber") ;
                        [void] $search.PropertiesToLoad.Add("DistinguishedName") ;
                        $exchServers = $search.FindAll() ;
                        $Aggr = @() ;
                        foreach($exServer in $exchServers){
                            $fqdn = ($exServer.Properties.networkaddress |
                                Where-Object{$_ -match '^ncacn_ip_tcp:'}).split(':')[1] ;
                            if($NoTest){} else {
                                $rsp = test-connection $fqdn -count 1 -ea 0 ;
                            } ;
                            $props = @{
                                Name = $exServer.Properties.name[0]
                                FQDN=$fqdn;
                                Version = $exServer.Properties.serialnumber
                                Roles = $exserver.Properties.msexchcurrentserverroles
                                RoleNames = $null ;
                                DistinguishedName = $exserver.Properties.distinguishedname;
                                Site = @("$($exserver.Properties.msexchserversite -Replace '^CN=|,.*$')") ;
                                ResponseTime = if($rsp){$rsp.ResponseTime} else { 0} ;
                            } ;
                            $props.RoleNames = switch ($exserver.Properties.msexchcurrentserverroles){
                                2       {"MBX"}
                                4       {"CAS"}
                                16      {"UM"}
                                20      {"CAS;UM".split(';')}
                                32      {"HUB"}
                                36      {"CAS;HUB".split(';')}
                                38      {"CAS;HUB;MBX".split(';')}
                                54      {"MBX"}
                                64      {"EDGE"}
                                16385   {"CAS"}
                                16439   {"CAS;HUB;MBX".split(';')}
                            }
                            if($NoTest){
                                $smsg = "(-NoTest:Defaulting Fast:`$true)" ;
                                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                                $props.add('Fast',$true) ;
                            }else {
                                $props.add('Fast',[boolean]($rsp.ResponseTime -le $SpeedThreshold)) ;
                            };
                            $Aggr += New-Object -TypeName PsObject -Property $props ;
                        } ;
                        $httmp = @{} ;
                        if($Roles){
                            [regex]$rgxRoles = ('(' + (($roles |%{[regex]::escape($_)}) -join '|') + ')') ;
                            $matched =  @( $aggr | ?{$_.Roles -match $rgxRoles}) ;
                            foreach($m in $matched){
                                if($httmp[$m.name]){} else {
                                    $httmp[$m.name] = $m ;
                                } ;
                            } ;
                        } ;
                        if($RoleNames){
                            foreach ($RoleName in $RoleNames){
                                $matched = @($Aggr | ?{$_.RoleNames -contains $RoleName} ) ;
                                foreach($m in $matched){
                                    if($httmp[$m.name]){} else {
                                        $httmp[$m.name] = $m ;
                                    } ;
                                } ;
                            } ;
                        } ;
                        if($Server){
                            foreach ($Name in $Server){
                                $matched = @($Aggr | ?{$_.Name -eq $Name} ) ;
                                foreach($m in $matched){
                                    if($httmp[$m.name]){} else {
                                        $httmp[$m.name] = $m ;
                                    } ;
                                } ;
                            } ;
                        } ;
                        if(($httmp.Values| measure).count -gt 0){
                            $Aggr  = $httmp.Values ;
                        } ;
                        $smsg = "Returning $((($Aggr|measure).count|out-string).trim()) match summaries to pipeline..." ;
                        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                        $Aggr | write-output ;
                    }CATCH{
                        Write-Error $_ ;
                    } ;
                } ;
                END{
                    $smsg = "$($_sBnr.replace('=v','=^').replace('v=','^='))" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level H1 }
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                } ;
            } ;
        }
        #*------^ END Function get-ADExchangeServerTDO ^------ ;
        #endregion GADEXSERVERTDO ; #*------^ END GADEXSERVERTDO ^------

        #region RVARIINVALIDCHARS ; #*------v RVARIINVALIDCHARS v------
        #*------v Function Remove-InvalidVariableNameChars v------
        if(-not (gcm Remove-InvalidVariableNameChars -ea 0)){
            Function Remove-InvalidVariableNameChars ([string]$Name) {
                ($Name.tochararray() -match '[A-Za-z0-9_]') -join '' | write-output ;
            };
        } ;
        #*------^ END Function Remove-InvalidVariableNameChars ^------
        #endregion RVARIINVALIDCHARS ; #*------^ END RVARIINVALIDCHARS ^------

        #endregion FUNCTIONS ; #*======^ END FUNCTIONS ^======

        #region START-LOG-HOLISTIC #*------v START-LOG-HOLISTIC v------
        # Single log for script/function example that accomodates detect/redirect from AllUsers scope'd installed code, and hunts a series of drive letters to find an alternate logging dir (defers to profile variables)
        #${CmdletName} = $rPSCmdlet.MyInvocation.MyCommand.Name ;
        if(!(get-variable LogPathDrives -ea 0)){$LogPathDrives = 'd','c' };
        foreach($budrv in $LogPathDrives){if(test-path -path "$($budrv):\scripts" -ea 0 ){break} } ;
        if(!(get-variable rgxPSAllUsersScope -ea 0)){
            $rgxPSAllUsersScope="^$([regex]::escape([environment]::getfolderpath('ProgramFiles')))\\((Windows)*)PowerShell\\(Scripts|Modules)\\.*\.(ps(((d|m))*)1|dll)$" ;
        } ;
        if(!(get-variable rgxPSCurrUserScope -ea 0)){
            $rgxPSCurrUserScope="^$([regex]::escape([Environment]::GetFolderPath('MyDocuments')))\\((Windows)*)PowerShell\\(Scripts|Modules)\\.*\.(ps((d|m)*)1|dll)$" ;
        } ;
        $pltSL=[ordered]@{Path=$null ;NoTimeStamp=$false ;Tag=$null ;showdebug=$($showdebug) ; Verbose=$($VerbosePreference -eq 'Continue') ; whatif=$($whatif) ;} ;
        # if using [CmdletBinding(SupportsShouldProcess)] + -WhatIf:$($WhatIfPreference):
        #$pltSL=[ordered]@{Path=$null ;NoTimeStamp=$false ;Tag=$null ;showdebug=$($showdebug) ; Verbose=$($VerbosePreference -eq 'Continue') ; whatif=$($WhatIfPreference) ;} ;
        #$pltSL=[ordered]@{Path=$null ;NoTimeStamp=$false ;Tag="$($ticket)-$($TenOrg)-LASTPASS-" ;showdebug=$($showdebug) ; Verbose=$($VerbosePreference -eq 'Continue') ; whatif=$($WhatIfPreference) ;} ;
        #$pltSL.Tag = $ModuleName ; 
        if($ticket){$pltSL.Tag = $ticket} ; 
        if($script:rPSCommandPath){ $prxPath = $script:rPSCommandPath }
        elseif($script:PSCommandPath){$prxPath = $script:PSCommandPath}
        if($rMyInvocation.MyCommand.Definition){$prxPath2 = $rMyInvocation.MyCommand.Definition }
        elseif($MyInvocation.MyCommand.Definition){$prxPath2 = $MyInvocation.MyCommand.Definition } ; 
        if($prxPath){
            if(($prxPath -match $rgxPSAllUsersScope) -OR ($prxPath -match $rgxPSCurrUserScope)){
                $bDivertLog = $true ; 
                switch -regex ($prxPath){
                    $rgxPSAllUsersScope{$smsg = "AllUsers"} 
                    $rgxPSCurrUserScope{$smsg = "CurrentUser"}
                } ;
                $smsg += " context script/module, divert logging into [$budrv]:\scripts" 
                write-verbose $smsg  ;
                if($bDivertLog){
                    if((split-path $prxPath -leaf) -ne $cmdletname){
                        # function in a module/script installed to allusers|cu - defer name to Cmdlet/Function name
                        $pltSL.Path = (join-path -Path "$($budrv):\scripts" -ChildPath "$($cmdletname).ps1") ;
                    } else {
                        # installed allusers|CU script, use the hosting script name
                        $pltSL.Path = (join-path -Path "$($budrv):\scripts" -ChildPath (split-path $prxPath -leaf)) ;
                    }
                } ;
            } else {
                $pltSL.Path = $prxPath ;
            } ;
       }elseif($prxPath2){
            if(($prxPath2 -match $rgxPSAllUsersScope) -OR ($prxPath2 -match $rgxPSCurrUserScope) ){
                 $pltSL.Path = (join-path -Path "$($budrv):\scripts" -ChildPath (split-path $prxPath2 -leaf)) ;
            } elseif(test-path $prxPath2) {
                $pltSL.Path = $prxPath2 ;
            } elseif($cmdletname){
                $pltSL.Path = (join-path -Path "$($budrv):\scripts" -ChildPath "$($cmdletname).ps1") ;
            } else {
                $smsg = "UNABLE TO RESOLVE A FUNCTIONAL `$CMDLETNAME, FROM WHICH TO BUILD A START-LOG.PATH!" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Warn } #Error|Warn|Debug 
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                BREAK ;
            } ; 
        } else{
            $smsg = "UNABLE TO RESOLVE A FUNCTIONAL `$CMDLETNAME, FROM WHICH TO BUILD A START-LOG.PATH!" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Warn } #Error|Warn|Debug 
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            BREAK ;
        }  ;
        write-verbose "start-Log w`n$(($pltSL|out-string).trim())" ; 
        $logspec = start-Log @pltSL ;
        $error.clear() ;
        TRY {
            if($logspec){
                #$logging=$logspec.logging ;
                $logging= $false ; # explicitly turned logfile writing off, just want to use it's path for exports
                $logfile=$logspec.logfile ;
                $transcript=$logspec.transcript ;
                <# 2:30 PM 9/27/2024 no transcript, just want solid logging path discovery
                $stopResults = try {Stop-transcript -ErrorAction stop} catch {} ;
                if($stopResults){
                    $smsg = "Stop-transcript:$($stopResults)" ; 
                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                } ; 
                $startResults = start-Transcript -path $transcript ;
                if($startResults){
                    $smsg = "start-transcript:$($startResults)" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                } ; 
                #>
            } else {throw "Unable to configure logging!" } ;
        } CATCH [System.Management.Automation.PSNotSupportedException]{
            if($host.name -eq 'Windows PowerShell ISE Host'){
                $smsg = "This version of $($host.name):$($host.version) does *not* support native (start-)transcription" ; 
            } else { 
                $smsg = "This host does *not* support native (start-)transcription" ; 
            } ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        } CATCH {
            $ErrTrapd=$Error[0] ;
            $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        } ;
        #endregion START-LOG-HOLISTIC #*------^ END START-LOG-HOLISTIC ^------


        #region SERVICE_CONNECTIONS #*======v SERVICE_CONNECTIONS v======
        # PRETUNE STEERING separately *before* pasting in balance of region
        #*------v STEERING VARIS v------
        $useO365 = $true ;
        $useEXO = $true ; 
        $UseOP=$true ; 
        $UseExOP=$true ;
        $useExopNoDep = $true ; # switch to use Connect-ExchangeServerTDO, vs connect-ex2010 (creds are assumed inherent to the account)
        $useForestWide = $true ; # flag to trigger cross-domain/forest-wide code in AD & EXoP
        $UseOPAD = $true ; 
        $UseMSOL = $false ; # should be hard disabled now in o365
        $UseAAD = $true  ; 
        $useO365 = [boolean]($useO365 -OR $useEXO -OR $UseMSOL -OR $UseAAD)
        $UseOP = [boolean]($UseOP -OR $UseExOP -OR $UseOPAD) ;
        #*------^ END STEERING VARIS ^------
        #*------v EXO V2/3 steering constants v------
        $EOMModName =  'ExchangeOnlineManagement' ;
        $EOMMinNoWinRMVersion = $MinNoWinRMVersion = '3.0.0' ; # support both names
        #*------^ END EXO V2/3 steering constants ^------
        # assert Org from Credential specs (if not param'd)
        # 1:36 PM 7/7/2023 and revised again -  revised the -AND, for both, logic wasn't working
        if($TenOrg){    
            $smsg = "Confirmed populated `$TenOrg" ; 
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        } elseif(-not($tenOrg) -and $Credential){
            $smsg = "(unconfigured `$TenOrg: asserting from credential)" ; 
            if($silent){}elseif ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Prompt } 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            if((get-command get-TenantTag).Parameters.keys -contains 'silent'){
                $TenOrg = get-TenantTag -Credential $Credential -silent ;;
            }else {
                $TenOrg = get-TenantTag -Credential $Credential ;
            }
        } else { 
            # if not using Credentials or a TargetTenants/TenOrg loop, default the $TenOrg on the $env:USERDOMAIN
            $smsg = "(unconfigured `$TenOrg & *NO* `$Credential: fallback asserting from `$env:USERDOMAIN)" ; 
            if($silent){}elseif ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Prompt } 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            switch -regex ($env:USERDOMAIN){
                ([regex]('(' + (( @($TORMeta.legacyDomain,$CMWMeta.legacyDomain)  |foreach-object{[regex]::escape($_)}) -join '|') + ')')).tostring() {$TenOrg = $env:USERDOMAIN.substring(0,3).toupper() } ;
                $TOLMeta.legacyDomain {$TenOrg = 'TOL' }
                default {throw "UNRECOGNIZED `$env:USERDOMAIN!:$($env:USERDOMAIN)" ; exit ; } ;
            } ; 
        } ; 
        #region useO365 ; #*------v useO365 v------
        #$useO365 = $false ; # non-dyn setting, drives variant EXO reconnect & query code
        #if($CloudFirst){ $useO365 = $true } ; # expl: steering on a parameter
        if($useO365){
            #region GENERIC_EXO_CREDS_&_SVC_CONN #*------v GENERIC EXO CREDS & SVC CONN BP v------
            # o365/EXO creds
            <### Usage: Type defaults to SID, if not spec'd - Note: there must be a *logged in & configured *profile*
            $o365Cred = get-TenantCredentials -TenOrg $TenOrg -verbose -userrole SID ;
            Returns a credential set for the $TenOrg Hybrid OnPrem Exchange Org
            .EXAMPLE
            $o365Cred = get-TenantCredentials -TenOrg $TenOrg -verbose -userrole CSVC ;
            Returns the CSVC Userrole credential for the $TenOrg Hybrid OnPrem Exchange Org
            .EXAMPLE
            $o365Cred = get-TenantCredentials -TenOrg $TenOrg -verbose -userrole B2BI ;
            Returns the B2BI Userrole credential for the $TenOrg Hybrid OnPrem Exchange Org
            ###>
            $o365Cred = $null ;
            if($Credential){
                $smsg = "`Credential:Explicit credentials specified, deferring to use..." ; 
                if($silent){}elseif ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                # get-TenantCredentials() return format: (emulating)
                $o365Cred = [ordered]@{
                    Cred=$Credential ; 
                    credType=$null ; 
                } ; 
                $uRoleReturn = resolve-UserNameToUserRole -UserName $Credential.username -verbose:$($VerbosePreference -eq "Continue") ; # Username
                #$uRoleReturn = resolve-UserNameToUserRole -Credential $Credential -verbose = $($VerbosePreference -eq "Continue") ;   # full Credential support
                if($uRoleReturn.UserRole){
                    $o365Cred.credType = $uRoleReturn.UserRole ; 
                } else { 
                    $smsg = "Unable to resolve `$credential.username ($($credential.username))"
                    $smsg += "`nto a usable 'UserRole' spec!" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    throw $smsg ;
                    Break ;
                } ; 
            } else { 
                $pltGTCred=@{TenOrg=$TenOrg ; UserRole=$null; verbose=$($verbose)} ;
                if($UserRole){
                    $smsg = "(`$UserRole specified:$($UserRole -join ','))" ; 
                    if($silent){}elseif ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    $pltGTCred.UserRole = $UserRole; 
                } else { 
                    $smsg = "(No `$UserRole found, defaulting to:'CSVC','SID' " ; 
                    if($silent){}elseif ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                    $pltGTCred.UserRole = 'CSVC','SID' ; 
                } ; 
                $smsg = "get-TenantCredentials w`n$(($pltGTCred|out-string).trim())" ; 
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level verbose } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                $o365Cred = get-TenantCredentials @pltGTCred
            } ; 
            if($o365Cred.credType -AND $o365Cred.Cred -AND $o365Cred.Cred.gettype().fullname -eq 'System.Management.Automation.PSCredential'){
                $smsg = "(validated `$o365Cred contains .credType:$($o365Cred.credType) & `$o365Cred.Cred.username:$($o365Cred.Cred.username)" ;
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                # 9:58 AM 6/13/2024 populate $credential with return, if not populated (may be required for follow-on calls that pass common $Credentials through)
                if((gv Credential) -AND $Credential -eq $null){
                    $credential = $o365Cred.Cred ;
                }elseif($credential.gettype().fullname -eq 'System.Management.Automation.PSCredential'){
                    $smsg = "(`$Credential is properly populated; explicit -Credential was in initial call)" ; 
                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                } else {
                    $smsg = "`$Credential is `$NULL, AND $o365Cred.Cred is unusable to populate!" ;
                    $smsg = "downstream commands will *not* properly pass through usable credentials!" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    throw $smsg ;
                    break ;
                } ;
            } else {
                $smsg = "UNABLE TO RESOLVE FUNCTIONAL CredType/UserRole from specified explicit -Credential:$($Credential.username)!" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                break ;
            } ; 
            if($o365Cred){
                # make it script scope, so we don't have to predetect & purge before using new-variable
                if(get-Variable -Name cred$($tenorg) -scope Script -ea 0 ){ remove-Variable -Name cred$($tenorg) -scope Script } ;
                New-Variable -Name cred$($tenorg) -scope Script -Value $o365Cred.cred ;
                $smsg = "Resolved $($Tenorg) `$o365cred:$($o365Cred.cred.username) (assigned to `$cred$($tenorg))" ;
                if($silent){}elseif ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } else {
                $statusdelta = ";ERROR"; # CHANGE|INCOMPLETE|ERROR|WARN|FAIL ;
                $script:PassStatus += $statusdelta ;
                set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatettus_$($tenorg)).value + $statusdelta) ;
                $smsg = "Unable to resolve $($tenorg) `$o365Cred value!"
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                throw "Unable to resolve $($tenorg) `$o365Cred value!`nEXIT!"
                Break ;
            } ;
            # if we get here, wo a $Credential, w resolved $o365Cred, assign it 
            if(-not $Credential -AND $o365Cred){$Credential = $o365Cred.cred } ; 
            # configure splat for connections: (see above useage)
            # downstream commands
            $pltRXO = [ordered]@{
                Credential = $Credential ;
                verbose = $($VerbosePreference -eq "Continue")  ;
            } ;
            if((get-command Reconnect-EXO).Parameters.keys -contains 'silent'){
                $pltRxo.add('Silent',$silent) ;
            } ;
            # default connectivity cmds - force silent 
            $pltRXOC = [ordered]@{} ; $pltRXO.GetEnumerator() | ?{ $_.Key -notmatch 'silent' }  | ForEach-Object { $pltRXOC.Add($_.Key, $_.Value) } ; $pltRXOC.Add('silent',$silent) ; 
            if((get-command Reconnect-EXO).Parameters.keys -notcontains 'silent'){
                $pltRxo.remove('Silent') ;
            } ; 
            #region EOMREV ; #*------v EOMREV Check v------
            #$EOMmodname = 'ExchangeOnlineManagement' ;
            $pltIMod = @{Name = $EOMmodname ; ErrorAction = 'Stop' ; verbose=$false} ;
            # do a gmo first, faster than gmo -list
            if([version]$EOMMv = (Get-Module @pltIMod).version){}
            elseif([version]$EOMMv = (Get-Module -ListAvailable @pltIMod).version){}
            else {
                $smsg = "$($EOMmodname) PowerShell v$($MinNoWinRMVersion) module is required, do you want to install it?" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Prompt }
                else{ $smsg = "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                $bRet = Read-Host "Enter YYY to continue. Anything else will exit"  ;
                if ($bRet.ToUpper() -eq "YYY") {
                    $smsg = "Installing $($EOMmodname) module..." ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    Install-Module $EOMmodname -Repository PSGallery -AllowClobber -Force ;
                } else {
                    $smsg = "Please install $($EOMmodname) PowerShell v$($MinNoWinRMVersion)  module." ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN }
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    #exit 1
                    break ;
                }  ;
            } ;
            $smsg = "(Checking for WinRM support in this EOM rev...)" ; 
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            if([version]$EOMMv -ge [version]$MinNoWinRMVersion){
                $MinNoWinRMVersion = $EOMMv.tostring() ;
                $IsNoWinRM = $true ;
            }elseif([version]$EOMMv -lt [version]$MinimumVersion){
                $smsg = "Installed $($EOMmodname) is v$($MinNoWinRMVersion): This module is obsolete!" ;
                $smsg += "`nAnd unsupported by this function!" ;
                $smsg += "`nPlease install $($EOMmodname) PowerShell v$($MinNoWinRMVersion)  module!" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                Break ;
            } else {
                $IsNoWinRM = $false ;
            } ;
            [boolean]$UseConnEXO = [boolean]([version]$EOMMv -ge [version]$MinNoWinRMVersion) ;
            #endregion EOMREV ; #*------^ END EOMREV Check  ^------
            #-=-=-=-=-=-=-=-=
            <### CALLS ARE IN FORM: (cred$($tenorg))
            # downstream commands
            $pltRXO = @{
                Credential = (Get-Variable -name cred$($tenorg) ).value ;
                #verbose = $($verbose) ;
                Verbose = $FALSE ; 
            } ;
            if((get-command Reconnect-EXO).Parameters.keys -contains 'silent'){
                $pltRxo.add('Silent',$false) ;
            } ; 
            # default connectivity cmds - force silent false
            $pltRXOC = [ordered]@{} ; $pltRXO.GetEnumerator() | ?{ $_.Key -notmatch 'silent' }  | ForEach-Object { $pltRXOC.Add($_.Key, $_.Value) } ; $pltRXOC.Add('silent',$true) ;
            if((get-command Reconnect-EXO).Parameters.keys -notcontains 'silent'){
                $pltRxo.remove('Silent') ;
            } ; 
            #$pltRXO creds & .username can also be used for AzureAD connections:
            #Connect-AAD @pltRXOC ;
            ###>
            #endregion GENERIC_EXO_CREDS_&_SVC_CONN #*------^ END GENERIC EXO CREDS & SVC CONN BP ^------

        } else {
            $smsg = "(`$useO365:$($useO365))" ; 
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        } ; # if-E if($useO365 ){
        #endregion useO365 ; #*------^ END useO365 ^------

        #region useEXO ; #*------v useEXO v------
        # 1:29 PM 9/15/2022 as of MFA & v205, have to load EXO *before* any EXOP, or gen get-steppablepipeline suffix conflict error
        if($useEXO){
            if ($script:useEXOv2 -OR $useEXOv2) { reconnect-eXO2 @pltRXOC }
            else { reconnect-EXO @pltRXOC } ;
        } else {
            $smsg = "(`$useEXO:$($useEXO))" ; 
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        } ; # if-E 
        #endregion  ; #*------^ END useEXO ^------
      
        #region GENERIC_EXOP_CREDS_&_SRVR_CONN #*------v GENERIC EXOP CREDS & SRVR CONN BP v------
        # steer all onprem code on $XXXMeta.ExOPAccessFromToro & Ex10Server values
        #$UseOP=$true ; 
        #$UseExOP=$true ;
        #$useExopNoDep = $true # switch to use Connect-ExchangeServerTDO, vs connect-ex2010 (creds are assumed inherent to the account) 
        #$useForestWide = $true ; # flag to trigger cross-domain/forest-wide code in AD & EXoP
        <# no onprem dep
        if((Get-Variable  -name "$($TenOrg)Meta").value.ExOPAccessFromToro -AND (Get-Variable  -name "$($TenOrg)Meta").value.Ex10Server){
            $UseOP = $UseExOP = $true ;
            $smsg = "$($TenOrg):Meta.ExOPAccessFromToro($((Get-Variable  -name "$($TenOrg)Meta").value.ExOPAccessFromToro)) -AND/OR Meta.Ex10Server($((Get-Variable  -name "$($TenOrg)Meta").value.Ex10Server)),`ENABLING use of OnPrem Ex system this pass." ;
            if($verbose){ if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
        } else {
            $UseOP = $UseExOP = $false ;
            $smsg = "$($TenOrg):Meta.ExOPAccessFromToro($((Get-Variable  -name "$($TenOrg)Meta").value.ExOPAccessFromToro)) -AND/OR Meta.Ex10Server($((Get-Variable  -name "$($TenOrg)Meta").value.Ex10Server)),`nDISABLING use of OnPrem Ex system this pass." ;
            if($verbose){ if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
        } ;
        #>
        if($UseOP){
            if($useExopNoDep){
                # Connect-ExchangeServerTDO use: creds are implied from the PSSession creds; assumed to have EXOP perms
            } else {
                #*------v GENERIC EXOP CREDS & SRVR CONN BP v------
                # do the OP creds too
                $OPCred=$null ;
                # default to the onprem svc acct
                # userrole='ESVC','SID'
                #$pltGHOpCred=@{TenOrg=$TenOrg ;userrole='ESVC','SID'; verbose=$($verbose)} ;
                # userrole='SID','ESVC'
                $pltGHOpCred=@{TenOrg=$TenOrg ;userrole='SID','ESVC'; verbose=$($verbose)} ;
                $smsg = "get-HybridOPCredentials w`n$(($pltGHOpCred|out-string).trim())" ; 
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level verbose } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                if($OPCred=(get-HybridOPCredentials @pltGHOpCred).cred){
                    # make it script scope, so we don't have to predetect & purge before using new-variable
                    if(get-Variable -Name "cred$($tenorg)OP" -scope Script -ea 0 ){ remove-Variable -Name "cred$($tenorg)OP" -scope Script } ;
                    New-Variable -Name "cred$($tenorg)OP" -scope Script -Value $OPCred ;
                    $smsg = "Resolved $($Tenorg) `$OPCred:$($OPCred.username) (assigned to `$cred$($tenorg)OP)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                } else {
                    $statusdelta = ";ERROR"; # CHANGE|INCOMPLETE|ERROR|WARN|FAIL ;
                    $script:PassStatus += $statusdelta ;
                    set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta) ;
                    $smsg = "Unable to resolve get-HybridOPCredentials -TenOrg $($TenOrg) -userrole 'ESVC' value!"
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    throw "Unable to resolve $($tenorg) `$OPCred value!`nEXIT!"
                    Break ;
                } ;
                $smsg= "Using OnPrem/EXOP cred:`$cred$($tenorg)OP:$((Get-Variable -name "cred$($tenorg)OP" ).value.username)" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                                                            <### CALLS ARE IN FORM: (cred$($tenorg))
                $pltRX10 = @{
                    Credential = (Get-Variable -name "cred$($tenorg)OP" ).value ;
                    #verbose = $($verbose) ;
                    Verbose = $FALSE ; 
                } ;
                $1stConn = $false ; # below uses silent suppr for both x10 & xo!
                if($1stConn){
                    $pltRX10.silent = $pltRXO.silent = $false ;
                } else {
                    $pltRX10.silent = $pltRXO.silent =$true ;
                } ;
                if($pltRX10){ReConnect-Ex2010 @pltRX10 }
                else {ReConnect-Ex2010 }
                #$pltRx10 creds & .username can also be used for local ADMS connections
                ###>
                $pltRX10 = @{
                    Credential = (Get-Variable -name "cred$($tenorg)OP" ).value ;
                    #verbose = $($verbose) ;
                    Verbose = $FALSE ; 
                } ;
                if((get-command Reconnect-Ex2010).Parameters.keys -contains 'silent'){
                    $pltRX10.add('Silent',$false) ;
                } ;
            } ; 
            # defer cx10/rx10, until just before get-recipients qry
            # connect to ExOP X10
            if($useEXOP){
                if($useExopNoDep){ 
                    $smsg = "(Using ExOP:Connect-ExchangeServerTDO())" ; 
                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;           
                    TRY{
                        $Site=[System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::GetComputerSite().Name 
                    }CATCH{$Site=$env:COMPUTERNAME} ;
                    $PSSession = Connect-ExchangeServerTDO -siteName $Site -RoleNames @('HUB','CAS') -verbose ; 
                } else {
                    if($pltRX10){
                        #ReConnect-Ex2010XO @pltRX10 ;
                        ReConnect-Ex2010 @pltRX10 ;
                    } else { Reconnect-Ex2010 ; } ;
                    #Add-PSSnapin -Name 'Microsoft.Exchange.Management.PowerShell.SnapIn'
                    #TK: add: test Exch & AD functional connections
                    TRY{
                        if(get-command -module (get-module |?{$_.name -like 'tmp_*'}).name -name 'get-OrganizationConfig'){} else {
                            $smsg = "(mangled Ex10 conn: dx10,rx10...)" ; 
                            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                            disconnect-ex2010 ; reconnect-ex2010 ; 
                        } ; 
                        if(-not ($OrgName = ((get-OrganizationConfig).DistinguishedName.split(',') |?{$_ -like 'DC=*'}) -join '.' -replace 'DC=','')){
                            $smsg = "Missing Exchange Connection! (no (Get-OrganizationConfig).name returned)" ; 
                            throw $smsg ; 
                            $smsg | write-warning  ; 
                        } ; 
                    } CATCH {
                        $ErrTrapd=$Error[0] ;
                        $smsg = $ErrTrapd ;
                        $smsg += "`n";
                        $smsg += $ErrTrapd.Exception.Message ;
                        if ($logging) { _write-log -LogContent $smsg -Path $logfile -useHost -Level WARN } 
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                        CONTINUE ;
                    } ;
                }
            } ; 
            if($useForestWide){
                #region  ; #*------v OPTIONAL CODE TO ENABLE FOREST-WIDE NATIVE EXCHANGE SUPPORT v------
                $smsg = "(`$useForestWide:$($useForestWide)):Enabling EXoP Forestwide)" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                $smsg = 'Set-AdServerSettings -ViewEntireForest `$True' ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                Set-AdServerSettings -ViewEntireForest $True ;
                #endregion  ; #*------^ END OPTIONAL CODE TO ENABLE FOREST-WIDE NATIVE EXCHANGE SUPPORT ^------
            } ;
        } else {
            $smsg = "(`$useOP:$($UseOP))" ; 
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        }  ;  # if-E $UseOP
        #endregion GENERIC_EXOP_CREDS_&_SRVR_CONN #*------^ END GENERIC EXOP CREDS & SRVR CONN BP ^------
            
            
        #region UseOPAD #*------v UseOPAD v------
        if($UseOP -OR $UseOPAD){
            #region GENERIC_ADMS_CONN_&_XO #*------v GENERIC ADMS CONN & XO  v------
            $smsg = "(loading ADMS...)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            # always capture load-adms return, it outputs a $true to pipeline on success
            $ADMTLoaded = load-ADMS -Verbose:$FALSE ;
            # 9:32 AM 4/20/2023 trimmed disabled/fw-borked cross-org code
            TRY {
                if(-not(Get-ADDomain  -ea STOP).DNSRoot){
                    $smsg = "Missing AD Connection! (no (Get-ADDomain).DNSRoot returned)" ; 
                    throw $smsg ; 
                    $smsg | write-warning  ; 
                } ; 
                $objforest = get-adforest -ea STOP ; 
                # Default new UPNSuffix to the UPNSuffix that matches last 2 elements of the forestname.
                $forestdom = $UPNSuffixDefault = $objforest.UPNSuffixes | ?{$_ -eq (($objforest.name.split('.'))[-2..-1] -join '.')} ; 
                if($useForestWide){
                    #region  ; #*------v OPTIONAL CODE TO ENABLE FOREST-WIDE AD GC QRY SUPPORT v------
                    $smsg = "(`$useForestWide:$($useForestWide)):Enabling AD Forestwide)" ; 
                    if($silent){}elseif ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    $smsg = 'Set-AdServerSettings -ViewEntireForest `$True' ;
                    if($silent){}elseif ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    #TK 9:44 AM 10/6/2022 need org wide for rolegrps in parent dom (only for onprem RBAC, not EXO)
                    $GcFwide = "$((Get-ADDomainController -Discover -Service GlobalCatalog).hostname):3268" ;        
                    #endregion  ; #*------^ END  OPTIONAL CODE TO ENABLE FOREST-WIDE AD GC QRY SUPPORT  ^------
                } ;    
            } CATCH {
                $ErrTrapd=$Error[0] ;
                $smsg = $ErrTrapd ;
                $smsg += "`n";
                $smsg += $ErrTrapd.Exception.Message ;
                if ($logging) { _write-log -LogContent $smsg -Path $logfile -useHost -Level WARN } 
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                CONTINUE ;
            } ;        
            #endregion GENERIC_ADMS_CONN_&_XO #*------^ END GENERIC ADMS CONN & XO ^------
        } else {
            $smsg = "(`$UseOP:$($UseOP))" ; 
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        }  ;
        #if (!$domaincontroller) { $domaincontroller = get-gcfast } ;
        #if(!$domaincontroller){ if(test-path function:get-gcfast){$domaincontroller = get-gcfast} else { throw "no get-gcfast()!" } ;} else {"(existing `$domaincontroller:$($domaincontroller))"} ;
        # use new get-GCFastXO cross-org dc finde
        # default to Op_ExADRoot forest from $TenOrg Meta
        #if($UseOP -AND -not $domaincontroller){
        if($UseOP -AND -not (get-variable domaincontroller -ea 0)){
            #$domaincontroller = get-GCFastXO -TenOrg $TenOrg -subdomain ((get-variable -name "$($TenOrg)Meta").value['OP_ExADRoot']) -verbose:$($verbose) |?{$_.length};
            # need to debug the above, credential issue?
            # just get it done
            $domaincontroller = get-GCFast
        }  else { 
            # have to defer to get-azuread, or use EXO's native cmds to poll grp members
            # TODO 1/15/2021
            $useEXOforGroups = $true ; 
            $smsg = "$($TenOrg):HAS NO ON-PREM ACTIVEDIRECTORY, DEFERRING ALL GROUP ACCESS & MGMT TO NATIVE EXO CMDS!" ; 
            if($silent){}elseif ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        } ; 
        if($useForestWide -AND -not $GcFwide){
            #region  ; #*------v OPTIONAL CODE TO ENABLE FOREST-WIDE ACTIVEDIRECTORY SUPPORT: v------
            $smsg = "`$GcFwide = Get-ADDomainController -Discover -Service GlobalCatalog" ;
            if($silent){}elseif ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            $GcFwide = "$((Get-ADDomainController -Discover -Service GlobalCatalog).hostname):3268" ;
            $smsg = "Discovered `$GcFwide:$($GcFwide)" ; 
            if($silent){}elseif ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            #endregion  ; #*------^ END OPTIONAL CODE TO ENABLE FOREST-WIDE ACTIVEDIRECTORY SUPPORT ^------
        } ;
        #endregion UseOPAD #*------^ END UseOPAD ^------

        #region MSOL_CONNECTION ; #*------v  MSOL CONNECTION v------
        #$UseMSOL = $false 
        if($UseMSOL){
            #$reqMods += "connect-msol".split(";") ;
            #if ( !(check-ReqMods $reqMods) ) { write-error "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Missing function. EXITING." ; Break ; }  ;
            $smsg = "(loading MSOL...)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } Error|Warn|Debug
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            #connect-msol ;
            connect-msol @pltRXOC ;
        } else {
            $smsg = "(`$UseMSOL:$($UseMSOL))" ; 
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        } ;
        #endregion MSOL_CONNECTION ; #*------^  MSOL CONNECTION ^------

        #region AZUREAD_CONNECTION ; #*------v AZUREAD CONNECTION v------
        #$UseAAD = $false 
        if($UseAAD){
            #$reqMods += "Connect-AAD".split(";") ;
            #if ( !(check-ReqMods $reqMods) ) { write-error "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Missing function. EXITING." ; Break ; }  ;
            $smsg = "(loading AAD...)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            Connect-AAD @pltRXOC ;
        } else {
            $smsg = "(`$UseAAD:$($UseAAD))" ; 
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        } ;
        #endregion AZUREAD_CONNECTION ; #*------^ AZUREAD CONNECTION ^------
      
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
        if ($script:useEXOv2) { reconnect-eXO2 @pltRXOC }
        else { reconnect-EXO @pltRXOC } ;
        # reenable VerbosePreference:Continue, if set, during mod loads
        if($VerbosePrefPrior -eq "Continue"){
            $VerbosePreference = $VerbosePrefPrior ;
            $verbose = ($VerbosePreference -eq "Continue") ;
        } ;
        #>
        #-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
        #endregion SERVICE_CONNECTIONS #*======^ END SERVICE_CONNECTIONS ^======

        # Configure the Get-xoMessageTrace splat 
        $pltGXMT=[ordered]@{
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
        if ($PSCmdlet.ParameterSetName -eq 'Dates') {
            if($EndDate -and -not $StartDate){
                $StartDate = (get-date $EndDate).addDays(-1 * $DaysLimit) ; 
            } ; 
            if($StartDate -and -not ($EndDate)){
                $smsg = "(StartDate w *NO* Enddate, asserting currenttime)" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level PROMPT } 
                else{ write-host -foregroundcolor YELLOW "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                $EndDate=(get-date) ;
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
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

        # convert dates to GMT .ToUniversalTime(
        $StartDate = ([datetime]$StartDate).ToUniversalTime() ; 
        $EndDate = ([datetime]$EndDate).ToUniversalTime() ; 
        $StartLocal = ([datetime]$StartDate).ToLocalTime() ; 
        $EndLocal = ([datetime]$EndDate).ToLocalTime() ; 
        
        # sanity test the start/end dates, just in case (won't throw an error in gxmt)
        if($StartDate -gt $EndDate){
            $smsg = "`-StartDate:$($StartDate) is GREATER THAN -EndDate:($EndDate)!" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            throw $smsg ; 
            break ; 
        } ; 

        $smsg = "`$StartDate:$(get-date -Date $StartLocal -format $sFulltimeStamp )" ;
        $smsg += "`n`$EndDate:$(get-date -Date $EndLocal -format $sFulltimeStamp )" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 

        if((New-TimeSpan -Start $StartDate -End (get-date)).days -gt $DaysLimit){
            $smsg = "Search span (between -StartDate & -EndDate, or- Days in use) *exceeds* MS supported days history limit!`nReduce the window below a historical 10d, or use get-HistoricalSearch instead!" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug 
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
            Break ; 
        } ; 

        TRY{
            $tendoms=Get-AzureADDomain ; 
            $Ten = ($tendoms |?{$_.name -like '*.mail.onmicrosoft.com'}).name.split('.')[0] ;
            $Ten = "$($Ten.substring(0,1).toupper())$($Ten.substring(1,$Ten.length-1).toLower())"
        }CATCH{
            $smsg = "NOT AAD CONNECTED!" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
          BREAK 
        } ;
       

    }  # BEG-E
    PROCESS {

        <#
        # default StartDate to -10 can't do more
        $pltGXMT=[ordered]@{
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
        #>

        <# #-=-=-=-=-=-=-=-=
        [Parameter(Mandatory=$false,HelpMessage="Ticket [-ticket 999999]")]
            [ValidateNotNullOrEmpty()]
            [string]$ticket,
        [Parameter(HelpMessage="Tag string that is used for Variables name construction. [-Tag 'LastDDGSend']")]
            [string]$Tag,
        [Parameter(HelpMessage="SenderAddress (an array runs search on each)[-SenderAddress addr@domain.com]")]
            [Alias('Sender')]
            [string[]]$SenderAddress, # MultiValuedProperty
        [Parameter(HelpMessage="RecipientAddress (an array runs search on each)[-RecipientAddress addr@domain.com]")]
            [Alias('Recipients')]
            [string[]]$RecipientAddress, # MultiValuedProperty
        [Parameter(ParameterSetName='Dates',HelpMessage="Start of range to be searched[-StartDate '11/5/2021 2:16 PM']")]
            [Alias('Start')]
            [DateTime]$StartDate,
        [Parameter(ParameterSetName='Dates',HelpMessage="End of range to be searched (defaults to current time if unspecified)[-EndDate '11/5/2021 5:16 PM']")]
            [Alias('End')]
            [DateTime]$EndDate=(get-date),
        [Parameter(ParameterSetName='Days',HelpMessage="Days to be searched, back from current time(Alt to use of StartDate & EndDate; Note:MS won't search -gt 10 days)[-Days 7]")]
            #[ValidateRange(0,[int]::MaxValue)]
            [ValidateRange(0,10)] # MS won't search beyond 10, and silently returns incomplete results
            [int]$Days,
        [Parameter(HelpMessage="Subject of target message (emulated via post filtering, not supported param of Get-xoMessageTrace) [-Subject 'Some subject']")]
            [Alias('MessageSubject')]
            [string]$subject,
        [Parameter(HelpMessage="Subject of target message [-Subject 'Some subject']")]
            [Alias('DeliveryStatus','EventId')]
            [ValidateSet('None','GettingStatus','Failed','Pending','Delivered','Expanded','Quarantined','FilteredAsSpam')]
            [string[]]$Status, # MultiValuedProperty
        [Parameter(HelpMessage="MessageId of target message(s) (include any <> and enclose in quotes; an array runs search on each)[-MessageId '<nnnn-nn.xxx....outlook.com>']")]
            # Get-xoMessageTrace specs <MultiValuedProperty>: "just means that you can provide multiple values (i.e. an array) as the argument to the parameter. If your users input something like alice@example.com,bob@example.com,charlie@example.com, you need to split the delims"
            [string[]]$MessageId, # MultiValuedProperty
        [Parameter(HelpMessage="MessageTraceId of target message [-MessageTraceId '[MessageTraceId string]']")]
            [Guid]$MessageTraceId,
        [Parameter(HelpMessage="The FromIP parameter filters the results by the source IP address. For incoming messages, the value of FromIP is the public IP address of the SMTP email server that sent the message. For outgoing messages from Exchange Online, the value is blank. [-FromIP '123.456.789.012']")]
            [string]$FromIP,
        [Parameter(HelpMessage="The ToIP parameter filters the results by the destination IP address. For outgoing messages, the value of ToIP is the public IP address in the resolved MX record for the destination domain. For incoming messages to Exchange Online, the value is blank. [-ToIP '123.456.789.012']")]
            [string]$ToIP,
        [Parameter(HelpMessage="switch to do Summarize & Expansion of any MTD TransportRule events (defaults true) [-DetailedReportRuleHits]")]
            [switch]$DetailedReportRuleHits= $true,
        [Parameter(HelpMessage="Integer number of maximum messages to be follow-up MessageTraceDetail'd [-MessageTraceDetailLimit 20]")]
            [int]$MessageTraceDetailLimit = 100,
        [Parameter(HelpMessage="switch to do Summarize & Expansion of any MTD TransportRule events (defaults true) [-DetailedReportRuleHits]")]
            [switch]$NoQuarCheck,
        [Parameter(HelpMessage="switch to perform MessageTrackingDetail pass, after intial MessageTrace (up to limti specified in -MessageTraceDetailLimit (defaults true) [-Detailed]")]
            [switch]$Detailed=$true
        #-=-=-=-=-=-=-=-=
        #>
        # 1:00 PM 11/20/2024 note, all the $ofile building until #1349:[string[]]$ofile=@() ; is a waste of time, it gets rebuilt at bottom; rem them all
        if($SenderAddress){
            if($SenderAddress -match '\*'){
                # To do wildcards (*@DOMAIN.COM), SPEC THE ADDRESS LIKE: -SenderAddress @('*@DOMAIN.COM') (forces as array)
                $pltGXMT.add('SenderAddress',@(($SenderAddress -split ' *, *')) ) ;
            }else{
                $pltGXMT.add('SenderAddress',($SenderAddress -split ' *, *')) ;
            } ; 
        } ;
        if($RecipientAddress){
            if($RecipientAddress -match '\*'){
                # To do wildcards (*@DOMAIN.COM), SPEC THE ADDRESS LIKE: -RecipientAddress @('*@DOMAIN.COM') (forces as array)
                $pltGXMT.add('RecipientAddress',@(($RecipientAddress -split ' *, *')) ) ;
            }else{
                $pltGXMT.add('RecipientAddress',($RecipientAddress -split ' *, *')) ;
            } ; 
        } ;
        if($StartDate){
            $pltGXMT.add('StartDate',$StartDate) ; 
        } ;
        if($EndDate){
            $pltGXMT.add('EndDate',$EndDate) ; 
        } ;
        if($Status){
            $pltGXMT.add('Status',($Status -split ' *, *')) ; 
        } ;
        if($MessageId){
            $pltGXMT.add('MessageId',($MessageId -split ' *, *')) ; 
        } ;
        if($MessageTraceId){
            $pltGXMT.add('MessageTraceId',$MessageTraceId) ; 
        } ;
        if($FromIP){
            $pltGXMT.add('FromIP',$FromIP) ; 
        } ;
        if($ToIP){
            $pltGXMT.add('ToIP',$ToIP) ; 
        } ;

        if($subject){
        } ;
        
        # use the updated psOfile build:
        #-=-=-=-=-=-=-=-=
        #region MSGTRKFILENAME ; #*------v MSGTRKFILENAME v------
        write-verbose "Keys off of typical msgtrk inputsplat" ; 
        <#
        $pltI=@{   ticket=$ticket ;
           Requestor=$requestor ;
           days=0 ;
           StartDate=$TargetMsg.Received.Adddays(-1) ;
           EndDate=$TargetMsg.Received.Adddays(+1) ;
           Sender="" ;
           Recipients=$TargetMsg.RecipientAddress ;
           Status='' ;
           MessageSubject="" ;
           MessageTraceId='' ;
           MessageId=$TargetMsg.MessageId ;
           FromIP='' ;
           NoQuarCheck='';
           Tag='LatestDDG' ;
         }   ;
          #>
        # default create a \logs\ dir below script dir
        $LogPath = split-path $logfile ; 
        $smsg = "Writing export files to discovered `$LogPath: $($LogPath)" ; 
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

        if (-not (test-path $LogPath )){mkdir $LogPath -verbose  }
        [string[]]$ofile=@() ; 
        write-verbose "Add comma-delimited elements" ; 
        #$ofile+=if($ticket -AND $Tag){@($ticket,$tag) -join '_'}else{$ticket} ;
        $ofile+= (@($ticket,$tag) | ?{$_}) -join '_' ; 
        $ofile+= (@($Ten,$Requestor,'EXOMsgTrc') | ?{$_} ) -join '-' ;
        $ofile+=if($SenderAddress){
            #"FROM_$((($SenderAddress | select -first 2) -join ',').replace('*','ANY'))"
            "FROM_$(( ($SenderAddress| select -first 2) -join ',').replace('*','ANY'))"
        }else{''} ;
        $ofile+=if($RecipientAddress){
            "TO_$(( ($RecipientAddress| select -first 2) -join ',').replace('*','ANY'))"
        }else{''} ;
        $ofile+=if($MessageId){
            #"MSGID_$($MessageId.replace('<','').replace('>',''))"
            if($MessageId -is [array]){
                "MSGID_$($MessageId[0] -replace '[\<\>]','')..."
            } else { 
                "MSGID_$($MessageId -replace '[\<\>]','')"
            } ; 
        }else{''} ;
        $ofile+=if($MessageTraceId){"MsgId_$($MessageTraceId)"}else{''} ;
        $ofile+=if($FromIP){"FIP_$($FromIP)"}else{''} ;
        $ofile+=if($MessageSubject){"SUBJ_$($MessageSubject.substring(0,[System.Math]::Min(10,$MessageSubject.Length)))..."}else{''} ;
        $ofile+=if($Status){
            "STATUS_$($Status -join ',')"
        }else{''} ;
        write-verbose "comma join the non-empty elements" ; 
        [string[]]$ofile=($ofile |  ?{$_} ) -join ',' ; 
        write-verbose "add the dash-delimited elements" ; 
        $ofile+=if($days){"$($days)d"}else{''} ;
        $ofile+=if($StartDate){"$(get-date $StartDate -format 'yyyyMMdd-HHmm')"}else{''} ;
        $ofile+=if($EndDate){$ofile+= "$(get-date $EndDate -format 'yyyyMMdd-HHmm')"}else{''} ;
        $ofile+=if($MessageSubject){"Subj_$($MessageSubject.replace("*"," ").replace("\"," "))"}else{''} ;
        $ofile+="run$(get-date -format 'yyyyMMdd-HHmm').csv" ;
        write-verbose "dash-join non-empty elems" ; 
        [string]$ofile=($ofile |  ?{$_} ) -join '-' ; 
        write-verbose "replace filesys illegal chars" ; 
        [string]$ofile=[RegEx]::Replace($ofile, "[{0}]" -f ([RegEx]::Escape(-join [System.IO.Path]::GetInvalidFileNameChars())), '') ;
        if($LogPath){
            write-verbose "add configured `LogPath" ; 
            $ofile = join-path $LogPath $ofile ; 
        } else { 
            write-verbose "add relative path" ; 
            $ofile=".\logs\$($ofile)" ;
        } ; 
        #$MSGSTRK | export-csv -noty $ofile -verbo ; 
        #write-host -foregroundcolor green "export-csv'd to:`n$((resolve-path $ofile).path)" ; 
        #endregion MSGTRKFILENAME ; #*------^ END MSGTRKFILENAME ^------
        #-=-=-=-=-=-=-=-=


        #$ofile=[RegEx]::Replace($ofile, "[{0}]" -f ([RegEx]::Escape(-join [System.IO.Path]::GetInvalidFileNameChars())), '') ;
        # use the tested redirected $logfile path
        #$ofile = join-path -path (split-path $logfile) -ChildPath $ofile ; 
        $hReports = [ordered]@{} ; 
        #rxo ;
        $error.clear() ;
        TRY {
            # prepurge empty hash value keys:
            #$pltGXMT=$pltGXMT.GetEnumerator()|? value ;
            # remove null keyed objects
            #$pltGXMT | Foreach {$p = $_ ;@($p.GetEnumerator()) | ?{ ($_.Value | Out-String).length -eq 0 } | Foreach-Object {$p.Remove($_.Key)} ;} ;
            # skip it, we're only adding populated items now
            #write-verbose "hashtype:$($pltGXMT.GetType().FullName)" ; 
            # and issue was first untested negative integer -Days; and 2nd GMT window for start/enddate, so the 'local' input needs to be converted to/from gmt to get the targeted content.

            $smsg = "Get-xoMessageTrace  w`n$(($pltGXMT|out-string).trim())" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-host "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
            $Page = 1  ;
            $Msgs=$null ;
            do {
                $smsg = "Collecting - Page $($Page)..."  ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                $pltGXMT.Page=$Page ;
                $PageMsgs = Get-xoMessageTrace @pltGXMT |  ?{$_.SenderAddress -notlike '*micro*' -or $_.SenderAddress -notlike '*root*' }  ;
                $Page++  ;
                $Msgs += @($PageMsgs)  ;
            } until ($PageMsgs -eq $null) ;
            $Msgs=$Msgs| Sort Received ;
            $smsg = "Raw sender/recipient events:$(($Msgs|measure).Count)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-host "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 

            if($subject){
                $smsg = "Post-Filtering on Subject:$($subject)" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                # detect whether to filter on -match (regex) or -like (asterisk, or default non-regex)
                if(test-IsRegexPattern -string $subject -verbose:$($VerbosePreference -eq "Continue")){
                    $smsg = "(detected -subject as regex - using -match comparison)" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    $MsgsFltrd = $Msgs | ?{$_.Subject -match $subject} ;
                    if(-not $MsgsFltrd){
                        $smsg = "Subject: regex -match comparison *FAILED* to return matches`nretrying Subject filter as -Like..." ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug 
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        $MsgsFltrd = $Msgs | ?{$_.Subject -like $subject} ;
                    } ; 
                } else { 
                    $smsg = "(detected -subject as NON-regex - using -like comparison)" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
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
            
            if($Msgs){
                # reselect with local time variant
                $Msgs = $Msgs | select $propsMTAll ; 
                if($DoExports){
                    $smsg = "($(($Msgs|measure).count) events | export-csv $($ofile))" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    TRY{
                        $Msgs | select $propsMT | export-csv -notype -path $ofile -ea STOP  ;
                        $smsg = "export-csv'd to:`n$((resolve-path $ofile).path)" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } CATCH {
                        $ErrTrapd=$Error[0] ;
                        $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                        else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ; 
                    $smsg = "(adding `$hReports.MTMessages)" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    
                    # add the csvfilename
                    $smsg = "(adding `$hReports.MTMessagesCSVFile:$($ofile))" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    $hReports.add('MTMessagesCSVFile',$ofile) ; 
                } 

                $hReports.add('MTMessages',$msgs) ; 

                if($Msgs){
                    $smsg = "`n#*------v Status DISTRIB v------`n`n" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Prompt } #Error|Warn|Debug 
                    else{ write-host -foregroundcolor blue "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                    #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success

                    $hReports.add('StatusHisto',($Msgs | select -expand Status | group | sort count,count -desc | select count,name)) ;

                    $smsg = "`n$(($hReports.StatusHisto|out-string).trim())" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-host "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                    $smsg = "`n`n#*------^ Status DISTRIB ^------`n`n" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Prompt } #Error|Warn|Debug 
                    else{ write-host -foregroundcolor blue "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                    $smsg = "`n`n#*------v MOST RECENT MATCH v------`n`n" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Prompt } #Error|Warn|Debug 
                    else{ write-host -foregroundcolor white "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 

                    $hReports.add('MsgLast',($msgs[-1]| fl $propsMsgDump)) ;
                    $smsg = "`n$(($hReports.MsgLast |out-string).trim())";
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-host "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                    $smsg = "`n`n#*------^ MOST RECENT MATCH ^------`n`n" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Prompt } #Error|Warn|Debug 
                    else{ write-host -foregroundcolor white "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 

                } ; 

                if($mFails = $msgs | ?{$_.status -eq 'Failed'} | select -first $MessageTraceDetailLimit){
                    $smsg = "Expanded analysis on first $($MessageTraceDetailLimit) Status:Failed messages..." ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 

                    if($mFails | ?{$_.Subject -notmatch '^Recall:\s' -AND $_.Subject -notmatch '^Automatic\sreply:\s'}){
                        $smsg = "Other Fails detected: Opening ExoP & ADMS connections..." ; 
                        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                        #region CXOP ; #*------v CXOP v------
                        $smsg = "Resolve ComputerSite..." ;
                        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                        TRY{
                            $Site=[System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::GetComputerSite().Name
                        }CATCH{
                            $Site=$env:COMPUTERNAME ;
                            $smsg = "Non-AD-Connected system, setting `$Site:`$env:COMPUTERNAME" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Prompt }
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        } ;
                        $smsg = "Resolved ComputerSite: $($Site)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Prompt }
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        $smsg = "Discovering and connecting to a local Exchange server in local AD Site"  ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Prompt }
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        $pltNPSS=[ordered]@{
                            siteName = $Site ;
                            RoleNames = @('HUB','CAS') ;
                            Verbose = ($PSBoundParameters['Verbose'] -eq $true) ;
                        } ;
                        $smsg = "Connect-ExchangeServerTDO w`n$(($pltNPSS|out-string).trim())" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        $PSSession = Connect-ExchangeServerTDO @pltNPSS ;
                        #endregion CXOP ; #*------^ END CXOP ^------
                        #region loadADMS ; #*------v loadADMS v------
                        if(-not (get-command -name get-aduser -ea 0)){
                            $env:ADPS_LoadDefaultDrive = 0 ; $sName="ActiveDirectory"; if (!(Get-Module | where {$_.Name -eq $sName})) {Import-Module $sName -ea Stop}
                        } ; 
                        #endregion loadADMS ; #*------^ END loadADMS ^------
                    } ; 
                    
                    $FailAggr = @() ;                                         
                    foreach($failed in $mFails){
                        # 'RunspaceId',
                        # $propsMTAll = 'Organization','MessageId','Received', @{N='ReceivedLocal';E={[datetime]$_.Received.ToLocalTime()}},'SenderAddress','RecipientAddress','Subject','Status','ToIP','FromIP','Size','MessageTraceId','StartDate','EndDate','Index'

                        if($host.version.major -ge 3){$FailMsgSummary=[ordered]@{Dummy = $null ;} }
                        else {$FailMsgSummary=@{Dummy = $null ;} ;}
                        If($FailMsgSummary.Contains("Dummy")){$FailMsgSummary.remove("Dummy")} ;
                        $fieldsnull = 'Organization','MessageId','Received','ReceivedLocal','SenderAddress',
                            'RecipientAddress','Subject','Status','ToIP','FromIP','Size','MessageTraceId',
                            'StartDate','EndDate','Index','FailXoRecipientType','FailXopRecipientType',
                            'FailDetailEvent','FailDetailDetail','ADUserTermOU' ;  ; 
                        $fieldsnull | % { $FailMsgSummary.add($_,$null) } ;
                        $fieldsbool = 'isFailed','ADUserDisabled'
                            #,'isFailedOOO','isFailRecall','isFailOther','isFailOtherAcctsBlock',
                            #'isFailSecBlock','isFailMailLoop','isFailBrokenTerm','isFailNoMailbox', ; 
                        # 4:11 PM 11/22/2024 pulled the is*, do it by parsing the FailCode instead
                        $fieldsbool | % { $FailMsgSummary.add($_,$false) } ;


                        <#$FailMsgSummary=[indexed]@{
                            Organization = $failed.Organization ; 
                            MessageId = $failed.MessageId ; 
                            Received = $failed.Received ; 
                            ReceivedLocal = $failed.ReceivedLocal ; 
                            SenderAddress = $failed.SenderAddress ; 
                            RecipientAddress = $failed.RecipientAddress ; 
                            Subject = $failed.Subject ; 
                            Status = $failed.Status ; 
                            ToIP = $failed.ToIP ; 
                            FromIP = $failed.FromIP ; 
                            Size = $failed.Size ; 
                            MessageTraceId = $failed.MessageTraceId ; 
                            StartDate = $failed.StartDate ; 
                            EndDate = $failed.EndDate ; 
                            Index = $failed.Index ;      
                            isFailed = $true ; 
                            isFailedOOO = $false ;
                            isFailRecall = $false ;
                            isFailOther = $false ;
                            isFailOtherAcctsBlock = $false ; 
                            isFailSecBlock = $false ; 
                            isFailMailLoop = $false ;
                            isFailBrokenTerm = $false ; 
                            isFailNoMailbox = $false ; 
                            FailXoRecipientType = $null ; 
                            FailXopRecipientType = $null ;
                            FailDetailEvent = $null ; 
                            FailDetailDetail = $null ; 
                            ADUserDisabled = $false ; 
                            ADUserTermOU = $null ; 
                        } ; 
                        #>
                        $FailMsgSummary.Organization = $failed.Organization ; 
                        $FailMsgSummary.MessageId = $failed.MessageId ; 
                        $FailMsgSummary.Received = $failed.Received ; 
                        $FailMsgSummary.ReceivedLocal = $failed.ReceivedLocal ; 
                        $FailMsgSummary.SenderAddress = $failed.SenderAddress ; 
                        $FailMsgSummary.RecipientAddress = $failed.RecipientAddress ; 
                        $FailMsgSummary.Subject = $failed.Subject ; 
                        $FailMsgSummary.Status = $failed.Status ; 
                        $FailMsgSummary.ToIP = $failed.ToIP ; 
                        $FailMsgSummary.FromIP = $failed.FromIP ; 
                        $FailMsgSummary.Size = $failed.Size ; 
                        $FailMsgSummary.MessageTraceId = $failed.MessageTraceId ; 
                        $FailMsgSummary.StartDate = $failed.StartDate ; 
                        $FailMsgSummary.EndDate = $failed.EndDate ; 
                        $FailMsgSummary.Index = $failed.Index ;      
                        $FailMsgSummary.isFailed = $true ; 
                        <#
                        $FailMsgSummary.isFailedOOO = $false ;
                        $FailMsgSummary.isFailRecall = $false ;
                        $FailMsgSummary.isFailOther = $false ;
                        $FailMsgSummary.isFailOtherAcctsBlock = $false ; 
                        $FailMsgSummary.isFailSecBlock = $false ; 
                        $FailMsgSummary.isFailMailLoop = $false ;
                        $FailMsgSummary.isFailBrokenTerm = $false ; 
                        $FailMsgSummary.isFailNoMailbox = $false ; 
                        #>
                        $FailMsgSummary.FailCode = $null ; 
                        $FailMsgSummary.FailXoRecipientType = $null ; 
                        $FailMsgSummary.FailXopRecipientType = $null ;
                        $FailMsgSummary.FailDetailEvent = $null ; 
                        $FailMsgSummary.FailDetailDetail = $null ; 
                        $FailMsgSummary.ADUserDisabled = $false ; 
                        $FailMsgSummary.ADUserTermOU = $null ;
                        #$rgxFailOOOSubj = '^Automatic\sreply:\s' ; 
                        if($failed | ?{$_.Subject -match $rgxFailOOOSubj}){
                            $FailMsgSummary.FailCode += @('FailOOO') ; 
                        } ; 
                        #$rgxFailRecallSubj = '^Recall:\s' ; 
                        if($failed | ?{$_.Subject -match $rgxFailRecallSubj}){
                            $FailMsgSummary.FailCode += @('FailRecall') ; 
                        } ; 
                        #$rgxFailOtherAcctBlock = 'OtherAccts-External-Mail-Rejection' ; 
                        #$FailOtherAcctBlockExemptionGroup = 'LYN-DL-OPExch-OtherAcctMbxs-ExternalMailOK@toro.com' ; 
                        if($failed | ?{$_.Subject -notmatch $rgxFailRecallSubj -AND $_.Subject -notmatch $rgxFailOOOSubj}){
                            #$FailMsgSummary.isFailOther = $true ; 
                            $FODetail =  $failed | Get-xoMessageTraceDetail -ea STOP; 
                            $FailMsgSummary.FailDetailEvent = $FODetail.event ; 
                            $FailMsgSummary.FailDetailDetail = $FODetail.Detail ; 
                            if($FODetail | ?{$_.event -eq 'Transport rule' -AND $_.Detail -match $rgxFailOtherAcctBlock}){
                                $FailMsgSummary.FailCode += @('FailOtherAcctBlock') ; 
                            } ; 
                            #$rgxFailConfRmExtBlock = 'ConfRm-External-Mail-Rejection' ; 
                            if($FODetail | ?{$_.event -eq 'Transport rule' -AND $_.Detail -match $rgxFailConfRmExtBlock}){
                                $FailMsgSummary.FailCode += @('FailConfRmExtBlock') ; 
                            } ; 
                            #$rgxFailSecBlock = '^Security(\s-\s|-)' ; 
                            if($FODetail | ?{$_.event -eq 'Transport rule' -AND $_.Detail -match $rgxFailSecBlock}){
                                if(($FODetail.detail | select -unique ) -match "Transport\srule:\s'(.*)',"){
                                    $TRule = $matches[0] ; 
                                    $FailMsgSummary.FailCode += @("FailSecBlock:$($Trule)") ; 
                                }else{
                                    $FailMsgSummary.FailCode += @('FailSecBlock') ; 
                                } ; 
                            }
                            if($FODetail | ?{$_.event -eq 'FAIL' -AND $_.Detail -match 'Hop\scount\sexceeded\s-\spossible\smail\sloop'}){
                                $FailMsgSummary.FailCode += @('FailMailLoop') ; 
                                $xopRcp = $xoRCP = $adu = $null ; 
                                $xopRcp = get-recipient $failed.RecipientAddress -ea 0;
                                $xoRCP = get-xorecipient $failed.RecipientAddress -ea 0 ; 
                                $adu = get-aduser -id $xoRCP.alias -ea 0 ; 
                                $FailMsgSummary.FailXoRecipientType = $xoRCP.RecipientTypeDetails  ; 
                                $FailMsgSummary.FailXopRecipientType = $xopRcp.RecipientTypeDetails  ; 
                                if($FailMsgSummary.FailXoRecipientType -eq 'MailUser' -AND $FailMsgSummary.FailXopRecipientType -eq 'RemoteUserMailbox'){
                                    $FailMsgSummary.FailCode += @('FailBrokenTerm') ; 
                                    $FailMsgSummary.FailCode += @('FailNoMailbox') ; 
                                } ; 
                                if($adu.DistinguishedName -match 'OU=(Disabled|TERMedUsers|TERMedUserSharedEmail),'){
                                    $FailMsgSummary.ADUserTermOU = ($adu.DistinguishedName.split(',') | select -skip 1) -join ','  
                                    $FailMsgSummary.FailCode += @('FailBrokenTerm','FailADUserTermOU') ; 
                                } ;
                                if($adu.Enabled -eq $false){
                                    $FailMsgSummary.ADUserDisabled = $true  ; 
                                    $FailMsgSummary.FailCode += @('FailBrokenTerm','FailADUserDisabled') ; 
                                } ;
                            } ; 
                            if($FODetail | ?{$_.event -eq 'FAIL' -AND $_.Detail -match 'Reason:\s'}){
                                # there's a Reason:\s in the mix, try to echo it
                                $FailMsgSummary.FailCode += @('FailReason') ; 
                                $FailMsgSummary.FailDetailDetail = ($FODetail | ?{$_.event -eq 'FAIL' -AND $_.Detail -match 'Reason:\s'}).Detail
                            } ; 
                        } ; 
                        if($FailMsgSummary.FailCode){
                            # reduce to single instance of each code
                            $FailMsgSummary.FailCode = $FailMsgSummary.FailCode | select -unique ;
                        } ; 
                        $FailAggr +=  New-Object -TypeName PsObject -Property $FailMsgSummary
                        #[pscustomobject]$FailMsgSummary ; 
                    } ; 
                    #---

                    $smsg = "adding:`$hReports.MsgsFail" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-host -foregroundcolor yellow "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                    $hReports.add('MsgsFail',$FailAggr) ; 
                    $ofileF = $ofile.replace('-EXOMsgTrc,','FAILMsgs,') ;
                    if($DoExports){
                        TRY{
                            $FailAggr | export-csv -notype -path $ofileF -ea STOP ;
                            $smsg = "export-csv'd to:`n$((resolve-path $ofileF).path)" ; 
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                            else{ write-host -foregroundcolor yellow "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                        } CATCH {
                            $ErrTrapd=$Error[0] ;
                            $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                            else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        } ; 
                    } ; 


                    # divide up the results & report on the types
                    $FailsVariants = $hReports.MsgsFail | group failcode | select -expand name; 

                    $prpFailMsg = 'ReceivedLocal','SenderAddress','RecipientAddress','Subject','Status' ; 

                    # do the single fails
                    $SingleFails = $hReports.MsgsFail | ?{-not ($_.failcode -is [array])} ; 
                    $FailVariantsSingle = $SingleFails | group failcode | select -expand name; 
                    foreach($FV in $FailVariantsSingle){
                        $theseFails = $SingleFails |?{$_.failcode -eq $FV} ; 
                        #$pltWH=get-colorcombo -Rand ;
                        #$Host.UI.RawUI.BackgroundColor = $pltWH.BackgroundColor ;
                        #$Host.UI.RawUI.ForegroundColor = $pltWH.ForegroundColor ;
                        $hsFailRpt = @"

*------v Delivery Status:Failed of type: $($FV) v------

$(($theseFails|ft -a $prpFailMsg  | out-string).trim())

$(
    switch -regex ($FV){
        'FailOtherAcctBlock' {
            "`n$($FV): Blocked by Security-mandated Transport rule:$($rgxFailOtherAcctBlock)"
            "`n -To Exempt from Block: Request SvcDesk add the mailbox to $($FailOtherAcctBlockExemptionGroup) group"
            "`n -To suppress DDG Membership for converted SharedMailboxes (to UserMailbox, w logon): set CustomAttribute4: DL-Exclude"
        }
        'FailConfRmExtBlock'{
            "`n$($FV): Blocked by Security-mandated Transport rule:$($rgxFailConfRmExtBlock)`n(Firm mandate: No exemption permitted for ResourceMailboxes)`n"
        }
        'FailOOO'{
            "`n$($FV): Blocked by Security Policy: Blocked external delivery of Out-Of-Office messages`n(Global Security policy mandate: No exemption permitted)`n"
        }
        'FailRecall' {
            "`n$($FV): Expected fail: Sender issued Outlook Recall of message`n"
        }
        'FailSecBlock' {
            "`n$($FV): Blocked by Security-configured Transport rule`n"
        }
        'FailBrokenTerm|FailNoMailbox|FailMailLoop' {
            "`n$($FV): No valid recipient found: Broken offboarded user: Email looped between environments until hop count exceeded, and Non-Delivery Notice (NDR) was issued`n"
        }
        'FailADUserTermOU' {
            "`n$($FV): No valid recipient found: Broken offboarded user:ADUser is in Term OU: Email looped between environments until hop count exceeded, and Non-Delivery Notice (NDR) was issued`n"
        }
        'FailADUserDisabled' {
            "`n$($FV): No valid recipient found: Broken offboarded user: ADUser is disabled: Email looped between environments until hop count exceeded, and Non-Delivery Notice (NDR) was issued`n"
        }  
        'FailReason' {
            "`n$($FV): Other error, with a 'Reason' specification`n"
            $theseFails.FailDetailDetail |%{"`n$($_)"}
        }
        default{
            "`n$($FV): Undefined error (not configured as a response in this script)`n"
            $theseFails.FailDetailDetail |%{"`n$($_)"}
        }   
    }
)

*------^ END Delivery Status:Failed of type: $($FV) ^------

"@ ; 
                        #write-host @pltWH $hsFailRpt ; 
                        $smsg = $hsFailRpt ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                    } ;  # loop-E

                    $ArrayFails = $hReports.MsgsFail | ?{$_.failcode -is [array]} ; 
                    $FailVariantsArray = $ArrayFails | group failcode | select -expand name; 
                    $BadTermFailCodes = 'FailBrokenTerm','FailNoMailbox','FailADUserTermOU','FailADUserDisabled' ; 
                    [regex]$rgxBadTerm = ('(' + (($BadTermFailCodes  |%{[regex]::escape($_)}) -join '|') + ')') ;
                    #$pltWH=get-colorcombo -Rand ;
                    #foreach($FV in $FailVariantsArray){
                    foreach($Fail in $ArrayFails){
                        # do the BadTermFails
                        #$BadTermFails = $hReports.MsgsFail | ?{$_.failcode -match $rgxbadterm}
                        #$OtherArrayFails = $hReports.MsgsFail | ?{$_.failcode -notmatch $rgxbadterm}

                        #$(($theseFails|ft -a $prpGXMTfta | out-string).trim())       
                        # MessageID: $($Fail.MessageID)                  
                        $hsFailRpt = @"

*------v Delivery Status:Failed : $($Fail.MessageID) v------


$(($Fail |ft -a $prpFailMsg  | out-string).trim())

$(
    
    if($Fail.FailDetailEvent){
        "`nFailDetailEvent: $(($Fail.FailDetailEvent) -join ',')`n"
    }
    if($Fail.FailDetailDetail){
        "`nFailDetailDetail:`n$(($Fail.FailDetailDetail | out-string).trim())`n"
    }
    if($Fail.FailXoRecipientType){
        "`nCloud Recipient"
        "`nFailXoRecipientType: $(($Fail.FailXoRecipientType | out-string).trim())"
        switch ($Fail.FailXoRecipientType){
            "UserMailbox" {
                "`nis a standard functional MAILBOX: UserMailbox" ; 
            }
            "SharedMailbox" {
                "`nis a standard functional MAILBOX: SharedMailbox" ; 
            }
            "EquipmentMailbox" {
                "`nis a standard functional MAILBOX: EquipmentMailbox" ; 
            }
            "RoomMailbox" {
                "`nis a standard functional MAILBOX: RoomMailbox" ; 
            }
            "MailUser" {
                "`nMAILUSER: Generally reflects a removed license: => MS immediately deletes mailbox"
                "`nis a NON-MAILBOX: MailUser *forwards* to matching OnPrem/external UserMailbox object" ; 
            } ;
            default {
                write-warning "$((get-date).ToString('HH:mm:ss')):Unsupported RecipientType:($Fail.recipienttype). EXITING!" ;
                Break ;
            }
        }
        
    }
    if($fail.FailXopRecipientType){
        "`n`nOnPrem Recipient"
        "`nFailXopRecipientType: $(($fail.FailXopRecipientType | out-string).trim())"
        switch ($fail.FailXopRecipientType){
            'RemoteUserMailbox' {
                "`nis a NON-MAILBOX: RemoteUserMailbox *forwards* to matching cloud UserMailbox object" ; 
            }
            'RemoteSharedMailbox' {
                "`nis a NON-MAILBOX: RemoteSharedMailbox *forwards* to matching cloud SharedMailbox object" ; 
            }
            'RemoteRoomMailbox' {
                "`nis a NON-MAILBOX: RemoteRoomMailbox *forwards* to matching cloud RoomMailbox object" ; 
            }
            'RemoteEquipmentMailbox' {
                "`nis a NON-MAILBOX: RemoteEquipmentMailbox *forwards* to matching cloud EquipmentMailbox object" ; 
            }
            'UserMailbox' {
                "`nis a MISCONFIGURED MAILBOX: UserMailbox objects should no longer remain OnPrem (longer than it takes to migrate them to cloud during onboarding)" ; 
            }
            'SharedMailbox' {
                "`nis a MISCONFIGURED MAILBOX: SharedMailbox objects should no longer remain OnPrem (longer than it takes to migrate them to cloud during onboarding)" ; 
            }
            'MailUser' {
                "`nMAILUSER WO RMBX DETECTED! - POSSIBLE NOBRAIN?"
                "`nis a NON-MAILBOX: MailUser forwards to matching cloud UserMailbox object" ; 
            }
            'MailUniversalDistributionGroup' {
                "`nis a NON-MAILBOX: MailUniversalDistributionGroup are DistrubutionGroup objects that distribute mail to a membership" ; 
            }
            'DynamicDistributionGroup'  {
                "`nis a NON-MAILBOX: DynamicDistributionGroup are Dynamic DistrubutionGroup objects that distribute mail to an on-demand query-populated membership" ; 
            }
            'MailContact' {
                "`nis a NON-MAILBOX: MailContact is a non-SecurityPrincipal, that forwards mail to an exteral email address" ; 
            }
            default{
                "`nUnable to resolve `$fail.FailXopRecipientType: $($fail.FailXopRecipientType)" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                throw $smsg ;
                break ;
            }
        }
    }
    if($Fail.ADUserDisabled){
        "`nADUserDisabled: $(($Fail.ADUserDisabled | out-string).trim())"
    }
    if($Fail.ADUserTermOU){
        "`nADUserTermOU: $(($Fail.ADUserTermOU | out-string).trim())"
    }
    if($fail.FailXopRecipientType -notmatch '^(Mailbox|SharedMailbox|RoomMailbox|EquipmentMailbox)' -AND $Fail.FailXoRecipientType -notmatch '^(Mailbox|SharedMailbox|RoomMailbox|EquipmentMailbox)'){
        "`n`n==> EXPECTED LOOP FAILURE: BOTH CLOUD AND ONPREM RECIPIENT OBJECTS ARE _NON-MAILBOXES_ `n=> THERE IS NO WHERE TO DELIVER ANY MESSAGES TO THE ADDRESS!`n" ; 
    } 
)

*------^ END  Delivery Status:Failed : $($Fail.MessageID)  ^------

"@ ; 
                        #write-host @pltWH $hsFailRpt ; 
                        $smsg = $hsFailRpt ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                    } ;  # loop-E

                } ;
                if(-not $NoQuarCheck -AND ($mQuars = $msgs | ?{$_.status -eq 'Quarantined'})){
                    $hReports.add('MsgsQuar',$mQuars) ;
                    $ofileQ = $ofile.replace('-EXOMsgTrc,','QUARMsgs,') ;
                    #set-variable -name "$($vn)_QUAR" -Value ($mQuars) -ea STOP;
                    if($DoExports){
                        TRY{
                            $mQuars | export-csv -notype -path $ofileQ -ea STOP ;
                            $smsg = "export-csv'd to:`n$((resolve-path $ofileQ).path)" ; 
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                            else{ write-host -foregroundcolor yellow "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                        } CATCH {
                            $ErrTrapd=$Error[0] ;
                            $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                            else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        } ; 
                    } ; 
                    $smsg = "EXPANDING QUARANTINES:`n$" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level PROMPT } 
                    else{ write-host -foregroundcolor YELLOW "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    $ttl = $mQuars |  measure | select -expand count ;
                    $prcd=0 ;
                    $mQuars |foreach-object{
                        $tmsg = $_ ;
                        $prcd++ ;
                        $smsg = $sBnrS="`n`n#*------v PROCESSING QUAR:($($prcd)/$($ttl)): $($tmsg.MessageID) v------" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        $qmsg = Get-xoQuarantineMessage -MessageId $tmsg.MessageID ;
                        $qmsg |foreach-object{
                            $qid = $_.identity ;
                            $smsg = "`n$(($qmsg|ft -a $prpGXQMfta | out-string).trim())`n" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            $hdr = Get-xoQuarantineMessageHeader -Identity $qid | select -expand header;
                            #$rgxReturnPath = "Return-Path:((\n|\r|\s)*)([0-9a-zA-Z]+[-._+&='])*[0-9a-zA-Z]+@([-0-9a-zA-Z]+[.])+[a-zA-Z]{2,63}" ;
                            if($hdr -match $rgxReturnPath){
                                $smsg = "$(( (($matches[0] -split ':' |foreach-object{$_.trim()} ) -join ': ') |out-string).trim())" ;
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level PROMPT } 
                                else{ write-host -foregroundcolor YELLOW "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            } ;
                            $hdrsp = $hdr.Split([Environment]::NewLine) ;
                            write-host  "$(($hdrsp | ?{$_ -match $rgxHdrSenderIDKeys}|out-string).trim())" ;
                            start-sleep -Milliseconds 500 ;
                        } ;
                        $smsg = "$($sBnrS.replace('-v','-^').replace('v-','^-'))" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ;
                } ;

                if( $msgs | ?{$_.status -eq 'GettingStatus'}){
                    $smsg = "Status:GettingStatus returned on some traces - INDETERMINANT STATUS THOSE ITEMS (PENDING TRACKABLE LOGGING), RERUN IN A FEW MINS TO GET FUNCTIONAL DATA! (EXO-SIDE ISSUE)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                } ;

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
                    $smsg = "`n[$(($msgs|measure).count)msgs]|=>Get-xoMessageTraceDetail:" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-host "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 

                    $mtds = $mtdmsgs | Get-xoMessageTraceDetail ;

                    $mtdRpt = @() ; 
                    if($DetailedReportRuleHits){
                        $TRules = Get-xotransportrule  ; 
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
                        if($DoExports){
                            $ofileMTD = $ofile.replace('-MsgTrc','-MTD') ;
                            $smsg = "($(($mtds|measure).count)mtds | export-csv $($ofileMTD))" ; 
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                            else{ write-host -foregroundcolor white "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                            TRY{
                                $mtds | select $propsMTD | export-csv -notype -path $ofileMTD -ea Stop ;
                                $smsg = "export-csv'd to:`n$((resolve-path $ofileMTD).path)" ; 
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                                else{ write-host -foregroundcolor white "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                            } CATCH {
                                $ErrTrapd=$Error[0] ;
                                $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                                else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            } ; 
                            if(test-path -path $ofileMTD){
                                $smsg = "(log file confirmed)" ;
                                $smsg += "`n$($mtds.count) MTD matches output to:`n'$($ofileMTD)" ;
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                                else{ write-host "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 

                            } else { 
                                $smsg = "MISSING MTD LOG FILE!" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                            } ;

                        } ; 

                        $hReports.add('MTDetails',$mtds) ; 

                        if($DoExports){
                            # add the csvfilename
                            $smsg = "(adding `$hReports.MTDCSVFile:$($ofileMTD))" ; 
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                            else{ write-host -foregroundcolor white "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                            $hReports.add('MTDCSVFile',$MTDCSVFile) ; 
                        } ; 

                        #$hReports.add('MTDReport',$ofileMTD) ; 
                        # mtdreport
                        $hReports.add('MTDReport', $mtdRpt) ; 

                        if($DoExports){
                            $ofileMTDRpt = $ofile.replace('-MsgTrc','-MTDRpt') ;
                            TRY{
                                $mtdRpt | export-csv -notype -path $ofileMTDRpt -ea Stop ;
                            } CATCH {
                                $ErrTrapd=$Error[0] ;
                                $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                                else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            } ; 
                            if(test-path -path $ofileMTD){
                                $smsg = "(log file confirmed)" ;
                                $smsg += "`n$($mtdRpt.count) MTDReport matches output to:`n'$($ofileMTDRpt)" ;
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                                else{ write-host "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                                # add the csvfilename
                                $smsg = "(adding `$hReports.MTDRptCSVFile:$($ofileMTDRpt))" ; 
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                                else{ write-host -foregroundcolor white "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                                $hReports.add('MTDRptCSVFile',$ofileMTDRpt) ; 

                            } else { 
                                $smsg = "MISSING MTD LOG FILE!" 
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug 
                                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                            } ;
                        } ; 
                   } ;
                } ;
            } else {
                $smsg = "NO MATCHES RETURNED from MT Query" ;
                $smsg += "`n$(($pltGXMT|out-string).trim())" ; 
                $smsg += "`n(net of any relevant ConnectorId or other postfilters)"  ; 
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
        if($SimpleTrack -AND ($hReports.Keys.Count -gt 0)){
            $smsg = "-SimpleTrack specified: Only returning net message tracking set to pipeline" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
            $msgs | write-output ; 
        } else { 
            $smsg = "(no -SimpleTrack: returning full summary object to pipeline)" ; 
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            if($hReports.Keys.Count -gt 0){
                # convert the hashtable to object for output to pipeline
                #$Rpt += New-Object PSObject -Property $hReports ;
                $smsg = "(Returning summary object to pipeline)" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                
                TRY{
                    New-Object -TypeName PsObject -Property $hReports | write-output ; 
                    # export the entire return object into xml
                    $smsg = "(exporting `$hReports summary object to xml:$($ofile.replace('.csv','.xml')))" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    New-Object -TypeName PsObject -Property $hReports | export-clixml -path $ofile.replace('.csv','.xml') -ea STOP
                } CATCH {
                    $ErrTrapd=$Error[0] ;
                    $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                    else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                } ; 
            } else { 
                $smsg = "Unpopulated `$hReports, skipping output to pipeline" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARNING } 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                $false | write-output ; 
            } ;  
        } ; 
    } ; 
}

#*------^ Get-EXOMessageTraceExportedTDO ^------