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
    * 4:25 PM 1/6/2026 pulled in latest CONNECT_O365SERVICES, CALL_CONNECT_O365SERVICES, CALL_CONNECT_OPSERVICES, START_LOG_OPTIONS; 
        fundemental retool, porting from AzureAD -> MgGraph (f M$!), working again, for first time since they disabled AAD access! fundemental retooling in the CONNECT_O365SERVICES block, added test-MGConnection() etc. 
    * 4:27 PM 10/21/2025 fixed /tested more: loop (M$ v2 pagination offload) ; add: -ResultSize; fundemental update to use Get-xoMessageTraceV2 & Get-xoMessageTraceDetailV2 mandates (gxmt & gxmtd are now borked; fail) ; 
        fixed connect-ExchangeServerTDO() return (brought over updated latest), post-import requirement; remove refs to removed reconnect-eXO2; typo fix, WARNING->WARN
    * 12:58 PM 5/5/2025 cx10 was out of date (prompting for manual creds); brought in fresh cbp copies of all internal funcs, and replaced svcs_conn block, logging etc from scratch. Working now, no prompts.
    * 2:28 PM 5/2/2025 main Catch wasn't returning underlying EOM cmdlet errors; added code to -force, dump; 
        -As gxmtd is now flaking out, adapting the $QuarExpandLimitPerSender support to per-Recipient (when -SenderAddress used), and per-Sender (when -RecipientAddress used), to cut down on repetitive 
            lengthy Get-xoMessageTraceDetail calls. If you want more, push up the QuarExpandLimitPerSender count; cleaned up rem'd code obso'd by pull-GetxoMessageTraceDetail(), along with other broad code rems.
        - functionalized Get-xoMessageTraceDetail w retry (pull-GetxoMessageTraceDetail()) to address null gxmtd back, & retrying pulled, rewrote all Get-xoMessageTraceDetail to use the func
        - updated all supporting core functions, moved functions block to top (matching issues-addressing seen w cmw boxes, unless funcs preloaded - no local mods)
        - prior: had issues getting gxmtd's out using pipeline, so expanded into a loop with a throttlems wait -> seems working; 
        - split Fail|Failed into 2 lines (as it's a lookup, both on one line never matches); expanded CBH with splat of full range of usable params (dupes psb-psMsgTrkEXO)
        - updated cbh with all useful params in demo.
    * 2:47 PM 5/1/2025 getting Status:GettingStatus on SAP confirmation passes, added Get-xoMessageTraceDetail pass on last 20 of the set, seems to expose actual delivery resolution wihere Get-xoMessageTrace has the bozo status. 
        Aggregates findings of the 20 and adds them to the returned vari. updated CBH with output sample 
    * 9:35 AM 4/23/2025 reduced MessageTraceDetailLimit default from 100-> 20 (too time consuming, if not really needed), flipped it's effect to filtering last xx, not first.
        added alias: 'Get-EXOMessageTraceTDO', matches concept on Get-MessageTrackingLogTDO() naming. 
    * 6:08 PM 4/22/2025 post cmw testing, spliced over updated svc_conn block, full write-log() (simplified lacks success lvl etc);
         ADD: resolve-environment() & support, and updated start-log support; TLS_LATEST_FORCE ; missing regions; SWRITELOG ; SSTARTLOG ; 
        updated -Version supporting Connect-ExchangeServerTDO  ; convertFrom-MarkdownTable() to support... ; Initialize-exoStatusTable; 
        fixed bug in -resultsize code; code to leverage Initialize-exoStatusTable and output uniqued eventid's returnedon gmtl passes (doc output inline)
        copied over latest service conn code & slog for renv()
    * # 8:57 AM 12/6/2024 it's taking *5mins* to Get-xoQuarantineMessage; there's no point in running that 15 times, for the same sender,
         w same header & senderID specs. We need to down group the SenderAddress, and just process the last most-recent 'x', $QuarExpandLimitPerSender
         Added: -QuarExpandLimitPerSender 
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

        -Status, underlying Get-xoMessageTrace supports: Delivered|Expanded|Failed|FilteredAsSpam|GettingStatus|None|Quarantined
            But the range of documented Status returns (via results post-filtering) is currently:
            Defer|Deliver|Delivered|Expand|Expanded|Fail|Failed|FilteredAsSpam|GettingStatus|None|Pending|Quarantined|Receive|Resolved|Send|Transfer
            
            Get-xoMessageTraceDetail also returns additional, undocumented: 'Submit|The message was submitted' (expanding GettingStatus items)

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
    .PARAMETER SubjectFilterType
    You specify how the value is evaluated in the message subject by using the SubjectFilterType parameter (Contains|EndsWith|StartsWith)
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
    .PARAMETER ResultSize
    The ResultSize parameter specifies the maximum number of results to return. A valid value is from 1 to 5000. The default value is 1000. Note: This parameter replaces the PageSize parameter that was available on the Get-MessageTrace cmdlet.
    .PARAMETER SimpleTrack
    Switch to just return the net messages on the initial track (no Fail/Quarantine, MTDetail or other post-processing summaries) [-simpletrack]
    .PARAMETER Detailed
    Switch to perform MessageTrackingDetail pass, after intial MessageTrace (up to limit specified in -MessageTraceDetailLimit (defaults true) [-Detailed]
    .PARAMETER DetailedOtherFails
    Switch to perform MessageTrackingDetail pass, for any 'Other' Fails (up to limit specified in -MessageTraceDetailLimit (defaults true) [-DetailedOtherFails]
    .PARAMETER DetailedReportRuleHits
    Switch to do Summarize & Expansion of any MTD TransportRule events (defaults true) [-DetailedReportRuleHits]
    .PARAMETER MessageTraceDetailLimit
    Integer number of maximum messages to be follow-up MessageTraceDetail'd (defaults to 20) [-MessageTraceDetailLimit 100]
    .PARAMETER NoQuarCheck
    NoQuarCheck
    Switch to DISABLE expansion of status:'Quarantined' messages into slow Get-QuarantineMessage & Get-QuarantineMessageHeader details[-NoQuarCheck]
    .PARAMETER QuarExpandLimitPerSender
    Integer number of maximum most-recent messages per SenderAddress, to be Expanded into Quarantine details & Quarantine Headers (defaults to 1)[-QuarExpandLimitPerSender 2]
    .PARAMETER DoExports
    Switch to perform configured csv exports of results (defaults true) [-DoExports]
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
        PS>    ToIP='' ;
        PS>    SimpleTrack = $false ;
        PS>    Detailed = $false ;
        PS>    DetailedReportRuleHits = $false ;
        PS>    DetailedOtherFails = $true ;
        PS>    MessageTraceDetailLimit = 20 ;
        PS>    NoQuarCheck='';
        PS>    QuarExpandLimitPerSender = 1 ;
        PS>    DoExports = $true ;
        PS>    TenOrg = $global:o365_TenOrgDefault ; 
        PS>    silent = $false ;      
        PS>    verbose = $true ; 
        PS> } ;
        PS> $pltGxMT = [ordered]@{} ;
        PS> $pltI.GetEnumerator() | ?{ $_.value}  | ForEach-Object { $pltGxMT.Add($_.Key, $_.Value) } ;
        PS> $vn = (@("xoMsgs$($pltI.ticket)",$pltI.Tag) | ?{$_}) -join '_' ;
        PS> write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):Get-EXOMessageTraceExportedTDO w`n$(($pltGxMT|out-string).trim())`n(assign to `$$($vn))" ;
        PS> if(gv $vn -ea 0){rv $vn} ;
        PS> if($tmsgs = Get-EXOMessageTraceExportedTDO @pltGxMT){sv -na $vn -va $tmsgs ;
        PS> write-host "(assigned to `$$vn)"} ;

            ...
            14:51:33:Raw sender/recipient events:1850
            14:51:33:(1850 events | export-csv d:\scripts\logs\900881_x2xxxx,Txxxxx-xxxxxx.xxxxxxx@xxxx.com-EXOMsgTrc,TO_xxxxxxx@xxxx.com-2d-20250429-1951-run20250501-1451.csv)
            14:51:34:export-csv'd to:
            D:\scripts\logs\900881_x2xxxx,Txxxxx-xxxxxx.xxxxxxx@xxxx.com-EXOMsgTrc,TO_xxxxxxx@xxxx.com-2d-20250429-1951-run20250501-1451.csv
            14:51:34:(adding $hReports.MTMessages)
            14:51:34:(adding $hReports.MTMessagesCSVFile:d:\scripts\logs\900881_x2xxxx,Txxxxx-xxxxxx.xxxxxxx@xxxx.com-EXOMsgTrc,TO_xxxxxxx@xxxx.com-2d-20250429-1951-run20250501-1451.csv)
            14:51:34:
            #*------v Status DISTRIB v------

            14:51:34:
            Count Name
            ----- ----
              963 Resolved
              881 Delivered
                5 FilteredAsSpam
                1 GettingStatus
            14:51:34:

            #*------^ Status DISTRIB ^------

            14:51:34:

            ## Status Definitions:
            Resolved The message was redirected to the new recipient address based on an Active Directory lookup. When this happens, the original recipient address will be listed in a separate row in the message trace along with the final delivery status for the message.
            Delivered The message was delivered to its destination.
            FilteredAsSpam The message was marked as spam (and moved to the mailbox 'Junk Email' folder).
            GettingStatus The message is waiting for status update.

            14:51:34:

            #*------v MOST RECENT MATCH v------

            14:51:34:
            ReceivedLocal    : 5/1/2025 2:49:11 PM
            Status           : Resolved
            SenderAddress    : xxxxxxx@xx-xxxxxxx.xxx
            RecipientAddress : xxxxxxx@xxxx.com
            Subject          : FW: help per below, need detail
            14:51:34:

            #*------^ MOST RECENT MATCH ^------

            WARNING: 14:51:34:Status:GettingStatus returned on some traces - INDETERMINANT STATUS THOSE ITEMS (PENDING TRACKABLE LOGGING), RERUN IN A FEW MINS TO GET FUNCTIONAL DATA! (EXO-SIDE ISSUE)
            14:51:34:

            #*------v GettingStatus's Attempt to Re-Resolve via Get-xoMessageTraceDetail (up to last 20 messages) v------

            14:51:40:

            ===#1: MsgId: <CH2PR04MB6619FCF5E2194B8622AAB01EED822@CH2PR04MB6619.namprd04.prod.outlook.com> : Status:GettingStatus
            Received            SenderAddress           RecipientAddress Subject
            --------            -------------           ---------------- -------
            5/1/2025 2:01:05 PM xxxxx.xxxxxxxx@xxxx.com xxxxxxx@xxxx.com xxxxxxxxxx xxxxxxxx xxxxxx      FW: xxxx - xxxxxxx  xxxxxx xxxx 40643310
            DetailDisposition:
            Date                Event  Detail
            ----                -----  ------
            5/1/2025 2:01:06 PM Submit The message was submitted.
            14:51:40:

            #*------^  GettingStatus's Attempt to Re-Resolve via Get-xoMessageTraceDetail (up to last 20 messages)  ^------

            14:51:40:(log file confirmed)
            14:51:40:1850 matches output to:
            'd:\scripts\logs\900881_x2xxxx,Txxxxx-xxxxxx.xxxxxxx@xxxx.com-EXOMsgTrc,TO_xxxxxxx@xxxx.com-2d-20250429-1951-run20250501-1451.csv'
            (copied to CB)
            14:51:40:(Returning summary object to pipeline)
            14:51:40:(exporting $hReports summary object to xml:d:\scripts\logs\900881_x2xxxx,Txxxxx-xxxxxx.xxxxxxx@xxxx.com-EXOMsgTrc,TO_xxxxxxx@xxxx.com-2d-20250429-1951-run20250501-1451.xml)

        Splatted demo, all configurable params, depict some common output profile features (conditional on content in the trace)
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
    [Alias('get-EXOMsgTraceDetailed','Get-EXOMessageTraceExported','Get-EXOMessageTraceTDO')]
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
        [Parameter(HelpMessage="You specify how the value is evaluated in the message subject by using the SubjectFilterType parameter (Contains|EndsWith|StartsWith) [-SubjectFilterType 'StartsWith']")]
            [ValidateSet("Contains","EndsWith","StartsWith")]
            #[Alias('MessageSubject')]
            [string]$SubjectFilterType,
        [Parameter(HelpMessage="The Status parameter filters the results by the delivery status of the message (None|GettingStatus|Failed|Pending|Delivered|Expanded|Quarantined|FilteredAsSpam),an array runs search on each, post-filter results to target full range of Status values). [-Status 'Failed']")]
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
            [Parameter(HelpMessage="The ResultSize parameter specifies the maximum number of results to return. A valid value is from 1 to 5000. The default value is 1000. Note: This parameter replaces the PageSize parameter that was available on the Get-MessageTrace cmdlet. [-ResultSize 2000]")]             
            [int32]$ResultSize,
        [Parameter(HelpMessage="Switch to just return the net messages on the initial track (no Fail/Quarantine, MTDetail or other post-processing summaries) [-simpletrack]")]
            [switch]$SimpleTrack,
        [Parameter(HelpMessage="Switch to do Summarize & Expansion of any MTD TransportRule events (defaults true) [-DetailedReportRuleHits]")]
            [switch]$DetailedReportRuleHits= $true,
        [Parameter(HelpMessage="Integer number of maximum messages to be follow-up MessageTraceDetail'd (defaults to 20) [-MessageTraceDetailLimit 100]")]
            [int]$MessageTraceDetailLimit = 20,
        [Parameter(HelpMessage="Switch to DISABLE expansion of status:'Quarantined' messages into slow Get-QuarantineMessage & Get-QuarantineMessageHeader details[-NoQuarCheck]")]
            [switch]$NoQuarCheck,
        [Parameter(HelpMessage="Integer number of maximum most-recent messages per SenderAddress, to be Expanded into Quarantine details & Quarantine Headers (defaults to 1)[-QuarExpandLimitPerSender 2]")]
            [int]$QuarExpandLimitPerSender = 1,
        [Parameter(HelpMessage="Switch to perform configured csv exports of results (defaults true) [-DoExports]")]
            [switch]$DoExports=$TRUE,
        [Parameter(HelpMessage="Switch to perform Get-xoMessageTraceDetail pass, after intial MessageTrace (up to limit specified in -MessageTraceDetailLimit (defaults false) [-Detailed]")]
            [switch]$Detailed,
        [Parameter(HelpMessage="Switch to perform Get-xoMessageTraceDetail pass, for any 'Other' Fails (up to limit specified in -MessageTraceDetailLimit (defaults true) [-DetailedOtherFails]")]
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
        
        # Pull the CUser mod dir out of psmodpaths:
        #$CUModPath = $env:psmodulepath.split(';')|?{$_ -like '*\Users\*'} ;
    
        # 2b4() 2b4c() & fb4() are located up in the CONSTANTS_AND_ENVIRO\ENCODED_CONTANTS block ( to convert Constant assignement strings)

        #region FUNCTIONS_FULLYEXTERNAL ; #*======v FUNCTIONS_FULLYEXTERNAL v======
        # Optional block that relies on local module installs (vs the FUNCTIONS_LOCAL integrated block that follows below, and the FUNCTIONS_LOCAL_INTERNAL that is used for completely non-shared local functions.)

        #region RESOLVE_ENVIRONMENTTDO ; #*------v verb-io\resolve-EnvironmentTDO v------
        if(-not(gi function:resolve-EnvironmentTDO -ea 0)){
            $smsg = "MISSING DEPENDANT: verb-io\resolve-EnvironmentTDO !" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
            break ; 
        } ;
        #endregion RESOLVE_ENVIRONMENTTDO ; #*------^ END verb-io\resolve-EnvironmentTDO ^------

        #region WRITE_LOG ; #*------v verb-logging\write-log v------
        if(-not(gi function:write-log -ea 0)){
            $smsg = "MISSING DEPENDANT: verb-logging\write-log !" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
            break ; 
        } ; 
        #endregion WRITE_LOG ; #*------^ END verb-logging\write-log  ^------
    
        #region START_LOG ; #*------v verb-logging\Start-Log v------
        if(-not(gi function:start-log -ea 0)){
            $smsg = "MISSING DEPENDANT: verb-logging\start-log !" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
            break ; 
        } ; 
        #endregion START_LOG ; #*------^ END verb-logging\start-log ^------
    
        #region RESOLVE_NETWORKLOCALTDO ; #*------v verb-Network\resolve-NetworkLocalTDO v------
        if(-not(gi function:resolve-NetworkLocalTDO -ea 0)){
            $smsg = "MISSING DEPENDANT: verb-Network\resolve-NetworkLocalTDO!" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
            break ; 
        }
        #endregion RESOLVE_NETWORKLOCALTDO ; #*------^ END verb-Network\resolve-NetworkLocalTDO ^------

        #region PUSH_TLSLATEST ; #*------v verb-Network\push-TLSLatest v------
        if(-not(gi function:push-TLSLatest -ea 0)){
            $smsg = "MISSING DEPENDANT: verb-Network\push-TLSLatest!" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
            break ; 
        } ; 
        #endregion PUSH_TLSLATEST ; #*------^ END verb-Network\push-TLSLatest ^------
    
        #region TEST_EXCHANGEINFO ; #*------v verb-Ex2010\test-LocalExchangeInfoTDO v------
        if(-not (get-item function:test-LocalExchangeInfoTDO -ea 0)){
            $smsg = "MISSING DEPENDANT: verb-Ex2010\test-LocalExchangeInfoTDO!" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
            break ; 
        } ; 
        #endregion TEST_EXCHANGEINFO ; #*------^ END verb-Ex2010\test-LocalExchangeInfoTDO ^------
    
        #region CONNECT_O365SERVICES ; #*======v verb-exo\connect-O365Services v======
        if(-not (get-childitem function:connect-O365Services -ea 0)){
            $smsg = "MISSING DEPENDANT: verb-exo\connect-O365Services!" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
            break ; 
        } ;
        #endregion CONNECT_O365SERVICES ; #*======^ END verb-exo\connect-o365services ^======

        #region OUT_CLIPBOARD ; #*------v verb-IO\out-Clipboard v------
        if(-not(gci function:out-Clipboard -ea 0)){
            $smsg = "MISSING DEPENDANT: verb-IO\out-Clipboard!" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
            break ; 
        } ; 
        #endregion OUT_CLIPBOARD ; #*------^ END verb-IO\out-Clipboard ^------

        #region START_SLEEPCOUNTDOWN ; #*------v verb-IO\start-sleepcountdown v------
        if (-not (get-command start-sleepcountdown -ea 0)) {
            $smsg = "MISSING DEPENDANT: verb-IO\start-sleepcountdown!" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
            break ; 
        } ;
        #endregion START_SLEEPCOUNTDOWN ; #*------^ END verb-IO\start-sleepcountdown ^------

        #region CONVERTFROM_MARKDOWNTABLE ; #*------v verb-IO\convertFrom-MarkdownTable v------
        if(-not(gci function:convertFrom-MarkdownTable -ea 0)){
            $smsg = "MISSING DEPENDANT: verb-IO\convertFrom-MarkdownTable!" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
            break ; 
        } ; 
        #endregion CONVERTFROM_MARKDOWNTABLE ; #*------^ END verb-IO\convertFrom-MarkdownTable ^------

        #region REMOVE_INVALIDVARIABLENAMECHARS ; #*------v verb-IO\Remove-InvalidVariableNameChars v------        
        if(-not (gcm Remove-InvalidVariableNameChars -ea 0)){
            Function Remove-InvalidVariableNameChars ([string]$Name) {
                ($Name.tochararray() -match '[A-Za-z0-9_]') -join '' | write-output ;
            };
        } ;
        #endregion REMOVE_INVALIDVARIABLENAMECHARS ; #*------^ END verb-IO\Remove-InvalidVariableNameChars ^------
        
        #endregion FUNCTIONS_FULLYEXTERNAL ; #*======^ END FUNCTIONS_FULLYEXTERNAL ^======

        #region FUNCTIONS_LOCAL_INTERNAL ; #*======v FUNCTIONS_LOCAL_INTERNAL v======

        #region INITIALIZE_EXOSTATUSTABLE ; #*------v Initialize-exoStatusTable v------
        #*------v Initialize-exoStatusTable.ps1 v------
        function Initialize-exoStatusTable {
            <#
            .SYNOPSIS
            Initialize-exoStatusTable - Builds an indexed hash tabl of Exchange Server Get-MessageTrackingLog Statuss
            .NOTES
            Version     : 1.0.0
            Author      : Todd Kadrie
            Website     : http://www.toddomation.com
            Twitter     : @tostka / http://twitter.com/tostka
            CreatedDate : 2025-04-22
            FileName    : Initialize-exoStatusTable
            License     : (none asserted)
            Copyright   : (none asserted)
            Github      : https://github.com/tostka/verb-Ex2010
            Tags        : Powershell,EmailAddress,Version
            AddedCredit : Bruno Lopes (brunokktro )
            AddedWebsite: https://www.linkedin.com/in/blopesinfo
            AddedTwitter: @brunokktro / https://twitter.com/brunokktro
            REVISIONS
            * 1:47 PM 7/9/2024 CBA github field correction
            * 1:22 PM 5/22/2024init
            .DESCRIPTION
            Initialize-exoStatusTable - Builds an indexed hash tabl of Exchange Server Get-MessageTrackingLog Statuses

            .OUTPUT
            String
            .EXAMPLE
            PS> $StatusLookupTbl = Initialize-StatusTable ; 
            PS> $smsg = "`n`n## Status Definitions:" ; 
            PS> $TrackMsgs | group Status | select -expand Name | foreach-object{                   
            PS>     $smsg += "`n$(($StatusLookupTbl[$_]| ft -hidetableheaders | out-string).trim())" ; 
            PS> } ; 
            PS> $smsg += "`n`n" ; 
            PS> if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Prompt } 
            PS> else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            Demo resolving histogram Status uniques, to MS documented meansings of each event id in the msgtrack.
            .EXAMPLE
            ps> Initialize-exoStatusTable -EmailAddress 'monitoring+SolarWinds@toro.com;notanemailaddresstoro.com,todd+spam@kadrie.net' -verbose ;
            PS> 
            Demo with comma and semicolon delimiting, and an invalid address (to force a regex match fail error).
            .LINK
            https://github.com/brunokktro/EmailAddress/blob/master/Get-ExchangeEnvironmentReport.ps1
            .LINK
            https://github.com/tostka/verb-Ex2010
            #>
            [CmdletBinding()]
            #[Alias('rvExVers')]
            PARAM() ;
            BEGIN {
                ${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name ;
                $verbose = $($VerbosePreference -eq "Continue")
                $sBnr="#*======v $($CmdletName): v======" ;
                write-verbose  "$((get-date).ToString('HH:mm:ss')):$($sBnr)" ;
                
                $StatussMD = @"
Status|Description
---|---
Defer|The message delivery to the intended recipient was postponed and might be re-attempted later
Deliver|The message was delivered to its destination.
Delivered|The message was delivered to its destination.
Expand|The message was sent to a distribution group that was recently expanded.
Expanded|There was no message delivery because the message was addressed to a distribution group and the membership of the distribution was expanded (to the individual recipients)
Fail|Message delivery was attempted and it failed or the message was filtered as spam or malware, or by transport rules.
Failed|Message delivery was attempted and it failed or the message was filtered as spam or malware, or by transport rules.
FilteredAsSpam|The message was marked as spam (and moved to the mailbox 'Junk Email' folder).
GettingStatus|The message is waiting for status update.
None|The message has no delivery status because it was rejected or redirected to a different recipient.
Pending|Message delivery is underway or was deferred and is being retried.
Quarantined|The message was quarantined.
Receive|The message was received by the service (via Outlook submission or via SMTP from another server).
Resolved|The message was redirected to the new recipient address based on an Active Directory lookup. When this happens, the original recipient address will be listed in a separate row in the message trace along with the final delivery status for the message.|
Send|The message was sent by the service (via SMTP to another server).
Transfer|The recipient was moved to a bifurcated message because of content conversion, message recipient limits, or agents.
"@ ;

                $Object = $StatussMD | convertfrom-MarkdownTable ; 
                $Key = 'Status' ; 
                $Hashtable = @{}
            }
            PROCESS {
                Foreach ($Item in $Object){
                    $Procd++ ; 
                    $Hashtable[$Item.$Key.ToString()] = $Item ; 
                    if($ShowProgress -AND ($Procd -eq $Every)){
                        write-host -NoNewline '.' ; $Procd = 0 
                    } ; 
                } ;                 
            } # PROC-E
            END{
                $Hashtable | write-output ; 
                write-verbose  "$((get-date).ToString('HH:mm:ss')):$($sBnr.replace('=v','=^').replace('v=','^='))" ;
            }
        }; 
        #*------^ Initialize-exoStatusTable.ps1 ^------
        #endregion INITIALIZE_EXOSTATUSTABLE ; #*------^ END INITIALIZE_EXOSTATUSTABLE ^------

        #region pull-GetxoMessageTraceDetail ; #*------v pull-GetxoMessageTraceDetail v------
        function pull-GetxoMessageTraceDetail {
            <#
            .SYNOPSIS
            pull-GetxoMessageTraceDetail - wrap Get-xoMessageTraceDetail, with retry around `$null returns
            .NOTES
            REVISIONS
            * 10:57 AM 5/2/2025 INIT
            .DESCRIPTION
            pull-GetxoMessageTraceDetail - wrap Get-xoMessageTraceDetail, with retry around `$null returns
            .PARAMETER  Messages
            Array of Get-xoMessageTrace Message returns to be expanded into Get-xoMessageTraceDetail 
            .INPUTS
            Array of Get-xoMessageTrace Message returns
            .OUTPUTS
            SystemObject Returns array of resolved Get-xoMessageTraceDetail results
            .EXAMPLE
            PS> $mtds = pull-GetxoMessageTraceDetail -Messages $mtdmsgs ; 
            EXSAMPLEOUTPUT
            Run with whatif & verbose
            #>
            [CmdletBinding()]
            PARAM(
                [Parameter(Mandatory=$True,HelpMessage="Array of Get-xoMessageTrace Message returns to be expanded into Get-xoMessageTraceDetail ")]
                [array]$Messages
            ) ; 
            BEGIN{
                if(-not $RetrySleep){$RetrySleep = 10 } ; # wait time between retries
                if(-not $DawdleWait){$DawdleWait = 30 } ; # wait time (secs) between dawdle checks
                if(-not $RetryThrottle){$RetryThrottle = 60 } ; # wait time (secs) after Throttle error $errtest.Exception ||Your recent queries have surpassed the permitted limit, please try again later
                if(-not $rgxEXOThrottle){$rgxEXOThrottle = 'Your\srecent\squeries\shave\ssurpassed\sthe\spermitted\slimit,\splease\stry\sagain\slater' } 
            } ; 
            PROCESS{
                $mtds = @() ; 
                foreach( $mtdm in  $Messages){
                    $smsg = "--Get-xoMessageTraceDetail: MsgID: $($mtdm.MessageId) : To: $($mtdm.recipientaddress)" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                    $pltgXMTD=[ordered]@{
                        MessageTraceId = $mtdm.MessageTraceId ;
                        RecipientAddress = $mtdm.RecipientAddress
                        erroraction = 'STOP' ;
                        #whatif = $($whatif) ;
                    } ;
                    $smsg = "Get-xoMessageTraceDetail w`n$(($pltgXMTD|out-string).trim())" ; 
                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                    $Exit = 0 ;
                    Do {
                        TRY {                            
                            #if($rmtd = Get-xoMessageTraceDetail @pltgXMTD){
                            # 1:21 PM 10/21/2025 try to splice in Get-xoMessageTraceDetailV2
                            # doesn't seem to support the same -StartingRecipientAddress etc params, no evidence does warning pushback, lacks the gxmt's native params for those features. 
                            # does throttle HARD, just kills the connection
                            if($rmtd = Get-xoMessageTraceDetailV2 @pltgXMTD){
                                $mtds += $rmtd ;
                            } else {
                                write-warning "No Return: #$($Exit):MTId: $($pltgXMTD.MessageTraceId) : To: $($pltgXMTD.RecipientAddress)" ; 
                                throw "no Get-xoMessageTraceDetail return" ; 
                            } ; 
                            $Exit = $Retries ;
                        } CATCH [System.Exception] {
                            $ErrTrapd=$Error[0] ;
                            if($ErrTrapd.Exception -match $rgxEXOThrottle){
                                $smsg = "MS 100-qry limit/5mins throttling detected, waiting $(RetryThrottle)s to retry..." ; 
                                $smsg += "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                                #Start-Sleep -Milliseconds $ThrottleMs 
                                start-sleepcountdown -seconds $RetryThrottle -Rolling ; 
                                $Exit ++ ;
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN }
                                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                $smsg= "Try #: $($Exit)" ;
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Error }  #Error|Warn|Debug 
                                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                If ($Exit -eq $Retries) {
                                    $smsg= "Unable to exec cmd!" ;
                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Error }  #Error|Warn|Debug 
                                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                    BREAK ; 
                                } ;
                            } ; 
                        } CATCH {
                            $ErrTrapd=$Error[0] ;
                            $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                            Start-Sleep -Milliseconds $ThrottleMs 
                            $Exit ++ ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN }
                            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            $smsg= "Try #: $($Exit)" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Error }  #Error|Warn|Debug 
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            If ($Exit -eq $Retries) {
                                $smsg= "Unable to exec cmd!" ;
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Error }  #Error|Warn|Debug 
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                BREAK ; 
                            } ;
                        }  ;
                    } Until ($Exit -eq $Retries) ; 
                    start-sleep -Milliseconds $ThrottleMs  ;
                } ; 
            } ;  # PROC-E
            END{
                $mtds | write-output 
            } ; 
        } ;
        #endregion pull-GetxoMessageTraceDetail ; #*------^ END pull-GetxoMessageTraceDetail ^------
                
        #endregion FUNCTIONS_LOCAL_INTERNAL ; #*======^ END FUNCTIONS_LOCAL_INTERNAL ^======

        #region CONSTANTS_AND_ENVIRO ; #*======v CONSTANTS_AND_ENVIRO v======
        #region ENVIRO_DISCOVER ; #*------v ENVIRO_DISCOVER v------
        push-TLSLatest
        $Verbose = [boolean]($VerbosePreference -eq 'Continue') ; 
        $rPSCmdlet = $PSCmdlet ; # an object that represents the cmdlet or advanced function that's being run. Available on functions w CmdletBinding (& $args will not be available). (Blank on non-CmdletBinding/Non-Adv funcs).
        $rPSScriptRoot = $PSScriptRoot ; # the full path of the executing script's parent directory., PS2: valid only in script modules (.psm1). PS3+:it's valid in all scripts. (Funcs: ParentDir of the file that hosts the func)
        $rPSCommandPath = $PSCommandPath ; # the full path and filename of the script that's being run, or file hosting the funct. Valid in all scripts.
        $rMyInvocation = $MyInvocation ; # populated only for scripts, function, and script blocks.
        # - $MyInvocation.MyCommand.Name returns name of a function, to identify the current command,  name of the current script (pop'd w func name, on Advfuncs)
        # - Ps3+:$MyInvocation.PSScriptRoot : full path to the script that invoked the current command. The value of this property is populated only when the caller is a script (blank on funcs & Advfuncs)
        # - Ps3+:$MyInvocation.PSCommandPath : full path and filename of the script that invoked the current command. The value of this property is populated only when the caller is a script (blank on funcs & Advfuncs)
        #     ** note: above pair contain information about the _invoker or calling script_, not the current script
        $rPSBoundParameters = $PSBoundParameters ; 
        #region PREF_VARI_DUMP ; #*------v PREF_VARI_DUMP v------
        <#$script:prefVaris = @{
            whatifIsPresent = $whatif.IsPresent
            whatifPSBoundParametersContains = $rPSBoundParameters.ContainsKey('WhatIf') ; 
            whatifPSBoundParameters = $rPSBoundParameters['WhatIf'] ;
            WhatIfPreferenceIsPresent = $WhatIfPreference.IsPresent ; # -eq $true
            WhatIfPreferenceValue = $WhatIfPreference;
            WhatIfPreferenceParentScopeValue = (Get-Variable WhatIfPreference -Scope 1).Value ;
            ConfirmPSBoundParametersContains = $rPSBoundParameters.ContainsKey('Confirm') ; 
            ConfirmPSBoundParameters = $rPSBoundParameters['Confirm'];
            ConfirmPreferenceIsPresent = $ConfirmPreference.IsPresent ; # -eq $true
            ConfirmPreferenceValue = $ConfirmPreference ;
            ConfirmPreferenceParentScopeValue = (Get-Variable ConfirmPreference -Scope 1).Value ; 
            VerbosePSBoundParametersContains = $rPSBoundParameters.ContainsKey('Confirm') ; 
            VerbosePSBoundParameters = $rPSBoundParameters['Verbose'] ;
            VerbosePreferenceIsPresent = $VerbosePreference.IsPresent ; # -eq $true
            VerbosePreferenceValue = $VerbosePreference ;
            VerbosePreferenceParentScopeValue = (Get-Variable VerbosePreference -Scope 1).Value;
            VerboseMyInvContains = '-Verbose' -in $rPSBoundParameters.UnboundArguments ; 
            VerbosePSBoundParametersUnboundArgumentContains = '-Verbose' -in $rPSBoundParameters.UnboundArguments 
        } ;
        write-verbose "`n$(($script:prefVaris.GetEnumerator() | Sort-Object Key | Format-Table Key,Value -AutoSize|out-string).trim())`n" ; 
        #>
        #endregion PREF_VARI_DUMP ; #*------^ END PREF_VARI_DUMP ^------
        #region RV_ENVIRO ; #*------v RV_ENVIRO v------
        $pltRvEnv=[ordered]@{
            PSCmdletproxy = $rPSCmdlet ; 
            PSScriptRootproxy = $rPSScriptRoot ; 
            PSCommandPathproxy = $rPSCommandPath ; 
            MyInvocationproxy = $rMyInvocation ;
            PSBoundParametersproxy = $rPSBoundParameters
            verbose = [boolean]($PSBoundParameters['Verbose'] -eq $true) ; 
        } ;
        write-verbose "(Purge no value keys from splat)" ; 
        $mts = $pltRVEnv.GetEnumerator() |?{$_.value -eq $null} ; $mts |%{$pltRVEnv.remove($_.Name)} ; rv mts -ea 0 -whatif:$false -confirm:$false; 
        $smsg = "resolve-EnvironmentTDO w`n$(($pltRVEnv|out-string).trim())" ; 
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        if(get-command resolve-EnvironmentTDO -ea STOP){}ELSE{
            $smsg = "UNABLE TO gcm resolve-EnvironmentTDO!" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
            BREAK ; 
        } ; 
        $rvEnv = resolve-EnvironmentTDO @pltRVEnv ; 
        $smsg = "`$rvEnv returned:`n$(($rvEnv |out-string).trim())" ; 
        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        #endregion RV_ENVIRO ; #*------^ END RV_ENVIRO ^------
        #region NETWORK_INFO ; #*======v NETWORK_INFO v======
        if(get-command resolve-NetworkLocalTDO  -ea STOP){}ELSE{
            $smsg = "UNABLE TO gcm resolve-NetworkLocalTDO !" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
            BREAK ; 
        } ; 
        $netsettings = resolve-NetworkLocalTDO ; 
        if($env:Userdomain){ 
            switch($env:Userdomain){
                'CMW'{
                    #$logon_SID = $CMW_logon_SID 
                }
                'TORO'{
                    #$o365_SIDUpn = $o365_Toroco_SIDUpn ; 
                    #$logon_SID = $TOR_logon_SID ; 
                }
                $env:COMPUTERNAME{
                    $smsg = "%USERDOMAIN% -EQ %COMPUTERNAME%: $($env:computername) => non-domain-connected, likely edge role Ex server!" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                    if($netsettings.Workgroup){
                        $smsg = "WorkgroupName:$($netsettings.Workgroup)" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;                    
                    } ; 
                } ; 
                default{
                    $smsg = "$($env:userdomain):UNRECOGIZED/UNCONFIGURED USER DOMAIN STRING!" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                    THROW $SMSG 
                    BREAK ; 
                }
            } ; 
        } ;  # $env:Userdomain-E
        #endregion NETWORK_INFO ; #*======^ END NETWORK_INFO ^======
        #region OS_INFO ; #*------v OS_INFO v------
        <# os detect, covers Server 2016, 2008 R2, Windows 10, 11
        if (get-command get-ciminstance -ea 0) {$OS = (Get-ciminstance -class Win32_OperatingSystem)} else {$Os = Get-WMIObject -class Win32_OperatingSystem } ;
        #$isWorkstationOS = $isServerOS = $isW2010 = $isW2011 = $isS2016 = $isS2008R2 = $false ;
        write-host "Detected:`$Os.Name:$($OS.name)`n`$Os.Version:$($Os.Version)" ;
        if ($OS.name -match 'Microsoft\sWindows\sServer') {
            $isServerOS = $true ;
            if ($os.name -match 'Microsoft\sWindows\sServer\s2016'){$isS2016 = $true ;} ;
            if ($os.name -match 'Microsoft\sWindows\sServer\s2008\sR2') { $isS2008R2 = $true ; } ;
        } else { 
            if ($os.name -match '^Microsoft\sWindows\s11') {
                $isWorkstationOS = $true ;
                if ($os.name -match 'Microsoft\sWindows\s11') { $isW2011 = $true ; } ;
            } elseif ($os.name -match '^Microsoft\sWindows\s10') {
                $isWorkstationOS = $true ; $isW2010 = $true
            } else {
                $isWorkstationOS = $true ;
            } ;         
        } ; 
        #>
        #endregion OS_INFO ; #*------^ END OS_INFO ^------
        #region TEST_EXOPLOCAL ; #*------v TEST_EXOPLOCAL v------
        if(get-command test-LocalExchangeInfoTDO -ea STOP){}ELSE{
            $smsg = "UNABLE TO gcm test-LocalExchangeInfoTDO !" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
            BREAK ; 
        } ; 
        $lclExOP = test-LocalExchangeInfoTDO ; 
        write-verbose "Expand returned NoteProperty properties into matching local variables" ; 
        if($host.version.major -gt 2){
            $lclExOP.PsObject.Properties | ?{$_.membertype -eq 'NoteProperty'} | foreach-object{set-variable -name $_.name -value $_.value -verbose -whatif:$false -Confirm:$false ;} ;
        }else{
            write-verbose "Psv2 lacks the above expansion capability; just create simpler variable set" ; 
            $ExVers = $lclExOP.ExVers ; $isLocalExchangeServer = $lclExOP.isLocalExchangeServer ; $IsEdgeTransport = $lclExOP.IsEdgeTransport ;
        } ;
        #
        #endregion TEST_EXOPLOCAL ; #*------^ END TEST_EXOPLOCAL ^------

        <#
        #region PsParams ; #*------v PSPARAMS v------
        $PSParameters = New-Object -TypeName PSObject -Property $rPSBoundParameters ;
        # DIFFERENCES $PSParameters vs $PSBoundParameters:
        # - $PSBoundParameters: System.Management.Automation.PSBoundParametersDictionary (native obj)
        # test/access: ($PSBoundParameters['Verbose'] -eq $true) ; $PSBoundParameters.ContainsKey('Referrer') #hash syntax
        # CAN use as a @PSBoundParameters splat to push through (make sure populated, can fail if wrong type of wrapping code)
        # - $PSParameters: System.Management.Automation.PSCustomObject (created obj)
        # test/access: ($PSParameters.verbose -eq $true) ; $PSParameters.psobject.Properties.name -contains 'SenderAddress' ; # cobj syntax
        # CANNOT use as a @splat to push through (it's a cobj)
        write-verbose "`$rPSBoundParameters:`n$(($rPSBoundParameters|out-string).trim())" ;
        # pre psv2, no $rPSBoundParameters autovari to check, so back them out:
        #>
        <# recycling $rPSBoundParameters into @splat calls: (can't use $psParams, it's a cobj, not a hash!)
        # rgx for filtering $rPSBoundParameters for params to pass on in recursive calls (excludes keys matching below)
        $rgxBoundParamsExcl = '^(Name|RawOutput|Server|Referrer)$' ; 
        if($rPSBoundParameters){
                $pltRvSPFRec = [ordered]@{} ;
                # add the specific Name for this call, and Server spec (which defaults, is generally not 
                $pltRvSPFRec.add('Name',"$RedirectRecord" ) ;
                $pltRvSPFRec.add('Referrer',$Name) ; 
                $pltRvSPFRec.add('Server',$Server ) ;
                $rPSBoundParameters.GetEnumerator() | ?{ $_.key -notmatch $rgxBoundParamsExcl} | foreach-object { $pltRvSPFRec.add($_.key,$_.value)  } ;
                write-host "Resolve-SPFRecord w`n$(($pltRvSPFRec|out-string).trim())" ;
                Resolve-SPFRecord @pltRvSPFRec  | write-output ;
        } else {
            $smsg = "unpopulated `$rPSBoundParameters!" ;
            write-warning $smsg ;
            throw $smsg ;
        };     
        #>
        #endregion PsParams ; #*------^ END PSPARAMS ^------    
        #endregion ENVIRO_DISCOVER ; #*------^ END ENVIRO_DISCOVER ^------

        #region COMMON_CONSTANTS ; #*------v COMMON_CONSTANTS v------
    
        if(-not $DoRetries){$DoRetries = 4 } ;    # # times to repeat retry attempts
        if(-not $RetrySleep){$RetrySleep = 10 } ; # wait time between retries
        if(-not $RetrySleep){$DawdleWait = 30 } ; # wait time (secs) between dawdle checks
        if(-not $DirSyncInterval){$DirSyncInterval = 30 } ; # AADConnect dirsync interval
        if(-not $ThrottleMs){$ThrottleMs = 50 ;}
        if(-not $rgxDriveBanChars){$rgxDriveBanChars = '[;~/\\\.:]' ; } ; # ;~/\.:,
        if(-not $rgxCertThumbprint){$rgxCertThumbprint = '[0-9a-fA-F]{40}' } ; # if it's a 40char hex string -> cert thumbprint  
        if(-not $rgxSmtpAddr){$rgxSmtpAddr = "^([0-9a-zA-Z]+[-._+&'])*[0-9a-zA-Z]+@([-0-9a-zA-Z]+[.])+[a-zA-Z]{2,63}$" ; } ; # email addr/UPN
        if(-not $rgxDomainLogon){$rgxDomainLogon = '^[a-zA-Z][a-zA-Z0-9\-\.]{0,61}[a-zA-Z]\\\w[\w\.\- ]+$' } ; # DOMAIN\samaccountname 
        if(-not $exoMbxGraceDays){$exoMbxGraceDays = 30} ; 
        if(-not $XOConnectionUri ){$XOConnectionUri = 'https://outlook.office365.com'} ; 
        if(-not $SCConnectionUri){$SCConnectionUri = 'https://ps.compliance.protection.outlook.com'} ; 
        if(-not $XODefaultPrefix){$XODefaultPrefix = 'xo' };
        if(-not $SCDefaultPrefix){$SCDefaultPrefix = 'sc' };
        #$rgxADDistNameGAT = ",$(($TORMeta.UnreplicatedOU -split ',' | select -skip 1 ) -join ',')" 
        #$rgxADDistNameAT = ",$(($TORMeta.UnreplicatedOU -split ',' | select -skip 2 ) -join ',')"

        write-verbose "Coerce configured but blank Resultsize to Unlimited" ; 
        if(get-variable -name resultsize -ea 0){
            if( ($null -eq $ResultSize) -OR ('' -eq $ResultSize) ){$ResultSize = 'unlimited' }
            elseif($Resultsize -is [int]){} else {throw "Resultsize must be an integer or the string 'unlimited' (or blank)"} ;
        } ; 
        #$ComputerName = $env:COMPUTERNAME ;
        #$NoProf = [bool]([Environment]::GetCommandLineArgs() -like '-noprofile'); # if($NoProf){# do this};
        # XXXMeta derived constants:
        # - AADU Licensing group checks
        # calc the rgxLicGrpName fr the existing $xxxmeta.rgxLicGrpDN: (get-variable tormeta).value.rgxLicGrpDN.split(',')[0].replace('^','').replace('CN=','')
        #$rgxLicGrpName = (get-variable -name "$($tenorg)meta").value.rgxLicGrpDN.split(',')[0].replace('^','').replace('CN=','')
        # use the dn vers LicGrouppDN = $null ; # | ?{$_ -match $tormeta.rgxLicGrpDN}
        #$rgxLicGrpDN = (get-variable -name "$($tenorg)meta").value.rgxLicGrpDN
        # email trigger vari, it will be semi-delimd list of mail-triggering events
        $script:PassStatus = $null ;
        # TenOrg or other looped-specific PassStatus (auto supported by 7pswlt)
        #New-Variable -Name PassStatus_$($tenorg) -scope Script -Value $null ;
        [array]$SmtpAttachment = $null ;
        #write-verbose "start-Timer:Master" ; 
        $swM = [Diagnostics.Stopwatch]::StartNew() ;
        # $ByPassLocalExchangeServerTest = $true # rough in, code exists below for exempting service/regkey testing on this variable status. Not yet implemented beyond the exemption code, ported in from orig source.
        #endregion COMMON_CONSTANTS ; #*------^ END COMMON_CONSTANTS ^------
              
        #region LOCAL_CONSTANTS ; #*------v LOCAL_CONSTANTS v------

        # BELOW TRIGGERS/DRIVES TEST_MODS: array of: "[modname];[modDLUrl,or pscmdline install]"    
        $tDepModules = @() ;
        $useVerbCore = $true ; 
        if($useVerbCore){
            $tDepModules += @('verb-logging;localRepo;write-log') ; #start-log; write-log ;
            $tDepModules += @('verb-io;localRepo;resolve-EnvironmentTDO') ; #resolve-EnvironmentTDO
            $tDepModules += @('verb-Network;localRepo;resolve-NetworkLocalTDO') ; #resolve-NetworkLocalTDO; Send-EmailNotif
        } ;
        <# NOTE: Svc modules are tested as needed by connect-O365Servicees() & connect-OPServices()
        if($useEXO){$tDepModules += @("ExchangeOnlineManagement;https://www.powershellgallery.com/packages/ExchangeOnlineManagement/;Get-xoOrganizationConfig",'verb-exo;localRepo;connect-exo')} ;
        if($UseMSOL){$tDepModules += @("MSOnline;https://www.powershellgallery.com/packages/MSOnline/;Get-MsolDomain")} ;
        if($UseAAD){$tDepModules += @("AzureAD;https://www.powershellgallery.com/packages/AzureAD/;Get-AzureADTenantDetail")} ;
        if($UseExOP){$tDepModules += @('verb-Ex2010;localRepo;Connect-Ex2010')} ;
        if($UseMG){$tDepModules += @("Microsoft.Graph.Authentication;https://www.powershellgallery.com/packages/Microsoft.Graph/;Get-MgOrganization")} ;
        if($UseOPAD){$tDepModules += @("ActiveDirectory;get-windowscapability -name RSAT* -Online | ?{$_.name -match 'Rsat\.ActiveDirectory'} | %{Add-WindowsCapability -online -name $_.name};Get-ADDomain")} ;
        #>

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

        #endregion LOCAL_CONSTANTS ; #*------^ END LOCAL_CONSTANTS ^------  
          
        #region ENCODED_CONTANTS ; #*------v ENCODED_CONTANTS v------
        # ENCODED CONsTANTS & SUPPORT FUNCTIONS:
        #region 2B4 ; #*------v 2B4 v------
        if(-not (get-command 2b4 -ea 0)){function 2b4{[CmdletBinding()][Alias('convertTo-Base64String')] PARAM([Parameter(ValueFromPipeline=$true)][string[]]$str) ; PROCESS{$str|%{[Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($_))}  };} ; } ; 
        #endregion 2B4 ; #*------^ END 2B4 ^------
        #region 2B4C ; #*------v 2B4C v------
        # comma-quoted return
        if(-not (get-command 2b4c -ea 0)){function 2b4c{ [CmdletBinding()][Alias('convertto-Base64StringCommaQuoted')] PARAM([Parameter(ValueFromPipeline=$true)][string[]]$str) ;BEGIN{$outs = @()} PROCESS{[array]$outs += $str | %{[Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($_))} ; } END {'"' + $(($outs) -join '","') + '"' | out-string | set-clipboard } ; } ; } ; 
        #endregion 2B4C ; #*------^ END 2B4C ^------
        #region FB4 ; #*------v FB4 v------
        # DEMO: $SitesNameList = 'THluZGFsZQ==','U3BlbGxicm9vaw==','QWRlbGFpZGU=' | fb4 ;
        if(-not (get-command fb4 -ea 0)){function fb4{[CmdletBinding()][Alias('convertFrom-Base64String')] PARAM([Parameter(ValueFromPipeline=$true)][string[]]$str) ; PROCESS{$str | %{ [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($_)) }; } ; } ; }; 
        #endregion FB4 ; #*------^ END FB4 ^------
        # FOLLOWING CONSTANTS ARE USED FOR DEPENDANCY-LESS CONNECTIONS
        if(-not $o365_Toroco_SIDUpn){$o365_Toroco_SIDUpn = 'cy10b2RkLmthZHJpZUB0b3JvLmNvbQ==' | fb4 } ;
        $o365_SIDUpn = $o365_Toroco_SIDUpn ; 
        switch($env:Userdomain){
            'CMW'{
                if(-not $CMW_logon_SID){$CMW_logon_SID = 'Q01XXGQtdG9kZC5rYWRyaWU=' | fb4 } ; 
                $logon_SID = $CMW_logon_SID ; 
            }
            'TORO'{
                if(-not $TOR_logon_SID){$TOR_logon_SID = 'VE9ST1xrYWRyaXRzcw==' | fb4 } ; 
                $logon_SID = $TOR_logon_SID ; 
            }
            $env:COMPUTERNAME{
                $smsg = "%USERDOMAIN% -EQ %COMPUTERNAME%: $($env:computername) => non-domain-connected, likely edge role Ex server!" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                if($WorkgroupName = (Get-WmiObject -Class Win32_ComputerSystem).Workgroup){
                    $smsg = "WorkgroupName:$($WorkgroupName)" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                }
                if(($isLocalExchangeServer = (Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\Setup')) -or (
                        $isLocalExchangeServer = (Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup')) -or
                            $ByPassLocalExchangeServerTest){
                            $smsg = "We are on Exchange Server"
                            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                            $IsEdgeTransport = $false
                            if((Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\EdgeTransportRole') -or (Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\EdgeTransportRole')){
                                $smsg = "We are on Exchange Edge Transport Server"
                                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                                $IsEdgeTransport = $true
                            } ; 
                } else {
                    $isLocalExchangeServer = $false 
                    $IsEdgeTransport = $false ;
                } ;
            } ; 
            default{
                $smsg = "$($env:userdomain):UNRECOGIZED/UNCONFIGURED USER DOMAIN STRING!" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                THROW $SMSG 
                BREAK ; 
            }
        } ; 
        #endregion ENCODED_CONTANTS ; #*------^ END ENCODED_CONTANTS ^------
    
        #endregion CONSTANTS_AND_ENVIRO ; #*======^ CONSTANTS_AND_ENVIRO ^======
        
        # moved FUNCTIONS block up top of BEGIN
        #if($ticket){$pltSL.Tag = $ticket} ; 
        #region START_LOG_OPTIONS #*======v START_LOG_OPTIONS v======
        $useSLogHOl = $true ; # one or 
        $useTransPath = $false ; # TRANSCRIPTPATH
        $useTransRotate = $false ; # TRANSCRIPTPATHROTATE
        $useStartTrans = $false ; # STARTTRANS
        $useTransNoDep = $false ; # TRANSCRIPT_NODEP
        $useTransBasicScript = $false ; # BASIC_SCRIPT_TRANSCRIPT
        #region START_LOG_HOLISTIC #*------v START_LOG_HOLISTIC v------
        if($useSLogHOl){
            # Single log for script/function example that accomodates detect/redirect from AllUsers scope'd installed code, and hunts a series of drive letters to find an alternate logging dir (defers to profile variables)
            #${CmdletName} = $rPSCmdlet.MyInvocation.MyCommand.Name ;
            if(-not (get-variable LogPathDrives -ea 0)){$LogPathDrives = 'd','c' };
            foreach($budrv in $LogPathDrives){if(test-path -path "$($budrv):\scripts" -ea 0 ){break} } ;
            if(-not (get-variable rgxPSAllUsersScope -ea 0)){$rgxPSAllUsersScope="^$([regex]::escape([environment]::getfolderpath('ProgramFiles')))\\((Windows)*)PowerShell\\(Scripts|Modules)\\.*\.(ps(((d|m))*)1|dll)$" ;} ;
            if(-not (get-variable rgxPSCurrUserScope -ea 0)){$rgxPSCurrUserScope="^$([regex]::escape([Environment]::GetFolderPath('MyDocuments')))\\((Windows)*)PowerShell\\(Scripts|Modules)\\.*\.(ps((d|m)*)1|dll)$" ;} ;
            $pltSL=[ordered]@{Path=$null ;NoTimeStamp=$false ;Tag=$null ;showdebug=$($showdebug) ; Verbose=$($VerbosePreference -eq 'Continue') ;} ;
            if($whatif.ispresent){$pltSL.add('whatif',$($whatif))}
            elseif($WhatIfPreference.ispresent ){$pltSL.add('whatif',$WhatIfPreferenc)} ;         
            # if using [CmdletBinding(SupportsShouldProcess)] + -WhatIf:$($WhatIfPreference):
            #$pltSL=[ordered]@{Path=$null ;NoTimeStamp=$false ;Tag=$null ;showdebug=$($showdebug) ; Verbose=$($VerbosePreference -eq 'Continue') ; whatif=$($WhatIfPreference) ;} ;
            #$pltSL=[ordered]@{Path=$null ;NoTimeStamp=$false ;Tag="$($ticket)-$($TenOrg)-LASTPASS-" ;showdebug=$($showdebug) ; Verbose=$($VerbosePreference -eq 'Continue') ; whatif=$($WhatIfPreference) ;} ;
            #$pltSL.Tag = $((@($ticket,$usr) |?{$_}) -join '-')
            if($ticket){$pltSL.Tag = $ticket} ; ####
            #$pltSL.Tag = $env:COMPUTERNAME ; 
            #$pltSL.Tag = $((@($ticket,$usr) |?{$_}) -join '-')
            $tagfields = 'ticket','UserPrincipalName','folderscope' ; # DomainName TenOrg ModuleName 
            $tagfields | foreach-object{$fld = $_ ; if(get-variable $fld -ea 0 |?{$_.value} ){$pltSL.Tag += @($((get-variable $fld).value))} } ; 
            if($pltSL.Tag -is [array]){$pltSL.Tag = $pltSL.Tag -join '-' } ; 
            #$transcript = ".\logs\$($Ticket)-$($DomainName)-$(split-path $rMyInvocation.InvocationName -leaf)-$(get-date -format 'yyyyMMdd-HHmmtt')-trans-log.txt" ; 
            #$pltSL.Tag += "-$($DomainName)"
            #
            if($rPSBoundParameters.keys){ # alt: leverage $rPSBoundParameters hash
                $sTag = @() ; 
                #$pltSL.TAG = $((@($rPSBoundParameters.keys) |?{$_}) -join ','); # join all params
                if($rPSBoundParameters['Summary']){ $sTag+= @('Summary') } ; # build elements conditionally, string
                if($rPSBoundParameters['Number']){ $sTag+= @("Number$($rPSBoundParameters['Number'])") } ; # and keyname,value
                $pltSL.Tag += "-$($sTag -join ',')" ; # 4:46 PM 7/16/2025 flipped to append, not assign
            } ; 
            #
            if($rvEnv.isScript){
                write-host "`$script:PSCommandPath:$($script:PSCommandPath)" ;
                write-host "`$PSCommandPath:$($PSCommandPath)" ;
                if($rvEnv.PSCommandPathproxy){ $prxPath = $rvEnv.PSCommandPathproxy }
                elseif($script:PSCommandPath){$prxPath = $script:PSCommandPath}
                elseif($rPSCommandPath){$prxPath = $rPSCommandPath} ; 
            } ; 
            if($rvEnv.isFunc){
                if($rvEnv.FuncDir -AND $rvEnv.FuncName){
                       $prxPath = join-path -path $rvEnv.FuncDir -ChildPath $rvEnv.FuncName ; 
                } else {
                    write-warning "Missing either `$rvEnv.FuncDir -OR `$rvEnv.FuncName!" ; 
                } ; 
            } ; 
            if(-not $rvEnv.isFunc){
                # under funcs, this is the scriptblock of the func, not a path
                if($rvEnv.MyInvocationproxy.MyCommand.Definition){$prxPath2 = $rvEnv.MyInvocationproxy.MyCommand.Definition }
                elseif($rvEnv.MyInvocationproxy.MyCommand.Definition){$prxPath2 = $rvEnv.MyInvocationproxy.MyCommand.Definition } ; 
            } ; 
            if($prxPath){
                # 12/12/2025 new code to patch no-ext $prxPath
                if(-not [System.IO.Path]::GetExtension($prxPath)){
                    write-verbose "no-extension `$prxpath, asserting fake ext (.ps1|.psm1 as approp)" ;                         
                    switch($rvEnv.runSource){
                        'Function'{$prxPath = "$($prxPath).psm1" }
                        'ExternalScript'{$prxPath = "$($prxPath).ps1" }
                        default {
                            $smsg = "NO RECOGNIZED `$rvEnv.runSource: '$($rvEnv.runSource)'`nUNABLE TO SAFELY TEST FOR AllUsers or CU SCOPE!: ABORTING (Could log into module hosting dir!)" ; 
                            write-warning $smsg ; throw $smsg ; 
                            BREAK ; 
                        }
                    } ; 
                } ; 
                if(($prxPath -match $rgxPSAllUsersScope) -OR ($prxPath -match $rgxPSCurrUserScope)){
                    $bDivertLog = $true ; 
                    switch -regex ($prxPath){
                        $rgxPSAllUsersScope{$smsg = "AllUsers"} 
                        $rgxPSCurrUserScope{$smsg = "CurrentUser"}
                    } ;
                    $smsg += " context script/module, divert logging into [$budrv]:\scripts" 
                    write-verbose $smsg  ;
                    if($bDivertLog){
                        if((split-path $prxPath -leaf) -ne $rvEnv.CmdletName){
                            # function in a module/script installed to allusers|cu - defer name to Cmdlet/Function name
                            $pltSL.Path = (join-path -Path "$($budrv):\scripts" -ChildPath "$($rvEnv.CmdletName).ps1") ;
                        } else {
                            # installed allusers|CU script, use the hosting script name
                            $pltSL.Path = (join-path -Path "$($budrv):\scripts" -ChildPath (split-path $prxPath -leaf)) ;
                        }
                    } ;
                } else {
                    $pltSL.Path = $prxPath ;
                } ;
            }elseif($prxPath2){
                # 12/12/2025 new code to patch no-ext $prxPath2
                if(-not [System.IO.Path]::GetExtension($prxPath2)){
                    write-verbose "no-extension `$prxPath2, asserting fake ext (.ps1|.psm1 as approp)" ;                         
                    switch($rvEnv.runSource){
                        'Function'{$prxPath2 = "$($prxPath2).psm1" }
                        'ExternalScript'{$prxPath2 = "$($prxPath2).ps1" }
                        default {
                            $smsg = "NO RECOGNIZED `$rvEnv.runSource: '$($rvEnv.runSource)'`nUNABLE TO SAFELY TEST FOR AllUsers or CU SCOPE!: ABORTING (Could log into module hosting dir!)" ; 
                            write-warning $smsg ; throw $smsg ; 
                            BREAK ; 
                        }
                    } ; 
                } ; 
                if(($prxPath2 -match $rgxPSAllUsersScope) -OR ($prxPath2 -match $rgxPSCurrUserScope) ){
                        $pltSL.Path = (join-path -Path "$($budrv):\scripts" -ChildPath (split-path $prxPath2 -leaf)) ;
                } elseif(test-path $prxPath2) {
                    $pltSL.Path = $prxPath2 ;
                } elseif($rvEnv.CmdletName){
                    $pltSL.Path = (join-path -Path "$($budrv):\scripts" -ChildPath "$($rvEnv.CmdletName).ps1") ;
                } else {
                    $smsg = "UNABLE TO RESOLVE A FUNCTIONAL `$rvEnv.CmdletName, FROM WHICH TO BUILD A START-LOG.PATH!" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Warn } #Error|Warn|Debug 
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    BREAK ;
                } ; 
            } else{
                $smsg = "UNABLE TO RESOLVE A FUNCTIONAL `$rvEnv.CmdletName, FROM WHICH TO BUILD A START-LOG.PATH!" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Warn } #Error|Warn|Debug 
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                BREAK ;
            }  ;
            write-verbose "start-Log w`n$(($pltSL|out-string).trim())" ; 
            $logspec = start-Log @pltSL ;
            $error.clear() ;
            TRY {
                if($logspec){
                    $logging=$logspec.logging ;
                    $logfile=$logspec.logfile ;
                    $transcript=$logspec.transcript ;
                    $stopResults = try {Stop-transcript -ErrorAction stop} catch {} ;
                    if($stopResults){
                        $smsg = "Stop-transcript:$($stopResults)" ; 
                        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                    } ; 
                    $startResults = start-Transcript -path $transcript -whatif:$false -confirm:$false;
                    if($startResults){
                        $smsg = "start-transcript:$($startResults)" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ; 
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

                #region SendMailAlert ; #*------v SendMailAlert v------
                $SmtpBody += "`n===FAIL Summary:" ;
                $SmtpBody += "`n$('-'*50)" ;
                $SmtpBody += "`n$('-'*50)" ;
                $smsg += "`n$(($smsg |out-string).trim())" ; 
                $sdEmail = @{
                    smtpFrom = $SMTPFrom ;
                    SMTPTo = $SMTPTo ;
                    SMTPSubj = $SMTPSubj ;
                    #SMTPServer = $SMTPServer ;
                    SmtpBody = $SmtpBody ;
                    SmtpAttachment = $SmtpAttachment ;
                    BodyAsHtml = $false ; # let the htmltag rgx in Send-EmailNotif flip on as needed
                    verbose = $($VerbosePreference -eq "Continue") ;
                } ;
                $smsg = "Send-EmailNotif w`n$(($sdEmail|out-string).trim())" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                Send-EmailNotif @sdEmail ;

                #endregion SendMailAlert ; #*------^ END SendMailAlert ^------
            } ;
        } ; 
        #endregion START_LOG_HOLISTIC #*------^ END START_LOG_HOLISTIC ^------
        # ...
        #endregion START_LOG_OPTIONS #*======^ START_LOG_OPTIONS ^======

        #region NETWORK_INFO ; #*======v NETWORK_INFO v======
        $netsettings = resolve-NetworkLocalTDO ; 
        #endregion NETWORK_INFO ; #*======^ END NETWORK_INFO ^======
    
        <#
        $useO365 = $true ;
        $useEXO = $true ; 
        $UseOP=$true ; 
        $UseExOP=$true ;
        $useExopNoDep = $true ; # switch to use Connect-ExchangeServerTDO, vs connect-ex2010 (creds are assumed inherent to the account)
        $ExopVers = 'Ex2010' # 'Ex2019','Ex2016','Ex2013','Ex2010','Ex2007','Ex2003','Ex2000', Null for All versions
        if($Version){
            $ExopVers = $Version ; #defer to local script $version if set
        } ; 
        $useForestWide = $true ; # flag to trigger cross-domain/forest-wide code in AD & EXoP
        $UseOPAD = $false ; 
        $UseMSOL = $false ; # should be hard disabled now in o365
        $UseAAD = $true ; 
        #>

        #region SERVICE_CONNECTIONS #*======v END SERVICE_CONNECTIONS v======
    
        #region BROAD_SVC_CONTROL_VARIS ; #*======v BROAD_SVC_CONTROL_VARIS  v======   
        $useO365 = $true ; 
        $useOP = $true ;     
        # (config individual svcs in each block)
        #endregion BROAD_SVC_CONTROL_VARIS ; #*======^ END BROAD_SVC_CONTROL_VARIS ^======

        #region CALL_CONNECT_O365SERVICES ; #*======v CALL_CONNECT_O365SERVICES v======
        #$useO365 = $true ; 
        if($useO365){
            $pltCco365Svcs=[ordered]@{
                # environment parameters:
                EnvSummary = $rvEnv ; 
                NetSummary = $netsettings ; 
                # service choices
                useEXO = $true ;
                useSC = $false ; 
                UseMSOL = $false ;
                UseAAD = $false ; # M$ is actively blocking all AAD access now: Message: Access blocked to AAD Graph API for this application. https://aka.ms/AzureADGraphMigration.
                UseMG = $true ;
                # Service Connection parameters
                TenOrg = $TenOrg ; # $global:o365_TenOrgDefault ; 
                Credential = $Credential ;
                AdminAccount = $AdminAccount ; 
                #[ValidateSet("SID","CSID","UID","B2BI","CSVC","ESVC","LSVC","ESvcCBA","CSvcCBA","SIDCBA")]
                UserRole = $UserRole ; # @('SID','CSVC') ;
                # svcAcct use: @('ESvcCBA','CSvcCBA','SIDCBA')
                silent = $silent ;
                MGPermissionsScope = $MGPermissionsScope ;
                MGCmdlets = $MGCmdlets ;
            } ;
            write-verbose "(Purge no value keys from splat)" ; 
            $mts = $pltCco365Svcs.GetEnumerator() |?{$_.value -eq $null} ; $mts |%{$pltCco365Svcs.remove($_.Name)} ; rv mts -ea 0 ; 
            if((get-command connect-O365Services -EA STOP).parameters.ContainsKey('whatif')){
                $pltCco365SvcsnDSR.add('whatif',$($whatif))
            } ; 
            $smsg = "connect-O365Services w`n$(($pltCco365Svcs|out-string).trim())" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            # add rertry on fail, up to $DoRetries
            $Exit = 0 ; # zero out $exit each new cmd try/retried
            # do loop until up to 4 retries...
            Do {
                $smsg = "connect-O365Services w`n$(($pltCco365Svcs|out-string).trim())" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                $ret_ccSO365 = connect-O365Services @pltCco365Svcs ; 
                #region CONFIRM_CCEXORETURN ; #*------v CONFIRM_CCEXORETURN v------
                # matches each: $plt.useXXX:$true to matching returned $ret.hasXXX:$true 
                $vplt = $pltCco365Svcs ; $vret = 'ret_ccSO365' ; $ACtionCommand = 'connect-O365Services' ; $vtests = @() ; $vFailMsgs = @()  ; 
                $vplt.GetEnumerator() |?{$_.key -match '^use' -ANd $_.value -match $true} | foreach-object{
                    $pltkey = $_ ;
                    $smsg = "$(($pltkey | ft -HideTableHeaders name,value|out-string).trim())" ; 
                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                    $tprop = $pltkey.name -replace '^use','has';
                    if($rProp = (gv $vret).Value.psobject.properties | ?{$_.name -match $tprop}){
                        $smsg = "$(($rprop | ft -HideTableHeaders name,value |out-string).trim())" ; 
                        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                        if($rprop.Value -eq $pltkey.value){
                            $vtests += $true ; 
                            $smsg = "Validated: $($pltKey.name):$($pltKey.value) => $($rprop.name):$($rprop.value)" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Success } 
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                        } else {
                            $smsg = "NOT VALIDATED: $($pltKey.name):$($pltKey.value) => $($rprop.name):$($rprop.value)" ;
                            $vtests += $false ; 
                            $vFailMsgs += "`n$($smsg)" ; 
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                        };
                    } else{
                        $smsg = "Unable to locate: $($pltKey.name):$($pltKey.value) to any matching $($rprop.name)!)" ;
                        $smsg = "" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                    } ; 
                } ; 
                if($vtests -notcontains $false){
                    $smsg = "==> $($ACtionCommand): confirmed specified connections *all* successful " ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Success } 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                    $Exit = $DoRetries ;
                } else {
                    $smsg = "==> $($ACtionCommand): FAILED SOME SPECIFIED CONNECTIONS" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                    $smsg = "MISSING SOME KEY CONNECTIONS. DO YOU WANT TO IGNORE, AND CONTINUE WITH CONNECTED SERVICES?" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                    $Exit ++ ;
                    $smsg = "Try #: $Exit" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Warn } 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                    if($Exit -eq $DoRetries){
                        $smsg = "Unable to exec cmd!"; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                        #-=-=-=-=-=-=-=-=
                        $sdEmail.SMTPSubj = "FAIL Rpt:$($ScriptBaseName):$(get-date -format 'yyyyMMdd-HHmmtt')"
                        $sdEmail.SmtpBody = "`n===Processing Summary:" ;
                        if($vFailMsgs){
                            $sdEmail.SmtpBody += "`n$(($vFailMsgs|out-string).trim())" ; 
                        } ; 
                        $sdEmail.SmtpBody += "`n" ;
                        if($SmtpAttachment){
                            $sdEmail.SmtpAttachment = $SmtpAttachment
                            $sdEmail.smtpBody +="`n(Logs Attached)" ;
                        };
                        $sdEmail.SmtpBody += "Pass Completed $([System.DateTime]::Now)" ;
                        $smsg = "Send-EmailNotif w`n$(($sdEmail|out-string).trim())" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        Send-EmailNotif @sdEmail ;
                        $bRet=Read-Host "Enter YYY to continue. Anything else will exit"  ;
                        if ($bRet.ToUpper() -eq "YYY") {
                            $smsg = "(Moving on), WITH THE FOLLOW PARTIAL CONNECTION STATUS" ;
                            $smsg += "`n`n$(($ret_CcOPSvcs|out-string).trim())" ; 
                            write-host -foregroundcolor green $smsg  ;
                        } else {
                            throw $smsg ; 
                            break ; #exit 1
                        } ;  
                    } ;        
                } ; 
                #endregion CONFIRM_CCEXORETURN ; #*------^ END CONFIRM_CCEXORETURN ^------
            } Until ($Exit -eq $DoRetries) ; 
        } ; #  useO365-E
        #endregion CALL_CONNECT_O365SERVICES ; #*======^ END CALL_CONNECT_O365SERVICES ^======
    
        #region TEST_EXO_CONN ; #*------v TEST_EXO_CONN v------
        # ALT: simplified verify EXO conn: ALT to full CONNECT_O365SERVICES block - USE ONE OR THE OTHER!
        $useEXO = $true ; 
        $useSC = $false ; 
        if(-not $XOConnectionUri ){$XOConnectionUri = 'https://outlook.office365.com'} ;
        if(-not $SCConnectionUri){$SCConnectionUri = 'https://ps.compliance.protection.outlook.com'} ;
        $EXOtestCmdlet = 'Get-xoOrganizationConfig' ; 
        if(gcm $EXOtestCmdlet -ea 0){
            $conns = Get-ConnectionInformation -ea STOP  ; 
            $hasEXO = $hasSC = $false ; 
            #if($conns | %{$_ | ?{$_.ConnectionUri -eq 'https://outlook.office365.com' -AND $_.State -eq 'Connected' -AND $_.TokenStatus -eq 'Active'}}){
            $conns | %{
                if($_ | ?{$_.ConnectionUri -eq $XOConnectionUri}){$hasEXO = $true } ; 
                if($_ | ?{$_.ConnectionUri -eq $SCConnectionUri}){$hasSC = $true } ; 
            }
            if($useEXO -AND $hasEXO){
                write-verbose "EXO ConnectionURI present" ; 
            }elseif(-not $useEXO){}else{
                $smsg = "No Active EXO connection: Run - Connect-ExchangeOnline -Prefix xo -  before running this script!" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                BREAK ; 
            } ; 
            if($useSC -AND $hasSC){
                write-verbose "SCI ConnectionURI present" ; 
            }elseif(-not $useSC){}else{
                $smsg = "No Active SC connection: Run - Connect-IPPSSession -Prefix SC -  before running this script!" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                BREAK ; 
            } ; 
        }else {
            $smsg = "Missing gcm get-xoMailboxFolderStatistics: ExchangeOnlineManagement module *not* loaded!" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
            BREAK ; 
        } ;     
        #endregion TEST_EXO_CONN ; #*------^ END TEST_EXO_CONN ^------
    
        #region CALL_CONNECT_OPSERVICES ; #*======v CALL_CONNECT_OPSERVICES v======
        #$useOP = $false ; 
        if($useOP){
            $pltCcOPSvcs=[ordered]@{
                # environment parameters:
                EnvSummary = $rvEnv ;
                NetSummary = $netsettings ;
                XoPSummary = $lclExOP ;
                # service choices
                UseExOP = $true ;
                useForestWide = $true ;
                useExopNoDep = $false ;
                ExopVers = 'Ex2010' ;
                UseOPAD = $true ;
                useExOPVers = $useExOPVers; # 'Ex2010' ;
                # Service Connection parameters
                TenOrg = $TenOrg ; # $global:o365_TenOrgDefault ;
                Credential = $Credential ;
                #[ValidateSet("SID","ESVC","LSVC")]
                #UserRole = $UserRole ; # @('SID','ESVC') ;
                # if inheriting same $userrole param/default, that was already used for cloud conn, filter out the op unsupported CBA roles
                # exclude csvc as well, go with filter on the supported ValidateSet from get-HybridOPCredentials: ESVC|LSVC|SID
                UserRole = ($UserRole -match '(ESVC|LSVC|SID)' -notmatch 'CBA') ; # @('SID','ESVC') ;
                # svcAcct use: @('ESvcCBA','CSvcCBA','SIDCBA')
                silent = $silent ;
            } ;

            write-verbose "(Purge no value keys from splat)" ;
            $mts = $pltCcOPSvcs.GetEnumerator() |?{$_.value -eq $null} ; $mts |%{$pltCcOPSvcs.remove($_.Name)} ; rv mts -ea 0 ;
            if((get-command connect-OPServices -EA STOP).parameters.ContainsKey('whatif')){
                $pltCcOPSvcsnDSR.add('whatif',$($whatif))
            } ;
            $smsg = "connect-OPServices w`n$(($pltCcOPSvcs|out-string).trim())" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            $ret_CcOPSvcs = connect-OPServices @pltCcOPSvcs ; 

            # #region CONFIRM_CCOPRETURN ; #*------v CONFIRM_CCOPRETURN v------
            # matches each: $plt.useXXX:$true to matching returned $ret.hasXXX:$true
            $vplt = $pltCcOPSvcs ; $vret = 'ret_CcOPSvcs' ;  ; $ACtionCommand = 'connect-OPServices' ; 
            $vplt.GetEnumerator() |?{$_.key -match '^use' -ANd $_.value -match $true} | foreach-object{
                $pltkey = $_ ;
                $smsg = "$(($pltkey | ft -HideTableHeaders name,value|out-string).trim())" ; 
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                $vtests = @() ;  $vFailMsgs = @()  ; 
                $tprop = $pltkey.name -replace '^use','has';
                if($rProp = (gv $vret).Value.psobject.properties | ?{$_.name -match $tprop}){
                    $smsg = "$(($rprop | ft -HideTableHeaders name,value |out-string).trim())" ; 
                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                    if($rprop.Value -eq $pltkey.value){
                        $vtests += $true ; 
                        $smsg = "Validated: $($pltKey.name):$($pltKey.value) => $($rprop.name):$($rprop.value)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Success } 
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                    } else {
                        $smsg = "NOT VALIDATED: $($pltKey.name):$($pltKey.value) => $($rprop.name):$($rprop.value)" ;
                        $vtests += $false ; 
                        $vFailMsgs += "`n$($smsg)" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                    };
                } else{
                    $smsg = "Unable to locate: $($pltKey.name):$($pltKey.value) to any matching $($rprop.name)!)" ;
                    $smsg = "" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                } ; 
            } ; 
            if($useOP -AND $vtests -notcontains $false){
                $smsg = "==> $($ACtionCommand): confirmed specified connections *all* successful " ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Success } 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
            }elseif($vtests -contains $false -AND (get-variable ret_CcOPSvcs) -AND (gv -name "$($tenorg)meta").value.o365_opdomain.split('.')[0].toupper() -ne $env:userdomain){
                $smsg = "==> $($ACtionCommand): FAILED SOME SPECIFIED CONNECTIONS" ; 
                $smsg += "`nCROSS-ORG ONPREM CONNECTION: ATTEMPTING TO CONNECT TO ONPREM '$((gv -name "$($tenorg)meta").value.o365_Prefix)' $((gv -name "$($tenorg)meta").value.o365_opdomain.split('.')[0].toupper()) domain, FROM $($env:userdomain)!" ;
                $smsg += "`nEXPECTED ERROR, SKIPPING ONPREM ACCESS STEPS (force `$useOP:$false)" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                $useOP = $false ; 
            }elseif(-not $useOP -AND -not (get-variable ret_CcOPSvcs)){
                $smsg = "-useOP: $($useOP), skipped connect-OPServices" ; 
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            } else {
                $smsg = "==> $($ACtionCommand): FAILED SOME SPECIFIED CONNECTIONS" ; 
                $smsg += "`n`$ret_CcOPSvcs:`n$(($ret_CcOPSvcs|out-string).trim())" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                $sdEmail.SMTPSubj = "FAIL Rpt:$($ScriptBaseName):$(get-date -format 'yyyyMMdd-HHmmtt')"
                $sdEmail.SmtpBody = "`n===Processing Summary:" ;
                if($vFailMsgs){
                    $sdEmail.SmtpBody += "`n$(($vFailMsgs|out-string).trim())" ; 
                } ; 
                $sdEmail.SmtpBody += "`n" ;
                if($SmtpAttachment){
                    $sdEmail.SmtpAttachment = $SmtpAttachment
                    $sdEmail.smtpBody +="`n(Logs Attached)" ;
                };
                $sdEmail.SmtpBody += "Pass Completed $([System.DateTime]::Now)" ;
                $smsg = "Send-EmailNotif w`n$(($sdEmail|out-string).trim())" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                Send-EmailNotif @sdEmail ;
                throw $smsg ; 
                BREAK ; 
            } ; 
            #endregion CONFIRM_CCOPRETURN ; #*------^ END CONFIRM_CCOPRETURN ^------
            
            #region CONFIRM_OPFORESTWIDE ; #*------v CONFIRM_OPFORESTWIDE v------    
            if($useOP -AND $pltCcOPSvcs.useForestWide -AND $ret_CcOPSvcs.hasForestWide -AND $ret_CcOPSvcs.AdGcFwide){
                $smsg = "==> $($ACtionCommand): confirmed has BOTH .hasForestWide & .AdGcFwide ($($ret_CcOPSvcs.AdGcFwide))" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Success } 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success        
            }elseif($pltCcOPSvcs.useForestWide -AND (get-variable ret_CcOPSvcs) -AND (gv -name "$($tenorg)meta").value.o365_opdomain.split('.')[0].toupper() -ne $env:userdomain){
                $smsg = "`nCROSS-ORG ONPREM CONNECTION: ATTEMPTING TO CONNECT TO ONPREM '$((gv -name "$($tenorg)meta").value.o365_Prefix)' $((gv -name "$($tenorg)meta").value.o365_opdomain.split('.')[0].toupper()) domain, FROM $($env:userdomain)!" ;
                $smsg += "`nEXPECTED ERROR, SKIPPING ONPREM FORESTWIDE SPEC" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                $useOP = $false ; 
            }elseif($useOP -AND $pltCcOPSvcs.useForestWide -AND -NOT $ret_CcOPSvcs.hasForestWide){
                $smsg = "==> $($ACtionCommand): MISSING CRITICAL FORESTWIDE SUPPORT COMPONENT:" ; 
                if(-not $ret_CcOPSvcs.hasForestWide){
                    $smsg += "`n----->$($ACtionCommand): MISSING .hasForestWide (Set-AdServerSettings -ViewEntireForest `$True) " ; 
                } ; 
                if(-not $ret_CcOPSvcs.AdGcFwide){
                    $smsg += "`n----->$($ACtionCommand): MISSING .AdGcFwide GC!:`n((Get-ADDomainController -Discover -Service GlobalCatalog).hostname):326) " ; 
                } ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                $smsg = "MISSING SOME KEY CONNECTIONS. DO YOU WANT TO IGNORE, AND CONTINUE WITH CONNECTED SERVICES?" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                $bRet=Read-Host "Enter YYY to continue. Anything else will exit"  ;
                if ($bRet.ToUpper() -eq "YYY") {
                    $smsg = "(Moving on), WITH THE FOLLOW PARTIAL CONNECTION STATUS" ;
                    $smsg += "`n`n$(($ret_CcOPSvcs|out-string).trim())" ; 
                    write-host -foregroundcolor green $smsg  ;
                } else {
                    throw $smsg ; 
                    break ; #exit 1
                } ;         
            }; 
            #endregion CONFIRM_OPFORESTWIDE ; #*------^ END CONFIRM_OPFORESTWIDE ^------
        } ; 
        #endregion CALL_CONNECT_OPSERVICES ; #*======^ END CALL_CONNECT_OPSERVICES ^======
    
        #endregion SERVICE_CONNECTIONS #*======^ END SERVICE_CONNECTIONS ^======

      
        # Configure the Get-xoMessageTrace splat 
        <# gxmt v1 params
        $pltGXMT=[ordered]@{
            Page= 1 ; # default it to 1 vs $null as we'll be purging empties further down
            ErrorAction = 'STOP' ;
            verbose = $($VerbosePreference -eq "Continue") ;
        } ;
        #>
        # 12:12 PM 10/21/2025 Get-xoMessageTraceV2 params - tossed out Page support and all native pagination
        #-ResultSize 5000 -StartDate $StartDate -EndDate $EndDate -WarningVariable MoreResultsAvailable 
        $pltGXMT=[ordered]@{
            #Page= 1 ; # default it to 1 vs $null as we'll be purging empties further down
            #ResultSize = 5000 ; 
            WarningVariable = "MoreResultsAvailable" ; 
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
            #$tendoms=Get-AzureADDomain ; 
            $tendoms = Get-MgDomain 
            #$Ten = ($tendoms |?{$_.name -like '*.mail.onmicrosoft.com'}).name.split('.')[0] ; #aad
            $Ten = ($tendoms |?{$_.id -like '*.mail.onmicrosoft.com'}).id.split('.')[0] ; # mg, why keep the same property, name == id? cuz fu!
            $Ten = "$($Ten.substring(0,1).toupper())$($Ten.substring(1,$Ten.length-1).toLower())"
        }CATCH{
            $smsg = "NOT MG CONNECTED! (dep: Get-MgDomain)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
          BREAK 
        } ;
       

    }  # BEG-E
    PROCESS {
        #region SPLAT_BUILD ; #*------v SPLAT_BUILD v------        
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
        # 4:01 PM 10/21/2025 add: $ResultSize,
        if($ResultSize){
            $pltGXMT.add('ResultSize',$ResultSize) ; 
        } else{
            # use default 5k msg limt, max allowed size by MS; given 100qrys/5min window throttling, pays to get max out of each qry.
            $pltGXMT.add('ResultSize',5000) ; 
        };
        # 12:28 PM 10/21/2025 new gxmtV2 -subject & SubjectFilterType params
        if($subject){
            $pltGXMT.add('Subject',$subject) ; 
        } ;
        <# You specify how the value is evaluated in the message subject by using the SubjectFilterType parameter.
            -SubjectFilterType
                The SubjectFilterType parameter specifies how the value of the Subject parameter is evaluated. Valid values are:
                    Contains
                    EndsWith
                    StartsWith
                We recommend using StartsWith or EndsWith instead of Contains whenever possible.
        #>
        if($SubjectFilterType){
            $pltGXMT.add('Subject',$SubjectFilterType) ; 
        } ;

        #endregion SPLAT_BUILD ; #*------^ END SPLAT_BUILD ^------

        # use the updated psOfile build:
        #-=-=-=-=-=-=-=-=
        #region MSGTRKFILENAME ; #*------v MSGTRKFILENAME v------
        write-verbose "Keys off of typical msgtrk inputsplat" ; 
        
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

        $statusLookupTbl = Initialize-exoStatusTable ;         

        #$ofile=[RegEx]::Replace($ofile, "[{0}]" -f ([RegEx]::Escape(-join [System.IO.Path]::GetInvalidFileNameChars())), '') ;
        # use the tested redirected $logfile path
        #$ofile = join-path -path (split-path $logfile) -ChildPath $ofile ; 
        $hReports = [ordered]@{} ; 
        #rxo ;
        $error.clear() ;
        $Exit = 0 ;
        # do retry the initial query
        Do {
            TRY {
                # prepurge empty hash value keys:
                #$pltGXMT=$pltGXMT.GetEnumerator()|? value ;
                # remove null keyed objects
                #$pltGXMT | Foreach {$p = $_ ;@($p.GetEnumerator()) | ?{ ($_.Value | Out-String).length -eq 0 } | Foreach-Object {$p.Remove($_.Key)} ;} ;
                # skip it, we're only adding populated items now
                #write-verbose "hashtype:$($pltGXMT.GetType().FullName)" ; 
                # and issue was first untested negative integer -Days; and 2nd GMT window for start/enddate, so the 'local' input needs to be converted to/from gmt to get the targeted content.

                <# as of 9:56 AM 10/21/2025: Microsoft is killing all requests to Get-xoMessageTrace, now returns:
                Write-ErrorMessage : ||Get-MessageTrace will start deprecating on September 1st, 2025. Please refer to: https://learn.microsoft.com/en-us/powershell/module/exchange/get-messagetracev2?view=exchange-ps to switch to 
                Get-MessageTraceV2.
                At C:\Users\kadriTSS\AppData\Local\Temp\2\tmpEXO_ah0pg1qz.hjm\tmpEXO_ah0pg1qz.hjm.psm1:1191 char:13
                +             Write-ErrorMessage $ErrorObject
                +             ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                    + CategoryInfo          : InvalidOperation: (:) [Get-MessageTrace], ValidationException
                    + FullyQualifiedErrorId : [Server=CY3PR04MB9691,RequestId=8c72957c-e95b-a089-8433-7f5e0e74f247,TimeStamp=Tue, 21 Oct 2025 14:40:45 GMT],Write-ErrorMessage
                [Announcing General Availability (GA) of the New Message Trace in Exchange Online | Microsoft Community Hub](https://techcommunity.microsoft.com/blog/exchange/announcing-general-availability-ga-of-the-new-message-trace-in-exchange-online/4420243)

                [New Message trace in EAC in Exchange Online | Microsoft Learn](https://learn.microsoft.com/en-us/exchange/monitoring/trace-an-email-message/new-message-trace)

                [Using the Get-MessageTraceV2 cmdlet to generate mail traffic statistics by user - Blog](https://michev.info/blog/post/6572/using-the-get-messagetracev2-cmdlet-to-generate-mail-traffic-statistics-by-user)

                [PowerShell/Get-DetailedMessageStatsV2.ps1 at master · michevnew/PowerShell · GitHub](https://github.com/michevnew/PowerShell/blob/master/Get-DetailedMessageStatsV2.ps1)
                
                # Changes: throttling, torn out usefulness, from Michev:

                ### Positives:
                - Let's start with the good news. The new message trace ensures feature parity 
                    with the "old" experience, while bringing some nice improvements. The most 
                    significant of these is the support for querying data up to 90 days in the past 
                    in a synchronous manner, whereas we were previously limited to just a handful 
                    of days, and had to run async queries. Up to 10 days of data is available for a 
                    single query, but this should not be a problem as in any tenant of meaningful 
                    size, you will hit the "page" limit early on and will have to issue additional 
                    queries anyway. 
                -  larger set of filters supported for the Get-MessageTraceV2 cmdlet, most 
                    notably the ability to filter based on a message's subject, made possible 
                    thanks to the -Subject parameter. 
                ### Negatives:
                - Microsoft changed the way "pagination" works, and the new experience is mildly 
                    annoying at best. It could have been implemented better IMO, or at the very 
                    least align with the pagination experience in the Graph API. Speaking of which, 
                    the next negative is the lack of support for Graph API endpoints/methods. One 
                    can argue that we do have a suitable replacement on the Graph via the 
                    **analyzedEmails** endpoint, but as we discussed in [our 
                    article](https://www.michev.info/blog/post/6181/first-look-at-the-analyzedemails-graph-api-endpoint) 
                    on said endpoint, it's use case is different. 
                -  incoming deprecation of the old experience. Microsoft is giving customers 
                    until September 1st to move away from the 
                    Get-MessageTrace/Get-MessageTraceDetail cmdlets. The same deadline applies to 
                    the MessageTrace report in the good old reporting web service, which is the 
                    only supported RESTful interface to query the message trace data programmatically.
                    No alternative is provided at this point, thus any customers 
                    and ISVs that still rely on the reporting web service need to move to using 
                    PowerShell instead, which might be an issue. The same can be safe in regard to 
                    the new throttling guidance, namely 100 requests per 5 minutes. 

                > The biggest change in the script's logic is in how it handles pagination. 
                I've opted for an approach that relies on the presence of the "hint" returned 
                by the service, which unfortunately is implemented via the warning stream. In 
                effect, we suppress the warning while making sure its content is stored in a 
                variable and then processed to extract the "next page" cmdlet. I'm not a big 
                fan of this implementation, as you can see from [my 
                comments](https://techcommunity.microsoft.com/blog/exchange/announcing-public-preview-of-the-new-message-trace-in-exchange-online/4356561/replies/4392248) 
                under the original blog article. The alternative is to create the "next page" 
                syntax yourself, by copying the properties of the last returned entry, like 
                Tony does in his [sample 
                script](https://github.com/12Knocksinna/Office365itpros/blob/master/Analyze-MailTraffic.PS1)

                > 
                > Another downside of the new "no pagination" approach is that we cannot have a 
                proper progress indicator, so instead I have added a "poor man's" variation of 
                it, just so you know whether the script is progressing. Once we fetch the 
                available message trace data, we largely follow the logic of the original 
                script and prepare a hashtable for each recipient, holding the count and size 
                of both inbound and outbound messages, per day. Lastly, we transform the output 
                and dump it into a CSV file in the working directory. And since HTML output 
                seems to be all the rage currently, I've asked Copilot to generate the 
                corresponding code. Can I play with the cool kids now? 🙂 

                TK: REQUIRES LATER REV
                10:18 AM 10/21/2025 jb running:
                #-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
                [PS]:D:\s\build $ Gcm Get-xoMessageTracev2

                CommandType     Name                                               Version    Source                                                                                                                                            
                -----------     ----                                               -------    ------                                                                                                                                            
                Function        Get-xoMessageTraceV2                               1.0        tmpEXO_ah0pg1qz.hjm

                [PS]:D:\s\build $ gmo exchangeonlinemanagement

                ModuleType Version    Name                                ExportedCommands                                                                                                                                                      
                ---------- -------    ----                                ----------------                                                                                                                                                      
                Script     3.6.0      ExchangeOnlineManagement            {Add-VivaModuleFeaturePolicy, Get-ConnectionInformation, Get-DefaultTenantBriefingConfig, Get-DefaultTenantMyAnalyticsFeatureConfig...}

                Curr Online rev:
                [PowerShell Gallery | ExchangeOnlineManagement 3.9.0](https://www.powershellgallery.com/packages/ExchangeOnlineManagement/3.9.0)

                    63,387,978 Downloads
                    12,868 Downloads of 3.9.1-Preview1
                    8/13/2025 Last Published

                Version                 | Downloads | Last updated
                ----------------------- | --------- | -------------
                3.9.1-Preview1          | 12,868    | a month ago
                3.9.0 (current version) | 1,424,135 | 2 months ago <=== 2 mos old!
                3.9.0-Preview1          | 9,617     | 3 months ago
                3.8.1-Preview1          | 48,821    | 5 months ago
                3.8.0                   | 2,050,555 | 5 months ago
                3.8.0-Preview2          | 14,309    | 6 months ago
                3.8.0-Preview1          | 10,600    | 7 months ago
                3.7.2                   | 1,234,251 | 7 months ago
                3.7.2-Preview1          | 6,924     | 8 months ago
                3.7.1                   | 3,181,614 | 9 months ago
                3.7.1-Preview1          | 9,542     | 10 months ago
                3.7.0                   | 2,950,466 | 12/2/2024
                3.7.0-Preview1          | 4,938     | 11/15/2024
                3.6.0                   | 3,537,886 | 9/25/2024

                #-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=

            #>
                <# my prior Get-xoMessageTrace code:
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
                #>

                # michev's code:
                #$MailTraffic = @{} ; 
                # static 10d window
                #$StartDate = (Get-Date).AddDays(-10) #max period we can cover in a single query is 10 days, if needed rerun multiple times to cover up to 90
                #$EndDate = (Get-Date)
                #Get the first "page"
                $Msgs = $null # aggregator
                #$PageMsgs = Get-xoMessageTraceV2 -ResultSize 5000 -StartDate $StartDate -EndDate $EndDate -WarningVariable MoreResultsAvailable -Verbose:$false 3>$null
                #pltGXMT
                $smsg = "Get-xoMessageTrace  w`n$(($pltGXMT|out-string).trim())" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                $PageMsgs = Get-xoMessageTraceV2 @pltGXMT 3>$null
                # The expression 3>$null in PowerShell redirects the Warning stream (stream number 3) to $null. required, because mich is using the warning to note restart points
                                                                                                <# Michev asked M$:
            [Announcing Public Preview of the New Message Trace in Exchange Online | Microsoft Community Hub](https://techcommunity.microsoft.com/blog/exchange/announcing-public-preview-of-the-new-message-trace-in-exchange-online/4356561/replies/4392248)

            Thanks. I see a new warning being generated now, is this the supposed "hint"?

            PS> Get-MessageTracev2 -resultsize 1 ; 
            out:> WARNING: There are more results, use the following command to get more. Get-MessageTraceV2 -StartDate "xxx" -EndDate "xxx" -StartingRecipientAddress "vasil@michev.info" 

            If so, might I suggest using a format that does not require any additional 
            transformation, for example by returning a separate array element with just the 
            "next page" cmdlet syntax? Or by adding a "dummy" entry to the general output 
            stream instead of using the warning one? I.e. like this: 

            Or just do it the "Graph way" with @odata.nextpage and $count?

            [YunjieCao](https://techcommunity.microsoft.com/users/yunjiecao/2896050)
            to VasilMichev

            Mar 12, 2025

            Hi,

            Yes, that is the hint we provide. Thank you for your suggestions! We value this 
            feedback and will work on improvements after gathering input from all our customers.
            In the meantime, you can refer to the FAQ section **How could 
            pagination from V1 be achieved in V2?**. If you see a warning message, you can 
            use the scripts provided in the FAQ section to compose your queries without 
            depending on the warning message itself. 

            #>
                #$Msgs += $PageMsgs | Select Received,SenderAddress,RecipientAddress,Size,Status
                # splice over my postfilter system messages removal
                $Msgs += $PageMsgs | ?{$_.SenderAddress -notlike '*micro*' -or $_.SenderAddress -notlike '*root*' } ; 
                $Exit = $Retries ;
            } CATCH [System.Exception] {
                $ErrTrapd=$Error[0] ;
                if($ErrTrapd.Exception -match $rgxEXOThrottle){
                    $smsg = "MS 100-qry limit/5mins throttling detected, waiting $(RetryThrottle)s to retry..." ; 
                    $smsg += "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                    #Start-Sleep -Milliseconds $ThrottleMs 
                    start-sleepcountdown -seconds $RetryThrottle -Rolling ; 
                    $Exit ++ ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN }
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    $smsg= "Try #: $($Exit)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Error }  #Error|Warn|Debug 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    If ($Exit -eq $Retries) {
                        $smsg= "Unable to exec cmd!" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Error }  #Error|Warn|Debug 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        BREAK ; 
                    } ;
                } else{
                    # different error, throw to the main catch
                    throw $ErrTrapd
                } ;     
            } CATCH {
                $ErrTrapd=$Error[0] ;
                $smsg = "$('*'*5)`nFailed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: `n$(($ErrTrapd|out-string).trim())`n$('-'*5)" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                # it's not outputting the underlying cmdlet error, try to force it :
                $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
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
        } Until ($Exit -eq $Retries) ; 
        # another do/try on the more's
        #If more results are available, as indicated by the presence of the WarningVariable, we need to loop until we get all results
        if ($MoreResultsAvailable) {
            $Exit = 0 ;
            # moreresults loop
            Do {
                # exit retries loop
                $Exit = 0 ;
                Do {
                    TRY {
                        #As we don't have a clue how many pages we will get, proper progress indicator is not feasible.
                        Write-Host "." -NoNewline
                        #Handling this via Warning output is beyong annoying...
                        $NextPage = ($MoreResultsAvailable -join "").TrimStart("There are more results, use the following command to get more. ")
                        # note the above lacks the PREFIX!, patch it in
                        $NextPage = ($MoreResultsAvailable -join "").TrimStart("There are more results, use the following command to get more. ") -replace 'Get-MessageTraceV2','Get-xoMessageTraceV2' ; 
                        $ScriptBlock = [ScriptBlock]::Create($NextPage)
                        $smsg = "MORE:($($Msgs.count)):$(($ScriptBlock|out-string).trim())" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                        else{ write-host "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                        $PageMsgs = Invoke-Command -ScriptBlock $ScriptBlock -WarningVariable MoreResultsAvailable -Verbose:$false 3>$null #MUST PASS WarningVariable HERE OR IT WILL NOT WORK
                        #$Msgs += $PageMsgs | Select Received,SenderAddress,RecipientAddress,Size,Status
                        # splice over my postfilter system messages removal 
                        #$Msgs += $PageMsgs | ?{$_.SenderAddress -notlike '*micro*' -or $_.SenderAddress -notlike '*root*' }
                        # Remove Exchange Online public folder hierarchy synchronization messages
                        # $Messages = $Messages | Where-Object {$_.Subject -NotLike "*HierarchySync*"}
                        $Msgs += $PageMsgs | ?{$_.SenderAddress -notlike '*micro*' -OR $_.SenderAddress -notlike '*root*' -AND ($_.Subject -NotLike "*HierarchySync*") }
                        if($MoreResultsAvailable.Count -eq 0){
                            # it didn't break on a repeat, could be the inner loop is preventing the outer until from triggering, so do it inside here.
                            Break ; 
                        } ; 
                    } CATCH [System.Exception] {
                        $ErrTrapd=$Error[0] ;
                        if($ErrTrapd.Exception -match $rgxEXOThrottle){
                            $smsg = "MS 100-qry limit/5mins throttling detected, waiting $(RetryThrottle)s to retry..." ; 
                            $smsg += "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                            #Start-Sleep -Milliseconds $ThrottleMs 
                            start-sleepcountdown -seconds $RetryThrottle -Rolling ; 
                            $Exit ++ ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN }
                            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            $smsg= "Try #: $($Exit)" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Error }  #Error|Warn|Debug 
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            If ($Exit -eq $Retries) {
                                $smsg= "Unable to exec cmd!" ;
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Error }  #Error|Warn|Debug 
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                BREAK ; 
                            } ;
                        } else{
                            # different error, throw to the main catch
                            throw $ErrTrapd
                        } ;     
                    } CATCH {
                        $ErrTrapd=$Error[0] ;
                        $smsg = "$('*'*5)`nFailed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: `n$(($ErrTrapd|out-string).trim())`n$('-'*5)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                        else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        # it's not outputting the underlying cmdlet error, try to force it :
                        $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
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
                } Until ($Exit -eq $Retries) ; 
            }until ($MoreResultsAvailable.Count -eq 0) #Arraylist
        }  # if-E More test
            
           
        #If no messages were found, exit
        if ($Msgs.Count -eq 0) {
            Write-Error "No messages found for the specified date range. Please check your permissions or update the date range above."
            return
        }
        # -----------
                                                                                                                            <# for comparison, here's how TonyRedmond handles the above
            [array]$Messages = $Null
            [int]$BatchSizeForMessages = 2000
                # original code [array]$MessagePage = Get-MessageTrace -StartDate $StartDate -EndDate $EndDate -PageSize 1000 -Page $i -Status "Delivered"
            Try {
                # The warning action is suppressed here because we don't want to see warnings when more data is available
                [array]$MessagePage = Get-MessageTraceV2 -StartDate $StartDate -EndDate $EndDate `
        	            -ResultSize $BatchSizeForMessages -Status "Delivered" -ErrorAction Stop -WarningAction SilentlyContinue
                $Messages += $MessagePage
            } Catch {
                Write-Host ("Error fetching message trace data: {0}" -f $_.Exception.Message)
                Break
            }
            If ($MessagePage.count -eq $BatchSizeForMessages) {
                Do {
                    Write-Host ("Fetched {0} messages so far" -f $Messages.count)
                    $LastMessageFetched = $MessagePage[-1]
                    $LastMessageFetchedDate = $LastMessageFetched.Received.ToString("O")
                    $LastMessageFetchedRecipient = $LastMessageFetched.RecipientAddress
                    # Fetch the next page of messages
                    [array]$MessagePage = Get-MessageTraceV2 -StartDate $StartDate -EndDate $LastMessageFetchedDate `
                        -StartingRecipientAddress $LastMessageFetchedRecipient -ResultSize $BatchSizeForMessages -Status "Delivered" -ErrorAction Stop -WarningAction SilentlyContinue
                    If ($MessagePage) {
                        $Messages += $MessagePage
                    }
                } While ($MessagePage.count -eq $BatchSizeForMessages)
            }
            # Remove Exchange Online public folder hierarchy synchronization messages
            $Messages = $Messages | Where-Object {$_.Subject -NotLike "*HierarchySync*"}
            #>
        # -----------
        # 12:25 PM 10/21/2025 as of Get-xoMessageTraceV2 -subject is a new [string] param, put it up in the splat - but it's not a regex, (removes my support below)
        <#
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
                #>
        
        # new try for balance of non-download work
        TRY {     
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

                    $smsg = "`n`n## Status Definitions:" ; 
                    $hReports.StatusHisto | select -expand Name | foreach-object{                   
                        $smsg += "`n$(($statusLookupTbl[$_] | ft -HideTableHeaders |out-string).trim())" ; 
                    } ; 
                    $smsg += "`n`n"
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success


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

                #region statFAIL ; #*------v statFAIL v------
                if($mFails = $msgs | ?{$_.status -eq 'Failed'} | select -last $MessageTraceDetailLimit){
                    $smsg = "Expanded analysis on last $($MessageTraceDetailLimit) Status:Failed messages..." ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 

                    if($mFails | ?{$_.Subject -notmatch '^Recall:\s' -AND $_.Subject -notmatch '^Automatic\sreply:\s'}){
                        $smsg = "Other Fails detected: Opening ExoP & ADMS connections (for get-recipient & get-aduser checks)..." ; 
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
                            #$FODetail =  $failed | Get-xoMessageTraceDetail -ea STOP; 
                            # 9:48 AM 5/2/2025 Get-xoMessageTraceDetail pipe fails, blow out into a wait loop
                            $FODetail = pull-GetxoMessageTraceDetail -Messages $failed ;                            

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
                #endregion statFAIL ; #*------^ END statFAIL ^------

                #region statQUAR ; #*------v statQUAR v------
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
                    # 8:57 AM 12/6/2024 it's taking *5mins* to Get-xoQuarantineMessage; there's no point in running that 15 times, for the same sender, w same header & senderID specs
                    # we need to down group the SenderAddress, and just process the last most-recent 'x', $QuarExpandLimitPerSender
                    $QuarSendersGrouped  = $mQuars | group SenderAddress | select Count,Name ; 
                    $smsg = "Status:Quarantined SenderAddress distribution:`n$(($QuarSendersGrouped |  ft -a count,name|out-string).trim())" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    $smsg = "EXPANDING QUARANTINES: (most recent $($QuarExpandLimitPerSender) Qurantine(s) per SenderAddress)`n$" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level PROMPT } 
                    else{ write-host -foregroundcolor YELLOW "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                    $procQuars = @() ; 
                    foreach($QSName in $QuarSendersGrouped.name){
                        $procQuars += @($mQuars | ?{$_.SenderAddress -eq $QSName} | sort Received | select -last $QuarExpandLimitPerSender) ; 
                    } ; 
                    #$ttl = $mQuars |  measure | select -expand count ;
                    $ttl = $procQuars |  measure | select -expand count ;
                    $prcd=0 ;
                    #$mQuars |foreach-object{
                    $procQuars |foreach-object{
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
                #endregion statQUAR ; #*------^ END statQUAR ^------

                #region statGETSTAT ; #*------v statGETSTAT  v------
                if( $mGetStat = $msgs|?{$_.Status -eq 'GettingStatus'}){
                    $smsg = "Status:GettingStatus returned on some traces - INDETERMINANT STATUS THOSE ITEMS (PENDING TRACKABLE LOGGING), RERUN IN A FEW MINS TO GET FUNCTIONAL DATA! (EXO-SIDE ISSUE)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 

                    $smsg = "`n`n#*------v GettingStatus's Attempt to Re-Resolve via Get-xoMessageTraceDetail (up to last $($MessageTraceDetailLimit) messages) v------`n`n" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-host -foregroundcolor yellow "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                    $midn = 20 ; $namn = 15 ; 
                    $gsmprcd = 0 ; 
                    $gsmAggr = @() ;
                    $prpgMsg = 'Received','SenderAddress','RecipientAddress','Subject' ; 
                    foreach($gsmsg in ($mGetStat | select -last $MessageTraceDetailLimit)){
                        $gsmprcd++ ; 
                        # just dump a quick summary for now
                        $smsg = "`n`n===#$($gsmprcd): MsgId: $($gsmsg.MessageId) : Status:$($gsmsg.Status)" ; 
                        $smsg += "`n$(($gsmsg | ft -a $prpgMsg|out-string).trim())" ;
                        # pop all the values but status, from the getstat event, (use the detail return'd updated)
                        $pxyEvent = [ordered]@{
                            Organization        = $gsmsg.Organization ; #$evtd.Organization ;#  toroco.onmicrosoft.com
                            MessageId           = $gsmsg.MessageId ; #$evtd.MessageId ;#  <ADR50000009071697200005056AEB0091FD089AFCAED106AF4B8@GRAINGER.COM>
                            Received            = $gsmsg.Received ; #$evtd.Date ;#  4/30/2025 4:48:27 AM
                            ReceivedLocal       = $gsmsg.ReceivedLocal ;#  4/29/2025 11:48:27 PM
                            SenderAddress       = $gsmsg.SenderAddress ; #$evtd.SenderAddress ;#  S_BTCEMAIL@GRAINGER.COM
                            RecipientAddress    = $gsmsg.RecipientAddress ; #$evtd.RecipientAddress ;#  ap@charlesmachineworks.com
                            Subject             = $gsmsg.Subject ;#  Grainger Inv # 9489372020 PO# 4501043337
                            Status              = $null ; #$evtd.Event ;#  GettingStatus
                            ToIP                = $gsmsg.ToIP ;#
                            FromIP              = $gsmsg.FromIP ;
                            Size                = $gsmsg.Size ;#  105464
                            MessageTraceId      = $gsmsg.MessageTraceId ; #$evtd.MessageTraceId ;#  f915afcc-f5ea-4f2a-3e0a-08dd87a23f8e
                            StartDate           = $gsmsg.StartDate ;#  4/29/2025 3:52:37 PM
                            EndDate             = $gsmsg.EndDate ;#  5/1/2025 3:52:37 PM
                            Index               = $gsmsg.Index ;#  9
                        } ; 

                        #if($gsmd = Get-xoMessageTrace -MessageId $gsmsg.MessageId | Get-xoMessageTraceDetail){
                        # 9:42 AM 5/2/2025 having issues with pipe into Get-xoMessageTraceDetail, expand it out, with a wait
                        # 11:18 AM 5/2/2025 shift to function
                        $gsmd = pull-GetxoMessageTraceDetail -Messages (Get-xoMessageTrace -MessageId $gsmsg.MessageId -ea STOP) ; 

                        if($gsmd){
                            # just dump a quick summary
                            $smsg += "`nDetailDisposition:`n$(($gsmd | ft -a|out-string).trim())" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Success } 
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                            $cndx = $gsmsg.Index ; 
                            foreach($evtd in $gsmd){
                                # build a proxy event to add to the msgs table
                                # will have multiple events - receive & Deliver etc - need to loop below
                                #$pxyEvent = [ordered]@{
                                    $pxyEvent.Organization        = if($evtd.Organization){$evtd.Organization} ; 
                                    $pxyEvent.MessageId           = if($evtd.MessageId){$evtd.MessageId} ;#  <ADR50000009071697200005056AEB0091FD089AFCAED106AF4B8@GRAINGER.COM>
                                    $pxyEvent.Received            = if($evtd.Date){$evtd.Date} ;#  4/30/2025 4:48:27 AM
                                    #ReceivedLocal       = if($gsmsg.ReceivedLocal){$gsmsg.ReceivedLocal} ;#  4/29/2025 11:48:27 PM
                                    $pxyEvent.SenderAddress       = if($evtd.SenderAddress){$evtd.SenderAddress} ;#  S_BTCEMAIL@GRAINGER.COM
                                    $pxyEvent.RecipientAddress    = if($evtd.RecipientAddress){$evtd.RecipientAddress} ;#  ap@charlesmachineworks.com
                                    #Subject             = if($gsmsg.Subject){$gsmsg.Subject} ;#  Grainger Inv # 9489372020 PO# 4501043337
                                    $pxyEvent.Status              = if($evtd.Event){$evtd.Event} ;#  GettingStatus
                                    $pxyEvent.ToIP                = if((([xml]$evtd.data).root.mep |?{$_.name -match 'MailboxServer|ServerHostName'}).string){
                                                                          (([xml]$evtd.data).root.mep |?{$_.name -match 'MailboxServer|ServerHostName'}).string ; 
                                                                    }
                                    # looked at resolving fqdn's at ms, to ips: there's no external dns support to ptr them
                                    $pxyEvent.FromIP              = if((([xml]$evtd.data).root.mep |?{$_.name -match 'ClientIP|ClientName'}).string ){
                                                                          (([xml]$evtd.data).root.mep |?{$_.name -match 'ClientIP|ClientName'}).string 
                                                                    } ;  
                                    #Size                = if($gsmsg.Size){} ;#  105464
                                    $pxyEvent.MessageTraceId      = if($evtd.MessageTraceId){$evtd.MessageTraceId} ;#  f915afcc-f5ea-4f2a-3e0a-08dd87a23f8e
                                    $pxyEvent.StartDate           = if($evtd.StartDate){$evtd.StartDate} ;#  4/29/2025 3:52:37 PM
                                    $pxyEvent.EndDate             = if($evtd.EndDate){$evtd.EndDate} ;#  5/1/2025 3:52:37 PM
                                    $pxyEvent.Index               = $cndx++ ;#  9
                                #} ; 
                                
                                $gsmAggr += [pscustomobject]$pxyEvent ;
                            } ; 


                        }else{
                            $smsg = "UNABLE TO Get-xoMessageTraceDetail on $($gsmsg.MessageId)!" ; 
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Warn } 
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                        } ; 

                        #$smsg = "$($sBnr3.replace('~v','~^').replace('v~','^~'))`n" ;
                        #if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level H3 } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;   
                    } ; 

                    $smsg = "`n`n#*------^  GettingStatus's Attempt to Re-Resolve via Get-xoMessageTraceDetail (up to last $($MessageTraceDetailLimit) messages)  ^------`n`n" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-host -foregroundcolor yellow "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 

                } ;
                $hReports.add('MsgsGetStatusDetail',$gsmAggr) ;
                #endregion statGETSTAT  ; #*------^ END statGETSTAT  ^------

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
                        #$mtdmsgs = $msgs | select -last $MessageTraceDetailLimit ; 
                        # should downgroup mtd senders,recipients just like quar senders
                        if($RecipientAddress -OR $SenderAddress){
                            if($RecipientAddress -AND -not $SenderAddress){
                                $smsg = "-RecipientAddress: $($RecipientAddress) with -Detail: limited SenderAddress gxmtd expansion to lastest $($QuarExpandLimitPerSender)/Sender" ; 
                                $smsg += "`n(condensing traffic...)" ; 
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                                $DtlSendersGrouped  = $msgs | group SenderAddress | select Count,Name ;
                                $mtdmsgs = @() ; 
                                foreach($QSName in $DtlSendersGrouped.name){
                                    $mtdmsgs += @($msgs | ?{$_.SenderAddress -eq $QSName} | sort Received | select -last $QuarExpandLimitPerSender) ; 
                                } ; 
                            }elseif($SenderAddress -AND -not $RecipientAddress){
                                $smsg = "-SenderAddress: $($SenderAddress) with -Detail: limited RecipientAddress gxmtd expansion to latest $($QuarExpandLimitPerSender)/RecipientAddress" ; 
                                $smsg += "`n(condensing traffic...)" ; 
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                                $DtlRecipientsGrouped  = $msgs | group recipientaddress | select Count,Name ;
                                foreach($QSName in $DtlRecipientsGrouped.name){
                                    $mtdmsgs += @($msgs | ?{$_.recipientaddress -eq $QSName} | sort Received | select -last $QuarExpandLimitPerSender) ; 
                                } ; 
                            }else{
                                # both, just do base limit                           
                                $mtdmsgs = $msgs | select -last $MessageTraceDetailLimit ; 
                            } ;
                            $smsg = "Reducing net Get-xoMessageTraceDetail lookups to last $($MessageTraceDetailLimit) messages " ; 
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                            $mtdmsgs = $mtdmsgs | select -last $MessageTraceDetailLimit ; 
                        }else{
                            # just do the last of whole set
                            $smsg = "Reducing net Get-xoMessageTraceDetail lookups to last $($MessageTraceDetailLimit) messages " ; 
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                            $mtdmsgs = $msgs | select -last $MessageTraceDetailLimit ; 
                        } ; 
                    } else { $mtdmsgs = $msgs }  ; 
                    $smsg = "`n[$(($msgs|measure).count)msgs]|=>Get-xoMessageTraceDetail:" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-host "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 

                    #$mtds = $mtdmsgs | Get-xoMessageTraceDetail ;
                    # 9:22 AM 5/2/2025 above isn't returning Get-xoMessageTraceDetail results, loop/throttle it
                    $mtds = pull-GetxoMessageTraceDetail -Messages $mtdmsgs ; 

                    $mtdRpt = @() ; 
                    if($DetailedReportRuleHits){
                        $TRules = Get-xotransportrule  ; 
                        $smsg = "Checking for `$mtds|`?{$_.Event -eq 'Transport rule'}:" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                        else{ write-host "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                    } ; 
                    $prpMTDSUM = 'DateLocal','Event','Action','Detail','TRuleName','MessageId','SenderAddress','RecipientAddress' ; 

                    foreach($mtd in $mtds){
                        $mtdsummary = [ordered]@{
                            Date = $mtd.Date ; 
                            DateLocal = ([datetime]$mtd.Date).ToLocalTime() ; 
                            Event = $mtd.Event ;
                            Action = $mtd.Action ;
                            Detail = $mtd.Detail ;
                            TRuleName = $null ; 
                            TRuleDetails = $null ; 
                            MessageId = $mtd.MessageId ; 
                            SenderAddress = if($mtd.SenderAddress){
                                                $mtd.SenderAddress ; 
                                            }else{
                                                $mtdm.SenderAddress ; 
                                            }
                            RecipientAddress =  if($mtd.RecipientAddress){
                                                $mtd.RecipientAddress
                                            }else{
                                                $mtdm.RecipientAddress ; 
                                            } 
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
                                #$smsg = "`n$(($mtdsummary| fl Date,Event,Action,Detail,TRuleName |out-string).trim())" ; 
                                # blank above
                                $smsg = "`n$(($mtdsummary | select $prpMTDSUM  | fl |out-string).trim())" ; 
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                                else{ write-host "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                            } else {
                                $smsg = "(no Event -eq 'Transport rule' matches in details run)" ; 
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                            }; 
                        } else {
                            $smsg = "`n$(($mtdsummary| fl Date,Event,Action,Detail|out-string).trim())" ; 
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                            else{ write-host "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                        }  ;
   
                        $mtdRpt += New-Object PSObject -Property $mtdsummary;
                        #[pscustomobject]$mtdsummary ; 
                    } ; 
                
                    if($mtds){
                        if($DoExports){
                            $ofileMTD = $ofile.replace('-MsgTrc','-MTD') ;
                            $smsg = "($(($mtds|measure).count) mtds | export-csv $($ofileMTD))" ; 
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
            # it's not outputting the underlying cmdlet error, try to force it :
            $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
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
                # 9:36 AM 10/21/2025 typo fix, WARNING->WARN
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                $false | write-output ; 
            } ;  
        } ; 
    } ; 
}

#*------^ Get-EXOMessageTraceExportedTDO ^------