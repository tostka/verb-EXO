﻿# convert-HistoricalSearchCSV.ps1

#*------v convert-HistoricalSearchCSV.ps1 v------
function convert-HistoricalSearchCSV {
    <#
    .SYNOPSIS
    convert-HistoricalSearchCSV - Summarize (to XML) or re-expand(to CSV), MS EXO HistoricalSearch (or Portal) MTSummary or MTDetail csv output files, to permit MessageTrace-style parsing of the output for delivery patterns.
    .NOTES
    Version     : 1.0.2
    Author      : Todd Kadrie
    Website     : http://www.toddomation.com
    Twitter     : @tostka / http://twitter.com/tostka
    CreatedDate : 2021-04-23
    FileName    : convert-HistoricalSearchCSV.ps1
    License     : MIT License
    Copyright   : (c) 2024 Todd Kadrie
    Github      : https://github.com/tostka/verb-EXO
    Tags        : Powershell,ExchangeOnline,MessageTrace,Conversion
    AddedCredit : REFERENCE
    AddedWebsite: URL
    AddedTwitter: URL
    REVISIONS
    * 4:08 PM 7/1/2024 spliced over code from convertfrom ConvertFrom-exoMTSummaryToEvents.ps1; wasn't sure portal and HistoricalSearch output same files, turns out, this already supports both; 
        Added wh color splats for banners; 
        added index field in MTSummary toCSV conversion (provides sub MessageID sort order, to reflect order in Recipient Status entries). 
        Added -Ticket param; Add alias:ConvertFrom-exoMTSummaryToEvents; added reverse on parsed reciient_status (put in proper recevie,resolve,delivery order); 
        pulled ParameterSets, defaulted unspec'd to -ToCSV ; added CBH exmple w Ticket
        Finally if/then'd out processing detail for dotcrawl, if DoDots is in use (default)
    * 1:58 PM 12/15/2021 revised expan code, implemented split MTDetail/MTSummary processing; normalized fieldnames against the MessageTrace output (goal is to make HS look and process more like MT)
    * 1:05 PM 12/14/2021 added full range of Expanded Rpt fields, tweaked the 
        non-recip statuses to look like recips (using primary recip & recipStat for the 
        record) ; fixed defaulted iscsv, modified param pipeline defaults; switched 
        Files from typeless to string[]; found extended gui trace had date fields with 
        diff names, added tests & support to suppress errors. ; updated Catch blocks to 
        curr spec (errors not being echoed). 
    * 11:21 AM 9/15/2021 updated Example to demo pipline-input, and post-processing to group Status (like you could a MessageTrace); added $DotsInterval param.
    * 2:54 PM 4/23/2021 wrote as freestanding .ps1, decided to flip it into func in verb-EXO
    .DESCRIPTION
    convert-HistoricalSearchCSV - Summarize (to XML) or re-expand(to CSV), MS EXO HistoricalSearch (or Portal) MTSummary or MTDetail csv output files, to permit MessageTrace-style parsing of the output for delivery patterns.

    Also works with stock Portal Mail Flow > Message Trace searches (which are the same underlying MTSummary or MTDetail format csv reports produced by Start-xoHistoricalSearch). 

    Issue is that HistoricalSearch/Portal output csv files summarize a lot of detail from the normal MessageTrace .csv output, into the single Recipient_status field,
    which is a concatonated combo of every recipient, double-hash (##) delimited with the following information per recipient
    <email address>##<status>
    And there can be a series of Status entries logged, for the single email address.

    - If ToXML is chosen, the RecipientAddress & RecipientEvents are nested as an array of CustomObjects in a field named 'RecipientStatuses'
    - If ToCsv is chosen, each transaction is unpacked back into separate 'Status' lines for each RecipientStatus (closer to the way get-MessageTrace returns records)

    The benefit of expanded CSV, over the native HS output, is you can do MessageTrace-like parsing of the results:
    $msgsx = import-csv -path path-to\MTSummary_History-expanded.csv ; 
    $msgsx | group status | ft -auto count,name
    Count Name
    ----- ----
      119 Receive
      117 Deliver
        2 Fail

    .PARAMETER  Files
    Array of HistoricalSearch/Portal download MTSummary or MTDetail .csv file paths[-Files c:\pathto\HistSearch.csv]
    .PARAMETER ToXML
    ToXML switch (generates nested summary XML)[-ToXML]
    .PARAMETER ToCSV
    ToCSV switch (Defaults True ; expands transactions into a logged entry per RecipientStatus)[-ToCSV]
    .PARAMETER Ticket
    Optional Ticket number (used for prefixing output file)[-ticket 123456]
    .PARAMETER DoDots
    Use progress dotcrawl over explicit x/y echo.
    .PARAMETER DotsInterval
    Progress dotcrawl interval (dot per every X proceessed, defaults to 3)[-DotsInterval 5]
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    outputs .csv or .xml with variant [originalname]-Expanded.[ext] filename of source .csv file.
    System.String is returned (filepath of each converted file)
    .EXAMPLE
    convert-HistoricalSearchCSV -ToXML -Files "C:\usr\work\incid\123456-fname.lname@domain.com-EXOHistSrch,-60D-History,From-ANY@mssociety.org,20210222-0000AM-20210423-0919AM,run-20210423-1007AM.csv" ; 
    Convert a HistoricalSearch .csv report, to XML (with filename:[originalname]-Expanded.xml)
    .EXAMPLE
    PS> $ifile = "C:\pathTo\MTSummary_History.csv" ;
    PS> $ofile = convert-HistoricalSearchCSV -ToCSV -Files $ifile  ; 
    PS> $msgsx = import-csv -path $ofile ; 
    PS> $msgsx | group status | ft -auto count,name
    Convert a HistoricalSearch .csv report, to -expanded.CSV, and then group the Status (as you could a normal MessageTrace). 
    .EXAMPLE
    PS> "HistReport1.csv","HistReport2.csv | convert-HistoricalSearchCSV -ToCSV ; 
    Pipeline convert multiple Hist reort csvs to xxx-expanded.csv files.
    .EXAMPLE
    PS> convert-HistoricalSearchCSV -Files 'D:\scripts\logs\849746_MTSummary_from-turfequipment@toro.com 90d - _2024-06-28T195527.734Z__67565bac-bf7a-46d7-85fe-7b1906326a27.csv' -ToCSV -ticket 849746
    Example specifying -Ticket ; 
    .LINK
    https://github.com/tostka/verb-exo
    .LINK
    https://docs.microsoft.com/en-us/powershell/module/exchange/get-historicalsearch
    .LINK
    https://docs.microsoft.com/en-us/powershell/module/exchange/get-messagetrace
    .LINK
    https://admin.exchange.microsoft.com/#/messagetrace
    #>
    #Requires -Version 3
    #[CmdletBinding(DefaultParameterSetName='CSV')]
    [CmdletBinding()]
    [Alias('convertfrom-HistoricalSearch')]
    PARAM(
        #[Parameter(Position=0,Mandatory=$True,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,HelpMessage="Array of HistoricalSearch .csv file paths[-Files c:\pathto\HistSearch.csv]")]
        [Parameter(Position=0,Mandatory=$True,ValueFromPipeline=$true,HelpMessage="Array of HistoricalSearch/Portal download MTSummary or MTDetail .csv file paths[-Files c:\pathto\HistSearch.csv]")]
            #[ValidateNotNullOrEmpty()]
            [ValidateScript({
                Test-Path $_ 
                $_ -match '\.csv$'
            })]
            #[string[]]$Files,
            [system.io.fileinfo[]]$Files,
        #[Parameter(ParameterSetName='XML',HelpMessage="ToXML switch (generates nested summary XML)[-ToXML]")]
        [Parameter(HelpMessage="ToXML switch (generates nested summary XML)[-ToXML]")]
            [switch] $ToXML,
        #[Parameter(ParameterSetName='CSV',HelpMessage="ToCSV switch (expands transactions into a line per RecipientStatus)[-ToCSV]")]
        [Parameter(HelpMessage="ToCSV switch (expands transactions into a line per RecipientStatus)[-ToCSV]")]
            [switch] $ToCSV,
        [Parameter(HelpMessage="Optional Ticket number (used for prefixing output file)[-ticket 123456]")]
        [string]$Ticket,
        [Parameter(HelpMessage="Use progress dotcrawl over explicit x/y echo switch[-DoDots]")]
            [switch]$DoDots=$true, 
        [Parameter(HelpMessage="Progress dotcrawl interval (dot per every X proceessed, defaults to 3)[-DotsInterval 5]")]
            [int]$DotsInterval=10
    ) ;
    BEGIN{
        if(-not ($ToCsv -OR $ToXML)){
            # default to csv
            write-host -foregroundcolor yellow "Neither -ToCSV or -ToXML specified: Defaulting to: -ToCSV" ; 
        } ; 
        if($ToCsv -AND $ToXML){
            write-warning "Both -ToCSV -AND -ToXML specified!`nPlease specify one or the other" ; 
        } ; 
        $verbose = ($VerbosePreference -eq "Continue") ; 
        $pltXCsv = [ordered]@{
            path = $null ; 
            NoTypeInformation = $true ;
        } ;
        $whBnr = @{BackgroundColor = 'Magenta' ; ForegroundColor = 'Black' } ;
        $whBnrS =@{BackgroundColor = 'Blue' ; ForegroundColor = 'Cyan' } ;
    } ; # BEG-E
    PROCESS{
        foreach($file in $files){
            $sBnr="#*======v PROCESSING : $($file.fullname) v======" ; 
            write-host @whBnr -obj "$((get-date).ToString('HH:mm:ss')):$($sBnr)" ;
        
            $error.clear() ;
            TRY {
                #$ifile= gci -path $file; 
                write-verbose "(import-csv:$($file.fullname))" ; 
                $records = import-csv -path $file.fullname -Encoding Unicode ; 
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
        
            <# Recipient_status: The status of the delivery of the message to the recipient. 
            If the message was sent to multiple recipients, it will show all the recipients 
            and the corresponding status for each, in the format: <email address>##<status>.
            For example: 
            ##Receive, Send means the message was received by the service and was sent to the intended destination.
            ##Receive, Fail means the message was received by the service but delivery to the intended destination failed.
            ##Receive, Deliver means the message was received by the service and was delivered to the recipient?s mailbox.
            Multi recipients appear like:
            Fname.Lname@domain.com##Receive, Deliver;Fname.Lname@domain.com##Receive, Deliver;Fname.Lname@domain.com##Receive, Deliver;Fname.Lname@domain.com##Receive, Deliver;Fname.Lname@domain.com##Receive, Deliver;Fname.Lname@domain.com##Receive, Deliver;Fname.Lname@domain.com##Receive, Deliver;Fname.Lname@domain.com##Receive, Deliver
            #>
            $aggreg = @() ; 
            $procd = 0 ; $ttl = (($records|measure).count) ; $ino=0 ; 
            if($DoDots){write-host -foregroundcolor Red "[" -NoNewline } ; 

            $isMTDetail = $false ; 
            # MTSummary has 'origin_timestamp_utc'
            # MTDetail has 'date_time_utc'
            if(($records[0] | gm | ?{$_.membertype -eq 'NoteProperty'}).name -contains 'origin_timestamp_utc'){
                $isMTDetail = $false ;
                $smsg = "(MTSummary csv file detected)" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
            }elseif(($records[0] | gm | ?{$_.membertype -eq 'NoteProperty'}).name -contains 'date_time_utc'){
                $isMTDetail = $true ;
                $smsg = "(MTDetail 'Extended' csv file detected)" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
            } else { 
                throw "Unable to determine if source is an MTSummary or MTDetail csv!"
                break ; 
            } ;  

            foreach ($record in $records){
                $procd++ ; 
                # echo every $DotsInterval'th record
                if(($procd % $DotsInterval) -eq 0){
                    if($DoDots){
                          $ino++ ; 
                          if(($ino % 80) -eq 0){
                            write-host "." ; $ino=0 ;
                          } else {write-host "." -NoNewLine} ;
                    } else { 
                        write-host -foregroundcolor gray "$((get-date).ToString('HH:mm:ss')):($($procd)/$($ttl)):" ; 
                    } ; 
                } ; 
                if(-not $DoDots){
                    #write-verbose "$((get-date).ToString('HH:mm:ss')):(record $($procd)/$($ttl)):"  ; 
                    $sBnrS="`n#*------v PROCESSING : $($procd)/$($ttl) v------" ; 
                    write-host @whBnrS -obj "$((get-date).ToString('HH:mm:ss')):$($sBnrS)" ;
                } ; 
            

                <# typical HistoricalSearch csv record & fields:
                origin_timestamp_utc : 2021-03-23T10:00:09.3284899Z
                sender_address       : Fname.Lname@domain.com
                recipient_status     : Fname.Lname@domain.com##Receive, Deliver;"Fname LName<fname.lname"@domain.com##Receive, Fail
                message_subject      : AW: Fwd: SOME SUBJECT 123456 22-03-2021
                total_bytes          : 49790
                message_id           : <PH0PR04MB73657A6BEBB3F89D9F4FC85A8C649@PH0PR04MB7365.namprd04.prod.outlook.com>
                network_message_id   : 81945af2-cab7-45ad-ba23-08d8ede2715d
                original_client_ip   : 123.456.789.012
                directionality       : Originating
                connector_id         : To_DefaultOpportunisticTLS
                delivery_priority    : Normal
                #>
                <# Extended report 8:59 AM 12/14/2021
                date_time_utc             : 2021-11-23T19:49:37.6050000Z
                client_ip                 :
                client_hostname           : CH0PR04MB8081.namprd04.prod.outlook.com
                server_ip                 :
                server_hostname           : BY5PR04MB6279.namprd04.prod.outlook.com
                source_context            : 08D9AE1AA4FEA27C;2021-11-23T19:49:37.215Z;ClientSubmitTime:2021-11-23T19:49:36.380Z
                connector_id              :
                source                    : STOREDRIVER
                event_id                  : DELIVER
                internal_message_id       : 132697
                message_id                : <CH0PR04MB8114A61AF981D65F07EA6A0C8B609@CH0PR04MB8114.namprd04.prod.outlook.com>
                network_message_id        : f955f718-d5ff-40a7-137f-08d9aeba6116
                recipient_address         : recip@domain.com
                recipient_status          :
                total_bytes               : 89464
                recipient_count           : 1
                related_recipient_address :
                reference                 :
                message_subject           : SENDER Last Day Details - List
                sender_address            : SENDER@domain.com
                return_path               : SENDER@domain.com
                message_info              : 2021-11-23T19:49:36.395Z;SRV=CH0PR04MB8114.namprd04.prod.outlook.com:TOTAL-SUB=0.218|SA=0.021|MTSS-PEN=0.197(MTSSD-PEN=0.197(MTSORGC=0.052|MTSSDC=0.073|MTSSDSDM=0.026 (MTSSDSDM-Mailbox Submission Filter
                                            Agent=0.025)|SDSSO-PEN=0.019(SMSC-PEN=0.019)));SRV=CH0PR04MB8081.namprd04.prod.outlook.com:TOTAL-HUB=0.504|SMRI=0.118(RENV=0.036|REOD=0.027|CMSGC=0.052|R-CMSG=0.026(R-CMSGC=0.023(R-HSRR=0.023
                                            )))|CAT=0.297(CATOS=0.068(CATSM=0.068(CATSM-DC Pre Content Filter Agent=0.062))|CATORES=0.187 (CATRS=0.187(CATRS-Transport Rule Agent=0.026(X-ETREX=0.022)|CATRS-DLP Policy Agent=0.043 (X-DLPEX=0.037)|CATRS-DC
                                            Content Filter Agent=0.106))|CATCC=0.024)|D-PEN=0.053(HSDSP=0.052
                                            (HSRR=0.051))|HSDN=0.031;SRV=BY5PR04MB6279.namprd04.prod.outlook.com:TOTAL-DEL=0.501|HSDR=0.113(HSDRR=0.097)|SDD=0.389(SDDPM=0.087(SDDPM-Mailbox Delivery Filter Agent=0.040|SDDPM-Inference Classification
                                            Agent=0.026)|SDDSDMG=0.268(SDDR=0.268)|X-SDDS=0.097)
                directionality            : Originating
                tenant_id                 : 549366ae-e80a-44b9-8adc-52d0c29ba08b
                original_client_ip        : 192.168.1.251
                original_server_ip        : 2603:10b6:610:f9::20
                custom_data               : S:IncludeInSla=True;S:MailboxDatabaseGuid=4ba0d02d-8b59-4bab-80e0-73f70ce64d61;S:ActivityId=77d7390c-af4d-4e43-99c5-aea5e353c61a;S:BCL=0;S:Mailboxes=f5436253-dbf4-428f-bb5c-08944e5f30e9;S:StoreObjectIds=AAAAAN
                                            4COUMvw7VMjllHB1/AorIHANlpuQRlrZxKlXO5Qqnh9vMAAAClXpkAAL34su7JVyNBoQgZmMcaJOoAAu1n1K8AAA==;S:FromEntity=Hosted;S:ToEntity=Hosted;S:P2RecipStat=0.008/9;S:MsgRecipCount=9;S:SubRecipCount=9;S:HttpRequestId=9cfd3b
                                            b0-f5cb-446d-b57e-a73440081811;S:DeliveredViaHttps=True;S:MapiMessageClass=IPM.Note;S:DeliveryLatency=1.207;S:AttachCount=1;S:E2ELatency=1.211;S:DeliveryPriority=Normal;S:PrioritizationReason=EnvelopePriority;
                                            S:AccountForest=NAMPR04A008.PROD.OUTLOOK.COM
                #>
        
                $error.clear() ;
                TRY {
                    <# fields from a typical MessageTrace (emulate the same names):
                    PSComputerName
                    RunspaceId
                    PSShowComputerName
                    Organization
                    MessageId
                    Received
                    SenderAddress
                    RecipientAddress
                    Subject
                    Status
                    ToIP
                    FromIP
                    Size
                    MessageTraceId
                    StartDate
                    EndDate
                    Index
                    #>
                
                    $TransSummary = [ordered]@{
                        Received=$null ;
                        ReceivedGMT=$null ;
                        SenderAddress=$record.sender_address ;
                        RecipientAddress= $null # $record.recipient_address ; only populated on MTDetail, imputed from recipinet_status for MTSummary
                        Status = $null ; 
                        Subject=$record.message_subject ;
                        Size=$record.total_bytes ;
                        MessageID=$record.message_id ;
                        OriginalClientIP=$record.original_client_ip ;
                        Directionality=$record.directionality ;
                        ConnectorID=$record.connector_id ;
                        DeliveryPriority=$record.delivery_priority ;
                        FromIP = $record.original_client_ip ; 
                        #ToIP = $record. ; 
                    } ; 
                
                    #if($record.origin_timestamp_utc){
                    if( -not $isMTDetail){
                        $TransSummary.Received=([datetime]$record.origin_timestamp_utc).ToLocalTime() ; # converting HistSearch GMT to LocalTime
                        $TransSummary.ReceivedGMT=$record.origin_timestamp_utc ;
                    #} elseif($record.date_time_utc){
                    } elseif($isMTDetail){
                        $TransSummary.Received=([datetime]$record.date_time_utc).ToLocalTime() ; # converting HistSearch GMT to LocalTime
                        $TransSummary.ReceivedGMT=$record.date_time_utc ;
                        write-verbose "(Expanded Report fields detected, and adding...)" ; 
                        # extended rpts include a raft of extra fields
                        #date_time_utc
                        $TransSummary.ADD('client_ip',$record.client_ip) ;
                        $TransSummary.ADD('client_hostname',$record.client_hostname) ;
                        $TransSummary.ADD('server_ip',$record.server_ip) ;
                        $TransSummary.ADD('server_hostname',$record.server_hostname) ;
                        $TransSummary.ADD('source_context',$record.source_context) ;
                        #$TransSummary.ADD('connector_id',$record.connector_id) ;
                        $TransSummary.ADD('source',$record.source) ;
                        $TransSummary.ADD('event_id',$record.event_id) ;
                        $TransSummary.ADD('internal_message_id',$record.internal_message_id) ;
                        #$TransSummary.ADD('message_id',$record.message_id) ;
                        $TransSummary.ADD('network_message_id',$record.network_message_id) ;

                        #$TransSummary.ADD('recipient_address',$record.recipient_address) ;
                        $TransSummary.RecipientAddress = $record.recipient_address ; 
                        #$TransSummary.ADD('recipient_status',$record.recipient_status) ;
                        $TransSummary.Status = $record.recipient_status ;  
                    
                        #$TransSummary.ADD('total_bytes',$record.total_bytes) ;
                        $TransSummary.ADD('recipient_count',$record.recipient_count) ;
                        $TransSummary.ADD('related_recipient_address',$record.related_recipient_address) ;
                        $TransSummary.ADD('reference',$record.reference) ;
                        #$TransSummary.ADD('message_subject',$record.message_subject) ;
                        #$TransSummary.ADD('sender_address',$record.sender_address) ;
                        #$TransSummary.SenderAddress = $record.sender_address
                        $TransSummary.ADD('return_path',$record.return_path) ;
                        $TransSummary.ADD('message_info',$record.message_info) ;
                        #$TransSummary.ADD('directionality',$record.directionality) ;
                        $TransSummary.ADD('tenant_id',$record.tenant_id) ;
                        $TransSummary.ADD('original_client_ip',$record.original_client_ip) ; # covered in base hash
                        $TransSummary.ADD('original_server_ip',$record.original_server_ip) ;
                        #$TransSummary.ADD('ToIP', $record.original_server_ip) ;
                        $TransSummary.ToIP = $record.server_ip ; 
                        $TransSummary.ADD('custom_data',$record.custom_data) ;

                    } ;

                    if($record.recipient_status.contains(";")){
                        $rcpRecs = $record.recipient_status.split(';') ; # if semi-delim'd we have multi recipients & status, split them for processing below
                    } else {
                         $rcpRecs = $record.recipient_status ; 
                    } ;  ; 
                    
                    if($ToXML){
                        if( -not $isMTDetail){
                            $RecipientStatuses=@() ; 
                            # the only one's that need expansion, are the one's delimited and with ##, all 
                            # others have a RecipientAddress & Status pulled from $record.recipient_address & 
                            # full $record.recipient_status value; 

                            #looks like non ## recipient_statu's have an entry corresponding to the number of $record.recipient_address's: [recipientAddr]:UserMailbox.Forwardable.Resolver.CreateRecipientItems.40
                            #split both and use/assign them in like order
                            if($record.recipient_status.contains(';')){
                                $rcpStatusSets = $record.recipient_status.split(';') ; 
                            } else { 
                                $rcpStatusSets = $record.recipient_status
                            } ; 
                            foreach($rcpStatusSet in $rcpStatusSets){
                                $statusRpt = [ordered]@{
                                    RecipientAddress = $null ; 
                                    Status = $null ; 
                                } ; 
                                if($rcpStatusSet.contains('##')){
                                    write-verbose "(RecipientAddress event)" ;
                                    $statusRpt.RecipientAddress =  ($rcpStatusSet -split '##')[0] ; 
                                    $statusRpt.Status = ($rcpStatusSet -split '##')[1] -split ', ' ; 
                                } else {
                                    $smsg = "MTSummary CSV that contains non-##-delimited recipient_status!"
                                    write-warning $smsg ; 
                                    throw $smsg ; 
                                    break ; 
                                    <# shouldn't have the below, all status should have ## delim ; 
                                    write-verbose "(RecipientEvent)" ;
                                    # fake the primary into the same format
                                    #$statusRpt.RecipientAddress =  $record.recipient_address ; 
                                    #$statusRpt.Status = $record.recipient_status ; 
                                    $statusRpt.RecipientAddress = $rcpRecipientSplit[$rcpRecNo] ; 
                                    $statusRpt.Status = $rcpStatusSet ; 
                                    #>
                                } ; 
                                $RecipientStatuses += New-Object PSObject -Property $statusRpt ; 
                            } ; 
                            $TransSummary.RecipientStatuses = $RecipientStatuses ; 
                        } else { 
                            # MTDetail report, has native recipient_address  & recipient_status
                            #$rcpRecipientSplit = $record.recipient_address.split(';') ; 
                            #$rcpStatusSets = $record.recipient_status.split(';') ; 
                            if($record.recipient_address.contains(';')){
                                $rcpRecipientSplit = $record.recipient_address.split(';') ; 
                            } else { 
                                $rcpRecipientSplit = $record.recipient_address ;
                            } ; 
                            if($record.recipient_status.contains(';')){
                                $rcpStatusSets = $record.recipient_status.split(';') ; 
                            } else { 
                                $rcpStatusSets = $record.recipient_status ;
                            } ; 
                            # if there's both -gt 1 recipient & -gt 1 status, do the loop, 
                            # otherwise, append the set (only reason to expand is per-recipoient status failure reporting/parsing)
                            if( ($rcpRecipientSplit|measure).count -gt 1 -AND ($rcpStatusSets|measure).count -gt 1){
                                $rcpRecNo = 0 ; 
                                foreach($rcp in $rcpRecipientSplit){
                                    $statusRpt = [ordered]@{
                                        RecipientAddress = $rcp ; 
                                        Status = $rcpStatusSets[$rcpRecNo] ; 
                                    } ; 
                                    $rcpRecNo ++ ; 
                                } ; 
                            } else { 
                                $TransSummary.RecipientAddress = $record.recipient_address ; 
                                $TransSummary.Status = $record.recipient_status ; 
                                $aggreg += New-Object PSObject -Property $TransSummary ; 
                            } ; 
                        } ; 
                        $aggreg += New-Object PSObject -Property $TransSummary ; 

                    } elseif($ToCSV){
                    
                        if( -not $isMTDetail){
                            #looks like non ## recipient_statu's have an entry corresponding to the number of $record.recipient_address's

                            if($record.recipient_status.contains(';')){
                                $rcpStatusSets = $record.recipient_status.split(';') ; 
                                [array]::Reverse($rcpStatusSets) ; # reverse the order of events, (otherwise deliver appears before receive & resolve)
                            } else { 
                                $rcpStatusSets = $record.recipient_status ;
                            } ; 
                            $index = 0 ; # we'll unpacking back into dupe message summaries with the key'd recipientaddr, one new event per logged status, assign a unique incrementing index to each
                            $TransSummary.add('index',$index) ; 
                            foreach($rcpStatusSet in $rcpStatusSets){
                                if($rcpStatusSet.contains('##')){
                                    write-verbose "(RecipientAddress event)" ;
                                    $TransSummary.RecipientAddress =  ($rcpStatusSet -split '##')[0] ; 
                                    #$statusRpt.Status = ($rcpStatusSet -split '##')[1] -split ', ' ; 
                                    foreach ($status in ($rcpStatusSet -split '##')[1] -split ', '){
                                        $index++ ;
                                        $TransSummary.index = $index ; 
                                        $TransSummary.Status = $status ;
                                        # add an entire new duped line for the status record
                                        $aggreg += New-Object PSObject -Property $TransSummary ; 
                                    } ; 
                                } else {
                                    $smsg = "MTSummary CSV that contains non-##-delimited recipient_status!"
                                    write-warning $smsg ; 
                                    throw $smsg ; 
                                    break ; 
                                } ; 
                            }

                        } else { 
                            # MTDetail report, has native recipient_address  & recipient_status
                            #$TransSummary.RecipientAddress = $rcpRecipientSplit[$rcpRecNo] ; 
                            #$TransSummary.Status = $rcpRec ; 
                            # ---
                            if($record.recipient_address.contains(';')){
                                $rcpRecipientSplit = $record.recipient_address.split(';') ; 
                            } else { 
                                $rcpRecipientSplit = $record.recipient_address ;
                            } ; 
                            if($record.recipient_status.contains(';')){
                                $rcpStatusSets = $record.recipient_status.split(';') ; 
                            } else { 
                                $rcpStatusSets = $record.recipient_status ;
                            } ; 
                            # if there's both -gt 1 recipient & -gt 1 status, do the loop, 
                            # otherwise, append the set (only reason to expand is per-recipoient status failure reporting/parsing)
                            if( ($rcpRecipientSplit|measure).count -gt 1 -AND ($rcpStatusSets|measure).count -gt 1){
                                $rcpRecNo = 0 ; 
                                foreach($rcp in $rcpRecipientSplit){
                                    #$statusRpt = [ordered]@{
                                        #RecipientAddress = $rcp ; 
                                        $TransSummary.RecipientAddress = $rcp ; 
                                        $TransSummary.Status = $rcpStatusSets[$rcpRecNo] ; 
                                    #} ; 
                                    # add a whole dupe status set for each variant 
                                    $aggreg += New-Object PSObject -Property $TransSummary ; 
                                    $rcpRecNo ++ ; 
                                } ; 
                            } else { 
                                $TransSummary.RecipientAddress = $record.recipient_address ; 
                                $TransSummary.Status = $record.recipient_status ; 
                                $aggreg += New-Object PSObject -Property $TransSummary ; 
                            } ; 
                            # ---
                            #$aggreg += New-Object PSObject -Property $TransSummary ; 
                        } ; 
                    } else { throw "neither ToCSV or ToXML specified!" } ; 
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
                
                if(-not $DoDots){
                    write-host @whBnrS -obj "$((get-date).ToString('HH:mm:ss')):$($sBnrS.replace('-v','-^').replace('v-','^-'))" ;
                } ; 
            } ; 
            if($DoDots){write-host -foregroundcolor Red "]" } ; 
            TRY{
                if($ToCSV){
                    if($ticket){
                        if(-not ($file.BaseName -match "^$($ticket)")){
                            write-host "-Ticket $($ticket) specified, and input path wasn't prefixed with Ticket#: Prefixing output filename with Ticket specified" ;
                            $pltXCsv.path = join-path -Path ($file.DirectoryName) -ChildPath "$($Ticket)-$($file.BaseName)-EXPANDED$($file.Extension)" ; 
                        } else { 
                            $pltXCsv.path = join-path -Path ($file.DirectoryName) -ChildPath "$($file.BaseName)-EXPANDED$($file.Extension)" ; 
                        } ; 
                    } else {
                        $pltXCsv.path = join-path -Path ($file.DirectoryName) -ChildPath "$($file.BaseName)-EXPANDED$($file.Extension)" ; 
                    } ; 
                    write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):export-csv w`n$(($pltXCsv |out-string).trim())" ; 
                    $aggreg | export-csv @pltXCsv ;
                    write-verbose "Returning exported csv path to pipeline" ;
                    $pltXCsv.path | write-output ;
                } elseif ($ToXML){
                    if($ticket){
                        if(-not ($file.BaseName -match "^$($ticket)")){
                            write-host "-Ticket $($ticket) specified, and input path wasn't prefixed with Ticket#: Prefixing output filename with Ticket specified" ;
                            $pltXCsv.path = join-path -Path ($file.DirectoryName) -ChildPath "$($Ticket)-$($file.BaseName)-EXPANDED.xml" ; 
                        } else { 
                            $pltXCsv.path = join-path -Path ($file.DirectoryName) -ChildPath "$($file.BaseName)-EXPANDED.xml" ; 
                        } ; 
                    } ; 
                    write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):export-cliXML to`n$(($opath|out-string).trim())" ; 
                    $aggreg | export-clixml -Path $opath  ;
                    write-verbose "Returning exported xml path to pipeline" ;
                    $opath | write-output ;
                } else { 

                } ; 
            } CATCH {
                $ErrTrapd=$Error[0] ;
                $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" ;
                Continue
            } ;

            write-host @whBnr -obj "$((get-date).ToString('HH:mm:ss')):$($sBnr.replace('=v','=^').replace('v=','^='))" ;
        } ;  # loop-E $files
    } ; # PROC-E 
} ;  

#*------^ convert-HistoricalSearchCSV.ps1 ^------
