#*------v convert-HistoricalSearchCSV.ps1 v------
function convert-HistoricalSearchCSV {
    <#
    .SYNOPSIS
    convert-HistoricalSearchCSV.ps1 - Summarize (to XML) or re-expand(to CSV), MS EXO HistoricalSearch csv output files, to permit MessageTrace-style parsing of the output for delivery patterns.
    .NOTES
    Version     : 1.0.0
    Author      : Todd Kadrie
    Website     : http://www.toddomation.com
    Twitter     : @tostka / http://twitter.com/tostka
    CreatedDate : 2021-04-23
    FileName    : convert-HistoricalSearchCSV.ps1
    License     : MIT License
    Copyright   : (c) 2021 Todd Kadrie
    Github      : https://github.com/tostka/verb-EXO
    Tags        : Powershell
    AddedCredit : REFERENCE
    AddedWebsite: URL
    AddedTwitter: URL
    REVISIONS
    * 2:54 PM 4/23/2021 wrote as freestanding .ps1, decided to flip it into func in verb-EXO
    .DESCRIPTION
    convert-HistoricalSearchCSV - Summarize (to XML) or re-expand(to CSV), MS EXO HistoricalSearch csv output files, to permit MessageTrace-style parsing of the output for delivery patterns.
    Issue is that HistoricalSearch csv files summarize a lot of detail from the normal MessageTrace .csv output, into the single Recipient_status field,
    which is a concatonated combo of every reciopient, double-hash (##) delimited with the following information per recipient
    <email address>##<status>
    And there can be a series of Status entries logged, for the single email address.

    - If ToXML is chosen, the RecipientAddress & RecipientEvents are nested as an array of CustomObjects in a field named 'RecipientStatuses'
    - If ToCsv is chosen, each transaction is unpacked back into separate 'Status' lines for each RecipientStatus (closer to the way get-MessageTrace returns records)

    The benefit of expanded CSV, over the native HS output, is you can do MessageTrace-like parsing of the output:
    $msgsx = import-csv -path path-to\file.csv ; 
    $msgsx | group status | ft -auto count,name
    Count Name
    ----- ----
      119 Receive
      117 Deliver
        2 Fail

    .PARAMETER  Files
    Array of HistoricalSearch .csv file paths[-Files c:\pathto\HistSearch.csv]
    .PARAMETER ToXML
    ToXML switch (generates nested summary XML)[-ToXML]
    .PARAMETER ToCSV
    ToCSV switch (Defaults True ; expands transactions into a logged entry per RecipientStatus)[-ToCSV]
    .PARAMETER DoDots
    Use progress dotcrawl over explicit x/y echo.
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    outputs .csv or .xml with variant [originalname]-Expanded.[ext] filename of source .csv file.
    Returns a string filepath to pipeline for each converted file.
    .EXAMPLE
    convert-HistoricalSearchCSV.ps1 -ToXML -Files "C:\usr\work\incid\123456-fname.lname@domain.com-EXOHistSrch,-60D-History,From-ANY@mssociety.org,20210222-0000AM-20210423-0919AM,run-20210423-1007AM.csv" ; 
    .EXAMPLE
    convert-HistoricalSearchCSV.ps1 -ToCSV -Files "C:\usr\work\incid\654321-MTSummary_fname.lname@domain.com-90D-History.csv" ; 
    .LINK
    https://github.com/tostka/verb-exo
    .LINK
    https://docs.microsoft.com/en-us/powershell/module/exchange/get-historicalsearch
    .LINK
    https://docs.microsoft.com/en-us/powershell/module/exchange/get-messagetrace
    #>
    #Requires -Version 3
    [CmdletBinding(DefaultParameterSetName='CSV')]
    PARAM(
        [Parameter(Position=0,Mandatory=$True,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,HelpMessage="Array of HistoricalSearch .csv file paths[-Files c:\pathto\HistSearch.csv]")]
        [ValidateNotNullOrEmpty()]$Files,
        [Parameter(ParameterSetName='XML',HelpMessage="ToXML switch (generates nested summary XML)[-ToXML]")]
        [switch] $ToXML,
        [Parameter(ParameterSetName='CSV',HelpMessage="ToCSV switch (expands transactions into a line per RecipientStatus)[-ToCSV]")]
        [switch] $ToCSV=$true,
        [Parameter(HelpMessage="Use progress dotcrawl over explicit x/y echo switch[-DoDots]")]
        [switch]$DoDots=$true 
    ) ;
    $pltXCsv = [ordered]@{
        path = $null ; 
        NoTypeInformation = $true ;
    } ; 
    foreach($file in $files){
        $error.clear() ;
        TRY {
            $ifile= gci -path $file; 
            $records = import-csv -path $ifile.fullname -Encoding Unicode ; 
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
        <# Recipient_status: The status of the delivery of the message to the recipient. 
        If the message was sent to multiple recipients, it will show all the recipients 
        and the corresponding status for each, in the format: <email address>##<status>.
        For example: 
        ##Receive, Send means the message was received by the service and was sent to the intended destination.
        ##Receive, Fail means the message was received by the service but delivery to the intended destination failed.
        ##Receive, Deliver means the message was received by the service and was delivered to the recipient�s mailbox.
        Multi recipients appear like:
        Fname.Lname@domain.com##Receive, Deliver;Fname.Lname@domain.com##Receive, Deliver;Fname.Lname@domain.com##Receive, Deliver;Fname.Lname@domain.com##Receive, Deliver;Fname.Lname@domain.com##Receive, Deliver;Fname.Lname@domain.com##Receive, Deliver;Fname.Lname@domain.com##Receive, Deliver;Fname.Lname@domain.com##Receive, Deliver
        #>
        $aggreg = @() ; 
        $procd = 0 ; $ttl = (($records|measure).count) ; $ino=0 ; 
        if($DoDots){write-host -foregroundcolor Red "[" -NoNewline } ; 
        foreach ($record in $records){
            $procd++ ; 
            # echo every 3rd record
            if(($procd % 3) -eq 0){
                if($DoDots){
                      $ino++ ; 
                      if(($ino % 80) -eq 0){
                        write-host "." ; $ino=0 ;
                      } else {write-host "." -NoNewLine} ;
                } else { 
                    write-host -foregroundcolor gray "$((get-date).ToString('HH:mm:ss')):($($procd)/$($ttl)):" ; 
                } ; 
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
                $rcpRecs = $record.recipient_status.split(';') ; # split recipients
                if($ToXML){
                    $TransSummary = [ordered]@{
                        Received=([datetime]$record.origin_timestamp_utc).ToLocalTime() ; # converting HistSearch GMT to LocalTime
                        ReceivedGMT=$record.origin_timestamp_utc ;
                        SenderAddress=$record.sender_address ;
                        RecipientAddress= $null ; 
                        Subject=$record.message_subject ;
                        Size=$record.total_bytes ;
                        MessageID=$record.message_id ;
                        OriginalClientIP=$record.original_client_ip ;
                        Directionality=$record.directionality ;
                        ConnectorID=$record.connector_id ;
                        DeliveryPriority=$record.delivery_priority ;
                    } ; 
                    $RecipientStatuses=@() ; 
                    foreach($rcpRec in $rcpRecs){
                        $statusRpt = [ordered]@{
                            RecipientAddress =  ($rcpRec -split '##')[0] ; 
                            RecipientEvents = ($rcpRec -split '##')[1] -split ', ' ; 
                        } ; 
                        $RecipientStatuses += New-Object PSObject -Property $statusRpt ; 
                    } ; 
                    $TransSummary.RecipientStatuses = $RecipientStatuses ; 
                    $aggreg += New-Object PSObject -Property $TransSummary ; 
                } elseif($ToCSV){
                    $TransSummary = [ordered]@{
                        Received=([datetime]$record.origin_timestamp_utc).ToLocalTime() ;
                        ReceivedGMT=$record.origin_timestamp_utc ;
                        SenderAddress=$record.sender_address ;
                        RecipientAddress= $null ; 
                        Subject=$record.message_subject ;
                        Size=$record.total_bytes ;
                        MessageID=$record.message_id ;
                        OriginalClientIP=$record.original_client_ip ;
                        Directionality=$record.directionality ;
                        ConnectorID=$record.connector_id ;
                        DeliveryPriority=$record.delivery_priority ;
                        Status= $null ; 
                    } ; 
                    foreach($rcpRec in $rcpRecs){
                        $TransSummary.RecipientAddress =  ($rcpRec -split '##')[0] ; 
                        foreach ($status in ($rcpRec -split '##')[1] -split ', '){
                            $TransSummary.Status = $status ;
                            $aggreg += New-Object PSObject -Property $TransSummary ; 
                        } ; 
                    } ; 
                } else { throw "neither ToCSV or ToXML specified!" } ; 
            } CATCH {
                $ErrTrapd=$Error[0] ;
                $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                Continue ; 
            } ; 
        } ; 
        if($DoDots){write-host -foregroundcolor Red "]" } ; 
        if($ToCSV){
            $pltXCsv.path = join-path -Path ($ifile.DirectoryName) -ChildPath "$($ifile.BaseName)-EXPANDED$($ifile.Extension)" ; 
            write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):export-csv w`n$(($pltXCsv |out-string).trim())" ; 
            $aggreg | export-csv @pltXCsv ;
            $pltXCsv.path | write-output ;
        } elseif ($ToXML){
            $opath = join-path -Path ($ifile.DirectoryName) -ChildPath "$($ifile.BaseName)-EXPANDED.XML" ; 
            write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):export-cliXML to`n$(($opath|out-string).trim())" ; 
            $aggreg | export-clixml -Path $opath  ;
            $opath | write-output ;
        } else { 

        } ; 
    } ;  # loop-E $files
}

#*------^ convert-HistoricalSearchCSV.ps1 ^------