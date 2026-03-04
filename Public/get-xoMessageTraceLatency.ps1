 # get-xoMessageTraceLatency_func.ps1



#region GET_XOMESSAGETRACELATENCYTDO ; #*------v get-xoMessageTraceLatencyTDO v------
Function get-xoMessageTraceLatencyTDO {
        <#
        .SYNOPSIS
        get-xoMessageTraceLatencyTDO - Returns the 'calculated' ExchangeOnline MessageLatency (approximateed from events & Get-xoMessageTraceDetailV2; as MS hides Total Latency on tracks) on a specified MessageID message
        .NOTES
        Version     : 0.0.
        Author      : Todd Kadrie
        Website     : http://www.toddomation.com
        Twitter     : @tostka / http://twitter.com/tostka
        CreatedDate : 2026-02-27
        FileName    : get-xoMessageTraceLatencyTDO.ps1
        License     : MIT License
        Copyright   : (c) 2026 Todd Kadrie
        Github      : https://github.org/tostka/verb-exo/
        Tags        : Powershell,Exchange,ExchangeOnline,MessageTrace,Tracking,Messasge
        AddedCredit : 
        AddedWebsite: 
        AddedTwitter: 
        REVISIONS
        * 4:03 PM 3/3/2026 updated, functional, outputs formated latency as well as raw timespan, and has optional -details
        * 4:27 PM 2/27/2026 init
        .DESCRIPTION
        get-xoMessageTraceLatencyTDO - Returns the 'calculated' MessageLatency (approximateed from events & Get-xoMessageTraceDetailV2; as MS hides Total Latency on tracks) on a specified MessageID message

        .PARAMETER MessageTraceData
        MessageTraceData Array (as returned by Get-MessageTraceV2) [-MessageTraceData `$Messages]
        .PARAMETER MessageId
        The MessageId parameter filters the results by the Message-ID header field of the message. Be sure to include the full Message ID string (which might include angle brackets) and enclose the value in quotation marks (for example, `"d9683b4c-127b-413a-ae2e-fa7dfb32c69d@DM3NAM06BG401.Eop-nam06.prod.protection.outlook.com`").[-File '<fb08fbe1a99c54827480cc6451de0d66@wince>']
        .INPUTS
        Accepts pipeline input of Get-MessageTraceV2 results (as -MessageTraceData)
        .OUTPUTS        
        System.Management.Automation.PSCustomObject Summary of Array of MessageLatency specifications.
        .EXAMPLE
        PS> $results = get-xoMessageTraceLatencyTDO -MessageId '<fb08fbe1a99c54827480cc6451de0d66@wince>' -MessageTraceData $xoMsgs_28538.MTMessages ; 
        PS> $smsg = "$(($results| fl MessageId,TotalLatencyFmt|out-string).trim())" ;
        PS> $smsg += "`n`n$(($results| select -expand DetailEvents|out-string).trim())"
        PS> write-host $smsg ;

            MessageId       : <fb08fbe1a99c54827480cc6451de0d66@wince>
            TotalLatencyFmt : 0d:00h:00m:08s.000

            Event   Date                  LatencyFromPriorEvt_hms
            -----   ----                  -----------------------
            Receive 2/23/2026 11:01:02 PM                        
            Resolve 2/23/2026 11:01:02 PM 00:00:00.3280000       
            Deliver 2/23/2026 11:01:09 PM 00:00:06.8650000            

        Cmdline parameter demo
        .EXAMPLE
        PS> $results = get-xoMessageTraceLatencyTDO -MessageId '<fb08fbe1a99c54827480cc6451de0d66@wince>' -MessageTraceData $xoMsgs_28538.MTMessages -Details ; 
        PS> $smsg = "$(($results| fl MessageId,TotalLatencyFmt,*address,Subject|out-string).trim())" ; 
        PS> $smsg += "`n`n$(($results| select -expand DetailEvents|out-string).trim())"        
        PS> write-host $smsg ;

            MessageId        : <fb08fbe1a99c54827480cc6451de0d66@wince>
            TotalLatencyFmt  : 0d:00h:00m:08s.000
            RecipientAddress : {Local.1291@aaaaaa.mail.onmicrosoft.com, Local.1291@aaaa.com}
            SenderAddress    : aaa-aa-aaaaaaaa@aaaa.aaa
            Subject          : Aaaaa Aaaaaaa

            Event   Date                  LatencyFromPriorEvt_hms
            -----   ----                  -----------------------
            Receive 2/23/2026 11:01:02 PM                        
            Resolve 2/23/2026 11:01:02 PM 00:00:00.3280000       
            Deliver 2/23/2026 11:01:09 PM 00:00:06.8650000        

        Cmdline parameter demo
        .EXAMPLE
        PS> $results = $xoMsgs_28538.MTMessages | get-xoMessageTraceLatencyTDO -MessageID '<fb08fbe1a99c54827480cc6451de0d66@wince>' ; 
        Pipeline demo, feeding captured Get-MessageTraceV2 output in, and specifying a specific messageID to be targeted (not working ATM)
        .EXAMPLE
        PS> $results = get-xoMessageTraceLatencyTDO -MessageTraceData $xoMsgs_28538.MTMessages ; 
        PS> $results2 | ft -a ReceivedGMT,RecipientAddress,TotalLatencyFmt

            ReceivedGMT         RecipientAddress                                              TotalLatencyFmt   
            -----------         ----------------                                              ---------------   
            02/23/2026 23:01:01 {Local.1291@aaaaaa.mail.onmicrosoft.com, Local.1291@aaaa.com} 0d:00h:00m:08s.000
            ...

        Cmdline parameter demo, wo MessageID specified, processes only the first 20 items. With summary tablular output for series
        .LINK
        https://github.org/tostka/verb-exo/
        #>
        [CmdletBinding()]
        [alias('get-xoMessageTraceLatency')]
        PARAM(
            [Parameter(Position=0,Mandatory=$True,ValueFromPipeline=$true,HelpMessage="MessageTraceData Array (as returned by Get-MessageTraceV2) [-MessageTraceData `$Messages]")]
                $MessageTraceData, # stored, it's System.Collections.ArrayList, single isolated item is System.Management.Automation.PSCustomObject
            [Parameter(Mandatory=$false,HelpMessage = "The MessageId parameter filters the results by the Message-ID header field of the message. Be sure to include the full Message ID string (which might include angle brackets) and enclose the value in quotation marks (for example, `"d9683b4c-127b-413a-ae2e-fa7dfb32c69d@DM3NAM06BG401.Eop-nam06.prod.protection.outlook.com`").[-File '<fb08fbe1a99c54827480cc6451de0d66@wince>']")]
                #[MultiValuedProperty]
                [string[]]$MessageId,
            [Parameter(Mandatory=$false,HelpMessage = "Switch to output MessageDetails.[-Details]")]
                [switch]$Details,
            [Parameter(Mandatory=$false,HelpMessage = "Limit for number of Get-MessageTraceDetailV2 passes to run (Defaults 20, to avoid MS throttling).[-DetailLimit 10]")]
                [int]$DetailLimit = 20
        ) ;
        BEGIN{
            #region FUNCTIONS_INTERNAL ; #*======v FUNCTIONS_INTERNAL v======

            #region GET_TRACEDETAILTOTALLATENCY ; #*------v Get-TraceDetailTotalLatency v------
            function Get-TraceDetailTotalLatency {
                <#                
                .SYNOPSIS
                Get-TraceDetailTotalLatency.ps1 - Resolves Get-MessageTraceDetailV2 events into a summary Latency series
                .NOTES
                Version     : 0.0.
                Author      : Todd Kadrie
                Website     : http://www.toddomation.com
                Twitter     : @tostka / http://twitter.com/tostka
                CreatedDate : 2025-
                FileName    : Get-TraceDetailTotalLatency.ps1
                License     : MIT License
                Copyright   : (c) 2026 Todd Kadrie
                Github      : https://github.com/tostka/verb-XXX
                Tags        : Powershell
                AddedCredit : REFERENCE
                AddedWebsite: URL
                AddedTwitter: URL
                REVISIONS
                .DESCRIPTION
                Get-TraceDetailTotalLatency.ps1 - Resolves Get-MessageTraceDetailV2 events into a summary Latency series
                .PARAMETER  TraceDetailRow
                get-messagetracedetailv2 outputfor processing                
                .INPUTS
                Acceptes Pipeline input (get-messagetracedetailv2 output for processing)
                .OUTPUTS
                PsCustomObject - Returns Event Latency summary                
                .EXAMPLE
                PS> Get-MessageTraceDetailV2 -MessageTraceId $MessageTraceId -RecipientAddress $Recipient | Where-Object Event -eq 'Deliver' |Get-TraceDetailTotalLatency ; 
                Usage: filter for Deliver and emit both seconds and TimeSpan
                .LINK
                https://github.com/tostka/verb-exo                
                #>
                [CmdletBinding()]
                PARAM(
                    [Parameter(Mandatory,ValueFromPipeline,helpmessage="get-messagetracedetailv2 outputfor processing")]
                        $TraceDetailRow
                )
                PROCESS {
                    if (-not $TraceDetailRow.Data) { return }

                    $xml = [xml]$TraceDetailRow.Data
                    $node = $xml.SelectSingleNode("//MEP[@Name='TotalLatency']")
                    if (-not $node) { return }

                    # Prefer Integer; fall back to Long or Double if ever present
                    $val = $null
                    foreach ($attrName in 'Integer','Long','Double') {
                        if ($node.Attributes[$attrName]) {
                            $val = $node.Attributes[$attrName].Value
                            break
                        }
                    }
                    if ($val -ne $null) {
                        [pscustomobject]@{
                            Date                 = $TraceDetailRow.Date
                            Event                = $TraceDetailRow.Event
                            #TotalLatencySeconds  = [double]$val
                            TotalLatency         = [TimeSpan]::FromSeconds([double]$val)
                        } | write-output 
                    }
                }
            }
            #endregion GET_TRACEDETAILTOTALLATENCY ; #*------^ END Get-TraceDetailTotalLatency ^------

            #endregion FUNCTIONS_INTERNAL ; #*======^ END FUNCTION FUNCTIONS_INTERNAL ^======

            #$global:prev = $null ;
            $ThrottleMs = 500 ; 
            $rgxEventTerminal = 'Deliver|Send\sexternal'  # ending event from which to calc EstLatency on Get-xoMessageTraceDetailV2 reseults
            $prev = $null ;            
            if($MessageId){
                $smsg = "-MessageID: using specified target values: $($MessageId -join ',')" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
            }else{
                #$MessageId = $MessageTraceData | select -unique MessageID
                $MessageId = ($MessageTraceData | select -unique MessageID).MessageID ; 
                $smsg = "No -MessageID: processing *ALL*: $($MessageId  |measure | select -expand count ) unique values" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
            }
            if($MessageId.count -gt $DetailLimit){
                $smsg = "More than $($DetailLimit) MessageIDs specified (or within MesssageTraceData where no MessageIDs specified)" ; 
                $smsg += "`nGet-MessageTraceDetailV2 will only be run on the LAST $($DetailLimit) MessageIDs! " ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                $MessageId = $MessageId | select -LAST $DetailLimit ; 
            } 
            $ReportArray = @()
            $Proc = 0 ; 
            $ttl = $MessageId|  measure | select -expand count 
        }
        PROCESS{
            foreach($mid in $MessageId){
                $hsReport = [ordered]@{
                    MessageId = $mid ;
                    ReceivedGMT = $null ;              
                } ; 
                if($Details){
                    $hsReport.add('RecipientAddress',$null)   
                    $hsReport.add('SenderAddress',$null)
                    $hsReport.add('Subject',$null)
                } ; 
                $hsReport.add('DetailEvents',@()) ; 
                $hsReport.add('TotalLatency',$null) ; 
                $hsReport.add('TotalLatencyFmt',$null) ; 
                $report = [pscustomobject]$hsReport
                $Proc++ ; 
                $smsg = $sBnrS="`n#*------v PROCESSING ($($Proc)/$($ttl)): $($mid) v------" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level H2 } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                TRY{
                    $thisTrace = @() ; 
                    #if($thisTrace = $MessageTraceData |  ?{$_.messageid -eq $mid } | sort Received){
                    # when all evts have same received, have to force the delivered out
                    if ( ($MessageTraceData |  ?{$_.messageid -eq $mid} | select -unique Received).received.count -eq 1){
                        $thisTrace += @($MessageTraceData |  ?{$_.messageid -eq $mid -ANd $_.status -ne 'Delivered'}  | SORT received) ; 
                        if($MessageTraceData |  ?{$_.messageid -eq $mid -ANd $_.status -eq 'Delivered'}){
                            $thisTrace += @($MessageTraceData |  ?{$_.messageid -eq $mid -ANd $_.status -eq 'Delivered'} | SORT received) ; 
                        } ; 
                    }else{
                        $thisTrace = $MessageTraceData |  ?{$_.messageid -eq $mid } | sort Received ; 
                    } ; 
                    if($thisTrace ){
                        $report.ReceivedGMT = (get-date ($thistrace.Received | select -unique | sort | select -first 1 ) -format 'MM/dd/yyyy HH:mm:ss');
                        if($Details){
                            $Report.RecipientAddress = ($thistrace.RecipientAddress | select -unique) ; 
                            $Report.SenderAddress = ($thistrace.senderaddress | select -unique) ; 
                            $Report.Subject = ($thistrace.Subject | select -unique) ; 
                        } ; 
                        if($thisEvt = $thistrace | ?{$_.Status -eq 'Receive'}){
                            write-host "(running gxmtd on Status:$($thisEvt.Status) event of MessageId:($mid))" ;
                            $thisDtl =$thistrace | ?{$_.Status -eq 'Receive'}| Get-xoMessageTraceDetailV2 -ea STOP
                        }elseif($thisEvt = $thistrace | select -first 1 ){
                            write-host "(running gxmtd on earliest Status:$($thisEvt.Status) event of MessageId:($mid))" ;
                        } ;
                        if($thisEvt){
                            if($thisDtl = $thisEvt | Get-xoMessageTraceDetailV2 -ea STOP){
                                $script:prev = $null ; 
                                #$thisDtl |
                                $report.DetailEvents = $thisDtl |
                                    Sort-Object Date |
                                        Select-Object Event, Date,
                                        @{n='LatencyFromPriorEvt_hms';e={
                                           #if ($prev) { $_.Date - $prev.Date } else { $null }
                                           #if ($global:prev) { $_.Date - $global:prev.Date } else { $null }
                                           if ($script:prev) { $_.Date - $script:prev.Date } else { $null }
                                           #$global:prev = $_ ;
                                           $script:prev = $_ ;
                                           #$prev = $_ ;
                                        }} ;
                                $DeliverEvent = $TotalLatency = $null ; 
                                #$rgxEventTerminal = 'Deliver|Send\sexternal'
                                #if($DeliverEvent = $thisDtl | ?{$_.Event -eq 'Deliver'}){
                                # add leaf sendexternal (might work)
                                if($DeliverEvent = $thisDtl | ?{$_.Event -match $rgxEventTerminal }){
                                    #$TotalLatency = $DeliverEvent  |Get-TraceDetailTotalLatency ; 
                                    if($EstLatency = ($DeliverEvent  |Get-TraceDetailTotalLatency)){
                                        #$report.TotalLatency = ($DeliverEvent  |Get-TraceDetailTotalLatency).TotalLatency ; 
                                        # timespan, expand it
                                        $ts = $EstLatency.TotalLatency| select days,hours,minutes,seconds,milliseconds ;
                                        [array]$fst=@() ;
                                        #$EstLatency.TotalLatency.psobject.properties |?{$_.value} |foreach-object{
                                        <#
                                        $EstLatency.TotalLatency.psobject.properties |
                                            ?{$_.value -AND $_.name  -match 'Days|Hours|Minutes|Seconds|Milliseconds' -AND $_.name -notmatch '^Total'} |foreach-object{
                                                $thisMeas = $_ ; 
                                                switch -regex($thisMeas.name){
                                                    'Days'{$fst += "$($thisMeas.value)d"} 
                                                    'Hours'{$fst += "$($thisMeas.value)h"}
                                                    'Minutes'{$fst += "$($thisMeas.value)m"} 
                                                    'Seconds'{$fst += "$($thisMeas.value)s"} 
                                                    'Milliseconds'{$fst += "$($thisMeas.value)ms"} 
                                                    default{} 
                                                } ; 
                                        } ;
                                        #>
                                        $ts.psobject.properties |?{$_.value} |foreach-object{switch ($_.name){ 'Days'{$fst += "{0:dd}d"} 'Hours'{$fst += "{0:hh}h"} 'Minutes'{$fst += "{0:mm}m"} 'Seconds'{$fst += "{0:ss}s"} 'Milliseconds'{$fst += "{0:fff}ms"} default{} } ; } ;
                                        #$smsg =  ("(Elapsed: $($fst -join " "))" -f $EstLatency.TotalLatency) ;
                                        #$smsg =  ("($($fst -join ":"))" -f $EstLatency.TotalLatency) ;
                                        $report.TotalLatency = $EstLatency.TotalLatency ; 
                                        $report.TotalLatencyFmt = $EstLatency.TotalLatency.ToString("d\d\:hh\h\:mm\m\:ss\s\.fff") ; 
                                        #$ssmsg = $EstLatency.TotalLatency.ToString("d\d\:hh\h\:mm\m\:ss\s\.fff")
                                        $smsg = "$($mid): TotalLatancy:$($report.TotalLatencyFmt)" ; 
                                        write-host $smsg ;                                        
                                    } else { 
                                        $smsg = "Unable to calculate TotalLatency on`n$(($DeliverEvent |out-string).trim())" ; 
                                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                                    } ; 
                                }else{
                                    $smsg = "No Event:Deliver from which to perform a TotalLatency calculation" ; 
                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                                }
                                start-sleep -Milliseconds $ThrottleMs ; 
                            }else{
                                throw "unable to return Get-xoMessageTraceDetailV2 on MessageId:$($mid)" ;
                            } ;
                        }else{
                            throw "unable to isolate earliest event in stream on MessageId:$($mid)" ;
                        } ;                        
                        $ReportArray+= $report ; 
                        $smsg = "$(($report | ft -a MessageID,TotalLatency|out-string).trim())" ; 
                        if($Details){                             
                            $smsg += "`n`nSenderAddress:$($Report.SenderAddress)" ; 
                            $smsg += "`nRecipientAddress:$($Report.RecipientAddress -join ',')" ; 
                            $smsg += "`nSubject:$($Report.Subject)" ; 
                        } ; 
                        $smsg += "`n`nEventDetails:`n$(($report.DetailEvents|out-string).trim())`n`n" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                        else{ write-host -foregroundcolor green $smsg } ;
                        #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                    }else{
                        throw "unable to isolate MessageId:$($mid) from input gxmt stream" ;
                    }
                } CATCH {$ErrTrapd=$Error[0] ;
                    write-host -foregroundcolor gray "TargetCatch:} CATCH [$($ErrTrapd.Exception.GetType().FullName)] {"  ;
                    $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                    write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" ;
                    CONTINUE ; 
                } ;
                $smsg = "$($sBnrS.replace('-v','-^').replace('v-','^-'))" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level H2 } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            }
        } # PROC-E
        END{
            if($ReportArray.count -gt 0){
                $ReportArray | write-output 
            }
        }
        
    } ;  
    #endregion GET_XOMESSAGETRACELATENCYTDO ; #*------^ END get-xoMessageTraceLatencyTDO ^------

