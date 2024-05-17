# get-ExoMessageTraceEventType.ps1

#*------v get-ExoMessageTraceEventType.ps1 v------
function get-ExoMessageTraceEventType {
<#
    .SYNOPSIS
    get-ExoMessageTraceEventType - Provides a prefab array indexed hash of Exchange-Online Get-xoMessageTrace Event Types Note: This is a static non-query-based list of events. The function must be manually updated to accomodate MS MessageTrace event types changes over time.
    .PARAMETER Mailboxes
    .NOTES
    Version     : 1.0.0
    Author      : Todd Kadrie
    Website     :	http://www.toddomation.com
    Twitter     :	@tostka / http://twitter.com/tostka
    CreatedDate : 2024-05-17
    FileName    : get-ExoMessageTraceEventType.ps1
    License     : MIT License
    Copyright   : (c) 2022 Todd Kadrie
    Github      : https://github.com/tostka/verb-ex2010
    Tags        : Powershell,ExchangeOnline,MessageTrace,Reporting
    REVISIONS
    * 3:55 PM 5/17/2024 adapted from vxo\get-ExoMailboxLicenses(); ren get-ExoGxMTEventType -> get-ExoMessageTraceEventType; added array support on the EventID & Type
    .DESCRIPTION
    get-ExoMessageTraceEventType - Provides a prefab array indexed hash of Exchange-Online Get-xoMessageTrace Event Types Note: This is a static non-query-based list of events. The function must be manually updated to accomodate MS MessageTrace event types changes over time.
    Run it without an -EventID and it returns the entire indexed hash to the pipeline

    I have assigned arbitrary 'Type's to each, to make it easy to filter MessageTrace output events into those relevent for demonstrating delivery (transport), those that indicate issues (problem), and those that provide additional information (info).

    .PARAMETER EventID
    Optional MessageTrace event-id(s) to be resolved to details[-EventID 'AGENTINFO']
    .PARAMETER Type
    Optional type classification subset(s) to be returned from the entire EventID set[-Type 'Problem','Transport']
    .EXAMPLE
    PS> $eventInfo = get-ExoMessageTraceEventType -eventid AGENTINFO -verbose ; 
    PS> $eventInfo | out-string ; 

        EventName Description                                                Type
        --------- -----------                                                ----
        AGENTINFO This event is used by transport agents to log custom data. info

    Retrieve Resolve 'AGENTINFO' to it's stock MS description & type, and output to console
    .EXAMPLE
    PS> $eventIdsDetail = get-ExoMessageTraceEventType ; 
    PS> $eventIdsDetail | out-string ; 

        Name                           Value                                                                                                                                                                                                                             
        ----                           -----                                                                                                                                                                                                                             
        AGENTINFO                      @{EventName=AGENTINFO; Description=This event is used by transport agents to log custom data.; Type=info}                   
        ...[TRIMMED]...
        TRANSFER                       @{EventName=TRANSFER; Description=Recipients were moved to a forked message because of content conversion, message recipient limits, or agents. Sources include ROUTING or QUEUE.; Type=info}                                     

    Retrive the entire indexed hash of EventID details, and output the results to console. 
    PS> $eventIdsDetail['AGENTINFO'] | out-string ; 

        EventName Description                                                Type
        --------- -----------                                                ----
        AGENTINFO This event is used by transport agents to log custom data. info

    PS> $eventIdsDetail.values | group type |  ft -a count,name 

        Count Name     
        ----- ----     
           18 info     
           18 problem  
            5 transport

    PS> $rgxEvtIDProblem = ('(' + (($eventIdsDetail.values | ?{$_.type -eq 'problem'} | select -expand EventName |%{[regex]::escape($_)}) -join '|') + ')') ;
    PS> $rgxEvtIDTransp = ('(' + (($eventIdsDetail.values | ?{$_.type -eq 'transport'} | select -expand EventName|%{[regex]::escape($_)}) -join '|') + ')') ;
    PS> $rgxEvetIDInfo =  ('(' + (($eventIdsDetail.values | ?{$_.type -eq 'info'} | select -expand EventName|%{[regex]::escape($_)}) -join '|') + ')') ;
    PS> $rgxEvtIDProblem

        (BADMAIL|DEFER|DELIVERFAIL|DROP|DSN|FAIL|HAREDIRECTFAIL|INITMESSAGECREATED|MODERATIONEXPIRE|MODERATORREJECT|MODERATORSALLNDR|POISONMESSAGE|RESUBMITDEFER|RESUBMITFAIL|SUBMITDEFER|SUBMITFAIL|SUPPRESSED|THROTTLE)

    PS> $rgxEvetIDInfo

        (AGENTINFO|CLIENTSUBMISSION|DUPLICATEDELIVER|DUPLICATEEXPAND|DUPLICATEREDIRECT|EXPAND|HADISCARD|HARECEIVE|HAREDIRECT|LOAD|NOTIFYMAPI|NOTIFYSHADOW|PROCESS|PROCESSMEETINGMESSAGE|RESOLVE|RESUBMIT|SUBMIT|TRANSFER)

    PS> $rgxEvtIDTransp 

        (DELIVER|MODERATORAPPROVE|RECEIVE|REDIRECT|SEND)

    Retrieve the entire set of defined event-id's returned as an indexed hash. Output the hash to console; lookup and return the AGENTINFO event-id details; group the assigned types; Build & output regexes for the Problem, Transport & Info types.
    .EXAMPLE
    PS> $eventInfo = get-ExoMessageTraceEventType -type problem ; 
    PS> $eventInfo | write-output ;

        EventName          Description                                                                                                                                                                                                           
        ---------          -----------                                                                                                                                                                                                           
        BADMAIL            A message submitted by the Pickup directory or the Replay directory that can't be delivered or returned.                                                                                                              
        DEFER              Message delivery was delayed.                                                                                                                                                                                         
        DELIVERFAIL        An agent tried to deliver the message to a folder that doesn't exist in the mailbox.                                                                                                                                  
        DROP               A message was dropped without a delivery status notification (also known as a DSN, bounce message, non-delivery report, or NDR). For example:\n* Completed moderation approval request messages.\n* Spam messages t...
        DSN                A delivery status notification (DSN) was generated.                                                                                                                                                                   
        FAIL               Message delivery failed. Sources include SMTP, DNS, QUEUE, and ROUTING.                                                                                                                                               
        HAREDIRECTFAIL     A shadow message failed to be created. The details are stored in the source-context field.                                                                                                                            
        INITMESSAGECREATED A message was sent to a moderated recipient, so the message was sent to the arbitration mailbox for approval. For more information, see Manage message approval.                                                      
        MODERATIONEXPIRE   A moderator for a moderated recipient never approved or rejected the message, so the message expired. For more information about moderated recipients, see Manage message approval.                                   
        MODERATORREJECT    A moderator for a moderated recipient rejected the message, so the message wasn't delivered to the moderated recipient.                                                                                               
        MODERATORSALLNDR   All approval requests sent to all moderators of a moderated recipient were undeliverable, and resulted in non-delivery reports (also known as NDRs or bounce messages).                                               
        POISONMESSAGE      A message was put in the poison message queue or removed from the poison message queue.                                                                                                                               
        RESUBMITDEFER      A message resubmitted from Safety Net was deferred.                                                                                                                                                                   
        RESUBMITFAIL       A message resubmitted from Safety Net failed.                                                                                                                                                                         
        SUBMITDEFER        The message transmission from the Mailbox Transport Submission service to the Transport service was deferred.                                                                                                         
        SUBMITFAIL         The message transmission from the Mailbox Transport Submission service to the Transport service failed.                                                                                                               
        SUPPRESSED         The message transmission was suppressed.                                                                                                                                                                              
        THROTTLE           The message was throttled.  

    Retrieve EventID details for all -type Problem events, and output results to console. 
    .EXAMPLE
    PS> $eventInfo = get-ExoMessageTraceEventType -type 'problem','transport' ; 
    PS> $eventInfo | write-output ;
    PS> $rgxEvtIDTransProb = ('(' + (($eventInfo.EventName |%{[regex]::escape($_)}) -join '|') + ')') ;
    Retrieve the Type:Problem & Transport EventIDs and build a regex out of the combonation
    .LINK
    https://github.com/tostka/verb-EXO
    #>
    #Requires -Modules verb-IO, verb-logging, verb-Text
    [OutputType('System.Collections.Hashtable')]
    [CmdletBinding()]
    PARAM(
        [Parameter(Mandatory=$FALSE,HelpMessage="Optional MessageTrace event-id(s) to be resolved to details[-EventID 'AGENTINFO']")]
            [ValidateNotNullOrEmpty()]
            [string[]]$EventID,
        [Parameter(Mandatory=$FALSE,HelpMessage="Optional type classification subset(s) to be returned from the entire EventID set[-Type 'Problem','Transport']")]
            [ValidateSet('transport','problem','info')]
            [string[]]$Type
    ) ;
    if($EventID -AND $Type){
        $smsg = "Both -EventID & -Type specified: Please specify one or the other (or none, to return all defined EventIDs)" ; 
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
    } ; 
    ${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name ;
    $verbose = ($VerbosePreference -eq "Continue") ;
    
    # check if using Pipeline input or explicit params:
    if ($PSCmdlet.MyInvocation.ExpectingInput) {
        write-verbose "Data received from pipeline input: '$($InputObject)'" ;
    } else {
        # doesn't actually return an obj in the echo
        #write-verbose "Data received from parameter input: '$($InputObject)'" ;
    } ;
    

    # input table of Exchange Online assignable licenses that include a UserMailbox:
    $ExoEventIDsTbl = @"
|EventName |Description|Type|
|---|---|---|
|AGENTINFO |This event is used by transport agents to log custom data.|info|
|BADMAIL |A message submitted by the Pickup directory or the Replay directory that can't be delivered or returned.|problem|
|CLIENTSUBMISSION |A message was submitted from the Outbox of a mailbox.|info|
|DEFER |Message delivery was delayed.|problem|
|DELIVER |A message was delivered to a local mailbox.|transport|
|DELIVERFAIL |An agent tried to deliver the message to a folder that doesn't exist in the mailbox.|problem|
|DROP |A message was dropped without a delivery status notification (also known as a DSN, bounce message, non-delivery report, or NDR). For example:<br/>* Completed moderation approval request messages.<br/>* Spam messages that were silently dropped without an NDR.|problem|
|DSN |A delivery status notification (DSN) was generated.|problem|
|DUPLICATEDELIVER |A duplicate message was delivered to the recipient. Duplication may occur if a recipient is a member of multiple nested distribution groups. Duplicate messages are detected and removed by the information store.|info|
|DUPLICATEEXPAND |During the expansion of the distribution group, a duplicate recipient was detected.|info|
|DUPLICATEREDIRECT |An alternate recipient for the message was already a recipient.|info|
|EXPAND |A distribution group was expanded.|info|
|FAIL |Message delivery failed. Sources include SMTP, DNS, QUEUE, and ROUTING.|problem|
|HADISCARD |A shadow message was discarded after the primary copy was delivered to the next hop. For more information, see Shadow redundancy in Exchange Server.|info|
|HARECEIVE |A shadow message was received by the server in the local database availability group (DAG) or Active Directory site.|info|
|HAREDIRECT |A shadow message was created.|info|
|HAREDIRECTFAIL |A shadow message failed to be created. The details are stored in the source-context field.|problem|
|INITMESSAGECREATED |A message was sent to a moderated recipient, so the message was sent to the arbitration mailbox for approval. For more information, see Manage message approval.|problem|
|LOAD |A message was successfully loaded at boot.|info|
|MODERATIONEXPIRE |A moderator for a moderated recipient never approved or rejected the message, so the message expired. For more information about moderated recipients, see Manage message approval.|problem|
|MODERATORAPPROVE |A moderator for a moderated recipient approved the message, so the message was delivered to the moderated recipient.|transport|
|MODERATORREJECT |A moderator for a moderated recipient rejected the message, so the message wasn't delivered to the moderated recipient.|problem|
|MODERATORSALLNDR |All approval requests sent to all moderators of a moderated recipient were undeliverable, and resulted in non-delivery reports (also known as NDRs or bounce messages).|problem|
|NOTIFYMAPI |A message was detected in the Outbox of a mailbox on the local server.|info|
|NOTIFYSHADOW |A message was detected in the Outbox of a mailbox on the local server, and a shadow copy of the message needs to be created.|info|
|POISONMESSAGE |A message was put in the poison message queue or removed from the poison message queue.|problem|
|PROCESS |The message was successfully processed.|info|
|PROCESSMEETINGMESSAGE |A meeting message was processed by the Mailbox Transport Delivery service.|info|
|RECEIVE |A message was received by the SMTP receive component of the transport service or from the Pickup or Replay directories (source: SMTP), or a message was submitted from a mailbox to the Mailbox Transport Submission service (source: STOREDRIVER).|transport|
|REDIRECT |A message was redirected to an alternative recipient after an Active Directory lookup.|transport|
|RESOLVE |A message's recipients were resolved to a different email address after an Active Directory lookup.|info|
|RESUBMIT |A message was automatically resubmitted from Safety Net. For more information, see Safety Net in Exchange Server.|info|
|RESUBMITDEFER |A message resubmitted from Safety Net was deferred.|problem|
|RESUBMITFAIL |A message resubmitted from Safety Net failed.|problem|
|SEND |A message was sent by SMTP between transport services.|transport|
|SUBMIT |The Mailbox Transport Submission service successfully transmitted the message to the Transport service. For SUBMIT events, the source-context property contains the following details:<br/> * MDB: The mailbox database GUID.<br/> * Mailbox: The mailbox GUID.<br/> * Event: The event sequence number.<br/> * MessageClass: The type of message. For example, IPM.Note.<br/> * CreationTime: Date-time of the message submission.<br/> * ClientType: For example, User, OWA, or ActiveSync.|info|
|SUBMITDEFER |The message transmission from the Mailbox Transport Submission service to the Transport service was deferred.|problem|
|SUBMITFAIL |The message transmission from the Mailbox Transport Submission service to the Transport service failed.|problem|
|SUPPRESSED |The message transmission was suppressed.|problem|
|THROTTLE |The message was throttled.|problem|
|TRANSFER |Recipients were moved to a forked message because of content conversion, message recipient limits, or agents. Sources include ROUTING or QUEUE.|info|
"@ ;
    $ExoEventIDs = $ExoEventIDsTbl | convertfrom-markdowntable ;

    # building a CustObj (actually an indexed hash) with the EventName|Descriptions. The 'index' for each event, is the EventName 
    $smsg = "(converting $(($ExoEventIDs|measure).count) Get-xoMessageTrace-supported Event-ID types, to indexed hash)" ;     
    if($verbose){
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
    } ; 
    if($host.version.major -gt 2){$hExoEventIDs = [ordered]@{} } 
    else { $hExoEventIDs = @{} } ;
    
    $ttl = ($ExoEventIDs|measure).count ; $Procd = 0 ; 
    foreach ($EID in $ExoEventIDs){
        $Procd ++ ; 
        $sBnrS="`n#*------v PROCESSING : ($($Procd)/$($ttl)) $($EID.EventName) v------" ; 
        $smsg = $sBnrS ; 
        if($verbose){
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        } ; 
        
        $name =$EID.EventName ; 
        $hExoEventIDs[$name] = $EID ; 
        if($EID.Description -match '<br/>'){
            $hExoEventIDs[$name].Description = $hExoEventIDs[$name].Description -replace '<br/>','\n'
        } ;

        $smsg = "$($sBnrS.replace('-v','-^').replace('v-','^-'))" ;
        if($verbose){
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        } ; 
    } ;  # loop-E

    if($hExoEventIDs){
        if($EventID){
            foreach($Evt in $EventID){
                if($hexoeventids[$Evt]){
                    $smsg = "(Returning matched EventID:$($Evt) details to pipeline)" ; 
                    if($verbose){
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ; 
                    $hexoeventids[$Evt] | Write-Output ; 
                }else {
                    $smsg = "Unable to resolve EventID: $($Evt) to a matching documented event-id string!" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                    throw $smsg ; 
                } ; 
            }  ; 
        }elseif($Type){
            foreach($Typ in $Type){
                $smsg = "(Returning matched Type:$($Typ) event details to pipeline)" ; 
                if($verbose){
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                } ; 
                $hexoeventids.values | ?{$_.type -eq $Typ} | write-output ;
            } ; 
        }else{
            $smsg = "(Returning full set of summary objects to pipeline)" ; 
            if($verbose){
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } ; 
            $hexoeventids | Write-Output ;
        } ; 
    } else {
        $smsg = "NO RETURNABLE `$hExoEventIDs OBJECT!" ; 
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug 
        else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
        THROW $smsg ;
    } ; 
}
#*------^ get-ExoMessageTraceEventType.ps1 ^------