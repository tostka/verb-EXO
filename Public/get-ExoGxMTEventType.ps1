#*------v get-ExoGxMTEventType.ps1 v------
function get-ExoGxMTEventType {
<#
    .SYNOPSIS
    get-ExoGxMTEventType - Provides a prefab array indexed hash of Exchange-Online Get-xoMessageTrace Event Types Note: This is a static non-query-based list of events. The function must be manually updated to accomodate MS MessageTrace event types changes over time).
    .PARAMETER Mailboxes
    .NOTES
    Version     : 1.0.0
    Author      : Todd Kadrie
    Website     :	http://www.toddomation.com
    Twitter     :	@tostka / http://twitter.com/tostka
    CreatedDate : 2022-02-25
    FileName    : get-ExoGxMTEventType.ps1
    License     : MIT License
    Copyright   : (c) 2022 Todd Kadrie
    Github      : https://github.com/tostka/verb-ex2010
    Tags        : Powershell
    REVISIONS
    * 10:27 AM 5/17/2024 addapted from vxo\get-ExoMailboxLicenses()
    .DESCRIPTION
    get-ExoGxMTEventType - Provides a prefab array indexed hash of Exchange-Online Get-xoMessageTrace Event Types Note: This is a static non-query-based list of events. The function must be manually updated to accomodate MS MessageTrace event types changes over time).
    .PARAMETER EventID
    MessageTrace event-id to be resolved to details[-EventID 'AGENTINFO']
    .EXAMPLE
    PS> $eventInfo = get-ExoGxMTEventType -eventid AGENTINFO -verbose ; 
    PS> $hQuotas['database2']
    Name           ProhibitSendReceiveQuotaGB ProhibitSendQuotaGB IssueWarningQuotaGB
    ----           -------------------------- ------------------- -------------------
    database2      12.000                     10.000              9.000
    Retrieve local org on-prem MailboxDatabase quotas and assign to a variable, with verbose outputs. Then output the retrieved quotas from the indexed hash returned, for the mailboxdatabase named 'database2'.
    .LINK
    https://github.com/tostka/verb-EXO
    #>
    #Requires -Modules verb-IO, verb-logging, verb-Text
    [OutputType('System.Collections.Hashtable')]
    [CmdletBinding()]
    PARAM(
        [Parameter(Mandatory=$FALSE,HelpMessage="MessageTrace event-id to be resolved to details[-EventID 'AGENTINFO']")]
        [ValidateNotNullOrEmpty()]
        [string]$EventID
    ) ;
    
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
|EventName |Description|
|---|---|
|AGENTINFO |This event is used by transport agents to log custom data.|
|BADMAIL |A message submitted by the Pickup directory or the Replay directory that can't be delivered or returned.|
|CLIENTSUBMISSION |A message was submitted from the Outbox of a mailbox.|
|DEFER |Message delivery was delayed.|
|DELIVER |A message was delivered to a local mailbox.|
|DELIVERFAIL |An agent tried to deliver the message to a folder that doesn't exist in the mailbox.|
|DROP |A message was dropped without a delivery status notification (also known as a DSN, bounce message, non-delivery report, or NDR). For example:<br/>* Completed moderation approval request messages.<br/>* Spam messages that were silently dropped without an NDR.|
|DSN |A delivery status notification (DSN) was generated.|
|DUPLICATEDELIVER |A duplicate message was delivered to the recipient. Duplication may occur if a recipient is a member of multiple nested distribution groups. Duplicate messages are detected and removed by the information store.|
|DUPLICATEEXPAND |During the expansion of the distribution group, a duplicate recipient was detected.|
|DUPLICATEREDIRECT |An alternate recipient for the message was already a recipient.|
|EXPAND |A distribution group was expanded.|
|FAIL |Message delivery failed. Sources include SMTP, DNS, QUEUE, and ROUTING.|
|HADISCARD |A shadow message was discarded after the primary copy was delivered to the next hop. For more information, see Shadow redundancy in Exchange Server.|
|HARECEIVE |A shadow message was received by the server in the local database availability group (DAG) or Active Directory site.|
|HAREDIRECT |A shadow message was created.|
|HAREDIRECTFAIL |A shadow message failed to be created. The details are stored in the source-context field.|
|INITMESSAGECREATED |A message was sent to a moderated recipient, so the message was sent to the arbitration mailbox for approval. For more information, see Manage message approval.|
|LOAD |A message was successfully loaded at boot.|
|MODERATIONEXPIRE |A moderator for a moderated recipient never approved or rejected the message, so the message expired. For more information about moderated recipients, see Manage message approval.|
|MODERATORAPPROVE |A moderator for a moderated recipient approved the message, so the message was delivered to the moderated recipient.|
|MODERATORREJECT |A moderator for a moderated recipient rejected the message, so the message wasn't delivered to the moderated recipient.|
|MODERATORSALLNDR |All approval requests sent to all moderators of a moderated recipient were undeliverable, and resulted in non-delivery reports (also known as NDRs or bounce messages).|
|NOTIFYMAPI |A message was detected in the Outbox of a mailbox on the local server.|
|NOTIFYSHADOW |A message was detected in the Outbox of a mailbox on the local server, and a shadow copy of the message needs to be created.|
|POISONMESSAGE |A message was put in the poison message queue or removed from the poison message queue.|
|PROCESS |The message was successfully processed.|
|PROCESSMEETINGMESSAGE |A meeting message was processed by the Mailbox Transport Delivery service.|
|RECEIVE |A message was received by the SMTP receive component of the transport service or from the Pickup or Replay directories (source: SMTP), or a message was submitted from a mailbox to the Mailbox Transport Submission service (source: STOREDRIVER).|
|REDIRECT |A message was redirected to an alternative recipient after an Active Directory lookup.|
|RESOLVE |A message's recipients were resolved to a different email address after an Active Directory lookup.|
|RESUBMIT |A message was automatically resubmitted from Safety Net. For more information, see Safety Net in Exchange Server.|
|RESUBMITDEFER |A message resubmitted from Safety Net was deferred.|
|RESUBMITFAIL |A message resubmitted from Safety Net failed.|
|SEND |A message was sent by SMTP between transport services.|
|SUBMIT |The Mailbox Transport Submission service successfully transmitted the message to the Transport service. For SUBMIT events, the source-context property contains the following details:<br/> * MDB: The mailbox database GUID.<br/> * Mailbox: The mailbox GUID.<br/> * Event: The event sequence number.<br/> * MessageClass: The type of message. For example, IPM.Note.<br/> * CreationTime: Date-time of the message submission.<br/> * ClientType: For example, User, OWA, or ActiveSync.|
|SUBMITDEFER |The message transmission from the Mailbox Transport Submission service to the Transport service was deferred.|
|SUBMITFAIL |The message transmission from the Mailbox Transport Submission service to the Transport service failed.|
|SUPPRESSED |The message transmission was suppressed.|
|THROTTLE |The message was throttled.|
|TRANSFER |Recipients were moved to a forked message because of content conversion, message recipient limits, or agents. Sources include ROUTING or QUEUE.|
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
    foreach ($Sku in $ExoEventIDs){
        $Procd ++ ; 
        $sBnrS="`n#*------v PROCESSING : ($($Procd)/$($ttl)) $($Sku.SKU) v------" ; 
        $smsg = $sBnrS ; 
        if($verbose){
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        } ; 
        
        $name =$($Sku | select -expand SKU) ; 
        $hExoEventIDs[$name] = $Sku ; 

        $smsg = "$($sBnrS.replace('-v','-^').replace('v-','^-'))" ;
        if($verbose){
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        } ; 
    } ;  # loop-E

    if($hExoEventIDs){
        $smsg = "(Returning summary objects to pipeline)" ; 
        if($verbose){
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        } ; 
        $hExoEventIDs | Write-Output ; 
    } else {
        $smsg = "NO RETURNABLE `$hExoEventIDs OBJECT!" ; 
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug 
        else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
        THROW $smsg ;
    } ; 
}

#*------^ get-ExoGxMTEventType.ps1 ^------