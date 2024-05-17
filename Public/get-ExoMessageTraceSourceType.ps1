# get-ExoMessageTraceSourceType.ps1

#*------v get-ExoMessageTraceSourceType.ps1 v------
function get-ExoMessageTraceSourceType {
<#
    .SYNOPSIS
    get-ExoMessageTraceSourceType - Provides a prefab array indexed hash of Exchange-Online Get-xoMessageTrace Source Types Note: This is a static non-query-based list of sourcess. The function must be manually updated to accomodate MS MessageTrace sources types changes over time.
    .PARAMETER Mailboxes
    .NOTES
    Version     : 1.0.0
    Author      : Todd Kadrie
    Website     :	http://www.toddomation.com
    Twitter     :	@tostka / http://twitter.com/tostka
    CreatedDate : 2024-05-17
    FileName    : get-ExoMessageTraceSourceType.ps1
    License     : MIT License
    Copyright   : (c) 2022 Todd Kadrie
    Github      : https://github.com/tostka/verb-ex2010
    Tags        : Powershell,ExchangeOnline,MessageTrace,Reporting
    REVISIONS
    * 3:54 PM 5/17/2024 adapted from vxo\get-ExoMessageTraceEventType(); ; added array support on the EventID & Type
    .DESCRIPTION
    get-ExoMessageTraceSourceType - Provides a prefab array indexed hash of Exchange-Online Get-xoMessageTrace Source Types Note: This is a static non-query-based list of Sources. The function must be manually updated to accomodate MS MessageTrace Source types changes over time.
    Run it without a -Source and it returns the entire indexed hash to the pipeline

    
    .PARAMETER Source
    MessageTrace Sourceid to be resolved to details[-Source 'AGENTINFO']
    .PARAMETER Type
    Optional type classification subset to be returned from the entire Source set[-Type 'Problem']
    .EXAMPLE
    PS> $SourceInfo = get-ExoMessageTraceSourceType -Source AGENTINFO -verbose ; 
    PS> $SourceInfo | out-string ; 

        SourceValue Description                                                                                    Type     
        ----------- -----------                                                                                    ----     
        SMTP        The message was submitted by the SMTP send or SMTP receive component of the transport service. transport

    Retrieve Resolve 'SMTP' to it's stock MS description & type, and output to console
    .EXAMPLE
    PS> write-verbose 'Retrive the entire indexed hash of Source details, and output the results to console' ; 
    PS> $SourcesDetail = get-ExoMessageTraceSourceType ; 
    PS> $SourcesDetail | out-string ; 

        Name                           Value                                                                                                                                                                                                                             
        ----                           -----                                                                                                                                                                                                                             
        AGENTINFO                      @{EventName=AGENTINFO; Description=This event is used by transport agents to log custom data.; Type=info}                   
        ...[TRIMMED]...
        TRANSFER                       @{EventName=TRANSFER; Description=Recipients were moved to a forked message because of content conversion, message recipient limits, or agents. Sources include ROUTING or QUEUE.; Type=info}                                     

    Retrive the entire indexed hash of EventID details, and output the results to console. 
    PS> $SourcesDetail['SMTP'] | out-string ; 

        SourceValue Description                                                                                    Type     
        ----------- -----------                                                                                    ----     
        SMTP        The message was submitted by the SMTP send or SMTP receive component of the transport service. transport

    PS> $SourcesDetail.values | group type |  ft -a count,name 

        Count Name     
        ----- ----     
            6 problem  
           12 info     
            2 transport

    PS> $rgxSrcProblem = ('(' + (($SourcesDetail.values | ?{$_.type -eq 'problem'} | select -expand SourceValue |%{[regex]::escape($_)}) -join '|') + ')') ;
    PS> $rgxSrcTransp = ('(' + (($SourcesDetail.values | ?{$_.type -eq 'transport'} | select -expand SourceValue|%{[regex]::escape($_)}) -join '|') + ')') ;
    PS> $rgxSrcInfo =  ('(' + (($SourcesDetail.values | ?{$_.type -eq 'info'} | select -expand SourceValue|%{[regex]::escape($_)}) -join '|') + ')') ;
    PS> $rgxSrcProblem

        (ADMIN|APPROVAL|BOOTLOADER|DSN|PICKUP|POISONMESSAGE)

    PS> $rgxSrcInfo

        (AGENT|DNS|GATEWAY|MAILBOXRULE|MEETINGMESSAGEPROCESSOR|ORAR|PUBLICFOLDER|QUEUE|REDUNDANCY|RESOLVER|ROUTING|SAFETYNET)

    PS> $rgxSrcTransp 

        (SMTP|STOREDRIVER)

    Retrieve the entire set of defined event-id's returned as an indexed hash. Output the hash to console; lookup and return the AGENTINFO event-id details; group the assigned types; Build & output regexes for the Problem, Transport & Info types.
    .EXAMPLE
    PS> $SrcInfo = get-ExoMessageTraceSourceType -type 'problem','transport' ; 
    PS> $SrcInfo | write-output ;

        SourceValue   Description                                                                                                                                                          Type     
        -----------   -----------                                                                                                                                                          ----     
        ADMIN         The event source was human intervention. For example, an administrator used Queue Viewer to delete a message, or submitted message files using the Replay directory. problem  
        APPROVAL      The event source was the approval framework that's used with moderated recipients. For more information, see Manage message approval.                                problem  
        BOOTLOADER    The event source was unprocessed messages that exist on the server at boot time. This is related to the LOAD event type.                                             problem  
        DSN           The event source was a delivery status notification (also known as a DSN, bounce message, non-delivery report, or NDR).                                              problem  
        PICKUP        The event source was the Pickup directory. For more information, see Pickup Directory and Replay Directory.                                                          problem  
        POISONMESSAGE The event source was the poison message identifier. For more information about poison messages and the poison message queue, see Queues and messages in queues       problem  
        SMTP          The message was submitted by the SMTP send or SMTP receive component of the transport service.                                                                       transport
    
    PS> write-verbose 'Create a regex from the Source names returned' ;
    PS> $rgxEvtIDTransProb = ('(' + (($SrcInfo.SourceValue |%{[regex]::escape($_)}) -join '|') + ')') ;
    PS> $rgxEvtIDTransProb

        (ADMIN|APPROVAL|BOOTLOADER|DSN|PICKUP|POISONMESSAGE|SMTP|STOREDRIVER)

    Retrieve the Type:Problem & Transport EventIDs and build a regex out of the combonation
    .LINK
    https://github.com/tostka/verb-EXO
    #>
    #Requires -Modules verb-IO, verb-logging, verb-Text
    [OutputType('System.Collections.Hashtable')]
    [CmdletBinding()]
    PARAM(
        [Parameter(Mandatory=$FALSE,HelpMessage="Optional MessageTrace Source Value to be resolved to details[-Source 'AGENTINFO']")]
            [ValidateNotNullOrEmpty()]
            [string[]]$Source,
        [Parameter(Mandatory=$FALSE,HelpMessage="Optional type classification subset to be returned from the entire Source set[-Type 'Problem']")]
            [ValidateSet('transport','problem','info')]
            [string[]]$Type
    ) ;
    if($Source -AND $Type){
        $smsg = "Both -Source & -Type specified: Please specify one or the other (or none, to return all defined Sources)" ; 
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
    $ExoSourcesTbl = @"
|SourceValue |Description|Type|
|---|---|---|
|ADMIN |The event source was human intervention. For example, an administrator used Queue Viewer to delete a message, or submitted message files using the Replay directory.|problem|
|AGENT |The event source was a transport agent.|info|
|APPROVAL |The event source was the approval framework that's used with moderated recipients. For more information, see Manage message approval.|problem|
|BOOTLOADER |The event source was unprocessed messages that exist on the server at boot time. This is related to the LOAD event type.|problem|
|DNS |The event source was DNS.|info|
|DSN |The event source was a delivery status notification (also known as a DSN, bounce message, non-delivery report, or NDR).|problem|
|GATEWAY |The event source was a Foreign connector. For more information, see Foreign Connectors.|info|
|MAILBOXRULE |The event source was an Inbox rule. For more information, see Inbox rules.|info|
|MEETINGMESSAGEPROCESSOR |The event source was the meeting message processor, which updates calendars based on meeting updates.|info|
|ORAR |The event source was an Originator Requested Alternate Recipient (ORAR). You can enable or disable support for ORAR on Receive connectors using the OrarEnabled parameter on the New-ReceiveConnector or Set-ReceiveConnector cmdlets.|info|
|PICKUP |The event source was the Pickup directory. For more information, see Pickup Directory and Replay Directory.|problem|
|POISONMESSAGE |The event source was the poison message identifier. For more information about poison messages and the poison message queue, see Queues and messages in queues|problem|
|PUBLICFOLDER |The event source was a mail-enabled public folder.|info|
|QUEUE |The event source was a queue.|info|
|REDUNDANCY |The event source was Shadow Redundancy. For more information, see Shadow redundancy in Exchange Server.|info|
|RESOLVER |The event source was the recipient resolution component of the categorizer in the Transport service. For more information, see Recipient resolution in Exchange Server.|info|
|ROUTING |The event source was the routing resolution component of the categorizer in the Transport service.|info|
|SAFETYNET |The event source was Safety Net. For more information, see Safety Net in Exchange Server.|info|
|SMTP |The message was submitted by the SMTP send or SMTP receive component of the transport service.|transport|
|STOREDRIVER |The event source was a MAPI submission from a mailbox on the local server.|transport|
"@ ;
    $ExoSources = $ExoSourcesTbl | convertfrom-markdowntable ;

    # building a CustObj (actually an indexed hash) with the SourceValue|Descriptions. The 'index' for each event, is the SourceValue 
    $smsg = "(converting $(($ExoSources|measure).count) Get-xoMessageTrace-supported Source Value types, to indexed hash)" ;     
    if($verbose){
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
    } ; 
    if($host.version.major -gt 2){$hExoSources = [ordered]@{} } 
    else { $hExoSources = @{} } ;
    
    $ttl = ($ExoSources|measure).count ; $Procd = 0 ; 
    foreach ($EID in $ExoSources){
        $Procd ++ ; 
        $sBnrS="`n#*------v PROCESSING : ($($Procd)/$($ttl)) $($EID.SourceValue) v------" ; 
        $smsg = $sBnrS ; 
        if($verbose){
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        } ; 
        
        $name =$EID.SourceValue ; 
        $hExoSources[$name] = $EID ; 
        if($EID.Description -match '<br/>'){
            $hExoSources[$name].Description = $hExoSources[$name].Description -replace '<br/>','\n'
        } ;

        $smsg = "$($sBnrS.replace('-v','-^').replace('v-','^-'))" ;
        if($verbose){
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        } ; 
    } ;  # loop-E

    if($hExoSources){
        if($Source){
            foreach($Src in $Source){
                if($hexoSources[$Src]){
                    $smsg = "(Returning matched Source:$($Src) details to pipeline)" ; 
                    if($verbose){
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ; 
                    $hexoSources[$Src] | Write-Output ; 
                }else {
                    $smsg = "Unable to resolve Source: $($Src) to a matching documented Source Value string!" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                    throw $smsg ; 
                } ; 
            } ; 
        }elseif($Type){
            foreach($Typ in $Type){
                $smsg = "(Returning matched Type:$($Typ) event details to pipeline)" ; 
                if($verbose){
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                } ; 
                $hexoSources.values | ?{$_.type -eq $typ} | write-output 
            } ; 
        }else{
            $smsg = "(Returning full set of summary objects to pipeline)" ; 
            if($verbose){
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } ; 
            $hexoSources | Write-Output ;
        } ; 
    } else {
        $smsg = "NO RETURNABLE `$hExoSources OBJECT!" ; 
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug 
        else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
        THROW $smsg ;
    } ; 
}
#*------^ get-ExoMessageTraceSourceType.ps1 ^------