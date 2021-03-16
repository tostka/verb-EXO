#*----------v Function get-MsgTrace() v----------
function get-MsgTrace {
    <#
    .SYNOPSIS
    get-MsgTrace.ps1 - Perform smart get-exoMessageTrace/MessageTrackingLog command, as appropriate to target location -Mailbox, on either Exchange on-premesis or Exchange Online.
    .NOTES
    Version     : 1.0.0.
    Author      : Todd Kadrie
    Website     : http://www.toddomation.com
    Twitter     : @tostka / http://twitter.com/tostka
    CreatedDate : 2021-03-12
    FileName    : get-MsgTrace.ps1
    License     : MIT License
    Copyright   : (c) 2020 Todd Kadrie
    Github      : https://github.com/tostka/verb-exo
    Tags        : Powershell,ExchangeOnline,Mailbox,Statistics,Reporting
    REVISIONS
    * 2:23 PM 3/16/2021 added multi-tenant support ; debugged both exOP & exo, added -ReportFail & -ReportRowsLimit params. At this point Exclusive params are only partially configured
    * 1:12 PM 3/15/2021 init work was done 3/12, removed recursive-err generating #Require on the hosting verb-EXO module
    .DESCRIPTION
    get-MsgTrace - Perform smart get-exoMessageTrace/MessageTrackingLog command, as appropriate to target location -Mailbox, on either Exchange on-premesis or Exchange Online.
    Dependancy on my verb-ex2010 Exchange onprem (and is within verb-exo EXO mod, which adds dependant EXO connection support).
    .PARAMETER TenOrg
    TenantTag values, indicating Tenant to Create DDG WITHIN[-TenOrg 'TOL']    
    .PARAMETER Recipients
    Recipient email addresses identifiers (comma-delimited)[-Recipients xxx@domain.com]
    .PARAMETER Sender
    Sender email address identifiers (EXO supports comma-delimited) [-Sender xxx@domain.com]
    .PARAMETER Subject
    "Message Subject string to be matched (post-filtered from broad query)[-Subject 'subject phrase']
    .PARAMETER Logon
    User Logon tag to be applied to output file[-Logon samaccountname]
    .PARAMETER Status
    Transport Status (EventID on-Prem)(RECEIVE|DELIVER|FAIL|SEND|RESOLVE|EXPAND|TRANSFER|DEFER) [-EventID SEND
    .PARAMETER Connectorid
    Connector identifier[-Connectorid SendConnX]
    .PARAMETER Source
    Source keyword to be used for filtering (STOREDRIVER|SMTP|DNS|ROUTING)[-Source SMTP]
    .PARAMETER MessageId
    "Target MessageId for search[-MessageId xxxxxxx]
    .PARAMETER MessageTraceId
    Target MessageId for search[-MessageTraceId xxxxxxx]
    .PARAMETER StartDate
    Start of time span to be searched[-StartDate 1/1/2021]
    .PARAMETER EndDate
    End of time span to be searched[-EndDate 1/7/2021]
    .PARAMETER Ticket
    Ticket # [-Ticket nnnnn]
    .PARAMETER useEXOP
    Switch to specify ONPREM Exch get-MessageTrackingLog trace (defaults `$false == EXO Message Search)[-useEXOP]
    .PARAMETER ReportRowsLimit
    Max number of rows to output to console when a -ReportXXX param is specified (defaults 100)[-ReportRowsLimit]
    .PARAMETER asObject
    Switch to return raw object rather than formated console report(defaults `$true)[-NonEmptyOnly]
    .INPUTS
    Accepts piped input.
    .OUTPUTS
    Outputs csv & console summary of mailbox folders content
    .EXAMPLE
    get-MsgTrace -Sender SENDER@DOMAIN.com -Ticket 99999 -days 7 -verbose ;
    Perform a default EXO trace last 7 days of traffic on specified sender, use specified Ticket number in csv file name, with verbose output
    .EXAMPLE
    $msgs = get-MsgTrace -Sender quotes@bossplow.com -Ticket 347298 -days 7 -asobject -verbose ;
    Above EXO MessageTrace returning an object for further postfiltering.
    .EXAMPLE
    get-msgtrace -sender monitoring@toro.com -useEXOP -ticket 99999 -d 1 -verbose ; 
    Run an ONPREM get-MessageTrackingLog search
    .EXAMPLE 
    $msgs = get-msgtrace -sender monitoring@toro.com -useEXOP -ticket 99999 -start (get-date).addhours(-1) -verbose -ReportFail; 
    Run an ONPREM get-MessageTrackingLog search, with specific -Start time (End will be asserted), with detailed dump of (first 100) EventID 'Fail' items
    .LINK
    https://github.com/tostka/verb-exo
    .LINK
    https://github.com/tostka/verb-ex2010
    #>
    #Requires -Version 3
    #Requires -Modules verb-ex2010
    # VALIDATORS: [ValidateNotNull()][ValidateNotNullOrEmpty()][ValidateLength(24,25)][ValidateLength(5)][ValidatePattern("(lyn|bcc|spb|adl)ms6(4|5)(0|1).(china|global)\.ad\.toro\.com")][ValidateSet("USEA","GBMK","AUSYD")][ValidateScript({Test-Path $_ -PathType 'Container'})][ValidateScript({Test-Path $_})][ValidateRange(21,65)][ValidateCount(1,3)]
    [CmdletBinding(DefaultParameterSetName='SendRec')]
    <# $isplt=@{  ticket="347298" ;  uid="wilinaj";  days=7 ;  Sender="quotes@bossplow.com" ;  Recipients="" ;  MessageSubject="" ;  EventID='' ;  Connectorid="" ;  Source="" ;} ; 
    #>
    Param(
        [Parameter(Mandatory=$FALSE,HelpMessage="TenantTag values, indicating Tenant to Create DDG WITHIN[-TenOrg 'TOL']")]
        [ValidateNotNullOrEmpty()]
        [string]$TenOrg = ('TOR'),
        [Parameter(ParameterSetName='SendRec',HelpMessage="Recipient email addresses identifiers (comma-delimited)[-Recipients xxx@domain.com]")]
        [string]$Recipients,    
        [Parameter(ParameterSetName='SendRec',HelpMessage="Sender email address identifier (EXO supports comma-delimited)")]
        [string]$Sender, 
        [Parameter(HelpMessage="Message Subject string to be matched (post-filtered from broad query)[-Subject 'subject phrase']")]
        [string]$Subject,
        [Parameter(HelpMessage="User Logon tag to be applied to output file[-Logon samaccountname]")]
        [string]$Logon,
        [Parameter(HelpMessage="Transport Status (EventID on-Prem)(RECEIVE|DELIVER|FAIL|SEND|RESOLVE|EXPAND|TRANSFER|DEFER) [-EventID SEND")]
        [ValidateSet("RECEIVE","DELIVER","FAIL","SEND","RESOLVE","EXPAND","TRANSFER","DEFER")]
        [string]$Status,
        [Parameter(HelpMessage="Connector identifier[-Connectorid SendConnX]")]
        [string]$Connectorid,
        [Parameter(HelpMessage="Source keyword to be used for filtering (STOREDRIVER|SMTP|DNS|ROUTING)[-Source SMTP]")]
        [ValidateSet("STOREDRIVER","SMTP","DNS","ROUTING")]
        [string]$Source,
        [Parameter(ParameterSetName='MsgID',HelpMessage="Target MessageId for search[-MessageId xxxxxxx]")]
        [string]$MessageId, 
        [Parameter(ParameterSetName='MsgTrcID',HelpMessage="Target MessageId for search[-MessageTraceId xxxxxxx]")]
        [string]$MessageTraceId,
        [Parameter(HelpMessage="Start of time span to be searched[-StartDate 1/1/2021]")]
        [string]$StartDate,
        [Parameter(HelpMessage="End of time span to be searched[-EndDate 1/7/2021]")]
        [string]$EndDate,
        [Parameter(HelpMessage="Days back to search[-Days 7]")]
        [int]$Days,
        [Parameter(Mandatory=$false,HelpMessage="Ticket # [-Ticket nnnnn]")]
        #[ValidateLength(5)] # non-mandatory
        [int]$Ticket,
        [Parameter(HelpMessage="Switch to specify ONPREM Exch get-MessageTrackingLog trace (defaults `$false == EXO Message Search)[-useEXOP]")]
        [switch] $useEXOP=$false,
        [Parameter(HelpMessage="Switch to return raw object rather than formated console report(defaults `$true)[-NonEmptyOnly]")]
        [switch] $asObject,
        [Parameter(HelpMessage="Switch to return detailed analysis of FAIL items[-ReportFail]")]
        [switch] $ReportFail,
        [Parameter(HelpMessage="Max number of rows to output to console when a -ReportXXX param is specified (defaults 100)[-ReportRowsLimit]")]
        [int]$ReportRowsLimit = 100  
    ) ;
    BEGIN {
        $Verbose=($VerbosePreference -eq 'Continue') ;  
        # Get parameters this function was invoked with
        $PSParameters = New-Object -TypeName PSObject -Property $PSBoundParameters ;
        $propsFldr = @{Name='Folder';Expression={$_.Identity.tostring()}},@{Name="Items";Expression={$_.ItemsInFolder}} ;
        $propsMsgEx10 = 'Timestamp',@{N='TimestampLocal';E={$_.Timestamp.ToLocalTime()}},'Source','EventId','RelatedRecipientAddress','Sender',@{N='Recipients';E={$_.Recipients}},"RecipientCount",@{N='RecipientStatus';E={$_.RecipientStatus}},"MessageSubject","TotalBytes",@{N='Reference';E={$_.Reference}},'MessageLatency','MessageLatencyType','InternalMessageId','MessageId','ReturnPath','ClientIp','ClientHostname','ServerIp','ServerHostname','ConnectorId','SourceContext','MessageInfo',@{N='EventData';E={$_.EventData}} ;
        $propsMsgEXO = @{N='ReceivedLocal';E={$_.Received.ToLocalTime()}},'SenderAddress','RecipientAddress','Subject','Status','ToIP','FromIP','Size','MessageId','MessageTraceId','Index' ;
        
        # pull settings per Tenant fr Meta
        $Meta = gv -name "$($TenOrg)Meta" ; 
        <# pull value fr meta
        if($Meta -is [system.array]){ throw "Unable to resolve unique `$xxxMeta! from `$TenOrg:$($TenOrg)" ; break} ; 
        if(!$Meta.value.DefaultObjectOwner){throw "Unable to resolve $($Meta.Name).value.DefaultObjectOwner from `$TenOrg:$($TenOrg)" ; break} 
        else { $ManagedBy=$Meta.value.DefaultObjectOwner} ;  ;
        #>

        $Retries = 4 ;
        $RetrySleep = 5 ;
        if(!$ThrottleMs){$ThrottleMs = 50 ;}
        $CredRole = 'CSVC' ; # role of svc to be dyn pulled from metaXXX if no -Credential spec'd, 
        if(!$rgxDriveBanChars){$rgxDriveBanChars = '[;~/\\\.:]' ; } ; # ;~/\.:, 
        
        if($useEXOP){
            $UseOP=$false ; 
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
        } else { 
            # o365/EXO creds
            $o365Cred=$null ;
            <# Usage: Type defaults to SID, if not spec'd - Note: there must be a *logged in & configured *profile* 
            $o365Cred=get-TenantCredentials -TenOrg $TenOrg -verbose -userrole SID ;
            Returns a credential set for the $TenOrg Hybrid OnPrem Exchange Org
            .EXAMPLE
            $o365Cred=get-TenantCredentials -TenOrg $TenOrg -verbose -userrole CSVC ;
            Returns the CSVC Userrole credential for the $TenOrg Hybrid OnPrem Exchange Org
            .EXAMPLE
            $o365Cred=get-TenantCredentials -TenOrg $TenOrg -verbose -userrole B2BI ;
            Returns the B2BI Userrole credential for the $TenOrg Hybrid OnPrem Exchange Org
            #>
            #if($o365Cred=(get-TenantCredentials -TenOrg $TenOrg -verbose:$($verbose))){
            # force it to use the csvc mapping from $xxxmeta.o365_CSvcUpn, failthrough to SID spec 
            if($o365Cred=(get-TenantCredentials -TenOrg $TenOrg -UserRole 'CSVC','SID' -verbose:$($verbose))){
                # make it script scope, so we don't have to predetect & purge before using new-variable
                New-Variable -Name cred$($tenorg) -scope Script -Value $o365Cred.cred ;
                $smsg = "Resolved $($Tenorg) `$o365cred:$($o365Cred.cred.username) (assigned to `$cred$($tenorg))" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } else {
                #-=-record a STATUS=-=-=-=-=-=-=
                $statusdelta = ";ERROR";
                $script:PassStatus += $statusdelta ;
                set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta) ;
                #-=-=-=-=-=-=-=-=
                $smsg = "Unable to resolve $($tenorg) `$o365Cred value!"
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                throw "Unable to resolve $($tenorg) `$o365Cred value!`nEXIT!"
                exit ;
            } ;
            <# CALLS ARE IN FORM: (cred$($tenorg))
            $pltRXO = @{Credential = $Credential ; verbose = $($verbose) ; }
            $pltRXO = @{
                Credential = (Get-Variable -name cred$($tenorg) ).value ;
                verbose = $($verbose) ; } ; 
            if ($script:useEXOv2) { reconnect-eXO2 @pltRXO }
            else { reconnect-EXO @pltRXO } ;
            # or with Tenant-specific cred($Tenorg) lookup
            #>
        } ; 

        if($UseOP){
            # do the OP creds too
            $OPCred=$null ;
            # default to the onprem svc acct
            $pltGHOpCred=@{TenOrg=$TenOrg ;userrole='ESVC' ;verbose=$($verbose)} ;
            if($OPCred=(get-HybridOPCredentials @pltGHOpCred).cred){
                # make it script scope, so we don't have to predetect & purge before using new-variable
                New-Variable -Name "cred$($tenorg)OP" -scope Script -Value $OPCred ;
                $smsg = "Resolved $($Tenorg) `$OPCred:$($OPCred.username) (assigned to `$cred$($tenorg)OP)" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } else {
                #-=-record a STATUS=-=-=-=-=-=-=
                $statusdelta = ";ERROR";
                $script:PassStatus += $statusdelta ;
                set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta) ;
                #-=-=-=-=-=-=-=-=
                $smsg = "Unable to resolve get-HybridOPCredentials -TenOrg $($TenOrg) -userrole 'ESVC' value!"
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                throw "Unable to resolve $($tenorg) `$OPCred value!`nEXIT!"
                exit ;
            } ;

            # === Exchange LEMS/REMS detect & connect code

            $smsg= "Using EXOP cred:`$cred$($tenorg)OP:$((Get-Variable -name "cred$($tenorg)OP" ).value.username)" ;  
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
        } ; 
        
    } ;  # BEGIN-E
    PROCESS {
        #$ofile=".\$($ticket)-$($Mailbox)-folder-sizes-$(get-date -format 'yyyyMMdd-HHmmtt').csv" ;
        $error.clear() ;
    
        switch ($useEXOP){
            $false {

                write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):PERFORMING AN EXO MSGTRACE" ;
                if($VerbosePreference = "Continue"){
                    $VerbosePrefPrior = $VerbosePreference ;
                    $VerbosePreference = "SilentlyContinue" ;
                    $verbose = ($VerbosePreference -eq "Continue") ;
                } ; 
                disconnect-exo ; # pre-disconnect    
                $pltRXO = @{
                    Credential = (Get-Variable -name cred$($tenorg) ).value ;
                    verbose = $($verbose) ; }
                if ($script:useEXOv2) { reconnect-eXO2 @pltRXO }
                else { reconnect-EXO @pltRXO } ;
                if($VerbosePrefPrior -eq "Continue"){
                    $VerbosePreference = $VerbosePrefPrior ;
                    $verbose = ($VerbosePreference -eq "Continue") ;
                } ;

                # recycle $pltRXO for the AAD connection
                connect-AAD @pltRXO ;

                set-alias ps1GetMsgTrace Get-exoMessageTrace  ; 
                $props = $propsMsgEXO ; 
                $msgtrk=[ordered]@{
                    PageSize=1000 ;
                    Page=$null ;
                    StartDate=$null ;
                    EndDate=$null ;
                } ;
                if($Days -AND -not($StartDate -AND $EndDate)){
                    $msgtrk.StartDate=(get-date ([datetime]::Now)).adddays(-1*$days);
                    $msgtrk.EndDate=(get-date) ;
                } ;
                if($StartDate -and !($days)){
                    $msgtrk.StartDate=$(get-date $StartDate)
                } ;
                if($EndDate -and !($days)){
                    $msgtrk.EndDate=$(get-date $EndDate)
                } elseif($StartDate -and !($EndDate)){
                    write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):
    (StartDate w *NO* Enddate, asserting currenttime)" ;
                    $msgtrk.EndDate=(get-date) ;
                } ;
                
                $error.clear() ;
                TRY {
                    #Connect-AAD ;
                    $tendoms=Get-AzureADDomain ;
                } CATCH {
                    $ErrTrapt=$Error[0] ;
                    Write-Warning "$(get-date -format 'HH:mm:ss'): Failed processing $($_.Exception.ItemName). `nError Message: $($_.Exception.Message)`nError Details: $($_)" ;
                    $smsg = "FULL ERROR TRAPPED (EXPLICIT CATCH BLOCK WOULD LOOK LIKE): } catch[$($ErrTrpd.Exception.GetType().FullName)]{" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    Exit #Opts: STOP(debug)|EXIT(close)|CONTINUE(move on in loop cycle)|BREAK(exit loop iteration)|THROW $_/'CustomMsg'(end script with Err output)
                } ; 
            
                $Ten = ($tendoms |?{$_.name -like '*.mail.onmicrosoft.com'}).name.split('.')[0] ;
                $ofile ="$($ticket)-$($Ten)-$($Logon)-EXOMsgTrk" ;
                if($Sender){
                    if($Sender -match '\*'){
                        "(wild-card Sender detected)" ;
                        $msgtrk.add("SenderAddress",$Sender) ;
                    } else {
                        $msgtrk.add("SenderAddress",$Sender) ;
                    } ;
                    $ofile+=",From-$($Sender.replace("*","ANY"))" ;
                } ;
                if($Recipients){
                    if($Recipients -match '\*'){        "(wild-card Recipient detected)" ;
                        $msgtrk.add("RecipientAddress",$Recipients) ;
                    } else {
                            $msgtrk.add("RecipientAddress",$Recipients) ;
                    } ;
                    $ofile+=",To-$($Recipients.replace("*","ANY"))" ;
                } ;
                if($MessageId){
                    $msgtrk.add("MessageId",$MessageId) ;
                    $ofile+=",MsgId-$($MessageId.replace('<','').replace('>',''))" ;
                } ;
                if($MessageTraceId){
                    $msgtrk.add("MessageTraceId",$MessageTraceId) ;
                    $ofile+=",MsgId-$($MessageTraceId.replace('<','').replace('>',''))" ;
                } ;
                if($Subject){    $ofile+=",Subj-$($Subject.substring(0,[System.Math]::Min(10,$Subject.Length)))..." ;
                } ;
                if($Status){
                    $msgtrk.add("Status",$Status)  ;
                    $ofile+=",Status-$($Status)" ;
                } ;
                if($days){$ofile+= "-$($days)d-" } ;
                if($StartDate){$ofile+= "-$(get-date $StartDate -format 'yyyyMMdd-HHmmtt')-" } ;
                if($EndDate){$ofile+= "$(get-date $EndDate -format 'yyyyMMdd-HHmmtt')" } ;
                
                write-verbose -verbose:$true  "$((get-date).ToString('HH:mm:ss')):Running MsgTrk:$($Ten)" ;
    $(($msgtrk|out-string).trim()|out-default) ;
  
                TRY {
                    $Page = 1  ;
                    $Msgs=$null ;
                    do {
                        write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):Collecting - Page $($Page)..."  ;
                        $msgtrk.Page=$Page ;
                        $PageMsgs = ps1GetMsgTrace @msgtrk |  ?{$_.SenderAddress -notlike '*micro*' -or $_.SenderAddress -notlike '*root*' }  ;
                        $Page++  ;
                        $Msgs += @($PageMsgs)  ;
                    } until ($PageMsgs -eq $null) ;
                } CATCH {
                    Write-Warning "$(get-date -format 'HH:mm:ss'): Failed processing $($_.Exception.ItemName). `nError Message: $($_.Exception.Message)`nError Details: $($_)" ;
                    Exit ;
                } ; 
                $Msgs=$Msgs| Sort Received ;
                write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):==Msgs Returned:$(($Msgs|measure).count)" ;
                write-host -foregroundcolor white "$((get-date).ToString('HH:mm:ss')):Raw matches:$(($Msgs|measure).Count)" ;
                if($Subject){
                    write-host -foregroundcolor gray "$((get-date).ToString('HH:mm:ss')):Post-Filtering on Subject:$($Subject)" ;
                    $Msgs = $Msgs | ?{$_.Subject -like $Subject} ;
                    $ofile+="-Subj-$($Subject.replace("*"," ").replace("\"," "))" ;
                    write-host -foregroundcolor gray "$((get-date).ToString('HH:mm:ss')):Post Subj filter matches:$(($Msgs|measure).Count)" ;
                } ;
                $ofile+= "-run$(get-date -format 'yyyyMMdd-HHmmtt').csv" ;
                $ofile=[RegEx]::Replace($ofile, "[{0}]" -f ([RegEx]::Escape(-join [System.IO.Path]::GetInvalidFileNameChars())), '') ;
                $ofile=".\logs\$($ofile)" ;
                if($Msgs){
                    $Msgs | select $props | export-csv -notype -path $ofile  ;
                    write-host -foregroundcolor yellow "Status Distrib:" ;
                    write-host -foregroundcolor yellow "$((get-date).ToString('HH:mm:ss')):`n#*------v MOST RECENT MATCH v------" ;
                    write-host -foregroundcolor white "$(($msgs[-1]| format-list ReceivedLocal,StatusSenderAddress,RecipientAddress,Subject|out-string).trim())";
                    write-host -foregroundcolor yellow "$((get-date).ToString('HH:mm:ss')):`n#*------^ MOST RECENT MATCH ^------" ;
                    write-host -foregroundcolor yellow "$((get-date).ToString('HH:mm:ss')):`n#*------v Status DISTRIB v------" ;
                    "$(($Msgs | select -expand Status | group | sort count,count -desc | select count,name |out-string).trim())";
                    write-host -foregroundcolor yellow "$((get-date).ToString('HH:mm:ss')):`n#*------^ Status DISTRIB ^------" ;
                    if(test-path -path $ofile){
                            write-host -foregroundcolor white "$((get-date).ToString('HH:mm:ss')):(log file confirmed)" ;
                            Resolve-Path -Path $ofile | select -expand Path | out-clipboard ;
                            write-host -foregroundcolor yellow "$((get-date).ToString('HH:mm:ss')):$($Msgs.count) matches output to:`n'$($ofile)'`n(copied to CB)" ;
                    } else { "MISSING LOG FILE!" } ;

                    if($ReportFail){
                        $sBnr3="`n#*------v Status:FAIL Traffic (up to 1st $($ReportRowsLimit)) v------" ; 
                        write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($sBnr3)" ;
                        write-host -foregroundcolor cyan "$(($MSGS|?{$_.Status -eq 'FAIL'} | select -first $($ReportRowsLimit) | fl recipients,recipientstatus,ServerHostname|out-string).trim())" ; 
                        write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($sBnr3.replace('~v','~^').replace('v~','^~'))`n" ;
                    } ; 
                    
                    if($asObject){
                        $Msgs | write-output ; 
                    } ; 
                } else {
                    write-host -foregroundcolor yellow "$((get-date).ToString('HH:mm:ss')):NO MATCHES FOUND from::`n$(($msgtrk|out-string).trim()|out-default)`n(with any relevant ConnectorId postfilter)" ;
                } ;
            } ; # end EXO switchblock

            $true {
                write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):PERFORMING AN ONPREM MSGTRACK" ;
                if($VerbosePreference = "Continue"){
                    $VerbosePrefPrior = $VerbosePreference ;
                    $VerbosePreference = "SilentlyContinue" ;
                    $verbose = ($VerbosePreference -eq "Continue") ;
                } ; 
                # connect OP
                $pltRX10 = @{
                    Credential = (Get-Variable -name "cred$($tenorg)OP" ).value ;
                    verbose = $($verbose) ; } ;     
                if($pltRX10){
                    Connect-Ex2010 @pltRX10 ;
                } else { connect-Ex2010 ; } ;

                # reenable VerbosePreference:Continue, if set, during mod loads 
                if($VerbosePrefPrior -eq "Continue"){
                    $VerbosePreference = $VerbosePrefPrior ;
                    $verbose = ($VerbosePreference -eq "Continue") ;
                } ;

                set-alias ps1GetMsgTrace get-messagetrackinglog  ; 
                $props = $propsMsgEx10 ; 
                $msgtrk=@{
                    Start=(get-date ([datetime]::Now)).adddays(-1*$days) ;
                    End=(get-date) ;
                    resultsize="UNLIMITED" ;
                } ;
                # Page=$null ;
                $msgtrk=[ordered]@{
                    resultsize="UNLIMITED" ;
                    Start=$null ;
                    End=$null ;
                } ;
                if($Days -AND -not($StartDate -AND $EndDate)){
                    $msgtrk.Start=(get-date ([datetime]::Now)).adddays(-1*$days);
                    $msgtrk.End=(get-date) ;
                } ;
                if($StartDate -and !($days)){
                    $msgtrk.Start=$(get-date $StartDate)
                } ;
                if($EndDate -and !($days)){
                    $msgtrk.End=$(get-date $EndDate)
                } elseif($StartDate -and !($EndDate)){
                    write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):
    (StartDate w *NO* End, asserting currenttime)" ;
                    $msgtrk.End=(get-date) ;
                } ;
                TRY {
                    $Site=[System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::GetComputerSite().Name ;
                    # "$($ticket)-$($uid)-$($Site.substring(0,3))-MsgTrk" ;
                    $ofile ="$($ticket)-$($Site.substring(0,3))-OPMsgTrk" ;
                    if($Sender){$msgtrk.add("Sender",$Sender) ;
                        $ofile+=",From-$($Sender)" ;
                        } ;
                    if($Recipients){$msgtrk.add("Recipients",$Recipients) ;
                        $ofile+=",To-$($Recipients)" ;
                    } ;
                    if($Subject){$msgtrk.add("MessageSubject",$Subject)  ;
                        $ofile+=",Subj-$($Subject.substring(0,[System.Math]::Min(10,$Subject.Length)))..." ;
                    } ;
                    if($EventID){$msgtrk.add("EventID",$Status)  ;
                        $ofile+=",Evt-$($Status)" ;
                    } ;
                    
                    write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$((get-alias ps1GetMsgTrace).ResolvedCommandName) w`n$(($msgtrk|out-string).trim())" ; 
                    $Srvrs=(Get-ExchangeServer | where { $_.isHubTransportServer -eq $true -and $_.Site -match ".*\/$($Site)$"} | select -expand Name) ;
                    #$Msgs=($Srvrs| get-messagetrackinglog @msgtrk) | sort Timestamp ;
                    $Msgs =@() ; # 
                    # loop the servers, to provide a status output
                    foreach($Srvr in $Srvrs){
                        write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):Tracking $($Srvr) server..." ; 
                        $sMsgs = ($Srvr| get-messagetrackinglog @msgtrk) ;
                        write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):(($Srvr):$(($sMsgs|measure).count) matched msgs)" ; 
                        $Msgs+=$sMsgs ; 
                        $sMsgs = $null ; 
                    } ; 
                    #$Msgs = $Msgs |  sort Timestamp ;
                    $Msgs=$Msgs| Sort Timestamp ;
                    write-verbose -verbose:$true  "$((get-date).ToString('HH:mm:ss')):Raw matches:$(($Msgs|measure).Count)" ;
                    if($Connectorid){
                        write-verbose -verbose:$true  "$((get-date).ToString('HH:mm:ss')):Filtering on Conn:$($Connectorid)" ;
                        $Msgs = $Msgs | ?{$_.connectorid -like $Connectorid} ;
                        $ofile+="-conn-$($Connectorid.replace("*"," ").replace("\"," "))" ;
                        write-verbose -verbose:$true  "$((get-date).ToString('HH:mm:ss')):Post Conn filter matches:$(($Msgs|measure).Count)" ;
                    } ;
                    if($Source){
                        write-verbose -verbose:$true  "$((get-date).ToString('HH:mm:ss')):Filtering on Source:$($Source)" ;
                        $Msgs = $Msgs | ?{$_.Source -like $Source} ;
                        write-verbose -verbose:$true  "$((get-date).ToString('HH:mm:ss')):Post Src filter matches:$(($Msgs|measure).Count)" ;
                        $ofile+="-src-$($Source)" ;
                    } ;
                    if($Days){$ofile+= "-$($days)d-run$(get-date -format 'yyyyMMdd-HHmmtt').csv" ;} 
                    else {
                        $ofile+= "-$(get-date $msgtrk.Start -format 'yyyyMMdd-HHmmtt')-$(get-date $msgtrk.End -format 'yyyyMMdd-HHmmtt')-run$(get-date -format 'yyyyMMdd-HHmmtt').csv" ;
                    } ;  
                    $ofile=[RegEx]::Replace($ofile, "[{0}]" -f ([RegEx]::Escape(-join [System.IO.Path]::GetInvalidFileNameChars())), '') ;
                    $ofile=".\logs\$($ofile)" ;
                    
                    if($Msgs){
                        $Msgs | SELECT $props| EXPORT-CSV -notype -path $ofile ;
                        write-host -foregroundcolor white "$((get-date).ToString('HH:mm:ss')):`n#*------v MOST RECENT MATCH v------" ;
                        write-host -foregroundcolor cyan "$(((($msgs[-1]| format-list Timestamp,EventId,Sender,Recipients,MessageSubject|out-string).trim())|out-string).trim())" ; 
                        write-host -foregroundcolor white "$((get-date).ToString('HH:mm:ss')):`n#*------^ MOST RECENT MATCH ^------" ;
                        write-host -foregroundcolor white "$((get-date).ToString('HH:mm:ss')):`n#*------v EVENTID DISTRIB v------" ;
                        write-host -foregroundcolor cyan "$(($Msgs | select -expand EventId | group | sort count,count -desc | select count,name |out-string).trim())" ; 
                        write-host -fore gray "(SEND=SMTP SEND,TRANSFER=Routing,RESOLVE=Recipient conversion)" ;
                        write-host -foregroundcolor white "$((get-date).ToString('HH:mm:ss')):`n#*------^ EVENTID DISTRIB ^------" ;
                        if(test-path -path $ofile){
                            write-host -foregroundcolor white "$((get-date).ToString('HH:mm:ss')):(log file confirmed)" ;
                            Resolve-Path -Path $ofile | select -expand Path | out-clipboard ;
                            write-host -foregroundcolor white "$((get-date).ToString('HH:mm:ss')):$($Msgs.count) matches output to:`n'$($ofile)'`n(copied to CB)" ;
                        } else { "MISSING LOG FILE!" } ;
                        
                        if($ReportFail){
                            $sBnr3="`n#*~~~~~~v -ReportFail specified: Status:FAIL Traffic (up to 1st $($ReportRowsLimit)): v~~~~~~" ; 
                            write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($sBnr3)" ;
                            write-host -foregroundcolor cyan "$(((($MSGS|?{$_.eventid -eq 'fail'} | select -first $($ReportRowsLimit) | fl recipients,recipientstatus,ServerHostname|out-string).trim())|out-string).trim())" ; 
                            write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($sBnr3.replace('~v','~^').replace('v~','^~'))`n" ;
                        } ; 

                        if($asObject){
                            $Msgs | SELECT $props | write-output ; 
                        } ; 
                    } else {    write-host -foregroundcolor white "$((get-date).ToString('HH:mm:ss')):NO MATCHES FOUND from::`n$(($msgtrk|out-string).trim()|out-default)`n(with any relevant ConnectorId postfilter)" ;
                    } ;
                } CATCH {
                    Write-Warning "$(get-date -format 'HH:mm:ss'): Failed processing $($_.Exception.ItemName). `nError Message: $($_.Exception.Message)`nError Details: $($_)" ;
                    Exit ;
                } ; 
            } ;
            default {
                throw "UNRECOGNIZED useEXOP value)" ; exit ; 
            } ; 
        } ; # SWITCH-E
        
    } ;  # PROC-E
    END {
        remove-alias ps1GetMsgTrace ;
    } ; 
} 
#*------^ END Function get-MsgTrace() ^------
