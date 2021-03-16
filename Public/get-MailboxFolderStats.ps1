#*------v get-MailboxFolderStats.ps1 v------
function get-MailboxFolderStats {
    <#
    .SYNOPSIS
    get-MailboxFolderStats.ps1 - Perform smart get-mailboxfolderstatistics command, as appropriate to target location -Mailbox, on either Exchange on-premesis or Exchange Online.
    .NOTES
    Version     : 1.0.0.
    Author      : Todd Kadrie
    Website     : http://www.toddomation.com
    Twitter     : @tostka / http://twitter.com/tostka
    CreatedDate : 2021-03-12
    FileName    : get-MailboxFolderStats
    License     : MIT License
    Copyright   : (c) 2020 Todd Kadrie
    Github      : https://github.com/tostka/verb-exo
    Tags        : Powershell,ExchangeOnline,Mailbox,Statistics,Reporting
    REVISIONS
    * 3:28 PM 3/16/2021 added multi-tenant support
    * 1:12 PM 3/15/2021 init work was done 3/12, removed recursive-err generating #Require on the hosting verb-EXO module
    .DESCRIPTION
    get-MailboxFolderStats.ps1 - Perform smart get-mailboxfolderstatistics command, as appropriate to target location -Mailbox, on either Exchange on-premesis or Exchange Online.
    Dependancy on my verb-ex2010 Exchange onprem (and is within verb-exo EXO mod, which adds dependant EXO connection support).
    .PARAMETER TenOrg
    TenantTag values, indicating Tenant to Create DDG WITHIN[-TenOrg 'TOL']    
    .PARAMETER  Mailbox
    Mailbox identifier [samaccountname,name,emailaddr,alias]
    .PARAMETER Ticket
    Ticket # [-Ticket nnnnn]
    .PARAMETER IncludeAge
    Switch to include Oldest/Newest message per folder information[-IncludeAge]
    .PARAMETER IncludeSize
    Switch to include aggregate size of each folder [-IncludeSize]
    .PARAMETER NonEmptyOnly
    Switch to display infor for only non-zero content folders (defaults `$true)[-NonEmptyOnly]
    .INPUTS
    Accepts piped input.
    .OUTPUTS
    Outputs csv & console summary of mailbox folders content
    .EXAMPLE
    get-MailboxFolderStats -Mailbox quotes@domain.com -Ticket 99999 -includeage -verbose ;
    Perform a mailbox stats summary report query, on the specified mailbox, and include specified ticket# in output csv (which is output below .\logs\ dir of current directory at runtime).
    .EXAMPLE
    $report = get-MailboxFolderStats -Mailbox quotes@domain.com -Ticket 99999 -includeage -asobject ;
    Return an object for the summary report, rather than console dump (in addition to csv export)
    .EXAMPLE
    get-MailboxFolderStats -Mailbox quotes@domain.com -Ticket 347298 -includeage -includesize ;
    Perform a mailbox stats, and include size per folder (in KB) in output report
    .LINK
    https://github.com/tostka/verb-exo
    .LINK
    https://github.com/tostka/verb-ex2010
    #>
    #Requires -Version 3
    #Requires -Modules verb-ex2010
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$FALSE,HelpMessage="TenantTag values, indicating Tenant to Create DDG WITHIN[-TenOrg 'TOL']")]
        [ValidateNotNullOrEmpty()]
        [string]$TenOrg = ('TOR'),
        [Parameter(Position=0,Mandatory=$True,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,HelpMessage="Mailbox identifier [samaccountname,name,emailaddr,alias]")]
        [ValidateNotNullOrEmpty()][string]$Mailbox,    
        [Parameter(Mandatory=$false,HelpMessage="Ticket # [-Ticket nnnnn]")]
        #[ValidateLength(5)] # non-mandatory
        [int]$Ticket,
        [Parameter(HelpMessage="Switch to include Oldest/Newest message per folder information[-IncludeAge]")]
        [switch] $IncludeAge,
        [Parameter(HelpMessage="Switch to include aggregate size of each folder [-IncludeSize]")]
        [switch] $IncludeSize,
        [Parameter(HelpMessage="Switch to display info for only non-zero content folders (defaults `$true)[-NonEmptyOnly]")]
        [switch] $NonEmptyOnly=$true,
        [Parameter(HelpMessage="Switch to return raw object rather than formated console report(defaults `$true)[-NonEmptyOnly]")]
        [switch] $asObject
    ) ;
    BEGIN {
        $Verbose=($VerbosePreference -eq 'Continue') ;  
        # Get parameters this function was invoked with
        $PSParameters = New-Object -TypeName PSObject -Property $PSBoundParameters ;  
        $pltGMFS=@{identity= $Mailbox ;} ; 
        $propsFldr = @{Name='Folder';Expression={$_.Identity.tostring()}},@{Name="Items";Expression={$_.ItemsInFolder}} ;
        $rgxSysFldrs = '.*\\(Versions|SubstrateHolds|DiscoveryHolds|Yammer.*|Social\sActivity\sNotifications|Suggested\sContacts|Recipient\sCache|PersonMetadata|Audits|Calendar\sLogging|Purges)$' ; 
        if($IncludeAge){ 
            $pltGMFS.add('IncludeOldestAndNewestItems',$true) ; 
            $propsFldr += @{Name="OldestItem";Expression={get-date $_.OldestItemReceivedDate}},@{Name="NewestItem";Expression={$_.NewestItemReceivedDate}} ; 
        } ;
        if($IncludeSize){ 
            $pltGMFS.add('IncludeAnalysis',$true) ; 
            # w dehydrated, raw parsing is: $mbxstats.TotalItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1MB ;
            $propsFldr += @{Name="SizeMB";Expression={[math]::round($_.FolderSize.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1MB,2)}} ; 
        } ;

        $Retries = 4 ;
        $RetrySleep = 5 ;
        if(!$ThrottleMs){$ThrottleMs = 50 ;}
        $CredRole = 'CSVC' ; # role of svc to be dyn pulled from metaXXX if no -Credential spec'd, 
        if(!$rgxDriveBanChars){$rgxDriveBanChars = '[;~/\\\.:]' ; } ; # ;~/\.:, 

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
        $ofile=".\$($ticket)-$($Mailbox)-folder-sizes-$(get-date -format 'yyyyMMdd-HHmmtt').csv" ;
        $error.clear() ;
        TRY {
            if(!(gcm get-recipient -ea 0)){rx10} ;
            $OpRcp=get-recipient $Mailbox ;
            switch ($OpRcp.recipienttype){
                "MailUser" {
                    write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($tmbx) IS AN EXO MBOX" ;
                    
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

                    set-alias ps1GetMbxFldrStat Get-exoMailboxFolderStatistics ; 
                } ;
                "UserMailbox" {
                    write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($tmbx) IS AN EX2010 MBOX" ;
                    
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

                    set-alias ps1GetMbxFldrStat Get-MailboxFolderStatistics ; 
                } ;
                default {
                    throw "UNRECOGNIZED ONPREM RECIPIENTTYPE:$($OpRcp.recipienttype)" ; exit ; 
                } ; 
            } ;
            write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$((get-alias ps1GetMbxFldrStat).definition) w`n$(($pltGMFS|out-string).trim())" ; 
            $fldrs = ps1GetMbxFldrStat @pltGMFS ;
            if($NonEmptyOnly){
                write-host -foregroundcolor white "$((get-date).ToString('HH:mm:ss')):(REPORTING NON-ZERO FOLDERS ONLY)" ; $fldrs = $fldrs | ?{$_.ItemsInFolder -gt 0}
            } ; 
            $fldrs | ?{$_.identity -notmatch $rgxSysFldrs } | select $propsFldr | export-csv  -path $ofile -notype ;
            if(!$asObject){
                import-csv $ofile | ft -auto | out-default ; 
            } else { 
                write-verbose "-asObject specified, returning object to pipeline (rather than console dump)" ; 
                import-csv $ofile | write-output ; 
            } ; 
            write-host -foregroundcolor white "$((get-date).ToString('HH:mm:ss')):`n===>`$ofile:$($ofile)`n" ;
        } CATCH {
            Write-Warning "$(get-date -format 'HH:mm:ss'): Failed processing $($_.Exception.ItemName). `nError Message: $($_.Exception.Message)`nError Details: $($_)" ;
            Exit ;
        } ; 
    } ;  # PROC-E
    END {
        remove-alias ps1GetMbxFldrStat ;
    } ; 
    
}

#*------^ get-MailboxFolderStats.ps1 ^------