#*----------v Function get-MailboxFolderStats() v----------
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
    * 1:12 PM 3/15/2021 init work was done 3/12, removed recursive-err generating #Require on the hosting verb-EXO module
    .DESCRIPTION
    get-MailboxFolderStats.ps1 - Perform smart get-mailboxfolderstatistics command, as appropriate to target location -Mailbox, on either Exchange on-premesis or Exchange Online.
    Dependancy on my verb-ex2010 Exchange onprem (and is within verb-exo EXO mod, which adds dependant EXO connection support).
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
    $Verbose=($VerbosePreference -eq 'Continue') ;  
    $pltGMFS=@{identity= $Mailbox ;} ; 
    $propsFldr = @{Name='Folder';Expression={$_.Identity.tostring()}},@{Name="Items";Expression={$_.ItemsInFolder}} ;
    $rgxSysFldrs = '.*\\(Versions|SubstrateHolds|DiscoveryHolds|Yammer.*|Social\sActivity\sNotifications|Suggested\sContacts|Recipient\sCache|PersonMetadata|Audits|Calendar\sLogging|Purges)$' ; 
    if($IncludeAge){ 
        $pltGMFS.add('IncludeOldestAndNewestItems',$true) ; 
        #$propsFldr += @{Name="OldestItem";Expression={get-date $_.OldestDeletedItemReceivedDate -f "yyyyMMdd"}},@{Name="NewestItem";Expression={$_.NewestItemReceivedDate -f "yyyyMMdd"}} ; 
        $propsFldr += @{Name="OldestItem";Expression={get-date $_.OldestItemReceivedDate}},@{Name="NewestItem";Expression={$_.NewestItemReceivedDate}} ; 
    } ;
    # # -IncludeAnalysis
    if($IncludeSize){ 
        $pltGMFS.add('IncludeAnalysis',$true) ; 
        #$propsFldr += @{Name="SizeMB";Expression={$_.FolderSize.ToMB()}} ; 
        # w dehydrated, raw parsing is: $mbxstats.TotalItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1MB ;
        $propsFldr += @{Name="SizeMB";Expression={[math]::round($_.FolderSize.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1MB,2)}} ; 
    } ;
    $ofile=".\$($ticket)-$($Mailbox)-folder-sizes-$(get-date -format 'yyyyMMdd-HHmmtt').csv" ;
    $error.clear() ;
    TRY {
        if(!(gcm get-recipient -ea 0)){rx10} ;
        $OpRcp=get-recipient $Mailbox ;
        switch ($OpRcp.recipienttype){
            "MailUser" {
                write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($tmbx) IS AN EXO MBOX" ;
                reconnect-exo ;
                set-alias ps1GetMbxFldrStat Get-exoMailboxFolderStatistics ; 
            } ;
            "UserMailbox" {
                write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($tmbx) IS AN EX2010 MBOX" ;
                reconnect-ex2010 ;
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
    remove-alias ps1GetMbxFldrStat ;
} #*------^ END Function get-MailboxFolderStats() ^------