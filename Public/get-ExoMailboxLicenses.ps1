# get-ExoMailboxLicenses.ps1

#*------v get-ExoMailboxLicenses.ps1 v------
function get-ExoMailboxLicenses {
<#
    .SYNOPSIS
    get-ExoMailboxLicenses - Provides a prefab indexed hash of Exchange-Online mailbox-supporting licenses (at least one of which is required to accomodate an EXO Usermailbox - This now dynamically calls Get-AzureADSubscribedSku and postfilters the ServicePlan list for matches on the $ServicePlanName array (which reflects Exchange mailbox ServicePlanNames). The ServicePlanName array must be manually updated to accomodate MS licensure changes over time).
    .PARAMETER Mailboxes
    .NOTES
    Version     : 1.0.0
    Author      : Todd Kadrie
    Website     :	http://www.toddomation.com
    Twitter     :	@tostka / http://twitter.com/tostka
    CreatedDate : 2022-02-25
    FileName    : get-ExoMailboxLicenses.ps1
    License     : MIT License
    Copyright   : (c) 2022 Todd Kadrie
    Github      : https://github.com/tostka/verb-ex2010
    Tags        : Powershell
    REVISIONS
    * 1:22 PM 6/18/2024 updated SERVICE_CONNECTIONS block; reflects latest variant; 
    * 4:16 PM 6/17/2024 add: ServicePlanName to the detailed output ; pulled transcript END block (unused, would kill other process logging) ; adding -OutDetail, need to implement to return avail etc full details in outobject
    * 5:12 PM 6/13/2024 update to make dynamic, querying for plans serviceplans.serviceplanname -match  EXCHANGE_S_ENTERPRISE','EXCHANGE_S_STANDARD','EXCHANGE_S_DESKLESS (-ne MCOCAP 	Common Area Phone)
    * 12:45 PM 6/21/2022 added cbh expl that rolls up a rgx to use for independant manual tests against 
    * 2:21 PM 3/1/2022 updated CBH
    * 4:27 PM 2/25/2022 init vers
    .DESCRIPTION
    get-ExoMailboxLicenses - Provides a prefab array indexed hash of Exchange-Online mailbox-supporting licenses (at least one of which is required to accomodate an EXO Usermailbox - This now dynamically calls Get-AzureADSubscribedSku and postfilters the ServicePlan list for matches on the $ServicePlanName array (which reflects Exchange mailbox ServicePlanNames). The ServicePlanName array must be manually updated to accomodate MS licensure changes over time).

    This feeds my test-EXOIsLicensed(), and is usable for feeding add-EXOLicense & remove-EXOLicense
    Underlying goal is to *dynamically* pursue the current supported licensure, but calling Get-AzureADSubscribedSku and postfiltering the ServicePlan list for matches on the $ServicePlanName array 
    (which reflects Exchange mailbox ServicePlanNames). The ServicePlanName array must be manually updated to accomodate MS licensure changes over time).

    ## Tracking changes in Microsoft Licensing Service Plan Names over time. 
    
    At the current time, there is a _five year out of date_ json here: 
    [Compare Microsoft Exchange Online Plans Microsoft 365](https://www.microsoft.com/en-us/microsoft-365/exchange/compare-microsoft-exchange-online-plans)

    ... which they claim is no longer needed because of online non-code-ingestable/non-filterable giant listing
    Or you can also use the GraphAPI and chase that ball of ugly over time. 
    So for now, I'm basing it off of the out of date json's plan names, and leaving gapi for a later more-freetimey time.

    ## MS Exchange plan compairson's posted (as of 6/17/2024)

    ### Exchange Online (Plan 1) $4.00 user/month
    - 50g mbx, OWA, inplace archive
    

    ### Exchange Online (Plan 2)  $8.00  user/month
    - above/owa ++100g mbx, DLP, Vmail

    ### Microsoft 365 Business Standard $12.50 user/month
    - ++Outlook
    - reflects use of the EXCHANGE_S_ENTERPRISE ServicePlanName

    ## Extracting ServicePlans broadly

    [Azure-AD-Licensing-DB/ProductLicensesDb.json at master · jpawlowski/Azure-AD-Licensing-DB · GitHub](https://github.com/jpawlowski/Azure-AD-Licensing-DB/blob/master/ProductLicensesDb.json) 
    has a json db of all plan details

    1. click the download link on the page, dl to file
    2. import & convert it
    ```powershell
    $licdb = gc C:\sc\powershell\EXOScripts\o365-ProductLicensesDb.json | ConvertFrom-Json ; 
    ```
    3. Filter for targets, exclude 'EXCHANGE_ANALYTICS, EXCHANGE_S_ARCHIVE, EXCHANGE_S_ARCHIVE_ADDON non-mailbox-granting
    ```powershell
    $licdb.items | ?{$_.ServicePlans -match 'Exchange'} | select -expand ServicePlans | ?{$_ -match 'EXCHANGE_' -AND $_-notmatch '_(ANALYTICS|ARCHIVE|FOUNDATION)'} | SELECT -UNIQUE  | sort ;
    EXCHANGE_B_STANDARD
    EXCHANGE_L_STANDARD
    EXCHANGE_S_DESKLESS
    EXCHANGE_S_ENTERPRISE
    EXCHANGE_S_ESSENTIALS
    EXCHANGE_S_STANDARD
    EXCHANGE_S_STANDARD_MIDMARKET

    ```
    4. So we can take any give Get-AzureADSubscribedSku and filter for licenses which include the above, to pick out suitable Exchange-Mailbox-supporting licenses. This filtered list goes in the default $ServicePlanName list


    ## Discussion of the topic - dynamically finding onboing licenses as they mess with the licenses

    [Service plans that indicate an exchange license? - Microsoft Q&A](https://learn.microsoft.com/en-us/answers/questions/967768/service-plans-that-indicate-an-exchange-license)

        #-=-=-=-=-=-=-=-=
        Service plans that indicate an exchange license?
        isaac parsons 6 Reputation points
        Aug 15, 2022, 2:34 PM
        From what I gather EXCHANGE_S_STANDARD, EXCHANGE_B_STANDARD, EXCHANGE_L_STANDARD, EXCHANGE_S_ENTERPRISE, EXCHANGE_S_STANDARD_GOV, EXCHANGE_S_ENTERPRISE_GOV, EXCHANGE_S_STANDARD_MIDMARKET are service plans that indicate an exchange license, are there any others that I'm missing?
        Microsoft Exchange Online Management 
        ---
        Dillon Silzer 54,926 Reputation points
        Aug 16, 2022, 9:58 PM
        Hey @isaac parsons
        If you navigate to https://learn.microsoft.com/en-us/azure/active-directory/enterprise-users/licensing-service-plan-reference you can scroll down to the Exchange plans under Enterprise Mobility + Security G5 GCC:
        You can also use CTRL+F and type the work EXCHANGE and search the page for all the plans that include it.
        You can also download the CSV version here.

        #-=-=-=-=-=-=-=-=

    ## [Mailbox plans in Exchange Online | Microsoft Learn](https://learn.microsoft.com/en-us/exchange/recipients-in-exchange-online/manage-user-mailboxes/mailbox-plans)

    ### Mailbox plans in Exchange Online

        Article
        02/21/2023
        The following table describes the mailbox plans that you're likely to see in Exchange Online.
        Subscription or license 	Mailbox plan display name
        Exchange Online Kiosk
        Microsoft 365 or Office 365 Enterprise F3
	        ExchangeOnlineDeskless
        Microsoft 365 Business Basic
        Microsoft 365 or Office 365 Enterprise E1
        Exchange Online Plan 1
	        ExchangeOnline
        Microsoft 365 or Office 365 Enterprise E3
        Microsoft 365 or Office 365 Enterprise E5
        Exchange Online Plan 2
    

    .PARAMETER ServicePlanName
    ServicePlanName values that identify Exchange-mailbox supporting licenses (defaults to EXCHANGE_S_DESKLESS|EXCHANGE_S_STANDARD|EXCHANGE_S_ENTERPRISE)[-ServicePlanName 'EXCHANGE_S_DESKLESS']
    .PARAMETER rgxbannedSPN
    Regular Expression for ServicePlanName values should be excluded from results (Common Area Phone license etc)[-rgxBannedSPN 'EXCHANGE_S_DESKLESS']
    .PARAMETER OutDetail
    Switch to enable expanded non-legacy return (legacy returned SKU|Label|Notes; -outdetail adds Enabled|Consumed|Available|Warning|Suspended specs)[-OutDetail]
    .PARAMETER Unfiltered
    Switch to suppress normal post-filter Availability checks (Available -gt 0 & Enabled -gt 0) and return *any* matched supporting licenses in the Tenant [-unfiltered]
    .PARAMETER TenOrg
    Tenant Tag (3-letter abbrebiation)[-TenOrg 'XYZ']
    .PARAMETER Credential
    Use specific Credentials (defaults to Tenant-defined SvcAccount)[-Credentials [credential object]]
    .PARAMETER UserRole
    Credential User Role spec (SID|CSID|UID|B2BI|CSVC|ESVC|LSVC|ESvcCBA|CSvcCBA|SIDCBA)[-UserRole @('SIDCBA','SID','CSVC')]
    .PARAMETER useEXOv2
    Use EXOv2 (ExchangeOnlineManagement) over basic auth legacy connection [-useEXOv2]
    .PARAMETER Silent
    Switch to specify suppression of all but warn/error echos.(unimplemented, here for cross-compat)
    .EXAMPLE
    PS> $hQuotas = get-ExoMailboxLicenses -verbose ; 
    PS> $hQuotas['database2']
    Name           ProhibitSendReceiveQuotaGB ProhibitSendQuotaGB IssueWarningQuotaGB
    ----           -------------------------- ------------------- -------------------
    database2      12.000                     10.000              9.000
    Retrieve local org on-prem MailboxDatabase quotas and assign to a variable, with verbose outputs. Then output the retrieved quotas from the indexed hash returned, for the mailboxdatabase named 'database2'.
    .EXAMPLE
    PS>  $pltGXML=[ordered]@{
    PS>      #TenOrg= $TenOrg;
    PS>      verbose=$($VerbosePreference -eq "Continue") ;
    PS>      #credential= $pltRXO.credential ;
    PS>      #(Get-Variable -name cred$($tenorg) ).value ;
    PS>  } ;
    PS>  $smsg = "$($tenorg):get-ExoMailboxLicenses w`n$(($pltGXML|out-string).trim())" ;
    PS>  if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
    PS>  else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
    PS>  $objRet = $null ;
    PS>  $objRet = get-ExoMailboxLicenses @pltGXML ;
    PS>  if( ($objRet|Measure-Object).count -AND $objRet.GetType().FullName -match $rgxHashTableTypeName ){
    PS>      $smsg = "get-ExoMailboxLicenses:$($tenorg):returned populated ExMbxLicenses" ;
    PS>      if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
    PS>      else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
    PS>      $ExMbxLicenses = $objRet ;
    PS>  } else {
    PS>      $smsg = "get-ExoMailboxLicenses:$($tenorg):FAILED TO RETURN populated [hashtable] ExMbxLicenses" ;
    PS>      if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Error } 
    PS>      else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
    PS>      THROW $SMSG ; 
    PS>      break ; 
    PS>  } ;
    PS>  $smsg = "$(($ExMbxLicenses.Values|measure).count) EXO UserMailbox-supporting License summaries returned)" ;
    PS>  if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
    PS>  else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;    
    PS>  $smsg = "$(($ExMbxLicenses.Values|measure).count) EXO UserMailbox-supporting License summaries returned)" ;
    PS>  if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
    PS>  else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
    PS> $aadu = get-azureaduser -obj someuser@domain.com ; 
    PS> $IsExoLicensed = $false ;
    PS> foreach($pLic in $aadu.AssignedLicenses){
    PS>     $smsg = "--(LicSku:$($plic): checking EXO UserMailboxSupport)" ; 
    PS>     if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
    PS>     else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;                                     
    PS>     if($ExMbxLicenses[$plic]){
    PS>         $hSummary.IsExoLicensed = $true ;
    PS>         $smsg = "$($mbx.userprincipalname) HAS EXO UserMailbox-supporting License:$($ExMbxLicenses[$sku].SKU)|$($ExMbxLicenses[$sku].Label)" ; 
    PS>         if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
    PS>         else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
    PS> } ; 
    PS> if(-not $hSummary.IsExoLicensed){
    PS>     $smsg = "$($mbx.userprincipalname) WAS FOUND TO HAVE *NO* EXO UserMailbox-supporting License!" ; 
    PS>     if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug 
    PS>     else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
    PS> } ;
    Expanded example with testing of returned object, and demoes use of the returned hash against a mailbox spec, steering via .UseDatabaseQuotaDefaults
    .EXAMPLE
    PS> $pltGXML=[ordered]@{
    PS>    #TenOrg= $TenOrg;
    PS>    verbose=$($VerbosePreference -eq "Continue") ;
    PS>    #credential= $pltRXO.credential ;
    PS>    #(Get-Variable -name cred$($tenorg) ).value ;
    PS>    Unfiltered = $true ; 
    PS>    OutDetail = $true ; 
    PS> } ;
    PS> $smsg = "$($tenorg):get-ExoMailboxLicenses w`n$(($pltGXML|out-string).trim())" ;
    PS> if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
    PS> else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
    PS> $objRet = $null ;
    PS> $objRet = get-ExoMailboxLicenses @pltGXML ;
    Demo that shows retrieving unfiltered full detail - as this now returns filtered usable/assignable licenses by default.
    . EXAMPLE
    PS> $ExMbxLicenses = get-ExoMailboxLicenses ;
    PS> [regex]$rgxExLics = ('(' + (($ExMbxLicenses.GetEnumerator().name |%{[regex]::escape($_)}) -join '|') + ')') ; 
    Demo pulling the underlying licenses list and building a regex for static use
    PS> $licOrdered = @() ; 
    PS> 'EXCHANGE_S_DESKLESS','EXCHANGE_S_STANDARD','EXCHANGE_S_ENTERPRISE' | %{
    PS>     $SPN = $_ ; 
    PS>     $licOrdered += $ExMbxLicenses.values | ?{$_.ServicePlanName -eq $SPN } | sort Enabled,Available -Descending; 
    PS> } ; 
    PS> $TenDom = (gv -name "$($TenOrg)Meta").value['o365_TenantDom'].tolower() ; 
    PS> $LicenseSkuIds = $licOrdered.sku  | %{"$($TenDom):$($_)"} ;
    Demo pushing the licenses into application preference order from cheapest to most $$ class, and sorted subs on Enabled & Available (lifted from Add-EXOLicense()).
    .EXAMPLE
    PS> TRY{
    PS>     $url = 'https://github.com/jpawlowski/Azure-AD-Licensing-DB/blob/master/ProductLicensesDb.json' ; 
    PS>     $tfile = join-path 'C:\usr\work\o365\scripts\' ($url.split('/')[-1]) ; 
    PS>     write-host "DL source json db of licenses to:`n$($tfile)" ; 
    PS>     Invoke-WebRequest -Uri $url -OutFile $tfile ; 
    PS>     write-verbose "import the json into a vari" ; 
    PS>     $licdb = gc $tfile | ConvertFrom-Json ; 
    PS>     write-host "filter licenses with ServicePlans named with Exchange, filter & exclude non-mailbox variants, select unique, and output a sorted list, for use in the `$ServicePlanName array" ; 
    PS>     $licdb.items | ?{$_.ServicePlans -match 'Exchange'} | select -expand ServicePlans | ?{$_ -match 'EXCHANGE_' -AND $_-notmatch '_(ANALYTICS|ARCHIVE|FOUNDATION)'} | SELECT -UNIQUE  | sort ;
    PS> } CATCH {
    PS>     $ErrTrapd=$Error[0] ;
    PS>     $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
    PS>     write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" ;
    PS> } ;
    Demo that downloads a json db of licenses and filters out the Exchange mailbox supporting licenses.
    .LINK
    https://github.com/tostka/verb-ex2010
    #>
    #Requires -Modules verb-IO, verb-logging, verb-Text, AzureAD
    [OutputType('System.Collections.Hashtable')]
    [CmdletBinding()]
    PARAM(
        [Parameter(Mandatory=$FALSE,HelpMessage="ServicePlanName values that identify Exchange-mailbox supporting licenses[-ServicePlanName 'EXCHANGE_S_DESKLESS']")]
            #[ValidateNotNullOrEmpty()]
            #[ValidatePattern("^\w{3}$")]
            [string[]]$ServicePlanName = @('EXCHANGE_B_STANDARD','EXCHANGE_L_STANDARD','EXCHANGE_S_DESKLESS','EXCHANGE_S_ENTERPRISE','EXCHANGE_S_ESSENTIALS','EXCHANGE_S_STANDARD','EXCHANGE_S_STANDARD_MIDMARKET'),
            # above is full list of ServicePlans extracted from o365-ProductLicensesDb.json, other than 'EXCHANGE_ANALYTICS','EXCHANGE_S_ARCHIVE','EXCHANGE_S_ARCHIVE_ADDON'
            # our list
            #@('EXCHANGE_S_DESKLESS','EXCHANGE_S_STANDARD','EXCHANGE_S_ENTERPRISE'),
        [Parameter(Mandatory=$FALSE,HelpMessage="Regular Expression for ServicePlanName values should be excluded from results (Common Area Phone license etc)[-rgxBannedSPN 'EXCHANGE_S_DESKLESS']")]
            #[ValidateNotNullOrEmpty()]
            #[ValidatePattern("^\w{3}$")]
            [string]$rgxbannedSPN = '^MCOCAP$',
        [Parameter(Mandatory=$FALSE,HelpMessage="Switch to enable expanded non-legacy return (legacy returned SKU|Label|Notes; -outdetail adds Enabled|Consumed|Available|Warning|Suspended specs)[-OutDetail]")]
            #[ValidateNotNullOrEmpty()]
            #[ValidatePattern("^\w{3}$")]
            [switch]$OutDetail,
        [Parameter(Mandatory=$FALSE,HelpMessage="Switch to suppress normal post-filter Availability checks (Available -gt 0 & Enabled -gt 0) and return *any* matched supporting licenses in the Tenant [-unfiltered]")]
            #[ValidateNotNullOrEmpty()]
            #[ValidatePattern("^\w{3}$")]
            [switch]$Unfiltered,
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
            [string[]]$UserRole = @('ESvcCBA','CSvcCBA','SIDCBA'),
            #@('SID','CSVC'),
            # svcAcct use: @('ESvcCBA','CSvcCBA','SIDCBA')
        [Parameter(HelpMessage="Use EXOv2 (ExchangeOnlineManagement) over basic auth legacy connection [-useEXOv2]")]
            [switch] $useEXOv2=$true,
        [Parameter(HelpMessage="Silent output (suppress status echos)[-silent]")]
            [switch] $silent
    ) ;
    BEGIN { 
        # for scripts wo support, can use regions to fake BEGIN;PROCESS;END:
        # ps1 faked:#region BEGIN ; #*------v BEGIN v------
        #region CONSTANTS_AND_ENVIRO #*======v CONSTANTS_AND_ENVIRO v======
        # Debugger:proxy automatic variables that aren't directly accessible when debugging (must be assigned and read back from another vari) ; 
        $rPSCmdlet = $PSCmdlet ; 
        $rPSScriptRoot = $PSScriptRoot ; 
        $rPSCommandPath = $PSCommandPath ; 
        $rMyInvocation = $MyInvocation ; 
        $rPSBoundParameters = $PSBoundParameters ; 
        [array]$score = @() ; 
        if($rPSCmdlet.MyInvocation.InvocationName){
            if($rPSCmdlet.MyInvocation.InvocationName -match '\.ps1$'){
                $score+= 'ExternalScript' 
            }elseif($rPSCmdlet.MyInvocation.InvocationName  -match '^\.'){
                write-warning "dot-sourced invocation detected!:$($rPSCmdlet.MyInvocation.InvocationName)`n(will be unable to leverage script path etc from MyInvocation objects)" ; 
                # dot sourcing is implicit scripot exec
                $score+= 'ExternalScript' ; 
            } else {$score+= 'Function' };
        } ; 
        if($rPSCmdlet.CommandRuntime){
            if($rPSCmdlet.CommandRuntime.tostring() -match '\.ps1$'){$score+= 'ExternalScript' } else {$score+= 'Function' }
        } ; 
        $score+= $rMyInvocation.MyCommand.commandtype.tostring() ; 
        $grpSrc = $score | group-object -NoElement | sort count ;
        if( ($grpSrc |  measure | select -expand count) -gt 1){
            write-warning  "$score mixed results:$(($grpSrc| ft -a count,name | out-string).trim())" ;
            if($grpSrc[-1].count -eq $grpSrc[-2].count){
                write-warning "Deadlocked non-majority results!" ;
            } else {
                $runSource = $grpSrc | select -last 1 | select -expand name ;
            } ;
        } else {
            write-verbose "consistent results" ;
            $runSource = $grpSrc | select -last 1 | select -expand name ;
        };
        write-verbose "Calculated `$runSource:$($runSource)" ;
        'score','grpSrc' | get-variable | remove-variable ; # cleanup temp varis

        # function self-name (equiv to script's: $MyInvocation.MyCommand.Path) ;
        ${CmdletName} = $rPSCmdlet.MyInvocation.MyCommand.Name ;
        $PSParameters = New-Object -TypeName PSObject -Property $rPSBoundParameters ;
        write-verbose "`$rPSBoundParameters:`n$(($rPSBoundParameters|out-string).trim())" ;
        $Verbose = ($VerbosePreference -eq 'Continue') ; 
        # pre psv2, no $rPSBoundParameters autovari to check, so back them out:
        if($rPSCmdlet.MyInvocation.InvocationName){
            if($rPSCmdlet.MyInvocation.InvocationName  -match '^\.'){
                $smsg = "detected dot-sourced invocation: Skipping `$PSCmdlet.MyInvocation.InvocationName-tied cmds..." ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
            } else { 
                write-verbose 'Collect all non-default Params (works back to psv2 w CmdletBinding)'
                $ParamsNonDefault = (Get-Command $rPSCmdlet.MyInvocation.InvocationName).parameters | Select-Object -expand keys | Where-Object{$_ -notmatch '(Verbose|Debug|ErrorAction|WarningAction|ErrorVariable|WarningVariable|OutVariable|OutBuffer)'} ;
            } ; 
        } else { 
            $smsg = "(blank `$rPSCmdlet.MyInvocation.InvocationName, skipping Parameters collection)" ; 
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        } ; 
        #region ENVIRO_DISCOVER ; #*------v ENVIRO_DISCOVER v------
        <#
        # Debugger:proxy automatic variables that aren't directly accessible when debugging ; 
        $rPSScriptRoot = $PSScriptRoot ; 
        $rPSCommandPath = $PSCommandPath ; 
        $rMyInvocation = $MyInvocation ; 
        $rPSBoundParameters = $PSBoundParameters ; 
        #>
        $ScriptDir = $scriptName = '' ;     
        if($ScriptDir -eq '' -AND ( (get-variable -name rPSScriptRoot -ea 0) -AND (get-variable -name rPSScriptRoot).value.length)){
            $ScriptDir = $rPSScriptRoot
        } ; # populated rPSScriptRoot
        if( (get-variable -name rPSCommandPath -ea 0) -AND (get-variable -name rPSCommandPath).value.length){
            $ScriptName = $rPSCommandPath
        } ; # populated rPSCommandPath
        if($ScriptDir -eq '' -AND $runSource -eq 'ExternalScript'){$ScriptDir = (Split-Path -Path $rMyInvocation.MyCommand.Source -Parent)} # Running from File
        # when $runSource:'Function', $rMyInvocation.MyCommand.Source is empty,but on functions also tends to pre-hit from the rPSCommandPath entFile.FullPath ;
        if( $scriptname -match '\.psm1$' -AND $runSource -eq 'Function'){
            write-host "MODULE-HOMED FUNCTION:Use `$CmdletName to reference the running function name for transcripts etc (under a .psm1 `$ScriptName will reflect the .psm1 file  fullname)"
            if(-not $CmdletName){write-warning "MODULE-HOMED FUNCTION with BLANK `$CmdletNam:$($CmdletNam)" } ;
        } # Running from .psm1 module
        if($ScriptDir -eq '' -AND (Test-Path variable:psEditor)) {
            write-verbose "Running from VSCode|VS" ; 
            $ScriptDir = (Split-Path -Path $psEditor.GetEditorContext().CurrentFile.Path -Parent) ; 
                if($ScriptName -eq ''){$ScriptName = $psEditor.GetEditorContext().CurrentFile.Path }; 
        } ;
        if ($ScriptDir -eq '' -AND $host.version.major -lt 3 -AND $rMyInvocation.MyCommand.Path.length -gt 0){
            $ScriptDir = $rMyInvocation.MyCommand.Path ; 
            write-verbose "(backrev emulating `$rPSScriptRoot, `$rPSCommandPath)"
            $ScriptName = split-path $rMyInvocation.MyCommand.Path -leaf ;
            $rPSScriptRoot = Split-Path $ScriptName -Parent ;
            $rPSCommandPath = $ScriptName ;
        } ;
        if ($ScriptDir -eq '' -AND $rMyInvocation.MyCommand.Path.length){
            if($ScriptName -eq ''){$ScriptName = $rMyInvocation.MyCommand.Path} ;
            $ScriptDir = $rPSScriptRoot = Split-Path $rMyInvocation.MyCommand.Path -Parent ;
        }
        if ($ScriptDir -eq ''){throw "UNABLE TO POPULATE SCRIPT PATH, EVEN `$rMyInvocation IS BLANK!" } ;
        if($ScriptName){
            if(-not $ScriptDir ){$ScriptDir = Split-Path -Parent $ScriptName} ; 
            $ScriptBaseName = split-path -leaf $ScriptName ;
            $ScriptNameNoExt = [system.io.path]::GetFilenameWithoutExtension($ScriptName) ;
        } ; 
        # blank $cmdlet name comming through, patch it for Scripts:
        if(-not $CmdletName -AND $ScriptBaseName){
            $CmdletName = $ScriptBaseName
        }
        # last ditch patch the values in if you've got a $ScriptName
        if($rPSScriptRoot.Length -ne 0){}else{ 
            if($ScriptName){$rPSScriptRoot = Split-Path $ScriptName -Parent }
            else{ throw "Unpopulated, `$rPSScriptRoot, and no populated `$ScriptName from which to emulate the value!" } ; 
        } ; 
        if($rPSCommandPath.Length -ne 0){}else{ 
            if($ScriptName){$rPSCommandPath = $ScriptName }
            else{ throw "Unpopulated, `$rPSCommandPath, and no populated `$ScriptName from which to emulate the value!" } ; 
        } ; 
        if(-not ($ScriptDir -AND $ScriptBaseName -AND $ScriptNameNoExt  -AND $rPSScriptRoot  -AND $rPSCommandPath )){ 
            throw "Invalid Invocation. Blank `$ScriptDir/`$ScriptBaseName/`ScriptNameNoExt" ; 
            BREAK ; 
        } ; 
        # echo results dyn aligned:
        $tv = 'runSource','CmdletName','ScriptName','ScriptBaseName','ScriptNameNoExt','ScriptDir','PSScriptRoot','PSCommandPath','rPSScriptRoot','rPSCommandPath' ; 
        $tvmx = ($tv| Measure-Object -Maximum -Property Length).Maximum * -1 ; 
        $tv | get-variable | %{  write-verbose ("`${0,$tvmx} : {1}" -f $_.name,$_.value) } ; 
        'tv','tvmx'|get-variable | remove-variable ; # cleanup temp varis
        
        #endregion ENVIRO_DISCOVER ; #*------^ END ENVIRO_DISCOVER ^------

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

        write-verbose "Coerce blank Resultsize to Unlimited" ; 
        if( ($null -eq $ResultSize) -OR ('' -eq $ResultSize) ){$ResultSize = 'unlimited' }
        else {throw "Resultsize must be an integer or the string 'unlimited' (or blank)"} ;
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
        [array]$SmtpAttachment = $null ;

        # local Constants:
        #endregion CONSTANTS_AND_ENVIRO ; #*------^ END CONSTANTS_AND_ENVIRO ^------
    
        #region BANNER ; #*------v BANNER v------
        $sBnr="#*======v $(${CmdletName}): v======" ;
        $smsg = $sBnr ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level H1 } #Error|Warn|Debug
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        #endregion BANNER ; #*------^ END BANNER ^------

        #region SERVICE_CONNECTIONS #*======v SERVICE_CONNECTIONS v======
        # PRETUNE STEERING separately *before* pasting in balance of region
        #*------v STEERING VARIS v------
        $useO365 = $true ;
        $useEXO = $true ; 
        $UseOP=$false ; 
        $UseExOP=$false ;
        $useForestWide = $false; # flag to trigger cross-domain/forest-wide code in AD & EXoP
        $UseOPAD = $false ; 
        $UseMSOL = $false ; # should be hard disabled now in o365
        $UseAAD = $false  ; 
        $useO365 = [boolean]($useO365 -OR $useEXO -OR $UseMSOL -OR $UseAAD)
        $UseOP = [boolean]($UseOP -OR $UseExOP -OR $UseOPAD) ;
        #*------^ END STEERING VARIS ^------
        #*------v EXO V2/3 steering constants v------
        $EOMModName =  'ExchangeOnlineManagement' ;
        $EOMMinNoWinRMVersion = $MinNoWinRMVersion = '3.0.0' ; # support both names
        #*------^ END EXO V2/3 steering constants ^------
        # assert Org from Credential specs (if not param'd)
        # 1:36 PM 7/7/2023 and revised again -  revised the -AND, for both, logic wasn't working
        if($TenOrg){    
            $smsg = "Confirmed populated `$TenOrg" ; 
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        } elseif(-not($tenOrg) -and $Credential){
            $smsg = "(unconfigured `$TenOrg: asserting from credential)" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Prompt } 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            $TenOrg = get-TenantTag -Credential $Credential ;
        } else { 
            # if not using Credentials or a TargetTenants/TenOrg loop, default the $TenOrg on the $env:USERDOMAIN
            $smsg = "(unconfigured `$TenOrg & *NO* `$Credential: fallback asserting from `$env:USERDOMAIN)" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Prompt } 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            switch -regex ($env:USERDOMAIN){
                ([regex]('(' + (( @($TORMeta.legacyDomain,$CMWMeta.legacyDomain)  |foreach-object{[regex]::escape($_)}) -join '|') + ')')).tostring() {$TenOrg = $env:USERDOMAIN.substring(0,3).toupper() } ;
                $TOLMeta.legacyDomain {$TenOrg = 'TOL' }
                default {throw "UNRECOGNIZED `$env:USERDOMAIN!:$($env:USERDOMAIN)" ; exit ; } ;
            } ; 
        } ; 
        #region useO365 ; #*------v useO365 v------
        #$useO365 = $false ; # non-dyn setting, drives variant EXO reconnect & query code
        #if($CloudFirst){ $useO365 = $true } ; # expl: steering on a parameter
        if($useO365){
            #region GENERIC_EXO_CREDS_&_SVC_CONN #*------v GENERIC EXO CREDS & SVC CONN BP v------
            # o365/EXO creds
            <### Usage: Type defaults to SID, if not spec'd - Note: there must be a *logged in & configured *profile*
            $o365Cred = get-TenantCredentials -TenOrg $TenOrg -verbose -userrole SID ;
            Returns a credential set for the $TenOrg Hybrid OnPrem Exchange Org
            .EXAMPLE
            $o365Cred = get-TenantCredentials -TenOrg $TenOrg -verbose -userrole CSVC ;
            Returns the CSVC Userrole credential for the $TenOrg Hybrid OnPrem Exchange Org
            .EXAMPLE
            $o365Cred = get-TenantCredentials -TenOrg $TenOrg -verbose -userrole B2BI ;
            Returns the B2BI Userrole credential for the $TenOrg Hybrid OnPrem Exchange Org
            ###>
            $o365Cred = $null ;
            if($Credential){
                $smsg = "`Credential:Explicit credentials specified, deferring to use..." ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                 # get-TenantCredentials() return format: (emulating)
                 $o365Cred = [ordered]@{
                    Cred=$Credential ; 
                    credType=$null ; 
                } ; 
                $uRoleReturn = resolve-UserNameToUserRole -UserName $Credential.username -verbose:$($VerbosePreference -eq "Continue") ; # Username
                #$uRoleReturn = resolve-UserNameToUserRole -Credential $Credential -verbose = $($VerbosePreference -eq "Continue") ;   # full Credential support
                if($uRoleReturn.UserRole){
                    $o365Cred.credType = $uRoleReturn.UserRole ; 
                } else { 
                    $smsg = "Unable to resolve `$credential.username ($($credential.username))"
                    $smsg += "`nto a usable 'UserRole' spec!" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    throw $smsg ;
                    Break ;
                } ; 
            } else { 
                $pltGTCred=@{TenOrg=$TenOrg ; UserRole=$null; verbose=$($verbose)} ;
                if($UserRole){
                    $smsg = "(`$UserRole specified:$($UserRole -join ','))" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    $pltGTCred.UserRole = $UserRole; 
                } else { 
                    $smsg = "(No `$UserRole found, defaulting to:'CSVC','SID' " ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                    $pltGTCred.UserRole = 'CSVC','SID' ; 
                } ; 
                $smsg = "get-TenantCredentials w`n$(($pltGTCred|out-string).trim())" ; 
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level verbose } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                $o365Cred = get-TenantCredentials @pltGTCred
            } ; 
            if($o365Cred.credType -AND $o365Cred.Cred -AND $o365Cred.Cred.gettype().fullname -eq 'System.Management.Automation.PSCredential'){
                $smsg = "(validated `$o365Cred contains .credType:$($o365Cred.credType) & `$o365Cred.Cred.username:$($o365Cred.Cred.username)" ;
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                # 9:58 AM 6/13/2024 populate $credential with return, if not populated (may be required for follow-on calls that pass common $Credentials through)
                if((gv Credential) -AND $Credential -eq $null){
                    $credential = $o365Cred.Cred ;
                }elseif($credential.gettype().fullname -eq 'System.Management.Automation.PSCredential'){
                    $smsg = "(`$Credential is properly populated; explicit -Credential was in initial call)" ; 
                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                } else {
                    $smsg = "`$Credential is `$NULL, AND $o365Cred.Cred is unusable to populate!" ;
                    $smsg = "downstream commands will *not* properly pass through usable credentials!" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    throw $smsg ;
                    break ;
                } ;
            } else {
                $smsg = "UNABLE TO RESOLVE FUNCTIONAL CredType/UserRole from specified explicit -Credential:$($Credential.username)!" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                break ;
            } ; 
            if($o365Cred){
                # make it script scope, so we don't have to predetect & purge before using new-variable
                if(get-Variable -Name cred$($tenorg) -scope Script -ea 0 ){ remove-Variable -Name cred$($tenorg) -scope Script } ;
                New-Variable -Name cred$($tenorg) -scope Script -Value $o365Cred.cred ;
                $smsg = "Resolved $($Tenorg) `$o365cred:$($o365Cred.cred.username) (assigned to `$cred$($tenorg))" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } else {
                $statusdelta = ";ERROR";
                $script:PassStatus += $statusdelta ;
                set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatettus_$($tenorg)).value + $statusdelta) ;
                $smsg = "Unable to resolve $($tenorg) `$o365Cred value!"
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                throw "Unable to resolve $($tenorg) `$o365Cred value!`nEXIT!"
                Break ;
            } ;
            # if we get here, wo a $Credential, w resolved $o365Cred, assign it 
            if(-not $Credential -AND $o365Cred){$Credential = $o365Cred.cred } ; 
            # configure splat for connections: (see above useage)
            # downstream commands
            $pltRXO = [ordered]@{
                Credential = $Credential ;
                verbose = $($VerbosePreference -eq "Continue")  ;
            } ;
            if((get-command Reconnect-EXO).Parameters.keys -contains 'silent'){
                $pltRxo.add('Silent',$silent) ;
            } ;
            # default connectivity cmds - force silent false
            $pltRXOC = [ordered]@{} ; $pltRXO.GetEnumerator() | ?{ $_.Key -notmatch 'silent' }  | ForEach-Object { $pltRXOC.Add($_.Key, $_.Value) } ; $pltRXOC.Add('silent',$true) ; 
            if((get-command Reconnect-EXO).Parameters.keys -notcontains 'silent'){
                $pltRxo.remove('Silent') ;
            } ; 
            #region EOMREV ; #*------v EOMREV Check v------
            #$EOMmodname = 'ExchangeOnlineManagement' ;
            $pltIMod = @{Name = $EOMmodname ; ErrorAction = 'Stop' ; verbose=$false} ;
            # do a gmo first, faster than gmo -list
            if([version]$EOMMv = (Get-Module @pltIMod).version){}
            elseif([version]$EOMMv = (Get-Module -ListAvailable @pltIMod).version){}
            else {
                $smsg = "$($EOMmodname) PowerShell v$($MinNoWinRMVersion) module is required, do you want to install it?" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Prompt }
                else{ write-host -foregroundcolor YELLOW "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                $bRet = Read-Host "Enter YYY to continue. Anything else will exit"  ;
                if ($bRet.ToUpper() -eq "YYY") {
                    $smsg = "Installing $($EOMmodname) module..." ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    Install-Module $EOMmodname -Repository PSGallery -AllowClobber -Force ;
                } else {
                    $smsg = "Please install $($EOMmodname) PowerShell v$($MinNoWinRMVersion)  module." ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN }
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    #exit 1
                    break ;
                }  ;
            } ;
            $smsg = "(Checking for WinRM support in this EOM rev...)" ; 
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            if([version]$EOMMv -ge [version]$MinNoWinRMVersion){
                $MinNoWinRMVersion = $EOMMv.tostring() ;
                $IsNoWinRM = $true ;
            }elseif([version]$EOMMv -lt [version]$MinimumVersion){
                $smsg = "Installed $($EOMmodname) is v$($MinNoWinRMVersion): This module is obsolete!" ;
                $smsg += "`nAnd unsupported by this function!" ;
                $smsg += "`nPlease install $($EOMmodname) PowerShell v$($MinNoWinRMVersion)  module!" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                Break ;
            } else {
                $IsNoWinRM = $false ;
            } ;
            [boolean]$UseConnEXO = [boolean]([version]$EOMMv -ge [version]$MinNoWinRMVersion) ;
            #endregion EOMREV ; #*------^ END EOMREV Check  ^------
            #-=-=-=-=-=-=-=-=
            <### CALLS ARE IN FORM: (cred$($tenorg))
            # downstream commands
            $pltRXO = @{
                Credential = (Get-Variable -name cred$($tenorg) ).value ;
                #verbose = $($verbose) ;
                Verbose = $FALSE ; 
            } ;
            if((get-command Reconnect-EXO).Parameters.keys -contains 'silent'){
                $pltRxo.add('Silent',$false) ;
            } ; 
            # default connectivity cmds - force silent false
            $pltRXOC = [ordered]@{} ; $pltRXO.GetEnumerator() | ?{ $_.Key -notmatch 'silent' }  | ForEach-Object { $pltRXOC.Add($_.Key, $_.Value) } ; $pltRXOC.Add('silent',$true) ;
            if((get-command Reconnect-EXO).Parameters.keys -notcontains 'silent'){
                $pltRxo.remove('Silent') ;
            } ; 
            #$pltRXO creds & .username can also be used for AzureAD connections:
            #Connect-AAD @pltRXOC ;
            ###>
            #endregion GENERIC_EXO_CREDS_&_SVC_CONN #*------^ END GENERIC EXO CREDS & SVC CONN BP ^------

        } else {
            $smsg = "(`$useO365:$($useO365))" ; 
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        } ; # if-E if($useO365 ){
        #endregion useO365 ; #*------^ END useO365 ^------

        #region useEXO ; #*------v useEXO v------
        # 1:29 PM 9/15/2022 as of MFA & v205, have to load EXO *before* any EXOP, or gen get-steppablepipeline suffix conflict error
        if($useEXO){
            if ($script:useEXOv2 -OR $useEXOv2) { reconnect-eXO2 @pltRXOC }
            else { reconnect-EXO @pltRXOC } ;
        } else {
            $smsg = "(`$useEXO:$($useEXO))" ; 
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        } ; # if-E 
        #endregion  ; #*------^ END useEXO ^------
        
        #region GENERIC_EXOP_CREDS_&_SRVR_CONN #*------v GENERIC EXOP CREDS & SRVR CONN BP v------
        # steer all onprem code on $XXXMeta.ExOPAccessFromToro & Ex10Server values
        #$UseOP=$true ; 
        #$UseExOP=$true ;
        #$useForestWide = $true ; # flag to trigger cross-domain/forest-wide code in AD & EXoP
        <# no onprem dep
        if((Get-Variable  -name "$($TenOrg)Meta").value.ExOPAccessFromToro -AND (Get-Variable  -name "$($TenOrg)Meta").value.Ex10Server){
            $UseOP = $UseExOP = $true ;
            $smsg = "$($TenOrg):Meta.ExOPAccessFromToro($((Get-Variable  -name "$($TenOrg)Meta").value.ExOPAccessFromToro)) -AND/OR Meta.Ex10Server($((Get-Variable  -name "$($TenOrg)Meta").value.Ex10Server)),`ENABLING use of OnPrem Ex system this pass." ;
            if($verbose){ if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
        } else {
            $UseOP = $UseExOP = $false ;
            $smsg = "$($TenOrg):Meta.ExOPAccessFromToro($((Get-Variable  -name "$($TenOrg)Meta").value.ExOPAccessFromToro)) -AND/OR Meta.Ex10Server($((Get-Variable  -name "$($TenOrg)Meta").value.Ex10Server)),`nDISABLING use of OnPrem Ex system this pass." ;
            if($verbose){ if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
        } ;
        #>
        if($UseOP){
            #*------v GENERIC EXOP CREDS & SRVR CONN BP v------
            # do the OP creds too
            $OPCred=$null ;
            # default to the onprem svc acct
            # userrole='ESVC','SID'
            #$pltGHOpCred=@{TenOrg=$TenOrg ;userrole='ESVC','SID'; verbose=$($verbose)} ;
            # userrole='SID','ESVC'
            $pltGHOpCred=@{TenOrg=$TenOrg ;userrole='SID','ESVC'; verbose=$($verbose)} ;
            $smsg = "get-HybridOPCredentials w`n$(($pltGHOpCred|out-string).trim())" ; 
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level verbose } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            if($OPCred=(get-HybridOPCredentials @pltGHOpCred).cred){
                # make it script scope, so we don't have to predetect & purge before using new-variable
                if(get-Variable -Name "cred$($tenorg)OP" -scope Script -ea 0 ){ remove-Variable -Name "cred$($tenorg)OP" -scope Script } ;
                New-Variable -Name "cred$($tenorg)OP" -scope Script -Value $OPCred ;
                $smsg = "Resolved $($Tenorg) `$OPCred:$($OPCred.username) (assigned to `$cred$($tenorg)OP)" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } else {
                $statusdelta = ";ERROR";
                $script:PassStatus += $statusdelta ;
                set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta) ;
                $smsg = "Unable to resolve get-HybridOPCredentials -TenOrg $($TenOrg) -userrole 'ESVC' value!"
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                throw "Unable to resolve $($tenorg) `$OPCred value!`nEXIT!"
                Break ;
            } ;
            $smsg= "Using OnPrem/EXOP cred:`$cred$($tenorg)OP:$((Get-Variable -name "cred$($tenorg)OP" ).value.username)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            <### CALLS ARE IN FORM: (cred$($tenorg))
            $pltRX10 = @{
                Credential = (Get-Variable -name "cred$($tenorg)OP" ).value ;
                #verbose = $($verbose) ;
                Verbose = $FALSE ; 
            } ;
            $1stConn = $false ; # below uses silent suppr for both x10 & xo!
            if($1stConn){
                $pltRX10.silent = $pltRXO.silent = $false ;
            } else {
                $pltRX10.silent = $pltRXO.silent =$true ;
            } ;
            if($pltRX10){ReConnect-Ex2010 @pltRX10 }
            else {ReConnect-Ex2010 }
            #$pltRx10 creds & .username can also be used for local ADMS connections
            ###>
            $pltRX10 = @{
                Credential = (Get-Variable -name "cred$($tenorg)OP" ).value ;
                #verbose = $($verbose) ;
                Verbose = $FALSE ; 
            } ;
            if((get-command Reconnect-Ex2010).Parameters.keys -contains 'silent'){
                $pltRX10.add('Silent',$false) ;
            } ;

            # defer cx10/rx10, until just before get-recipients qry
            #endregion GENERIC_EXOP_CREDS_&_SRVR_CONN #*------^ END GENERIC EXOP CREDS & SRVR CONN BP ^------
            # connect to ExOP X10
            if($useEXOP){
                if($pltRX10){
                    #ReConnect-Ex2010XO @pltRX10 ;
                    ReConnect-Ex2010 @pltRX10 ;
                } else { Reconnect-Ex2010 ; } ;
                #Add-PSSnapin -Name 'Microsoft.Exchange.Management.PowerShell.SnapIn'
                #TK: add: test Exch & AD functional connections
                TRY{
                    if(get-command -module (get-module |?{$_.name -like 'tmp_*'}).name -name 'get-OrganizationConfig'){} else {
                        $smsg = "(mangled Ex10 conn: dx10,rx10...)" ; 
                        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                        disconnect-ex2010 ; reconnect-ex2010 ; 
                    } ; 
                    if(-not ($OrgName = ((get-OrganizationConfig).DistinguishedName.split(',') |?{$_ -like 'DC=*'}) -join '.' -replace 'DC=','')){
                        $smsg = "Missing Exchange Connection! (no (Get-OrganizationConfig).name returned)" ; 
                        throw $smsg ; 
                        $smsg | write-warning  ; 
                    } ; 
                } CATCH {
                    $ErrTrapd=$Error[0] ;
                    $smsg = $ErrTrapd ;
                    $smsg += "`n";
                    $smsg += $ErrTrapd.Exception.Message ;
                    if ($logging) { _write-log -LogContent $smsg -Path $logfile -useHost -Level WARN } 
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                    CONTINUE ;
                } ;
            } else { 
            
            } ; 
            if($useForestWide){
                #region  ; #*------v OPTIONAL CODE TO ENABLE FOREST-WIDE NATIVE EXCHANGE SUPPORT v------
                $smsg = "(`$useForestWide:$($useForestWide)):Enabling EXoP Forestwide)" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                $smsg = 'Set-AdServerSettings -ViewEntireForest `$True' ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                Set-AdServerSettings -ViewEntireForest $True ;
                #endregion  ; #*------^ END OPTIONAL CODE TO ENABLE FOREST-WIDE NATIVE EXCHANGE SUPPORT ^------
            } ;
        } else {
            $smsg = "(`$useOP:$($UseOP))" ; 
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        }  ;  # if-E $UseOP

        #region UseOPAD #*------v UseOPAD v------
        if($UseOP -OR $UseOPAD){
            #region GENERIC_ADMS_CONN_&_XO #*------v GENERIC ADMS CONN & XO  v------
            $smsg = "(loading ADMS...)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            # always capture load-adms return, it outputs a $true to pipeline on success
            $ADMTLoaded = load-ADMS -Verbose:$FALSE ;
            # 9:32 AM 4/20/2023 trimmed disabled/fw-borked cross-org code
            TRY {
                if(-not(Get-ADDomain  -ea STOP).DNSRoot){
                    $smsg = "Missing AD Connection! (no (Get-ADDomain).DNSRoot returned)" ; 
                    throw $smsg ; 
                    $smsg | write-warning  ; 
                } ; 
                $objforest = get-adforest -ea STOP ; 
                # Default new UPNSuffix to the UPNSuffix that matches last 2 elements of the forestname.
                $forestdom = $UPNSuffixDefault = $objforest.UPNSuffixes | ?{$_ -eq (($objforest.name.split('.'))[-2..-1] -join '.')} ; 
                if($useForestWide){
                    #region  ; #*------v OPTIONAL CODE TO ENABLE FOREST-WIDE AD GC QRY SUPPORT v------
                    $smsg = "(`$useForestWide:$($useForestWide)):Enabling AD Forestwide)" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    $smsg = 'Set-AdServerSettings -ViewEntireForest `$True' ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    #TK 9:44 AM 10/6/2022 need org wide for rolegrps in parent dom (only for onprem RBAC, not EXO)
                    $GcFwide = "$((Get-ADDomainController -Discover -Service GlobalCatalog).hostname):3268" ;        
                    #endregion  ; #*------^ END  OPTIONAL CODE TO ENABLE FOREST-WIDE AD GC QRY SUPPORT  ^------
                } ;    
            } CATCH {
                $ErrTrapd=$Error[0] ;
                $smsg = $ErrTrapd ;
                $smsg += "`n";
                $smsg += $ErrTrapd.Exception.Message ;
                if ($logging) { _write-log -LogContent $smsg -Path $logfile -useHost -Level WARN } 
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                CONTINUE ;
            } ;        
            #endregion GENERIC_ADMS_CONN_&_XO #*------^ END GENERIC ADMS CONN & XO ^------
        } else {
            $smsg = "(`$UseOP:$($UseOP))" ; 
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        }  ;
        #if (!$domaincontroller) { $domaincontroller = get-gcfast } ;
        #if(!$domaincontroller){ if(test-path function:get-gcfast){$domaincontroller = get-gcfast} else { throw "no get-gcfast()!" } ;} else {"(existing `$domaincontroller:$($domaincontroller))"} ;
        # use new get-GCFastXO cross-org dc finde
        # default to Op_ExADRoot forest from $TenOrg Meta
        #if($UseOP -AND -not $domaincontroller){
        if($UseOP -AND -not (get-variable domaincontroller -ea 0)){
            #$domaincontroller = get-GCFastXO -TenOrg $TenOrg -subdomain ((get-variable -name "$($TenOrg)Meta").value['OP_ExADRoot']) -verbose:$($verbose) |?{$_.length};
            # need to debug the above, credential issue?
            # just get it done
            $domaincontroller = get-GCFast
        }  else { 
            # have to defer to get-azuread, or use EXO's native cmds to poll grp members
            # TODO 1/15/2021
            $useEXOforGroups = $true ; 
            $smsg = "$($TenOrg):HAS NO ON-PREM ACTIVEDIRECTORY, DEFERRING ALL GROUP ACCESS & MGMT TO NATIVE EXO CMDS!" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        } ; 
        if($useForestWide -AND -not $GcFwide){
            #region  ; #*------v OPTIONAL CODE TO ENABLE FOREST-WIDE ACTIVEDIRECTORY SUPPORT: v------
            $smsg = "`$GcFwide = Get-ADDomainController -Discover -Service GlobalCatalog" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            $GcFwide = "$((Get-ADDomainController -Discover -Service GlobalCatalog).hostname):3268" ;
            $smsg = "Discovered `$GcFwide:$($GcFwide)" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            #endregion  ; #*------^ END OPTIONAL CODE TO ENABLE FOREST-WIDE ACTIVEDIRECTORY SUPPORT ^------
        } ;
        #endregion UseOPAD #*------^ END UseOPAD ^------

        #region MSOL_CONNECTION ; #*------v  MSOL CONNECTION v------
        #$UseMSOL = $false 
        if($UseMSOL){
            #$reqMods += "connect-msol".split(";") ;
            #if ( !(check-ReqMods $reqMods) ) { write-error "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Missing function. EXITING." ; Break ; }  ;
            $smsg = "(loading MSOL...)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } Error|Warn|Debug
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            #connect-msol ;
            connect-msol @pltRXOC ;
        } else {
            $smsg = "(`$UseMSOL:$($UseMSOL))" ; 
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        } ;
        #endregion MSOL_CONNECTION ; #*------^  MSOL CONNECTION ^------

        #region AZUREAD_CONNECTION ; #*------v AZUREAD CONNECTION v------
        #$UseAAD = $false 
        if($UseAAD){
            #$reqMods += "Connect-AAD".split(";") ;
            #if ( !(check-ReqMods $reqMods) ) { write-error "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Missing function. EXITING." ; Break ; }  ;
            $smsg = "(loading AAD...)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            Connect-AAD @pltRXOC ;
        } else {
            $smsg = "(`$UseAAD:$($UseAAD))" ; 
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        } ;
        #endregion AZUREAD_CONNECTION ; #*------^ AZUREAD CONNECTION ^------
        
        <# defined above
        # EXO connection
        $pltRXO = @{
            Credential = (Get-Variable -name cred$($tenorg) ).value ;
            verbose = $($verbose) ; } ;
        #>
        <#
        if($VerbosePreference = "Continue"){
            $VerbosePrefPrior = $VerbosePreference ;
            $VerbosePreference = "SilentlyContinue" ;
            $verbose = ($VerbosePreference -eq "Continue") ;
        } ;
        disconnect-exo ;
        if ($script:useEXOv2) { reconnect-eXO2 @pltRXOC }
        else { reconnect-EXO @pltRXOC } ;
        # reenable VerbosePreference:Continue, if set, during mod loads
        if($VerbosePrefPrior -eq "Continue"){
            $VerbosePreference = $VerbosePrefPrior ;
            $verbose = ($VerbosePreference -eq "Continue") ;
        } ;
        #>
        #-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
        #endregion SERVICE_CONNECTIONS #*======^ END SERVICE_CONNECTIONS ^======

        # check if using Pipeline input or explicit params:
        if ($rPSCmdlet.MyInvocation.ExpectingInput) {
            $smsg = "Data received from pipeline input: '$($InputObject)'" ;
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        } else {
            # doesn't actually return an obj in the echo
            #$smsg = "Data received from parameter input: '$($InputObject)'" ;
            #if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
            #else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        } ;

        #endregion SUBMAIN ; #*======^ END SUB MAIN ^======
    } ;  # BEGIN-E
    PROCESS {
        $Error.Clear() ; 
        <#$propsAADL = 'SkuId',  'SkuPartNumber',  @{name='Enabled';Expression={$_.PrepaidUnits.enabled }},
            @{name='Consumed';Expression={$_.ConsumedUnits} }, @{name='Available';Expression={$_.PrepaidUnits.enabled - $_.ConsumedUnits} },
            @{name='Warning';Expression={$_.PrepaidUnits.warning} }, @{name='Suspended';Expression={$_.PrepaidUnits.suspended} } ;
            #>
        $propsAADL = 'SkuId',  'SkuPartNumber',  @{name='Enabled';Expression={$_.PrepaidUnits.enabled }},
            @{name='Consumed';Expression={$_.ConsumedUnits} }, @{name='Available';Expression={$_.PrepaidUnits.enabled - $_.ConsumedUnits} },
            @{name='Warning';Expression={$_.PrepaidUnits.warning} }, @{name='Suspended';Expression={$_.PrepaidUnits.suspended} },
            @{name='ServicePlanName';Expression={(($_.ServicePlans).ServicePlanName |?{$_ -match 'EXCHANGE_'})}} ;

        $rgxExSvcPlans = ('(' + (($ServicePlanName |%{[regex]::escape($_)}) -join '|') + ')') ; 

        $ExMbxLicenses = @() ; 
        foreach($SPName in $ServicePlanName){
            $smsg = $sBnrS="`n#*------v PROCESSING $($spname): v------" ; 
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            
            TRY{
                if($thisPlan = Get-AzureADSubscribedSku | ?{$_.serviceplans.serviceplanname -match $SPName} | select $propsAADL | ?{$_.SkuPartNumber -notmatch $rgxbannedSPN}){
                    if(-not $Unfiltered){
                        $smsg = "Postfilter:`$_.available -gt 0 -AND `$_.Enabled -gt 0 " ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        $thisplan = $thisplan |?{$_.available -gt 0 -AND $_.Enabled -gt 0} ; 
                    } else {} ; 
                    if($thisplan){
                        #$OutDetail, $Unfiltered
                        foreach($item in $thisplan){
                            $smsg = $sBnr3="`n#*~~~~~~v $($item.SkuPartNumber) : v~~~~~~" ; 
                            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                            if($OutDetail){
                                $oReturn = [ordered]@{
                                    SKU = $item.SkuPartNumber ;                                 
                                    Label= $item.SkuPartNumber.split('_')[-1] ; 
                                    Notes= "$($item.SkuPartNumber):Detailed usage" ; 
                                    ServicePlanName = $item.ServicePlanName ; 
                                    Enabled  = $item.Enabled ; 
                                    Consumed = $item.Consumed ; 
                                    Available = $item.Available ; 
                                    warning = $item.Warning ; 
                                    Suspended = $item.Suspended ; 
                                } ; 
                            }else {
                                $oReturn = [ordered]@{
                                    SKU = $item.SkuPartNumber ;                                 
                                    Label= $item.SkuPartNumber.split('_')[-1] ; 
                                    Notes= "Enabled:{0}|Consumed:{1}|Avail:{2}|Warn:{3}|Susp:{4}" -f $item.Enabled,$item.Consumed,$item.Available,$item.WArning,$item.Suspended ; 
                                } ; 
                            } ; 
                            $ExMbxLicenses += New-Object PSObject -Property $oReturn ;

                            $smsg = "$($sBnr3.replace('~v','~^').replace('v~','^~'))`n" ;
                            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                        } ; 
                    } else { 
                        #$smsg = "NOTE:serviceplans.serviceplanname -match $($SPName) yielded *NONE* with Available and Enabled -gt 0!" ; 
                        #if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } 
                        #else{ write-host -foregroundcolor yellow "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                    } ; 
                } else { } ; 

            } CATCH {
                $ErrTrapd=$Error[0] ;
                $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } ; 
    

            

            
            $smsg = "$($sBnrS.replace('-v','-^').replace('v-','^-'))" ;
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 

        } ; 

    # ===========

<#
        # input table of Exchange Online assignable licenses that include a UserMailbox:
    $ExMbxLicensesTbl = @"
|SKU|Label|Notes|
|ENTERPRISEPACK|Office 365 Enterprise E3|OfficE; EXO (OL,OWA,OM,100G mbx)|
|EXCHANGESTANDARD|Exchange Online Plan 1|No Office; no Services; 50G mbx, No ArchiveMbx|
|SPE_F1|Microsoft 365 F3| OfficeWeb, OfficeMobile; EXO (OWA,OM 2G Mbx)|(formerly Microsoft 365 F1, renamed Mar2020)
|STANDARDPACK|OFFICE 365 E1| OfficeWeb, OfficeMobile; EXO (OWA,OM 50G Mbx)
|EXCHANGEENTERPRISE_FACULTY|Exch Online Plan 2 for Faculty|No Office; no Services; 100G mbx, +ArchiveMbx, +vmail, +DLP|
|EXCHANGE_L_STANDARD|Exchange Online (Plan 1)|No Office; no Services; 50G mbx, No ArchiveMbx|
|EXCHANGE_S_ENTERPRISE|Exchange Online Plan 2 S|No Office; no Services; 100G mbx, +ArchiveMbx, +vmail, +DLP|
|EXCHANGEENTERPRISE|Exchange Online Plan 2|No Office; no Services; 50G mbx, +ArchiveMbx, +vmail, +DLP|
|STANDARDWOFFPACK_STUDENT|O365 Education E1 for Students|OfficeWeb, OfficeMobile; EXO (OWA,OM 50G Mbx)|
|STANDARDWOFFPACK_IW_FACULTY|O365 Education for Faculty||
|STANDARDWOFFPACK_IW_STUDENT|O365 Education for Students||
|STANDARDPACK_STUDENT|Office 365 (Plan A1) for Students||
|ENTERPRISEPACKLRG|Office 365 (Plan E3)||
|STANDARDWOFFPACK_FACULTY|Office 365 Education E1 for Faculty|OfficeWeb, OfficeMobile; EXO (OWA,OM 50G Mbx)|
|ENTERPRISEWITHSCAL_FACULTY|Office 365 Education E4 for Faculty||
|ENTERPRISEWITHSCAL_STUDENT|Office 365 Education E4 for Students||
|STANDARDPACK|Office 365 Enterprise E1|OfficeWeb, OfficeMobile; EXO (OWA,OM 50G Mbx)|
|STANDARDWOFFPACK|Office 365 Enterprise E2|OfficeWeb, OfficeMobile; EXO (OWA,OM 50G Mbx), No ArchiveMbx|
|ENTERPRISEPACKWITHOUTPROPLUS|Office 365 Enterprise E3 without ProPlus Add-on||
|ENTERPRISEWITHSCAL|Office 365 Enterprise E4||
|ENTERPRISEPREMIUM|Office 365 Enterprise E5|OfficE; EXO (OL,OWA,OM,100G mbx),AAD P1 & P2, Az Info Protection Plan 2; UC; ATP|
|DESKLESSPACK_YAMMER|Office 365 Enterprise K1 with Yammer||
|DESKLESSPACK|Office 365 Enterprise K1 without Yammer||
|DESKLESSWOFFPACK|Office 365 Enterprise K2||
|MIDSIZEPACK|Office 365 Midsize Business||
|STANDARDWOFFPACKPACK_FACULTY|Office 365 Plan A2 for Faculty||
|STANDARDWOFFPACKPACK_STUDENT|Office 365 Plan A2 for Students||
|ENTERPRISEPACK_FACULTY|Office 365 Plan A3 for Faculty||
|ENTERPRISEPACK_STUDENT|Office 365 Plan A3 for Students||
|OFFICESUBSCRIPTION_FACULTY|Office 365 ProPlus for Faculty||
|LITEPACK_P2|Office 365 Small Business Premium||
|SPE_E3|MICROSOFT 365 E3|OfficeWeb, OfficeMobile; EXO (OL,OWA,OM 2G Mbx)||
|SPE_E5|MICROSOFT 365 E5||
"@ ;
        $ExMbxLicenses = $ExMbxLicensesTbl | convertfrom-markdowntable ;
#>
        # building a CustObj (actually an indexed hash) with the data. The 'index' for each license, is the Sku/SkuPartNumber
        $smsg = "(converting $(($ExMbxLicenses|measure).count) UserMailbox-supporting o365 Licenses to indexed hash)" ;     
        if($verbose){
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
        } ; 
        if($host.version.major -gt 2){$hExMbxLicenses = [ordered]@{} } 
        else { $hExMbxLicenses = @{} } ;
    
        $ttl = ($ExMbxLicenses|measure).count ; $Procd = 0 ; 
        foreach ($Sku in $ExMbxLicenses){
            $Procd ++ ; 
            $sBnrS="`n#*------v PROCESSING : ($($Procd)/$($ttl)) $($Sku.SKU) v------" ; 
            $smsg = $sBnrS ; 
            if($verbose){
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } ; 
        
            $name =$($Sku | select -expand SKU) ; 
            $hExMbxLicenses[$name] = $Sku ; 

            $smsg = "$($sBnrS.replace('-v','-^').replace('v-','^-'))" ;
            if($verbose){
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } ; 
        } ;  # loop-E

        if($hExMbxLicenses){
            $smsg = "(Returning summary objects to pipeline)" ; 
            if($verbose){
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } ; 
            $hExMbxLicenses | Write-Output ; 
        } else {
            $smsg = "NO RETURNABLE `$hExMbxLicenses OBJECT!" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug 
            else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
            THROW $smsg ;
        } ; 
    } ;  # PROC-E
    END {
        $smsg = "$($sBnr.replace('=v','=^').replace('v=','^='))" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level H1 } #Error|Warn|Debug
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
    } ;  # END-E
}

#*------^ get-ExoMailboxLicenses.ps1 ^------