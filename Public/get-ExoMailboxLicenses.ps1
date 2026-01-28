# get-ExoMailboxLicenses.ps1

#*------v get-ExoMailboxLicenses.ps1 v------
function get-ExoMailboxLicenses {
<#
    .SYNOPSIS
    get-ExoMailboxLicenses - Provides a prefab indexed hash of Exchange-Online mailbox-supporting licenses (at least one of which is required to accomodate an EXO Usermailbox - This now dynamically calls Get-MgSubscribedSku and postfilters the ServicePlan list for matches on the $ServicePlanName array (which reflects Exchange mailbox ServicePlanNames). The ServicePlanName array must be manually updated to accomodate MS licensure changes over time).
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
    * 12:41 PM 1/27/2026 latest conn_svcs block updated
    * 10:48 AM 1/19/2026 bugfix: $pltCcOPSvcs.UserRole (postfilter, not match test)
    # 2:23 PM 1/12/2026 fixed op break: #804, FORCE:  $useOP = $false ; : This has NO NEED FOR ONPREM, IT JUST DOES A Get-MgSubscribedSku AND RUNS THE STATIC SERVICEPLANLIST CHECK, NOTHING ELSE!
    * 1:30 PM 1/7/2026 WIP unupdated port from AADLicense -> MGUDLicense
    * 10:23 AM 11/27/2024 cleaned up CBH expls, pulled spurious quota #1, added in its place a simple call & review; updated #2's comment to actually reflect what's going on (vs the original indexed hash pull for quotas, it wass lifted from).
    * 1:22 PM 6/18/2024 updated SERVICE_CONNECTIONS block; reflects latest variant; 
    * 4:16 PM 6/17/2024 add: ServicePlanName to the detailed output ; pulled transcript END block (unused, would kill other process logging) ; adding -OutDetail, need to implement to return avail etc full details in outobject
    * 5:12 PM 6/13/2024 update to make dynamic, querying for plans serviceplans.serviceplanname -match  EXCHANGE_S_ENTERPRISE','EXCHANGE_S_STANDARD','EXCHANGE_S_DESKLESS (-ne MCOCAP 	Common Area Phone)
    * 12:45 PM 6/21/2022 added cbh expl that rolls up a rgx to use for independant manual tests against 
    * 2:21 PM 3/1/2022 updated CBH
    * 4:27 PM 2/25/2022 init vers
    .DESCRIPTION
    get-ExoMailboxLicenses - Provides a prefab indexed hash of Exchange-Online mailbox-supporting licenses (at least one of which is required to accomodate an EXO Usermailbox - This now dynamically calls Get-MgSubscribedSku and postfilters the ServicePlan list for matches on the $ServicePlanName array (which reflects Exchange mailbox ServicePlanNames). The ServicePlanName array must be manually updated to accomodate MS licensure changes over time).

    This feeds my test-EXOIsLicensed(), and is usable for feeding add-EXOLicense & remove-EXOLicense
    Underlying goal is to *dynamically* pursue the current supported licensure, but calling Get-MgSubscribedSku and postfiltering the ServicePlan list for matches on the $ServicePlanName array 
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
    4. So we can take any give Get-MgSubscribedSku and filter for licenses which include the above, to pick out suitable Exchange-Mailbox-supporting licenses. This filtered list goes in the default $ServicePlanName list


    ## Discussion of the topic - dynamically finding ongoing licenses as they mess with the licenses

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
    PS> $ExMbxLicenses = get-ExoMailboxLicenses ;
    PS> $ExMbxLicenses

        Name                           Value
        ----                           -----
        SPE_F1                         @{SKU=SPE_F1; Label=F1; Notes=Enabled:5681|Consumed:5179|Avail:502|Warn:0|Susp:0}
        SPE_E3                         @{SKU=SPE_E3; Label=E3; Notes=Enabled:5326|Consumed:5298|Avail:28|Warn:0|Susp:0}
        STANDARDPACK                   @{SKU=STANDARDPACK; Label=STANDARDPACK; Notes=Enabled:268|Consumed:241|Avail:27|Warn:0|Susp:0}    

    Simplest demo, unformatted output, that is still fast-readable to find usables
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
    Expanded example with testing of returned object, and demoes use of the returned hash against a AzureADUser spec
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
    .EXAMPLE
    PS> $ExMbxLicenses = get-ExoMailboxLicenses ;
    PS> [regex]$rgxExLics = ('(' + (($ExMbxLicenses.GetEnumerator().name |%{[regex]::escape($_)}) -join '|') + ')') ; 
    Demo pulling the underlying licenses list and building a regex for static use
    .EXAMPLE
    PS> $TenOrg = 'ABC' ; 
    PS> $ExMbxLicenses = get-ExoMailboxLicenses ;
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
            [string[]]$UserRole = @('ESvcCBA','CSvcCBA','SIDCBA','SID'),
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
        #region CONSTANTS_AND_ENVIRO ; #*======v CONSTANTS_AND_ENVIRO v======
        #region ENVIRO_DISCOVER ; #*------v ENVIRO_DISCOVER v------
        push-TLSLatest
        $Verbose = [boolean]($VerbosePreference -eq 'Continue') ; 
        $rPSCmdlet = $PSCmdlet ; # an object that represents the cmdlet or advanced function that's being run. Available on functions w CmdletBinding (& $args will not be available). (Blank on non-CmdletBinding/Non-Adv funcs).
        $rPSScriptRoot = $PSScriptRoot ; # the full path of the executing script's parent directory., PS2: valid only in script modules (.psm1). PS3+:it's valid in all scripts. (Funcs: ParentDir of the file that hosts the func)
        $rPSCommandPath = $PSCommandPath ; # the full path and filename of the script that's being run, or file hosting the funct. Valid in all scripts.
        $rMyInvocation = $MyInvocation ; # populated only for scripts, function, and script blocks.
        # - $MyInvocation.MyCommand.Name returns name of a function, to identify the current command,  name of the current script (pop'd w func name, on Advfuncs)
        # - Ps3+:$MyInvocation.PSScriptRoot : full path to the script that invoked the current command. The value of this property is populated only when the caller is a script (blank on funcs & Advfuncs)
        # - Ps3+:$MyInvocation.PSCommandPath : full path and filename of the script that invoked the current command. The value of this property is populated only when the caller is a script (blank on funcs & Advfuncs)
        #     ** note: above pair contain information about the _invoker or calling script_, not the current script
        $rPSBoundParameters = $PSBoundParameters ; 
        #region PREF_VARI_DUMP ; #*------v PREF_VARI_DUMP v------
        <#$script:prefVaris = @{
            whatifIsPresent = $whatif.IsPresent
            whatifPSBoundParametersContains = $rPSBoundParameters.ContainsKey('WhatIf') ; 
            whatifPSBoundParameters = $rPSBoundParameters['WhatIf'] ;
            WhatIfPreferenceIsPresent = $WhatIfPreference.IsPresent ; # -eq $true
            WhatIfPreferenceValue = $WhatIfPreference;
            WhatIfPreferenceParentScopeValue = (Get-Variable WhatIfPreference -Scope 1).Value ;
            ConfirmPSBoundParametersContains = $rPSBoundParameters.ContainsKey('Confirm') ; 
            ConfirmPSBoundParameters = $rPSBoundParameters['Confirm'];
            ConfirmPreferenceIsPresent = $ConfirmPreference.IsPresent ; # -eq $true
            ConfirmPreferenceValue = $ConfirmPreference ;
            ConfirmPreferenceParentScopeValue = (Get-Variable ConfirmPreference -Scope 1).Value ; 
            VerbosePSBoundParametersContains = $rPSBoundParameters.ContainsKey('Confirm') ; 
            VerbosePSBoundParameters = $rPSBoundParameters['Verbose'] ;
            VerbosePreferenceIsPresent = $VerbosePreference.IsPresent ; # -eq $true
            VerbosePreferenceValue = $VerbosePreference ;
            VerbosePreferenceParentScopeValue = (Get-Variable VerbosePreference -Scope 1).Value;
            VerboseMyInvContains = '-Verbose' -in $rPSBoundParameters.UnboundArguments ; 
            VerbosePSBoundParametersUnboundArgumentContains = '-Verbose' -in $rPSBoundParameters.UnboundArguments 
        } ;
        write-verbose "`n$(($script:prefVaris.GetEnumerator() | Sort-Object Key | Format-Table Key,Value -AutoSize|out-string).trim())`n" ; 
        #>
        #endregion PREF_VARI_DUMP ; #*------^ END PREF_VARI_DUMP ^------
        #region RV_ENVIRO ; #*------v RV_ENVIRO v------
        $pltRvEnv=[ordered]@{
            PSCmdletproxy = $rPSCmdlet ; 
            PSScriptRootproxy = $rPSScriptRoot ; 
            PSCommandPathproxy = $rPSCommandPath ; 
            MyInvocationproxy = $rMyInvocation ;
            PSBoundParametersproxy = $rPSBoundParameters
            verbose = [boolean]($PSBoundParameters['Verbose'] -eq $true) ; 
        } ;
        write-verbose "(Purge no value keys from splat)" ; 
        $mts = $pltRVEnv.GetEnumerator() |?{$_.value -eq $null} ; $mts |%{$pltRVEnv.remove($_.Name)} ; rv mts -ea 0 -whatif:$false -confirm:$false; 
        $smsg = "resolve-EnvironmentTDO w`n$(($pltRVEnv|out-string).trim())" ; 
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        if(get-command resolve-EnvironmentTDO -ea STOP){}ELSE{
            $smsg = "UNABLE TO gcm resolve-EnvironmentTDO!" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
            BREAK ; 
        } ; 
        $rvEnv = resolve-EnvironmentTDO @pltRVEnv ; 
        $smsg = "`$rvEnv returned:`n$(($rvEnv |out-string).trim())" ; 
        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        #endregion RV_ENVIRO ; #*------^ END RV_ENVIRO ^------
        #region NETWORK_INFO ; #*======v NETWORK_INFO v======
        if(get-command resolve-NetworkLocalTDO  -ea STOP){}ELSE{
            $smsg = "UNABLE TO gcm resolve-NetworkLocalTDO !" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
            BREAK ; 
        } ; 
        $netsettings = resolve-NetworkLocalTDO ; 
        if($env:Userdomain){ 
            switch($env:Userdomain){
                'CMW'{
                    #$logon_SID = $CMW_logon_SID 
                }
                'TORO'{
                    #$o365_SIDUpn = $o365_Toroco_SIDUpn ; 
                    #$logon_SID = $TOR_logon_SID ; 
                }
                $env:COMPUTERNAME{
                    $smsg = "%USERDOMAIN% -EQ %COMPUTERNAME%: $($env:computername) => non-domain-connected, likely edge role Ex server!" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                    if($netsettings.Workgroup){
                        $smsg = "WorkgroupName:$($netsettings.Workgroup)" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;                    
                    } ; 
                } ; 
                default{
                    $smsg = "$($env:userdomain):UNRECOGIZED/UNCONFIGURED USER DOMAIN STRING!" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                    THROW $SMSG 
                    BREAK ; 
                }
            } ; 
        } ;  # $env:Userdomain-E
        #endregion NETWORK_INFO ; #*======^ END NETWORK_INFO ^======
        #region OS_INFO ; #*------v OS_INFO v------
        <# os detect, covers Server 2016, 2008 R2, Windows 10, 11
        if (get-command get-ciminstance -ea 0) {$OS = (Get-ciminstance -class Win32_OperatingSystem)} else {$Os = Get-WMIObject -class Win32_OperatingSystem } ;
        #$isWorkstationOS = $isServerOS = $isW2010 = $isW2011 = $isS2016 = $isS2008R2 = $false ;
        write-host "Detected:`$Os.Name:$($OS.name)`n`$Os.Version:$($Os.Version)" ;
        if ($OS.name -match 'Microsoft\sWindows\sServer') {
            $isServerOS = $true ;
            if ($os.name -match 'Microsoft\sWindows\sServer\s2016'){$isS2016 = $true ;} ;
            if ($os.name -match 'Microsoft\sWindows\sServer\s2008\sR2') { $isS2008R2 = $true ; } ;
        } else { 
            if ($os.name -match '^Microsoft\sWindows\s11') {
                $isWorkstationOS = $true ;
                if ($os.name -match 'Microsoft\sWindows\s11') { $isW2011 = $true ; } ;
            } elseif ($os.name -match '^Microsoft\sWindows\s10') {
                $isWorkstationOS = $true ; $isW2010 = $true
            } else {
                $isWorkstationOS = $true ;
            } ;         
        } ; 
        #>
        #endregion OS_INFO ; #*------^ END OS_INFO ^------
        #region TEST_EXOPLOCAL ; #*------v TEST_EXOPLOCAL v------
        if(get-command test-LocalExchangeInfoTDO -ea STOP){}ELSE{
            $smsg = "UNABLE TO gcm test-LocalExchangeInfoTDO !" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
            BREAK ; 
        } ; 
        $lclExOP = test-LocalExchangeInfoTDO ; 
        write-verbose "Expand returned NoteProperty properties into matching local variables" ; 
        if($host.version.major -gt 2){
            $lclExOP.PsObject.Properties | ?{$_.membertype -eq 'NoteProperty'} | foreach-object{set-variable -name $_.name -value $_.value -verbose -whatif:$false -Confirm:$false ;} ;
        }else{
            write-verbose "Psv2 lacks the above expansion capability; just create simpler variable set" ; 
            $ExVers = $lclExOP.ExVers ; $isLocalExchangeServer = $lclExOP.isLocalExchangeServer ; $IsEdgeTransport = $lclExOP.IsEdgeTransport ;
        } ;
        #
        #endregion TEST_EXOPLOCAL ; #*------^ END TEST_EXOPLOCAL ^------

        <#
        #region PsParams ; #*------v PSPARAMS v------
        $PSParameters = New-Object -TypeName PSObject -Property $rPSBoundParameters ;
        # DIFFERENCES $PSParameters vs $PSBoundParameters:
        # - $PSBoundParameters: System.Management.Automation.PSBoundParametersDictionary (native obj)
        # test/access: ($PSBoundParameters['Verbose'] -eq $true) ; $PSBoundParameters.ContainsKey('Referrer') #hash syntax
        # CAN use as a @PSBoundParameters splat to push through (make sure populated, can fail if wrong type of wrapping code)
        # - $PSParameters: System.Management.Automation.PSCustomObject (created obj)
        # test/access: ($PSParameters.verbose -eq $true) ; $PSParameters.psobject.Properties.name -contains 'SenderAddress' ; # cobj syntax
        # CANNOT use as a @splat to push through (it's a cobj)
        write-verbose "`$rPSBoundParameters:`n$(($rPSBoundParameters|out-string).trim())" ;
        # pre psv2, no $rPSBoundParameters autovari to check, so back them out:
        #>
        <# recycling $rPSBoundParameters into @splat calls: (can't use $psParams, it's a cobj, not a hash!)
        # rgx for filtering $rPSBoundParameters for params to pass on in recursive calls (excludes keys matching below)
        $rgxBoundParamsExcl = '^(Name|RawOutput|Server|Referrer)$' ; 
        if($rPSBoundParameters){
                $pltRvSPFRec = [ordered]@{} ;
                # add the specific Name for this call, and Server spec (which defaults, is generally not 
                $pltRvSPFRec.add('Name',"$RedirectRecord" ) ;
                $pltRvSPFRec.add('Referrer',$Name) ; 
                $pltRvSPFRec.add('Server',$Server ) ;
                $rPSBoundParameters.GetEnumerator() | ?{ $_.key -notmatch $rgxBoundParamsExcl} | foreach-object { $pltRvSPFRec.add($_.key,$_.value)  } ;
                write-host "Resolve-SPFRecord w`n$(($pltRvSPFRec|out-string).trim())" ;
                Resolve-SPFRecord @pltRvSPFRec  | write-output ;
        } else {
            $smsg = "unpopulated `$rPSBoundParameters!" ;
            write-warning $smsg ;
            throw $smsg ;
        };     
        #>
        #endregion PsParams ; #*------^ END PSPARAMS ^------    
        #endregion ENVIRO_DISCOVER ; #*------^ END ENVIRO_DISCOVER ^------

        #region COMMON_CONSTANTS ; #*------v COMMON_CONSTANTS v------
    
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
        if(-not $XOConnectionUri ){$XOConnectionUri = 'https://outlook.office365.com'} ; 
        if(-not $SCConnectionUri){$SCConnectionUri = 'https://ps.compliance.protection.outlook.com'} ; 
        if(-not $XODefaultPrefix){$XODefaultPrefix = 'xo' };
        if(-not $SCDefaultPrefix){$SCDefaultPrefix = 'sc' };
        #$rgxADDistNameGAT = ",$(($TORMeta.UnreplicatedOU -split ',' | select -skip 1 ) -join ',')" 
        #$rgxADDistNameAT = ",$(($TORMeta.UnreplicatedOU -split ',' | select -skip 2 ) -join ',')"

        write-verbose "Coerce configured but blank Resultsize to Unlimited" ; 
        if(get-variable -name resultsize -ea 0){
            if( ($null -eq $ResultSize) -OR ('' -eq $ResultSize) ){$ResultSize = 'unlimited' }
            elseif($Resultsize -is [int]){} else {throw "Resultsize must be an integer or the string 'unlimited' (or blank)"} ;
        } ; 
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
        # TenOrg or other looped-specific PassStatus (auto supported by 7pswlt)
        #New-Variable -Name PassStatus_$($tenorg) -scope Script -Value $null ;
        [array]$SmtpAttachment = $null ;
        #write-verbose "start-Timer:Master" ; 
        $swM = [Diagnostics.Stopwatch]::StartNew() ;
        # $ByPassLocalExchangeServerTest = $true # rough in, code exists below for exempting service/regkey testing on this variable status. Not yet implemented beyond the exemption code, ported in from orig source.
        #endregion COMMON_CONSTANTS ; #*------^ END COMMON_CONSTANTS ^------
              
        #region LOCAL_CONSTANTS ; #*------v LOCAL_CONSTANTS v------

        # BELOW TRIGGERS/DRIVES TEST_MODS: array of: "[modname];[modDLUrl,or pscmdline install]"    
        $tDepModules = @() ;
        $useVerbCore = $true ; 
        if($useVerbCore){
            $tDepModules += @('verb-logging;localRepo;write-log') ; #start-log; write-log ;
            $tDepModules += @('verb-io;localRepo;resolve-EnvironmentTDO') ; #resolve-EnvironmentTDO
            $tDepModules += @('verb-Network;localRepo;resolve-NetworkLocalTDO') ; #resolve-NetworkLocalTDO; Send-EmailNotif
        } ;
        <# NOTE: Svc modules are tested as needed by connect-O365Servicees() & connect-OPServices()
        if($useEXO){$tDepModules += @("ExchangeOnlineManagement;https://www.powershellgallery.com/packages/ExchangeOnlineManagement/;Get-xoOrganizationConfig",'verb-exo;localRepo;connect-exo')} ;
        if($UseMSOL){$tDepModules += @("MSOnline;https://www.powershellgallery.com/packages/MSOnline/;Get-MsolDomain")} ;
        if($UseAAD){$tDepModules += @("AzureAD;https://www.powershellgallery.com/packages/AzureAD/;Get-AzureADTenantDetail")} ;
        if($UseExOP){$tDepModules += @('verb-Ex2010;localRepo;Connect-Ex2010')} ;
        if($UseMG){$tDepModules += @("Microsoft.Graph.Authentication;https://www.powershellgallery.com/packages/Microsoft.Graph/;Get-MgOrganization")} ;
        if($UseOPAD){$tDepModules += @("ActiveDirectory;get-windowscapability -name RSAT* -Online | ?{$_.name -match 'Rsat\.ActiveDirectory'} | %{Add-WindowsCapability -online -name $_.name};Get-ADDomain")} ;
        #>

        $prpMGConnDeleg = 'Account','ClientId','TenantId','AuthType','ContextScope' ; 
        $prpMGConnCBA = 'CertificateSubjectName','CertificateThumbprint','Certificate' ; 
        $prpMGConnRet = $($prpMGConnDeleg;$prpMGConnCBA) ; 

        #endregion LOCAL_CONSTANTS ; #*------^ END LOCAL_CONSTANTS ^------  
          
        #region ENCODED_CONTANTS ; #*------v ENCODED_CONTANTS v------
        # ENCODED CONsTANTS & SUPPORT FUNCTIONS:
        #region 2B4 ; #*------v 2B4 v------
        if(-not (get-command 2b4 -ea 0)){function 2b4{[CmdletBinding()][Alias('convertTo-Base64String')] PARAM([Parameter(ValueFromPipeline=$true)][string[]]$str) ; PROCESS{$str|%{[Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($_))}  };} ; } ; 
        #endregion 2B4 ; #*------^ END 2B4 ^------
        #region 2B4C ; #*------v 2B4C v------
        # comma-quoted return
        if(-not (get-command 2b4c -ea 0)){function 2b4c{ [CmdletBinding()][Alias('convertto-Base64StringCommaQuoted')] PARAM([Parameter(ValueFromPipeline=$true)][string[]]$str) ;BEGIN{$outs = @()} PROCESS{[array]$outs += $str | %{[Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($_))} ; } END {'"' + $(($outs) -join '","') + '"' | out-string | set-clipboard } ; } ; } ; 
        #endregion 2B4C ; #*------^ END 2B4C ^------
        #region FB4 ; #*------v FB4 v------
        # DEMO: $SitesNameList = 'THluZGFsZQ==','U3BlbGxicm9vaw==','QWRlbGFpZGU=' | fb4 ;
        if(-not (get-command fb4 -ea 0)){function fb4{[CmdletBinding()][Alias('convertFrom-Base64String')] PARAM([Parameter(ValueFromPipeline=$true)][string[]]$str) ; PROCESS{$str | %{ [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($_)) }; } ; } ; }; 
        #endregion FB4 ; #*------^ END FB4 ^------
        # FOLLOWING CONSTANTS ARE USED FOR DEPENDANCY-LESS CONNECTIONS
        if(-not $o365_Toroco_SIDUpn){$o365_Toroco_SIDUpn = 'cy10b2RkLmthZHJpZUB0b3JvLmNvbQ==' | fb4 } ;
        $o365_SIDUpn = $o365_Toroco_SIDUpn ; 
        switch($env:Userdomain){
            'CMW'{
                if(-not $CMW_logon_SID){$CMW_logon_SID = 'Q01XXGQtdG9kZC5rYWRyaWU=' | fb4 } ; 
                $logon_SID = $CMW_logon_SID ; 
            }
            'TORO'{
                if(-not $TOR_logon_SID){$TOR_logon_SID = 'VE9ST1xrYWRyaXRzcw==' | fb4 } ; 
                $logon_SID = $TOR_logon_SID ; 
            }
            $env:COMPUTERNAME{
                $smsg = "%USERDOMAIN% -EQ %COMPUTERNAME%: $($env:computername) => non-domain-connected, likely edge role Ex server!" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                if($WorkgroupName = (Get-WmiObject -Class Win32_ComputerSystem).Workgroup){
                    $smsg = "WorkgroupName:$($WorkgroupName)" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                }
                if(($isLocalExchangeServer = (Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\Setup')) -or (
                        $isLocalExchangeServer = (Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup')) -or
                            $ByPassLocalExchangeServerTest){
                            $smsg = "We are on Exchange Server"
                            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                            $IsEdgeTransport = $false
                            if((Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\EdgeTransportRole') -or (Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\EdgeTransportRole')){
                                $smsg = "We are on Exchange Edge Transport Server"
                                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                                $IsEdgeTransport = $true
                            } ; 
                } else {
                    $isLocalExchangeServer = $false 
                    $IsEdgeTransport = $false ;
                } ;
            } ; 
            default{
                $smsg = "$($env:userdomain):UNRECOGIZED/UNCONFIGURED USER DOMAIN STRING!" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                THROW $SMSG 
                BREAK ; 
            }
        } ; 
        #endregion ENCODED_CONTANTS ; #*------^ END ENCODED_CONTANTS ^------
    
        #endregion CONSTANTS_AND_ENVIRO ; #*======^ CONSTANTS_AND_ENVIRO ^======

        #region SUBMAIN ; #*======v SUB MAIN v======
    
        #region BANNER ; #*------v BANNER v------
        $sBnr="#*======v $(${CmdletName}): v======" ;
        $smsg = $sBnr ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level H1 } #Error|Warn|Debug
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        #endregion BANNER ; #*------^ END BANNER ^------

        <# prior specs: 
           $pltCco365Svcs=[ordered]@{
                # environment parameters:
                EnvSummary = $rvEnv ; 
                NetSummary = $netsettings ; 
                # service choices
                useEXO = $true ;
                useSC = $false ; 
                UseMSOL = $false ;
                UseAAD = $false ; # M$ is actively blocking all AAD access now: Message: Access blocked to AAD Graph API for this application. https://aka.ms/AzureADGraphMigration.
                UseMG = $true ;
                # Service Connection parameters
                TenOrg = $TenOrg ; # $global:o365_TenOrgDefault ; 
                Credential = $Credential ;
                AdminAccount = $AdminAccount ; 
                #[ValidateSet("SID","CSID","UID","B2BI","CSVC","ESVC","LSVC","ESvcCBA","CSvcCBA","SIDCBA")]
                UserRole = $UserRole ; # @('SID','CSVC') ;
                # svcAcct use: @('ESvcCBA','CSvcCBA','SIDCBA')
                silent = $silent ;
                MGPermissionsScope = $MGPermissionsScope ;
                MGCmdlets = $MGCmdlets ;
            } ;
            $pltCcOPSvcs=[ordered]@{
                # environment parameters:
                EnvSummary = $rvEnv ;
                NetSummary = $netsettings ;
                XoPSummary = $lclExOP ;
                # service choices
                UseExOP = $true ;
                useForestWide = $true ;
                useExopNoDep = $false ;
                ExopVers = 'Ex2010' ;
                UseOPAD = $true ;
                useExOPVers = $useExOPVers; # 'Ex2010' ;
                # Service Connection parameters
                TenOrg = $TenOrg ; # $global:o365_TenOrgDefault ;
                Credential = $Credential ;
                #[ValidateSet("SID","ESVC","LSVC")]
                #UserRole = $UserRole ; # @('SID','ESVC') ;
                # if inheriting same $userrole param/default, that was already used for cloud conn, filter out the op unsupported CBA roles
                # exclude csvc as well, go with filter on the supported ValidateSet from get-HybridOPCredentials: ESVC|LSVC|SID
                UserRole = $UserRole |?{$_ -match '(ESVC|LSVC|SID)' -AND $_ -notmatch 'CBA'} ; 
                # svcAcct use: @('ESvcCBA','CSvcCBA','SIDCBA')
                silent = $silent ;
            } ;
        #>

        #region SERVICE_CONNECTIONS #*======v SERVICE_CONNECTIONS v======
    
    #region BROAD_SVC_CONTROL_VARIS ; #*======v BROAD_SVC_CONTROL_VARIS  v======   
    $useO365 = $true ; 
    $useOP = $true ;     
    # (config individual svcs in each block)
    #endregion BROAD_SVC_CONTROL_VARIS ; #*======^ END BROAD_SVC_CONTROL_VARIS ^======

    #region CALL_CONNECT_O365SERVICES ; #*======v CALL_CONNECT_O365SERVICES v======
    #$useO365 = $true ; 
    if($useO365){
        $pltCco365Svcs=[ordered]@{
                # environment parameters:
                EnvSummary = $rvEnv ; 
                NetSummary = $netsettings ; 
                # service choices
                useEXO = $true ;
                useSC = $false ; 
                UseMSOL = $false ;
                UseAAD = $false ; # M$ is actively blocking all AAD access now: Message: Access blocked to AAD Graph API for this application. https://aka.ms/AzureADGraphMigration.
                UseMG = $true ;
                # Service Connection parameters
                TenOrg = $TenOrg ; # $global:o365_TenOrgDefault ; 
                Credential = $Credential ;
                AdminAccount = $AdminAccount ; 
                #[ValidateSet("SID","CSID","UID","B2BI","CSVC","ESVC","LSVC","ESvcCBA","CSvcCBA","SIDCBA")]
                UserRole = $UserRole ; # @('SID','CSVC') ;
                # svcAcct use: @('ESvcCBA','CSvcCBA','SIDCBA')
                silent = $silent ;
                MGPermissionsScope = $MGPermissionsScope ;
                MGCmdlets = $MGCmdlets ;
            } ;
        write-verbose "(Purge no value keys from splat)" ; 
        $mts = $pltCco365Svcs.GetEnumerator() |?{$_.value -eq $null} ; $mts |%{$pltCco365Svcs.remove($_.Name)} ; rv mts -ea 0 ; 
        if((get-command connect-O365Services -EA STOP).parameters.ContainsKey('whatif')){
            $pltCco365SvcsnDSR.add('whatif',$($whatif))
        } ; 
        # add rertry on fail, up to $DoRetries
        $Exit = 0 ; # zero out $exit each new cmd try/retried
        # do loop until up to 4 retries...
        Do {
            $smsg = "connect-O365Services w`n$(($pltCco365Svcs|out-string).trim())" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            $ret_ccSO365 = connect-O365Services @pltCco365Svcs ; 
            #region CONFIRM_CCEXORETURN ; #*------v CONFIRM_CCEXORETURN v------
            # matches each: $plt.useXXX:$true to matching returned $ret.hasXXX:$true 
            $vplt = $pltCco365Svcs ; $vret = 'ret_ccSO365' ; $ACtionCommand = 'connect-O365Services' ; $vtests = @() ; $vFailMsgs = @()  ; 
            $vplt.GetEnumerator() |?{$_.key -match '^use' -ANd $_.value -match $true} | foreach-object{
                $pltkey = $_ ;
                $smsg = "$(($pltkey | ft -HideTableHeaders name,value|out-string).trim())" ; 
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                $tprop = $pltkey.name -replace '^use','has';
                if($rProp = (gv $vret).Value.psobject.properties | ?{$_.name -match $tprop}){
                    $smsg = "$(($rprop | ft -HideTableHeaders name,value |out-string).trim())" ; 
                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                    if($rprop.Value -eq $pltkey.value){
                        $vtests += $true ; 
                        $smsg = "Validated: $($pltKey.name):$($pltKey.value) => $($rprop.name):$($rprop.value)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Success } 
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                    } else {
                        $smsg = "NOT VALIDATED: $($pltKey.name):$($pltKey.value) => $($rprop.name):$($rprop.value)" ;
                        $vtests += $false ; 
                        $vFailMsgs += "`n$($smsg)" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                    };
                } else{
                    $smsg = "Unable to locate: $($pltKey.name):$($pltKey.value) to any matching $($rprop.name)!)" ;
                    $smsg = "" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                } ; 
            } ; 
            if($vtests -notcontains $false){
                $smsg = "==> $($ACtionCommand): confirmed specified connections *all* successful " ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Success } 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                # populuate the $pltRXO.credential
                if($ret_ccSO365.CredentialO365){
                    $pltRXO = [ordered]@{
                        Credential = $ret_ccSO365.CredentialO365 ;
                        verbose = $($VerbosePreference -eq "Continue")  ;
                    } ;
                }else{
                    $smsg = "Unpopulated returnd:connect-O365Services.CredentialO365!" ; 
                    $smsg += "`nUNABLE TO POPULATE `$pltRXO!" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                }
                $Exit = $DoRetries ;
            } else {
                $smsg = "==> $($ACtionCommand): FAILED SOME SPECIFIED CONNECTIONS" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                $smsg = "MISSING SOME KEY CONNECTIONS. DO YOU WANT TO IGNORE, AND CONTINUE WITH CONNECTED SERVICES?" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                $Exit ++ ;
                $smsg = "Try #: $Exit" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Warn } 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                if($Exit -eq $DoRetries){
                    $smsg = "Unable to exec cmd!"; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                    #-=-=-=-=-=-=-=-=
                    $sdEmail.SMTPSubj = "FAIL Rpt:$($ScriptBaseName):$(get-date -format 'yyyyMMdd-HHmmtt')"
                    $sdEmail.SmtpBody = "`n===Processing Summary:" ;
                    if($vFailMsgs){
                        $sdEmail.SmtpBody += "`n$(($vFailMsgs|out-string).trim())" ; 
                    } ; 
                    $sdEmail.SmtpBody += "`n" ;
                    if($SmtpAttachment){
                        $sdEmail.SmtpAttachment = $SmtpAttachment
                        $sdEmail.smtpBody +="`n(Logs Attached)" ;
                    };
                    $sdEmail.SmtpBody += "Pass Completed $([System.DateTime]::Now)" ;
                    $smsg = "Send-EmailNotif w`n$(($sdEmail|out-string).trim())" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    Send-EmailNotif @sdEmail ;
                    $bRet=Read-Host "Enter YYY to continue. Anything else will exit"  ;
                    if ($bRet.ToUpper() -eq "YYY") {
                        $smsg = "(Moving on), WITH THE FOLLOW PARTIAL CONNECTION STATUS" ;
                        $smsg += "`n`n$(($ret_CcOPSvcs|out-string).trim())" ; 
                        write-host -foregroundcolor green $smsg  ;
                    } else {
                        throw $smsg ; 
                        break ; #exit 1
                    } ;  
                } ;        
            } ; 
            #endregion CONFIRM_CCEXORETURN ; #*------^ END CONFIRM_CCEXORETURN ^------
        } Until ($Exit -eq $DoRetries) ; 
    } ; #  useO365-E
    #endregion CALL_CONNECT_O365SERVICES ; #*======^ END CALL_CONNECT_O365SERVICES ^======

    #region TEST_EXO_CONN ; #*------v TEST_EXO_CONN v------
    # ALT: simplified verify EXO conn: ALT to full CONNECT_O365SERVICES block - USE ONE OR THE OTHER!
    $useEXO = $FALSE ; 
    $useSC = $false ; 
    if(-not $XOConnectionUri ){$XOConnectionUri = 'https://outlook.office365.com'} ;
    if(-not $SCConnectionUri){$SCConnectionUri = 'https://ps.compliance.protection.outlook.com'} ;
    $EXOtestCmdlet = 'Get-xoOrganizationConfig' ; 
    if($useEXO){
        if(gcm $EXOtestCmdlet -ea 0){
            $conns = Get-ConnectionInformation -ea STOP  ; 
            $hasEXO = $hasSC = $false ; 
            #if($conns | %{$_ | ?{$_.ConnectionUri -eq 'https://outlook.office365.com' -AND $_.State -eq 'Connected' -AND $_.TokenStatus -eq 'Active'}}){
            $conns | %{
                if($_ | ?{$_.ConnectionUri -eq $XOConnectionUri}){$hasEXO = $true } ; 
                if($_ | ?{$_.ConnectionUri -eq $SCConnectionUri}){$hasSC = $true } ; 
            }
            if($useEXO -AND $hasEXO){
                write-verbose "EXO ConnectionURI present" ; 
            }elseif(-not $useEXO){}else{
                $smsg = "No Active EXO connection: Run - Connect-ExchangeOnline -Prefix xo -  before running this script!" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                BREAK ; 
            } ; 
            if($useSC -AND $hasSC){
                write-verbose "SCI ConnectionURI present" ; 
            }elseif(-not $useSC){}else{
                $smsg = "No Active SC connection: Run - Connect-IPPSSession -Prefix SC -  before running this script!" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                BREAK ; 
            } ; 
        }else {
            $smsg = "Missing gcm get-xoMailboxFolderStatistics: ExchangeOnlineManagement module *not* loaded!" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
            BREAK ; 
        } ;    
    }else{
        $smsg = "useEXO:$($useEXO): skipping EXO tests" ; 
        if($VerbosePreference -eq "Continue"){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
    } 
    #endregion TEST_EXO_CONN ; #*------^ END TEST_EXO_CONN ^------

    #region CALL_CONNECT_OPSERVICES ; #*======v CALL_CONNECT_OPSERVICES v======
    #$useOP = $false ; 
    if($useOP){
        $pltCcOPSvcs=[ordered]@{
            # environment parameters:
            EnvSummary = $rvEnv ;
            NetSummary = $netsettings ;
            XoPSummary = $lclExOP ;
            # service choices
            UseExOP = $true ;
            useForestWide = $true ;
            useExopNoDep = $false ;
            ExopVers = 'Ex2010' ;
            UseOPAD = $true ;
            useExOPVers = $useExOPVers; # 'Ex2010' ;
            # Service Connection parameters
            TenOrg = $TenOrg ; # $global:o365_TenOrgDefault ;
            Credential = $Credential ;
            #[ValidateSet("SID","ESVC","LSVC")]
            #UserRole = $UserRole ; # @('SID','ESVC') ;
            # if inheriting same $userrole param/default, that was already used for cloud conn, filter out the op unsupported CBA roles
            # exclude csvc as well, go with filter on the supported ValidateSet from get-HybridOPCredentials: ESVC|LSVC|SID
            UserRole = $UserRole |?{$_ -match '(ESVC|LSVC|SID)' -AND $_ -notmatch 'CBA'} ; 
            # svcAcct use: @('ESvcCBA','CSvcCBA','SIDCBA')
            silent = $silent ;
        } ;

        write-verbose "(Purge no value keys from splat)" ;
        $mts = $pltCcOPSvcs.GetEnumerator() |?{$_.value -eq $null} ; $mts |%{$pltCcOPSvcs.remove($_.Name)} ; rv mts -ea 0 ;
        if((get-command connect-OPServices -EA STOP).parameters.ContainsKey('whatif')){
            $pltCcOPSvcsnDSR.add('whatif',$($whatif))
        } ;
        $smsg = "connect-OPServices w`n$(($pltCcOPSvcs|out-string).trim())" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        $ret_CcOPSvcs = connect-OPServices @pltCcOPSvcs ;

        #region CONFIRM_CCOPRETURN ; #*------v CONFIRM_CCOPRETURN v------
        # matches each: $plt.useXXX:$true to matching returned $ret.hasXXX:$true
        $vplt = $pltCcOPSvcs ; $vret = 'ret_CcOPSvcs' ;  ; $ACtionCommand = 'connect-OPServices' ; 
        $vplt.GetEnumerator() |?{$_.key -match '^use' -ANd $_.value -match $true} | foreach-object{
            $pltkey = $_ ;
            $smsg = "$(($pltkey | ft -HideTableHeaders name,value|out-string).trim())" ; 
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            $vtests = @() ;  $vFailMsgs = @()  ; 
            $tprop = $pltkey.name -replace '^use','has';
            if($rProp = (gv $vret).Value.psobject.properties | ?{$_.name -match $tprop}){
                $smsg = "$(($rprop | ft -HideTableHeaders name,value |out-string).trim())" ; 
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                if($rprop.Value -eq $pltkey.value){
                    $vtests += $true ; 
                    $smsg = "Validated: $($pltKey.name):$($pltKey.value) => $($rprop.name):$($rprop.value)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Success } 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                } else {
                    $smsg = "NOT VALIDATED: $($pltKey.name):$($pltKey.value) => $($rprop.name):$($rprop.value)" ;
                    $vtests += $false ; 
                    $vFailMsgs += "`n$($smsg)" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                };
            } else{
                $smsg = "Unable to locate: $($pltKey.name):$($pltKey.value) to any matching $($rprop.name)!)" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
            } ; 
        } ; 
        if($useOP -AND $vtests -notcontains $false){
            $smsg = "==> $($ACtionCommand): confirmed specified connections *all* successful " ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Success } 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
            # 11:01 AM 1/27/2026 populate followon:            
            if($ret_CcOPSvcs.CredentialOP){
                $pltRx10 = [ordered]@{
                    Credential = $ret_CcOPSvcs.CredentialOP ;
                    verbose = $($VerbosePreference -eq "Continue")  ;
                } ;
            }else{
                $smsg = "Unpopulated returned:connect-OPServices.CredentialOP!" ;
                $smsg += "`nUNABLE TO POPULATE `$pltRx10!" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            }
        }elseif($vtests -contains $false -AND (get-variable ret_CcOPSvcs) -AND (gv -name "$($tenorg)meta").value.o365_opdomain.split('.')[0].toupper() -ne $env:userdomain){
            $smsg = "==> $($ACtionCommand): FAILED SOME SPECIFIED CONNECTIONS" ; 
            $smsg += "`nCROSS-ORG ONPREM CONNECTION: ATTEMPTING TO CONNECT TO ONPREM '$((gv -name "$($tenorg)meta").value.o365_Prefix)' $((gv -name "$($tenorg)meta").value.o365_opdomain.split('.')[0].toupper()) domain, FROM $($env:userdomain)!" ;
            $smsg += "`nEXPECTED ERROR, SKIPPING ONPREM ACCESS STEPS (force `$useOP:$false)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
            $useOP = $false ; 
        }elseif(-not $useOP -AND -not (get-variable ret_CcOPSvcs)){
            $smsg = "-useOP: $($useOP), skipped connect-OPServices" ; 
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        } else {
            $smsg = "==> $($ACtionCommand): FAILED SOME SPECIFIED CONNECTIONS" ; 
            $smsg += "`n`$ret_CcOPSvcs:`n$(($ret_CcOPSvcs|out-string).trim())" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
            $sdEmail.SMTPSubj = "FAIL Rpt:$($ScriptBaseName):$(get-date -format 'yyyyMMdd-HHmmtt')"
            $sdEmail.SmtpBody = "`n===Processing Summary:" ;
            if($vFailMsgs){
                $sdEmail.SmtpBody += "`n$(($vFailMsgs|out-string).trim())" ; 
            } ; 
            $sdEmail.SmtpBody += "`n" ;
            if($SmtpAttachment){
                $sdEmail.SmtpAttachment = $SmtpAttachment
                $sdEmail.smtpBody +="`n(Logs Attached)" ;
            };
            $sdEmail.SmtpBody += "Pass Completed $([System.DateTime]::Now)" ;
            $smsg = "Send-EmailNotif w`n$(($sdEmail|out-string).trim())" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            Send-EmailNotif @sdEmail ;
            throw $smsg ; 
            BREAK ; 
        } ; 
        #endregion CONFIRM_CCOPRETURN ; #*------^ END CONFIRM_CCOPRETURN ^------
            
        #region CONFIRM_OPFORESTWIDE ; #*------v CONFIRM_OPFORESTWIDE v------    
        if($useOP -AND $pltCcOPSvcs.useForestWide -AND $ret_CcOPSvcs.hasForestWide -AND $ret_CcOPSvcs.AdGcFwide){
            $smsg = "==> $($ACtionCommand): confirmed has BOTH .hasForestWide & .AdGcFwide ($($ret_CcOPSvcs.AdGcFwide))" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Success } 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success        
        }elseif($pltCcOPSvcs.useForestWide -AND (get-variable ret_CcOPSvcs) -AND (gv -name "$($tenorg)meta").value.o365_opdomain.split('.')[0].toupper() -ne $env:userdomain){
            $smsg = "`nCROSS-ORG ONPREM CONNECTION: ATTEMPTING TO CONNECT TO ONPREM '$((gv -name "$($tenorg)meta").value.o365_Prefix)' $((gv -name "$($tenorg)meta").value.o365_opdomain.split('.')[0].toupper()) domain, FROM $($env:userdomain)!" ;
            $smsg += "`nEXPECTED ERROR, SKIPPING ONPREM FORESTWIDE SPEC" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
            $useOP = $false ; 
        }elseif($useOP -AND $pltCcOPSvcs.useForestWide -AND -NOT $ret_CcOPSvcs.hasForestWide){
            $smsg = "==> $($ACtionCommand): MISSING CRITICAL FORESTWIDE SUPPORT COMPONENT:" ; 
            if(-not $ret_CcOPSvcs.hasForestWide){
                $smsg += "`n----->$($ACtionCommand): MISSING .hasForestWide (Set-AdServerSettings -ViewEntireForest `$True) " ; 
            } ; 
            if(-not $ret_CcOPSvcs.AdGcFwide){
                $smsg += "`n----->$($ACtionCommand): MISSING .AdGcFwide GC!:`n((Get-ADDomainController -Discover -Service GlobalCatalog).hostname):326) " ; 
            } ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
            $smsg = "MISSING SOME KEY CONNECTIONS. DO YOU WANT TO IGNORE, AND CONTINUE WITH CONNECTED SERVICES?" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
            $bRet=Read-Host "Enter YYY to continue. Anything else will exit"  ;
            if ($bRet.ToUpper() -eq "YYY") {
                $smsg = "(Moving on), WITH THE FOLLOW PARTIAL CONNECTION STATUS" ;
                $smsg += "`n`n$(($ret_CcOPSvcs|out-string).trim())" ; 
                write-host -foregroundcolor green $smsg  ;
            } else {
                throw $smsg ; 
                break ; #exit 1
            } ;         
        }; 
        #endregion CONFIRM_OPFORESTWIDE ; #*------^ END CONFIRM_OPFORESTWIDE ^------
    } ; 
    #endregion CALL_CONNECT_OPSERVICES ; #*======^ END CALL_CONNECT_OPSERVICES ^======
    
    #endregion SERVICE_CONNECTIONS #*======^ END SERVICE_CONNECTIONS ^======
    
    <# Service Conditional if/thens: Tests above should BREAK on any fail, but these are for critical dependancy calls
    # o365 calls
    if($ret_ccSO365.hasAAD){ }else{
        $smsg = "`$ret_ccSO365.hasAAD:$($ret_ccSO365.hasAAD): MISSING DEPENDANT CONNECTION!" ; 
        $smsg += "`n(SKIPPING EXECUTION OF DEPENDANT COMMANDS)" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
    } ;  ;
    if($ret_ccSO365.hasEXO){ }else{
        $smsg = "`$ret_ccSO365.hasEXO:$($ret_ccSO365.hasEXO): MISSING DEPENDANT CONNECTION!" ; 
        $smsg += "`n(SKIPPING EXECUTION OF DEPENDANT COMMANDS)" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
    } ; 
    if($ret_ccO365S.hasSC){ }else{
        $smsg = "`$ret_ccO365S.hasSC:$($ret_ccO365S.hasSC): MISSING DEPENDANT CONNECTION!" ; 
        $smsg += "`n(SKIPPING EXECUTION OF DEPENDANT COMMANDS)" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
    } ;         
    if($ret_ccSO365.hasMG){ }else{
        $smsg = "`$ret_ccSO365.hasMG:$($ret_ccSO365.hasMG): MISSING DEPENDANT CONNECTION!" ; 
        $smsg += "`n(SKIPPING EXECUTION OF DEPENDANT COMMANDS)" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
    } ; 
    if($ret_ccSO365.hasMSOL){ }else{
        $smsg = "`$ret_ccSO365.hasMSOL:$($ret_ccSO365.hasMSOL): MISSING DEPENDANT CONNECTION!" ; 
        $smsg += "`n(SKIPPING EXECUTION OF DEPENDANT COMMANDS)" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
    } ; 
    # XOP calls
    if($ret_ccOPSvcs.UseExOP){ }else{
        $smsg = "`$ret_ccOPSvcs.UseExOP:$($ret_ccOPSvcs.UseExOP): MISSING DEPENDANT CONNECTION!" ; 
        $smsg += "`n(SKIPPING EXECUTION OF DEPENDANT COMMANDS)" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
    } ; 
    if($ret_ccOPSvcs.UseOPAD){ }else{
        $smsg = "`$ret_ccOPSvcs.UseOPAD:$($ret_ccOPSvcs.UseOPAD): MISSING DEPENDANT CONNECTION!" ; 
        $smsg += "`n(SKIPPING EXECUTION OF DEPENDANT COMMANDS)" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
    } ; 
    #>

    # connect to ExOP X10
    <#
    if($pltRX10){
        ReConnect-Ex2010 @pltRX10 ;
    } else { Reconnect-Ex2010 ; } ;
    #>


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
        # 1:48 PM 1/7/2026 MG vers
        $propsMGL = 'SkuId',  'SkuPartNumber',  @{name='Enabled';Expression={$_.PrepaidUnits.enabled }},  
            @{name='Consumed';Expression={$_.ConsumedUnits} }, @{name='Available';Expression={$_.PrepaidUnits.enabled - $_.ConsumedUnits} }, 
            @{name='Warning';Expression={$_.PrepaidUnits.warning} }, @{name='Suspended';Expression={$_.PrepaidUnits.suspended} } ;

        $rgxExSvcPlans = ('(' + (($ServicePlanName |%{[regex]::escape($_)}) -join '|') + ')') ; 

        $ExMbxLicenses = @() ; 
        foreach($SPName in $ServicePlanName){
            $smsg = $sBnrS="`n#*------v PROCESSING $($spname): v------" ; 
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            
            TRY{
                #if($thisPlan = Get-AzureADSubscribedSku | ?{$_.serviceplans.serviceplanname -match $SPName} | select $propsAADL | ?{$_.SkuPartNumber -notmatch $rgxbannedSPN}){
                if($thisPlan = Get-MgSubscribedSku | ?{$_.serviceplans.serviceplanname -match $SPName} | select $propsMGL | ?{$_.SkuPartNumber -notmatch $rgxbannedSPN}){
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