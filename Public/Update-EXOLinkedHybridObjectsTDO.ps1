# Update-EXOLinkedHybridObjectsTDO.ps1
# Update-LinkHybridObjects.ps1

#region UPDATE_EXOLINKEDHYBRIDOBJECTSTDO ; #*------v Update-EXOLinkedHybridObjectsTDO v------
function Update-EXOLinkedHybridObjectsTDO {
    <#
    .SYNOPSIS
    Update-EXOLinkedHybridObjectsTDO - Tests & Repairs an EXO mailbox for critical hybrid matches: mgUser,ADUser,HardMatch,RemoteMailbox,ExGuidMatch, creates ADUser and RemoteMailbox if possible, reports on any conflicting OnPrem mailbox, Updates ImmutableID and ExchangeGuid to bring into sync.
    .NOTES
    Version     : 0.0.1
    Author      : Todd Kadrie
    Website     : http://www.toddomation.com
    Twitter     : @tostka / http://twitter.com/tostka
    CreatedDate : 2026-
    FileName    : Update-EXOLinkedHybridObjectsTDO.ps1
    License     : MIT License
    Copyright   : (c) 2026 Todd Kadrie
    Github      : https://github.com/tostka/verb-XXX
    Tags        : Powershell
    AddedCredit : REFERENCE
    AddedWebsite: URL
    AddedTwitter: URL
    REVISIONS
    * 9:21 AM 11/2/2026 added $MGImmutableOrUpnChange to MG_CONNECT block of SVCCONN_LITE: an option to force an SID - required for all Immutable or UPN updates (CBA perms won't cover) 
        Also spliced over latest full SVCCONN_LITE, primarily adds switch support for edge/local or rems.
    * 2:03 PM 10/30/2026 added --ea 0 to disconnect-mg's (err suppress when no connection); ren'd ThisXoMbx -> EXOMailbox
    * 10:20 AM 10/29/2026 updated, added pretest on get-mgcontext.scopes, abort Update-MgUser if has a blocked prop update, wo the elevated perm reqd; 
        Updated CBH to fully doc the cmdlines.
        Added test- alias, only effective under function use, but this is intended to be a resident of verb-exo in the long term (when done debugging).
    * 2:49 PM 10/28/2026 updated mg connect code to use connect-mggraph v connect-mg, that *ensures* the connection is SID and non-CBA - CBA is BLOCKED FROM ANY IMMUTABLE/UPN ETC SECURITY TIED CHANGES
        to user objects. 
        - fixed issue with blank ADU.msExchMailboxGuid (not OnPrem recipient mapping, attempts to get-mailbox get-remotemailbox -id on the blank value throws an error)
    * 2:05 PM 4/14/2026 updated dnam -> samaccountname conversion
    * 12:48 PM 4/2/2026 added colorcombo coding for readability
    * 3:34 PM 3/31/2026 updated conflcit op mbx align testing, conditional rmbx output in hsreport
    * 2:56 PM 3/26/2026 ADD: -DisableConflictMailbox to do immed purge on conficting OnPrem mbx;
          updated echos, cleared rem'd blocks, stdized the catch blocks;
         added 'employeeType' to spot if these are offic or production onboards
    * 5:39 PM 3/25/2026 updated logic to check for OPMbx on Rmbx missing, also 
        added code grab Get-xoMailboxStatistics -id & Get-MailboxStatistics in that 
        case, and output conditional report for comparison, along with echo'd cojmmand 
        to manually delete conflicting mailbox.  Used to repair Tx 29719
    * 3:58 PM 3/20/2026 tweaked output reporting to make it clar status, added email address and CA5 status reporting ; 
         rplfcd disk disco api calls with non-aliess (-whatifpreference was echoing whatifs setting the aliases); 
         flipped back to a .ps1 from func, easier to ensure only this copy is running ; don't really need this occaisional item in verb-exo
    * 1:52 PM 3/20/2026 latest rev, mostly func, missing some outlier parts, did work past the update-mguser scope block. Needs more debugging
    * 5:36 PM 3/16/2026 init
    .DESCRIPTION

    Update-EXOLinkedHybridObjectsTDO - Tests & Repairs an EXO mailbox for critical hybrid matches: mgUser,ADUser,HardMatch,RemoteMailbox,ExGuidMatch, creates ADUser and RemoteMailbox if possible, reports on any conflicting OnPrem mailbox, Updates ImmutableID and ExchangeGuid to bring into sync.
    If user mailbox is properly configured, no changes are attempted, and status is echoed.

    # NOTE: UPDATE-MGUSER -onpremisesimmutableid (& +'Employee Type', 'Employee Hire Date', usagelocation) 
    requires BEYOND globalADMIN defaults: throws 'Access Denied' 
    unless scopes includes Directory.AccessAsUser.All     
    which reflects performing anything the user themselves can do.

    "Directory.AccessAsUser.All",'Directory.ReadWrite.All','User.ReadWrite.All','AuditLog.Read.All','openid','profile','User.Read','User.Read.All','email' ;

    Tried backing out MG scopes:
    PS> $MGScopesRqd = verb-mg\get-MGCodeCmdletPermissionsTDO -path D:\scripts\Update-EXOLinkedHybridObjectsTDO.ps1 ;

    [How do i run a bulk update for the 'Employee Type' and 'Employee Hire Date' attribute for users in you tenant using a CSV file in MS graph PowerShell - Microsoft Q&A](https://learn.microsoft.com/en-us/answers/questions/1471000/how-do-i-run-a-bulk-update-for-the-employee-type-a)
    says: 
    ```
    Update-MgUser is a graph command and based upon the link you have attached is 
    meant for cloud users only, could you please check if the user is cloud user only.
    if not, you can use permissions Directory.ReadWrite.All, User.ReadWrite.All (application)     

    In case nothing works, maybe you can also give a try to use the following permission.
    Directory.AccessAsUser.All
    ```

    [Insufficient privileges to complete the operation. update-mguser usagelocation - Google Search - www.google.com/](https://www.google.com/search?client=firefox-b-1-d&q=%22Insufficient+privileges+to+complete+the+operation.%22+update-mguser+usagelocation)
    ```
    The "Insufficient privileges" error when running Update-MgUser -UsageLocation usually means the connected session lacks User.ReadWrite.All or Directory.ReadWrite.All permissions. To fix this, reconnect with appropriate scopes, ensure your admin role is sufficient, or use PIM to elevate permissions. 
    Common Causes & Solutions

        Insufficient Scopes (Most Likely): The session requires elevated delegated or application permissions.
            Fix: Run Connect-MgGraph -Scopes "User.ReadWrite.All".
        Updating Administrators: If you are a standard admin attempting to update a Global Administrator, the operation may fail.
            Fix: Use an account with higher or equal privileges (e.g., Global Admin).
        Missing PIM Elevation: If you are a privileged user, you may need to activate your role via PIM (Privileged Identity Management) first.
        Service Principal Issues: If running an app, ensure the Service Principal has Directory.ReadWrite.All assigned in Microsoft Entra ID. 
    ```
    Vasil Michev 126.5K Reputation points • MVP • Volunteer Moderator
    Aug 30, 2022, 1:25 AM

    As the error message suggests, you don't have permissions to update the password on said user. For this operation to succeed when run against a privileged (admin) user, you need to run it as a Global Administrator assigned the Directory.AccessAsUser.All permission.

    fe4e49e8-b04c-4ff3-8119-3e036c67dbb0
    ```

    => Nothing except  Directory.AccessAsUser.All actually works, even for a Global Admin. 

    .PARAMETER EXOMailbox
    EXO Mailbox Object (product of get-xoMailbox) to be checked[-EXOMailbox (get-xomailbox -id USERID)]
    .PARAMETER ticket
    TicketNumber[-ticket '99999']
    .PARAMETER DisableConflictMailbox
    Switch to delete conflicting OnPrem mailbox when discovered[-DisableConflictMailbox]
    .PARAMETER waitPostChangeSecs
    Seconds to wait after change made, to refresh object[-waitPostChangeSecs 15]
    .PARAMETER emlCoTag
    Tag string to be appended to conflicting email address objects, to create unique addresses[-emlCoTag '-INT']
    .PARAMETER RequiredScopes
    Scopes required for planned cmdlets to be executed[-RequiredScopes @('User.Read.All', 'Group.Read.All', 'Domain.Read.All')]
    .PARAMETER Force
    Force (Confirm-override switch, overrides ShouldProcess testing, executes somewhat like legacy -whatif:`$false)[-force]
    .INPUTS
    None. Does not accepted piped input.(.NET types, can add description)
    .OUTPUTS
    None. Returns no objects or output        
    .EXAMPLE
    PS> .\Update-EXOLinkedHybridObjectsTDO.ps1 -EXOMailbox (get-xomailbox -id Aaaa.Aaaaaa@aaaa.aaa) -ticket 99999 -DisableConflictMailbox -WhatIf ; 
    Demo use of optional -DisableConflictMailbox parameter to purge conflicting OnPrem mailbox after discovery
    .LINK
    https://github.com/tostka/powershellbb/
    #>
    [Alias('Test-EXOLinkedHybridObjectsTDO')]
    [CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact = 'High')]
    #[CmdletBinding()]
    PARAM(
        [Parameter(Mandatory=$TRUE,HelpMessage="EXO Mailbox Object (product of get-xoMailbox) to be checked[-EXOMailbox (get-xomailbox -id USERID)]")]
            [psobject]$EXOMailbox,
        #[Parameter(HelpMessage="CustomAttribute5 Update Value[-ca5 'Spartanmowers']")]
        #    [string]$ca5 = 'Spartanmowers',
        [Parameter(HelpMessage="TicketNumber[-ticket '99999']")]
            [string]$ticket,
        #[Parameter(HelpMessage="Domain Name to be used for constructing ADUser proxyAddress Additions (for missing onmicrosoft.com)[-newdom 'spartanmowers.com']")]
        #    [string]$newdom = 'spartanmowers.com',
        [Parameter(HelpMessage="Switch to delete conflicting OnPrem mailbox when discovered[-DisableConflictMailbox]")]
            [switch]$DisableConflictMailbox,
        [Parameter(HelpMessage="Seconds to wait after change made, to refresh object[-waitPostChangeSecs 15]")]
            [int]$waitPostChangeSecs = 30,
        [Parameter(HelpMessage="Tag string to be appended to conflicting email address objects, to create unique addresses[-emlCoTag '-INT']")]
            [string]$emlCoTag = '-INT',
        [Parameter(Mandatory=$False,HelpMessage="Scopes required for planned cmdlets to be executed[-RequiredScopes @('User.Read.All', 'Group.Read.All', 'Domain.Read.All')]")]
            [Alias('scopes')] # alias the connect-mggraph underlying param, for passthru
            [array]$RequiredScopes = @("Directory.AccessAsUser.All",'Directory.ReadWrite.All','User.ReadWrite.All','AuditLog.Read.All','openid','profile','User.Read','User.Read.All','email'),
        [Parameter(HelpMessage="Force (Confirm-override switch, overrides ShouldProcess testing, executes somewhat like legacy -whatif:`$false)[-force]")]
            [switch]$Force,
        #[switch]$WHATIF = $true # whatif is implied by SSP, throws: A parameter with the name 'WhatIf' was defined multiple times for the command. if both $whatif param and SupportsShouldProcess=$true
        [Parameter(HelpMessage="switch to suppress non-essential echos[-silent]")]
            [switch]$silent
    )
    BEGIN{
        #region ENVIRO_DATA ; #*------v ENVIRO_DATA v------
        $Verbose = [boolean]($VerbosePreference -eq 'Continue') ; 
        # USE: -whatif:$($whatifswitch) # proxy for -whatif when SupportsShouldProcess
        [boolean]$whatIfSwitch = ($WhatIf.IsPresent -or $whatif -eq $true -OR $WhatIfPreference -eq $true);  $smsg = "-Verbose:$($Verbose)`t-Whatif:$($whatifswitch) " ;  write-host -foregroundcolor yellow $smsg 
        $rPSCmdlet = $PSCmdlet ; # an object that represents the cmdlet or advanced function that's being run. Available on functions w CmdletBinding (& $args will not be available). (Blank on non-CmdletBinding/Non-Adv funcs).
        $rPSScriptRoot = $PSScriptRoot ; # the full path of the executing script's parent directory., PS2: valid only in script modules (.psm1). PS3+:it's valid in all scripts. (Funcs: ParentDir of the file that hosts the func)
        $rPSCommandPath = $PSCommandPath ; # the full path and filename of the script that's being run, or file hosting the funct. Valid in all scripts.
        $rMyInvocation = $MyInvocation ; # populated only for scripts, function, and script blocks.
        # - $MyInvocation.MyCommand.Name returns name of a function, to identify the current command,  name of the current script (pop'd w func name, on Advfuncs)
        # - Ps3+:$MyInvocation.PSScriptRoot : full path to the script that invoked the current command. The value of this property is populated only when the caller is a script (blank on funcs & Advfuncs)
        # - Ps3+:$MyInvocation.PSCommandPath : full path and filename of the script that invoked the current command. The value of this property is populated only when the caller is a script (blank on funcs & Advfuncs)
        #     ** note: above pair contain information about the _invoker or calling script_, not the current script
        $rPSBoundParameters = $PSBoundParameters ; 
        # CONTINUES @ ENVIRO_DISCOVER
        #endregion ENVIRO_DATA ; #*------^ END ENVIRO_DATA ^------

        if($rMyInvocation.InvocationName -like '*test-EXOLinkedHybridObjectsTDO'){
            write-verbose 'Alias:test-EXOLinkedHybridObjectsTDO was used (setting -whatif:`$true).' ;
            $whatif = $true ;
        } ;
        #region MODULES_FORCE_LOAD ; #*------v MODULES_FORCE_LOAD v------
        # core modes in dep order
        $tmods = @('verb-IO','verb-Text','verb-logging','verb-Desktop','verb-dev','verb-Mods','verb-Network','verb-Auth','verb-ADMS','VERB-ex2010','verb-EXO','VERB-mg') ; 
        # task mods in dep order
        $tmods += @('ExchangeOnlineManagement','ActiveDirectory','Microsoft.Graph.Users')
        $oWPref = $WarningPreference ; $WarningPreference = 'SilentlyContinue' ; 
        $tmods | %{ $thismod = $_ ; TRY{$thismod | ipmo -fo  -ea STOP }CATCH{write-host -foregroundcolor yellow "Missing module:$($thismod)`ntrying find-module lookup..." ; find-module $thismod}} ; $WarningPreference = $oWPref ; 
        #endregion MODULES_FORCE_LOAD ; #*------^ END MODULES_FORCE_LOAD ^------
    
        #region CONSTANTS_AND_ENVIRO ; #*======v CONSTANTS_AND_ENVIRO v======
        #region LOCAL_CONSTANTS ; #*------v LOCAL_CONSTANTS v------
        $prpOPMbx = 'Name','Alias','Database','ExchangeGuid','UserPrincipalName','whencreated','whenchanged' ; 
        # resolved functional, per [How do i run a bulk update for the 'Employee Type' and 'Employee Hire Date' attribute for users in you tenant using a CSV file in MS graph PowerShell - Microsoft Q&A](https://learn.microsoft.com/en-us/answers/questions/1471000/how-do-i-run-a-bulk-update-for-the-employee-type-a)
        #$RequiredScopes = "Directory.AccessAsUser.All",'Directory.ReadWrite.All','User.ReadWrite.All','AuditLog.Read.All','openid','profile','User.Read','User.Read.All','email' ; 
        # non-func list using get-mgcommandlet permissions: bogus. 
        #'User.ReadBasic.All','User.ReadWrite.All','DeviceManagementApps.ReadWrite.All','User.Read.All','Directory.ReadWrite.All','Directory.Read.All','DeviceManagementServiceConfig.ReadWrite.All','User.ReadWrite.CrossCloud','DeviceManagementManagedDevices.ReadWrite.All','DeviceManagementManagedDevices.Read.All','DeviceManagementConfiguration.ReadWrite.All','DeviceManagementConfiguration.Read.All','DeviceManagementServiceConfig.Read.All','DeviceManagementApps.Read.All','User.ReadWrite','User-Mail.ReadWrite.All','User-PasswordProfile.ReadWrite.All','User-Phone.ReadWrite.All','User.EnableDisableAccount.All','User.ManageIdentities.All' ; 
        #region WH_COLORCOMBOS ; #*------v WH_COLORCOMBOS v------
        # USE:L write-host $smsg @whCFAIL 
        $whcBnr = if(gcm get-colorcombo -ea 0){get-colorcombo 62 | write-output }else{ @{BackgroundColor='Magenta';ForegroundColor='Black'}|write-host } ; 
        $whcFAIL = if(gcm get-colorcombo -ea 0){get-colorcombo 20 | write-output }else{ @{BackgroundColor='DarkRed';ForegroundColor='Yellow'}|write-host } ; 
        $whcSucc = if(gcm get-colorcombo -ea 0){get-colorcombo 49 | write-output }else{ @{BackgroundColor='Green';ForegroundColor='Blue'}|write-host } ; 
        $whcACT = if(gcm get-colorcombo -ea 0){get-colorcombo 52 | write-output }else{ @{BackgroundColor='Yellow';ForegroundColor='Black'}|write-host } ; 
        $whcRpt = if(gcm get-colorcombo -ea 0){get-colorcombo 48 | write-output }else{ @{BackgroundColor='White';ForegroundColor='Black'}|write-host } ; 
        $whcStat = if(gcm get-colorcombo -ea 0){get-colorcombo 28 | write-output }else{ @{BackgroundColor='Gray';ForegroundColor='Blue'}|write-host } ; 
        #endregion WH_COLORCOMBOS ; #*------^ END WH_COLORCOMBOS ^------
        #endregion LOCAL_CONSTANTS ; #*------^ END LOCAL_CONSTANTS ^------  
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

        #region TEST_METAS ; #*------v TEST_METAS v------
        # critical dependancy Meta variables
        $MetaNames = 'TOR','CMW','TOL' # ,'NOSUCH' ; 
        # critical dependancy Meta variable properties
        $MetaProps = 'legacyDomain','o365_TenantDomain' #,'DOESNTEXIST' ; 
        $isBased = $true ; $gvMiss = @() ; $ppMiss = @() ; 
        foreach($met in $metanames){
            write-verbose "chk:`$$($met)Meta" ; 
            if(-not (gv -name "$($met)Meta" -ea 0)){
                $isBased = $false; $gvMiss += "$($met)Meta" ; 
            } ; 
            foreach($mp in $MetaProps){
                write-verbose "chk:`$$($met)Meta.$($mp)" ; 
                #if(-not (gv -name "$($met)Meta" -ea 0).value[$mp]){ # testing has a value, not is present as a spec!
                if(-not (gv -name "$($met)Meta" -ea 0).value.keys -contains $mp){$isBased = $false; $ppMiss += "$($met)Meta.$($mp)" ; } ; 
            } ; 
        } ; 
        if($gvmiss){write-warning "Missing Dependant Meta variables:`n$(($gvMiss |%{"`$$($_)" }) -join ',')" } ; 
        if($ppMiss){write-warning "Missing Dependant Meta vari properties:`n$(($ppMiss |%{"`$$($_)" }) -join ',')" } ; 
        if(-not $isBased){ write-warning  "missing critical dependancy profile config!" } ; 
        #endregion TEST_METAS ; #*------^ END TEST_METAS ^------

        write-verbose "Coerce configured but blank Resultsize to Unlimited" ; 
        if(get-variable -name resultsize -ea 0){ if( ($null -eq $ResultSize) -OR ('' -eq $ResultSize) ){$ResultSize = 'unlimited' } elseif($Resultsize -is [int]){} else {throw "Resultsize must be an integer or the string 'unlimited' (or blank)"} ; } ;      
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

        #region FUNCTIONS_INTERNAL ; #*======v FUNCTIONS_INTERNAL v======
        
        #region UPDATE_MGUIMMUT ; #*------v Update-MGUIMmut v------
        Function Update-MGUIMmut{
            <# call: Update-MGUIMmut -thisMGU $thisMGU -thisADU $rptOutput.ADUser -Verbose:$($VerbosePreference -eq 'Continue') -whatif:$($whatifswitch) 
            #>
            [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'HIGH')] 
            PARAM(
                [Parameter(Mandatory=$true)]
                    $thisMGU, 
                [Parameter(Mandatory=$true)]
                    $thisADU
            ) ; 
            TRY{
                $OpImmutableId = [System.Convert]::ToBase64String($ThisADU.ObjectGUID.ToByteArray())
                $pltUdMgu = [ordered]@{
                    UserId = $ThisMgu.Id ;
                    OnPremisesImmutableId = $OpImmutableId ;
                    ErrorAction = 'Stop' ;
                    WhatIf = $whatIfSwitch ;
                }
                $MGUPropsThatRequireDirAccAsUsrAll = @('OnPremisesImmutableId','UserPrincipalName','UsageLocation','EmployeeHireDate','EmployeeType') ; 
                $MGCtxt = get-mgcontext ;
                [regex]$rgxMGUPropsThatRequireDirAccAsUsrAll = ('(' + (($MGUPropsThatRequireDirAccAsUsrAll |%{[regex]::escape($_)}) -join '|') + ')') ;
                write-verbose "check splat for elevated/restricted properties" ; 
                if($BlockedProp = $pltUdMgu.keys -match $rgxMGUPropsThatRequireDirAccAsUsrAll){
                    $smsg = "Updates are being made to restricted MGUser properties:`n$($BlockedProp -join ',')" ; 
                    $smsg += "`nWITHOUT required (App-blocked) permission/scope 'Directory.AccessAsUser.All'!" ;
                    $smsg += "`n=>ABORTING Update-MgUser ATTEMPT (will FAIL with 'Insufficient privileges to complete the operation')" ;
                    write-warning $smsg ; 
                    $stopResults = try {Stop-transcript -ErrorAction stop} CATCH {} ;
                    if($stopResults){
                        $smsg = "Stop-transcript:$($stopResults)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                    } ;
                    return ; 
                } ; 
                write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):Update-MgUser w`n$(($pltUdMgu|out-string).trim())" ;
                if ($Force -or $PSCmdlet.ShouldProcess($ThisMgu.displayname, "Update-MgUser")) {
                    #Update-MgUser -UserId $ThisMgu.Id -OnPremisesImmutableId $OpImmutableId -ErrorAction Stop -WhatIf:$whatIfSwitch
                    Update-MgUser @pltUdMgu ;
                    $doUpdtMGUOnPremImmut = $true ;
                    write-verbose "refresh thisMGU" ;
                    $ThisMgu = Get-MgUserFull -UserId $rptOutput.xoMailbox.ExternalDirectoryObjectId -ErrorAction Stop
                    $rptOutput.hasImmutSync = $true ;
                    $ThisMgu | write-output  ; 
                } else {
                    Write-Host "(-Whatif or `"No`" to the prompt)"
                    $ThisMgu | write-output  ; 
                } ; 
            } CATCH [System.Exception] {
                $ErrTrapd=$Error[0] ;
                if($ErrTrapd.ErrorDetails -match 'Insufficient\sprivileges\sto\scomplete\sthe\soperation'){
                    $smsg = "Update-MgUse: PERMISSION FAIL!" ;
                    $smsg += "`nIF USING ISE, RUN THE CMD OUTSIDE OF ISE IN A FULL PS SESSION!" ;
                    $smsg += "`n`nPS> Update-MgUser -UserId $($ThisMgu.Id) -OnPremisesImmutableId $($OpImmutableId) -ErrorAction Stop -WhatIf"
                    write-warning $smsg ;
                    write-warning $smsg ;
                }else{
                    THROW $ErrTrapd
                } ;
            }CATCH {                  
                #write-warning "POPULATED onPremisesImmutableId and no MGUser.UPN matching ADuser"
                $ErrTrapd=$Error[0] ;
                write-host -foregroundcolor gray "TargetCatch:} CATCH [$($ErrTrapd.Exception.GetType().FullName)] {"  ;
                $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                write-warning "$($smsg)" ;
            }
        } ;
        #endregion UPDATE_MGUIMMUT ; #*------^ END Update-MGUIMmut ^------

        #endregion FUNCTIONS_INTERNAL ; #*------^ END FUNCTIONS_INTERNAL ^------

        #region TRANSCRIPT_NODEPLITE ; #*------v TRANSCRIPT_NODEPLITE v------
        TRY{
            #if([system.io.fileinfo]$rPSCmdlet.MyInvocation.MyCommand.Source){
            if($ps1Path = [system.io.fileinfo]$rMyInvocation.mycommand.definition){
                $transcript = (join-path -path $ps1Path.DirectoryName -ChildPath 'logs') ; 
                if(-not (test-path $transcript  -PathType Container -ea 0)){ mkdir $transcript -verbose }
                $transcript = join-path -path $transcript -childpath $ps1Path.BaseName ;                 
            }else{$throw} ;
        }CATCH{
            if($rPSCmdlet.MyInvocation.InvocationName){                
                if(gcm get-ciminstance -ea 0){$drvs = get-ciminstance Win32_LogicalDisk }elseif(gcm Get-WmiObject -ea 0){$drvs = Get-WmiObject Win32_LogicalDisk} 
                if($drvs = $drvs |?{$_.deviceid -match '[A-Z]:' -AND $_.drivetype -eq 3}){
                    foreach($logdrive in @('D:','C:')){if($drvs |?{$_.deviceid -eq $logdrive}){break} } ; 
                }else{write-warning "unable to gcim/gwmi Win32_LogicalDisk class!" ; break } ; 
                $transcript = (join-path -path (join-path -path $logdrive -ChildPath 'scripts') -childpath 'logs') ; 
                if(-not (test-path $transcript  -PathType Container -ea 0)){ mkdir $transcript -verbose }
                $transcript = join-path -path $transcript -childpath $rPSCmdlet.MyInvocation.InvocationName ;                 
            } ELSE{
                $smsg = "FUNCTION: Unable to resolve the function name (blank `$rPSCmdlet.MyInvocation.InvocationName)" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                BREAK ; 
            }; 
        }
        if($ticket){$transcript += "-$($ticket)" }
        if($whatif -OR $WhatIf.IsPresent -OR $WhatIfPreference.IsPresent){$transcript += "-WHATIF"}ELSE{$transcript += "-EXEC"} ; 
        if($EXOMailbox.userprincipalname){$transcript += "-$($EXOMailbox.userprincipalname)"} ; 
        $transcript += "-$(get-date -format 'yyyyMMdd-HHmmtt')-trans-log.txt" ;
        #$transcript = "d:\scripts\24381-CA5-$($ca5)-updates-$(get-date -format 'yyyyMMdd-HHmmtt')-trans-log.txt" ;
        $stopResults = try {Stop-transcript -ErrorAction stop} CATCH {} ;
        if($stopResults){
            $smsg = "Stop-transcript:$($stopResults)" ;
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
        } ;
        $startResults = start-Transcript -path $transcript -whatif:$false -confirm:$false;
        if($startResults){
            $smsg = "start-transcript:$($startResults)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        } ;        
        #endregion TRANSCRIPT_NODEPLITE ; #*------^ END TRANSCRIPT_NODEPLITE ^------
        
        $VerbosePreference = 'Continue'

        #region SVCCONN_LITE ; #*------v SVCCONN_LITE v------
        if ((get-service MSExchangeEdgeCredential -ea 0) -AND ($hklmPath = (resolve-path "HKLM:\SOFTWARE\Microsoft\ExchangeServer\v*\EdgeTransportRole").path)) {$isXOPLocal = $isXOPEdge = $true
        } elseif ((get-service MSExchangeADTopology -ea 0)  -AND ($hklmPath = (resolve-path "HKLM:\SOFTWARE\Microsoft\ExchangeServer\v*\Setup").path)) {$isXOPLocal = $isXOPRole = $true
        } else{$isXOPRole = $isXOPEdge = $false}; 
        if($hklmPath){
            switch -regex ($hklmPath) {
                '\\v14\\' { $isEx2010 = $true ; $ExVers = 'Ex2010' ; write-host "\v14\Setup == Ex2010" ; }
                '\\v15\\' { write-host "\v15\Setup == Ex2013/Ex2016/Ex2019/ExSE" }
                default {write-host "Unable to manually resolve $($hklmPath) to a known version path!"} ; 
            } ;  
        } ; 
        if($isXOPRole){
            write-verbose "load AD before LEMS" ; 
            get-pssession | Remove-PSSession ; 
            import-module -fo -verb ActiveDirectory ; 
            $domaincontroller = get-gcfast ;
            TRY{
                . "$($env:ExchangeInstallPath)\bin\RemoteExchange.ps1" ;
                Connect-ExchangeServer -auto ;
            }CATCH {throw "Failed to Load Exchange PowerShell Module..."} ;
        } elseif($isXOPEdge){
            TRY{
                write-host "Loading Edge Role PSSnapins..." ;
                [xml]$PSSnapIns = Get-Content -Path "$env:ExchangeInstallPath\Bin\exshell.psc1" -ErrorAction Stop ;
                ForEach($PSSnapIn in $PSSnapIns.PSConsoleFile.PSSnapIns.PSSnapIn){
                    write-verbose ("Trying to add PSSnapIn: {0}" -f $PSSnapIn.Name) ;
                    Add-PSSnapin -Name $PSSnapIn.Name -ErrorAction Stop ;
                } ;
                Import-Module $env:ExchangeInstallPath\bin\Exchange.ps1 -ErrorAction Stop ;
            }CATCH {throw "Failed to Load Exchange PowerShell Module..."} ;   
        }else{
            $smsg = "Running Remomte EMS Connection: Resolve ComputerSite..." ;
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
            TRY{$Site=[System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::GetComputerSite().Name }CATCH{
                $Site=$env:COMPUTERNAME ;
                $smsg = "Non-AD-Connected system, setting `$Site:`$env:COMPUTERNAME" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Prompt }
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } ;
            $smsg = "Resolved ComputerSite: $($Site)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Prompt }
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            $smsg = "Discovering and connecting to a local Exchange server in local AD Site"  ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Prompt }
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            $pltNPSS=[ordered]@{ siteName = $Site ; RoleNames = @('HUB','CAS') ; Verbose = ($PSBoundParameters['Verbose'] -eq $true) ; } ; 
            $smsg = "Connect-ExchangeServerTDO w`n$(($pltNPSS|out-string).trim())" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            $PSSession = Connect-ExchangeServerTDO @pltNPSS ;        
        };
        if(-not ($OrgName = ((get-OrganizationConfig).DistinguishedName.split(',') |?{$_ -like 'DC=*'}) -join '.' -replace 'DC=','')){
            $smsg = "Missing Exchange Connection! (no (Get-OrganizationConfig).name returned)" ;
            throw $smsg ;
            $smsg | write-warning  ;
        } else {
            $smsg = "(connected to EXOP.Org:$($orgName))" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Success }
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        } ;
        #region XOPRVW_CONN ; #*------v XOPRVW_CONN v------
        if(-not $rgxConnectionUriEXO){$rgxConnectionUriEXO = 'https://outlook\.office365\.com'} ; 
        if(-not $rgxConnectionUriPrvw){$rgxConnectionUriPrvw = 'https://ps\.compliance\.protection\.outlook\.com'} ; 
        $usePrvw = $false; $xoPrefix = 'xo'  ; 
        if(gcm Get-ConnectionInformation -ea 0){$xoStat = (Get-ConnectionInformation -ea 0 )}else{write-warning "mssing ExchangeOnlineManagement module!`n(no Get-ConnectionInformation)"} ; 
        $xoTest = @() ; $xoTest+= [boolean]($xoStat | ?{ $_.ConnectionUri -match $rgxConnectionUriEXO -AND $_.State -eq 'Connected' -AND $_.TokenStatus -eq 'Active'}) ; 
        if($xoPrefix){$xoTest+= [boolean]($xoStat | ?{ $_.ModulePrefix -eq $xoPrefix})} ; 
        $isXoConn = [boolean]($xoTest -contains $true -AND $xoTest -notcontains $false) ; 
        if (-not $isXoConn) {$pltCXO = @{silent=$($silent)} ; if($xoPrefix){$pltCXO.add('Prefix',$xoPrefix)} ; 
                Connect-EXO @pltCXO ; 
            } else { write-verbose "EXO connected" };
        $isPrvwConn = [boolean]($xoStat | ? { $_.ConnectionUri -eq $rgxConnectionUriPrvw -AND $_.State -eq 'Connected' -AND $_.TokenStatus -eq 'Active' }) ; if($usePrvw -AND -not $isPrvwConn) {
                Connect-IPPSSession  -ErrorAction stop ; 
            } else { write-verbose "EXO connected" } ; 
        #endregion XOPRVW_CONN ; #*------^ END XOPRVW_CONN ^------
        #region MG_CONNECT ; #*------v MG_CONNECT v------
	    #$RequiredScopes = "Directory.AccessAsUser.All",'Directory.ReadWrite.All','User.ReadWrite.All','AuditLog.Read.All','openid','profile','User.Read','User.Read.All','email' ;        
	    $MGImmutableOrUpnChange = $true ; # 2:10 PM 10/28/2026 add option to force an SID - required for all Immutable or UPN updates (CBA perms won't cover)
	    if(gcm test-mgconnection -ea 0){
		    $MGConn = test-mgconnection -Verbose:($VerbosePreference -eq 'Continue') -silent:$($silent) -ea STOP;
		    if ($RequiredScopes) { $addScopes = @() ; $RequiredScopes | foreach-object { $thisPerm = $_ ; if ($mgconn.scopes -contains $thisPerm) { write-verbose "has scope: $($thisPerm)" } else { $addScopes += @($thisPerm) ; write-verbose "ADD scope: $($thisPerm)" } } ; } ;
		    $pltCcMG = [ordered]@{NoWelcome = $true; ErrorAction = 'STOP'; silent = $($silent) } ;
		    if ($addScopes) { $pltCcMG.add('RequiredScopes', $addscopes); $pltCcMG.add('ContextScope', 'Process'); $pltCCMG.silent = $false; write-verbose "Adding non-default Scopes, setting non-persistant single-process ContextScope" } ;
		    if($MGImmutableOrUpnChange -OR ($MGConn.isConnected -AND $addScopes -AND $mgconn.CertificateThumbprint)) {
			    # update to use connect-mggraph, not connect-mg - forces SID interactive logon
			    $smsg = "CBA cert lacking scopes :$($addscopes -join ',')!"  ; $smsg += "`nDisconnecting to use interactive connection: connect-mggraph -Scopes `"'$($addscopes -join "','")'`"" ; write-warning $smsg ;
			    disconnect-mggraph -ErrorAction SilentlyContinue             
		    } elseif ($MGConn.isConnected -AND $addScopes -and -not ($mgconn.CertificateThumbprint)) {write-verbose "Connect via Account with specified scopes" 
		    } elseif (-NOT ($MGConn.isConnected) -AND $addScopes -and -not ($mgconn.CertificateThumbprint)) { $pltCCMG.add('Credential', $credO365TORSID) 
		    } elseif (-NOT ($MGConn.isConnected)) { write-verbose "not isMGConnected" ; 
		    }else { write-verbose "(currently connected with any specifically specified required Scopes)" ; 
			    $pltCcMG = $null ; }
		    if ($pltCcMG) {
			    if($pltCCMG.RequiredScopes){
				    $pltCCMG.add('Scopes',$pltCCMG.RequiredScopes) ;
			    } ; 
			    if($pltCcMG.keys -contains 'RequiredScopes'){$pltCCMG.remove('RequiredScopes')} ;  ; 
			    if($pltCcMG.keys -contains 'silent'){$pltCCMG.remove('silent') } ; 
			    if($pltCcMG.keys -contains 'Credential'){$pltCCMG.remove('Credential') } ; 
			    #$smsg = "connect-mg w`n$(($pltCCMG.getenumerator() | ?{$_.name -notmatch 'requiredscopes'} | ft -a | out-string|out-string).trim())" ; 
			    $smsg = "connect-mgGraph w`n$(($pltCCMG.getenumerator() | ?{$_.name -notmatch 'requiredscopes'} | ft -a | out-string|out-string).trim())" ; 
			    if($pltCCMG.requiredscopes){$smsg += "`n`n-requiredscopes:`n$(($pltCCMG.requiredscopes|out-string).trim())`n"} ; 
			    if ($silent) {} else { if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }else { write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
			    #connect-mg @pltCCMG ;                
			    connect-mggraph @pltCCMG ;
		    } ;
	    }elseif(gcm get-mgcontext -ea STOP){
		    write-host -foregroundcolor yellow "no verb-mg: using native mg cmdlets" ; 
		    $MGCtxt = get-mgcontext ;
		    if ($RequiredScopes) { $addScopes = @() ; $RequiredScopes | foreach-object { $thisPerm = $_ ; if (($MGCtxt.scopes|?{$_}) -contains $thisPerm) { write-verbose "has scope: $($thisPerm)" } else { $addScopes += @($thisPerm) ; write-verbose "ADD scope: $($thisPerm)" } } ; } ;
		    $pltCcMG = [ordered]@{NoWelcome = $true; ErrorAction = 'STOP'} ;
		    if ($addScopes) { $pltCcMG.add('Scopes', $addscopes); $pltCcMG.add('ContextScope', 'Process'); $pltCCMG.silent = $false; write-verbose "Adding non-default Scopes, setting non-persistant single-process ContextScope" } ;
		    if($MGImmutableOrUpnChange -OR ($MGCtxt -AND $addScopes -AND $MGCtxt.CertificateThumbprint) ) {              
			    $smsg = "CBA cert lacking scopes :$($addscopes -join ',')!"  ; $smsg += "`nDisconnecting to use interactive connection: connect-mggraph -Scopes `"'$($addscopes -join "','")'`"" ; $smsg += "`n(alt: : connect-mggraph -Scopes `"'$($addscopes -join "','")'`" )" ; write-warning $smsg ;
			    disconnect-mggraph -ErrorAction SilentlyContinue ; 
		    } elseif ($MGCtxt -AND $addScopes -and -not ($MGCtxt.CertificateThumbprint)) {write-verbose "(isMGConnected via Account)"
		    } elseif (-NOT ($MGCtxt) -AND $addScopes -and -not ($MGCtxt.CertificateThumbprint)) {
		    } elseif (-NOT ($MGCtxt)) { 
		    }else { write-verbose "(currently connected with any specifically specified required Scopes)" ; 
			    $pltCcMG = $null ; }
		    if ($pltCcMG) {
			    $smsg = "connect-mggraph w`n$(($pltCCMG.getenumerator() | ?{$_.name -notmatch 'scopes'} | ft -a | out-string|out-string).trim())" ; 
			    if($pltCCMG.scopes){$smsg += "`n`n-scopes:`n$(($pltCCMG.requiredscopes|out-string).trim())`n"} ; 
			    if ($silent) {} else { if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }else { write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
			    connect-mggraph @pltCCMG ;
		    } ;
	    }
	    if (-not (get-command Get-MgUser)) {
		    $smsg = "Missing Get-MgUser!" ; $smsg += "`nPre-connect to Microsoft.Graph via:" ; $smsg += "`nConnect-MgGraph -Scopes `'$($requiredscopes -join "','")`'" ; if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent }else { write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
		    BREAK ;
	    } ;
	    #endregion MG_CONNECT ; #*------^ END MG_CONNECT ^------
        $isXoPConn = [boolean]( (gcm get-pssession -ea 0) -AND (get-pssession -ea 0 | ? { $_.State -eq 'Opened' -AND $_.Availability -eq 'Available' })); if (-not $isXoPConn) {
            Reconnect-Ex2010 -silent:$($silent)
        } else { write-verbose "XOP connected" };
        $isADConn = [boolean](gcm get-aduser -ea 0) ; if (!$isADConn) { $env:ADPS_LoadDefaultDrive = 0 ; $sName = "ActiveDirectory"; if (!(Get-Module | where { $_.Name -eq $sName })) { Import-Module $sName -ea Stop } }else { write-verbose "ADMS connected" };
        #endregion SVCCONN_LITE ; #*------^ END SVCCONN_LITE ^------
        #region SVCCONN_TEST ; #*------v SVCCONN_TEST v------
        $isXoConn = [boolean]( (gcm Get-ConnectionInformation -ea 0) -AND (Get-ConnectionInformation -ea 0 | ? { $_.State -eq 'Connected' -AND $_.TokenStatus -eq 'Active' })) ; if (-not $isXoConn) { write-host "No EXO connection! RUN: connect-excchangeonline or rxo" ; BREAK }else { write-verbose "EXO connected" };
        $isMgConn = [boolean](get-mgcontext) ; if (-not $isMgConn) { write-host "No MG connection! RUN: Connect-MgGraph or Connect-Mg" ; BREAK }else { write-verbose "MG connected" };
        $isXoPConn = [boolean]( (gcm get-pssession -ea 0) -AND (get-pssession -ea 0 | ? { $_.State -eq 'Opened' -AND $_.Availability -eq 'Available' })); if (-not $isXoPConn) { WRITE-WARNING "NO XOP CONNECTION! run rx10" ; BREAK }else { write-verbose "XOP connected" };
        $isADConn = [boolean](gcm get-aduser -ea 0) ; if (!$isADConn) { TRY { $env:ADPS_LoadDefaultDrive = 0 ; $sName = "ActiveDirectory"; if (!(Get-Module | where { $_.Name -eq $sName })) { Import-Module $sName -ea Stop } } CATCH { $ErrTrapd = $Error[0] ; WRITE-WARNING "$(($ErrTrapd | fl * -Force|out-string).trim())" ; BREAK } }else { write-verbose "ADMS connected" };
        #endregion SVCCONN_TEST ; #*------^ END SVCCONN_TEST ^------
        
    } ;  # BEG-E
    PROCESS{
        $sBnr="#*======v PROCESSING : $($EXOMailbox.userprincipalname) v======" ; 
        write-host "$((get-date).ToString('HH:mm:ss')):$($sBnr)" @$whcBnr ;

        <#
        $hasXoMbx = $false ;
        $hasXoMbxDirSync = $false ; 
        $hasMgUser = $false ;
        $hasADUser = $false ; 
        $hasImmutSync = $false ; 
        $hasRmbx = $false ; 
        $hasRmbxExGuidMatch = $false ; 
        $hasBadOPMbx = $false ; 
        $hasAduDupe = $false ; 
    
        $doAddADUser = $false ; 
        $doUpdtMGUOnPremImmut = $false ; 
        $doAddRmbx = $false ; 
        $doRmbxExGuidMatch = $false ; 
        $doOpMbxConfictDisable = $false 
        #>
        # shift to exportable summary object for above, plus data fieldds
        $rptOutput = [ordered]@{
            hasXoMbx = $false ;
            hasXoMbxDirSync = $false ; 
            hasMgUser = $false ;
            hasADUser = $false ;             
            hasImmutSync = $false ; 
            hasRmbx = $false ; 
            hasRmbxExGuidMatch = $false ; 
            hasBadOPMbx = $false ; 
            hasAduDupe = $false ;     
            doAddADUser = $false ; 
            doUpdtMGUOnPremImmut = $false ; 
            doAddRmbx = $false ; 
            doRmbxExGuidMatch = $false ; 
            doOpMbxConfictDisable = $false
            MGUser = $null ; 
            xoMailbox = $null ; 
            ADUser = $null ;             
            opMailbox = $null ; 
            RemoteMailbox  = $null ; 
        } ; 

        # 0. Resolve xombx to live object (csv collapsed)
        try {
            $DC = GET-GCFAST ;
            $rptOutput.xoMailbox = get-xomailbox -id $EXOMailbox.ExchangeGuid -ea STOP ; 
            $rptOutput.hasXoMbx = $true ;
            $wh = get-colorcombo -rand ; 
            write-host "->Resolved XOMailbox: $($($rptOutput.xoMailbox.userprincipalname))" @wh ; 
        } CATCH {
            $ErrTrapd=$Error[0] ;
            $smsg = "Failed to resolve XOMailbox"                        
            $smsg += "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
            write-warning "$($smsg)" ;
            write-host -foregroundcolor gray "TargetCatch:} CATCH [$($ErrTrapd.Exception.GetType().FullName)] {"  ;
            throw $smsg ; 
        } ;
        # 1. Resolve MG User from ExternalDirectoryObjectId
        TRY {
            if($rptOutput.MGUser = Get-MgUserFull -UserId $rptOutput.xoMailbox.ExternalDirectoryObjectId -ErrorAction Stop){
            $rptOutput.hasMgUser = $true ;}
            $wh = get-colorcombo -rand ; 
            write-host "->Resolved MGUser: $($($rptOutput.MGUser.userprincipalname))" @wh ; 
        } CATCH {
            $ErrTrapd=$Error[0] ;
            $smsg = "Failed to resolve MG user for ExternalDirectoryObjectId:"                        
            $smsg += "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
            write-warning "$($smsg)" ;
            write-host -foregroundcolor gray "TargetCatch:} CATCH [$($ErrTrapd.Exception.GetType().FullName)] {"  ;
            throw $smsg ; 
        } ;

        # 2. If onPremisesImmutableId exists, convert base64 -> GUID, then find AD user
        $ThisOnPremisesImmutableId = (Get-MgUser -UserId $rptOutput.xoMailbox.ExternalDirectoryObjectId -Property OnPremisesImmutableId -ErrorAction Stop | Select-Object -ExpandProperty OnPremisesImmutableId) 
        write-host "->Resolved MGUser.OnPremisesImmutableId: $($ThisOnPremisesImmutableId)" ; 

        $rptOutput.ADUser = $null
        if ($ThisOnPremisesImmutableId) {
            $wh = get-colorcombo -rand ; 
            try {
                $GuidBytes = [System.Convert]::FromBase64String($ThisOnPremisesImmutableId)
                $GuidObj = New-Object -TypeName guid -ArgumentList (,$GuidBytes)
                $rptOutput.ADUser = Get-ADUser -Identity $GuidObj.Guid -Properties * -server $dc -ErrorAction Stop
                $rptOutput.hasADUser = $true ;                 
                write-host "->Found matching AD user by immutableId: $($rptOutput.ADUser.DistinguishedName)" @wh ; 
            }CATCH {
                write-host "onPremisesImmutableId present but resolving AD user failed: $($_.Exception.Message)" @whcFAIL; 
                try {            
                    #$rptOutput.ADUser = Get-ADUser -Identity $GuidObj.Guid -Properties * -ErrorAction Stop
                    $thisUPN = $rptOutput.xoMailbox.userprincipalname ; 
                    write-warning "->onPremisesImmutableId present, UNMATCHED, checking for an ADUser with MGUser.UPN:$($thisUPN)" ; 
                    $rptOutput.ADUser = get-aduser -Filter {userprincipalname -eq $thisUPN} -Properties * -server $dc -ErrorAction STOP ;
                    if($rptOutput.ADUser){
                        $rptOutput.hasADUser = $true ; 
                        $smsg = "Found matching AD user by UPN: $($rptOutput.ADUser.DistinguishedName)" ; 
                        $smsg += "`nUpdating MgUser OnPremisesImmutableId to match ADUser OpImmutableId" ;
                        write-warning $smsg ;                        
                        $rptOutput.MGUser = Update-MGUIMmut -thisMGU $rptOutput.MGUser -thisADU $rptOutput.ADUser -Verbose:$($VerbosePreference -eq 'Continue') -whatif:$($whatifswitch) ;
                    } ; 
                <#
                } CATCH [System.Exception] {
                    $ErrTrapd=$Error[0] ;
                    if($ErrTrapd.ErrorDetails -match 'Insufficient\sprivileges\sto\scomplete\sthe\soperation'){
                        $smsg = "Update-MgUse: PERMISSION FAIL!" ;
                        $smsg += "`nIF USING ISE, RUN THE CMD OUTSIDE OF ISE IN A FULL PS SESSION!" ;
                        $smsg += "`n`nPS> Update-MgUser -UserId $($rptOutput.MGUser.Id) -OnPremisesImmutableId $($OpImmutableId) -ErrorAction Stop -WhatIf"
                        write-warning $smsg ;
                        write-warning $smsg ;
                    }else{
                        THROW $ErrTrapd
                    } ;
                #>
                }CATCH {                  
                    #write-warning "POPULATED onPremisesImmutableId and no MGUser.UPN matching ADuser"
                    $ErrTrapd=$Error[0] ;
                    write-host -foregroundcolor gray "TargetCatch:} CATCH [$($ErrTrapd.Exception.GetType().FullName)] {"  ;
                    $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                    write-warning "$($smsg)" ;
                }
            } # CATCH
        }else{
            TRY {            
                #$rptOutput.ADUser = Get-ADUser -Identity $GuidObj.Guid -Properties * -ErrorAction Stop
                $thisUPN = $rptOutput.xoMailbox.userprincipalname ; 
                write-warning "->NO onPremisesImmutableId present, checking for an ADUser with MGUser.UPN:$($thisUPN)" ; 
                $wh = get-colorcombo -rand ; 
                $rptOutput.ADUser = get-aduser -Filter {userprincipalname -eq $thisUPN} -Properties * -server $dc -ErrorAction STOP -Server $dc ;            
                if($rptOutput.ADUser){
                    $rptOutput.hasADUser = $true ; 
                    $smsg = "->Found matching AD user by UPN: $($rptOutput.ADUser.DistinguishedName)"  ; 
                    if(-not $whatIfSwitch){
                        $smsg += "`n->Updating MgUser OnPremisesImmutableId to match ADUser OpImmutableId" ;
                    } else{
                        $smsg += "`n->NEEDS: Updating MgUser OnPremisesImmutableId to match ADUser OpImmutableId" ;
                    } ; 
                    write-warning $smsg ; 
                    # Hard-match MG user to newly created AD user
                    
                    $rptOutput.MGUser = Update-MGUIMmut -thisMGU $rptOutput.MGUser -thisADU $rptOutput.ADUser -Verbose:$($VerbosePreference -eq 'Continue') -whatif:$($whatifswitch) 
                    write-host "->Resolved MGUser: $($($rptOutput.MGUser.userprincipalname))" ; 
                } ; 
            <#
            } CATCH [System.Exception] {
                    $ErrTrapd=$Error[0] ;
                    if($ErrTrapd.ErrorDetails -match 'Insufficient\sprivileges\sto\scomplete\sthe\soperation'){
                        $smsg = "Update-MgUse: PERMISSION FAIL!" ;
                        $smsg += "`nIF USING ISE, RUN THE CMD OUTSIDE OF ISE IN A FULL PS SESSION!" ;
                        $smsg += "`n`nPS> Update-MgUser -UserId $($rptOutput.MGUser.Id) -OnPremisesImmutableId $($OpImmutableId) -ErrorAction Stop -WhatIf"
                        write-warning $smsg ;
                        write-warning $smsg ;
                    }else{
                        THROW $ErrTrapd
                    } ;
            #>
            }CATCH {
                write-warning "NO onPremisesImmutableId and no MGUser.UPN matching ADuser"
            }
        }

        # 3. If mailbox is DirSynced, check for RemoteMailbox by ExchangeGuid
        $rptOutput.RemoteMailbox = $null
        if ($rptOutput.xoMailbox.IsDirSynced) {
            $rptOutput.hasXoMbxDirSync = $true ; 
            try {
                $wh = get-colorcombo -rand ; 
                $rptOutput.RemoteMailbox = Get-RemoteMailbox -Identity $rptOutput.xoMailbox.ExchangeGuid.Guid -DomainController $dc -ErrorAction SilentlyContinue
                if ($rptOutput.RemoteMailbox) { 
                    write-host "->Found RemoteMailbox: $($rptOutput.RemoteMailbox.Identity)" @WH ; 
                    $rptOutput.hasRmbx = $true ; 
                    $rptOutput.hasRmbxExGuidMatch = $true ; 
                }else{
                    $rptOutput.RemoteMailbox = Get-RemoteMailbox -Identity $rptOutput.xoMailbox.UserPrincipalName -DomainController $dc -ErrorAction SilentlyContinue
                    if ($rptOutput.RemoteMailbox) { 
                        write-host "->Found RemoteMailbox (via ExGuid): $($rptOutput.RemoteMailbox.Identity)" @wh ; 
                        $rptOutput.hasRmbx = $true ;                     
                    } ; 
                } ; 
            } CATCH {
                $ErrTrapd=$Error[0] ;
                $smsg = "Get-RemoteMailbox lookup error:"                        
                $smsg += "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" ;
                write-host -foregroundcolor gray "TargetCatch:} CATCH [$($ErrTrapd.Exception.GetType().FullName)] {"  ;
                throw $smsg ; 
            } ;
        }ELSE{
            # intersync, it's possible for there to be an RMBX, and the XoMbx to be .isDirsynced:$false
            $wh = get-colorcombo -rand ; 
            try {
                $rptOutput.RemoteMailbox = Get-RemoteMailbox -Identity $rptOutput.xoMailbox.UserPrincipalName -DomainController $dc -ErrorAction SilentlyContinue
                if ($rptOutput.RemoteMailbox) { 
                    write-host "->Found RemoteMailbox (via UPN): $($rptOutput.RemoteMailbox.Identity)" @wh ; 
                    $rptOutput.hasRmbx = $true ;                     
                } ;             
            }CATCH {
                write-host "Get-RemoteMailbox lookup error: $($_.Exception.Message)"
            }
        }

        # if no rmbx, check for a blocking xop mailbox
        if (-not $rptOutput.RemoteMailbox -and $rptOutput.xoMailbox.IsDirSynced) {
            #[guid]$inputobject.msExchMailboxGuid
            #if($OPMailbox = get-mailbox -identity $inputobject.guid ){
            # with non-hardmatch, there won't be a populated msExchMailboxGuid, test it
            if($rptOutput.ADUser.msExchMailboxGuid){
                if($rptOutput.opMailbox = Get-Mailbox -Identity ([guid]$rptOutput.ADUser.msExchMailboxGuid).guid -DomainController $dc -ErrorAction SilentlyContinue){
                    $smsg = "NO RMBX: FOUND CONFLICTING OP MAILBOX!`n(resolved from this ADUser.msExchMailboxGuid)" ; 
                }elseif($rptOutput.opMailbox = Get-Mailbox -Identity $rptOutput.xoMailbox.UserprincipalName -DomainController $dc -ErrorAction SilentlyContinue){
                     $smsg = "NO RMBX: FOUND CONFLICTING OP MAILBOX!`n(resolved from this ADUser.Userprincipalname)" ; 
                }; 
            } else {
                write-host -foregroundcolor yellow "unpoplated ADU.msExchMailboxGuid: No ADU-mounted Rmbx/Mbx onprem" ; 
                $rptOutput.opMailbox = Get-Mailbox -Identity $rptOutput.xoMailbox.UserprincipalName -DomainController $dc -ErrorAction SilentlyContinue
            }
            if($rptOutput.opMailbox){ 
                # check account mapping on OPMbx
                if($rptOutput.opMailbox.ExchangeGuid.GetType().fullname -eq 'System.Guid'){
                    $exOPGuid = $rptOutput.opMailbox.ExchangeGuid
                    if($rptOutput.ADUserConflict = Get-ADUser -Filter { msExchMailboxGuid -eq $exOPGuid } -Properties msExchMailboxGuid){
                        #$ADUser| write-output  ; 
                        if($rptOutput.ADUserConflict.ObjectGUID -eq $rptOutput.ADUser.ObjectGUID){
                            $smsg += "`n~~>Mounted on SAME ADUser object resolved MGUser(matching ObjectGuid)" ;
                        }else{
                            $smsg += "`n~~>Mounted on DIFFERENT/DUPLICATE ADUser than object resolved MGUser(matching ObjectGuid)" ;
                            $rptOutput.hasAduDupe = $true ; 
                        } 
                    }else{
                        write-warning "Unable to resolve a matching ADUser for conflicting OPMbx:`nGet-ADUser -Filter { msExchMailboxGuid -eq $($exOPGuid) }" ; 
                    }
                }
                write-WARNING $smsg  ; 
                $rptOutput.hasBadOPMbx = $TRUE ; 
                $rptOutput.xoMailboxStats = $rptOutput.xoMailbox | Get-xoMailboxStatistics -ea STOP ; 
                $rptOutput.OPMailboxStats = $rptOutput.opMailbox | Get-MailboxStatistics -ea STOP ; 

                if($DisableConflictMailbox){
                    $smsg = "-DisableConflictMailbox SPECIFIED!" ;                     
                    $smsg += "`nMoving to Disable-Mailbox the conflict!" ;                    
                    write-warning $smsg
                    TRY{
                        # disable-mailbox -identity 3f15a958-4d99-4741-89e6-96f74f9439ca -domaincontroller LYNMS8102 -VERBOSE -WHATIF
                        $pltDOpMbx = [ordered]@{
                            identity = $rptOutput.opMailbox.ExchangeGuid.guid ;
                            domaincontroller = $dc ;
                            VERBOSE = $true ;
                            WHATIF = $($whatifswitch) ;
                            ErrorAction = 'STOP' ;                          
                        }
                        $smsg = "Disable-Mailbox w`n$(($pltDOpMbx|out-string).trim())" ; 
                        write-warning $smsg ; 
                        if ($Force -OR $PSCmdlet.ShouldProcess($rptOutput.opMailbox.Identity, 'Disable-Mailbox ExchangeGuid')) {
                            Write-Host "EXEC:Disable-Mailbox $($rptOutput.opMailbox.database)\$($rptOutput.opMailbox.userprincipalname)" ; 
                            Disable-Mailbox @pltDOpMbx 
                            if(-not ($rptOutput.opMailbox = Get-Mailbox -Identity $pltDOpMbx.identity -DomainController $dc -ErrorAction SilentlyContinue)){
                                $rptOutput.doOpMbxConfictDisable = $true ;
                                $rptOutput.hasBadOPMbx = $false ;
                                $smsg = "CONFLICTING MAILBOX DISABLED/REMOVED!" ; 
                                $smsg = "(follow-on code should create and sync RemoteMailbox object)" ; 
                                write-host $smsg @whcSucc ; 
                            }else{
                                $smsg = "FAILED TO REMOVE CONFLICTING MAILBOX!" ; 
                                write-warning $smsg ; 
                                $rptOutput.doOpMbxConfictDisable = $FALSE ;
                                $rptOutput.hasBadOPMbx = $true ;
                            }
                        } else {
                            Write-Host "(-Whatif or `"No`" to the prompt)" ; 
                        }  ;                      
                    } CATCH {$ErrTrapd=$Error[0] ;
                        write-host -foregroundcolor gray "TargetCatch:} CATCH [$($ErrTrapd.Exception.GetType().FullName)] {"  ;
                        $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                        write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" ;
                    } ;

                }
            }else{
                $smsg = "->NO RMBX: found *NO* conflicting OP mailbox!" ; 
                write-host  $smsg @whcSucc ; 
            } 
        } ; 

        # PreCheck ExchangeGuid on RemoteMailbox matches Exchange Online mailbox's ExchangeGuid
        if ($rptOutput.RemoteMailbox -and $rptOutput.xoMailbox.ExchangeGuid) {
            if ($rptOutput.RemoteMailbox.ExchangeGuid -ne $rptOutput.xoMailbox.ExchangeGuid.Guid) {
                write-host "->Setting RemoteMailbox ExchangeGuid to match XO mailbox ExchangeGuid" @whcACT
                #$whatIfSwitch = $WhatIf.IsPresent        
                #if ($PSCmdlet.ShouldProcess($rptOutput.RemoteMailbox.Identity, 'Set-RemoteMailbox ExchangeGuid')) {
                if ($Force -OR $PSCmdlet.ShouldProcess($rptOutput.RemoteMailbox.Identity, 'Set-RemoteMailbox ExchangeGuid')) {
                    Write-Host "Write Action Here: $InputObject" ; 
                    Set-RemoteMailbox -Identity $rptOutput.RemoteMailbox.Identity -ExchangeGuid $rptOutput.xoMailbox.ExchangeGuid.Guid -DomainController $dc -WhatIf:$whatIfSwitch -ErrorAction Stop
                    $rptOutput.doRmbxExGuidMatch = $false ; 
                    # refresh the rmbx for trailing report
                    $rptOutput.RemoteMailbox = Get-RemoteMailbox -Identity $rptOutput.RemoteMailbox.Identity -DomainController $dc -ErrorAction STOP
                    $rptOutput.hasRmbxExGuidMatch = $true ; 
                } else {
                    Write-Host "(-Whatif or `"No`" to the prompt)" ; 
                }  ;                      
            }else{
                $rptOutput.hasRmbxExGuidMatch = $true ; 
                write-host -foregroundcolor green "->Has hasRmbxExGuidMatch" ; 
            }
        }

        # 4. If AD user exists but MG user is not hard-matched, set OnPremisesImmutableId to AD.ObjectGUID
        if ($rptOutput.ADUser -and $rptOutput.MGUser) {
            # Compare base64 of AD GUID to MG OnPremisesImmutableId
            $OpImmutableId = [System.Convert]::ToBase64String($rptOutput.ADUser.ObjectGUID.ToByteArray())
            if (-not $rptOutput.MGUser.OnPremisesImmutableId -or ($rptOutput.MGUser.OnPremisesImmutableId -ne $OpImmutableId)) {
                write-host "->Updating MG user OnPremisesImmutableId to AD ObjectGUID base64 to hard-match"
               
                $rptOutput.MGUser = Update-MGUIMmut -thisMGU $rptOutput.MGUser -thisADU $rptOutput.ADUser -Verbose:$($VerbosePreference -eq 'Continue') -whatif:$($whatifswitch) 
            }else{
                $rptOutput.hasImmutSync = $true ; 
            }
        } ; 

        # 5. If no AD user exists, create one under the specified OU
        if (-not $rptOutput.ADUser) {
            $smsg = "->NO ADUser exists, Creating one" ; 
            write-host -foregroundcolor yellow $smsg ;

            # Build attributes from MG user
            $ouPath = 'OU=Azure Email Enabled,OU=System Accounts,OU=Other Accounts,OU=LYN,DC=global,DC=ad,DC=toro,DC=com'
            $displayName = $rptOutput.MGUser.DisplayName
            $givenName = $rptOutput.MGUser.GivenName
            $surname = $rptOutput.MGUser.Surname
            $userPrincipalName = $rptOutput.MGUser.UserPrincipalName

            # SamAccountName: first 20 word characters from displayname
            #$SamAccountNameBase = (($displayName.ToCharArray() | Where-Object { $_ -match '\w' } | Select-Object -First 20) -join '')
            #$SamAccountName = $SamAccountNameBase
            $SamAccountName =  -join (($displayName -replace '[^a-zA-Z0-9]', '').ToCharArray()[0..19]) ; 
            $suffix = 1
            while (Get-ADUser -Filter { SamAccountName -eq $SamAccountName } -server $dc -ErrorAction SilentlyContinue) {
                $SamAccountName = "$SamAccountNameBase$suffix"
                $suffix++
            }

            $pltNADU = [ordered]@{
                Name = $displayName ;
                GivenName = $givenName ;
                Surname = $surname ;
                SamAccountName = $SamAccountName ;
                UserPrincipalName = $userPrincipalName ;
                Path = $ouPath ;
                AccountPassword = $secureRandomPassword ;
                Enabled = $false ;       
                PasswordNeverExpires = $true ;
                ChangePasswordAtLogon = $False ;
                Server = $dc ;
                #WhatIf = $whatIfSwitch 
                ErrorAction = 'stop' ; 
            } ;        
            #write-host "Creating AD user $displayName in $ouPath with SamAccountName $SamAccountName"
            if ($PSCmdlet.ShouldProcess($displayName, 'New-ADUser,Update-MgUser')) {
                $secureRandomPassword = (ConvertTo-SecureString -String (('Pfx' + [guid]::NewGuid().ToString()) ) -AsPlainText -Force)
                write-host -foregroundcolor YELLOW "`n->new-ADUser w:`n$(($pltNADU|out-string).trim())`n" ;

                #New-ADUser -Name $displayName -GivenName $givenName -Surname $surname -SamAccountName $SamAccountName -UserPrincipalName $userPrincipalName -Path $ouPath -AccountPassword $secureRandomPassword -Enabled $false -ErrorAction Stop
                TRY{
                    New-ADUser @pltNADU 
                    $rptOutput.doAddADUser = $true ; 
                    # Re-fetch created AD user
                    $rptOutput.ADUser = Get-ADUser -Filter { SamAccountName -eq $SamAccountName } -server $dc -Properties * -ErrorAction Stop
                } CATCH {
                    $ErrTrapd=$Error[0] ;
                    write-host -foregroundcolor gray "TargetCatch:} CATCH [$($ErrTrapd.Exception.GetType().FullName)] {"  ;
                    $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                    write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" ;
                } ;
                
                <#
                # Hard-match MG user to newly created AD user
                $OpImmutableId = [System.Convert]::ToBase64String($rptOutput.ADUser.ObjectGUID.ToByteArray())
                Update-MgUser -UserId $rptOutput.MGUser.Id -OnPremisesImmutableId $OpImmutableId -ErrorAction Stop  -WhatIf:$whatIfSwitch ; 
                $rptOutput.doUpdtMGUOnPremImmut = $true ; 
                write-verbose "refresh thisMGU" ; 
                $rptOutput.MGUser = Get-MgUserFull -UserId $rptOutput.xoMailbox.ExternalDirectoryObjectId -ErrorAction Stop
                $rptOutput.hasImmutSync = $true ; 
                #>
                ADU -Verbose:$($VerbosePreference -eq 'Continue') -whatif:$($whatifswitch) 
            } else {
                Write-Host "(-Whatif or `"No`" to the prompt)"
            } ; 
        } # if-E thisADU

        # 6. If no RemoteMailbox exists and AD user exists and is matched, create RemoteMailbox and set ExchangeGuid
        # add $rptOutput.opMailbox test - won't run wo conflicting object
        if (-not $rptOutput.RemoteMailbox -AND $rptOutput.ADUser -AND $rptOutput.MGUser -AND -not $rptOutput.opMailbox) {
            $SMSG = "-> NO RemoteMailbox, ADUser exists and is matched -> create RemoteMailbox and set ExchangeGuid" ; 
            WRITE-HOST -FOREGROUNDCOLOR YELLOW $smsg ; 
            # Find an .onmicrosoft.com routing address from XO mailbox
            #$RemRouteAddr = ($rptOutput.xoMailbox.EmailAddresses | Where-Object { $_ -match '\.onmicrosoft\.com$' } | ForEach-Object { $_ -replace '^smtp:' , '' } | Select-Object -Last 1) 
            # above matches sip: addresses too
            $RemRouteAddr = ($rptOutput.xoMailbox.EmailAddresses | Where-Object { $_ -match '^smtp:.*\.onmicrosoft\.com$' } | ForEach-Object { $_ -replace '^smtp:' , '' } | Select-Object -Last 1)
            #if (-not $RemRouteAddr) {
            if(-not $RemRouteAddr -and $rptOutput.ADUser){
                $smsg = "->No suitable .onmicrosoft.com remote routing address found on $($rptOutput.xoMailbox.Identity)"
                write-warning $smsg ;
                $smsg = "->$($rptOutput.xoMailbox.userprincipalname).isDirsync, but has *no* onmicrosoft.com suitable address!";
                $smsg += "`n->Attempting to calculate an address and push it into the ADUser.proxyaddresses list (to populate on the xmbx, for a future pass, after ADC sync)" ;
                WRITE-WARNING $SMSG ;
                if($rptOutput.xoMailbox.EmailAddressPolicyEnabled -eq $false){
                    $dirname = "$($rptOutput.xoMailbox.primarysmtpaddress.split('@')[0])$($emlCoTag)" ;
                    $newdom = 'toroco.onmicrosoft.com' ;
                    $newpEml = @($dirname,$newdom) -join '@' ;
                    $newpEml = "smtp:$($newpEml)" ;
                    $xproxy = $rptOutput.ADUser.proxyaddresses  | ?{$_ -match '^smtp\:'} ;
                    if($xproxy -contains $newpEml){
                        $smsg = "->ADU.$($rptOutput.xoMailbox.userprincipalname) already has the necessary RemoteRouting Address added, skipping dupe addition ";
                        write-warning $smsg ;
                    }else{
                        #$rptOutput.xoMailbox | set-xomailbox -EmailAddresses @{add="smtp:$($newpEml)"} -whatif:$($whatif) -ea STOP ;
                        #Set-ADUser -Identity $rptOutput.ADUser.DistinguishedName -Add @{proxyAddresses=$newpEml}  -whatif:$($whatif) -ea STOP -server $dc -VERBOSE  ;
                        if ($Force -or $PSCmdlet.ShouldProcess($rptOutput.ADUser.userprincipalname, "set-ADUser")) {
                            Set-ADUser -Identity $rptOutput.ADUser.objectguid.guid -Add @{proxyAddresses=$newpEml}  -whatif:$($whatifswitch) -ea STOP -server $dc -VERBOSE  ;
                            # refresh the updated obj
                            $rptOutput.ADUser = get-aduser -Identity $rptOutput.ADUser.objectguid.guid -ea STOP -server $dc
                        } else {
                            Write-Host "(-Whatif or `"No`" to the prompt)"
                        } ; 
                    } ;
                    $ADUFix = $true ; 
                    $smsg = "->WAIT FULL ADC CYCLE AND RECHECK IF THE XMBX HAS A SUITABLE ONMICROSOFT.COM ADDRESS (and MGU has matched OPimmuntable), THEN RERUN AN UPDATE TO CREATE RMBX" ;
                    WRITE-WARNING $SMSG ;
                } ;
            } ; 

            if($RemRouteAddr -and $rptOutput.ADUser){
                write-host "->Enabling RemoteMailbox for $($rptOutput.ADUser.UserPrincipalName) with RemoteRoutingAddress $RemRouteAddr"
                if ($PSCmdlet.ShouldProcess($rptOutput.ADUser.UserPrincipalName, 'Enable-RemoteMailbox')) {
                    $rptOutput.RemoteMailbox = Enable-RemoteMailbox -Identity $rptOutput.ADUser.UserPrincipalName -RemoteRoutingAddress $RemRouteAddr -DomainController $dc -ErrorAction Stop ; 
                    $rptOutput.doAddRmbx = $true ; 
                    <#
                    VERBOSE: Enabling RemoteMailbox for Marketing@intimidatorutv.com with RemoteRoutingAddress Marketing1@toroco.onmicrosoft.com
                    This task does not support recipients of this type. The specified recipient global.ad.toro.com/LYN/Other Accounts/System Accounts/Azure Email Enabled/Marketing - Intimidator is of type RemoteUserMailbox. Please make sure
                    that this recipient matches the required recipient type for this task.
                    + CategoryInfo          : InvalidArgument: (global.ad.toro....g - Intimidator:ADObjectId) [Enable-RemoteMailbox], RecipientTaskException
                    + FullyQualifiedErrorId : 2882FF3D,Microsoft.Exchange.Management
                    #>
                    # above throws error, but still created rmbx: redisco it as UPN
                    try {
                        #$rptOutput.RemoteMailbox = Get-RemoteMailbox -Identity $rptOutput.xoMailbox.ExchangeGuid.Guid -ErrorAction SilentlyContinue
                        # newly created won't have the guid match
                        $rptOutput.RemoteMailbox = Get-RemoteMailbox -Identity $rptOutput.xoMailbox.userprincipalname -DomainController $dc -ErrorAction SilentlyContinue
                        if ($rptOutput.RemoteMailbox) { 
                            write-host "->Found RemoteMailbox: $($rptOutput.RemoteMailbox.Identity)" 
                            $rptOutput.hasRmbx = $true ;
                        }                    
                    } CATCH {
                        $ErrTrapd=$Error[0] ;
                        $smsg = "Get-RemoteMailbox lookup error:"                        
                        $smsg += "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                        write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" ;
                        write-host -foregroundcolor gray "TargetCatch:} CATCH [$($ErrTrapd.Exception.GetType().FullName)] {"  ;
                    } ;
                }
            }else{
                $smsg = "->this ADUser but No matched RemoteRoutingAddress, deferrring Rmbx creation until RRA is available (future pass)" ; 
                write-warning $smsg;
            }
            

            # Ensure ExchangeGuid on RemoteMailbox matches Exchange Online mailbox's ExchangeGuid
            if ($rptOutput.RemoteMailbox -and $rptOutput.xoMailbox.ExchangeGuid) {
                if ($rptOutput.RemoteMailbox.ExchangeGuid -ne $rptOutput.xoMailbox.ExchangeGuid.Guid) {
                    write-host "->Setting RemoteMailbox ExchangeGuid to match XO mailbox ExchangeGuid" @whcACT ; 
                    if ($PSCmdlet.ShouldProcess($rptOutput.RemoteMailbox.Identity, 'Set-RemoteMailbox ExchangeGuid')) {
                        Set-RemoteMailbox -Identity $rptOutput.RemoteMailbox.Identity -ExchangeGuid $rptOutput.xoMailbox.ExchangeGuid.Guid -DomainController $dc -WhatIf:$whatIfSwitch -ErrorAction Stop
                        $rptOutput.doRmbxExGuidMatch = $false ; 
                        # refresh the rmbx for trailing report
                        $rptOutput.RemoteMailbox = Get-RemoteMailbox -Identity $rptOutput.RemoteMailbox.Identity -DomainController $dc -ErrorAction STOP
                        $rptOutput.hasRmbxExGuidMatch = $true ; 
                    } else {
                        Write-Host "(-Whatif or `"No`" to the prompt)"
                    } ; 
                }else{
                    $rptOutput.hasRmbxExGuidMatch = $true ; 
                    write-host "->hasRmbxExGuidMatch" ; 
                }
            }
        }elseif($rptOutput.RemoteMailbox -and $rptOutput.ADUser -and $rptOutput.MGUser -AND ($rptOutput.RemoteMailbox.ExchangeGuid -eq $rptOutput.xoMailbox.ExchangeGuid)) {
            $smsg = "Has XoMbx & Rmbx, with ExchangeGuid Alighnment: " ; # $rptOutput.RemoteMailbox -and $rptOutput.xoMailbox.ExchangeGuid
            $smsg += "this Rmbx w`n$(($rptOutput.RemoteMailbox | ft -a |out-string).trim())" ; 
            $smsg += "This XoMbx.ExchangeGuid:`n$($rptOutput.xoMailbox.ExchangeGuid)" ; 
            write-host "$((get-date).ToString('HH:mm:ss')):$($smsg)" @whcSucc ;
        }ELSE{
            $smsg = "MISSING COMPONENT: " ; # $rptOutput.RemoteMailbox -and $rptOutput.xoMailbox.ExchangeGuid
            $smsg += "this Rmbx w`n$(($rptOutput.RemoteMailbox | ft -a |out-string).trim())" ; 
            $smsg += "This XoMbx.ExchangeGuid:`n$($rptOutput.xoMailbox.ExchangeGuid)" ; 
            if($rptOutput.opMailbox){
                #$prpOPMbx = 'Name','Alias','Database','ExchangeGuid','UserPrincipalName' ; 
                $smsg += "`n`n**CONFLICTING** this OPMbx w`n$(($rptOutput.opMailbox | fl $prpOPMbx |out-string).trim())`n`n" ; 
            } ; 
            write-warning $smsg ; 
        }


        $hsReport = @"

### This XoMailbox: 
$(($rptOutput.xoMailbox| ft -a Name,Alias,ServerName,isDirSynced|out-string).trim())
$(($rptOutput.xoMailbox| ft -a ExchangeGuid,CustomAttribute5 |out-string).trim())
$(($rptOutput.xoMailbox| ft -a UserprincipalName,Alias|out-string).trim())
$(($rptOutput.xoMailbox| ft -a PrimarySmtpAddress |out-string).trim())
-> EmailAddresses:
$(($rptOutput.xoMailbox.emailaddresses | ?{$_ -match 'smtp'} | sort| fl |out-string).trim())

    $(
        if($rptOutput.hasBadOPMbx){

            "`n`n### This opMailbox **CONFLICTING!!**`n" | write-output     
            $(($rptOutput.opMailbox| ft -a Name,Alias,ServerName|out-string).trim()) ; "`n"| write-output         
            $(($rptOutput.opMailbox| ft -a ExchangeGuid,CustomAttribute5 |out-string).trim()) ; "`n"| write-output 
            $(($rptOutput.opMailbox| ft -a UserprincipalName,Alias|out-string).trim()) ; "`n"| write-output 
            $(($rptOutput.opMailbox| ft -a PrimarySmtpAddress |out-string).trim()) ; "`n"| write-output 
            '-> EmailAddresses:`n`n' | write-output 
            $(($rptOutput.opMailbox.emailaddresses | ?{$_ -match 'smtp'} | sort| fl |out-string).trim()) ; "`n"| write-output 

            "`n`n### This XoMbxStats`n" | write-output     
            $(($rptOutput.xoMailboxStats| ft -a|out-string).trim()) ; "`n"| write-output 

            "`n`n### This OPMbxStats`n" | write-output     
            $(($rptOutput.OPMailboxStats| ft -a|out-string).trim())  ; "`n"| write-output 
            "`n`n" | write-output  

            "`n`n===>***TO REMOVE CONFLICING ONPREM MAILBOX RUN:***``n`n" | write-output 
            "`n`nPS>  disable-mailbox -identity $($rptOutput.opMailbox.ExchangeGuid) -domaincontroller $($dc) -VERBOSE -WHATIF`n`n" ; 
            "`n`n(or specify -DisableConflictMailbox when running this script)`n`n" ; 
            "`n`n(REQUIRED before Enable-RemoteMailbox will work! - clear OpMbx and rerun this script again)`n`n" ;
        }      
    )

### This Mgu: 
$(($rptOutput.MGUser| ft -a DisplayName,Id,Mail,UserPrincipalName  |out-string).trim())
$(($rptOutput.MGUser| ft -a CreatedDateTime,EmployeeHireDate,EmployeeType,JobTitle  |out-string).trim())
$(($rptOutput.MGUser| ft -a OnPremisesImmutableId,OnPremisesSyncEnabled,OnPremisesLastSyncDateTime|out-string).trim())
$(($rptOutput.MGUser| ft -a OnPremisesDistinguishedName,OnPremisesProvisioningErrors|out-string).trim())
-> ProxyAddresses:
$(($rptOutput.MGUser.proxyaddresses | ?{$_ -match 'smtp'} | sort|out-string).trim())

### This ADU: 
$(($rptOutput.ADUser| ft -a name,DistinguishedName,Enabled |out-string).trim())
$(($rptOutput.ADUser| ft -a 'GivenName','Surname','Name' |out-string).trim())
$(($rptOutput.ADUser| ft -a 'SamAccountName','UserPrincipalName','ObjectClass','ObjectGUID' |out-string).trim())
$(($rptOutput.ADUser| ft -a 'whenCreated','whenChanged','lastlogondate','EmployeeID','employeeType' |out-string).trim())
$(($rptOutput.ADUser| ft -a 'description','title' |out-string).trim())
-> ProxyAddresses:
$(($rptOutput.ADUser.proxyaddresses | ?{$_ -match 'smtp'} | sort|out-string).trim())


### This OnPremisesImmutableId: $(($ThisOnPremisesImmutableId|out-string).trim())
###(equiv converted ADUser:OpImmutableId:$(($OpImmutableId|out-string).trim())

$(
    if($rptOutput.hasAduDupe){
        "### This OPMbx **CONFLICTING!!**`n" | write-output 
        $(($rptOutput.ADUserConflict| ft -a name,DistinguishedName,Enabled|out-string).trim()) ; "`n"| write-output         
        $(($rptOutput.ADUserConflict| ft -a 'GivenName','Surname','Name'  |out-string).trim()) ; "`n"| write-output 
        $(($rptOutput.ADUserConflict| ft -a 'SamAccountName','UserPrincipalName','ObjectClass','ObjectGUID'|out-string).trim()) ; "`n"| write-output 
        $(($rptOutput.ADUserConflict| ft -a 'whenCreated','whenChanged','lastlogondate','EmployeeID','employeeType'|out-string).trim()) ; "`n"| write-output 
        $(($rptOutput.ADUserConflict| ft -a 'description','title'|out-string).trim()) ; "`n"| write-output 
        '-> ProxyAddresses:`n`n' | write-output 
        $(($rptOutput.ADUserConflict.proxyaddresses | ?{$_ -match 'smtp'} | sort| fl |out-string).trim()) ; "`n"| write-output 
    
        "`n`n===>***TO CONFLICING ONPREM ADUSER MAY BE WORKDAY SYNCED OBJECT!:***``n`n" | write-output     
    }
)

$(
    if($rptOutput.RemoteMailbox){         
        "### This Rmbx:`n" | write-output 
        $(($rptOutput.RemoteMailbox| ft -a Name,RecipientTypeDetails,RemoteRecipientType |out-string).trim()); "`n"| write-output  
        $(($rptOutput.RemoteMailbox| ft -a UserprincipalName,Alias,PrimarySMTPAddress|out-string).trim()); "`n"| write-output  
        $(($rptOutput.RemoteMailbox| ft -a ExchangeGuid,CustomAttribute5 |out-string).trim()); "`n"| write-output  
        '-> EmailAddresses:`n`n' | write-output 
        $(($rptOutput.RemoteMailbox.emailaddresses | ?{$_ -match 'smtp'} | sort| fl |out-string).trim()) ; "`n"| write-output 
    }else{
         "`n`n### *MISING Rmbx*!`n`n" | write-output 

    }
)

"@;
        #WRITE-HOST -FOREGROUNDCOLOR GREEN $hsReport ; 
        write-host $hsReport @whcRpt

        # 2:43 PM 10/30/2026 moving to hash props breaks raw name lookup
        $actions = @('doAddADUser','doUpdtMGUOnPremImmut','doAddRmbx','doRmbxExGuidMatch','doOpMbxConfictDisable')    
        $actions | foreach-object{
            $thisActName = $_ ; 
            #if((gv -Name $thisActName -ea 0).value -eq $true){
            if($rptOutput[$thisActName] -eq $true){
                #write-warning "ACTION:`$$($thisActName):$((gv -Name $thisActName).value)!"
                write-warning "ACTION:`$$($thisActName):$($rptOutput[$thisActName])!"                
            } else{
                #write-host "ACTION:$($thisActName):$((gv -Name $thisActName).value)" @whcSucc ; 
                write-host "ACTION:$($thisActName):$($rptOutput[$thisActName])" @whcSucc ; 
            } ; 
        } ; 
        write-host "`n" ; 
        $tests = @('hasXoMbx','hasXoMbxDirSync','hasMgUser','hasADUser','hasImmutSync','hasRmbx','hasRmbxExGuidMatch'); 
        $rptOutput.testResults = @() ; 
        $tests | foreach-object{
            $thisTestName = $_ ; 
            #if((gv -Name $thisTestName -ea 0).value -eq $false){
            if($rptOutput[$thisTestName] -eq $false){
                #$smsg = "TEST:`$$($thistestName):$((gv -Name $thistestname).value)!"
                $smsg = "TEST:`$$($thistestName):$($rptOutput[$thisTestName])!"
                write-warning $SMSG ;
                $rptOutput.testResults+=$false ; 
            } else{
                #$smsg = "TEST:$($thistestName):$((gv -Name $thistestname).value)"
                $smsg = "TEST:$($thistestName):$($rptOutput[$thisTestName])"                
                WRITE-HOST $SMSG @whcSucc ; 
                $rptOutput.testResults+=$true ; 
            } ; 
        } ; 
        # $TRUE IS BAD TESTS
        write-host "`n" ; 
        $BADtests = @('hasBadOPMbx','hasAduDupe'); 
        $rptOutput.BADtestResults = @() ; 
        $BADtests | foreach-object{
            $thisTestName = $_ ; 
            #if((gv -Name $thisTestName -ea 0).value -eq $TRUE){
            if($rptOutput[$thisTestName] -eq $TRUE){
                #$smsg = "TEST:`$$($thistestName):$((gv -Name $thistestname).value)!"
                $smsg = "TEST:`$$($thistestName):$($rptOutput[$thisTestName])!"
                write-warning $SMSG ;
                $rptOutput.BADtestResults += $true ; 
            } else{
                #$smsg = "TEST:$($thistestName):$((gv -Name $thistestname).value)"
                $smsg = "TEST:$($thistestName):$($rptOutput[$thisTestName])"                
                WRITE-HOST $SMSG @whcSucc ; 
                $rptOutput.BADtestResults += $false ; 
            } ; 
        } ; 

        if( ($rptOutput.testResults -contains $false) -OR ($rptOutput.BADtestResults -contains $true)){        
            $smsg = "`n==> **MAILBOX **FAILS** SYNC TESTING" ;
            write-warning $smsg ; 
        } else{
            $smsg = "`n==> **MAILBOX PASSES SYNC TESTING" ;
            WRITE-HOST $smsg @whcSucc; 
        } ;
        write-host -foregroundcolor yellow "Returning Summary Object to Pipeline" ; 
        [pscustomobject]$rptOutput | write-output ; 

        if($rMyInvocation.InvocationName -like '*test-EXOLinkedHybridObjectsTDO'){
            write-host "`nTest-EXOLinkedHybridObjectsTDO completed for $($rptOutput.xoMailbox.Identity)" @whcStat ; 
        }else{
            write-host "`nUpdate-EXOLinkedHybridObjectsTDO completed for $($rptOutput.xoMailbox.Identity)" @whcStat ; 
        }

        write-host "$((get-date).ToString('HH:mm:ss')):$($sBnr.replace('=v','=^').replace('v=','^='))" @whcBnr ;
    } ;  # PROC-E
    END{
        $stopResults = try {Stop-transcript -ErrorAction stop} CATCH {} ;
        if($stopResults){
            $smsg = "Stop-transcript:$($stopResults)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
        } ;
    }
}
#endregion UPDATE_EXOLINKEDHYBRIDOBJECTSTDO ; #*------^ END Update-EXOLinkedHybridObjectsTDO ^------