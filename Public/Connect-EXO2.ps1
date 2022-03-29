﻿#*------v Connect-EXO2.ps1 v------
Function Connect-EXO2 {
    <#
    .SYNOPSIS
    Connect-EXO2 - Establish connection to Exchange Online (via EXO V2 graph-api module)
    .NOTES
    Version     : 1.0.0
    Author      : Todd Kadrie
    Website     :	http://www.toddomation.com
    Twitter     :	@tostka / http://twitter.com/tostka
    CreatedDate : 2020-07-29
    FileName    :
    License     :
    Copyright   :
    Github      : https://github.com/tostka
    Tags        : Powershell,ExchangeOnline,Exchange,RemotePowershell,Connection,MFA
    AddedCredit : ExactMike Perficient
    AddedWebsite:	https://docs.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2?view=exchange-ps
    AddedTwitter:
    AddedCredit2 : Jeremy Bradshaw
    AddedWebsite2:	https://github.com/JeremyTBradshaw
    AddedTwitter2:
    REVISIONS   :
    * 4:10 PM 3/29/2022 finished getting test-EXOToken interaction and EOM module/.dll load code squared up. 
    3:58 PM 3/28/2022 add: DefaultParameterSetName='UPN', avoid: 'Parameter set cannot be resolved using the specified named parameters.'
    - updated module path code, to support new .netcore/.netframework subdir forking of the .dll storage in the EOm module. 
    - all of the write-* to incl wl support. 
    - trying to sort out use of the test-ActiveToken() - wants a new -TokenExpiryTime, in other code in the EOM .psm1 it's called as 
    $hasActiveToken = Test-ActiveToken -TokenExpiryTime $script:PSSession.TokenExpiryTime
    $sessionIsOpened = $script:PSSession.Runspace.RunspaceStateInfo.State -eq 'Opened'
    if (($hasActiveToken -eq $false) -or ($sessionIsOpened -ne $true))
    {
        #If there is no active user token or opened session then ensure that we remove the old session
        $shouldRemoveCurrentSession = $true;
    }
    * 1:24 PM 3/15/2022 moved $minvers to a param: -MinimumVersion
    * 2:40 PM 12/10/2021 more cleanup 
    # 11:23 AM 9/16/2021 string
    # 1:31 PM 7/21/2021 revised Add-PSTitleBar $sTitleBarTag with TenOrg spec (for prompt designators)
    * 11:53 AM 4/2/2021 updated with rlt & recstat support, updated catch blocks
    # 8:34 AM 3/31/2021 added verbose suppress to all import-mods
    * 11:36 AM 3/5/2021 updated colorcode, subed wv -verbose with just write-verbose, added cred.uname echo
    * 1:15 PM 3/1/2021 added org-level color-coded console
    * 8:55 AM 11/11/2020 added fake -Username block, to make -Credential, *also* auto-renew sessions! (above from: https://ingogegenwarth.wordpress.com/2018/02/02/exo-ps-mfa/)
    * 2:01 PM 11/10/2020 swap connect-exo2 to connect-exo2old (uses connect-ExchangeOnline), and ren this "Connect-EXO2A" to connect-exo2 ; fixed get-module tests (sub'd off the .dll from the modname)
    * 9:56 AM 11/10/2020 variant of cxo2, that has direct ported-in low-level code from the ExchangeOnlineManagement:connect-ExchangeOnlin(). debugs functional so far, haven't tested concurrent CCMS + EXO overlap & tokens yet. 
    * 8:30 AM 10/22/2020 ren'd $TentantTag -> $TenOrg, swapped looping meta resolve with 1-liner approach ; added AcceptedDom caching to the middle status test (suppress one more get-exoaccepteddomain call if possible)
    * 4:41 PM 10/8/2020 implemented AcceptedDomain caching, in connect-exo2 to match rxo2
    * 1:18 PM 8/11/2020 fixed typo in *broken *closed varis in use; updated ExoV1 conn filter, to specificly target v1 (old matched v1 & v2) ; trimmed entire rem'd MFA block ; added trailing test-EXOToken confirm
    * 12:57 PM 8/4/2020 sorted ExchangeOnlineMgmt mod issues (splatting wo using splat char), if MS hadn't completely rewritten the access, this rewrite wouldn't have been necessary in the 1st place. I'm not looking forward to the org wide rewrites to recode verb-exoNoun -> verb-xoNoun, to accomodate the breaking-change blocking -Prefix 'exo'. ; # 1:04 PM 8/4/2020 cute: now the above error's stopped occuring on the problem tenant. Can't do further testing of the workaround, unless/until it breaks again ; * 2:39 PM 8/4/2020 fixed -match "^(Session|WinRM)\d*" rgx (lacked ^, mismatched EXOv2 conns)
    * 12:20 PM 7/29/2020 rewrite/port from connect-EXO to replace import-pssession with new connect-ExchangeOnline cmdlet (supports MFA natively) - #127 # *** LEFT OFF HERE 5:01 PM 7/29/2020 *** not sure if it supports allowclobber, if it's actually wrapping pssession, it sure as shit does!
    * 11:21 AM 7/28/2020 added Credential -> AcceptedDomains Tenant validation, also testing existing conn, and skipping reconnect unless unhealthy or wrong tenant to match credential
    * 3:24 PM 7/24/2020 updated to support tenant-alignment & sub'd out showdebug for verbose
    * 7:13 AM 7/22/2020 replaced codeblock w get-TenantTag()
    * 5:12 PM 7/21/2020 added ven supp
    * 11:50 AM 5/27/2020 added alias:cxo win func
    * 8:38 AM 4/17/2020 added a new test of $global:EOLSession, to detect initial cred fail (pw chg, outofdate creds, locked out)
    * 8:45 AM 3/3/2020 public cleanup, refactored Connect-EXO2 for Meta's
    * 9:52 PM 1/16/2020 cleanup
    * 10:55 AM 12/6/2019 Connect-EXO2:added suffix to TitleBar tag for other tenants, also config'd a central tab vari
    * 9:17 AM 12/4/2019 CONSISTENTLY failing to load properly in lab, on lynms6200d - wont' get-module xxxx -listinstalled, even after load, so I rewrote an exemption diverting into the locally installed $env:userprofile\documents\WindowsPowerShell\Modules\exoMFAModule\ copy.
    * 5:14 PM 11/27/2019 repl $MFA code with get-TenantMFARequirement
    * 1:07 PM 11/25/2019 added tenant-specific alias variants for connect & reconnect
    # 1:26 PM 11/19/2019 added MFA detection fr infastrings .ps1 globals, lifted from Jeremy Bradshaw (https://github.com/JeremyTBradshaw)'s Connect-Exchange()
    # 10:35 AM 6/20/2019 added $pltiSess splat dump to the import-pssession cmd block; hard-typed the $Credential [System.Management.Automation.PSCredential]
    # 8:22 AM 11/20/2017 spliced in retry loop into reConnect-EXO2 as well, to avoid using any state testing in scripts, localize it 1x here.
    # 1:49 PM 11/2/2017 coded around non-profile gaps from not having get-admincred() - added the prompt code in to fake it
    # 12:26 PM 9/11/2017 debugged retry - catch doesn't fire properly on new-Pssession, have to test the $error state, to detect auth fails (assuming the bad pw error# is specific). $error test is currently targeting specific error returned on a bad password. Added retry, for when connection won't hold and fails breaks - need to watch out that bad pw doesn't lock out the acct!
    # 12:50 PM 6/2/2017 expanded pshelp, added Examples, cleaned up deadwood
    * # 12:10 PM 3/15/2017 Connect-EXO2 typo, disable prefix auto spec, unless actually blanked. switch ProxyEnabled to non-Mandatory.
    .DESCRIPTION
    Connect-EXO2 - Establish PSS to EXO V2 Modern Auth
    .PARAMETER  Prefix
    [verb]-PREFIX[command] PREFIX string for clearly marking cmdlets sourced in this connection [-Prefix tag]
    .PARAMETER  Credential
    Credential to use for this connection [-credential 'SOMEACCT@DOMAIN.COM']
    .PARAMETER UserPrincipalName
    User Principal Name or email address of the user
    .PARAMETER
    ConnectionUri
    Connection Uri for the Remote PowerShell endpoint [-ConnectionUri 'https://outlook.office365.com/powershell-liveid/']
    .PARAMETER ExchangeEnvironmentName
    Exchange Environment name [-ExchangeEnvironmentName 'O365Default']
    .PARAMETER MinimumVersion
    MinimumVersion to be used when ipmo'ing ExchangeOnlineManagement module[-MinimumVersion '2.0.5']
    .PARAMETER PSSessionOption
    PowerShell session options to be used when opening the Remote PowerShell session
    .PARAMETER BypassMailboxAnchoring
    Switch to bypass use of mailbox anchoring hint.
    .PARAMETER UseMultithreading
    Switch to enable/disable Multi-threading in the EXO cmdlets
    .PARAMETER ShowProgress
    Flag to enable or disable showing the number of objects written
    .PARAMETER Pagesize
    Pagesize Param
    .PARAMETER showDebug
    Debugging Flag [-showDebug]
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    None. Returns no objects or output.
    .EXAMPLE
    Connect-EXO2 -cred $credO365TORSID ;
    Connect using defaults, and leverage any pre-set $global:credo365TORSID variable
    .EXAMPLE
    Connect-EXO2 -Prefix exo -credential (Get-Credential -credential user@domain.com)  ;
    Connect an explicit credential, and use 'exolab' as the cmdlet prefix
    .EXAMPLE
    $cred = get-credential -credential $o365_Torolab_SIDUpn ;
    Connect-EXO2 -credential $cred ;
    Pass in a prefab credential object (useful for auto-shifting to MFA - the function will autoresolve MFA reqs based on the cred domain)
    .LINK
    #>
    [CmdletBinding(DefaultParameterSetName='UPN')]
    [Alias('cxo2')]
    Param(
        # try pulling all the ParameterSetName's - just need to get through it now. - no got through it with a defaultparametersetname (avoids 
        [Parameter(HelpMessage = "[verb]-PREFIX[command] PREFIX string for clearly marking cmdlets sourced in this connection [-Prefix tag]")]
        [string]$Prefix = 'xo',
        [Parameter(ParameterSetName = 'Cred', HelpMessage = "Credential to use for this connection [-credential [credential obj variable]")]
        [System.Management.Automation.PSCredential]$Credential = $global:credo365TORSID,
        [Parameter(ParameterSetName = 'UPN',HelpMessage = "User Principal Name or email address of the user[-UserPrincipalName logon@domain.com]")]
        [string]$UserPrincipalName,
        [Parameter(HelpMessage = "Connection Uri for the Remote PowerShell endpoint [-ConnectionUri 'https://outlook.office365.com/powershell-liveid/']")]
        [string] $ConnectionUri,
        [Parameter(HelpMessage = "Azure AD Authorization endpoint Uri that can issue the OAuth2 access tokens [-AzureADAuthorizationEndpointUri 'https://XXX']")]
        [string] $AzureADAuthorizationEndpointUri,
        [Parameter(HelpMessage = "Exchange Environment name [-ExchangeEnvironmentName 'O365Default']")]
        [Microsoft.Exchange.Management.RestApiClient.ExchangeEnvironment]
        $ExchangeEnvironmentName = 'O365Default',
        [Parameter(HelpMessage = "MinimumVersion to be used when ipmo'ing ExchangeOnlineManagement module[-MinimumVersion '2.0.5']")]
        [string] $MinimumVersion = '2.0.5',
        [Parameter(HelpMessage = "PowerShell session options to be used when opening the Remote PowerShell session [-PSSessionOption `$PsSessObj]")]
        [System.Management.Automation.Remoting.PSSessionOption]
        $PSSessionOption = $null,
        [Parameter(HelpMessage = "Switch to bypass use of mailbox anchoring hint. [-BypassMailboxAnchoring]")]
        [switch] $BypassMailboxAnchoring = $false,
        [Parameter(HelpMessage = "Switch to enable/disable Multi-threading in the EXO cmdlets [-UseMultithreading]")]
        [switch]$UseMultithreading=$true,
        [Parameter(HelpMessage = "Switch to enable or disable showing the number of objects written (defaults `$true)[-ShowProgress]")]
        [switch]$ShowProgress=$true,
        [Parameter(HelpMessage = "Pagesize Param[-PageSize 500]")]
        [uint32]$PageSize = 1000,
        [Parameter(HelpMessage = "Debugging Flag [-showDebug]")]
        [switch] $showDebug
    ) ;
    BEGIN {
        $verbose = ($VerbosePreference -eq "Continue") ;
        if (-not $rgxExoPsHostName) { $rgxExoPsHostName = "^(ps\.outlook\.com|outlook\.office365\.com)$" } ;

          # defer to verb-text if avail
          if(-not(get-command test-uri)){
            function Test-Uri {
                [CmdletBinding()]
                [OutputType([bool])]
                Param(
                    # Uri to be validated
                    [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true, Position=0)]
                    [string]$UriString
                ) ; 
                [Uri]$uri = $UriString -as [Uri]
                $uri.AbsoluteUri -ne $null -and $uri.Scheme -eq 'https'
            } ; 
        } ;
        
        # validate params
        if($ConnectionUri -and $AzureADAuthorizationEndpointUri){
            throw "BOTH -Connectionuri & -AzureADAuthorizationEndpointUri specified, use ONE or the OTHER!";
        }

        # disable prefix spec, unless actually blanked (e.g. centrally spec'd in profile).
        if (-not $Prefix) {
            $Prefix = 'xo' ; # 4:31 PM 7/29/2020 MS has RESERVED use of the 'exo' prefix [facepalm]
            $smsg = "(asserting Prefix:$($Prefix)" ;
            if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        } ;
        if (($Prefix) -and ($Prefix -eq 'EXO')) {
            throw "Prefix 'EXO' is a reserved Prefix, please use a different prefix."
        }

        if (($ConnectionUri) -and (-not (Test-Uri $ConnectionUri))) {
            throw "Invalid ConnectionUri parameter '$ConnectionUri'"
        }
        if (($AzureADAuthorizationEndpointUri) -and (-not (Test-Uri $AzureADAuthorizationEndpointUri))) {
            throw "Invalid AzureADAuthorizationEndpointUri parameter '$AzureADAuthorizationEndpointUri'"
        }

        $TenOrg = get-TenantTag -Credential $Credential ;
        $sTitleBarTag = @("EXO2") ;
        $sTitleBarTag += $TenOrg ;

        $MFA = get-TenantMFARequirement -Credential $Credential ;

        # admin/SID module auto-install code (myBoxes UID split-perm CU, all else t AllUsers)
        $modname = 'ExchangeOnlineManagement' ;
        #$MinimumVersion = '1.0.1' ;
        Try { Get-Module -name $modname -listavailable -ErrorAction Stop | out-null } Catch {
            $pltInMod = [ordered]@{Name = $modname ; verbose=$false ;} ;
            if ( $env:COMPUTERNAME -match $rgxMyBoxUID ) { $pltInMod.add('scope', 'CurrentUser') } else { $pltInMod.add('scope', 'AllUsers') } ;
            $smsg = "Install-Module w scope:$($pltInMod.scope)`n$(($pltInMod|out-string).trim())" ;
            if($silent){}elseif ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            Install-Module @pltIMod ;
        } ; # IsInstalled
        $pltIMod = @{Name = $modname ; ErrorAction = 'Stop' ; verbose=$false} ;
        if ($MinimumVersion) { $pltIMod.add('MinimumVersion', $MinimumVersion) } ;
        Try { Get-Module $modname -ErrorAction Stop | out-null } Catch {
            $smsg = "Import-Module w`n$(($pltIMod|out-string).trim())" ;
            if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            Import-Module @pltIMod ;
        } ; # IsImported

        # .dll etc loads, from connect-exchangeonline: (should be installed with the above)
        if (-not($ExchangeOnlineMgmtPath)) {
            $EOMgmtModulePath = split-path (get-module ExchangeOnlineManagement -list).Path ;
            <# Current ExchangeOnlineMgmt (v2.0.5) has 2 versions:
                gci $EOMgmtModulePath -Recurse |?{$_.Name -eq 'Microsoft.Exchange.Management.RestApiClient.dll'}
                    Directory: C:\Program Files\WindowsPowerShell\Modules\ExchangeOnlineManagement\2.0.5\netCore
                Mode                LastWriteTime         Length Name
                ----                -------------         ------ ----
                -a----         5/8/2021   6:53 AM         192888 Microsoft.Exchange.Management.RestApiClient.dll
                    Directory: C:\Program Files\WindowsPowerShell\Modules\ExchangeOnlineManagement\2.0.5\netFramework
                Mode                LastWriteTime         Length Name
                ----                -------------         ------ ----
                -a----         5/8/2021   8:04 AM         194952 Microsoft.Exchange.Management.RestApiClient.dll

                - netCore for powershellCore
                - netFramework for PS
                PSCore auto varis: 
                if($IsCoreCLR){"PSCore"} else {"PSWin"} ;
                if($IsWindows){"Win"} ;
                if($IsLinux){"Linux"} ; 
                if($IsMacOS){"MacOS"}
            #>
            if($IsCoreCLR){
                $EOMgmtModulePath = resolve-path -Path $EOMgmtModulePath\netcore ;
                $smsg = "(.netcore path in use:" ; 
            } else { 
                $EOMgmtModulePath = resolve-path -Path $EOMgmtModulePath\netFramework
                $smsg = "(.netnetFramework path in use:" ;                 
            } ; 
            $smsg += "$($EOMgmtModulePath))" ; 
            if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        } ;

        if (-not $RestModule) { $RestModule = "Microsoft.Exchange.Management.RestApiClient.dll" } ;
        # stock uses $PSScriptRoot, which will be the verb-EXO path, not the EXOMgmt module have to dyn locate it
        if (-not $RestModulePath) {
            # prior pathing: - they moved the DLL, patched the subdir above
            $RestModulePath = join-path -path $EOMgmtModulePath -childpath $RestModule  ;
            if(-not (test-path $RestModulePath)){
                $smsg = "UNABLE TO test-path `$RestModulePath!:`n$($RestModulePathz)" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } 
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                THROW $smsg 
                BREAK ; 
            } ; 
        } ;
        # paths to proper Module path: Name lists as: Microsoft.Exchange.Management.RestApiClient
        if (-not(get-module $RestModule.replace('.dll',''))) {
            Import-Module $RestModulePath -verbose:$false -force ;
        } ;
        if (-not $ExoPowershellGalleryModule) { $ExoPowershellGalleryModule = "Microsoft.Exchange.Management.ExoPowershellGalleryModule.dll" } ;
        if (-not $ExoPowershellGalleryModulePath) {
            $ExoPowershellGalleryModulePath = join-path -path $EOMgmtModulePath -childpath $ExoPowershellGalleryModule ;
            if(-not (test-path $ExoPowershellGalleryModulePath)){
                $smsg = "UNABLE TO test-path `$ExoPowershellGalleryModulePath!:`n$($ExoPowershellGalleryModulePathz)" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } 
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                THROW $smsg 
                BREAK ; 
            } ; 
        } ;
        # full path: C:\Users\USER\Documents\WindowsPowerShell\Modules\ExchangeOnlineManagement\1.0.1\Microsoft.Exchange.Management.ExoPowershellGalleryModule.dll
        # Name: Microsoft.Exchange.Management.ExoPowershellGalleryModule
        if (-not(get-module $ExoPowershellGalleryModule.replace('.dll','') )) {
            Import-Module $ExoPowershellGalleryModulePath -Verbose:$false ;
        } ;

    } ; # BEG-E
    PROCESS {
        $bExistingEXOGood = $false ;

                # Keep track of error count at beginning.
        $errorCountAtStart = $global:Error.Count;
        $global:_EXO_TelemetryFilePath = $null;

        <# Get-PSSession | fl ConfigurationName,name,state,availability,computername
        -legacy remote-ps New-PSSession results in this PSS:
          ConfigurationName : Microsoft.Exchange
          Name              : WinRM2 (seems to increment)
          State             : Opened
          Availability      : Available
          ComputerName      : ps.outlook.com

        - legacy remote from jumpbox:
        ConfigurationName : Microsoft.Exchange
          Name              : Session6
          State             : Opened
          Availability      : Available
          ComputerName      : ps.outlook.com

        -while a connect-ExchangeOnline (non-MFA, haven't verified) connect results in this PSS:
          ConfigurationName : Microsoft.Exchange
          Name              : ExchangeOnlineInternalSession_4
          State             : Opened
          Availability      : Available
          ComputerName      : outlook.office365.com

        -CCMS session via Connect-IPPSSession
        ConfigurationName : Microsoft.Exchange
        ComputerName      : nam02b.ps.compliance.protection.outlook.com
        Name              : ExchangeOnlineInternalSession_1
        State             : Opened
        Availability      : Available
        #>
        # clear any existing legacy EXO sessions:
        #if( Get-PSSession|Where-Object{($_.ComputerName -match $rgxExoPsHostName) -AND ($_.State -eq 'Opened') -AND ($_.Availability -eq 'Available')}){
        # Get-PSSession | fl ConfigurationName,name,state,availability
        # legacy non-OAuth EXOv2 sessions
        if ( $existingPSSession = Get-PSSession | where-object { $_.ConfigurationName -like "Microsoft.Exchange" -and $_.Name -match "^(Session|WinRM)\d*" } ) {
            # ignore state & Avail, close the conflicting legacy conn's
            if ($existingPSSession.count -gt 0) {
                $smsg = "(closing $($existingPSSession.count) legacy EXO sessions...)" ;
                if($silent){}elseif ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                for ($index = 0; $index -lt $existingPSSession.count; $index++) {
                    $session = $existingPSSession[$index] ;
                    Remove-PSSession -session $session ;
                    $smsg = "Removed the PSSession $($session.Name) connected to $($session.ComputerName)" ;
                    if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                } ;
                $bExistingEXOGood = $false ;
            } ;
        } ;
        #if ( Get-PSSession | where-object { $_.ConfigurationName -like "Microsoft.Exchange" -and $_.Name -like "ExchangeOnlineInternalSession*" -and $_.State -like "*Opened*" -AND ($_.Availability -eq 'Available') } ) {
        # update to *not* tamper with CCMS connects
        if (-not $rgxExoPsHostName) { $rgxExoPsHostName = "^(ps\.outlook\.com|outlook\.office365\.com)$" } ;
        if ( Get-PSSession | where-object { $_.ConfigurationName -like "Microsoft.Exchange" -and $_.Name -like "ExchangeOnlineInternalSession*" -and $_.State -like "*Opened*" -AND ($_.Availability -eq 'Available') -AND ($_.ComputerName -match $rgxExoPsHostName) } ) {
            if( get-command Get-xoAcceptedDomain -ea 0) {
                 #$TenOrg = get-TenantTag -Credential $Credential ;
                if(-not (Get-Variable  -name "$($TenOrg)Meta").value.o365_AcceptedDomains){
                    set-Variable  -name "$($TenOrg)Meta" -value ( (Get-Variable  -name "$($TenOrg)Meta").value += @{'o365_AcceptedDomains' = (Get-xoAcceptedDomain).domainname} )
                } ;
                #if ((Get-xoAcceptedDomain).domainname.contains($Credential.username.split('@')[1].tostring())) {
                if((Get-Variable  -name "$($TenOrg)Meta").value.o365_AcceptedDomains.contains($Credential.username.split('@')[1].tostring())){
                    # validate that the connected EXO is to the $Credential tenant
                    $smsg = "(Existing EXO Authenticated & Functional:$($Credential.username.split('@')[1].tostring()))" ;
                    if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                    $bExistingEXOGood = $true ;
                # issue: found fresh bug in cxo: svcacct UPN suffix @tenantname.onmicrosoft.com, but testing against AccepteDomain, it's not in there (tho @DOMAIN.mail.onmicrosoft.comis)
                }elseif((Get-Variable  -name "$($TenOrg)Meta").value.o365_TenantDomain -eq ($Credential.username.split('@')[1].tostring())){
                    $smsg = "(EXO Authenticated & Functional(TenDom):$($Credential.username.split('@')[1].tostring()))" ; 
                    if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                    $bExistingEXOGood = $true ;
                } else {
                    $smsg = "(Credential mismatch:disconnecting from existing EXO:$($eEXO.Identity) tenant)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } 
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                    DisConnect-EXO2 ;
                    $bExistingEXOGood = $false ;
                } ;
            } else {
                # capture outlier: shows a session wo the test cmdlet, force reset
                DisConnect-EXO2 ;
                $bExistingEXOGood = $false ;
            } ;
        } ;

        if ($bExistingEXOGood -eq $false) {
            # open a new EXOv2 session
            # EXOMgt bits:
            # stock globals recording the session
            $global:_EXO_ExchangeEnvironmentName = $ExchangeEnvironmentName;
            $global:_EXO_ConnectionUri = $ConnectionUri;
            $global:_EXO_AzureADAuthorizationEndpointUri = $AzureADAuthorizationEndpointUri;
            $global:_EXO_PSSessionOption = $PSSessionOption;
            $global:_EXO_BypassMailboxAnchoring = $BypassMailboxAnchoring;
            $global:_EXO_DelegatedOrganization = $DelegatedOrganization;
            $global:_EXO_Prefix = $Prefix;
            $global:_EXO_UserPrincipalName = $UserPrincipalName;
            $global:_EXO_Credential = $Credential;
            $global:_EXO_EnableErrorReporting = $EnableErrorReporting;
            # import the ExoPowershellGalleryModule .dll
            if(-not (get-module $ExoPowershellGalleryModule.replace('.dll','') )){ Import-Module $ExoPowershellGalleryModulePath -verbose:$false} ;
            $global:_EXO_ModulePath = $ExoPowershellGalleryModulePath;

            <# prior module code
            #Connect-ExchangeOnline -Credential $credO365TORSID -Prefix 'xo' -ShowBanner:$false ;
            # *** LEFT OFF HERE 5:01 PM 7/29/2020 *** not sure if it supports allowclobber, if it's actually wrapping pssession, it sure as shit does!

            $pltCXO = @{
                Prefix     = [string]$Prefix ;
                ShowBanner = [switch]$false ;
            } ;

            ==1:52 PM 3/29/2022: v2.0.5 examples
            -------------------------- Example 1 --------------------------
            $UserCredential = Get-Credential
            Connect-ExchangeOnline -Credential $UserCredential
            This example connects to Exchange Online PowerShell using modern authentication for an account that doesn't use
            multi-factor authentication (MFA).
            The first command gets the username and password and stores them in the $UserCredential variable.
            The second command connects the current PowerShell session using the credentials in $UserCredential.
            After the Connect-ExchangeOnline command is complete, the password key in the $UserCredential variable is emptied,
            and you can run Exchange Online PowerShell cmdlets.
            -------------------------- Example 2 --------------------------
            Connect-ExchangeOnline -UserPrincipalName chris@contoso.com
            This example connects to Exchange Online PowerShell using modern authentication for the account chris@contoso.com,
            which uses MFA.
            After the command is successful, you can run ExO V2 module cmdlets and older remote PowerShell cmdlets.
            -------------------------- Example 3 --------------------------
            Connect-ExchangeOnline -AppId <%App_id%> -CertificateFilePath "C:\users\navin\Documents\TestCert.pfx" -Organization
            "contoso.onmicrosoft.com"
            This example connects to Exchange Online in an unattended scripting scenario using the public key of a certificate.
            -------------------------- Example 4 --------------------------
            Connect-ExchangeOnline -AppId <%App_id%> -CertificateThumbprint <%Thumbprint string of certificate%> -Organization
            "contoso.onmicrosoft.com"
            This example connects to Exchange Online in an unattended scripting scenario using a certificate thumbprint.
            -------------------------- Example 5 --------------------------
            Connect-ExchangeOnline -AppId <%App_id%> -Certificate <%X509Certificate2 object%> -Organization
            "contoso.onmicrosoft.com"
            This example connects to Exchange Online in an unattended scripting scenario using a certificate file. This method is
            best suited for scenarios where the certificate is stored in remote machines and fetched at runtime. For example, the
            certificate is stored in the Azure Key Vault.
            #>

            <# new-exopssession params:
            new-exopssession -ConnectionUri -AzureADAuthorizationEndpointUri -BypassMailboxAnchoring -ExchangeEnvironmentName 
            -Credential -DelegatedOrganization -Device -PSSessionOption -UserPrincipalName -Reconnect -CertificateFilePath -CertificatePassword 
            -CertificateThumbprint -AppId -Organization -WhatIf
            #>
            $pltNEXOS = @{
                ExchangeEnvironmentName         = $ExchangeEnvironmentName ;
                #ConnectionUri                   = $ConnectionUri ;
                #AzureADAuthorizationEndpointUri = $AzureADAuthorizationEndpointUri ;
                #UserPrincipalName               = $UserPrincipalName ;
                PSSessionOption                 = $PSSessionOption ;
                #Credential                      = $Credential ;
                BypassMailboxAnchoring          = $($BypassMailboxAnchoring) ;
                #ShowProgress                    = $($showProgress) # isn't a param of new-exopssessoin, is used with set-exo
                #DelegatedOrg                    = $DelegatedOrganization ;
                Verbose                          = $false ;
            }
            <# v2.0.5 updated params as a splat
            $pltNEXOS=[ordered]@{
                ExchangeEnvironmentName = $ExchangeEnvironmentName ;
                ConnectionUri = $ConnectionUri ;
                AzureADAuthorizationEndpointUri = $AzureADAuthorizationEndpointUri ;
                UserPrincipalName = $UserPrincipalName.Value ;
                PSSessionOption = $PSSessionOption ;
                Credential = $Credential.Value ;
                BypassMailboxAnchoring = $BypassMailboxAnchoring ;
                DelegatedOrg = $DelegatedOrganization ;
                # new CBA support
                Certificate = $Certificate.Value ;
                CertificateFilePath = $CertificateFilePath.Value ;
                CertificatePassword = $CertificatePassword.Value ;
                CertificateThumbprint = $CertificateThumbprint.Value ;
                AppId = $AppId.Value ;
                Organization = $Organization.Value ;
                # browser ps7 options
                Device = $Device.Value ;
                InlineCredential = $InlineCredential.Value
            } ; 
            #>
            if ($MFA) {
                # -UserPrincipalName
                #$pltCXO.Add("UserPrincipalName", [string]$Credential.username);
                if ($UserPrincipalName) {
                    $pltNEXOS.Add("UserPrincipalName", [string]$UserPrincipalName);
                    $smsg = "(using cred:$([string]$UserPrincipalName))" ; 
                    if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                } elseif ($Credential -AND -not $UserPrincipalName){
                    $pltNEXOS.Add("UserPrincipalName", [string]$Credential.username);
                    $smsg = "(using cred:$($credential.username))" ; 
                    if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                };
            } else {
                # just use the passed $Credential vari
                #$pltCXO.Add("Credential", [System.Management.Automation.PSCredential]$Credential);
                $pltNEXOS.Add("Credential", [System.Management.Automation.PSCredential]$Credential);
                $smsg = "(using cred:$($credential.username))" ; 
                if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            } ;

            if ($AzureADAuthorizationEndpointUri) { $pltNEXOS.Add("AzureADAuthorizationEndpointUri", [string]$AzureADAuthorizationEndpointUri) } ;
            if ($ConnectionUri) { $pltNEXOS.Add("ConnectionUri", [string]$ConnectionUri) } ;

            $smsg = "Connecting to EXOv2:($($credential.username))"  ;
            if($silent){}elseif ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            $smsg = "New-ExoPSSession w`n$(($pltNEXOS|out-string).trim())" ;
            if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            Try {
                #$global:ExoPSSession = New-PSSession @pltCXO ;
                # looks like connect-exchangonline does create a global: $global:_EXO_PreviousModuleName on successful connect (later: likely added in the $global_EXO block below)
                # - but haven't spotted it in debugging tho', so have to gcm for 1st cmdlt in the module to confirm connected, and then get-xoacceptedomain, to verify connected to desired tenant
                $PSSession = New-ExoPSSession @pltNEXOS ;
            } catch [System.ArgumentException] {
                <# post an attempt fail w conn-exo properly stacks the error into $error[0]:
                    Connect-ExchangeOnline -Credential $credO365VENCSID -Prefix 'xo' -ShowBanner:$false ;
                    Removed the PSSession ExchangeOnlineInternalSession_3 connected to outlook.office365.com
                    Import-PSSession : Data returned by the remote Get-FormatData command is not in the expected format.
                    At C:\Program Files\WindowsPowerShell\Modules\ExchangeOnlineManagement\1.0.1\ExchangeOnlineManagement.psm1:454 char:40
                    + ... oduleInfo = Import-PSSession $PSSession -AllowClobber -DisableNameChe ...
                    +                 ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                    + CategoryInfo          : InvalidResult: (:) [Import-PSSession], ArgumentException
                    + FullyQualifiedErrorId : ErrorMalformedDataFromRemoteCommand,Microsoft.PowerShell.Commands.ImportPSSessionCommand

                    +[SIDS]::[PS]:D:\scripts$ $error[0]
                    Import-PSSession : Data returned by the remote Get-FormatData command is not in the expected format.
                    At C:\Program Files\WindowsPowerShell\Modules\ExchangeOnlineManagement\1.0.1\ExchangeOnlineManagement.psm1:454 char:40
                    + ... oduleInfo = Import-PSSession $PSSession -AllowClobber -DisableNameChe ...
                    +                 ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                    + CategoryInfo          : InvalidResult: (:) [Import-PSSession], ArgumentException
                    + FullyQualifiedErrorId : ErrorMalformedDataFromRemoteCommand,Microsoft.PowerShell.Commands.ImportPSSessionCommand

                    Should be trappable, even external function

                    # 1:04 PM 8/4/2020 cute: now the above error's stopped occuring on the problem tenant. Can't do further testing of the workaround, unless/until it breaks again
                #>
                #$pltCXO.Add('ConnectionUri', [string]'https://outlook.office365.com/powershell-liveid?SerializationLevel=Full') ;
                $pltNEXOS.Add('ConnectionUri', [string]'https://outlook.office365.com/powershell-liveid?SerializationLevel=Full') ;
                $smsg = "'Get-FormatData command is not in the expected format' EXO bug: Retrying with '&SerializationLevel=Full'ConnectionUri`n(details at https://answers.microsoft.com/en-us/msoffice/forum/all/cannot-connect-to-exchange-online-via-powershell/)" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } 
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                $smsg = "`n$((get-date).ToString('HH:mm:ss')):Connect-ExchangeOnline w`n$(($pltCXO|out-string).trim())" ;
                if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                <# when this crashes, it leaves an open PSS matching below that TIES UP YOUR CONN QUOTA!
                Get-PSSession | where-object {$_.ConfigurationName -like "Microsoft.Exchange" -and $_.Name -like "ExchangeOnlineInternalSession*" -and $_.State -like "*Opened*" -AND ($_.Availability -eq 'Available')}
                #>
                $error.clear() ;
                TRY {
                    # cleanup the borked attempt left half-functioning
                    #Disconnect-ExchangeOnline -confirm:$false ;
                    #Connect-ExchangeOnline @pltCXO ;
                    $smsg = "New-ExoPSSession w`n$(($pltNEXOS|out-string).trim())" ;
                    if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                    $PSSession = New-ExoPSSession @pltNEXOS ;
                    #Add-PSTitleBar $sTitleBarTag ;
                } CATCH {
                    $ErrTrapd=$Error[0] ;
                    $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    #-=-record a STATUSWARN=-=-=-=-=-=-=
                    $statusdelta = ";WARN"; # CHANGE|INCOMPLETE|ERROR|WARN|FAIL ;
                    if(gv passstatus -scope Script -ea 0){$script:PassStatus += $statusdelta } ;
                    if(gv -Name PassStatus_$($tenorg) -scope Script -ea 0){set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta)} ; 
                    #-=-=-=-=-=-=-=-=
                    $smsg = "FULL ERROR TRAPPED (EXPLICIT CATCH BLOCK WOULD LOOK LIKE): } catch[$($ErrTrapd.Exception.GetType().FullName)]{" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    Break #Opts: STOP(debug)|EXIT(close)|CONTINUE(move on in loop cycle)|BREAK(exit loop iteration)|THROW $_/'CustomMsg'(end script with Err output)
                } ; 
            } CATCH [System.Management.Automation.RuntimeException] {
                # see if we can trap the weird blank ConnnectionURI error
                #$pltCXO.Add('ConnectionUri', [string]'https://outlook.office365.com/powershell-liveid/') ;
                $pltNEXOS.Add('ConnectionUri', [string]'https://outlook.office365.com/powershell-liveid/') ;
                $smsg = "'Blank ConnectionUri EXOv2 bug: Retrying with explicit 'ConnectionUri" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } 
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                $smsg = "`n$((get-date).ToString('HH:mm:ss')):Connect-ExchangeOnline w`n$(($pltCXO|out-string).trim())" ;
                if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                TRY {
                    #Disconnect-ExchangeOnline -confirm:$false ;
                    #Connect-ExchangeOnline @pltCXO ;
                    $PSSession = New-ExoPSSession @pltNEXOS ;
                    #Add-PSTitleBar $sTitleBarTag ;
                } CATCH {
                    $ErrTrapd=$Error[0] ;
                    $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    #-=-record a STATUSWARN=-=-=-=-=-=-=
                    $statusdelta = ";WARN"; # CHANGE|INCOMPLETE|ERROR|WARN|FAIL ;
                    if(gv passstatus -scope Script -ea 0){$script:PassStatus += $statusdelta } ;
                    if(gv -Name PassStatus_$($tenorg) -scope Script -ea 0){set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta)} ; 
                    #-=-=-=-=-=-=-=-=
                    $smsg = "FULL ERROR TRAPPED (EXPLICIT CATCH BLOCK WOULD LOOK LIKE): } catch[$($ErrTrapd.Exception.GetType().FullName)]{" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    Break #Opts: STOP(debug)|EXIT(close)|CONTINUE(move on in loop cycle)|BREAK(exit loop iteration)|THROW $_/'CustomMsg'(end script with Err output)
                } ; 
            } CATCH {
                $ErrTrapd=$Error[0] ;
                $smsg = "failed to connect to EXO V2 PS module`nFailed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                #-=-record a STATUSWARN=-=-=-=-=-=-=
                $statusdelta = ";WARN"; # CHANGE|INCOMPLETE|ERROR|WARN|FAIL ;
                if(gv passstatus -scope Script -ea 0){$script:PassStatus += $statusdelta } ;
                if(gv -Name PassStatus_$($tenorg) -scope Script -ea 0){set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta)} ; 
                #-=-=-=-=-=-=-=-=
                $smsg = "FULL ERROR TRAPPED (EXPLICIT CATCH BLOCK WOULD LOOK LIKE): } catch[$($ErrTrapd.Exception.GetType().FullName)]{" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                Break #Opts: STOP(debug)|EXIT(close)|CONTINUE(move on in loop cycle)|BREAK(exit loop iteration)|THROW $_/'CustomMsg'(end script with Err output)
            } ; 
            if ($error.count -ne 0) {
                if ($error[0].FullyQualifiedErrorId -eq '-2144108477,PSSessionOpenFailed') {
                    $smsg = "AUTH FAIL BAD PASSWORD? ABORTING TO AVOID LOCKOUT!" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    #-=-record a STATUSWARN=-=-=-=-=-=-=
                    $statusdelta = ";WARN"; # CHANGE|INCOMPLETE|ERROR|WARN|FAIL ;
                    if(gv passstatus -scope Script -ea 0){$script:PassStatus += $statusdelta } ;
                    if(gv -Name PassStatus_$($tenorg) -scope Script -ea 0){set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta)} ; 
                    #-=-=-=-=-=-=-=-=
                    Break ;
                } ;
            } ;

            if ($PSSession -ne $null ) {

                # hack in coverage to fake use of -UserPrincipalName, which auto-renews sessions, (and creates this global vari to feed renewal), while -Credential use *does not*
                # If UserPrincipal is NULL, but a PSSession exist set variable to refresh token from cache - NICE it pulls the username *right  out  of the session/token!*
                if ([System.String]::IsNullOrEmpty($global:UserPrincipalName) -and (-not [System.String]::IsNullOrEmpty($script:PSSession.Runspace.ConnectionInfo.Credential.UserName))){
                    Write-PSImplicitRemotingMessage ('Set global variable UserPrincialName ...') ; 
                    $global:UserPrincipalName = $script:PSSession.Runspace.ConnectionInfo.Credential.UserName ; 
                } ; 
                # above from: https://ingogegenwarth.wordpress.com/2018/02/02/exo-ps-mfa/

                $PSSessionModuleInfo = Import-PSSession $PSSession -AllowClobber -DisableNameChecking

                # Import the above module globally. This is needed as with using psm1 files,
                # any module which is dynamically loaded in the nested module does not reflect globally.
                Import-Module $PSSessionModuleInfo.Path -Global -DisableNameChecking -Prefix $Prefix -verbose:$false ;
                # haven't checked into what this does - looks like it configures should-reload etc on the tmp_ module
                UpdateImplicitRemotingHandler ;

                # Import the REST module .dll
                $RestPowershellModule = "Microsoft.Exchange.Management.RestApiClient.dll";
                $RestModulePath = [System.IO.Path]::Combine($EOMgmtModulePath, $RestPowershellModule);
                Import-Module $RestModulePath -Cmdlet Set-ExoAppSettings;

                # Set the AppSettings disabling the logging
                #Set-ExoAppSettings -ShowProgress $ShowProgress.Value -PageSize $PageSize.Value -UseMultithreading $UseMultithreading.Value -TrackPerformance $TrackPerformance.Value -ExchangeEnvironmentName $ExchangeEnvironmentName -ConnectionUri $ConnectionUri -AzureADAuthorizationEndpointUri $AzureADAuthorizationEndpointUri -EnableErrorReporting $false ;
                <# 3:18 PM 3/28/2022: Set-ExoAppSettings : A parameter cannot be found that matches parameter name 'ExchangeEnvironmentName'.
                    At C:\Program Files\WindowsPowerShell\Modules\verb-exo\3.2.4\verb-EXO.psm1:2562 char:182
                    + ... kPerformance $TrackPerformance.Value -ExchangeEnvironmentName $Exchan ...
                    +                                          ~~~~~~~~~~~~~~~~~~~~~~~~
                        + CategoryInfo          : InvalidArgument: (:) [Set-ExoAppSettings], ParameterBindingException
                        + FullyQualifiedErrorId : NamedParameterNotFound,Microsoft.Exchange.Management.RestApiClient.SetExoAppSettings
                    #>
                    # checking nope, that param's been dropped since above, remove it:
                    # -ExchangeEnvironmentName $ExchangeEnvironmentName 
                    # I don't see -AzureADAuthorizationEndpointUri either
                    # -AzureADAuthorizationEndpointUri $AzureADAuthorizationEndpointUri 
                #>
                $pltSXAS=[ordered]@{
                  ShowProgress=$false ;
                  PageSize=1000 ;
                  UseMultithreading=$true ;
                  TrackPerformance=$false ;
                  EnableErrorReporting=$false ;
                } ;
                $smsg = "Set-ExoAppSettings w`n$(($pltSXAS|out-string).trim())" ; 
                if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                Set-ExoAppSettings @pltSXAS ;                
                Add-PSTitleBar $sTitleBarTag #-verbose:$($VerbosePreference -eq "Continue");;
            } # if-e $PSSession -ne $null
        } ; #  # if-E $bExistingEXOGood
    } ; # PROC-E
    END {
        if ($bExistingEXOGood -eq $false) {
            # verify the exov2 cmdlets actually imported as a tmp_ module w specifid prefix & 1st cmdlet
            if ( (get-module -name tmp_* | ForEach-Object { Get-Command -module $_.name -name 'Add-xoAvailabilityAddressSpace' -ea 0 }) -AND (test-EXOToken) ) {
                $bExistingEXOGood = $true ;
            } else { $bExistingEXOGood = $false ; }
            # implement caching of accepteddoms into the XXXMeta, in the session (cut back on queries to EXO on acceptedom)
            # swap in non-looping
            if( get-command Get-xoAcceptedDomain) {
                 #$TenOrg = get-TenantTag -Credential $Credential ;
                if(-not (Get-Variable  -name "$($TenOrg)Meta").value.o365_AcceptedDomains){
                    set-Variable  -name "$($TenOrg)Meta" -value ( (Get-Variable  -name "$($TenOrg)Meta").value += @{'o365_AcceptedDomains' = (Get-xoAcceptedDomain).domainname} )
                } ;
            } ;

            if((Get-Variable  -name "$($TenOrg)Meta").value.o365_AcceptedDomains.contains($Credential.username.split('@')[1].tostring())){
                # validate that the connected EXO is to the $Credential tenant
                $smsg = "(EXO Authenticated & Functional:$($Credential.username.split('@')[1].tostring())),($($Credential.username))" ;
                if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                $bExistingEXOGood = $true ;
            # issue: found fresh bug in cxo: svcacct UPN suffix @tenantname.onmicrosoft.com, but testing against AccepteDomain, it's not in there (tho @DOMAIN.mail.onmicrosoft.comis)
            }elseif((Get-Variable  -name "$($TenOrg)Meta").value.o365_TenantDomain -eq ($Credential.username.split('@')[1].tostring())){
                $smsg = "(EXO Authenticated & Functional(TenDom):$($Credential.username.split('@')[1].tostring()))" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                $bExistingEXOGood = $true ;
            } else {
                $smsg = "(Credential mismatch:disconnecting from existing EXO:$($eEXO.Identity) tenant)" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } 
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                Disconnect-exo ;
                $bExistingEXOGood = $false ;
            } ;
        } ;
        $bExistingEXOGood | write-output ;
        # splice in console color scheming
        <# borked by psreadline v1/v2 breaking changes
        if(($PSFgColor = (Get-Variable  -name "$($TenOrg)Meta").value.PSFgColor) -AND ($PSBgColor = (Get-Variable  -name "$($TenOrg)Meta").value.PSBgColor)){
            write-verbose "(setting console colors:$($TenOrg)Meta.PSFgColor:$($PSFgColor),PSBgColor:$($PSBgColor))" ; 
            $Host.UI.RawUI.BackgroundColor = $PSBgColor
            $Host.UI.RawUI.ForegroundColor = $PSFgColor ; 
        } ;
        #>
    }  # END-E
}

#*------^ Connect-EXO2.ps1 ^------