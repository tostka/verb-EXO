#*------v Connect-EXO2.ps1 v------
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
    Github      : https://github.com/tostka/verb-exo
    Tags        : Powershell,ExchangeOnline,Exchange,RemotePowershell,Connection,MFA
    REVISIONS   :
    * 4:08 PM 4/17/2023 ported into connect-exo(), aliased cxo2,connect-exo2 in it.
    * 2:02 PM 4/17/2023 rev: $MinNoWinRMVersion from 2.0.6 => 3.0.0.
    * 2:40 PM 4/5/2023: force the Connect-ExchangeOnline banner hidden:$pltCEO.ShowBanner = $false ;
    * 11:02 AM 4/4/2023 reduced the ipmo and vers chk block, removed the lengthy gmo -list; and any autoinstall. Assume EOM is installed, & break if it's not; added support for EMOv3 get-connectioninfo() etc, and differentiate EMOv2 from EMOv3 connections
    * 3:14 pm 3/29/2023: REN'D $modname => $EOMModName
    * 3:59 PM 8/2/2022 got through dbugging EOM v205 SID interactive pass, working ; added Conn-EOM missing -prefix spec; fully works from mybox w v206p6, cEOM connection, with functional prefix.
    * 4:07 PM 7/26/2022 found that MS code doesn't chk for multi vers's installed, when building .dll paths: wrote in code to take highest version.
    * 3:30 PM 7/25/2022 tests against CBA & SID interactive creds on EOM v205, need to debug now against EOM v206p6, to accomodate PSSession-less connect & test code.
    * 10:50 AM 6/27/2022 missing: $rgxCertThumbprint; validated works with userp interactive mfa
    * 3:27 PM 6/24/2022 dbgd through w x10 connected, looks like it properly disconnects and reconnects; shifted valid code into test-EXOv2Connection(); patched in CBA support
    * 11:27 AM 6/7/2022 cbh cleanup
    * 3:54 PM 4/1/2022 add missing $silent param (had support, but no param)
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
    .PARAMETER usePSSLegacy
    Switch to force use of -UseRPSSession legacy PSSession Basic-Auth connection (new with EMO v2.0.6preview6+; deprecates 5/2023)[-usePSSLegacy]
    .PARAMETER ConnectionUri
    Connection Uri for the Remote PowerShell endpoint [-ConnectionUri 'https://outlook.office365.com/powershell-liveid/']
    .PARAMETER ExchangeEnvironmentName
    Exchange Environment name [-ExchangeEnvironmentName 'O365Default']
    .PARAMETER MinimumVersion
    MinimumVersion required for ExchangeOnlineManagement module (defaults to '2.0.5')[-MinimumVersion '2.0.6']
    .PARAMETER MinNoWinRMVersion
    MinimumVersion required for Non-WinRM connections (of ExchangeOnlineManagement module (defaults to '3.0.0')[-MinimumVersion '2.0.6']
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
    .PARAMETER silent
    Switch to suppress all non-error echos
    .PARAMETER showDebug
    Debugging Flag [-showDebug]
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    None. Returns no objects or output.
    .EXAMPLE
    PS>  Connect-EXO2 -cred $credO365TORSID ;
    Connect using defaults, and leverage any pre-set $global:credo365TORSID variable
    .EXAMPLE
    PS>  Connect-EXO2 -Prefix exolab -credential (Get-Credential -credential user@domain.com)  ;
    Connect an explicit credential, and use 'exolab' as the cmdlet prefix
    .EXAMPLE
    PS>  $cred = get-credential -credential $o365_Torolab_SIDUpn ;
    PS>  Connect-EXO2 -credential $cred ;
    Pass in a prefab credential object (useful for auto-shifting to MFA - the function will autoresolve MFA reqs based on the cred domain)
    .EXAMPLE
    PS>  $pltCXOCThmb=[ordered]@{
    PS>  	CertificateThumbPrint = $credO365TORSIDCBA.UserName ;
    PS>  	AppID = $credO365TORSIDCBA.GetNetworkCredential().Password ;
    PS>  	Organization = 'toroco.onmicrosoft.com' ;
    PS>  	Prefix = 'xo' ;
    PS>  	ShowBanner = $false ;
    PS>  };
    PS>  write-host "Connect-ExchangeOnline w $(($pltCXOCThmb|out-string).trim())" ;
    PS>  Connect-ExchangeOnline @pltCXOCThmb ;
    Example of native connect-ExchangeOnline syntax leveraging a CBA certificate stored locally, with AppID and CertificateThumbPrint pulled from a local global-scope credential object (with AppID stored as password & Thumprint as username)
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
        # implment param for new v206p6+ -UseRPSSession
        [Parameter(HelpMessage = "Switch to force use of -UseRPSSession legacy PSSession Basic-Auth connection (new with EMO v2.0.6preview6+; deprecates 5/2023)[-usePSSLegacy]")]
        [switch] $usePSSLegacy, 
        [Parameter(HelpMessage = "Connection Uri for the Remote PowerShell endpoint [-ConnectionUri 'https://outlook.office365.com/powershell-liveid/']")]
        [string] $ConnectionUri,
        [Parameter(HelpMessage = "Azure AD Authorization endpoint Uri that can issue the OAuth2 access tokens [-AzureADAuthorizationEndpointUri 'https://XXX']")]
        [string] $AzureADAuthorizationEndpointUri,
        [Parameter(HelpMessage = "Exchange Environment name [-ExchangeEnvironmentName 'O365Default']")]
        [Microsoft.Exchange.Management.RestApiClient.ExchangeEnvironment]
        <# error: typedef missing, pre ipmo the mod. 
        Unable to find type [Microsoft.Exchange.Management.RestApiClient.ExchangeEnvironment].
        At D:\scripts\connect-exo2_func.ps1:132 char:9
        +         [Microsoft.Exchange.Management.RestApiClient.ExchangeEnvironm ...
        +         ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            + CategoryInfo          : InvalidOperation: (Microsoft.Excha...angeEnvironment:TypeName) [], RuntimeException
            + FullyQualifiedErrorId : TypeNotFound
        #>
        $ExchangeEnvironmentName = 'O365Default',
        [Parameter(HelpMessage = "MinimumVersion required for ExchangeOnlineManagement module (defaults to '2.0.5')[-MinimumVersion '2.0.6']")]
        [version] $MinimumVersion = '2.0.5',
        [Parameter(HelpMessage = "MinimumVersion required for Non-WinRM connections (of ExchangeOnlineManagement module (defaults to '3.0.0')[-MinimumVersion '2.0.6']")]
        [version] $MinNoWinRMVersion = '3.0.0',
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
        [Parameter(HelpMessage="Silent output (suppress status echos)[-silent]")]
        [switch] $silent,
        [Parameter(HelpMessage = "Debugging Flag [-showDebug]")]
        [switch] $showDebug
    ) ;
    BEGIN {
        $verbose = ($VerbosePreference -eq "Continue") ;

        if(-not (gv rgxCertThumbprint -ea 0)){$rgxCertThumbprint = '[0-9a-fA-F]{40}' ; } ;
        if(-not (gv rgxCertFNameSuffix -ea 0)){$rgxCertFNameSuffix = '-([A-Z]{3})$' ; } ; 

        #*------v PSS & GMO VARIS v------
        # move into a param
        #$MinNoWinRMVersion = '3.0.0' ; 
        # get-pssession session varis
        # select key differentiating properties:
        $pssprops = 'Id','ComputerName','ComputerType','State','ConfigurationName','Availability', 
            'Description','Guid','Name','Path','PrivateData','RootModuleModule', 
            @{name='runspace.ConnectionInfo.ConnectionUri';Expression={$_.runspace.ConnectionInfo.ConnectionUri} },  
            @{name='runspace.ConnectionInfo.ComputerName';Expression={$_.runspace.ConnectionInfo.ComputerName} },  
            @{name='runspace.ConnectionInfo.Port';Expression={$_.runspace.ConnectionInfo.Port} },  
            @{name='runspace.ConnectionInfo.AppName';Expression={$_.runspace.ConnectionInfo.AppName} },  
            @{name='runspace.ConnectionInfo.Credentialusername';Expression={$_.runspace.ConnectionInfo.Credential.username} },  
            @{name='runspace.ConnectionInfo.AuthenticationMechanism';Expression={$_.runspace.ConnectionInfo.AuthenticationMechanism } },  
            @{name='runspace.ExpiresOn';Expression={$_.runspace.ExpiresOn} } ; 
        $EOMmodname = 'ExchangeOnlineManagement' ;
        $EXOv1ConfigurationName = $EXOv2ConfigurationName = $EXoPConfigurationName = "Microsoft.Exchange" ;
        <#
        if(-not (gv EXOv1ConfigurationName -ea 0 )){$EXOv1ConfigurationName = "Microsoft.Exchange" };
        if(-not (gv EXOv2ConfigurationName -ea 0 )){$EXOv2ConfigurationName = "Microsoft.Exchange" };
        if(-not (gv EXoPConfigurationName -ea 0 )){$EXoPConfigurationName = "Microsoft.Exchange" };
        #>
        if(-not (gv EXOv1ComputerName -ea 0 )){$EXOv1ComputerName = 'ps.outlook.com' };
        if(-not (gv EXOv1runspaceConnectionInfoAppName -ea 0 )){$EXOv1runspaceConnectionInfoAppName = '/PowerShell-LiveID'  };
        if(-not (gv EXOv1runspaceConnectionInfoPort -ea 0 )){$EXOv1runspaceConnectionInfoPort -eq '443' };

        if(-not (gv EXOv2ComputerName -ea 0 )){$EXOv2ComputerName = 'outlook.office365.com' ;}
        if(-not (gv EXOv2Name -ea 0 )){$EXOv2Name = "ExchangeOnlineInternalSession*" ; }
        if(-not (gv rgxEXoPrunspaceConnectionInfoAppName -ea 0 )){$rgxEXoPrunspaceConnectionInfoAppName = '^/(exadmin|powershell)$'}; 
        if(-not (gv EXoPrunspaceConnectionInfoPort -ea 0 )){$EXoPrunspaceConnectionInfoPort = '80' } ; 
        # gmo varis
        if(-not (gv rgxExoPsHostName -ea 0 )){ $rgxExoPsHostName = "^(ps\.outlook\.com|outlook\.office365\.com)$" } ;
        if(-not (gv rgxEXOv1gmoDescription -ea 0 )){$rgxEXOv1gmoDescription = "^Implicit\sremoting\sfor\shttps://ps\.outlook\.com/PowerShell" }; 
        if(-not (gv EXOv1gmoprivatedataImplicitRemoting -ea 0 )){$EXOv1gmoprivatedataImplicitRemoting = $true };
        if(-not (gv rgxEXOv2gmoDescription -ea 0 )){$rgxEXOv2gmoDescription = "^Implicit\sremoting\sfor\shttps://outlook\.office365\.com/PowerShell" }; 
        if(-not (gv EXOv2gmoprivatedataImplicitRemoting -ea 0 )){$EXOv2gmoprivatedataImplicitRemoting = $true } ;
        if(-not (gv rgxExoPsessionstatemoduleDescription -ea 0 )){$rgxExoPsessionstatemoduleDescription = '/(exadmin|powershell)$' };
        if(-not (gv EXOv2StateOK -ea 0 )){$EXOv2StateOK = 'Opened'} ; 
        if(-not (gv EXOv2AvailabilityOK -ea 0 )){$EXOv2AvailabilityOK = 'Available'} ; 
        if(-not (gv EXOv2RunStateBad -ea 0 )){ $EXOv2RunStateBad = 'Broken'} ;
        if(-not (gv EXOv1GmoFilter -ea 0 )){$EXOv1GmoFilter = 'tmp_*' } ; 
        if(-not (gv EXOv2GmoNoWinRMFilter -ea 0 )){$EXOv2GmoNoWinRMFilter = 'tmpEXO_*' };
        #*------^ END PSS & GMO VARIS ^------

        # defer to verb-text if avail
        if(-not(get-command test-uri -ea 0)){
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

        # 12:08 PM 8/2/2022 scrap the msal.net material: it's fundementally incompatible with EXO - sure you can pull and auth a token into the PS EXO clientid, but you can't spec a prefix on the returned cmdlets.
        # 4:45 PM 7/7/2022 workaround msal.ps bug: always ipmo it FIRST: "Get-msaltoken : The property 'Authority' cannot be found on this object. Verify that the property exists."
        # admin/SID module auto-install code (myBoxes UID split-perm CU, all else t AllUsers)
        <#$modname = 'MSAL.PS' ;
        $error.clear() ;
        Try { Get-Module -name $modname -listavailable -ErrorAction Stop | out-null } Catch {
            $pltInMod = [ordered]@{Name = $modname ; verbose=$false ;} ;
            if ( $env:COMPUTERNAME -match $rgxMyBoxUID ) { $pltInMod.add('scope', 'CurrentUser') } else { $pltInMod.add('scope', 'AllUsers') } ;
            $smsg = "Install-Module w scope:$($pltInMod.scope)`n$(($pltInMod|out-string).trim())" ;
            if($silent){}elseif ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            Install-Module @pltIMod ;
        } ; # IsInstalled
        $pltIMod = @{Name = $modname ; ErrorAction = 'Stop' ; verbose=$false} ;
        # this forces a specific rev into the ipmo! (and the spec is for EOM, not MSAL.ps!)
        #if ($MinimumVersion) { $pltIMod.add('MinimumVersion', $MinimumVersion.tostring()) } ;
        $error.clear() ;
        Try { Get-Module $modname -ErrorAction Stop | out-null } Catch {
            $smsg = "Import-Module w`n$(($pltIMod|out-string).trim())" ;
            if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            Import-Module @pltIMod ;
        } ; # IsImported
        #>

        # * 11:02 AM 4/4/2023 reduced the ipmo and vers chk block, removed the lengthy gmo -list; and any autoinstall. Assume EOM is installed, & break if it's not
        #region EOMREV ; #*------v EOMREV Check v------
        #$EOMmodname = 'ExchangeOnlineManagement' ;
        $pltIMod = @{Name = $EOMmodname ; ErrorAction = 'Stop' ; verbose=$false} ;
        if($xmod = Get-Module $EOMmodname -ErrorAction Stop){ } else {
            $smsg = "Import-Module w`n$(($pltIMod|out-string).trim())" ;
            if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
            Try {
                Import-Module @pltIMod | out-null ;
                $xmod = Get-Module $EOMmodname -ErrorAction Stop ;
            } Catch {
                $ErrTrapd=$Error[0] ;
                $smsg = "$('*'*5)`nFailed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: `n$(($ErrTrapd|out-string).trim())`n$('-'*5)" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                $smsg = $ErrTrapd.Exception.Message ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                Break ;
            } ;
        } ; # IsImported
        if([version]$xmod.version -ge $MinNoWinRMVersion){
            $MinNoWinRMVersion = $xmod.version.tostring() ;
            $IsNoWinRM = $true ; 
        }
        [boolean]$UseConnEXO = [boolean]([version]$xmod.version -ge $MinNoWinRMVersion) ; 
        #endregion EOMREV ; #*------^ END EOMREV Check  ^------

        if(-not $UseConnEXO){
            # EOM -lt 2.0.5preview6 .dll etc loads, from connect-exchangeonline: (should be installed with the above)
            $EOMgmtModulePath = split-path (get-module $EOMmodname -list | sort Version | select -last 1).Path ;
            if($IsCoreCLR){
                $EOMgmtModulePath = resolve-path -Path $EOMgmtModulePath\netcore ;
                $smsg = "(.netcore path in use:" ; 
            } else { 
                $EOMgmtModulePath = resolve-path -Path $EOMgmtModulePath\netFramework ; 
                $smsg = "(.netnetFramework path in use:" ;                 
            } ;         
            $RestModule = "Microsoft.Exchange.Management.RestApiClient.dll" ;
            $RestModulePath = join-path -path $EOMgmtModulePath -childpath $RestModule  ;
            # paths to proper Module path: Name lists as: Microsoft.Exchange.Management.RestApiClient
            if (-not(get-module $RestModule.replace('.dll',''))) {
                $error.clear() ;
                Import-Module $RestModulePath -verbose:$false -force -ErrorAction 'STOP';
            } ;
            $ExoPowershellGalleryModule = "Microsoft.Exchange.Management.ExoPowershellGalleryModule.dll" ;
            $ExoPowershellGalleryModulePath = join-path -path $EOMgmtModulePath -childpath $ExoPowershellGalleryModule ;
            # full path: C:\Users\USER\Documents\WindowsPowerShell\Modules\ExchangeOnlineManagement\1.0.1\Microsoft.Exchange.Management.ExoPowershellGalleryModule.dll
            if (-not(get-module $ExoPowershellGalleryModule.replace('.dll','') )) {
                $error.clear() ;
                Import-Module $ExoPowershellGalleryModulePath -Verbose:$false -ErrorAction 'stop' ;
            } ;
        
        } else { 
            # $UseConnEXO => we're doing native connect-ExchangeOnline connectivity, no PSSession etc
            $smsg = "native connect-ExchangeOnline specified..." ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        } ; 

    } ; # BEG-E
    PROCESS {
        $bExistingEXOGood = $false ;
        $certUname = $null ; 

        # Keep track of error count at beginning.
        $errorCountAtStart = $global:Error.Count;
        $global:_EXO_TelemetryFilePath = $null;

        <# EXOv1:
        Get-PSSession | fl ConfigurationName,name,state,availability,computername
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
        
        #EXOv2 MFA: 4/4/2022
        TokenProvider          : Microsoft.Exchange.Management.AdminApiProvider.Authentication.MSALTokenProvider
        ConnectionUri          : https://outlook.office365.com:443/PowerShell-LiveID?BasicAuthToOAuthConversion=true&HideBannerMessage=true&ConnectionId=c93cad7f-d8f5-4cce-8ac2-24de6c28518e&ClientProcessId=10808&ExoModuleVersion=2.0.5&OSVersion=
                                 Microsoft+Windows+NT+10.0.14393.0&email=s-email%40domain.com
        PSSessionOption        :
        TokenExpiryTime        : 3/29/2022 8:21:45 PM +00:00
        CurrentModuleName      : tmp_j2itmjec.1iw
        State                  : Opened
        IdleTimeout            : 900000
        OutputBufferingMode    : None
        DisconnectedOn         :
        ExpiresOn              :
        ComputerType           : RemoteMachine
        ComputerName           : outlook.office365.com
        ContainerId            :
        VMName                 :
        VMId                   :
        ConfigurationName      : Microsoft.Exchange
        InstanceId             : 7b793cd7-33de-451d-92a3-bdb3e154bd35
        Id                     : 1
        Name                   : ExchangeOnlineInternalSession_1
        Availability           : Available
        ApplicationPrivateData : {SupportedVersions, ImplicitRemoting, PSVersionTable}
        Runspace               : System.Management.Automation.RemoteRunspace

        -CCMS session via Connect-IPPSSession
        ConfigurationName : Microsoft.Exchange
        ComputerName      : nam02b.ps.compliance.protection.outlook.com
        Name              : ExchangeOnlineInternalSession_1
        State             : Opened
        Availability      : Available
        #>

        <# due to bug in ExchangeOnlineManagement (still in v2.0.5)...
            [Issue using ExchangeOnlineManagement v2.0.4 module to connect to Exchange Online remote powershell (EXO) and Exchange On-Prem remote powershell (EXOP) in same powershell window - Microsoft Q&A - docs.microsoft.com/](https://docs.microsoft.com/en-us/answers/questions/451786/issue-using-exchangeonlinemanagement-v204-module-t.html)
            ...we need to detect and pre-disconnect any existing EXoP implicit remoting sessions
            Because EMO is so badly written it can't properly differentiate the ExOP implicit-remote session(s) from it's own *prior*
            implicit-remote session (which is used for all legacy EXO cmdlets, other than the 9 new 'toy' get-exo[noun] graph-api based cmdlets)
            net-result, if you don't pre-disconnect ExOP implicit-remote pss, EMOs import-pssession cmd throws a 'steppable error' error, 
            commonly, in our case, due to a blank -prefix param, lifted off of the prior PSS connect
            triggered in ExchangeOnlineManagement.psm1:ln143 in global:UpdateImplicitRemotingHandler()
            $PSSessionModuleInfo = Import-PSSession $session -AllowClobber -DisableNameChecking -CommandName $script:MyModule.CommandName -FormatTypeName $script:MyModule.FormatTypeName
            throws:
            ```
            Exception calling "GetSteppablePipeline" with "1" argument(s): "Cannot validate argument on parameter 'Prefix'. The argument is null. Provide a valid value for the argument, and then try running the command again."
            At C:\Users\USER\AppData\Local\Temp\2\tmp_jlykdki2.vpm\tmp_jlykdki2.vpm.psm1:29929 char:13
            +             $steppablePipeline = $scriptCmd.GetSteppablePipeline($myI ...
            +             ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                + CategoryInfo          : NotSpecified: (:) [], ParentContainsErrorRecordException
                + FullyQualifiedErrorId : CmdletInvocationException
            ```
        #>

        if(-not $UseConnEXO){
            
            # all the EXOP PsSession hybrid bug conflicts are only nece3ssary with v2.0.5 or less of EMO...
            $bPreExoPPss= $false ;
            if($pssEXOP = Get-PSSession | 
                    where-object { ($_.ConfigurationName -eq $EXoPConfigurationName) -AND (
                    $_.runspace.ConnectionInfo.AppName -match $rgxEXoPrunspaceConnectionInfoAppName) -AND (
                    $_.runspace.ConnectionInfo.Port -eq $EXoPrunspaceConnectionInfoPort) } ){
                # If EXOP pssession, we're going to have a stepablepipeline conflict, with refresh reconnects of EXOv2 v205 or less
                # this block pre-tags the conflicting Ex10 session, and disconnects it, for later post-refresh.
                $smsg = "(EXOv2 bug-workaround: existing Exchange OnPrem PSSession detected..." ; 
                $smsg += "`nwill remove & reinstate to workaround EMO hybrid coexistance bug)" ; 
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                $bPreExoPPss = $true ; 
                # have to pre-acquire suitable ExOP creds
                #region useEXOP ; #*------v useEXOP v------
                $useEXOP = $true ; 
                if($useEXOP){
                    #*------v GENERIC EXOP CREDS & SRVR CONN BP v------
                    # do the OP creds too
                    $OPCred=$null ;
                    # default to the onprem svc acct
                    $pltGHOpCred=@{TenOrg=$TenOrg ;userrole='ESVC' ;verbose=$($verbose)} ;
                    if($OPCred=(get-HybridOPCredentials @pltGHOpCred).cred){
                        # make it script scope, so we don't have to predetect & purge before using new-variable - nope, still conflicts, need to purge to suppress errors
                        if(get-Variable -Name "cred$($tenorg)OP" -ea 0){
                            set-Variable -Name "cred$($tenorg)OP" -Value $OPCred ;
                        } else { 
                            New-Variable -Name "cred$($tenorg)OP" -Value $OPCred ;
                        } ; 
                        $smsg = "Resolved $($Tenorg) `$OPCred:$($OPCred.username) (assigned to `$cred$($tenorg)OP)" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } else {
                        #-=-record a STATUSERROR=-=-=-=-=-=-=
                        $statusdelta = ";ERROR"; # CHANGE|INCOMPLETE|ERROR|WARN|FAIL ;
                        if(gv passstatus -scope Script){$script:PassStatus += $statusdelta } ;
                        if(gv -Name PassStatus_$($tenorg) -scope Script){set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta)} ; 
                        #-=-=-=-=-=-=-=-=
                        $smsg = "Unable to resolve get-HybridOPCredentials -TenOrg $($TenOrg) -userrole 'ESVC' value!"
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        throw "Unable to resolve $($tenorg) `$OPCred value!`nEXIT!"
                        BREAK ;
                    } ;
                    $smsg= "Using EXOP cred:`$cred$($tenorg)OP:$((Get-Variable -name "cred$($tenorg)OP" ).value.username)" ;  
                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;                 
                    <# CALLS ARE IN FORM: (cred$($tenorg))
                    $pltRX10 = @{
                        Credential = (Get-Variable -name "cred$($tenorg)OP" ).value ;
                        verbose = $($verbose) ; }
                    ReConnect-Ex2010XO @pltRX10 ; # cross-prem conns
                    Reconnect-Ex2010 @pltRX10 ; # local org conns
                    #$pltRx10 creds & .username can also be used for local ADMS connections
                    #>
                    $pltRX10 = @{
                        Credential = (Get-Variable -name "cred$($tenorg)OP" ).value ;
                        verbose = $($verbose) ; silent = $false ; } ;   
                    # enable forest-wide support
                    #enable-forestview -verbose:$($VerbosePreference -eq "Continue") ;
  
                    # TEST
                    #if( !(check-ReqMods $reqMods) ) {write-error "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Missing function. EXITING." ; BREAK ;}  ;
                    # defer cx10/rx10, until just before get-recipients qry
                    #*------^ END GENERIC EXOP CREDS & SRVR CONN BP ^------
                    # connect to ExOP X10
                    <#
                    if($pltRX10){
                        ReConnect-Ex2010 @pltRX10 ;
                    } else { Reconnect-Ex2010 ; } ; 
                    #>
                } ;  # if-E $useEXOP
                #endregion useEXOP ; #*------^ END useEXOP ^------
                Disconnect-Ex2010 -verbose:$($VerbosePreference -eq "Continue") ;
            } ; 
        } else { 
            $smsg = "(native connect-ExchangeOnline specified...)" ; 
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        }; 

        # clear any existing legacy EXO sessions:
        # legacy non-OAuth EXOv2 sessions (AKA EXOv1 basic-auth PSsession-based connections) distinguished on the Computername etc
        if ( $pssEXOv1 = Get-PSSession | 
            where-object {$_.ConfigurationName -like $EXOv1ConfigurationName -AND ($_.ComputerName -eq $EXOv1ComputerName) -AND (
                $_.runspace.ConnectionInfo.AppName -eq $EXOv1runspaceConnectionInfoAppName) -AND (
                $_.runspace.ConnectionInfo.Port -eq $EXOv1runspaceConnectionInfoPort) }  ) {
            # ignore state & Avail, close the conflicting legacy conn's
            if ($pssEXOv1.count -gt 0) {
                $smsg = "(closing $($pssEXOv1.count) legacy EXOv1 sessions...)" ;
                if($silent){}elseif ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                for ($index = 0; $index -lt $pssEXOv1.count; $index++) {
                    $session = $pssEXOv1[$index] ;
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
        #if ( Get-PSSession | where-object { $_.ConfigurationName -like "Microsoft.Exchange" -and $_.Name -like "ExchangeOnlineInternalSession*" -and $_.State -like "*Opened*" -AND ($_.Availability -eq 'Available') -AND ($_.ComputerName -match $rgxExoPsHostName) } ) {
        # simpler - MS uses - very simple detect: 
        # $pssEXOv2 = Get-PSSession | Where-Object {$_.ConfigurationName -like "Microsoft.Exchange" -and $_.Name -like "ExchangeOnlineInternalSession*"} ;
       
        # use test-EXOConnection - cxo2 *only* drives compliant eXOv2 connections, not legacy basicAuth
        #$IsNoWinRM = $false ; 
        $oRet = test-EXOv2Connection -Credential $credential -verbose:$($verbose) ; 
        $bExistingEXOGood = $oRet.Valid ; 
        if($oRet.Valid){
            $pssEXOv2 = $oRet.PsSession ; 
            $IsNoWinRM = $oRet.IsNoWinRM ; 
            $smsg = "(Validated EXOv2 Connected to Tenant aligned with specified Credential)" ; 
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        } else { 
            $smsg = "NO VALID EXOV2/3 PSSESSION FOUND! (DISCONNECTING...)"
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } 
            else{ write-host -ForegroundColor yellow "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
            # capture outlier: shows a session wo the test cmdlet, force reset
            DisConnect-EXO2 ;
            $bExistingEXOGood = $false ;
        } ;     
    
        if ($bExistingEXOGood -eq $false) {
            # open a new EXOv2 session
            if(-not $UseConnEXO){
                # -----------
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

                ==2:43 PM 4/5/2023: V3.1.0 examples
                -------------------------- Example 1 --------------------------
                Connect-ExchangeOnline -UserPrincipalName chris@contoso.com
                This example connects to Exchange Online PowerShell using modern authentication, with or without multi-factor authentication (MFA). We
                aren't using the UseRPSSession parameter, so the connection uses REST and doesn't require Basic authentication to be enabled in WinRM
                on the local computer.
                -------------------------- Example 2 --------------------------
                Connect-ExchangeOnline -UserPrincipalName chris@contoso.com -UseRPSSession
                This example connects to Exchange Online PowerShell using modern authentication, with or without MFA. We're using the UseRPSSession
                parameter, so the connection requires Basic authentication to be enabled in WinRM on the local computer.
                -------------------------- Example 3 --------------------------
                Connect-ExchangeOnline -AppId <%App_id%> -CertificateFilePath "C:\users\navin\Documents\TestCert.pfx" -Organization
                "contoso.onmicrosoft.com"
                This example connects to Exchange Online PowerShell in an unattended scripting scenario using the public key of a certificate.
                -------------------------- Example 4 --------------------------
                Connect-ExchangeOnline -AppId <%App_id%> -CertificateThumbprint <%Thumbprint string of certificate%> -Organization
                "contoso.onmicrosoft.com"
                This example connects to Exchange Online PowerShell in an unattended scripting scenario using a certificate thumbprint.
                -------------------------- Example 5 --------------------------
                Connect-ExchangeOnline -AppId <%App_id%> -Certificate <%X509Certificate2 object%> -Organization "contoso.onmicrosoft.com"
                This example connects to Exchange Online PowerShell in an unattended scripting scenario using a certificate file. This method is best
                suited for scenarios where the certificate is stored in remote machines and fetched at runtime. For example, the certificate is stored
                in the Azure Key Vault.
                -------------------------- Example 6 --------------------------
                Connect-ExchangeOnline -Device
                In PowerShell 7.0.3 or later using version 2.0.4 or later of the module, this example connects to Exchange Online PowerShell in
                interactive scripting scenarios on computers that don't have web browsers.
                The command returns a URL and unique code that's tied to the session. You need to open the URL in a browser on any computer, and then
                enter the unique code. After you complete the login in the web browser, the session in the Powershell 7 window is authenticated via
                the regular Azure AD authentication flow, and the Exchange Online cmdlets are imported after few seconds.
                -------------------------- Example 7 --------------------------
                Connect-ExchangeOnline -InlineCredential
                In PowerShell 7.0.3 or later using version 2.0.4 or later of the module, this example connects to Exchange Online PowerShell in
                interactive scripting scenarios by passing credentials directly in the PowerShell window.

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
                    if($credential.username -match $rgxCertThumbprint){
                        $smsg =  "(UserName:Certificate Thumbprint detected)"
                        if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                        # cert CBA non-basic auth
                        <# CertificateThumbPrint = $Credential.UserName ;
	                        AppID = $Credential.GetNetworkCredential().Password ;
	                        Organization = 'toroco.onmicrosoft.com' ; # org is on $xxxmeta.o365_TenantDomain
                        #>
                        $pltNEXOS.Add("CertificateThumbPrint", [string]$Credential.UserName);                    
                        $pltNEXOS.Add("AppID", [string]$Credential.GetNetworkCredential().Password);
                        if($TenDomain = (Get-Variable  -name "$($TenOrg)Meta").value.o365_TenantDomain){
                            $pltNEXOS.Add("Organization", [string]$TenDomain);
                        } else { 
                            $smsg = "UNABLE TO RESOLVE `$TENORG:$($TenOrg) TO FUNCTIONAL `$$($TenOrg)meta.o365_TenantDomain!" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } 
                            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                            throw $smsg ; 
                            Break ; 
                        } ; 
                        # want the friendlyname to display the cred source in use #$tcert.friendlyname
                        if($tcert = get-childitem -path "Cert:\CurrentUser\My\$($credential.username)"){
                            $certUname = $tcert.friendlyname ; 
                            $certTag = [regex]::match($certUname,$rgxCertFNameSuffix).captures[0].groups[1].value ; 
                            $smsg = "(using CBA:cred:$($certTag):$([string]$tcert.friendlyname))" ; 
                            if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                        } else { 
                            $smsg = "UNABLE TO RESOLVE `$TENORG:$($TenOrg) TO FUNCTIONAL `$$($TenOrg)meta.o365_TenantDomain!" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } 
                            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                            throw $smsg ; 
                            Break ; 
                        } ;
                    
                    } else { 
                        # interactive ModernAuth -UserPrincipalName
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
                    } 
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
                if($certUname){
                    $smsg = "Connecting to EXOv2 w CBA cert:($($certUname))"  ;
                } else { 
                    $smsg = "Connecting to EXOv2:($($credential.username))"  ;
                } ; 
                if($silent){}elseif ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                $smsg = "New-ExoPSSession w`n$(($pltNEXOS|out-string).trim())" ;
                if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                Try {
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

                    # only applies to non CBA cert auth
                    if($credential.username -match $rgxCertThumbprint){
                        $smsg = "(CBA cert auth: skipping `$global:UserPrincipalName populate)" ; 
                        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                    } else { 
                        # hack in coverage to fake use of -UserPrincipalName, which auto-renews sessions, (and creates this global vari to feed renewal), while -Credential use *does not*
                        # If UserPrincipal is NULL, but a PSSession exist set variable to refresh token from cache - NICE it pulls the username *right  out  of the session/token!*
                        if ([System.String]::IsNullOrEmpty($global:UserPrincipalName) -and (-not [System.String]::IsNullOrEmpty($script:PSSession.Runspace.ConnectionInfo.Credential.UserName))){
                            Write-PSImplicitRemotingMessage ('Set global variable UserPrincialName ...') ; 
                            $global:UserPrincipalName = $script:PSSession.Runspace.ConnectionInfo.Credential.UserName ; 
                        } ; 
                        # above from: https://ingogegenwarth.wordpress.com/2018/02/02/exo-ps-mfa/
                    } ; 

                    $PSSessionModuleInfo = Import-PSSession $PSSession -AllowClobber -DisableNameChecking

                    # Import the above module globally. This is needed as with using psm1 files,
                    # any module which is dynamically loaded in the nested module does not reflect globally.
                    $error.clear() ;
                    Import-Module $PSSessionModuleInfo.Path -Global -DisableNameChecking -Prefix $Prefix -verbose:$false -ErrorAction 'stop' ;
                    # haven't checked into what this does - looks like it configures should-reload etc on the tmp_ module
                    UpdateImplicitRemotingHandler ;

                    # Import the REST module .dll
                    $RestPowershellModule = "Microsoft.Exchange.Management.RestApiClient.dll";
                    $RestModulePath = [System.IO.Path]::Combine($EOMgmtModulePath, $RestPowershellModule);
                    $error.clear() ;
                    Import-Module $RestModulePath -Cmdlet Set-ExoAppSettings -erroraction 'Stop' ;

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

            } else { 
                # $UseConnEXO 
                <#
                ==4:21 PM 6/30/2022: v2.0.6p6 examples
                -------------------------- Example 1 --------------------------
                    Connect-ExchangeOnline -UserPrincipalName chris@contoso.com
                    This example connects to Exchange Online PowerShell using modern authentication, with or without multi-factor authentication (MFA). We
                    aren't using the UseRPSSession parameter, so the connection uses REST and doesn't require Basic authentication to be enabled in WinROM
                    on the local computer. But, only the subset of frequently used REST API parameters are available.
                    -------------------------- Example 2 --------------------------
                    Connect-ExchangeOnline -UserPrincipalName chris@contoso.com -UseRPSSession
                    This example connects to Exchange Online PowerShell using modern authentication, with or without MFA. We're using the UseRPSSession
                    parameter, so the connection requires Basic authentication to be enabled in WinRM on the local computer. But, all Exchange Online
                    PowerShell cmdlets are available using traditional remote PowerShell access.
                    -------------------------- Example 3 --------------------------
                    Connect-ExchangeOnline -AppId <%App_id%> -CertificateFilePath "C:\users\navin\Documents\TestCert.pfx" -Organization
                    "contoso.onmicrosoft.com"
                    This example connects to Exchange Online PowerShell in an unattended scripting scenario using the public key of a certificate.
                    -------------------------- Example 4 --------------------------
                    Connect-ExchangeOnline -AppId <%App_id%> -CertificateThumbprint <%Thumbprint string of certificate%> -Organization
                    "contoso.onmicrosoft.com"
                    This example connects to Exchange Online PowerShell in an unattended scripting scenario using a certificate thumbprint.
                    -------------------------- Example 5 --------------------------
                    Connect-ExchangeOnline -AppId <%App_id%> -Certificate <%X509Certificate2 object%> -Organization "contoso.onmicrosoft.com"
                    This example connects to Exchange Online PowerShell in an unattended scripting scenario using a certificate file. This method is best
                    suited for scenarios where the certificate is stored in remote machines and fetched at runtime. For example, the certificate is stored
                    in the Azure Key Vault.
                    -------------------------- Example 6 --------------------------
                    Connect-ExchangeOnline -Device
                    In PowerShell 7.0.3 or later using the EXO V2 module version 2.0.4 or later, this example connects to Exchange Online PowerShell in
                    interactive scripting scenarios on computers that don't have web browsers.
                    The command returns a URL and unique code that's tied to the session. You need to open the URL in a browser on any computer, and then
                    enter the unique code. After you complete the login in the web browser, the session in the Powershell 7 window is authenticated via
                    the regular Azure AD authentication flow, and the Exchange Online cmdlets are imported after few seconds.
                    -------------------------- Example 7 --------------------------
                    Connect-ExchangeOnline -InlineCredential
                    In PowerShell 7.0.3 or later using the EXO V2 module version 2.0.4 or later, this example connects to Exchange Online PowerShell in
                    interactive scripting scenarios by passing credentials directly in the PowerShell window.
                #>

                $pltCEO=[ordered]@{                    
                    erroraction = 'STOP' ;
                    ShowBanner = $false ; # force the fugly banner hidden
                } ;
                
                # 9:43 AM 8/2/2022 add defaulted prefix spec
                if($Prefix){
                    $smsg = "(adding specified Connect-ExchangeOnline -Prefix:$($Prefix))" ; 
                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                    $pltCEO.add('Prefix',$Prefix) ; 
                } ; 

                if ($MFA) {
                    if($credential.username -match $rgxCertThumbprint){
                        $smsg =  "(UserName:Certificate Thumbprint detected)"
                        if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                        # cert CBA non-basic auth
                        <# CertificateThumbPrint = $Credential.UserName ;
	                        AppID = $Credential.GetNetworkCredential().Password ;
	                        Organization = 'toroco.onmicrosoft.com' ; # org is on $xxxmeta.o365_TenantDomain
                        #>
                        $pltCEO.Add("CertificateThumbPrint", [string]$Credential.UserName);                    
                        $pltCEO.Add("AppID", [string]$Credential.GetNetworkCredential().Password);
                        if($TenDomain = (Get-Variable  -name "$($TenOrg)Meta").value.o365_TenantDomain){
                            $pltCEO.Add("Organization", [string]$TenDomain);
                        } else { 
                            $smsg = "UNABLE TO RESOLVE `$TENORG:$($TenOrg) TO FUNCTIONAL `$$($TenOrg)meta.o365_TenantDomain!" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } 
                            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                            throw $smsg ; 
                            Break ; 
                        } ; 
                        # want the friendlyname to display the cred source in use #$tcert.friendlyname
                        if($tcert = get-childitem -path "Cert:\CurrentUser\My\$($credential.username)"){
                            $certUname = $tcert.friendlyname ; 
                            $certTag = [regex]::match($certUname,$rgxCertFNameSuffix).captures[0].groups[1].value ; 
                            $smsg = "(using CBA:cred:$($certTag):$([string]$tcert.friendlyname))" ; 
                            if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                        } else { 
                            $smsg = "UNABLE TO RESOLVE `$TENORG:$($TenOrg) TO FUNCTIONAL `$$($TenOrg)meta.o365_TenantDomain!" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } 
                            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                            throw $smsg ; 
                            Break ; 
                        } ;
                    
                    } else { 
                        # interactive ModernAuth -UserPrincipalName
                        #$pltCXO.Add("UserPrincipalName", [string]$Credential.username);
                        if ($UserPrincipalName) {
                            $pltCEO.Add("UserPrincipalName", [string]$UserPrincipalName);
                            $smsg = "(using cred:$([string]$UserPrincipalName))" ; 
                            if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                        } elseif ($Credential -AND -not $UserPrincipalName){
                            $pltCEO.Add("UserPrincipalName", [string]$Credential.username);
                            $smsg = "(using cred:$($credential.username))" ; 
                            if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                        };
                    } 
                } else {
                    # just use the passed $Credential vari
                    #$pltCXO.Add("Credential", [System.Management.Automation.PSCredential]$Credential);
                    $pltCEO.Add("Credential", [System.Management.Automation.PSCredential]$Credential);
                    $smsg = "(using cred:$($credential.username))" ; 
                    if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                } ;

                $smsg = "Connect-ExchangeOnline w`n$(($pltCEO|out-string).trim())" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                TRY {
                    #Disconnect-ExchangeOnline -confirm:$false ;
                    #Connect-ExchangeOnline @pltCXO ;
                    Connect-ExchangeOnline @pltCEO ;
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

                <# END{} has it's own pass, defer to it, don't need it run 2x
                # use test-EXOConnection
                $oRet = test-EXOv2Connection -Credential $credential -verbose:$($verbose) ; 
                $bExistingEXOGood = $oRet.Valid ; 
                if($oRet.Valid){
                    $pssEXOv2 = $oRet.PsSession ; 
                    $IsNoWinRM = $oRet.IsNoWinRM ; 
                    $smsg = "(Validated EXOv2 Connected to Tenant aligned with specified Credential)" ; 
                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                } else { 
                    $smsg = "NO VALID EXOV2 PSSESSION FOUND! (DISCONNECTING...)"
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } 
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                    # capture outlier: shows a session wo the test cmdlet, force reset
                    DisConnect-EXO2 ;
                    $bExistingEXOGood = $false ;
                } ;     
                #>
                # -------- $UseConnEXO 
            } ; 
        } ; #  # if-E $bExistingEXOGood
    } ; # PROC-E
    END {
        
        $smsg = "Existing PSSessions:`n$((get-pssession|out-string).trim())" ; 
        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 

        if ($bExistingEXOGood -eq $false) {
            
            # defer into test-EXOv2Connection()

            $oRet = test-EXOv2Connection -Credential $credential -verbose:$($verbose) ; 
            $bExistingEXOGood = $oRet.Valid ;
            if($oRet.Valid){
	            $pssEXOv2 = $oRet.PsSession ;
                $IsNoWinRM = $oRet.IsNoWinRM ; 
	            $smsg = "(Validated EXOv2 Connected to Tenant aligned with specified Credential)`n`$IsNoWinRM:$($IsNoWinRM )" ;
	            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
	            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
            } else {
	            $smsg = "NO VALID EXOV2 PSSESSION FOUND! (DISCONNECTING...)"
	            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN }
	            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
	            # capture outlier: shows a session wo the test cmdlet, force reset
	            DisConnect-EXO2 ;
	            $bExistingEXOGood = $false ;
            } ;       

        } else {
            if($bPreExoPPss){
                $smsg = "(EMO bug-workaround: reconnecting prior ExOP PssSession,"
                $smsg += "`nreconnect-Ex2010 -Credential $($pltRX10.Credential.username) -verbose:$($VerbosePreference -eq "Continue"))" ; 
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                reconnect-Ex2010 -Credential $pltRX10.Credential -verbose:$($VerbosePreference -eq "Continue") ; 
            } ;
        } ; 

        if($VerbosePreference -eq "Continue"){
            # echo Exchange-tied PSS summary
            if($pssEXOP = Get-PSSession | 
                where-object { ($_.ConfigurationName -eq $EXoPConfigurationName) -AND (
                    $_.runspace.ConnectionInfo.AppName -match '^/(exadmin|powershell)$') -AND (
                    $_.runspace.ConnectionInfo.Port -eq '80') } ){
                $smsg = "`nExOP PSSessions:`n$(($pssEXOP | fl $pssprops|out-string).trim())" ; 
            } ;
            if($pssEXOv1 = Get-PSSession | 
                    where-object {$_.ConfigurationName -like $EXOv1ConfigurationName -AND (
                        $_.ComputerName -eq 'ps.outlook.com') -AND ($_.runspace.ConnectionInfo.AppName -eq '/PowerShell-LiveID') -AND (
                        $_.runspace.ConnectionInfo.Port -eq '443') }){
                $smsg += "`n`nEXOv1 PSSessions:`n$(($pssEXOv1 | fl $pssprops|out-string).trim())" ; 
            } ; 
            if($pssEXOv2 = Get-PSSession | 
                    where-object {$_.ConfigurationName -like $EXOv2ConfigurationName -AND (
                        $_.Name -like $EXOv2Name) -AND ($_.ComputerName -eq $EXOv2ComputerName)} ){
                $smsg += "`n`nEXOv2 PSSessions:`n$(($pssEXOv2 | fl $pssprops|out-string).trim())" ; 
            } ;
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            if($bPreExoPPss -AND -not $pssEXOP){
                $smsg = "(EMO bug-workaround: reconnecting prior ExOP PssSession,"
                $smsg += "`nreconnect-Ex2010 -Credential $($pltRX10.Credential.username) -verbose:$($VerbosePreference -eq "Continue"))" ; 
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                reconnect-Ex2010 -Credential $pltRX10.Credential -verbose:$($VerbosePreference -eq "Continue") ; 
            } else { 
                $smsg = "(no bPreExoPPss, no Rx10 conn restore)" ; 
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            } ; 

            if($IsNoWinRM -AND ((get-module $EXOv2GmoNoWinRMFilter) -AND (get-module $EOMModName))){
                $smsg = "(native non-WinRM/Non-PSSession-based EXO connection detected." ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
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