﻿# Connect-EXO

#*------v Connect-EXO.ps1 v------
Function Connect-EXO {
    <#
    .SYNOPSIS
    Connect-EXO - Establish connection to Exchange Online (via EXO V2 graph-api module)
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
    * 11:10 AM 6/27/2024 wip updated for functionalized verb-AAD:Update-AADAppRegistrationKeyCertificate(); need to debug the S&C conn, haven't revisited since initial hybrid coding attempt ; odd, it lost the cxo alias def (added back, did I lose a rev in the mix?) CBA certs expired, error in connect-ExchangeOnline doesn't cite the expiration, just crashes out. So added code to precheck local cert NotAfter, and premptively feed problem cert into Update-AADAppRegistrationKeyCertificate 
    (not debugged yet; need to reroll the certs & creds)
    * 4:28 PM 6/26/2024 interrum, functional 
    * 9:55 AM 6/21/2024 add: prereq checks, and $isBased support, to devert into most basic Get-ConnectionInformation , Connect-ExchangeOnline fall back support
    * 11:26 AM 4/12/2024 validated connect-exo -prefix xo -verbose ; 
    * 9:09 AM 4/2/2024 rem'd citations of $bPreExoPPss
    * 1:05 PM 4/1/2024 validates functional jb    
    * 1:48 PM 3/1/2024  with v340 support for proper/native S&C conn, I can finally remove the legacy EOM connection bits from this (*substantial* simplification):
        - removed raft of pre EOMv3xx code, basic auth is fully blocked now, independantly, test-EXOv2Connection() got some updates (TenOrg tweak, likewise removed code < EOM3xx support)
    * 2:51 PM 2/26/2024 add | sort version | select -last 1  on gmos, LF installed 3.4.0 parallel to 3.1.0 and broke auth: caused mult versions to come back and conflict with the assignement of [version] type (would require [version[]] to accom both, and then you get to code everything for mult handling)
    * 1:32 PM 5/30/2023 Updates to support either -Credential, or -UserRole + -TenOrg, to support fully portable downstream credentials: 
        - Add -UserRole & explicit -TenOrg params; working. 
        - Drive TenOrg defaulted $global:o365_TenOrgDefault, or on $env:userdomain
        - use the combo thru get-TenantCredential(), then set result to $Credential
        - if using Credential, the above are backed out via get-TenantTag() on the $credential 
        - CBA identifiers are resolve always via $uRoleReturn = resolve-UserNameToUserRole -Credential $Credential ;
    * 2:02 PM 5/25/2023 updated CBH example to reflect use of $pltRXOC std conn splat
    * 1:08 PM 5/23/2023 fixed typo (-eq vs =, dumping $false into pipe)     
    * 4:15 PM 5/22/2023 cleaned cbh ; removed some rem'd 
    * 10:06 AM 5/19/2023 add: code to run resolve-UserNameToUserRole  wi $Credential or $UserPrincipalName; sub'd out direct cert-parsing & value assignements with resolve-UserNameToUserRole outputs; 
    * 3:21 PM 4/25/2023 add CBA CBH demo ; added code to pass through calc'd $CertTag as test-EXOv2Connection() -CertTag (used for validating credential alignment w Tenant)
    * 10:59 AM 4/18/2023 step debugs ; 
    * 10:16 AM 4/18/2023 rem'd out unused $ConnectionUri;$AzureADAuthorizationEndpointUri;$PSSessionOption;$BypassMailboxAnchoring;$DelegatedOrganization;
    rem'd boolean dump into pipeline in END{}
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
    Connect-EXO - Establish connection to Exchange Online (via EXO V2 graph-api module)
    .PARAMETER  Prefix
    [verb]-PREFIX[command] PREFIX string for clearly marking cmdlets sourced in this connection [-Prefix tag]
    .PARAMETER Credential
    Credential to use for this connection [-credential [credential obj variable]
    .PARAMETER UserPrincipalName
    User Principal Name or email address of the user
    .PARAMETER UserRole
    Credential Optional User Role spec for credential discovery (wo -Credential)(SID|CSID|UID|B2BI|CSVC|ESVC|LSVC|ESvcCBA|CSvcCBA|SIDCBA)[-UserRole @('SIDCBA','SID','CSVC')]
    .PARAMETER TenOrg
    Optional Tenant Tag (wo -Credential)[-TenOrg 'XYZ']
    .PARAMETER ExchangeEnvironmentName
    Exchange Environment name [-ExchangeEnvironmentName 'O365Default']
    .PARAMETER MinimumVersion
    MinimumVersion required for ExchangeOnlineManagement module (defaults to '3.0.0')[-MinimumVersion '2.0.6']
    .PARAMETER MinNoWinRMVersion
    MinimumVersion required for Non-WinRM connections (of ExchangeOnlineManagement module (defaults to '3.0.0')[-MinimumVersion '2.0.6']
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
    PS>  Connect-EXO -cred $credO365TORSID ;
    Connect using defaults, and leverage any pre-set $global:credo365TORSID variable
    .EXAMPLE
    PS>  Connect-EXO -Prefix exolab -credential (Get-Credential -credential user@domain.com)  ;
    Connect an explicit credential, and use 'exolab' as the cmdlet prefix
    .EXAMPLE 
    connect-exo2 -credential $credO365xxxCBA -verbose ; 
    Connect using a CBA credential variable (prestocked from profile automation). Script opens and recycles the cred cert specs emulating the native CBA connection below, but pulling source info from a stored dpapi-encrypted .xml credential file.
    .EXAMPLE
    connect-exo -UserRole SIDCBA -TenOrg ABC -verbose  ; 
    Demo use of UserRole (specifying a CBA variant), AND TenOrg spec, to connect (autoresolves against preconfigured credentials in profile)
    .EXAMPLE
    PS>  $cred = get-credential -credential $o365_Torolab_SIDUpn ;
    PS>  Connect-EXO -credential $cred ;
    Pass in a prefab credential object (useful for auto-shifting to MFA - the function will autoresolve MFA reqs based on the cred domain)
    .EXAMPLE
    PS> $pltRXOC = [ordered]@{
    PS>     Credential = $Credential ;
    PS>     verbose = $($VerbosePreference -eq "Continue")  ;
    PS>     Silent = $silent ; 
    PS> } ;
    PS> if ($script:useEXOv2 -OR $useEXOv2) { Connect-EXO2 @pltRXOC }
    PS> else { Connect-EXO @pltRXOC } ;    
    Splatted example leveraging prefab $pltRXOC splat, derived from local variables & $VerbosePreference value.
    .EXAMPLE
    PS>  $pltCXOCThmb=[ordered]@{
    PS>  	CertificateThumbPrint = $credO365TORSIDCBA.UserName ;
    PS>  	AppID = $credO365TORSIDCBA.GetNetworkCredential().Password ;
    PS>  	Organization = 'TENANTNAME.onmicrosoft.com' ;
    PS>  	Prefix = 'xo' ;
    PS>  	ShowBanner = $false ;
    PS>  };
    PS>  write-host "Connect-ExchangeOnline w $(($pltCXOCThmb|out-string).trim())" ;
    PS>  Connect-ExchangeOnline @pltCXOCThmb ;
    Example of native connect-ExchangeOnline syntax leveraging a CBA certificate stored locally, with AppID and CertificateThumbPrint pulled from a local global-scope credential object (with AppID stored as password & Thumprint as username)
    .LINK
    #>
    [CmdletBinding(DefaultParameterSetName='UPN')]
    [Alias('cxo','cxo2','Connect-EXO2' )]
    PARAM(
        # try pulling all the ParameterSetName's - just need to get through it now. - no got through it with a defaultparametersetname (avoids 
        [Parameter(HelpMessage = "[verb]-PREFIX[command] PREFIX string for clearly marking cmdlets sourced in this connection [-Prefix tag]")]
            [string]$Prefix = 'xo',
        [Parameter(ParameterSetName = 'Cred', HelpMessage = "Credential to use for this connection [-credential [credential obj variable]")]
            [System.Management.Automation.PSCredential]$Credential,
            # = $global:credo365TORSID, # defer to TenOrg & UserRole resolution
        [Parameter(ParameterSetName = 'UPN',HelpMessage = "User Principal Name or email address of the user[-UserPrincipalName logon@domain.com]")]
            [string]$UserPrincipalName,
        [Parameter(Mandatory = $false, HelpMessage = "Credential User Role spec (SID|CSID|UID|B2BI|CSVC|ESVC|LSVC|ESvcCBA|CSvcCBA|SIDCBA)[-UserRole @('SIDCBA','SID','CSVC')]")]
            # sourced from get-admincred():#182: $targetRoles = 'SID', 'CSID', 'ESVC','CSVC','UID','ESvcCBA','CSvcCBA','SIDCBA' ; 
            #[ValidateSet("SID","CSID","UID","B2BI","CSVC","ESVC","LSVC","ESvcCBA","CSvcCBA","SIDCBA")]
            # pulling the pattern from global vari w friendly err
            [ValidateScript({
                if(-not $rgxPermittedUserRoles){$rgxPermittedUserRoles = '(SID|CSID|UID|B2BI|CSVC|ESVC|LSVC|ESvcCBA|CSvcCBA|SIDCBA)'} ;
                if(-not ($_ -match $rgxPermittedUserRoles)){throw "'$($_)' doesn't match `$rgxPermittedUserRoles:`n$($rgxPermittedUserRoles.tostring())" ; } ; 
                return $true ; 
            })]
            [string[]]$UserRole = @('SIDCBA','SID','CSVC'),
            # CCMS: [string[]]$UserRole = @('SID'),
        [Parameter(Mandatory=$FALSE,HelpMessage="TenantTag value, indicating Tenants to connect to[-TenOrg 'TOL']")]
            [ValidateNotNullOrEmpty()]
            #[ValidatePattern("^\w{3}$")]
            [string]$TenOrg = $global:o365_TenOrgDefault,
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

        #region CHKPREREQ ; #*------v CHKPREREQ v------
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
                if(-not (gv -name "$($met)Meta" -ea 0).value[$mp]){
                    $isBased = $false; $ppMiss += "$($met)Meta.$($mp)" ; 
                } ; 
            } ; 
        } ; 
        if($gvmiss){write-warning "Missing Dependant Meta variables:`n$(($gvMiss |%{"`$$($_)" }) -join ',')" } ; 
        if($ppMiss){write-warning "Missing Dependant Meta vari properties:`n$(($ppMiss |%{"`$$($_)" }) -join ',')" } ; 
        if(-not $isBased){ write-warning  "missing critical dependancy profile config!" } ; 
        #endregion CHKPREREQ ; #*------^ END CHKPREREQ ^------

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
        if(-not (gv EXOv1ComputerName -ea 0 )){$EXOv1ComputerName = 'ps.outlook.com' };
        if(-not (gv EXOv1runspaceConnectionInfoAppName -ea 0 )){$EXOv1runspaceConnectionInfoAppName = '/PowerShell-LiveID'  };
        if(-not (gv EXOv1runspaceConnectionInfoPort -ea 0 )){$EXOv1runspaceConnectionInfoPort = '443' };

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
        # add get-connectioninformation.ConnectionURI targeting rgxs for CCMS vs EXO
        if(-not $rgxConnectionUriEXO){$rgxConnectionUriEXO = 'https://outlook\.office365\.com'} ; 
        if(-not $rgxConnectionUriEXO){$rgxConnectionUriCCMS = 'https://ps\.compliance\.protection\.outlook\.com'} ; 
        $sTitleBarTag = @("EXO2") ;
        #*------^ END PSS & GMO VARIS ^------

        #*======v FUNCTIONS v======
        if(-not(get-command test-uri -ea 0)){
            #*------v Function Test-Uri v------
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
            #*------^ END Function Test-Uri ^------
        } ;
        #*======^ END FUNCTIONS ^======

        
        # disable prefix spec, unless actually blanked (e.g. centrally spec'd in profile).
        if (-not $Prefix) {
            $Prefix = 'xo' ; # 4:31 PM 7/29/2020 MS has RESERVED use of the 'exo' prefix [facepalm]
            #$Prefix = 'cc' ; # ccms variant
            $smsg = "(asserting Prefix:$($Prefix)" ;
            if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        } ;
        if (($Prefix) -and ($Prefix -eq 'EXO')) {
            throw "Prefix 'EXO' is a reserved Prefix, please use a different prefix."
        }
        if($Prefix -eq 'cc'){
            # build in hybrid xo & ccms support, switch on the prefix spec
            $useCCMSConn = $true ; 
        }; 
        if($useCCMSConn){
            # respec userrole
            $UserRole = @('SID') ; 
            $sTitleBarTag = @("CCMS") ;
        } ; 

        <#
        $TenOrg = get-TenantTag -Credential $Credential ;
        if($Credential){
            $uRoleReturn = resolve-UserNameToUserRole -Credential $Credential
        } elseif($UserPrincipalName){
            $uRoleReturn = resolve-UserNameToUserRole -UserName $UserPrincipalName
        } ; 
        if($uRoleReturn.TenOrg){
            $CertTag = $uRoleReturn.TenOrg
        } ; 
        #>

        if(-not $isBased){
            # default to most basic rudimentary connection
            $Status = Get-ConnectionInformation -ErrorAction SilentlyContinue
            If (-not ($Status)) {
              #Connect-ExchangeOnline -SkipLoadingCmdletHelp
              Connect-ExchangeOnline -SkipLoadingCmdletHelp -ShowBanner:$false ; 
            }
        }else {

            # transplat fr rxo ---
            if(-not $Credential){
                if($UserRole){
                    $smsg = "Using specified -UserRole:$( $UserRole -join ',' )" ;
                    if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                } else { $UserRole = @('SID','CSVC') } ;
                if($TenOrg){
                    $smsg = "Using explicit -TenOrg:$($TenOrg)" ; 
                } else { 
                    switch -regex ($env:USERDOMAIN){
                        ([regex]('(' + (( @($TORMeta.legacyDomain,$CMWMeta.legacyDomain)  |foreach-object{[regex]::escape($_)}) -join '|') + ')')).tostring() {$TenOrg = $env:USERDOMAIN.substring(0,3).toupper() } ;
                        $TOLMeta.legacyDomain {$TenOrg = 'TOL' }
                        default {throw "UNRECOGNIZED `$env:USERDOMAIN!:$($env:USERDOMAIN)" ; exit ; } ;
                    } ;  
                    $smsg = "Imputed `$TenOrg from logged on USERDOMAIN:$($TenOrg)" ;             
                } ; 
                if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;             
            
                $sTitleBarTag += $TenOrg ;

                $o365Cred = $null ;
                $pltGTCred=@{TenOrg=$TenOrg ; UserRole= $UserRole; verbose=$($verbose)} ;
                $smsg = "get-TenantCredentials w`n$(($pltGTCred|out-string).trim())" ; 
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level verbose } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                $o365Cred = get-TenantCredentials @pltGTCred ;

                if($o365Cred.credType -AND $o365Cred.Cred -AND $o365Cred.Cred.gettype().fullname -eq 'System.Management.Automation.PSCredential'){
                    $smsg = "(validated `$o365Cred contains .credType:$($o365Cred.credType) & `$o365Cred.Cred.username:$($o365Cred.Cred.username)" ; 
                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                    $Credential = $o365Cred.Cred ;
                } else { 
                    $smsg = "UNABLE TO RESOLVE FUNCTIONAL CredType/UserRole from specified explicit -Credential:$($Credential.username)!" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 

                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                    break ; 
                } ; 
            }elseif(-not $Credential -AND -not $isBaseed){    
                $smsg = "Missing Profile config to drive connection automation, defa" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
            } else { 
                # test-exotoken only applies if $UseConnEXO  $false
                $TenOrg = get-TenantTag -Credential $Credential ;
            } ;
            # build the cred etc once, for all below:
            $pltCXO=[ordered]@{
                Credential = $Credential ;
                verbose = $($verbose) ; 
                erroraction = 'STOP' ;
            } ;
            if((gcm connect-EXO).Parameters.keys -contains 'silent'){
                $pltCXO.add('Silent',$false) ;
            } ;

            $uRoleReturn = resolve-UserNameToUserRole -Credential $Credential ; 
            if($credential.username -match $rgxCertThumbprint){
                $certTag = $uRoleReturn.TenOrg ; 
            } ; 
            # ---

            $MFA = get-TenantMFARequirement -Credential $Credential ;

            # 12:08 PM 8/2/2022 scrap the msal.net material: it's fundementally incompatible with EXO - sure you can pull and auth a token into the PS EXO clientid, but you can't spec a prefix on the returned cmdlets.
            # 4:45 PM 7/7/2022 workaround msal.ps bug: always ipmo it FIRST: "Get-msaltoken : The property 'Authority' cannot be found on this object. Verify that the property exists."

            # * 11:02 AM 4/4/2023 reduced the ipmo and vers chk block, removed the lengthy gmo -list; and any autoinstall. Assume EOM is installed, & break if it's not
            #region EOMREV ; #*------v EOMREV Check v------
            # reduced the ipmo and vers chk block, removed the lengthy gmo -list; and any autoinstall. Assume EOM is installed, & break if it's not
            $pltIMod = @{Name = $EOMmodname ; ErrorAction = 'Stop' ; verbose=$false} ;
            if($xmod = Get-Module $EOMmodname -ErrorAction Stop| sort version | select -last 1 ){ } else {
                $smsg = "Import-Module w`n$(($pltIMod|out-string).trim())" ;
                if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                Try {
                    Import-Module @pltIMod | out-null ;
                    $xmod = Get-Module $EOMmodname -ErrorAction Stop | sort version | select -last 1 ;
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
                $smsg = "NON-connect-ExchangeOnline() version of ExchangeOnlineManagement installed, update to vers:$($MinNoWinRMVersion) or higher!" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                throw $smsg ; 
                break ; 

                # dropping all support/legacy code for EOMv1 (PSSession native-only connections)
                # code below is used *solely* for EOM v205 connections (uses cached creds, integrates Connect-ExchangeOnline underlying commands)
                # EOM -lt 2.0.5preview6 .dll etc loads, from connect-exchangeonline: (should be installed with the above)
                # removed 12:23 PM 3/1/2024
        
            } else { 
                # $UseConnEXO => we're doing native connect-ExchangeOnline connectivity, no PSSession etc
                $smsg = "native connect-ExchangeOnline specified..." ; 
                if($silent){}elseif ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } ; 

        }  # if-E $isBased 

    } ; # BEG-E
    PROCESS {
        if($isBased){

            $bExistingEXOGood = $bExistingCCMSGood = $false ;
            $certUname = $null ; 

            # Keep track of error count at beginning.
            $errorCountAtStart = $global:Error.Count;
            $global:_EXO_TelemetryFilePath = $null;

                                                                                                                                                                                                                        <# EXOv1: fully deprecated 12:24 PM 3/1/2024
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

                                                                                <#
        if(-not $UseConnEXO){
            
            # all the EXOP PsSession hybrid bug conflicts are only nece3ssary with v2.0.5 or less of EMO...

            $bPreExoPPss= $false ;
            $smsg = "NON-connect-ExchangeOnline() version of ExchangeOnlineManagement installed, update to vers:$($MinNoWinRMVersion) or higher!" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
            throw $smsg ; 
            break ; 
            # removed all legacy code: 12:25 PM 3/1/2024
            
        } else { 
            $smsg = "(native connect-ExchangeOnline specified...)" ; 
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        }; 
        #>

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
            # use test-EXOConnection - cxo2 *only* drives compliant eXOv2 connections, not legacy basicAuth
            #$IsNoWinRM = $false ; 
            # 11:18 AM 4/25/2023 add support for passing calc'd CertTag "Cert FriendlyName Suffix to be used for validating credential alignment(Optional but required for CBA calls)[-CertTag `$certtag]")][string]$CertTag
            $tIn = '4/18/2024;D:\scripts\test-EXOv2Connection_func.ps1' ;
            $tdt,$tsrc =  $tIn.split(';') ;
            $tdt=[datetime]$tdt ;
            if($psise -and (get-date ).date -eq $tdt){
                $gcm = gcm (split-path $tsrc -leaf).replace('_func.ps1','') ;
                if( $gcm -AND $gcm.source -ne ''){
                    gci function:$((split-path $tsrc -leaf).replace('_func.ps1','')) -ea 0| remove-item -force ;
                    ipmo -fo -verb $tsrc;
                } else {write-host "(non-Mod vers loaded)"} ;
            } ; 
            if($CertTag -ne $null){
                $smsg = "(specifying detected `$CertTag:$($CertTag))" ; 
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                # keeps diverting to verb-exo, put in a today-only test/force reload _func
                if($psise -AND ( (gcm test-EXOv2Connection).parameters.keys -notcontains 'Prefix' )){
                    Do{
                        gci function:test-EXOv2Connection | remove-item -force ; 
                        #ipmo -fo -verb D:\scripts\Connect-EXO_func.ps1
                        epbp ; $psise.CurrentFile.FullPath | ipmo -fo -verb ; 
                    }while($null -ne (gcm test-EXOv2Connection).Module)
                        
                    #gci function:test-EXOv2Connection | remove-item -force ; 
                    #ipmo -fo -verb D:\scripts\Connect-EXO_func.ps1
                } ; 
                $oRet = test-EXOv2Connection -Credential $credential -CertTag $certtag -Prefix $Prefix -verbose:$($verbose) ; 
            } else { 
                $oRet = test-EXOv2Connection -Credential $credential -Prefix $Prefix -verbose:$($verbose) ; 
            } ; 
            <# 1:46 PM 4/5/2024 updated test-EXOv2Connection return obj:
            $oReturn = [ordered]@{
                PSSession = $null ; 
                IsNoWinRM = $false ; 
                Valid = $false ; 
                Prefix = $Prefix ; 
                isEXO = $false ; 
                isCCMS = $false ;
                ConnectionUri = $null ; 
            } ; 
            #>
            if($oReturn.isEXO){
                $bExistingEXOGood = $oRet.Valid ; 
                $smsg = "(Validated EXOv2 Connected to Tenant aligned with specified Credential)" ;
            }elseif($oReturn.isCCMS){
                $bExistingCCMSGood = $oRet.Valid ; 
                $smsg = "(Validated CCMSv2 Connected to Tenant aligned with specified Credential)" ;
            } ; 
            if($oRet.Valid){
                $pssEXOv2 = $oRet.PsSession ; 
                $IsNoWinRM = $oRet.IsNoWinRM ; 
                #$smsg = "(Validated EXOv2 Connected to Tenant aligned with specified Credential)" ; 
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            } else { 
                $smsg = "NO VALID EXOV2/3 PSSESSION FOUND! (DISCONNECTING...)"
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } 
                else{ write-host -ForegroundColor yellow "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                # capture outlier: shows a session wo the test cmdlet, force reset
                $pltDXO=[ordered]@{
                    confirm = $false ;
                    erroraction = 'STOP' ;
                    whatif = $($whatif) ;
                } ;
                if($Prefix){
                    $pltDXO.add('ModulePrefix',$Prefix) ;
                }
                $smsg = "Disconnect-ExchangeOnline w`n$(($pltDXO|out-string).trim())" ;
                if($silent){}elseif ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                Disconnect-ExchangeOnline @pltDXO ;
                #  DisConnect-CCMS ; # CCMS
                #DisConnect-EXO ;
                $bExistingEXOGood = $bExistingCCMSGood = $false ;
            } ;     
        
            #$bExistingCCMSGood
            #if ($bExistingEXOGood -eq $false) {
            # $UseConnEXO indicates it's a MinNoWinRMVersino, not necc xo-only conn; $useCCMSConn indicates it's a prefix cc/CCMS connection, solely
            if( ($useCCMSConn -AND ($bExistingCCMSGood -eq $false)) -OR (-not($useCCMSConn) -AND $bExistingEXOGood -eq $false) ){
                # open a new EXOv2 session
                if(-not $UseConnEXO){
                
                    # removed all legacy code: 12:25 PM 3/1/2024

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
                    <# CCMS connect
                    ==2:04 PM 4/1/2024: v3.4.0 examples
                    -------------------------- Example 1 --------------------------
                    Connect-IPPSSession -UserPrincipalName michelle@contoso.onmicrosoft.com
                    This example connects to Security & Compliance PowerShell using the specified account and modern authentication, with or without MFA. In v3.2.0 or later of the module, we're connecting in REST API mode, so Basic authentication in WinRM isn't required on the
                    local computer.
                    -------------------------- Example 2 --------------------------
                    Connect-IPPSSession -UserPrincipalName michelle@contoso.onmicrosoft.com -UseRPSSession
                    This example connects to Security & Compliance PowerShell using the specified account and modern authentication, with or without MFA. In v3.2.0 or later of the module, we're connecting in remote PowerShell mode, so Basic authentication in WinRM is required
                    on the local computer.
                    -------------------------- Example 3 --------------------------
                    Connect-IPPSSession -AppId <%App_id%> -CertificateThumbprint <%Thumbprint string of certificate%> -Organization "contoso.onmicrosoft.com"
                    This example connects to Security & Compliance PowerShell in an unattended scripting scenario using a certificate thumbprint.
                    -------------------------- Example 4 --------------------------
                    Connect-IPPSSession -AppId <%App_id%> -Certificate <%X509Certificate2 object%> -Organization "contoso.onmicrosoft.com"
                    This example connects to Security & Compliance PowerShell in an unattended scripting scenario using a certificate file. This method is best suited for scenarios where the certificate is stored in remote machines and fetched at runtime. For example, the
                    certificate is stored in the Azure Key Vault.            
                    #>

                    $pltCEO=[ordered]@{                    
                        erroraction = 'STOP' ;
                        ShowBanner = $false ; # force the fugly banner hidden
                    } ;
                
                    # 9:43 AM 8/2/2022 add defaulted prefix spec
                    if($Prefix){
                        if($useCCMSConn){
                            $smsg = "(adding specified  Connect-IPPSSession -Prefix:$($Prefix))" ; 
                        } else { 
                            $smsg = "(adding specified Connect-ExchangeOnline -Prefix:$($Prefix))" ; 
                        } ; 
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
	                            Organization = 'TENANTNAME.onmicrosoft.com' ; # org is on $xxxmeta.o365_TenantDomain
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
                            <# want the friendlyname to display the cred source in use #$tcert.friendlyname
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
                            #>
                            $certUname = $uRoleReturn.FriendlyName ; 
                            $certTag = $uRoleReturn.TenOrg ; 

                            <# 9:35 AM 6/25/2024 expired auth cert, need to proactively test and warn
                            # CertificateThumbPrint = $Credential.UserName ;
	                            AppID = $Credential.GetNetworkCredential().Password ;
	                            Organization = 'TENANTNAME.onmicrosoft.com' ; # org is on $xxxmeta.o365_TenantDomain
                            # warn at 2wks
                            # warn high pause at 7 days
                            #if((gci Cert:\CurrentUser\My\[string]$Credential.UserName).NotAfter -lt (get-date )){write-warning "Expired Cert!"} ;
                            "cert:$(((gci Cert:\CurrentUser\My\C5672B2D81CC828F78A93CE81CF436CC8C861F8F -ea STOP).pspath -split('::'))[1])"
                            #>
                            $prpCertgci = 'FriendlyName','Subject','Thumbprint','NotBefore','NotAfter',@{Name='Path';Expression={( "cert:$(($_.pspath -split('::'))[-1])" )}} ; 
                            $certWarnDays = 14 ; 
                            $certAlarmDays = 7 ; 
                            $oCert = gci "Cert:\CurrentUser\My\$([string]$Credential.UserName)" -ea STOP ; 
                            $certLifeDays = (new-timespan -start (get-date ) -end $oCert.NotAfter -ea STOP).days ; `
                            $hsRollCert = @"

## To roll over manually out of band:

gci "Cert:\CurrentUser\My\$([string]$Credential.UserName)" | Update-AADAppRegistrationKeyCertificate 

"@ ; 
                            if($certLifeDays -lt $certAlarmDays){
                                $smsg = "`n`n*** CERTIFICATE $($Credential.UserName) ($($certUname)) EXPIRES IN $($certLifeDays) DAYS! ***" ; 
                                $smsg += "`n$(($oCert | fl $prpCertgci |out-string).trim())`n`n" ; 
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 

                                $smsg = "DO YOU WANT TO ROLLOVER AND REPLACE THE CERTIFICATE & KEYCRED ON THE APP REGISTRATION? " ; 
                                if($certLifeDays -lt 0){
                                    $SMSG += "`nCERT IS ALREADY EXPIRED, THIS PROCESS WILL CRASH OUT UNTIL YOU REPLACE THE CERT!" 
                                } ; 
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Prompt } 
                                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                                $bRet=Read-Host "Enter YYY to continue. Anything else will exit"  ; 
                                if ($bRet.ToUpper() -eq "YYY") {
                                    $smsg = "(Moving on)" ; 
                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                                    #code to rollover cert
                                    # updated name, verb-aad:Update-AADAppRegistrationKeyCertificate()
                                    #if($rolltool = (get-command -name Rollover-AADAppRegistrationCBAAuth.ps1 -ea STOP ).source){
                                    if(get-command Update-AADAppRegistrationKeyCertificate){
                                        #. $rolltool -certificate $ocert  ; 
                                        # another: & runs the script in it's own scope
                                        #& "C:\AzureFileShare\MEDsys\Powershell Scripts\B.ps1" -ServerName medsys-dev ; 
                                        #$smsg = "RUNNING:`n& $($rolltool) -certificate `$ocert ; " ; 
                                        #& $rolltool -certificate $ocert ; 
                                        # shift to func
                                        $smsg = "Running:`nUpdate-AADAppRegistrationKeyCertificate -certificate `$ocert " ; 
                                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                        #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                                        #Update-AADAppRegistrationKeyCertificate -certificate $ocert
                                        if($results = Update-AADAppRegistrationKeyCertificate -certificate $ocert){
                                            if($results.Certificate){ 
                                                $smsg = "Updated Certificate`n$(($results.Certificate| ft -a Subject,NotAfter,Thumbprint|out-string).trim())" ; 
                                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                                                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                                #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                                            }else{
                                                $smsg = "NO SUMMARY CERTIFICATE RETURNED!" ; 
                                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                                                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                                            } ; 
                                        }else{
                                            $smsg = "NO SUMMARY RETURNED!" ; 
                                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                                            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 

                                        } ; 
                                        
                                    } else {
                                        $smsg = "Unable to: get-command Update-AADAppRegistrationKeyCertificate!" ; 
                                        $smsg += "`nManually resolve location issue and run:" ; 
                                        $smsg += $hsRollCert ; 
                                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                                    } ;
                                } else {
                                    $smsg = "(Dropping through, continuing to attempt execution...)" ; 
                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;  
                                } ;
                            }elseif($certLifeDays -lt $certWarnDays){
                                $smsg = "`n`n*** CERTIFICATE $($Credential.UserName) ($($certUname) EXPIRES IN $($certLifeDays) DAYS! ***" ; 
                                $smsg += "`n$(($oCert | fl $prpCertgci |out-string).trim())`n`n" ; 
                                $smsg += $hsRollCert ; 
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                            } else{
                                $smsg = "(Auth Certificate $($Credential.UserName) ($($certUname) remaining lifespan:$($certLifeDays) days)" ; 
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            
                            }; 
                            if($certLifeDays -lt $certWarnDays){
                                
                                

                            } ; # $certLifeDays -lt $certWarnDays

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

                    if($useCCMSConn){
                        $smsg = "connect-IPPSSession w`n$(($pltCEO|out-string).trim())" ; 
                    } else { 
                        $smsg = "Connect-ExchangeOnline w`n$(($pltCEO|out-string).trim())" ; 
                    } ;                 
                    if($silent){}elseif ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                    TRY {
                        if($useCCMSConn){
                            connect-IPPSSession @pltCEO ;
                        } else {
                            Connect-ExchangeOnline @pltCEO ;
                        } ; 
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
                    # -------- $UseConnEXO 
                } ; 
            } ; #  # if-E $bExistingEXOGood

        } else { 
            $smsg = "(-not:`$isBased: running most basic Get-ConnectionInformation , Connect-ExchangeOnline connectivity)" ; 
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        } ; # $isBased
    } ; # PROC-E
    END {
        
        <# 1:10 PM 3/1/2024 there are no more pss's in eom, rem it
        $smsg = "Existing PSSessions:`n$((get-pssession|out-string).trim())" ; 
        if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        #>
        if($isBased){
            if( ($useCCMSConn -AND ($bExistingCCMSGood -eq $false)) -OR (-not($useCCMSConn) -AND $bExistingEXOGood -eq $false) ){
                $tIn = '4/18/2024;D:\scripts\test-EXOv2Connection_func.ps1' ;
                $tdt,$tsrc =  $tIn.split(';') ;
                $tdt=[datetime]$tdt ;
                if($psise -and (get-date ).date -eq $tdt){
                    $gcm = gcm (split-path $tsrc -leaf).replace('_func.ps1','') ;
                    if( $gcm -AND $gcm.source -ne ''){
                        gci function:$((split-path $tsrc -leaf).replace('_func.ps1','')) -ea 0| remove-item -force ;
                        ipmo -fo -verb $tsrc;
                    } else {write-host "(non-Mod vers loaded)"} ;
                } ; 
                if($CertTag -ne $null){
                    $smsg = "(specifying detected `$CertTag:$($CertTag))" ; 
                    if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                    $oRet = test-EXOv2Connection -Credential $credential -CertTag $certtag -Prefix $Prefix -verbose:$($verbose) ; 
                } else { 
                    $oRet = test-EXOv2Connection -Credential $credential -Prefix $Prefix -verbose:$($verbose) ; 
                } ; 

                if($useCCMSConn){
                    $bExistingCCMSGood = $oRet.Valid ;
                } else { 
                    $bExistingEXOGood = $oRet.Valid ;
                } ; 

                if($oRet.Valid){
                    $pssEXOv2 = $oRet.PsSession ;
                    $IsNoWinRM = $oRet.IsNoWinRM ; 
                    <#$oRet
                    PSSession     :
                    IsNoWinRM     : True
                    Valid         : True
                    Prefix        : xo
                    isEXO         : True
                    isCCMS        : False
                    ConnectionUri : https://outlook.office365.com
                    #>        
                    switch -regex ($oRet.ConnectionUri){
                        $rgxConnectionUriEXO {
                            if ($oRet.Valid -AND $oRet.isEXO){
                                $smsg = "connected to EXO" ;
                                $bExistingEXOGood = $isEXOValid = $true ;
                            } ;
                        }
                        $rgxConnectionUriCCMS {
                            if ($oRet.Valid -AND $oRet.isCCMS){
                                $smsg = "connected to CCMS" ;
                                $bExistingCCMSGood = $isCCMSValid = $true ;
                            } ;
                        }
                        default {
                            $bExistingEXOGood = $isEXOValid = $bExistingCCMSGood = $isCCMSValid = $FALSE ;
                            $smsg = "unreconized test-EXOv2Connection returned:`$oRet.ConnectionUri:$($oRet.ConnectionUri)!" 
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                        }
                    } ; 
      
                    $smsg += ":Validated Connected to Tenant aligned with specified Credential: `$IsNoWinRM:$($IsNoWinRM)" ;
                    if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                } else {
                    $smsg = "NO VALID EXOV2/3 SESSION FOUND! (DISCONNECTING...)"
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } 
                    else{ write-host -ForegroundColor yellow "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                    # capture outlier: shows a session wo the test cmdlet, force reset
                    $pltDXO=[ordered]@{
                        confirm = $false ;
                        erroraction = 'STOP' ;
                        whatif = $($whatif) ;
                    } ;
                    if($Prefix){
                        $pltDXO.add('ModulePrefix',$Prefix) ;
                    }
                    $smsg = "Disconnect-ExchangeOnline w`n$(($pltDXO|out-string).trim())" ;
                    if($silent){}elseif ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                    Disconnect-ExchangeOnline @pltDXO ;

                    #  DisConnect-CCMS ; # CCMS
                    #DisConnect-EXO ;
                    $bExistingEXOGood = $false ;
                } ;       
            } else {
              <# 9:08 AM 4/2/2024 unneded post v205p5 - disabled the cod that sets vari, above
              if($bPreExoPPss){
                  $smsg = "(EMO bug-workaround: reconnecting prior ExOP PssSession,"
                  $smsg += "`nreconnect-Ex2010 -Credential $($pltRX10.Credential.username) -verbose:$($VerbosePreference -eq "Continue"))" ; 
                  if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                  else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                  reconnect-Ex2010 -Credential $pltRX10.Credential -verbose:$($VerbosePreference -eq "Continue") ; 
              } ;
              #>
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
                if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            
                if($IsNoWinRM -AND ((get-module $EXOv2GmoNoWinRMFilter) -AND (get-module $EOMModName))){
                    $smsg = "(native non-WinRM/Non-PSSession-based EXO connection detected.)" ; 
                    if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                } ; 
            } ; 

            # 10:37 AM 4/18/2023: Rem this: Been seldom capturing returns: that's bound to contaiminate pipeline! May have planned to grab and compare, but never really implemented
            #$bExistingEXOGood | write-output ;
            # splice in console color scheming

        } else { 
            $smsg = "(-not:`$isBased: running most basic Get-ConnectionInformation , Connect-ExchangeOnline connectivity)" ; 
            #if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
            #else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        } ; 
    }  # END-E
} ; 
#*------^ Connect-EXO.ps1 ^------