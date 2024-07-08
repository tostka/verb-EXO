# Connect-CCMS

#*------v Connect-CCMS.ps1 v------
Function Connect-CCMS {
    <#
    .SYNOPSIS
    Connect-CCMS - Establish Connection to Sec & Compliance V3 Modern Auth (https://ps.compliance.protection.outlook.com/powershell-liveid/)
    .NOTES
    Version     : 1.0.0
    Author      : Todd Kadrie
    Website     :	http://www.toddomation.com
    Twitter     :	@tostka / http://twitter.com/tostka
    CreatedDate : 2020-05-27
    FileName    : Connect-CCMS.ps1
    License     : MIT License
    Copyright   : (c) 2020 Todd Kadrie
    Github      : https://github.com/tostka/verb-CCMS
    Tags        : Powershell,SecurityAndCompliance,ExchangeOnline,Exchange,RemotePowershell,Connection,MFA
    AddedCredit : REFERENCE
    AddedWebsite:	URL
    AddedTwitter:	URL    
    REVISIONS   :
    # 4:47 PM 7/8/2024 this is obsoleted; shifted all (re|dis)connect-CCMS functions into connect-exo & reconnect-exo: CCMS Sec & Compl connection mgmt is triggered via the -Prefix cc parameter (any other param is assumed to be native EXO; but -Prefix cc will always generate a connection to Sec & Compliance); 
    * 1:05 PM 4/1/2024 repurp updated functional connect-exo to ccms (again)
    * 1:48 PM 3/1/2024  with v340 support for proper/native S&C conn, I can finally remove the legacy EOM connection bits from this (*substantial* simplification):
        - removed raft of pre EOMv3xx code, basic auth is fully blocked now, independantly, test-EXOv2Connection() got some updates (TenOrg tweak, likewise removed code < EOM3xx support)
    * 2:44 PM 3/2/2021 added console TenOrg color support
    * 7:13 AM 7/22/2020 replaced codeblock w get-TenantTag()
    * 12:18 PM 5/27/2020 updated cbh, moved alias:cccms win func
    * 4:17 PM 5/14/2020 fixed fundemental typos, in port over from verb-exo, mfa is just sketched in... we don't have it enabled, so it needs live debugging to update
    * 10:55 AM 12/6/2019 Connect-CCMS: added suffix to TitleBar tag for non-TOR tenants, also config'd a central tab vari
    * 5:14 PM 11/27/2019 repl $MFA code with get-TenantMFARequirement
    * 1:07 PM 11/25/2019 added *tol/*tor/*cmw alias variants for connect & reconnect
    # 1:31 PM 7/9/2018 added suffix hint: if($CommandPrefix){ '(Connected to CCMS: Cmds prefixed [verb]-cc[Noun])' ; } ;
    # 12:25 PM 6/20/2018 port from cxo:     Primary diff from EXO connect is the "-ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid/" all else is the same, repurpose connect-EXO to this
    .DESCRIPTION
    Connect-CCMS - Establish Connection to Sec & Compliance V3 Modern Auth (https://ps.compliance.protection.outlook.com/powershell-liveid/)

    revised 2/27/24: [Connect to Security & Compliance PowerShell | Microsoft Learn](https://learn.microsoft.com/en-us/powershell/exchange/connect-to-scc-powershell?view=exchange-ps)

    - USES eom: 
    Import-Module ExchangeOnlineManagement ;
    Connect-IPPSSession -UserPrincipalName <UPN> [-ConnectionUri <URL>] [-AzureADAuthorizationEndpointUri <URL>] [-DelegatedOrganization <String>] [-PSSessionOption $ProxyOptions]
    
    ## PARAMS: 
        ENVIRO: Microsoft 365 or Microsoft 365 GCC:
            -ConnectionUri: None. 
                The required value https://ps.compliance.protection.outlook.com/powershell-liveid/ is also the default value, so you don't need to use the ConnectionUri parameter in Microsoft 365 or Microsoft 365 GCC environments.
            -AzureADAuthorizationEndpointUri: None. 
                The required value https://login.microsoftonline.com/common is also the default value, so you don't need to use the AzureADAuthorizationEndpointUri parameter in Microsoft 365 or Microsoft 365 GCC environments.

    
    [Connect-IPPSSession (ExchangePowerShell) | Microsoft Learn](https://learn.microsoft.com/en-us/powershell/module/exchange/connect-ippssession?view=exchange-ps)
    0.
        -AppId
            The AppId parameter specifies the application ID of the service principal that's used in certificate based authentication (CBA). A valid value is the GUID of the application ID (service principal). For example, 36ee4c6c-0812-40a2-b820-b22ebd02bce3.
        -AzureADAuthorizationEndpointUri
            The AzureADAuthorizationEndpointUri parameter specifies the Microsoft Entra Authorization endpoint that can issue OAuth2 access tokens. The following PowerShell environments and related values are supported:
                Security & Compliance PowerShell in Microsoft 365 or Microsoft 365 GCC: Don't use this parameter. The required value is https://login.microsoftonline.com/common, but that's also the default value, so you don't need to use this parameter.
                ...
        -Certificate
            The Certificate parameter specifies the certificate that's used for certificate-based authentication (CBA). A valid value is the X509Certificate2 object value of the certificate.
            Don't use this parameter with the CertificateFilePath or CertificateThumbprint parameters.
        -CertificateFilePath
            The CertificateFilePath parameter specifies the certificate that's used for CBA. A valid value is the complete public path to the certificate file. Use the CertificatePassword parameter with this parameter.
            Don't use this parameter with the Certificate or CertificateThumbprint parameters.
        -CertificatePassword
            The CertificatePassword parameter specifies the password that's required to open the certificate file when you use the CertificateFilePath parameter to identify the certificate that's used for CBA.
            You can use the following methods as a value for this parameter:
                (ConvertTo-SecureString -String '<password>' -AsPlainText -Force).
                Before you run this command, store the password as a variable (for example, $password = Read-Host "Enter password" -AsSecureString), and then use the variable ($password) for the value.
                (Get-Credential).password to be prompted to enter the password securely when you run this command.
        -CertificateThumbprint
            The CertificateThumbprint parameter specifies the certificate that's used for CBA. A valid value is the thumbprint value of the certificate. For example, 83213AEAC56D61C97AEE5C1528F4AC5EBA7321C1.
            Don't use this parameter with the Certificate or CertificateFilePath parameters.
            Note: The CertificateThumbprint parameter is supported only in Microsoft Windows.    
        -CommandName
            The CommandName parameter specifies the comma separated list of commands to import into the session. Use this parameter for applications or scripts that use a specific set of cmdlets. Reducing the number of cmdlets in the session helps improve performance and reduces the memory footprint of the application or script.
        -ConnectionUri
        The ConnectionUri parameter specifies the connection endpoint for the PowerShell session. The following PowerShell environments and related values are supported:
            Security & Compliance PowerShell in Microsoft 365 or Microsoft 365 GCC: Don't use this parameter. The required value is https://ps.compliance.protection.outlook.com/powershell-liveid/, but that's also the default value, so you don't need to use this parameter
            ...
        -Credential
            The Credential parameter specifies the username and password that's used to connect to Exchange Online PowerShell. Typically, you use this parameter in scripts or when you need to provide different credentials that have the required permissions. Don't use this parameter for accounts that use multi-factor authentication (MFA).
            Before you run the Connect-IPPSSession command, store the username and password in a variable (for example, $UserCredential = Get-Credential). Then, use the variable name ($UserCredential) for this parameter.
            After the Connect-IPPSSession command is complete, the password key in the variable is emptied.
        -Organization
        The Organization parameter specifies the organization when you connect using CBA. You must use the primary .onmicrosoft.com domain of the organization for the value of this parameter.
        -Prefix
            The Prefix parameter specifies a text value to add to the names of Security & Compliance PowerShell cmdlets when you connect. For example, Get-ComplianceCase becomes Get-ContosoComplianceCase when you use the value Contoso for this parameter.
                The Prefix value can't contain spaces or special characters like underscores or asterisks.
                You can't use the Prefix value EXO. That value is reserved for the nine exclusive Get-EXO* cmdlets that are built into the module.
                The Prefix parameter affects only imported Security & Compliance cmdlet names. It doesn't affect the names of cmdlets that are built into the module (for example, Disconnect-IPPSSession).
        -PSSessionOption
            Note: This parameter doesn't work in REST API connections.
            The PSSessionOption parameter specifies the remote PowerShell session options to use in your connection to Security & Compliance PowerShell. This parameter works only if you also use the UseRPSSession switch in the same command.
            Store the output of the New-PSSessionOption command in a variable (for example, $PSOptions = New-PSSessionOption <Settings>), and use the variable name as the value for this parameter (for example, $PSOptions).
        -UserPrincipalName
            The UserPrincipalName parameter specifies the account that you want to use to connect (for example, navin@contoso.onmicrosoft.com). Using this parameter allows you to skip entering a username in the modern authentication credentials prompt (you're prompted to enter a password).
            If you use the UserPrincipalName parameter, you don't need to use the AzureADAuthorizationEndpointUri parameter for MFA or federated users in environments that normally require it (UserPrincipalName or AzureADAuthorizationEndpointUri is required; OK to use both).
        -UseRPSSession
            This parameter is available in version 3.2.0 or later of the module.
            Note: Remote PowerShell connections to Security & Compliance are deprecated. For more information, see Deprecation of Remote PowerShell in Security and Compliance PowerShell.
            The UseRPSSession switch allows you to connect to Security & Compliance PowerShell using traditional remote PowerShell access to all cmdlets. You don't need to specify a value with this switch.
            This switch requires that Basic authentication is enabled in WinRM on the local computer. For more information, see Turn on Basic authentication in WinRM.
            If you don't use this switch, Basic authentication in WinRM is not required.


    .PARAMETER Prefix
[verb]-PREFIX[command] PREFIX string for clearly marking cmdlets sourced in this connection [-Prefix tag]
.PARAMETER Credential
Credential to use for this connection [-credential [credential obj variable]
.PARAMETER UserPrincipalName
User Principal Name or email address of the user[-UserPrincipalName logon@domain.com]
.PARAMETER UserRole
Credential User Role spec (SID|CSID|UID|B2BI|CSVC|ESVC|LSVC|ESvcCBA|CSvcCBA|SIDCBA)[-UserRole @('SIDCBA','SID','CSVC')]
.PARAMETER TenOrg
TenantTag value, indicating Tenants to connect to[-TenOrg 'TOL']
.PARAMETER ExchangeEnvironmentName
Exchange Environment name [-ExchangeEnvironmentName 'O365Default']
.PARAMETER MinimumVersion
MinimumVersion required for ExchangeOnlineManagement module (defaults to '2.0.5')[-MinimumVersion '2.0.6']
.PARAMETER MinNoWinRMVersion
MinimumVersion required for Non-WinRM connections (of ExchangeOnlineManagement module (defaults to '3.0.0')[-MinimumVersion '2.0.6']
.PARAMETER UseMultithreading
Switch to enable/disable Multi-threading in the EXO cmdlets [-UseMultithreading]
.PARAMETER ShowProgress
Switch to enable or disable showing the number of objects written (defaults `$true)[-ShowProgress]
.PARAMETER PageSize
Pagesize Param[-PageSize 500]
.PARAMETER silent
Silent output (suppress status echos)[-silent]
.PARAMETER showDebug
Debugging Flag [-showDebug]
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    None. Returns no objects or output.
    .EXAMPLE
    PS>  Connect-CCMS -cred $credO365TORSID ;
    Connect using defaults, and leverage any pre-set $global:credo365TORSID variable
    .EXAMPLE
    PS>  Connect-CCMS -Prefix exolab -credential (Get-Credential -credential user@domain.com)  ;
    Connect an explicit credential, and use 'exolab' as the cmdlet prefix
    .EXAMPLE 
    Connect-CCMS2 -credential $credO365xxxCBA -verbose ; 
    Connect using a CBA credential variable (prestocked from profile automation). Script opens and recycles the cred cert specs emulating the native CBA connection below, but pulling source info from a stored dpapi-encrypted .xml credential file.
    .EXAMPLE
    Connect-CCMS -UserRole SIDCBA -TenOrg ABC -verbose  ; 
    Demo use of UserRole (specifying a CBA variant), AND TenOrg spec, to connect (autoresolves against preconfigured credentials in profile)
    .EXAMPLE
    PS>  $cred = get-credential -credential $o365_Torolab_SIDUpn ;
    PS>  Connect-CCMS -credential $cred ;
    Pass in a prefab credential object (useful for auto-shifting to MFA - the function will autoresolve MFA reqs based on the cred domain)
    .EXAMPLE
    PS> $pltRXOC = [ordered]@{
    PS>     Credential = $Credential ;
    PS>     verbose = $($VerbosePreference -eq "Continue")  ;
    PS>     Silent = $silent ; 
    PS> } ;
    PS> if ($script:useEXOv2 -OR $useEXOv2) { Connect-CCMS2 @pltRXOC }
    PS> else { Connect-CCMS @pltRXOC } ;    
    Splatted example leveraging prefab $pltRXOC splat, derived from local variables & $VerbosePreference value.
    .EXAMPLE
    PS>  $pltCXOCThmb=[ordered]@{
    PS>  	CertificateThumbPrint = $credO365TORSIDCBA.UserName ;
    PS>  	AppID = $credO365TORSIDCBA.GetNetworkCredential().Password ;
    PS>  	Organization = 'TENANTNAME.onmicrosoft.com' ;
    PS>  	Prefix = 'cc' ;
    PS>  	ShowBanner = $false ;
    PS>  };
    PS>  write-host "connect-IPPSSession w $(($pltCXOCThmb|out-string).trim())" ;
    PS>  connect-IPPSSession @pltCXOCThmb ;
    Example of native connect-IPPSSession syntax leveraging a CBA certificate stored locally, with AppID and CertificateThumbPrint pulled from a local global-scope credential object (with AppID stored as password & Thumprint as username)
    .LINK
    https://learn.microsoft.com/en-us/powershell/exchange/connect-to-scc-powershell?view=exchange-ps
    #>
    [CmdletBinding(DefaultParameterSetName='UPN')]
    [Alias('cccms')]
    PARAM(
        # try pulling all the ParameterSetName's - just need to get through it now. - no got through it with a defaultparametersetname (avoids 
        [Parameter(HelpMessage = "[verb]-PREFIX[command] PREFIX string for clearly marking cmdlets sourced in this connection [-Prefix tag]")]
            [string]$Prefix = 'cc',
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
            # cba's don't have perms to s&c: shift to mfa'd sid only
            [string[]]$UserRole = @('SID'),
            #@('SIDCBA','SID','CSVC'),
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
        write-warning "OBSOLETE! shifted all (re|dis)connect-CCMS functions into connect-exo & reconnect-exo: CCMS Sec & Compl connection mgmt is triggered via the -Prefix cc parameter (any other param is assumed to be native EXO; but -Prefix cc will always generate a connection to Sec & Compliance)!" ; 
        BREAK ; 
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
            $Prefix = 'cc' ; # 4:31 PM 7/29/2020 MS has RESERVED use of the 'exo' prefix [facepalm]
            $smsg = "(asserting Prefix:$($Prefix)" ;
            if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        } ;
        if (($Prefix) -and ($Prefix -eq 'EXO')) {
            throw "Prefix 'EXO' is a reserved Prefix, please use a different prefix."
        }

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

        # transplat fr rxo ---
        if(-not $Credential){
            if($UserRole){
                $smsg = "Using specified -UserRole:$( $UserRole -join ',' )" ;
                if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            #} else { $UserRole = @('SID','CSVC') } ;
            # S&C doesnt generally have cert or svc support
            } else { $UserRole = @('SID') } ;
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
        if((gcm Connect-CCMS).Parameters.keys -contains 'silent'){
            $pltCXO.add('Silent',$false) ;
        } ;

        $uRoleReturn = resolve-UserNameToUserRole -Credential $Credential ; 
        if($credential.username -match $rgxCertThumbprint){
            $certTag = $uRoleReturn.TenOrg ; 
        } ; 
        # ---

        $sTitleBarTag = @("CCMS") ;
        $sTitleBarTag += $TenOrg ;

        $MFA = get-TenantMFARequirement -Credential $Credential ;

        # 12:08 PM 8/2/2022 scrap the msal.net material: it's fundementally incompatible with EXO - sure you can pull and auth a token into the PS EXO clientid, but you can't spec a prefix on the returned cmdlets.
        # 4:45 PM 7/7/2022 workaround msal.ps bug: always ipmo it FIRST: "Get-msaltoken : The property 'Authority' cannot be found on this object. Verify that the property exists."

        # * 11:02 AM 4/4/2023 reduced the ipmo and vers chk block, removed the lengthy gmo -list; and any autoinstall. Assume EOM is installed, & break if it's not
        #region EOMREV ; #*------v EOMREV Check v------
        #$EOMmodname = 'ExchangeOnlineManagement' ;
        $pltIMod = @{Name = $EOMmodname ; ErrorAction = 'Stop' ; verbose=$false} ;
        # do a gmo first, faster than gmo -list
        if([version]$EOMMv = (Get-Module @pltIMod | sort version | select -last 1 ).version){}
        elseif([version]$EOMMv = (Get-Module -ListAvailable @pltIMod | sort version | select -last 1 ).version){} 
        else { 
            $smsg = "$($EOMmodname) PowerShell v$($MinNoWinRMVersion) module is required, do you want to install it?" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Prompt } 
            else{ write-host -foregroundcolor YELLOW "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            $bRet=Read-Host "Enter YYY to continue. Anything else will exit"  ; 
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
        $smsg = "Checking for WinRM support in this EOM rev..." 
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

        if(-not $UseConnEXO){
            $smsg = "NON-connect-IPPSSession() version of ExchangeOnlineManagement installed, update to vers:$($MinNoWinRMVersion) or higher!" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
            throw $smsg ; 
            break ; 

            # dropping all support/legacy code for EOMv1 (PSSession native-only connections)
            # code below is used *solely* for EOM v205 connections (uses cached creds, integrates connect-IPPSSession underlying commands)
            # EOM -lt 2.0.5preview6 .dll etc loads, from connect-IPPSSession: (should be installed with the above)
            # removed 12:23 PM 3/1/2024
        
        } else { 
            # $UseConnEXO => we're doing native connect-IPPSSession connectivity, no PSSession etc
            $smsg = "native connect-IPPSSession specified..." ; 
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

        -while a connect-IPPSSession (non-MFA, haven't verified) connect results in this PSS:
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
            $smsg = "NON-connect-IPPSSession() version of ExchangeOnlineManagement installed, update to vers:$($MinNoWinRMVersion) or higher!" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
            throw $smsg ; 
            break ; 
            # removed all legacy code: 12:25 PM 3/1/2024
            
        } else { 
            $smsg = "(native connect-IPPSSession specified...)" ; 
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
        #if ( Get-PSSession | where-object { $_.ConfigurationName -like "Microsoft.Exchange" -and $_.Name -like "ExchangeOnlineInternalSession*" -and $_.State -like "*Opened*" -AND ($_.Availability -eq 'Available') } ) {
        # update to *not* tamper with CCMS connects
        #if ( Get-PSSession | where-object { $_.ConfigurationName -like "Microsoft.Exchange" -and $_.Name -like "ExchangeOnlineInternalSession*" -and $_.State -like "*Opened*" -AND ($_.Availability -eq 'Available') -AND ($_.ComputerName -match $rgxExoPsHostName) } ) {
        # simpler - MS uses - very simple detect: 
        # $pssEXOv2 = Get-PSSession | Where-Object {$_.ConfigurationName -like "Microsoft.Exchange" -and $_.Name -like "ExchangeOnlineInternalSession*"} ;
       
        # use test-EXOConnection - cxo2 *only* drives compliant eXOv2 connections, not legacy basicAuth
        #$IsNoWinRM = $false ; 
        # 11:18 AM 4/25/2023 add support for passing calc'd CertTag "Cert FriendlyName Suffix to be used for validating credential alignment(Optional but required for CBA calls)[-CertTag `$certtag]")][string]$CertTag
        if($CertTag -ne $null){
            $smsg = "(specifying detected `$CertTag:$($CertTag))" ; 
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            # add prefix to steer into ccms
            $oRet = test-EXOv2Connection -Credential $credential -CertTag $certtag -verbose:$($verbose) -prefix 'cc' ; 
        } else { 
            # add prefix to steer into ccms
            $oRet = test-EXOv2Connection -Credential $credential -verbose:$($verbose)  -prefix 'cc'  ; 
        } ; 
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
    
        if ($bExistingEXOGood -eq $false) {
            # open a new EXOv2 session
            # removed all legacy code: 12:25 PM 3/1/2024
            if(-not $UseConnEXO){
                
                # removed all legacy code: 12:25 PM 3/1/2024

            } else { 
                # $UseConnEXO 
                <#
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
                    $smsg = "(adding specified connect-IPPSSession -Prefix:$($Prefix))" ; 
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
                        $certTag = $uRoleReturn.TenOrg
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

                $smsg = "connect-IPPSSession w`n$(($pltCEO|out-string).trim())" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                TRY {
                    #Disconnect-IPPSSession -confirm:$false ;
                    #connect-IPPSSession @pltCXO ;
                    connect-IPPSSession @pltCEO ;
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
    } ; # PROC-E
    END {
        
        <# 1:10 PM 3/1/2024 there are no more pss's in eom, rem it
        $smsg = "Existing PSSessions:`n$((get-pssession|out-string).trim())" ; 
        if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        #>

        if ($bExistingEXOGood -eq $false) {
            
            # defer into test-EXOv2Connection()
            # 11:18 AM 4/25/2023 add support for passing calc'd CertTag "Cert FriendlyName Suffix to be used for validating credential alignment(Optional but required for CBA calls)[-CertTag `$certtag]")][string]$CertTag
            if($CertTag -ne $null){
                $smsg = "(specifying detected `$CertTag:$($CertTag))" ; 
                if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                $oRet = test-EXOv2Connection -Credential $credential -CertTag $certtag -verbose:$($verbose) -Prefix:$($Prefix) ; 
            } else { 
                $oRet = test-EXOv2Connection -Credential $credential -verbose:$($verbose) -Prefix:$($Prefix) ; ; 
            } ; 

            $bExistingEXOGood = $oRet.Valid ;
            if($oRet.Valid){
	            $pssEXOv2 = $oRet.PsSession ;
                $IsNoWinRM = $oRet.IsNoWinRM ; 
	            $smsg = "(Validated EXOv2 Connected to Tenant aligned with specified Credential)`n`$IsNoWinRM:$($IsNoWinRM )" ;
	            if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
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
                $bExistingEXOGood = $false ;
            } ;       

        } else {
            <# shouldn't need the bug workaround post v205p5
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
            <# shouldn't need the bug workaround post v205p5
            if($bPreExoPPss -AND -not $pssEXOP){
                $smsg = "(EMO bug-workaround: reconnecting prior ExOP PssSession,"
                $smsg += "`nreconnect-Ex2010 -Credential $($pltRX10.Credential.username) -verbose:$($VerbosePreference -eq "Continue"))" ; 
                if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                reconnect-Ex2010 -Credential $pltRX10.Credential -verbose:$($VerbosePreference -eq "Continue") ; 
            } else { 
                $smsg = "(no bPreExoPPss, no Rx10 conn restore)" ; 
                if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            } ; 
            #>

            if($IsNoWinRM -AND ((get-module $EXOv2GmoNoWinRMFilter) -AND (get-module $EOMModName))){
                $smsg = "(native non-WinRM/Non-PSSession-based EXO connection detected." ; 
                if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            } ; 
        } ; 

        # 10:37 AM 4/18/2023: Rem this: Been seldom capturing returns: that's bound to contaiminate pipeline! May have planned to grab and compare, but never really implemented
        #$bExistingEXOGood | write-output ;
        # splice in console color scheming
    }  # END-E
} ; 
#*------^ Connect-CCMS.ps1 ^------