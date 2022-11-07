#*------v connect-EXOv2RAW.ps1 v------
function connect-EXOv2RAW {
    <#
    .SYNOPSIS
    Connect-ExchangeOnlineTargetedPurge.ps1 - Stripped to basics version of the Exchangeonlinemanagement module:connect-ExchangeOnline(), uses RemoveExistingEXOPSSession (vs RemoveExistingPSSession) to leave CCMS sessions intact, and permit run of concurrent EXO & CCMS sessions
    .NOTES
    Version     : 1.0.0
    Author      : Todd Kadrie
    Website     :	http://www.toddomation.com
    Twitter     :	@tostka / http://twitter.com/tostka
    CreatedDate : 20201109-0833AM
    FileName    : Connect-ExchangeOnlineTargetedPurge.ps1
    License     : [none specified]
    Copyright   : [none specified]
    Github      : https://github.com/tostka/verb-XXX
    Tags        : Powershell
    AddedCredit : Microsoft (edited version of published commands in the module)
    AddedWebsite:	https://docs.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2
    REVISIONS
    # 8:34 AM 3/31/2021 added verbose suppress to all import-mods
    * 3:36 PM 11/9/2020 init debugged to basic function
    .DESCRIPTION
    Connect-ExchangeOnlineTargetedPurge.ps1 - Stripped to basics version of the Exchangeonlinemanagement module:connect-ExchangeOnline(), uses RemoveExistingEXOPSSession (vs RemoveExistingPSSession) to leave CCMS sessions intact, and permit run of concurrent EXO & CCMS sessions
    .PARAMETER ConnectionUri
    Connection Uri for the Remote PowerShell endpoint
    .PARAMETER AzureADAuthorizationEndpointUri,
    Azure AD Authorization endpoint Uri that can issue the OAuth2 access tokens
    .PARAMETER ExchangeEnvironmentName = 'O365Default',
    Exchange Environment name
    .PARAMETER PSSessionOption
    PowerShell session options to be used when opening the Remote PowerShell session
    .PARAMETER BypassMailboxAnchoring
    Switch to bypass use of mailbox anchoring hint.
    .PARAMETER DelegatedOrganization
    Delegated Organization Name
    .PARAMETER Prefix
    Command Prefix
    .PARAMETER ShowBanner
    Show Banner of Exchange cmdlets Mapping and recent updates
    .PARAMETER UserPrincipalName
    User Principal Name or email address of the user
    .PARAMETER Credential
    User Credential to Logon
    .PARAMETER EnableErrorReporting
    Switch to collect telemetry on command execution. - NOPE
    .PARAMETER TrackPerformance
    Switch to track perfomance
    .PARAMETER ShowProgress = $false
    Flag to enable or disable showing the number of objects written
    .PARAMETER UseMultithreading
    Switch to enable/disable Multi-threading in the EXO cmdlets
    .PARAMETER Pagesize
    Pagesize Param
    .PARAMETER ShowDebug
    Parameter to display Debugging messages [-ShowDebug switch]
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    None. Returns no objects or output.
    .EXAMPLE
    $PSSession = New-ExoPSSession -ExchangeEnvironmentName $ExchangeEnvironmentName -ConnectionUri $ConnectionUri -AzureADAuthorizationEndpointUri $AzureADAuthorizationEndpointUri -UserPrincipalName $UserPrincipalName.Value -PSSessionOption $PSSessionOption -Credential $Credential.Value -BypassMailboxAnchoring:$BypassMailboxAnchoring -DelegatedOrg $DelegatedOrganization
    .EXAMPLE
    connect-exov2Raw -credential $credO365TORSID -prefix xo
    .LINK
    https://github.com/tostka/verb-EXO
    .LINK
    https://docs.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2
    #>
    [CmdletBinding()]
    param(
        # stock params
        # Connection Uri for the Remote PowerShell endpoint
        [string] $ConnectionUri,
        # Azure AD Authorization endpoint Uri that can issue the OAuth2 access tokens
        [string] $AzureADAuthorizationEndpointUri,
        # Exchange Environment name
        [Microsoft.Exchange.Management.RestApiClient.ExchangeEnvironment] $ExchangeEnvironmentName = 'O365Default',
        # PowerShell session options to be used when opening the Remote PowerShell session
        [System.Management.Automation.Remoting.PSSessionOption] $PSSessionOption = $null,
        # Switch to bypass use of mailbox anchoring hint.
        [switch] $BypassMailboxAnchoring = $false,
        # Delegated Organization Name
        [string] $DelegatedOrganization,
        # Prefix
        [string] $Prefix,
        # Show Banner of Exchange cmdlets Mapping and recent updates
        [switch] $ShowBanner,

        # add back the dynamic paras as explicit paras:
        # User Principal Name or email address of the user
        [string]$UserPrincipalName,
        # User Credential to Logon
        [System.Management.Automation.PSCredential]$Credential,
        # Switch to collect telemetry on command execution. - NOPE
        #[switch]$EnableErrorReporting
        # Switch to track perfomance
        [switch]$TrackPerformance,
        # Flag to enable or disable showing the number of objects written
        [switch]$ShowProgress,
        # Switch to enable/disable Multi-threading in the EXO cmdlets
        [switch]$UseMultithreading = $true,
        # Pagesize Param
        [uint32]$PageSize = 1000
    )

    # intent is to strip down the ExchangeOnlineManagement module's Connect-ExchangeOnline and distill it into the lowest level non-wrapped commands available

    # drop all the cloudshell support variants
    # just straight path to new-EXOPsSession

    BEGIN {
        # TSK:add a BEGIN block & stick THE ExchangOnlineManagement.psm1 'above-the mods' variable/load specs in here, with tests added
        # Import the REST module so that the EXO* cmdlets are present before Connect-ExchangeOnline in the powershell instance.

        if (-not($ExchangeOnlineMgmtPath)) {
            $EOMgmtModulePath = split-path (get-module ExchangeOnlineManagement -list).Path ;
        } ;
        if (!$RestModule) { $RestModule = "Microsoft.Exchange.Management.RestApiClient.dll" } ;
        # stock uses $PSScriptRoot, which will be the verb-EXO path, not the EXOMgmt module have to dyn locate it
        if (!$RestModulePath) {
            $RestModulePath = [System.IO.Path]::Combine($EOMgmtModulePath, $RestModule)
        } ;
        # paths to proper Module path: Name lists as: Microsoft.Exchange.Management.RestApiClient
        if (-not(get-module Microsoft.Exchange.Management.RestApiClient)) {
            Import-Module $RestModulePath -verbose:$false ;
        } ;

        if (!$ExoPowershellModule) { $ExoPowershellModule = "Microsoft.Exchange.Management.ExoPowershellGalleryModule.dll" } ;
        if (!$ExoPowershellModulePath) {
            $ExoPowershellModulePath = [System.IO.Path]::Combine($EOMgmtModulePath, $ExoPowershellModule) ;
        } ;
        # full path: C:\Users\SIDs\Documents\WindowsPowerShell\Modules\ExchangeOnlineManagement\1.0.1\Microsoft.Exchange.Management.ExoPowershellGalleryModule.dll
        # Name: Microsoft.Exchange.Management.ExoPowershellGalleryModule
        if (-not(get-module Microsoft.Exchange.Management.ExoPowershellGalleryModule)) {
            Import-Module $ExoPowershellModulePath -Verbose:$false ;
        } ;
    }
    PROCESS {
        # Validate parameters
        if (($ConnectionUri) -and (-not (Test-Uri $ConnectionUri))) {
            throw "Invalid ConnectionUri parameter '$ConnectionUri'"
        }
        if (($AzureADAuthorizationEndpointUri) -and (-not (Test-Uri $AzureADAuthorizationEndpointUri))) {
            throw "Invalid AzureADAuthorizationEndpointUri parameter '$AzureADAuthorizationEndpointUri'"
        }
        if (($Prefix) -and ($Prefix -eq 'EXO')) {
            throw "Prefix 'EXO' is a reserved Prefix, please use a different prefix."
        }
        if ($ShowBanner -eq $true) {
            Print-Details;
        }
        if (($ConnectionUri) -and (-not($AzureADAuthorizationEndpointUri))) {
            Write-Host -ForegroundColor Green "Using ConnectionUri:'$ConnectionUri', in the environment:'$ExchangeEnvironmentName'."
        }
        if (($AzureADAuthorizationEndpointUri) -and (-not($ConnectionUri))) {
            Write-Host -ForegroundColor Green "Using AzureADAuthorizationEndpointUri:'$AzureADAuthorizationEndpointUri', in the environment:'$ExchangeEnvironmentName'."
        }
        # Keep track of error count at beginning.
        $errorCountAtStart = $global:Error.Count;
        $global:_EXO_TelemetryFilePath = $null;

        try {
            # Cleanup old exchange online PSSessions
            #RemoveExistingPSSession
            RemoveExistingEXOPSSession
            $ExoPowershellModule = "Microsoft.Exchange.Management.ExoPowershellGalleryModule.dll";
            $ModulePath = [System.IO.Path]::Combine($EOMgmtModulePath, $ExoPowershellModule);
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
            # import the ExoPowershellModule .dll
            Import-Module $ModulePath -verbose:$false;
            $global:_EXO_ModulePath = $ModulePath;
            # $PSSession = New-ExoPSSession -ExchangeEnvironmentName $ExchangeEnvironmentName -ConnectionUri $ConnectionUri -AzureADAuthorizationEndpointUri $AzureADAuthorizationEndpointUri -UserPrincipalName $UserPrincipalName.Value -PSSessionOption $PSSessionOption -Credential $Credential.Value -BypassMailboxAnchoring:$BypassMailboxAnchoring -DelegatedOrg $DelegatedOrganization

            $pltNEXOS = @{
                ExchangeEnvironmentName         = $ExchangeEnvironmentName ;
                ConnectionUri                   = $ConnectionUri ;
                AzureADAuthorizationEndpointUri = $AzureADAuthorizationEndpointUri ;
                UserPrincipalName               = $UserPrincipalName ;
                PSSessionOption                 = $PSSessionOption ;
                Credential                      = $Credential ;
                BypassMailboxAnchoring          = $($BypassMailboxAnchoring) ;
                DelegatedOrg                    = $DelegatedOrganization ;
            }
            write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):New-ExoPSSession w`n$(($pltNEXOS|out-string).trim())" ;
            $PSSession = New-ExoPSSession @pltNEXOS ;

            if ($PSSession -ne $null ) {
                $PSSessionModuleInfo = Import-PSSession $PSSession -AllowClobber -DisableNameChecking
                $pltIMod=@{Global=$true;DisableNameChecking=$true ; verbose=$false} ; # force verbose off, suppress spam in console
                if($Prefix){
                    $pltIMod.add('Prefix',$CommandPrefix) ;
                } ;
                # Import the above module globally. This is needed as with using psm1 files,
                # any module which is dynamically loaded in the nested module does not reflect globally.
                Import-Module $PSSessionModuleInfo.Path @pltIMod ;
                # haven't checked into what this does
                UpdateImplicitRemotingHandler ;

                # Import the REST module .dll
                $RestPowershellModule = "Microsoft.Exchange.Management.RestApiClient.dll";
                $RestModulePath = [System.IO.Path]::Combine($EOMgmtModulePath, $RestPowershellModule);
                Import-Module $RestModulePath -Cmdlet Set-ExoAppSettings -verbose:$false;

                # Set the AppSettings disabling the logging
                Set-ExoAppSettings -ShowProgress $ShowProgress.Value -PageSize $PageSize.Value -UseMultithreading $UseMultithreading.Value -TrackPerformance $TrackPerformance.Value -ExchangeEnvironmentName $ExchangeEnvironmentName -ConnectionUri $ConnectionUri -AzureADAuthorizationEndpointUri $AzureADAuthorizationEndpointUri -EnableErrorReporting $false ;

            }

        } CATCH {
            Write-Warning "$(get-date -format 'HH:mm:ss'): Failed processing $($_.Exception.ItemName). `nError Message: $($_.Exception.Message)`nError Details: $($_)" ;
            Exit #Opts: STOP(debug)|EXIT(close)|CONTINUE(move on in loop cycle)|BREAK(exit loop iteration)|THROW $_/'CustomMsg'(end script with Err output)
        } ;

    }

}

#*------^ connect-EXOv2RAW.ps1 ^------
