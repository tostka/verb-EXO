#*------v Connect-EXOPSSession.ps1 v------
function Connect-EXOPSSession {
    <#
    .SYNOPSIS
   Connect-EXOPSSession.ps1 - Stripped to basics version of the Exchangeonlinemanagement module:connect-ExchangeOnline(), uses RemoveExistingEXOPSSession (vs RemoveExistingPSSession) to leave CCMS sessions intact, and permit run of concurrent EXO & CCMS sessions
    .NOTES
    Version     : 1.0.0
    Author      : Todd Kadrie
    Website     :	http://www.toddomation.com
    Twitter     :	@tostka / http://twitter.com/tostka
    CreatedDate : 20201109-0833AM
    FileName    :Connect-EXOPSSession.ps1
    License     : [none specified]
    Copyright   : [none specified]
    Github      : https://github.com/tostka/verb-XXX
    Tags        : Powershell
    AddedCredit : Microsoft (edited version of published commands in the module)
    AddedWebsite: https://www.powershellgallery.com/packages/CreateExoPsSession/0.1/Content/CreateExoPsSession.psm1
    REVISIONS
    * 12:01 PM 7/15/2024 long obso pssession target func, delete
    * 3:36 PM 11/9/2020 init debugged to basic function
    .DESCRIPTION
   Connect-EXOPSSession.ps1 - *another* take on a stripped to basics version of the Exchangeonlinemanagement module:connect-ExchangeOnline(), uses RemoveExistingEXOPSSession (vs RemoveExistingPSSession) to leave CCMS sessions intact, and permit run of concurrent EXO & CCMS sessions
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
    https://www.powershellgallery.com/packages/CreateExoPsSession/0.1/Content/CreateExoPsSession.psm1
    #>

    param(
        # Connection Uri for the Remote PowerShell endpoint
        [string] $ConnectionUri = 'https://outlook.office365.com/PowerShell-LiveId',
        # Azure AD Authorization endpoint Uri that can issue the OAuth2 access tokens
        [string] $AzureADAuthorizationEndpointUri = 'https://login.windows.net/common',
        # User Principal Name or email address of the user
        [string] $UserPrincipalName = '',
        # PowerShell session options to be used when opening the Remote PowerShell session
        [System.Management.Automation.Remoting.PSSessionOption] $PSSessionOption = $null,
        # User Credential to Logon
        [System.Management.Automation.PSCredential] $Credential = $null
    )

    # Validate parameters
    if (-not (Test-Uri $ConnectionUri)){throw "Invalid ConnectionUri parameter '$ConnectionUri'"}
    if (-not (Test-Uri $AzureADAuthorizationEndpointUri)){throw "Invalid AzureADAuthorizationEndpointUri parameter '$AzureADAuthorizationEndpointUri'"}

    try{
        # Cleanup old ps sessions - TSK this is gonna kill EVERYTHING! not good
        Get-PSSession | Remove-PSSession

        # TSK, don't use psscript, pull it dyn from profile
        if(!$PSExoPowershellModuleRoot){$PSExoPowershellModuleRoot = (Get-ChildItem -Path $($env:LOCALAPPDATA+"\Apps\2.0\") -Filter CreateExoPSSession.ps1 -Recurse -ErrorAction SilentlyContinue -Force | Select -Last 1).DirectoryName } ; 
        if(!$ExoPowershellModule){$ExoPowershellModule = "Microsoft.Exchange.Management.ExoPowershellModule.dll"} ; 
        if(!$ExoPowershellModulePath){$ExoPowershellModulePath = [System.IO.Path]::Combine($PSExoPowershellModuleRoot, $ExoPowershellModule)} ; 

        $global:_EXO_ConnectionUri = $ConnectionUri;
        $global:_EXO_AzureADAuthorizationEndpointUri = $AzureADAuthorizationEndpointUri;
        $global:_EXO_UserPrincipalName = $UserPrincipalName;
        $global:_EXO_PSSessionOption = $PSSessionOption;
        $global:_EXO_Credential = $Credential;

        Import-Module $ExoPowershellModulePath -verbose:$false;
        $PSSession = New-ExoPSSession -UserPrincipalName $UserPrincipalName -ConnectionUri $ConnectionUri -AzureADAuthorizationEndpointUri $AzureADAuthorizationEndpointUri -PSSessionOption $PSSessionOption -Credential $Credential
    
        if ($PSSession -ne $null)
        {
            Import-PSSession $PSSession -AllowClobber -Prefix $Prefix ;
            UpdateImplicitRemotingHandler
        }
    }catch{
        throw $_
    }
}

#*------^ Connect-EXOPSSession.ps1 ^------