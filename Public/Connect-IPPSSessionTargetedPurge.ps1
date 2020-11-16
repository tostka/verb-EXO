#*------v Connect-IPPSSessionTargetedPurge.ps1 v------
function Connect-IPPSSessionTargetedPurge{
    <#
    .SYNOPSIS
    Connect-IPPSSessionTargetedPurge.ps1 - localized verb-EXO vers of non-'$global:' funct from ExchangeOnlineManagement. The globals export fine, these don't and appear to need to be loaded manually
    .NOTES
    Version     : 1.0.0
    Author      : Todd Kadrie
    Website     :	http://www.toddomation.com
    Twitter     :	@tostka / http://twitter.com/tostka
    CreatedDate : 20201109-0833AM
    FileName    : Connect-IPPSSessionTargetedPurge.ps1
    License     : [none specified]
    Copyright   : [none specified]
    Github      : https://github.com/tostka/verb-exo
    Tags        : Powershell
    AddedCredit : Microsoft (edited version of published commands in the module)
    AddedWebsite:	https://docs.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2
    REVISIONS
    * 8:34 AM 11/9/2020 init
    .DESCRIPTION
    Connect-IPPSSessionTargetedPurge.ps1 - Extract organization name from UserPrincipalName ; localized verb-EXO vers of non-'$global:' helper funct from ExchangeOnlineManagement. The globals export fine, these don't and appear to need to be loaded manually
    .PARAMETER ConnectionUri = 'https://ps.compliance.protection.outlook.com/PowerShell-LiveId',
    Connection Uri for the Remote PowerShell endpoint
    .PARAMETER AzureADAuthorizationEndpointUri = 'https://login.windows.net/common',
    Azure AD Authorization endpoint Uri that can issue the OAuth2 access tokens
    .PARAMETER DelegatedOrganization = '',
    Delegated Organization Name
    .PARAMETER PSSessionOption = $null,
    PowerShell session options to be used when opening the Remote PowerShell session
    .PARAMETER BypassMailboxAnchoring = $false
    Switch to bypass use of mailbox anchoring hint.
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    None. Returns no objects or output.
    .EXAMPLE
    Connect-IPPSSessionTargetedPurge
    Stock call
    .LINK
    https://github.com/tostka/verb-EXO
    .LINK
    https://docs.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2
    #>
    [CmdletBinding()]
    param(
        # Connection Uri for the Remote PowerShell endpoint
        [string] $ConnectionUri = 'https://ps.compliance.protection.outlook.com/PowerShell-LiveId',

        # Azure AD Authorization endpoint Uri that can issue the OAuth2 access tokens
        [string] $AzureADAuthorizationEndpointUri = 'https://login.windows.net/common',

        # Delegated Organization Name
        [string] $DelegatedOrganization = '',

        # PowerShell session options to be used when opening the Remote PowerShell session
        [System.Management.Automation.Remoting.PSSessionOption] $PSSessionOption = $null,

        # Switch to bypass use of mailbox anchoring hint.
        [switch] $BypassMailboxAnchoring = $false
    )
    DynamicParam
    {
        if (($isCloudShell = IsCloudShellEnvironment) -eq $false)
        {
            $attributes = New-Object System.Management.Automation.ParameterAttribute
            $attributes.Mandatory = $false

            $attributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
            $attributeCollection.Add($attributes)

            # User Principal Name or email address of the user
            $UserPrincipalName = New-Object System.Management.Automation.RuntimeDefinedParameter('UserPrincipalName', [string], $attributeCollection)
            $UserPrincipalName.Value = ''

            # User Credential to Logon
            $Credential = New-Object System.Management.Automation.RuntimeDefinedParameter('Credential', [System.Management.Automation.PSCredential], $attributeCollection)
            $Credential.Value = $null

            $paramDictionary = New-object System.Management.Automation.RuntimeDefinedParameterDictionary
            $paramDictionary.Add('UserPrincipalName', $UserPrincipalName)
            $paramDictionary.Add('Credential', $Credential)
            return $paramDictionary
        }
        else
        {
            $attributes = New-Object System.Management.Automation.ParameterAttribute
            $attributes.Mandatory = $false

            $attributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
            $attributeCollection.Add($attributes)

            # Switch to MSI auth 
            $Device = New-Object System.Management.Automation.RuntimeDefinedParameter('Device', [switch], $attributeCollection)
            $Device.Value = $false

            $paramDictionary = New-object System.Management.Automation.RuntimeDefinedParameterDictionary
            $paramDictionary.Add('Device', $Device)
            return $paramDictionary
        }
    }
        BEGIN {
        # TSK:add a BEGIN block & stick THE ExchangOnlineManagement.psm1 'above-the mods' variable/load specs in here, with tests added
        # Import the REST module so that the EXO* cmdlets are present before Connect-ExchangeOnline in the powershell instance.
        
        if(-not($ExchangeOnlineMgmtPath)){
            $EOMgmtModulePath = split-path (get-module ExchangeOnlineManagement -list).Path ; 
        } ; 
        if(!$RestModule){$RestModule = "Microsoft.Exchange.Management.RestApiClient.dll"} ;
        # $PSScriptRoot will be the verb-EXO path, not the EXOMgmt module have to dyn locate it
        if(!$RestModulePath){
            #$RestModulePath = [System.IO.Path]::Combine($PSScriptRoot, $RestModule)
            $RestModulePath = [System.IO.Path]::Combine($EOMgmtModulePath, $RestModule)
        } ;
        # paths to proper Module path: Name lists as: Microsoft.Exchange.Management.RestApiClient
        if(-not(get-module Microsoft.Exchange.Management.RestApiClient)){
            Import-Module $RestModulePath
        } ;

        if(!$ExoPowershellModule){$ExoPowershellModule = "Microsoft.Exchange.Management.ExoPowershellGalleryModule.dll"} ;
        if(!$ExoPowershellModulePath){
            $ExoPowershellModulePath = [System.IO.Path]::Combine($EOMgmtModulePath, $ExoPowershellModule) ;
        } ;
        # full path: C:\Users\kadritss\Documents\WindowsPowerShell\Modules\ExchangeOnlineManagement\1.0.1\Microsoft.Exchange.Management.ExoPowershellGalleryModule.dll
        # Name: Microsoft.Exchange.Management.ExoPowershellGalleryModule
        if(-not(get-module Microsoft.Exchange.Management.ExoPowershellGalleryModule)){
            Import-Module $ExoPowershellModulePath
        } ; 
    } 
    process 
    {
        [string]$newUri = $null;

        if (![string]::IsNullOrWhiteSpace($DelegatedOrganization))
        {
            [UriBuilder] $uriBuilder = New-Object -TypeName UriBuilder -ArgumentList $ConnectionUri;
            [string] $queryToAppend = "DelegatedOrg={0}" -f $DelegatedOrganization;
            if ($uriBuilder.Query -ne $null -and $uriBuilder.Query.Length -gt 0)
            {
                [string] $existingQuery = $uriBuilder.Query.Substring(1);
                $uriBuilder.Query = $existingQuery + "&" + $queryToAppend;
            }
            else
            {
                $uriBuilder.Query = $queryToAppend;
            }

            $newUri = $uriBuilder.ToString();
        }
        else
        {
           $newUri = $ConnectionUri;
        }

        if ($isCloudShell -eq $false)
        {
            Connect-ExchangeOnline -ConnectionUri $newUri -AzureADAuthorizationEndpointUri $AzureADAuthorizationEndpointUri -UserPrincipalName $UserPrincipalName.Value -PSSessionOption $PSSessionOption -Credential $Credential.Value -BypassMailboxAnchoring:$BypassMailboxAnchoring -ShowBanner:$false
        }
        else
        {
            Connect-ExchangeOnline -ConnectionUri $newUri -AzureADAuthorizationEndpointUri $AzureADAuthorizationEndpointUri -PSSessionOption $PSSessionOption -BypassMailboxAnchoring:$BypassMailboxAnchoring -Device:$Device.Value -ShowBanner:$false
        }
    }
}

#*------^ Connect-IPPSSessionTargetedPurge.ps1 ^------