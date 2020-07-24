# verb-EXO.psm1


  <#
  .SYNOPSIS
  verb-EXO - Powershell Exchange Online generic functions module
  .NOTES
  Version     : 1.0.12.0
  Author      : Todd Kadrie
  Website     :	https://www.toddomation.com
  Twitter     :	@tostka
  CreatedDate : 3/3/2020
  FileName    : verb-EXO.psm1
  License     : MIT
  Copyright   : (c) 3/3/2020 Todd Kadrie
  Github      : https://github.com/tostka
  REVISIONS
  * 4:38 PM 3/16/2020 public cleanup
  * 8:45 AM 3/3/2020 1.0.0.0 public cleanup
  * 9:52 PM 1/16/2020 cleanup
  * 11:36 AM 12/30/2019 ran vsc alias-expan
  * 10:55 AM 12/6/2019 Connect-EXO:added suffix to TitleBar tag for non-TOR tenants, also config'd a central tab vari
  * 5:14 PM 11/27/2019 repl $MFA code with get-TenantMFARequirement
  * 1:07 PM 11/25/2019 added 3-letter alias variants for connect & reconnect
  # 9:57 AM 11/20/2019 added Credential param to reconnect, with passthru.
  # 1:26 PM 11/19/2019 added MFA detection fr infastrings .ps1 globals
  * 10:10 AM 6/20/2019 added local $rgxExoPsHostName, swapped dxo to use the vari, added showdebug to rxo & cxo, added $pltPSS wplat dump to the import-pssession cmd block
  * 1:02 PM 11/7/2018 added Disconnect-PssBroken
  * 2:55 PM 10/11/2018 connect-exo: added pre sleep skip on tryno 1
  # 9:24 PM 7/16/2018 broad cleanup & tightening
  # 9:04 PM 7/11/2018 synced to tsksid-incl-ServerApp.ps1
  .DESCRIPTION
  verb-EXO - Powershell Exchange Online generic functions module
  .LINK
  https://github.com/tostka/verb-EXO
  #>


$script:ModuleRoot = $PSScriptRoot ;
$script:ModuleVersion = (Import-PowerShellDataFile -Path (get-childitem $script:moduleroot\*.psd1).fullname).moduleversion ;

#*======v FUNCTIONS v======



#*------v Connect-EXO.ps1 v------
Function Connect-EXO {
    <#
    .SYNOPSIS
    Connect-EXO - Establish PSS to https://ps.outlook.com/powershell/
    .NOTES
    Version     : 1.0.0
    Author      : Todd Kadrie
    Website     :	http://www.toddomation.com
    Twitter     :	@tostka / http://twitter.com/tostka
    CreatedDate : 2020-03-03
    FileName    :
    License     :
    Copyright   :
    Github      : https://github.com/tostka
    Tags        : Powershell,ExchangeOnline,Exchange,RemotePowershell,Connection,MFA
    AddedCredit : ExactMike Perficient
    AddedWebsite:	https://social.technet.microsoft.com/Forums/msonline/en-US/f3292898-9b8c-482a-86f0-3caccc0bd3e5/exchange-powershell-monitoring-remote-sessions?forum=onlineservicesexchange
    AddedTwitter:
    AddedCredit2 : Jeremy Bradshaw
    AddedWebsite2:	https://github.com/JeremyTBradshaw
    AddedTwitter2:
    REVISIONS   :
    * 7:13 AM 7/22/2020 replaced codeblock w get-TenantTag()
    * 5:12 PM 7/21/2020 added ven supp
    * 11:50 AM 5/27/2020 added alias:cxo win func
    * 8:38 AM 4/17/2020 added a new test of $global:EOLSession, to detect initial cred fail (pw chg, outofdate creds, locked out)
    * 8:45 AM 3/3/2020 public cleanup, refactored connect-exo for Meta's
    * 9:52 PM 1/16/2020 cleanup
    * 10:55 AM 12/6/2019 Connect-EXO:added suffix to TitleBar tag for other tenants, also config'd a central tab vari
    * 9:17 AM 12/4/2019 CONSISTENTLY failing to load properly in lab, on lynms6200d - wont' get-module xxxx -listinstalled, even after load, so I rewrote an exemption diverting into the locally installed $env:userprofile\documents\WindowsPowerShell\Modules\exoMFAModule\ copy.
    * 5:14 PM 11/27/2019 repl $MFA code with get-TenantMFARequirement
    * 1:07 PM 11/25/2019 added tenant-specific alias variants for connect & reconnect
    # 1:26 PM 11/19/2019 added MFA detection fr infastrings .ps1 globals, lifted from Jeremy Bradshaw (https://github.com/JeremyTBradshaw)'s Connect-Exchange()
    # 10:35 AM 6/20/2019 added $pltPSS splat dump to the import-pssession cmd block; hard-typed the $Credential [System.Management.Automation.PSCredential]
    # 8:22 AM 11/20/2017 spliced in retry loop into reconnect-exo as well, to avoid using any state testing in scripts, localize it 1x here.
    # 1:49 PM 11/2/2017 coded around non-profile gaps from not having get-admincred() - added the prompt code in to fake it
    # 12:26 PM 9/11/2017 debugged retry - catch doesn't fire properly on new-Pssession, have to test the $error state, to detect auth fails (assuming the bad pw error# is specific). $error test is currently targeting specific error returned on a bad password. Added retry, for when connection won't hold and fails breaks - need to watch out that bad pw doesn't lock out the acct!
    # 12:50 PM 6/2/2017 expanded pshelp, added Examples, cleaned up deadwood
    * # 12:10 PM 3/15/2017 Connect-EXO typo, disable prefix auto spec, unless actually blanked. switch ProxyEnabled to non-Mandatory.
    .DESCRIPTION
    Connect-EXO - Establish PSS to https://ps.outlook.com/powershell/
    Original concept based on 'overlapping functions' concept by: ExactMike Perficient, Global Knowl... (Partner)
    Website:	https://social.technet.microsoft.com/Forums/msonline/en-US/f3292898-9b8c-482a-86f0-3caccc0bd3e5/exchange-powershell-monitoring-remote-sessions?forum=onlineservicesexchange
    Extended with Jeremy Bradshaw's on-the-fly EXO MFA module concept (https://github.com/JeremyTBradshaw)'s Connect-Exchange()
    .PARAMETER  ProxyEnabled
    Use Proxy-Aware SessionOption settings [-ProxyEnabled]
    .PARAMETER  CommandPrefix
    [verb]-PREFIX[command] PREFIX string for clearly marking cmdlets sourced in this connection [-CommandPrefix tag]
    .PARAMETER  Credential
    Credential to use for this connection [-credential 's-todd.kadrie@toro.com']
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    None. Returns no objects or output.
    .EXAMPLE
    connect-exo
    Connect using defaults, and leverage any pre-set $global:credo365TORSID variable
    .EXAMPLE
    connect-exo -CommandPrefix exo -credential (Get-Credential -credential s-todd.kadrie@torolab.com)  ;
    Connect an explicit credential, and use 'exolab' as the cmdlet prefix
    .EXAMPLE
    $cred = get-credential -credential $o365_Torolab_SIDUpn ;
    connect-exo -credential $cred ;
    Pass in a prefab credential object (useful for auto-shifting to MFA - the function will autoresolve MFA reqs based on the cred domain)
    .LINK
    https://social.technet.microsoft.com/Forums/msonline/en-US/f3292898-9b8c-482a-86f0-3caccc0bd3e5/exchange-powershell-monitoring-remote-sessions?forum=onlineservicesexchange
    .LINK
    https://github.com/JeremyTBradshaw
    #>
    [CmdletBinding()]
    [Alias('cxo')]
    Param(
        [Parameter(HelpMessage = "Use Proxy-Aware SessionOption settings [-ProxyEnabled]")][boolean]$ProxyEnabled = $False,
        [Parameter(HelpMessage = "[verb]-PREFIX[command] PREFIX string for clearly marking cmdlets sourced in this connection [-CommandPrefix tag]")][string]$CommandPrefix = 'exo',
        [Parameter(HelpMessage = "Credential to use for this connection [-credential [credential obj variable]")][System.Management.Automation.PSCredential]$Credential = $global:credo365TORSID,
        [Parameter(HelpMessage = "Debugging Flag [-showDebug]")]
        [switch] $showDebug
    ) ;

    $MFA = get-TenantMFARequirement -Credential $Credential ;

    # disable prefix spec, unless actually blanked (e.g. centrally spec'd in profile).
    if (!$CommandPrefix) {
        $CommandPrefix = 'exo' ;
        write-verbose -verbose:$true  "(asserting Prefix:$($CommandPrefix)" ;
    } ;

    $sTitleBarTag = "EXO" ;
    $TentantTag=get-TenantTag -Credential $Credential ; 
    if($TentantTag -ne 'TOR'){
        # explicitly leave this tenant (default) untagged
        $sTitleBarTag += $TentantTag ;
    } ; 

    $ImportPSSessionProps = @{
        AllowClobber        = $true ;
        DisableNameChecking = $true ;
        Prefix              = $CommandPrefix ;
        ErrorAction         = 'Stop' ;
    } ;

    if ($MFA) {
        try {
            $ExoPSModuleSearchProperties = @{
                Path        = "$($env:LOCALAPPDATA)\Apps\2.0\" ;
                Filter      = 'Microsoft.Exchange.Management.ExoPowerShellModule.dll' ;
                Recurse     = $true ;
                ErrorAction = 'Stop' ;
            } ;

            if ($showDebug) { write-host -foregroundcolor green "Get-ChildItem w`n$(($ExoPSModuleSearchProperties|out-string).trim())" } ;
            $ExoPSModule = Get-ChildItem @ExoPSModuleSearchProperties |
            Where-Object { $_.FullName -notmatch '_none_' } |
            Sort-Object LastWriteTime |
            Select-Object -Last 1 ;
            Import-Module $ExoPSModule.FullName -ErrorAction:Stop ;
            $ExoPSModuleManifest = $ExoPSModule.FullName -replace '\.dll', '.psd1' ;
            if (!(Get-Module $ExoPSModule.FullName -ListAvailable -ErrorAction 0 )) {
                write-verbose -verbose:$true  "Unable to`nGet-Module $($ExoPSModule.FullName) -ListAvailable`ndiverting to hardcoded exoMFAModule`nRequires that it be locally copied below`n$env:userprofile\documents\WindowsPowerShell\Modules\exoMFAModule\`n " ;
                # go to a hard load path $env:userprofile\documents\WindowsPowerShell\Modules\exoMFAModule\
                $ExoPSModuleSearchProperties = @{
                    Path        = "$($env:userprofile)\documents\WindowsPowerShell\Modules\exoMFAModule\" ;
                    Filter      = 'Microsoft.Exchange.Management.ExoPowerShellModule.dll' ;
                    Recurse     = $true ;
                    ErrorAction = 'Stop' ;
                } ;
                $ExoPSModule = Get-ChildItem @ExoPSModuleSearchProperties |
                Where-Object { $_.FullName -notmatch '_none_' } |
                Sort-Object LastWriteTime |
                Select-Object -Last 1 ;
                # roll an otf psd1+psm1 module
                # pull the broken ModuleVersion   = "$((Get-Module $ExoPSModule.FullName -ListAvailable).Version.ToString())" ;
                $NewExoPSModuleManifestProps = @{
                    Path        = $ExoPSModuleManifest ;
                    RootModule  = $ExoPSModule.Name
                    Author      = 'Jeremy Bradshaw (https://github.com/JeremyTBradshaw)' ;
                    CompanyName = 'jb365' ;
                } ;
                if (Get-Content "$($env:userprofile)\Documents\WindowsPowerShell\Modules\exoMFAModule\Microsoft.Exchange.Management.ExoPowershellModule.manifest" | Select-String '<assemblyIdentity\sname="mscorlib"\spublicKeyToken="b77a5c561934e089"\sversion="(\d\.\d\.\d\.\d)"\s/>' | Where-Object { $_ -match '(\d\.\d\.\d\.\d)' }) {
                    $NewExoPSModuleManifestProps.add('ModuleVersion', $matches[0]) ;
                } ;
            } else {
                # roll an otf psd1+psm1 module
                $NewExoPSModuleManifestProps = @{
                    Path          = $ExoPSModuleManifest ;
                    RootModule    = $ExoPSModule.Name
                    ModuleVersion = "$((Get-Module $ExoPSModule.FullName -ListAvailable).Version.ToString())" ;
                    Author        = 'Jeremy Bradshaw (https://github.com/JeremyTBradshaw)' ;
                    CompanyName   = 'jb365' ;
                } ;
            } ;
            if ($showDebug) { write-host -foregroundcolor green "New-ModuleManifest w`n$(($NewExoPSModuleManifestProps|out-string).trim())" } ;
            New-ModuleManifest @NewExoPSModuleManifestProps ;
            Import-Module $ExoPSModule.FullName -Global -ErrorAction:Stop ;
            $CreateExoPSSessionPs1 = Get-ChildItem -Path $ExoPSModule.PSParentPath -Filter 'CreateExoPSSession.ps1' ;
            $CreateExoPSSessionManifest = $CreateExoPSSessionPs1.FullName -replace '\.ps1', '.psd1' ;
            $CreateExoPSSessionPs1 = $CreateExoPSSessionPs1 |
            Get-Content | Where-Object { -not ($_ -like 'Write-Host*') } ;
            $CreateExoPSSessionPs1 -join "`n" |
            Set-Content -Path "$($CreateExoPSSessionManifest -replace '\.psd1','.psm1')" ;
            $NewCreateExoPSSessionManifest = @{
                Path          = $CreateExoPSSessionManifest ;
                RootModule    = Split-Path -Path ($CreateExoPSSessionManifest -replace '\.psd1', '.psm1') -Leaf ;
                ModuleVersion = '1.0' ;
                Author        = 'Todd Kadrie (https://github.com/tostka)' ;
                CompanyName   = 'toddomation.com' ;
            } ;
            if ($showDebug) { write-host -foregroundcolor green "New-ModuleManifest w`n$(($NewCreateExoPSSessionManifest|out-string).trim())" } ;
            New-ModuleManifest @NewCreateExoPSSessionManifest ;
            Import-Module "$($ExoPSModule.PSParentPath)\CreateExoPSSession.psm1" -Global -ErrorAction:Stop ;
        } catch {
            Write-Warning -Message "Tried but failed to import the EXO PS module.`n`nError message:" ;
            throw $_ ;
        } ;

        try {
            $global:UserPrincipalName = $Credential.Username ;
            $global:ConnectionUri = 'https://outlook.office365.com/PowerShell-LiveId' ;
            $global:AzureADAuthorizationEndpointUri = 'https://login.windows.net/common' ;
            $global:PSSessionOption = New-PSSessionOption -CancelTimeout 5000 -IdleTimeout 43200000 ;
            $global:BypassMailboxAnchoring = $false ;
            $ExoPSSession = @{
                UserPrincipalName               = $global:UserPrincipalName ;
                ConnectionUri                   = $global:ConnectionUri ;
                AzureADAuthorizationEndpointUri = $global:AzureADAuthorizationEndpointUri ;
                PSSessionOption                 = $global:PSSessionOption ;
                BypassMailboxAnchoring          = $global:BypassMailboxAnchoring ;
            } ;
            if ($showDebug) { write-host -foregroundcolor green "New-ExoPSSession w`n$(($ExoPSSession|out-string).trim())" } ;
            $ExoPSSession = New-ExoPSSession @ExoPSSession -ErrorAction:Stop ;
            if ($showDebug) { write-host -foregroundcolor green "Import-PSSession w`n$(($ImportPSSessionProps|out-string).trim())" } ;
            Import-Module (Import-PSSession $ExoPSSession @ImportPSSessionProps) -Prefix $CommandPrefix -Global -DisableNameChecking -ErrorAction:Stop ;
            UpdateImplicitRemotingHandler ;
            Add-PSTitleBar $sTitleBarTag ;
        } catch {
            Write-Warning -Message "Failed to connect to EXO via the imported EXO PS module.`n`nError message:" ;
            throw $_ ;
        } ;

    } else {
        $EXOsplat = @{
            ConfigurationName = "Microsoft.Exchange" ;
            ConnectionUri     = "https://ps.outlook.com/powershell/" ;
            Authentication    = "Basic" ;
            AllowRedirection  = $true;
        } ;

        # just use the passed $Credential vari
        $EXOsplat.Add("Credential", $Credential);

        If ($ProxyEnabled) {
            $EXOsplat.Add("sessionOption", $(New-PsSessionOption -ProxyAccessType IEConfig -ProxyAuthentication basic));
            Write-Host "Connecting to Exchange Online via Proxy"  ;
        } Else {
            Write-Host "Connecting to Exchange Online"  ;
        } ;
        if ($showDebug) {
            write-host -foregroundcolor green "`n$((get-date).ToString('HH:mm:ss')):New-PSSession w`n$(($EXOsplat|out-string).trim())" ;
        } ;
        Try {
            #$global:ExoPSSession = New-PSSession @EXOsplat ;
            $global:EOLSession = New-PSSession @EXOsplat ;
        } catch {
            Write-Warning -Message "Tried but failed to import the EXO PS module.`n`nError message:" ;
            throw $_ ;
        } ;
        if ($error.count -ne 0) {
            if ($error[0].FullyQualifiedErrorId -eq '-2144108477,PSSessionOpenFailed') {
                write-warning "$((get-date).ToString('HH:mm:ss')):AUTH FAIL BAD PASSWORD? ABORTING TO AVOID LOCKOUT!" ;
                throw "$((get-date).ToString('HH:mm:ss')):AUTH FAIL BAD PASSWORD? ABORTING TO AVOID LOCKOUT!" ;
                EXIT ;
            } ;
        } ;
        if(!$global:EOLSession){
            write-warning "$((get-date).ToString('HH:mm:ss')):FAILED TO RETURN PSSESSION!`nAUTH FAIL BAD PASSWORD? ABORTING TO AVOID LOCKOUT!" ;
            throw "$((get-date).ToString('HH:mm:ss')):AUTH FAIL BAD PASSWORD? ABORTING TO AVOID LOCKOUT!" ;
            EXIT ;
        } ; 
        $pltPSS = [ordered]@{
            Session             = $global:EOLSession ;
            Prefix              = $CommandPrefix ;
            DisableNameChecking = $true  ;
            AllowClobber        = $true ;
            ErrorAction         = 'Stop' ;
        } ;
        if ($showDebug) {
            write-host -foregroundcolor green "`n$((get-date).ToString('HH:mm:ss')):Import-PSSession w`n$(($pltPSS|out-string).trim())" ;
        } ;
        Try {
            $Global:EOLModule = Import-Module (Import-PSSession @pltPSS) -Global -Prefix $CommandPrefix -PassThru -DisableNameChecking   ;
            Add-PSTitleBar $sTitleBarTag ;
        } catch {
            Write-Warning -Message "Tried but failed to import the EXO PS module.`n`nError message:" ;
            throw $_ ;
        } ;
    } ;

} ; #*------^ END Function Connect-EXO ^------
if(!(get-alias | Where-Object{$_.name -like "cxo"})) {Set-Alias 'cxo' -Value 'Connect-EXO' ; }

#*------^ Connect-EXO.ps1 ^------

#*------v Connect-EXO2.ps1 v------
Function Connect-EXO2 {
    <#
    .SYNOPSIS
    Connect-EXO2 - Establish session with EXO v2 module
    .NOTES
    Version     : 1.0.0
    Author      : Todd Kadrie
    Website     :	http://www.toddomation.com
    Twitter     :	@tostka / http://twitter.com/tostka
    CreatedDate : 2020-04-28
    FileName    :
    License     :
    Copyright   :
    Github      : https://github.com/tostka
    Tags        : Powershell,ExchangeOnline,Exchange,RemotePowershell,Connection,MFA
    AddedCredit : ExactMike Perficient
    AddedWebsite:	https://social.technet.microsoft.com/Forums/msonline/en-US/f3292898-9b8c-482a-86f0-3caccc0bd3e5/exchange-powershell-monitoring-remote-sessions?forum=onlineservicesexchange
    AddedTwitter:
    AddedCredit2 : Jeremy Bradshaw
    AddedWebsite2:	https://github.com/JeremyTBradshaw
    AddedTwitter2:
    REVISIONS   :
    * 7:13 AM 7/22/2020 replaced codeblock w get-TenantTag()
    * 5:14 PM 7/21/2020 added VEN supp
    * 3:42 PM 4/28/2020 update to EXOv2
    * 8:38 AM 4/17/2020 added a new test of $global:EOLSession, to detect initial cred fail (pw chg, outofdate creds, locked out)
    * 8:45 AM 3/3/2020 public cleanup, refactored Connect-EXO2 for Meta's
    * 9:52 PM 1/16/2020 cleanup
    * 10:55 AM 12/6/2019 Connect-EXO2:added suffix to TitleBar tag for other tenants, also config'd a central tab vari
    * 9:17 AM 12/4/2019 CONSISTENTLY failing to load properly in lab, on lynms6200d - wont' get-module xxxx -listinstalled, even after load, so I rewrote an exemption diverting into the locally installed $env:userprofile\documents\WindowsPowerShell\Modules\exoMFAModule\ copy.
    * 5:14 PM 11/27/2019 repl $MFA code with get-TenantMFARequirement
    * 1:07 PM 11/25/2019 added tenant-specific alias variants for connect & reconnect
    # 1:26 PM 11/19/2019 added MFA detection fr infastrings .ps1 globals, lifted from Jeremy Bradshaw (https://github.com/JeremyTBradshaw)'s Connect-Exchange()
    # 10:35 AM 6/20/2019 added $pltPSS splat dump to the import-pssession cmd block; hard-typed the $Credential [System.Management.Automation.PSCredential]
    # 8:22 AM 11/20/2017 spliced in retry loop into reConnect-EXO2 as well, to avoid using any state testing in scripts, localize it 1x here.
    # 1:49 PM 11/2/2017 coded around non-profile gaps from not having get-admincred() - added the prompt code in to fake it
    # 12:26 PM 9/11/2017 debugged retry - catch doesn't fire properly on new-Pssession, have to test the $error state, to detect auth fails (assuming the bad pw error# is specific). $error test is currently targeting specific error returned on a bad password. Added retry, for when connection won't hold and fails breaks - need to watch out that bad pw doesn't lock out the acct!
    # 12:50 PM 6/2/2017 expanded pshelp, added Examples, cleaned up deadwood
    * # 12:10 PM 3/15/2017 Connect-EXO2 typo, disable prefix auto spec, unless actually blanked. switch ProxyEnabled to non-Mandatory.
    .DESCRIPTION
    Connect-EXO2 - Establish session with EXO v2 module
    Original concept based on 'overlapping functions' concept by: ExactMike Perficient, Global Knowl... (Partner)
    Website:	https://social.technet.microsoft.com/Forums/msonline/en-US/f3292898-9b8c-482a-86f0-3caccc0bd3e5/exchange-powershell-monitoring-remote-sessions?forum=onlineservicesexchange
    Extended with Jeremy Bradshaw's on-the-fly EXO MFA module concept (https://github.com/JeremyTBradshaw)'s Connect-Exchange()
    .PARAMETER  ProxyEnabled
    Use Proxy-Aware SessionOption settings [-ProxyEnabled]
    .PARAMETER  CommandPrefix
    [verb]-PREFIX[command] PREFIX string for clearly marking cmdlets sourced in this connection [-CommandPrefix tag]
    .PARAMETER  Credential
    Credential to use for this connection [-credential 's-todd.kadrie@toro.com']
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    None. Returns no objects or output.
    .EXAMPLE
    Connect-EXO2
    Connect using defaults, and leverage any pre-set $global:credo365TORSID variable
    .EXAMPLE
    Connect-EXO2 -CommandPrefix exo -credential (Get-Credential -credential s-todd.kadrie@torolab.com)  ;
    Connect an explicit credential, and use 'exolab' as the cmdlet prefix
    .EXAMPLE
    $cred = get-credential -credential $o365_Torolab_SIDUpn ;
    Connect-EXO2 -credential $cred ;
    Pass in a prefab credential object (useful for auto-shifting to MFA - the function will autoresolve MFA reqs based on the cred domain)
    .LINK
    https://social.technet.microsoft.com/Forums/msonline/en-US/f3292898-9b8c-482a-86f0-3caccc0bd3e5/exchange-powershell-monitoring-remote-sessions?forum=onlineservicesexchange
    .LINK
    https://github.com/JeremyTBradshaw
    #>
    Param(
        [Parameter(HelpMessage = "Use Proxy-Aware SessionOption settings [-ProxyEnabled]")][boolean]$ProxyEnabled = $False,
        [Parameter(HelpMessage = "[verb]-PREFIX[command] PREFIX string for clearly marking cmdlets sourced in this connection [-CommandPrefix tag]")][string]$CommandPrefix = 'exo',
        [Parameter(HelpMessage = "Credential to use for this connection [-credential [credential obj variable]")][System.Management.Automation.PSCredential]$Credential = $global:credo365TORSID,
        [Parameter(HelpMessage = "Debugging Flag [-showDebug]")]
        [switch] $showDebug
    ) ;

    $MFA = get-TenantMFARequirement -Credential $Credential ;

    # disable prefix spec, unless actually blanked (e.g. centrally spec'd in profile).
    if (!$CommandPrefix) {
        $CommandPrefix = 'exo' ;
        write-verbose -verbose:$true  "(asserting Prefix:$($CommandPrefix)" ;
    } ;

    $sTitleBarTag = "EXO" ;
    $TentantTag=get-TenantTag -Credential $Credential ; 
    if($TentantTag -ne 'TOR'){
        # explicitly leave this tenant (default) untagged
        $sTitleBarTag += $TentantTag ;
    } ; 

    $ImportPSSessionProps = @{
        AllowClobber        = $true ;
        DisableNameChecking = $true ;
        Prefix              = $CommandPrefix ;
        ErrorAction         = 'Stop' ;
    } ;

    if ($MFA) {
        try {
            $ExoPSModuleSearchProperties = @{
                Path        = "$($env:LOCALAPPDATA)\Apps\2.0\" ;
                Filter      = 'Microsoft.Exchange.Management.ExoPowerShellModule.dll' ;
                Recurse     = $true ;
                ErrorAction = 'Stop' ;
            } ;

            if ($showDebug) { write-host -foregroundcolor green "Get-ChildItem w`n$(($ExoPSModuleSearchProperties|out-string).trim())" } ;
            $ExoPSModule = Get-ChildItem @ExoPSModuleSearchProperties |
            Where-Object { $_.FullName -notmatch '_none_' } |
            Sort-Object LastWriteTime |
            Select-Object -Last 1 ;
            Import-Module $ExoPSModule.FullName -ErrorAction:Stop ;
            $ExoPSModuleManifest = $ExoPSModule.FullName -replace '\.dll', '.psd1' ;
            if (!(Get-Module $ExoPSModule.FullName -ListAvailable -ErrorAction 0 )) {
                write-verbose -verbose:$true  "Unable to`nGet-Module $($ExoPSModule.FullName) -ListAvailable`ndiverting to hardcoded exoMFAModule`nRequires that it be locally copied below`n$env:userprofile\documents\WindowsPowerShell\Modules\exoMFAModule\`n " ;
                # go to a hard load path $env:userprofile\documents\WindowsPowerShell\Modules\exoMFAModule\
                $ExoPSModuleSearchProperties = @{
                    Path        = "$($env:userprofile)\documents\WindowsPowerShell\Modules\exoMFAModule\" ;
                    Filter      = 'Microsoft.Exchange.Management.ExoPowerShellModule.dll' ;
                    Recurse     = $true ;
                    ErrorAction = 'Stop' ;
                } ;
                $ExoPSModule = Get-ChildItem @ExoPSModuleSearchProperties |
                Where-Object { $_.FullName -notmatch '_none_' } |
                Sort-Object LastWriteTime |
                Select-Object -Last 1 ;
                # roll an otf psd1+psm1 module
                # pull the broken ModuleVersion   = "$((Get-Module $ExoPSModule.FullName -ListAvailable).Version.ToString())" ;
                $NewExoPSModuleManifestProps = @{
                    Path        = $ExoPSModuleManifest ;
                    RootModule  = $ExoPSModule.Name
                    Author      = 'Jeremy Bradshaw (https://github.com/JeremyTBradshaw)' ;
                    CompanyName = 'jb365' ;
                } ;
                if (Get-Content "$($env:userprofile)\Documents\WindowsPowerShell\Modules\exoMFAModule\Microsoft.Exchange.Management.ExoPowershellModule.manifest" | Select-String '<assemblyIdentity\sname="mscorlib"\spublicKeyToken="b77a5c561934e089"\sversion="(\d\.\d\.\d\.\d)"\s/>' | Where-Object { $_ -match '(\d\.\d\.\d\.\d)' }) {
                    $NewExoPSModuleManifestProps.add('ModuleVersion', $matches[0]) ;
                } ;
            } else {
                # roll an otf psd1+psm1 module
                $NewExoPSModuleManifestProps = @{
                    Path          = $ExoPSModuleManifest ;
                    RootModule    = $ExoPSModule.Name
                    ModuleVersion = "$((Get-Module $ExoPSModule.FullName -ListAvailable).Version.ToString())" ;
                    Author        = 'Jeremy Bradshaw (https://github.com/JeremyTBradshaw)' ;
                    CompanyName   = 'jb365' ;
                } ;
            } ;
            if ($showDebug) { write-host -foregroundcolor green "New-ModuleManifest w`n$(($NewExoPSModuleManifestProps|out-string).trim())" } ;
            New-ModuleManifest @NewExoPSModuleManifestProps ;
            Import-Module $ExoPSModule.FullName -Global -ErrorAction:Stop ;
            $CreateExoPSSessionPs1 = Get-ChildItem -Path $ExoPSModule.PSParentPath -Filter 'CreateExoPSSession.ps1' ;
            $CreateExoPSSessionManifest = $CreateExoPSSessionPs1.FullName -replace '\.ps1', '.psd1' ;
            $CreateExoPSSessionPs1 = $CreateExoPSSessionPs1 |
            Get-Content | Where-Object { -not ($_ -like 'Write-Host*') } ;
            $CreateExoPSSessionPs1 -join "`n" |
            Set-Content -Path "$($CreateExoPSSessionManifest -replace '\.psd1','.psm1')" ;
            $NewCreateExoPSSessionManifest = @{
                Path          = $CreateExoPSSessionManifest ;
                RootModule    = Split-Path -Path ($CreateExoPSSessionManifest -replace '\.psd1', '.psm1') -Leaf ;
                ModuleVersion = '1.0' ;
                Author        = 'Todd Kadrie (https://github.com/tostka)' ;
                CompanyName   = 'toddomation.com' ;
            } ;
            if ($showDebug) { write-host -foregroundcolor green "New-ModuleManifest w`n$(($NewCreateExoPSSessionManifest|out-string).trim())" } ;
            New-ModuleManifest @NewCreateExoPSSessionManifest ;
            Import-Module "$($ExoPSModule.PSParentPath)\CreateExoPSSession.psm1" -Global -ErrorAction:Stop ;
        } catch {
            Write-Warning -Message "Tried but failed to import the EXO PS module.`n`nError message:" ;
            throw $_ ;
        } ;

        try {
            $global:UserPrincipalName = $Credential.Username ;
            $global:ConnectionUri = 'https://outlook.office365.com/PowerShell-LiveId' ;
            $global:AzureADAuthorizationEndpointUri = 'https://login.windows.net/common' ;
            $global:PSSessionOption = New-PSSessionOption -CancelTimeout 5000 -IdleTimeout 43200000 ;
            $global:BypassMailboxAnchoring = $false ;
            $ExoPSSession = @{
                UserPrincipalName               = $global:UserPrincipalName ;
                ConnectionUri                   = $global:ConnectionUri ;
                AzureADAuthorizationEndpointUri = $global:AzureADAuthorizationEndpointUri ;
                PSSessionOption                 = $global:PSSessionOption ;
                BypassMailboxAnchoring          = $global:BypassMailboxAnchoring ;
            } ;
            if ($showDebug) { write-host -foregroundcolor green "New-ExoPSSession w`n$(($ExoPSSession|out-string).trim())" } ;
            $ExoPSSession = New-ExoPSSession @ExoPSSession -ErrorAction:Stop ;
            if ($showDebug) { write-host -foregroundcolor green "Import-PSSession w`n$(($ImportPSSessionProps|out-string).trim())" } ;
            Import-Module (Import-PSSession $ExoPSSession @ImportPSSessionProps) -Prefix $CommandPrefix -Global -DisableNameChecking -ErrorAction:Stop ;
            UpdateImplicitRemotingHandler ;
            Add-PSTitleBar $sTitleBarTag ;
        } catch {
            Write-Warning -Message "Failed to connect to EXO via the imported EXO PS module.`n`nError message:" ;
            throw $_ ;
        } ;

    } else {
        $EXOsplat = @{
            ConfigurationName = "Microsoft.Exchange" ;
            ConnectionUri     = "https://ps.outlook.com/powershell/" ;
            Authentication    = "Basic" ;
            AllowRedirection  = $true;
        } ;

        # just use the passed $Credential vari
        $EXOsplat.Add("Credential", $Credential);

        If ($ProxyEnabled) {
            $EXOsplat.Add("sessionOption", $(New-PsSessionOption -ProxyAccessType IEConfig -ProxyAuthentication basic));
            Write-Host "Connecting to Exchange Online via Proxy"  ;
        } Else {
            Write-Host "Connecting to Exchange Online"  ;
        } ;
        if ($showDebug) {
            write-host -foregroundcolor green "`n$((get-date).ToString('HH:mm:ss')):New-PSSession w`n$(($EXOsplat|out-string).trim())" ;
        } ;
        Try {
            #$global:ExoPSSession = New-PSSession @EXOsplat ;
            $global:EOLSession = New-PSSession @EXOsplat ;
        } catch {
            Write-Warning -Message "Tried but failed to import the EXO PS module.`n`nError message:" ;
            throw $_ ;
        } ;
        if ($error.count -ne 0) {
            if ($error[0].FullyQualifiedErrorId -eq '-2144108477,PSSessionOpenFailed') {
                write-warning "$((get-date).ToString('HH:mm:ss')):AUTH FAIL BAD PASSWORD? ABORTING TO AVOID LOCKOUT!" ;
                throw "$((get-date).ToString('HH:mm:ss')):AUTH FAIL BAD PASSWORD? ABORTING TO AVOID LOCKOUT!" ;
                EXIT ;
            } ;
        } ;
        if(!$global:EOLSession){
            write-warning "$((get-date).ToString('HH:mm:ss')):FAILED TO RETURN PSSESSION!`nAUTH FAIL BAD PASSWORD? ABORTING TO AVOID LOCKOUT!" ;
            throw "$((get-date).ToString('HH:mm:ss')):AUTH FAIL BAD PASSWORD? ABORTING TO AVOID LOCKOUT!" ;
            EXIT ;
        } ; 
        $pltPSS = [ordered]@{
            Session             = $global:EOLSession ;
            Prefix              = $CommandPrefix ;
            DisableNameChecking = $true  ;
            AllowClobber        = $true ;
            ErrorAction         = 'Stop' ;
        } ;
        if ($showDebug) {
            write-host -foregroundcolor green "`n$((get-date).ToString('HH:mm:ss')):Import-PSSession w`n$(($pltPSS|out-string).trim())" ;
        } ;
        Try {
            $Global:EOLModule = Import-Module (Import-PSSession @pltPSS) -Global -Prefix $CommandPrefix -PassThru -DisableNameChecking   ;
            Add-PSTitleBar $sTitleBarTag ;
        } catch {
            Write-Warning -Message "Tried but failed to import the EXO PS module.`n`nError message:" ;
            throw $_ ;
        } ;
    } ;

} ; #*------^ END Function Connect-EXO2 ^------
if(!(get-alias | Where-Object{$_.name -like "cxo"})) {Set-Alias 'cxo' -Value 'Connect-EXO2' ; }

#*------^ Connect-EXO2.ps1 ^------

#*------v cxocmw.ps1 v------
function cxoCMW {Connect-EXO -cred $credO365CMWCSID}

#*------^ cxocmw.ps1 ^------

#*------v cxotol.ps1 v------
function cxoTOL {Connect-EXO -cred $credO365TOLSID}

#*------^ cxotol.ps1 ^------

#*------v cxotor.ps1 v------
function cxoTOR {Connect-EXO -cred $credO365TORSID}

#*------^ cxotor.ps1 ^------

#*------v cxoVEN.ps1 v------
function cxoVEN {Connect-EXO -cred $credO365VENCSID}

#*------^ cxoVEN.ps1 ^------

#*------v Disconnect-EXO.ps1 v------
Function Disconnect-EXO {
    <#
    .SYNOPSIS
    Disconnect-EXO - Disconnects any PSS to https://ps.outlook.com/powershell/ (cleans up session after a batch or other temp work is done)
    .NOTES
    Version     : 1.0.0
    Author      : Todd Kadrie
    Website     :	http://www.toddomation.com
    Twitter     :	@tostka / http://twitter.com/tostka
    CreatedDate : 2020-03-03
    FileName    : 
    License     : 
    Copyright   : 
    Github      : https://github.com/tostka
    Tags        : Powershell,ExchangeOnline,Exchange,RemotePowershell,Connection,MFA
    AddedCredit : ExactMike Perficient
    AddedWebsite:	https://social.technet.microsoft.com/Forums/msonline/en-US/f3292898-9b8c-482a-86f0-3caccc0bd3e5/exchange-powershell-monitoring-remote-sessions?forum=onlineservicesexchange
    AddedTwitter:	
    REVISIONS   :
    * 11:50 AM 5/27/2020 added alias:dxo win func
    * 2:34 PM 4/20/2020 added local $rgxExoPsHostName
    * 8:45 AM 3/3/2020 public cleanup
    * 9:52 AM 11/20/2019 reviewed for credential matl, no way to see the credential on a given pssession, so there's no way to target and disconnect discretely. It's a shotgun close.
    # 10:27 AM 6/20/2019 switched to common $rgxExoPsHostName
    # 1:12 PM 11/7/2018 added Disconnect-PssBroken
    # 11:23 AM 7/10/2018: made exo-only (was overlapping with CCMS)
    # 8:47 AM 6/2/2017 cleaned up deadwood, simplified pshelp
    * 8:49 AM 3/15/2017 Disconnect-EXO: add Remove-PSTitleBar 'EXO' to clean up on disconnect
    * 2/10/14 posted version
    .DESCRIPTION
    Used to smoothly cleanup connections (at end, or when expired, to purge for a fresh pass).
    Mike's original notes:
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    None. Returns no objects or output.
    .EXAMPLE
    Disconnect-EXO;
    .LINK
    https://social.technet.microsoft.com/Forums/msonline/en-US/f3292898-9b8c-482a-86f0-3caccc0bd3e5/exchange-powershell-monitoring-remote-sessions?forum=onlineservicesexchange
    #>
    [CmdletBinding()]
    [Alias('dxo')]
    Param() 
    if(!$rgxExoPsHostName){$rgxExoPsHostName="^(ps\.outlook\.com|outlook\.office365\.com)$" } ;
    if($Global:EOLModule){$Global:EOLModule | Remove-Module -Force ; } ;
    if($global:EOLSession){$global:EOLSession | Remove-PSSession ; } ;
    Get-PSSession |Where-Object{$_.ComputerName -match $rgxExoPsHostName } | Remove-PSSession ;
    Disconnect-PssBroken ;
    Remove-PSTitlebar 'EXO' ;
}

#*------^ Disconnect-EXO.ps1 ^------

#*------v Reconnect-EXO.ps1 v------
Function Reconnect-EXO {
   <#
    .SYNOPSIS
    Reconnect-EXO - Test and reestablish PSS to https://ps.outlook.com/powershell/
    .NOTES
    Updated By: : Todd Kadrie
    Website:	http://toddomation.com
    Twitter:	@tostka http://twitter.com/tostka
    Based on original function Author: ExactMike Perficient, Global Knowl... (Partner)
    Website:	https://social.technet.microsoft.com/Forums/msonline/en-US/f3292898-9b8c-482a-86f0-3caccc0bd3e5/exchange-powershell-monitoring-remote-sessions?forum=onlineservicesexchange
    REVISIONS   :
    * 11:48 AM 5/27/2020 added func alias:rxo within the func
    * 2:38 PM 4/20/2020 added local $rgxExoPsHostName
    * 8:45 AM 3/3/2020 public cleanup
    * 9:52 PM 1/16/2020 cleanup
    * 1:07 PM 11/25/2019 added *tol/*tor/*cmw alias variants for connect & reconnect
    * 9:52 AM 11/20/2019 spliced in credential matl
    * 2:55 PM 10/11/2018 connect-exo: added pre sleep skip on tryno 1
    * 8:04 AM 11/20/2017 code in a loop in the reconnect-exo, until it hits or 100% times out
    # 8:47 AM 6/2/2017 cleaned up deadwood, simplified pshelp
    * 7:58 AM 3/15/2017 ren Disconnect/Connect/Reconnect-EXO => Disconnect/Connect/Reconnect-EXO, added pss, doing tweaks to put into prod use
    * 2/10/14 posted version
    .DESCRIPTION
    I use this for routine test/reconnect of EXO. His orig use was within batches, to break up and requeue chunks of commands.
    Mike's original comment: Below is one
    example of how I batch items for processing and use the
    Reconnect-EXO function.  I'm still experimenting with how to best
    batch items and you can see here I'm using a combination of larger batches for
    Write-Progress and actually handling each individual item within the
    foreach-object script block.  I was driven to this because disconnections
    happen so often/so unpredictably in my current customer's environment:
    .PARAMETER  Credential
    Credential to use for this connection [-credential 'account@domain.com']
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    None. Returns no objects or output.
    .EXAMPLE
    Reconnect-EXO;
    Reconnect EXO connection
    .EXAMPLE
    $batchsize = 10 ;
    $RecordCount=$mr.count #this is the array of whatever you are processing ;
    $b=0 #this is the initialization of a variable used in the do until loop below ;
    $mrs = @() ;
    do {
        Write-Progress -Activity "Getting move request statistics for all $wave move requests." -Status "Processing Records $b through $($b+$batchsize) of $RecordCount." -PercentComplete ($b/$RecordCount*100) ;
        $mrs += $mr | Select-Object -skip $b -first $batchsize | foreach-object {Reconnect-EXO; $_ | Get-OLMoveRequestStatistics} ;
        $b=$b+$batchsize ;
        } ;
    until ($b -gt $RecordCount) ;
    
    .LINK
    https://social.technet.microsoft.com/Forums/msonline/en-US/f3292898-9b8c-482a-86f0-3caccc0bd3e5/exchange-powershell-monitoring-remote-sessions?forum=onlineservicesexchange
    #>
    [CmdletBinding()]
    [Alias('rxo')]
    Param(
      [Parameter(HelpMessage="Use Proxy-Aware SessionOption settings [-ProxyEnabled]")]
      [boolean]$ProxyEnabled = $False,
      [Parameter(HelpMessage="Credential to use for this connection [-credential [credential obj variable]")][System.Management.Automation.PSCredential]$Credential = $global:credo365TORSID,
      [Parameter(HelpMessage="Debugging Flag [-showDebug]")]
      [switch] $showDebug
    ) ;

    if(!$rgxExoPsHostName){$rgxExoPsHostName="^(ps\.outlook\.com|outlook\.office365\.com)$" } ;
    
    # fault tolerant looping exo connect, don't let it exit until a connection is present, and stable, or return error for hard time out
    $tryNo=0 ;
    Do {
        $tryNo++ ;
        write-host "." -NoNewLine; if($tryNo -gt 1){Start-Sleep -m (1000 * 5)} ;
        # appears MFA may not properly support passing back a session vari, so go right to strict hostname matches
        if( !(Get-PSSession|Where-Object{($_.ComputerName -match $rgxExoPsHostName) -AND ($_.State -eq 'Opened') -AND ($_.Availability -eq 'Available')}) ){
            if($showdebug){ write-host -foregroundcolor yellow "$((get-date).ToString('HH:mm:ss')):Reconnecting:No existing PSSESSION matching $($rgxExoPsHostName) with valid Open/Availability:$((Get-PSSession|Where-Object{$_.ComputerName -match $rgxExoPsHostName}| Format-Table -a State,Availability |out-string).trim())" } ;
            Disconnect-Exo ; Disconnect-PssBroken ;Start-Sleep -Seconds 3;
            if(!$Credential){
                connect-EXO ;
            } else {
                connect-EXO -credential:$($Credential) ;
            } ;
        }  ;
        if($tryNo -gt $DoRetries ){throw "RETRIED EXO CONNECT $($tryNo) TIMES, ABORTING!" } ;
    } Until ((Get-PSSession |Where-Object{$_.ComputerName -match $rgxExoPsHostName -AND $_.State -eq "Opened" -AND $_.Availability -eq "Available"}))

}

#*------^ Reconnect-EXO.ps1 ^------

#*------v rxocmw.ps1 v------
function rxoCMW {Reconnect-EXO -cred $credO365CMWCSID}

#*------^ rxocmw.ps1 ^------

#*------v rxotol.ps1 v------
function rxoTOL {Reconnect-EXO -cred $credO365TOLSID}

#*------^ rxotol.ps1 ^------

#*------v rxotor.ps1 v------
function rxoTOR {Reconnect-EXO -cred $credO365TORSID}

#*------^ rxotor.ps1 ^------

#*------v rxoVEN.ps1 v------
function rxoVEN {Reconnect-EXO -cred $credO365VENCSID}

#*------^ rxoVEN.ps1 ^------

#*======^ END FUNCTIONS ^======

Export-ModuleMember -Function Connect-EXO,Connect-EXO2,cxoCMW,cxoTOL,cxoTOR,cxoVEN,Disconnect-EXO,Reconnect-EXO,rxoCMW,rxoTOL,rxoTOR,rxoVEN -Alias *


# SIG # Begin signature block
# MIIELgYJKoZIhvcNAQcCoIIEHzCCBBsCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU82pmblFfNYCSyE0gDhbxyR1Z
# eNWgggI4MIICNDCCAaGgAwIBAgIQWsnStFUuSIVNR8uhNSlE6TAJBgUrDgMCHQUA
# MCwxKjAoBgNVBAMTIVBvd2VyU2hlbGwgTG9jYWwgQ2VydGlmaWNhdGUgUm9vdDAe
# Fw0xNDEyMjkxNzA3MzNaFw0zOTEyMzEyMzU5NTlaMBUxEzARBgNVBAMTClRvZGRT
# ZWxmSUkwgZ8wDQYJKoZIhvcNAQEBBQADgY0AMIGJAoGBALqRVt7uNweTkZZ+16QG
# a+NnFYNRPPa8Bnm071ohGe27jNWKPVUbDfd0OY2sqCBQCEFVb5pqcIECRRnlhN5H
# +EEJmm2x9AU0uS7IHxHeUo8fkW4vm49adkat5gAoOZOwbuNntBOAJy9LCyNs4F1I
# KKphP3TyDwe8XqsEVwB2m9FPAgMBAAGjdjB0MBMGA1UdJQQMMAoGCCsGAQUFBwMD
# MF0GA1UdAQRWMFSAEL95r+Rh65kgqZl+tgchMuKhLjAsMSowKAYDVQQDEyFQb3dl
# clNoZWxsIExvY2FsIENlcnRpZmljYXRlIFJvb3SCEGwiXbeZNci7Rxiz/r43gVsw
# CQYFKw4DAh0FAAOBgQB6ECSnXHUs7/bCr6Z556K6IDJNWsccjcV89fHA/zKMX0w0
# 6NefCtxas/QHUA9mS87HRHLzKjFqweA3BnQ5lr5mPDlho8U90Nvtpj58G9I5SPUg
# CspNr5jEHOL5EdJFBIv3zI2jQ8TPbFGC0Cz72+4oYzSxWpftNX41MmEsZkMaADGC
# AWAwggFcAgEBMEAwLDEqMCgGA1UEAxMhUG93ZXJTaGVsbCBMb2NhbCBDZXJ0aWZp
# Y2F0ZSBSb290AhBaydK0VS5IhU1Hy6E1KUTpMAkGBSsOAwIaBQCgeDAYBgorBgEE
# AYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwG
# CisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMCMGCSqGSIb3DQEJBDEWBBS6uK9Y
# Pw5/Hj8TViNYrHDLi27cADANBgkqhkiG9w0BAQEFAASBgESJdrDZ9OGwh25Slv6t
# 2aeXXr6HdkqhSxrGDvWn40HcEK4/06qsxbEOpnYYhRGf2j8/trYyXxSELSoav2Z3
# vWNmW5SthHRZ0BYcnzvlI/GvUGy3gLv9lVQ0r8eNhGmirFdHanlTs9aUfIgxdk4J
# bd3naV/XJzts5mKJFHkjpTgT
# SIG # End signature block
