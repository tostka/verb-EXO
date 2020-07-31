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
    * 4:27 PM 7/29/2020 added Catch workaround for EXO bug here:https://answers.microsoft.com/en-us/msoffice/forum/all/cannot-connect-to-exchange-online-via-powershell/25ca1cc2-e23a-470e-9c73-e6c56c4fbb46?page=7 Workaround 1) Use EXO V2 module - but it breaks historical use of -suffix 'exo' 2) use ?SerializationLevel=Full with the ConnectionURI: -ConnectionUri "https://outlook.office365.com/powershell-liveid?SerializationLevel=Full". Added Beg/Proc/End with trailing Tenant -cred align validation. Need to rewrite MFA, as the EXO V2 fundementally conflicts on a cmdlet that was part of the exoMFA mod, now uninstalled
    * 11:21 AM 7/28/2020 added Credential -> AcceptedDomains Tenant validation, also testing existing conn, and skipping reconnect unless unhealthy or wrong tenant to match credential
    * 3:24 PM 7/24/2020 updated to support tenant-alignment & sub'd out showdebug for verbose
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
        [Parameter(HelpMessage = "Use Proxy-Aware SessionOption settings [-ProxyEnabled]")]
        [boolean]$ProxyEnabled = $False,
        [Parameter(HelpMessage = "[verb]-PREFIX[command] PREFIX string for clearly marking cmdlets sourced in this connection [-CommandPrefix tag]")]
        [string]$CommandPrefix = 'exo',
        [Parameter(HelpMessage = "Credential to use for this connection [-credential [credential obj variable]")]
        [System.Management.Automation.PSCredential]$Credential = $global:credo365TORSID,
        [Parameter(HelpMessage = "Debugging Flag [-showDebug]")]
        [switch] $showDebug
    ) ;
    BEGIN{
        $verbose = ($VerbosePreference -eq "Continue") ; 
        if(!$rgxExoPsHostName){$rgxExoPsHostName="^(ps\.outlook\.com|outlook\.office365\.com)$" } ;
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
    } ;  # BEG-E
    PROCESS{
        $bExistingEXOGood = $false ; 
        if( Get-PSSession|Where-Object{($_.ComputerName -match $rgxExoPsHostName) -AND ($_.State -eq 'Opened') -AND ($_.Availability -eq 'Available')}){
            if( get-command Get-exoAcceptedDomain) {
                if ((Get-exoAcceptedDomain).domainname.contains($Credential.username.split('@')[1].tostring())){
                    # validate that the connected EXO is to the $Credential tenant    
                    write-verbose "(Existing EXO Authenticated & Functional:$($Credential.username.split('@')[1].tostring()))" ; 
                    $bExistingEXOGood = $true ; 
                } else { 
                    write-verbose "(Credential mismatch:disconnecting from existing EXO:$($eEXO.Identity) tenant)" ; 
                    Disconnect-exo ; 
                    $bExistingEXOGood = $false ; 
                } ; 
            } else { 
                # capture outlier: shows a session wo the test cmdlet, force reset
                Disconnect-exo ; 
                $bExistingEXOGood = $false ; 
            } ; 
        } ; 

        if($bExistingEXOGood -eq $false){
    
            $ImportPSSessionProps = @{
                AllowClobber        = $true ;
                DisableNameChecking = $true ;
                Prefix              = $CommandPrefix ;
                ErrorAction         = 'Stop' ;
            } ;

            if ($MFA) {
                
                throw "MFA is not currently supported by the connect-exo cmdlet!. Use connect-exo2 instead" ; 
                Exit 
                <# 4:24 PM 7/30/2020 HAD TO UNINSTALL THE EXOMFA module, a bundled cmdlet fundementally conflicted with ExchangeOnlineManagement
                try {
                    $ExoPSModuleSearchProperties = @{
                        Path        = "$($env:LOCALAPPDATA)\Apps\2.0\" ;
                        Filter      = 'Microsoft.Exchange.Management.ExoPowerShellModule.dll' ;
                        Recurse     = $true ;
                        ErrorAction = 'Stop' ;
                    } ;

                    write-verbose "Get-ChildItem w`n$(($ExoPSModuleSearchProperties|out-string).trim())" ;
                    $ExoPSModule = Get-ChildItem @ExoPSModuleSearchProperties |
                        Where-Object { $_.FullName -notmatch '_none_' } | Sort-Object LastWriteTime |
                            Select-Object -Last 1 ;
                    # suppress VerbosePreference:Continue, if set, during mod loads (VERY NOISEY)
                    if($VerbosePreference = "Continue"){
                        $VerbosePrefPrior = $VerbosePreference ;
                        $VerbosePreference = "SilentlyContinue" ;
                        $verbose = ($VerbosePreference -eq "Continue") ;
                    } ; 
                    Import-Module $ExoPSModule.FullName -ErrorAction:Stop ;
                    # reenable VerbosePreference:Continue, if set, during mod loads 
                    if($VerbosePrefPrior -eq "Continue"){
                        $VerbosePreference = $VerbosePrefPrior ;
                        $verbose = ($VerbosePreference -eq "Continue") ;
                    } ; 
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
                                Sort-Object LastWriteTime | Select-Object -Last 1 ;
                        # roll an otf psd1+psm1 module
                        # pull the broken ModuleVersion   = "$((Get-Module $ExoPSModule.FullName -ListAvailable).Version.ToString())" ;
                        $NewExoPSModuleManifestProps = @{
                            Path        = $ExoPSModuleManifest ;
                            RootModule  = $ExoPSModule.Name
                            Author      = 'Jeremy Bradshaw (https://github.com/JeremyTBradshaw)' ;
                            CompanyName = 'jb365' ;
                        } ;
                        if (Get-Content "$($env:userprofile)\Documents\WindowsPowerShell\Modules\exoMFAModule\Microsoft.Exchange.Management.ExoPowershellModule.manifest" | 
                            Select-String '<assemblyIdentity\sname="mscorlib"\spublicKeyToken="b77a5c561934e089"\sversion="(\d\.\d\.\d\.\d)"\s/>' | 
                                Where-Object { $_ -match '(\d\.\d\.\d\.\d)' }) {
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
                    write-verbose "New-ModuleManifest w`n$(($NewExoPSModuleManifestProps|out-string).trim())" ;
                    New-ModuleManifest @NewExoPSModuleManifestProps ;
                    # suppress VerbosePreference:Continue, if set, during mod loads (VERY NOISEY)
                    if($VerbosePreference = "Continue"){
                        $VerbosePrefPrior = $VerbosePreference ;
                        $VerbosePreference = "SilentlyContinue" ;
                        $verbose = ($VerbosePreference -eq "Continue") ;
                    } ; 
                    Import-Module $ExoPSModule.FullName -Global -ErrorAction:Stop ;
                    # reenable VerbosePreference:Continue, if set, during mod loads 
                    if($VerbosePrefPrior -eq "Continue"){
                        $VerbosePreference = $VerbosePrefPrior ;
                        $verbose = ($VerbosePreference -eq "Continue") ;
                    } ; 
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
                    write-verbose "New-ModuleManifest w`n$(($NewCreateExoPSSessionManifest|out-string).trim())"  ;
                    New-ModuleManifest @NewCreateExoPSSessionManifest ;
                    # suppress VerbosePreference:Continue, if set, during mod loads (VERY NOISEY)
                    if($VerbosePreference = "Continue"){
                        $VerbosePrefPrior = $VerbosePreference ;
                        $VerbosePreference = "SilentlyContinue" ;
                        $verbose = ($VerbosePreference -eq "Continue") ;
                    } ; 
                    Import-Module "$($ExoPSModule.PSParentPath)\CreateExoPSSession.psm1" -Global -ErrorAction:Stop ;
                    # reenable VerbosePreference:Continue, if set, during mod loads 
                    if($VerbosePrefPrior -eq "Continue"){
                        $VerbosePreference = $VerbosePrefPrior ;
                        $verbose = ($VerbosePreference -eq "Continue") ;
                    } ; 
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
                    write-verbose "New-ExoPSSession w`n$(($ExoPSSession|out-string).trim())" ;
                    $ExoPSSession = New-ExoPSSession @ExoPSSession -ErrorAction:Stop ;
                    write-verbose "Import-PSSession w`n$(($ImportPSSessionProps|out-string).trim())" ;
                    # suppress VerbosePreference:Continue, if set, during mod loads (VERY NOISEY)
                    if($VerbosePreference = "Continue"){
                        $VerbosePrefPrior = $VerbosePreference ;
                        $VerbosePreference = "SilentlyContinue" ;
                        $verbose = ($VerbosePreference -eq "Continue") ;
                    } ; 
                    Import-Module (Import-PSSession $ExoPSSession @ImportPSSessionProps) -Prefix $CommandPrefix -Global -DisableNameChecking -ErrorAction:Stop ;
                    # reenable VerbosePreference:Continue, if set, during mod loads 
                    if($VerbosePrefPrior -eq "Continue"){
                        $VerbosePreference = $VerbosePrefPrior ;
                        $verbose = ($VerbosePreference -eq "Continue") ;
                    } ; 
                    UpdateImplicitRemotingHandler ;
                    Add-PSTitleBar $sTitleBarTag ;
                } catch {
                    Write-Warning -Message "Failed to connect to EXO via the imported EXO PS module.`n`nError message:" ;
                    throw $_ ;
                } ;
                #>

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
                    Write-Host "Connecting to Exchange Online ($($credential.username.split('@')[1])) via Proxy"  ;
                } Else {
                    Write-Host "Connecting to Exchange Online ($($credential.username.split('@')[1]))"  ;
                } ;
                write-verbose "`n$((get-date).ToString('HH:mm:ss')):New-PSSession w`n$(($EXOsplat|out-string).trim())" ;
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
                write-verbose "`n$((get-date).ToString('HH:mm:ss')):Import-PSSession w`n$(($pltPSS|out-string).trim())" ;
                Try {
                    # Verbose:Continue is VERY noisey for module loads. Bracketed suppress:
                    # suppress VerbosePreference:Continue, if set, during mod loads (VERY NOISEY)
                    if($VerbosePreference = "Continue"){
                        $VerbosePrefPrior = $VerbosePreference ;
                        $VerbosePreference = "SilentlyContinue" ;
                        $verbose = ($VerbosePreference -eq "Continue") ;
                    } ; 
                    $Global:EOLModule = Import-Module (Import-PSSession @pltPSS) -Global -Prefix $CommandPrefix -PassThru -DisableNameChecking   ;
                    # reenable VerbosePreference:Continue, if set, during mod loads 
                    if($VerbosePrefPrior -eq "Continue"){
                        $VerbosePreference = $VerbosePrefPrior ;
                        $verbose = ($VerbosePreference -eq "Continue") ;
                    } ; 
                    Add-PSTitleBar $sTitleBarTag ;
                } catch [System.ArgumentException] {
                    <# 8:45 AM 7/29/2020 VEN tenant now throwing error:
                        WARNING: Tried but failed to import the EXO PS module.
                        Error message:
                        Import-PSSession : Data returned by the remote Get-FormatData command is not in the expected format.
                        At C:\Program Files\WindowsPowerShell\Modules\verb-exo\1.0.14\verb-EXO.psm1:370 char:52
                        + ...   $Global:EOLModule = Import-Module (Import-PSSession @pltPSS) -Globa ...
                        +                                          ~~~~~~~~~~~~~~~~~~~~~~~~
                            + CategoryInfo          : InvalidResult: (:) [Import-PSSession], ArgumentException
                            + FullyQualifiedErrorId : ErrorMalformedDataFromRemoteCommand,Microsoft.PowerShell.Commands.ImportPSSessionCommand
                    
                    EXO bug here:https://answers.microsoft.com/en-us/msoffice/forum/all/cannot-connect-to-exchange-online-via-powershell/25ca1cc2-e23a-470e-9c73-e6c56c4fbb46?page=7
                    Workaround 1) Use EXO V2 module - but it breaks historical use of -suffix 'exo'
                    2) use ?SerializationLevel=Full with the ConnectionURI: -ConnectionUri "https://outlook.office365.com/powershell-liveid?SerializationLevel=Full"
                    #>
                    $EXOsplat.ConnectionUri = 'https://outlook.office365.com/powershell-liveid?SerializationLevel=Full' ;
                    write-warning -verbose:$true "$((get-date).ToString('HH:mm:ss')):'Get-FormatData command is not in the expected format' EXO bug: Retrying with '&SerializationLevel=Full'ConnectionUri`n(details at https://answers.microsoft.com/en-us/msoffice/forum/all/cannot-connect-to-exchange-online-via-powershell/)" ;
                    write-verbose -verbose:$true "`n$((get-date).ToString('HH:mm:ss')):New-PSSession w`n$(($EXOsplat|out-string).trim())" ;
                    TRY{
                        $global:EOLSession | Remove-PSSession; ; 
                        $global:EOLSession = New-PSSession @EXOsplat ;
                    } CATCH {
                        $ErrTrapd = $_ ; 
                        Write-Warning "$(get-date -format 'HH:mm:ss'): Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                        Exit #STOP(debug)|EXIT(close)|Continue(move on in loop cycle) ; 
                    } ; 
                    $pltPSS = [ordered]@{
                        Session             = $global:EOLSession ;
                        Prefix              = $CommandPrefix ;
                        DisableNameChecking = $true  ;
                        AllowClobber        = $true ;
                        ErrorAction         = 'Stop' ;
                    } ;
                    write-verbose -verbose:$true "`n$((get-date).ToString('HH:mm:ss')):Import-PSSession w`n$(($pltPSS|out-string).trim())" ;
                    TRY{
                        $Global:EOLModule = Import-Module (Import-PSSession @pltPSS) -Global -Prefix $CommandPrefix -PassThru -DisableNameChecking   ;
                    } CATCH {
                        $ErrTrapd = $_ ; 
                        Write-Warning "$(get-date -format 'HH:mm:ss'): Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                        Exit #STOP(debug)|EXIT(close)|Continue(move on in loop cycle) ; 
                    } ; 
                    # reenable VerbosePreference:Continue, if set, during mod loads 
                    if($VerbosePrefPrior -eq "Continue"){
                        $VerbosePreference = $VerbosePrefPrior ;
                        $verbose = ($VerbosePreference -eq "Continue") ;
                    } ; 
                    Add-PSTitleBar $sTitleBarTag ;

                } catch {
                    Write-Warning -Message "Tried but failed to import the EXO PS module.`n`nError message:" ;
                    throw $_ ;
                } ;
            
            } ;

        } ; #  # if-E $bExistingEXOGood
    } ;  # PROC-E
    END {
        if($bExistingEXOGood -eq $false){ 
            if ((Get-exoAcceptedDomain).domainname.contains($Credential.username.split('@')[1].tostring())){
                # validate that the connected EXO is to the $Credential tenant    
                write-verbose "(EXO Authenticated & Functional:$($Credential.username.split('@')[1].tostring()))" ; 
                $bExistingEXOGood = $true ; 
            } else { 
                write-error "(Credential mismatch:disconnecting from existing EXO:$($eEXO.Identity) tenant)" ; 
                Disconnect-exo ; 
                $bExistingEXOGood = $false ; 
            } ;
        } ; 
    }  # END-E 
} ; 
#*------^ Connect-EXO.ps1 ^------