﻿# verb-exo.psm1


  <#
  .SYNOPSIS
  verb-EXO - Powershell Exchange Online generic functions module
  .NOTES
  Version     : 1.0.55.0
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



#*------v check-EXOLegalHold.ps1 v------
Function check-EXOLegalHold {
    <#
    .SYNOPSIS
    check-EXOLegalHold - check passed in EXO mailbox object for Legal Hold status
    .NOTES
    Version     : 1.0.0
    Author      : Todd Kadrie
    Website     :	http://www.toddomation.com
    Twitter     :	@tostka / http://twitter.com/tostka
    CreatedDate : 12:36 PM 11/6/2020
    FileName    :
    License     :
    Copyright   :
    Github      : https://github.com/tostka/verb-exo
    Tags        : Powershell,ExchangeOnline,Exchange,Legal
    REVISIONS   :
    # 11:21 AM 3/31/2021 added TenDom test, after AccDom test ;  added verbose suppress to all import-mods
    * 12:37 PM 11/6/2020 init version 
    .DESCRIPTION
    check-EXOLegalHold - check passed in EXO mailbox object for Legal Hold status
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
    check-EXOLegalHold
    Connect using defaults, and leverage any pre-set $global:credo365TORSID variable
    .EXAMPLE
    check-EXOLegalHold -CommandPrefix exo -credential (Get-Credential -credential s-todd.kadrie@torolab.com)  ;
    Connect an explicit credential, and use 'exolab' as the cmdlet prefix
    .EXAMPLE
    $cred = get-credential -credential $o365_Torolab_SIDUpn ;
    check-EXOLegalHold -credential $cred ;
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
        $TenOrg=get-TenantTag -Credential $Credential ; 
        if($TenOrg -ne 'TOR'){
            # explicitly leave this tenant (default) untagged
            $sTitleBarTag += $TenOrg ;
        } ; 
    } ;  # BEG-E
    PROCESS{

        # if we're using EXOv1-style BasicAuth, clear incompatible existing EXOv2 PSS's
        $exov2Good = Get-PSSession | where-object {$_.ConfigurationName -like "Microsoft.Exchange" -and $_.Name -like "ExchangeOnlineInternalSession*" -and $_.State -like "*Opened*" -AND ($_.Availability -eq 'Available')} ; 
        $exov2Broken = Get-PSSession | where-object {$_.ConfigurationName -like "Microsoft.Exchange" -and $_.Name -eq "ExchangeOnlineInternalSession*" -and $_.State -like "*Broken*"}
        $exov2Closed = Get-PSSession | where-object {$_.ConfigurationName -like "Microsoft.Exchange" -and $_.Name -eq "ExchangeOnlineInternalSession*" -and $_.State -like "*Closed*"}

        if($exov2Good  ){
            write-verbose "EXOv1:Disconnecting conflicting EXOv2 connection" ; 
            Discheck-EXOLegalHold2 ; 
        } ; 
        if ($exov2Broken.count -gt 0){for ($index = 0 ;$index -lt $exov2Broken.count ;$index++){Remove-PSSession -session $exov2Broken[$index]} };
        if ($exov2Closed.count -gt 0){for ($index = 0 ;$index -lt $exov2Closed.count ; $index++){Remove-PSSession -session $exov2Closed[$index] } } ; 
    
        $bExistingEXOGood = $false ; 
        # $existingPSSession = Get-PSSession | where-object { $_.ConfigurationName -like "Microsoft.Exchange" -and $_.Name -match "^(Session|WinRM)\d*" }
        #if( Get-PSSession|Where-Object{($_.ComputerName -match $rgxExoPsHostName) -AND ($_.State -eq 'Opened') -AND ($_.Availability -eq 'Available')}){
        # EXOv1 & v2 both use ComputerName -match $rgxExoPsHostName, need to use the distinctive differentiators instead
        if(Get-PSSession | where-object { $_.ConfigurationName -like "Microsoft.Exchange" -AND $_.Name -match "^(Session|WinRM)\d*" -AND $_.State -eq 'Opened' -AND $_.Availability -eq 'Available' }){
            if( get-command Get-exoAcceptedDomain -ea 0) {
                #if ((Get-exoAcceptedDomain).domainname.contains($Credential.username.split('@')[1].tostring())){
                #-=-=-=-=-=-=-=-=
                #$TenOrg = get-TenantTag -Credential $Credential ;
                if(!(Get-Variable  -name "$($TenOrg)Meta").value.o365_AcceptedDomains){
                    set-Variable  -name "$($TenOrg)Meta" -value ( (Get-Variable  -name "$($TenOrg)Meta").value += @{'o365_AcceptedDomains' = (Get-exoAcceptedDomain).domainname} )
                } ;
                if((Get-Variable  -name "$($TenOrg)Meta").value.o365_AcceptedDomains.contains($Credential.username.split('@')[1].tostring())){
                    # validate that the connected EXO is to the $Credential tenant    
                    $smsg = "(EXO Authenticated & Functional(AccDom):$($Credential.username.split('@')[1].tostring()))" ; 
                    $bExistingEXOGood = $true ; 
                # issue: found fresh bug in cxo: svcacct UPN suffix @tenantname.onmicrosoft.com, but testing against AccepteDomain, it's not in there (tho @toroco.mail.onmicrosoft.comis)
                }elseif((Get-Variable  -name "$($TenOrg)Meta").value.o365_TenantDomain -eq ($Credential.username.split('@')[1].tostring())){
                    $smsg = "(EXO Authenticated & Functional(TenDom):$($Credential.username.split('@')[1].tostring()))" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    $bExistingEXOGood = $true ;
                } else { 
                    write-verbose "(Credential mismatch:disconnecting from existing EXO:$($eEXO.Identity) tenant)" ; 
                    Discheck-EXOLegalHold ; 
                    $bExistingEXOGood = $false ; 
                } ; 
            } else { 
                # capture outlier: shows a session wo the test cmdlet, force reset
                Discheck-EXOLegalHold ; 
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
                
                throw "MFA is not currently supported by the check-EXOLegalHold cmdlet!. Use connect/disconnect/recheck-EXOLegalHold2 instead" ; 
                Break 
                <# 4:24 PM 7/30/2020 HAD TO UNINSTALL THE EXOMFA module, a bundled cmdlet fundementally conflicted with ExchangeOnlineManagement#>

            } else {
                $EXOsplat = @{
                    ConfigurationName = "Microsoft.Exchange" ;
                    ConnectionUri     = "https://ps.outlook.com/powershell/" ;
                    Authentication    = "Basic" ;
                    AllowRedirection  = $true;
                } ;
                $EXOsplat.Add("Credential", $Credential); # just use the passed $Credential vari

                $cMsg = "Connecting to Exchange Online ($($credential.username.split('@')[1]))"; 
                If ($ProxyEnabled) {
                    $EXOsplat.Add("sessionOption", $(New-PsSessionOption -ProxyAccessType IEConfig -ProxyAuthentication basic)) ;
                    $cMsg += " via Proxy"  ;
                } ;
                Write-Host $cMsg ;
                write-verbose "`n$((get-date).ToString('HH:mm:ss')):New-PSSession w`n$(($EXOsplat|out-string).trim())" ;
                Try { $global:EOLSession = New-PSSession @EXOsplat ;
                } catch {
                    Write-Warning -Message "Tried but failed to import the EXO PS module.`n`nError message:" ;
                    throw $_ ;
                } ;
                if ($error.count -ne 0) {
                    if ($error[0].FullyQualifiedErrorId -eq '-2144108477,PSSessionOpenFailed') {
                        write-warning "$((get-date).ToString('HH:mm:ss')):AUTH FAIL BAD PASSWORD? ABORTING TO AVOID LOCKOUT!" ;
                        throw "$((get-date).ToString('HH:mm:ss')):AUTH FAIL BAD PASSWORD? ABORTING TO AVOID LOCKOUT!" ;
                        Break ;
                    } ;
                } ;
                if(!$global:EOLSession){
                    write-warning "$((get-date).ToString('HH:mm:ss')):FAILED TO RETURN PSSESSION!`nAUTH FAIL BAD PASSWORD? ABORTING TO AVOID LOCKOUT!" ;
                    throw "$((get-date).ToString('HH:mm:ss')):AUTH FAIL BAD PASSWORD? ABORTING TO AVOID LOCKOUT!" ;
                    Break ;
                } ; 
                $pltiSess = [ordered]@{Session = $global:EOLSession ; DisableNameChecking = $true  ; AllowClobber = $true ; ErrorAction  = 'Stop' ;} ;
                $pltIMod=@{Global=$true;PassThru=$true;DisableNameChecking=$true ; verbose=$false} ; # force verbose off, suppress spam in console
                if($CommandPrefix){
                    $pltIMod.add('Prefix',$CommandPrefix) ;
                    $pltISess.add('Prefix',$CommandPrefix) ;
                } ;
                write-verbose "`n$((get-date).ToString('HH:mm:ss')):Import-PSSession w`n$(($pltiSess|out-string).trim())" ;
                Try {
                    # Verbose:Continue is VERY noisey for module loads. Bracketed suppress:
                    if($VerbosePreference = "Continue"){
                        $VerbosePrefPrior = $VerbosePreference ;
                        $VerbosePreference = "SilentlyContinue" ;
                        $verbose = ($VerbosePreference -eq "Continue") ;
                    } ; 
                    $Global:EOLModule = Import-Module (Import-PSSession @pltiSess) @pltIMod ;
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
                        + ...   $Global:EOLModule = Import-Module (Import-PSSession @pltiSess) -Globa ...
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
                        Break #STOP(debug)|EXIT(close)|Continue(move on in loop cycle) ; 
                    } ; 
                    $pltiSess = [ordered]@{Session = $global:EOLSession ; DisableNameChecking = $true  ; AllowClobber = $true ; ErrorAction = 'Stop' ;} ;
                    $pltIMod=@{Global=$true;PassThru=$true;DisableNameChecking=$true ; verbose=$false} ; # force verbose off, suppress spam in console
                    if($CommandPrefix){
                        $pltIMod.add('Prefix',$CommandPrefix) ;
                        $pltISess.add('Prefix',$CommandPrefix) ;
                    } ;
                    write-verbose -verbose:$true "`n$((get-date).ToString('HH:mm:ss')):Import-PSSession w`n$(($pltiSess|out-string).trim())" ;
                    TRY{
                        $Global:EOLModule = Import-Module (Import-PSSession @pltiSess) @pltIMod   ;
                    } CATCH {
                        $ErrTrapd = $_ ; 
                        Write-Warning "$(get-date -format 'HH:mm:ss'): Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                        Break #STOP(debug)|EXIT(close)|Continue(move on in loop cycle) ; 
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
            # implement caching of accepteddoms into the XXXMeta, in the session (cut back on queries to EXO on acceptedom)
            <#
            $credDom = ($Credential.username.split("@"))[1] ;
            $Metas=(get-variable *meta|Where-Object{$_.name -match '^\w{3}Meta$'}) ;
            foreach ($Meta in $Metas){
                if( ($credDom -eq $Meta.value.legacyDomain) -OR ($credDom -eq $Meta.value.o365_TenantDomain) -OR ($credDom -eq $Meta.value.o365_OPDomain)){
                    if(!$Meta.value.o365_AcceptedDomains){
                        set-variable -Name $meta.name -Value ((get-variable -name $meta.name).value  += @{'o365_AcceptedDomains' = (Get-exoAcceptedDomain).domainname} )
                    } ; 
                    break ;
                } ;
            } ;
            #>
            # simpler non-looping version of testing for meta value, and adding/caching where absent
            #$TenOrg = get-TenantTag -Credential $Credential ;
            if( get-command Get-exoAcceptedDomain -ea 0) {
                if(!(Get-Variable  -name "$($TenOrg)Meta").value.o365_AcceptedDomains){
                    
                    set-Variable  -name "$($TenOrg)Meta" -value ( (Get-Variable  -name "$($TenOrg)Meta").value += @{'o365_AcceptedDomains' = (Get-exoAcceptedDomain).domainname} )
                } ; 
            } ; 
            #if ((Get-exoAcceptedDomain).domainname.contains($Credential.username.split('@')[1].tostring())){
            # do caching & check cached value, not qry unless unpopulated (first pass in global session)
            #if($Meta.value.o365_AcceptedDomains.contains($Credential.username.split('@')[1].tostring())){
            if((Get-Variable  -name "$($TenOrg)Meta").value.o365_AcceptedDomains.contains($Credential.username.split('@')[1].tostring())){
                # validate that the connected EXO is to the $Credential tenant    
                write-verbose "(EXO Authenticated & Functional:$($Credential.username.split('@')[1].tostring()))" ; 
                $bExistingEXOGood = $true ; 
            # issue: found fresh bug in cxo: svcacct UPN suffix @tenantname.onmicrosoft.com, but testing against AccepteDomain, it's not in there (tho @toroco.mail.onmicrosoft.comis)
            }elseif((Get-Variable  -name "$($TenOrg)Meta").value.o365_TenantDomain -eq ($Credential.username.split('@')[1].tostring())){
                $smsg = "(EXO Authenticated & Functional(TenDom):$($Credential.username.split('@')[1].tostring()))" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                $bExistingEXOGood = $true ;
            } else { 
                write-error "(Credential mismatch:disconnecting from existing EXO:$($eEXO.Identity) tenant)" ; 
                Discheck-EXOLegalHold ; 
                $bExistingEXOGood = $false ; 
            } ;
        } ; 
    }  # END-E 
}

#*------^ check-EXOLegalHold.ps1 ^------

#*------v Connect-ExchangeOnlineTargetedPurge.ps1 v------
function Connect-ExchangeOnlineTargetedPurge {
<#
.SYNOPSIS
Connect-ExchangeOnlineTargetedPurge.ps1 - Tweaked version of the Exchangeonline module:connect-ExchangeOnline(), uses variant RemoveExistingPSSession() - RemoveExistingPSSessionTargeted - to avoid purging CCMW sessions on connect. Intent is to permit *concurrent* EXO & CCMS sessions.
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
* 8:34 AM 11/9/2020 init
.DESCRIPTION
Connect-ExchangeOnlineTargetedPurge.ps1 - Tweaked version of the Exchangeonline module:connect-ExchangeOnline(), uses variant RemoveExistingPSSession() to avoid purging CCMW sessions on connect. Intent is to permit *concurrent* EXO & CCMS sessions.
.PARAMETER ConnectionUri
Connection Uri for the Remote PowerShell endpoint
.PARAMETER AzureADAuthorizationEndpointUri = '',
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
Prefix 
.PARAMETER ShowBanner
Show Banner of Exchange cmdlets Mapping and recent updates
.PARAMETER ShowDebug
Parameter to display Debugging messages [-ShowDebug switch]
.INPUTS
None. Does not accepted piped input.
.OUTPUTS
None. Returns no objects or output.
.EXAMPLE
.EXAMPLE
.LINK
https://github.com/tostka/verb-EXO
.LINK
https://docs.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2
#>
    [CmdletBinding()]
    param(

        # Connection Uri for the Remote PowerShell endpoint
        [string] $ConnectionUri = '',

        # Azure AD Authorization endpoint Uri that can issue the OAuth2 access tokens
        [string] $AzureADAuthorizationEndpointUri = '',

        # Exchange Environment name
        [Microsoft.Exchange.Management.RestApiClient.ExchangeEnvironment] $ExchangeEnvironmentName = 'O365Default',

        # PowerShell session options to be used when opening the Remote PowerShell session
        [System.Management.Automation.Remoting.PSSessionOption] $PSSessionOption = $null,

        # Switch to bypass use of mailbox anchoring hint.
        [switch] $BypassMailboxAnchoring = $false,

        # Delegated Organization Name
        [string] $DelegatedOrganization = '',

        # Prefix 
        [string] $Prefix = '',

        # Show Banner of Exchange cmdlets Mapping and recent updates
        [switch] $ShowBanner = $true
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

            # Switch to collect telemetry on command execution. 
            $EnableErrorReporting = New-Object System.Management.Automation.RuntimeDefinedParameter('EnableErrorReporting', [switch], $attributeCollection)
            $EnableErrorReporting.Value = $false
            
            # Where to store EXO command telemetry data. By default telemetry is stored in the directory "%TEMP%/EXOTelemetry" in the file : EXOCmdletTelemetry-yyyymmdd-hhmmss.csv.
            $LogDirectoryPath = New-Object System.Management.Automation.RuntimeDefinedParameter('LogDirectoryPath', [string], $attributeCollection)
            $LogDirectoryPath.Value = ''

            # Create a new attribute and valiate set against the LogLevel
            $LogLevelAttribute = New-Object System.Management.Automation.ParameterAttribute
            $LogLevelAttribute.Mandatory = $false
            $LogLevelAttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
            $LogLevelAttributeCollection.Add($LogLevelAttribute)
            $LogLevelList = @([Microsoft.Online.CSE.RestApiPowerShellModule.Instrumentation.LogLevel]::Default, [Microsoft.Online.CSE.RestApiPowerShellModule.Instrumentation.LogLevel]::All)
            $ValidateSet = New-Object System.Management.Automation.ValidateSetAttribute($LogLevelList)
            $LogLevel = New-Object System.Management.Automation.RuntimeDefinedParameter('LogLevel', [Microsoft.Online.CSE.RestApiPowerShellModule.Instrumentation.LogLevel], $LogLevelAttributeCollection)
            $LogLevel.Attributes.Add($ValidateSet)

# EXO params start

            # Switch to track perfomance 
            $TrackPerformance = New-Object System.Management.Automation.RuntimeDefinedParameter('TrackPerformance', [bool], $attributeCollection)
            $TrackPerformance.Value = $false

            # Flag to enable or disable showing the number of objects written
            $ShowProgress = New-Object System.Management.Automation.RuntimeDefinedParameter('ShowProgress', [bool], $attributeCollection)
            $ShowProgress.Value = $false

            # Switch to enable/disable Multi-threading in the EXO cmdlets
            $UseMultithreading = New-Object System.Management.Automation.RuntimeDefinedParameter('UseMultithreading', [bool], $attributeCollection)
            $UseMultithreading.Value = $true

            # Pagesize Param
            $PageSize = New-Object System.Management.Automation.RuntimeDefinedParameter('PageSize', [uint32], $attributeCollection)
            $PageSize.Value = 1000

# EXO params end
            $paramDictionary = New-object System.Management.Automation.RuntimeDefinedParameterDictionary
            $paramDictionary.Add('UserPrincipalName', $UserPrincipalName)
            $paramDictionary.Add('Credential', $Credential)
            $paramDictionary.Add('EnableErrorReporting', $EnableErrorReporting)
            $paramDictionary.Add('LogDirectoryPath', $LogDirectoryPath)
            $paramDictionary.Add('LogLevel', $LogLevel)
            $paramDictionary.Add('TrackPerformance', $TrackPerformance)
            $paramDictionary.Add('ShowProgress', $ShowProgress)
            $paramDictionary.Add('UseMultithreading', $UseMultithreading)
            $paramDictionary.Add('PageSize', $PageSize)
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
            Import-Module $RestModulePath -Verbose:$false ;
        } ;

        if(!$ExoPowershellModule){$ExoPowershellModule = "Microsoft.Exchange.Management.ExoPowershellGalleryModule.dll"} ;
        if(!$ExoPowershellModulePath){
            $ExoPowershellModulePath = [System.IO.Path]::Combine($EOMgmtModulePath, $ExoPowershellModule) ;
        } ;
        # full path: C:\Users\kadritss\Documents\WindowsPowerShell\Modules\ExchangeOnlineManagement\1.0.1\Microsoft.Exchange.Management.ExoPowershellGalleryModule.dll
        # Name: Microsoft.Exchange.Management.ExoPowershellGalleryModule
        if(-not(get-module Microsoft.Exchange.Management.ExoPowershellGalleryModule)){
            Import-Module $ExoPowershellModulePath -verbose:$false ;
        } ; 
    } 
    process {

        # Validate parameters
        if($ConnectionUri -eq 'False'){$ConnectionUri = ''}
        if (($ConnectionUri -ne '') -and (-not (Test-Uri $ConnectionUri)))
        {
            throw "Invalid ConnectionUri parameter '$ConnectionUri'"
        }
        if (($AzureADAuthorizationEndpointUri -ne '') -and (-not (Test-Uri $AzureADAuthorizationEndpointUri)))
        {
            throw "Invalid AzureADAuthorizationEndpointUri parameter '$AzureADAuthorizationEndpointUri'"
        }
        if (($Prefix -ne '') -and ($Prefix -eq 'EXO'))
        {
            throw "Prefix 'EXO' is a reserved Prefix, please use a different prefix."
        }

        if ($ShowBanner -eq $true)
        {
            Print-Details;
        }

        if (($ConnectionUri -ne '') -and ($AzureADAuthorizationEndpointUri -eq ''))
        {
            Write-Host -ForegroundColor Green "Using ConnectionUri:'$ConnectionUri', in the environment:'$ExchangeEnvironmentName'."
        }
        if (($AzureADAuthorizationEndpointUri -ne '') -and ($ConnectionUri -eq ''))
        {
            Write-Host -ForegroundColor Green "Using AzureADAuthorizationEndpointUri:'$AzureADAuthorizationEndpointUri', in the environment:'$ExchangeEnvironmentName'."
        }

        # Keep track of error count at beginning.
        $errorCountAtStart = $global:Error.Count;
        $global:_EXO_TelemetryFilePath = $null;

        try
        {
            # Cleanup old exchange online PSSessions
            #RemoveExistingPSSession
            RemoveExistingPSSessionTargeted

            $ExoPowershellModule = "Microsoft.Exchange.Management.ExoPowershellGalleryModule.dll";
            #$ModulePath = [System.IO.Path]::Combine($PSScriptRoot, $ExoPowershellModule);
            $ModulePath = [System.IO.Path]::Combine($EOMgmtModulePath, $ExoPowershellModule);

            $global:_EXO_ExchangeEnvironmentName = $ExchangeEnvironmentName;
            $global:_EXO_ConnectionUri = $ConnectionUri;
            $global:_EXO_AzureADAuthorizationEndpointUri = $AzureADAuthorizationEndpointUri;
            $global:_EXO_PSSessionOption = $PSSessionOption;
            $global:_EXO_BypassMailboxAnchoring = $BypassMailboxAnchoring;
            $global:_EXO_DelegatedOrganization = $DelegatedOrganization;
            $global:_EXO_Prefix = $Prefix;

            if ($isCloudShell -eq $false)
            {
                $global:_EXO_UserPrincipalName = $UserPrincipalName.Value;
                $global:_EXO_Credential = $Credential.Value;
                $global:_EXO_EnableErrorReporting = $EnableErrorReporting.Value;
            }
            else
            {
                $global:_EXO_Device = $Device.Value;
            }

            Import-Module $ModulePath -Verbose:$false ;

            $global:_EXO_ModulePath = $ModulePath;

            if ($isCloudShell -eq $false)
            {
                $PSSession = New-ExoPSSession -ExchangeEnvironmentName $ExchangeEnvironmentName -ConnectionUri $ConnectionUri -AzureADAuthorizationEndpointUri $AzureADAuthorizationEndpointUri -UserPrincipalName $UserPrincipalName.Value -PSSessionOption $PSSessionOption -Credential $Credential.Value -BypassMailboxAnchoring:$BypassMailboxAnchoring -DelegatedOrg $DelegatedOrganization
            }
            else
            {
                $PSSession = New-ExoPSSession -ExchangeEnvironmentName $ExchangeEnvironmentName -ConnectionUri $ConnectionUri -AzureADAuthorizationEndpointUri $AzureADAuthorizationEndpointUri -PSSessionOption $PSSessionOption -BypassMailboxAnchoring:$BypassMailboxAnchoring -Device:$Device.Value -DelegatedOrg $DelegatedOrganization
            }

            if ($PSSession -ne $null)
            {
                $PSSessionModuleInfo = Import-PSSession $PSSession -AllowClobber -DisableNameChecking

                # Import the above module globally. This is needed as with using psm1 files, 
                # any module which is dynamically loaded in the nested module does not reflect globally.
                Import-Module $PSSessionModuleInfo.Path -Global -DisableNameChecking -Prefix $Prefix -Verbose:$false ;

                UpdateImplicitRemotingHandler

                # Import the REST module
                $RestPowershellModule = "Microsoft.Exchange.Management.RestApiClient.dll";
                #$RestModulePath = [System.IO.Path]::Combine($PSScriptRoot, $RestPowershellModule);
                $RestModulePath = [System.IO.Path]::Combine($EOMgmtModulePath, $RestPowershellModule);

                Import-Module $RestModulePath -Cmdlet Set-ExoAppSettings -Verbose:$false ;

                # If we are configured to collect telemetry, add telemetry wrappers. 
                if ($EnableErrorReporting.Value -eq $true)
                {
                    $FilePath = Add-EXOClientTelemetryWrapper -Organization (Get-OrgNameFromUPN -UPN $UserPrincipalName.Value) -PSSessionModuleName $PSSessionModuleInfo.Name -LogDirectoryPath $LogDirectoryPath.Value
                    $global:_EXO_TelemetryFilePath = $FilePath[0]
                    Import-Module $FilePath[1] -DisableNameChecking -Verbose:$false

                    Push-EXOTelemetryRecord -TelemetryFilePath $global:_EXO_TelemetryFilePath -CommandName Connect-ExchangeOnlineTargetedPurge -CommandParams $PSCmdlet.MyInvocation.BoundParameters -OrganizationName  $global:_EXO_ExPSTelemetryOrganization -ScriptName $global:_EXO_ExPSTelemetryScriptName  -ScriptExecutionGuid $global:_EXO_ExPSTelemetryScriptExecutionGuid

                    # Set the AppSettings
                    Set-ExoAppSettings -ShowProgress $ShowProgress.Value -PageSize $PageSize.Value -UseMultithreading $UseMultithreading.Value -TrackPerformance $TrackPerformance.Value -ExchangeEnvironmentName $ExchangeEnvironmentName -ConnectionUri $ConnectionUri -AzureADAuthorizationEndpointUri $AzureADAuthorizationEndpointUri -EnableErrorReporting $true -LogDirectoryPath $LogDirectoryPath.Value -LogLevel $LogLevel.Value
                }
                else 
                {
                    # Set the AppSettings disabling the logging
                    Set-ExoAppSettings -ShowProgress $ShowProgress.Value -PageSize $PageSize.Value -UseMultithreading $UseMultithreading.Value -TrackPerformance $TrackPerformance.Value -ExchangeEnvironmentName $ExchangeEnvironmentName -ConnectionUri $ConnectionUri -AzureADAuthorizationEndpointUri $AzureADAuthorizationEndpointUri -EnableErrorReporting $false
                }
            }
        }
        catch
        {
            # If telemetry is enabled, log errors generated from this cmdlet also. 
            if ($EnableErrorReporting.Value -eq $true)
            {
                $errorCountAtProcessEnd = $global:Error.Count 

                if ($global:_EXO_TelemetryFilePath -eq $null)
                {
                    $global:_EXO_TelemetryFilePath = New-EXOClientTelemetryFilePath -LogDirectoryPath $LogDirectoryPath.Value

                    # Import the REST module
                    $RestPowershellModule = "Microsoft.Exchange.Management.RestApiClient.dll";
                    #$RestModulePath = [System.IO.Path]::Combine($PSScriptRoot, $RestPowershellModule);
                    $RestModulePath = [System.IO.Path]::Combine($EOMgmtModulePath, $RestPowershellModule);
                    Import-Module $RestModulePath -Cmdlet Set-ExoAppSettings -Verbose:$false;

                    # Set the AppSettings
                    Set-ExoAppSettings -ShowProgress $ShowProgress.Value -PageSize $PageSize.Value -UseMultithreading $UseMultithreading.Value -TrackPerformance $TrackPerformance.Value -ExchangeEnvironmentName $ExchangeEnvironmentName -ConnectionUri $ConnectionUri -AzureADAuthorizationEndpointUri $AzureADAuthorizationEndpointUri -EnableErrorReporting $true -LogDirectoryPath $LogDirectoryPath.Value -LogLevel $LogLevel.Value
                }

                # Log errors which are encountered during Connect-ExchangeOnlineTargetedPurge execution. 
                Write-Warning("Writing Connect-ExchangeOnlineTargetedPurge error log to " + $global:_EXO_TelemetryFilePath)
                Push-EXOTelemetryRecord -TelemetryFilePath $global:_EXO_TelemetryFilePath -CommandName Connect-ExchangeOnlineTargetedPurge -CommandParams $PSCmdlet.MyInvocation.BoundParameters -OrganizationName  $global:_EXO_ExPSTelemetryOrganization -ScriptName $global:_EXO_ExPSTelemetryScriptName  -ScriptExecutionGuid $global:_EXO_ExPSTelemetryScriptExecutionGuid -ErrorObject $global:Error -ErrorRecordsToConsider ($errorCountAtProcessEnd - $errorCountAtStart) 
            }

            throw $_
        }
    }
}

#*------^ Connect-ExchangeOnlineTargetedPurge.ps1 ^------

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
    # 2:56 PM 3/31/2021 typo/mispaste fix: had $E10Sess assigning on the import ;  bugfix: @toroco.onmicr...com, isn't in EXO.AccDoms, so added a 2nd test for match to TenDom ; added verbose suppress to all import-mods
    * 11:36 AM 3/5/2021 updated colorcode, subed wv -verbose with just write-verbose, added cred.uname echo
    * 1:15 PM 3/1/2021 added org-level color-coded console
    * 8:30 AM 10/22/2020 ren'd $TentantTag -> $TenOrg, swapped looping meta resolve with 1-liner approach ; added AcceptedDom caching to the middle status test (suppress one more get-exoaccepteddomain call if possible)
    * 3:45 PM 10/8/2020 added AcceptedDomain caching to connect-exo as well
    * 1:18 PM 8/11/2020 fixed typo in *broken *closed varis in use; updated ExoV1 conn filter, to specificly target v1 (old matched v1 & v2) ; trimmed entire rem'd MFA block 
    * 4:52 PM 8/4/2020 fixed regex for id'ing legacy pss's
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
    # 10:35 AM 6/20/2019 added $pltiSess splat dump to the import-pssession cmd block; hard-typed the $Credential [System.Management.Automation.PSCredential]
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
            write-host -foregroundcolor white  "(asserting Prefix:$($CommandPrefix)" ;
        } ;

        $sTitleBarTag = "EXO" ;
        $TenOrg=get-TenantTag -Credential $Credential ; 
        if($TenOrg -ne 'TOR'){
            # explicitly leave this tenant (default) untagged
            $sTitleBarTag += $TenOrg ;
        } ; 
    } ;  # BEG-E
    PROCESS{

        # if we're using EXOv1-style BasicAuth, clear incompatible existing EXOv2 PSS's
        $exov2Good = Get-PSSession | where-object {
            $_.ConfigurationName -like "Microsoft.Exchange" -AND $_.Name -like "ExchangeOnlineInternalSession*" -AND (
            $_.State -like "*Opened*") -AND ($_.Availability -eq 'Available')} ; 
        $exov2Broken = Get-PSSession | where-object {$_.ConfigurationName -like "Microsoft.Exchange" -AND (
            $_.Name -eq "ExchangeOnlineInternalSession*") -AND ($_.State -like "*Broken*")}
        $exov2Closed = Get-PSSession | where-object {$_.ConfigurationName -like "Microsoft.Exchange" -AND (
            $_.Name -eq "ExchangeOnlineInternalSession*") -AND ($_.State -like "*Closed*")}

        if($exov2Good  ){
            write-verbose "EXOv1:Disconnecting conflicting EXOv2 connection" ; 
            DisConnect-EXO2 ; 
        } ; 
        if ($exov2Broken.count -gt 0){for ($index = 0 ;$index -lt $exov2Broken.count ;$index++){Remove-PSSession -session $exov2Broken[$index]} };
        if ($exov2Closed.count -gt 0){for ($index = 0 ;$index -lt $exov2Closed.count ; $index++){Remove-PSSession -session $exov2Closed[$index] } } ; 
    
        $bExistingEXOGood = $false ; 
        # $existingPSSession = Get-PSSession | where-object { $_.ConfigurationName -like "Microsoft.Exchange" -AND$_.Name -match "^(Session|WinRM)\d*" }
        #if( Get-PSSession|Where-Object{($_.ComputerName -match $rgxExoPsHostName) -AND ($_.State -eq 'Opened') -AND ($_.Availability -eq 'Available')}){
        # EXOv1 & v2 both use ComputerName -match $rgxExoPsHostName, need to use the distinctive differentiators instead
        if(Get-PSSession | where-object { $_.ConfigurationName -like "Microsoft.Exchange" -AND $_.Name -match "^(Session|WinRM)\d*" -AND $_.State -eq 'Opened' -AND $_.Availability -eq 'Available' }){
            if( get-command Get-exoAcceptedDomain -ea 0) {
                #if ((Get-exoAcceptedDomain).domainname.contains($Credential.username.split('@')[1].tostring())){
                #-=-=-=-=-=-=-=-=
                #$TenOrg = get-TenantTag -Credential $Credential ;
                if(!(Get-Variable  -name "$($TenOrg)Meta").value.o365_AcceptedDomains){
                    set-Variable  -name "$($TenOrg)Meta" -value ( (Get-Variable  -name "$($TenOrg)Meta").value += @{'o365_AcceptedDomains' = (Get-exoAcceptedDomain).domainname} )
                } ;
                if((Get-Variable  -name "$($TenOrg)Meta").value.o365_AcceptedDomains.contains($Credential.username.split('@')[1].tostring())){
                    # validate that the connected EXO is to the $Credential tenant    
                    $smsg = "(EXO Authenticated & Functional(AccDom):$($Credential.username.split('@')[1].tostring()))" ; 
                    $bExistingEXOGood = $true ; 
                # issue: found fresh bug in cxo: svcacct UPN suffix @tenantname.onmicrosoft.com, but testing against AccepteDomain, it's not in there (tho @toroco.mail.onmicrosoft.comis)
                }elseif((Get-Variable  -name "$($TenOrg)Meta").value.o365_TenantDomain -eq ($Credential.username.split('@')[1].tostring())){
                    $smsg = "(EXO Authenticated & Functional(TenDom):$($Credential.username.split('@')[1].tostring()))" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
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
                
                throw "MFA is not currently supported by the connect-exo cmdlet!. Use connect/disconnect/reconnect-exo2 instead" ; 
                Break 
                <# 4:24 PM 7/30/2020 HAD TO UNINSTALL THE EXOMFA module, a bundled cmdlet fundementally conflicted with ExchangeOnlineManagement#>

            } else {
                $EXOsplat = @{
                    ConfigurationName = "Microsoft.Exchange" ;
                    ConnectionUri     = "https://ps.outlook.com/powershell/" ;
                    Authentication    = "Basic" ;
                    AllowRedirection  = $true;
                } ;
                if ($Credential) {
                    $EXOsplat.Add("Credential", $Credential); # just use the passed $Credential vari
                    write-verbose "(using cred:$($credential.username))" ; 
                } ;

                $cMsg = "Connecting to Exchange Online ($($credential.username.split('@')[1]))"; 
                If ($ProxyEnabled) {
                    $EXOsplat.Add("sessionOption", $(New-PsSessionOption -ProxyAccessType IEConfig -ProxyAuthentication basic)) ;
                    $cMsg += " via Proxy"  ;
                } ;
                Write-Host $cMsg ;
                write-verbose "`n$((get-date).ToString('HH:mm:ss')):New-PSSession w`n$(($EXOsplat|out-string).trim())" ;
                Try { $global:EOLSession = New-PSSession @EXOsplat ;
                } catch {
                    Write-Warning -Message "Tried but failed to import the EXO PS module.`n`nError message:" ;
                    throw $_ ;
                } ;
                if ($error.count -ne 0) {
                    if ($error[0].FullyQualifiedErrorId -eq '-2144108477,PSSessionOpenFailed') {
                        write-warning "$((get-date).ToString('HH:mm:ss')):AUTH FAIL BAD PASSWORD? ABORTING TO AVOID LOCKOUT!" ;
                        throw "$((get-date).ToString('HH:mm:ss')):AUTH FAIL BAD PASSWORD? ABORTING TO AVOID LOCKOUT!" ;
                        Break ;
                    } ;
                } ;
                if(!$global:EOLSession){
                    write-warning "$((get-date).ToString('HH:mm:ss')):FAILED TO RETURN PSSESSION!`nAUTH FAIL BAD PASSWORD? ABORTING TO AVOID LOCKOUT!" ;
                    throw "$((get-date).ToString('HH:mm:ss')):AUTH FAIL BAD PASSWORD? ABORTING TO AVOID LOCKOUT!" ;
                    Break ;
                } ; 
                $pltiSess = [ordered]@{Session = $global:EOLSession ; DisableNameChecking = $true  ; AllowClobber = $true ; ErrorAction = 'Stop' ;} ;
                $pltIMod=@{Global=$true;PassThru=$true;DisableNameChecking=$true ; verbose=$false} ; # force verbose off, suppress spam in console
                if($CommandPrefix){
                    $pltIMod.add('Prefix',$CommandPrefix) ;
                    $pltISess.add('Prefix',$CommandPrefix) ;
                } ;
                write-verbose "`n$((get-date).ToString('HH:mm:ss')):Import-PSSession w`n$(($pltiSess|out-string).trim())`n$((get-date).ToString('HH:mm:ss')):Import-Module w`n$(($pltIMod|out-string).trim())" ;
                Try {
                    # Verbose:Continue is VERY noisey for module loads. Bracketed suppress:
                    if($VerbosePreference = "Continue"){
                        $VerbosePrefPrior = $VerbosePreference ;
                        $VerbosePreference = "SilentlyContinue" ;
                        $verbose = ($VerbosePreference -eq "Continue") ;
                    } ; 
                    #$Global:EOLModule = Import-Module (Import-PSSession @pltiSess) -Global -Prefix $CommandPrefix -PassThru -DisableNameChecking -Verbose:$false  ;
                    $Global:EOLModule = Import-Module (Import-PSSession @pltiSess) @pltIMod  ;
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
                        + ...   $Global:EOLModule = Import-Module (Import-PSSession @pltiSess) -Globa ...
                        +                                          ~~~~~~~~~~~~~~~~~~~~~~~~
                            + CategoryInfo          : InvalidResult: (:) [Import-PSSession], ArgumentException
                            + FullyQualifiedErrorId : ErrorMalformedDataFromRemoteCommand,Microsoft.PowerShell.Commands.ImportPSSessionCommand
                    
                    EXO bug here:https://answers.microsoft.com/en-us/msoffice/forum/all/cannot-connect-to-exchange-online-via-powershell/25ca1cc2-e23a-470e-9c73-e6c56c4fbb46?page=7
                    Workaround 1) Use EXO V2 module - but it breaks historical use of -suffix 'exo'
                    2) use ?SerializationLevel=Full with the ConnectionURI: -ConnectionUri "https://outlook.office365.com/powershell-liveid?SerializationLevel=Full"
                    #>
                    $EXOsplat.ConnectionUri = 'https://outlook.office365.com/powershell-liveid?SerializationLevel=Full' ;
                    write-warning -verbose:$true "$((get-date).ToString('HH:mm:ss')):'Get-FormatData command is not in the expected format' EXO bug: Retrying with '&SerializationLevel=Full'ConnectionUri`n(details at https://answers.microsoft.com/en-us/msoffice/forum/all/cannot-connect-to-exchange-online-via-powershell/)" ;
                    write-host -foregroundcolor white "`n$((get-date).ToString('HH:mm:ss')):New-PSSession w`n$(($EXOsplat|out-string).trim())" ;
                    TRY{
                        $global:EOLSession | Remove-PSSession; ; 
                        $global:EOLSession = New-PSSession @EXOsplat ;
                    } CATCH {
                        $ErrTrapd = $_ ; 
                        Write-Warning "$(get-date -format 'HH:mm:ss'): Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                        Break #STOP(debug)|EXIT(close)|Continue(move on in loop cycle) ; 
                    } ; 
                    $pltiSess = [ordered]@{Session = $global:EOLSession ; Prefix = $CommandPrefix ; DisableNameChecking = $true  ; AllowClobber = $true ; ErrorAction = 'Stop' ;} ;
                    $pltIMod=@{Global=$true;PassThru=$true;DisableNameChecking=$true ; verbose=$false} ; # force verbose off, suppress spam in console
                    if($CommandPrefix){
                        $pltIMod.add('Prefix',$CommandPrefix) ;
                        $pltISess.add('Prefix',$CommandPrefix) ;
                    } ;
                    write-host -foregroundcolor white "`n$((get-date).ToString('HH:mm:ss')):Import-PSSession w`n$(($pltiSess|out-string).trim())" ;
                    TRY{
                        $Global:EOLModule = Import-Module (Import-PSSession @pltiSess) @pltIMod   ;
                    } CATCH {
                        $ErrTrapd = $_ ; 
                        Write-Warning "$(get-date -format 'HH:mm:ss'): Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                        Break #STOP(debug)|EXIT(close)|Continue(move on in loop cycle) ; 
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
            # implement caching of accepteddoms into the XXXMeta, in the session (cut back on queries to EXO on acceptedom)
            <#
            $credDom = ($Credential.username.split("@"))[1] ;
            $Metas=(get-variable *meta|Where-Object{$_.name -match '^\w{3}Meta$'}) ;
            foreach ($Meta in $Metas){
                if( ($credDom -eq $Meta.value.legacyDomain) -OR ($credDom -eq $Meta.value.o365_TenantDomain) -OR ($credDom -eq $Meta.value.o365_OPDomain)){
                    if(!$Meta.value.o365_AcceptedDomains){
                        set-variable -Name $meta.name -Value ((get-variable -name $meta.name).value  += @{'o365_AcceptedDomains' = (Get-exoAcceptedDomain).domainname} )
                    } ; 
                    break ;
                } ;
            } ;
            #>
            # simpler non-looping version of testing for meta value, and adding/caching where absent
            #$TenOrg = get-TenantTag -Credential $Credential ;
            if( get-command Get-exoAcceptedDomain -ea 0) {
                if(!(Get-Variable  -name "$($TenOrg)Meta").value.o365_AcceptedDomains){
                    
                    set-Variable  -name "$($TenOrg)Meta" -value ( (Get-Variable  -name "$($TenOrg)Meta").value += @{'o365_AcceptedDomains' = (Get-exoAcceptedDomain).domainname} )
                } ; 
            } ; 
            #if ((Get-exoAcceptedDomain).domainname.contains($Credential.username.split('@')[1].tostring())){
            # do caching & check cached value, not qry unless unpopulated (first pass in global session)
            #if($Meta.value.o365_AcceptedDomains.contains($Credential.username.split('@')[1].tostring())){
            if((Get-Variable  -name "$($TenOrg)Meta").value.o365_AcceptedDomains.contains($Credential.username.split('@')[1].tostring())){
                # validate that the connected EXO is to the $Credential tenant    
                $smsg = "(EXO Authenticated & Functional(AccDom):$($Credential.username.split('@')[1].tostring()))" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                $bExistingEXOGood = $true ; 
            # issue: found fresh bug in cxo: svcacct UPN suffix @tenantname.onmicrosoft.com, but testing against AccepteDomain, it's not in there (tho @toroco.mail.onmicrosoft.comis)
            }elseif((Get-Variable  -name "$($TenOrg)Meta").value.o365_TenantDomain -eq ($Credential.username.split('@')[1].tostring())){
                $smsg = "(EXO Authenticated & Functional(TenDom):$($Credential.username.split('@')[1].tostring()))" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                $bExistingEXOGood = $true ; 
            } else { 
                write-error "(Credential mismatch:disconnecting from existing EXO:$($eEXO.Identity) tenant)" ; 
                Disconnect-exo ; 
                $bExistingEXOGood = $false ; 
                # splice in console color scheming
                <# borked by psreadline v1/v2 breaking changes
                if(($PSFgColor = (Get-Variable  -name "$($TenOrg)Meta").value.PSFgColor) -AND ($PSBgColor = (Get-Variable  -name "$($TenOrg)Meta").value.PSBgColor)){
                    write-verbose "(setting console colors:$($TenOrg)Meta.PSFgColor:$($PSFgColor),PSBgColor:$($PSBgColor))" ; 
                    $Host.UI.RawUI.BackgroundColor = $PSBgColor
                    $Host.UI.RawUI.ForegroundColor = $PSFgColor ; 
                } ;
                #>
            } ;
        } ; 
    }  # END-E 
}

#*------^ Connect-EXO.ps1 ^------

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
    Github      : https://github.com/tostka
    Tags        : Powershell,ExchangeOnline,Exchange,RemotePowershell,Connection,MFA
    AddedCredit : ExactMike Perficient
    AddedWebsite:	https://docs.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2?view=exchange-ps
    AddedTwitter:
    AddedCredit2 : Jeremy Bradshaw
    AddedWebsite2:	https://github.com/JeremyTBradshaw
    AddedTwitter2:
    REVISIONS   :
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
    Credential to use for this connection [-credential 's-todd.kadrie@toro.com']
    .PARAMETER UserPrincipalName
    User Principal Name or email address of the user
    .PARAMETER
    ConnectionUri
    Connection Uri for the Remote PowerShell endpoint [-ConnectionUri 'https://outlook.office365.com/powershell-liveid/']
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
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    None. Returns no objects or output.
    .EXAMPLE
    Connect-EXO2 -cred $credO365TORSID ;
    Connect using defaults, and leverage any pre-set $global:credo365TORSID variable
    .EXAMPLE
    Connect-EXO2 -Prefix exo -credential (Get-Credential -credential s-todd.kadrie@torolab.com)  ;
    Connect an explicit credential, and use 'exolab' as the cmdlet prefix
    .EXAMPLE
    $cred = get-credential -credential $o365_Torolab_SIDUpn ;
    Connect-EXO2 -credential $cred ;
    Pass in a prefab credential object (useful for auto-shifting to MFA - the function will autoresolve MFA reqs based on the cred domain)
    .LINK
    #>
    [CmdletBinding()]
    [Alias('cxo2')]
    Param(
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
        if (!$rgxExoPsHostName) { $rgxExoPsHostName = "^(ps\.outlook\.com|outlook\.office365\.com)$" } ;

        # validate params
        if($ConnectionUri -and $AzureADAuthorizationEndpointUri){
            throw "BOTH -Connectionuri & -AzureADAuthorizationEndpointUri specified, use ONE or the OTHER!";
        }

        # disable prefix spec, unless actually blanked (e.g. centrally spec'd in profile).
        if (!$Prefix) {
            $Prefix = 'xo' ; # 4:31 PM 7/29/2020 MS has RESERVED use of the 'exo' prefix [facepalm]
            write-verbose -verbose:$true  "(asserting Prefix:$($Prefix)" ;
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

        $sTitleBarTag = "EXO2" ;
        $TenOrg = get-TenantTag -Credential $Credential ;
        if ($TenOrg -ne 'TOR') {
            # explicitly leave this tenant (default) untagged
            $sTitleBarTag += $TenOrg ;
        } ;

        $MFA = get-TenantMFARequirement -Credential $Credential ;

        # admin/SID module auto-install code (myBoxes UID split-perm CU, all else t AllUsers)
        $modname = 'ExchangeOnlineManagement' ;
        $minvers = '1.0.1' ;
        Try { Get-Module -name $modname -listavailable -ErrorAction Stop | out-null } Catch {
            $pltInMod = [ordered]@{Name = $modname ; verbose=$false ;} ;
            if ( $env:COMPUTERNAME -match $rgxMyBoxUID ) { $pltInMod.add('scope', 'CurrentUser') } else { $pltInMod.add('scope', 'AllUsers') } ;
            write-host -foregroundcolor YELLOW "$((get-date).ToString('HH:mm:ss')):Install-Module w scope:$($pltInMod.scope)`n$(($pltInMod|out-string).trim())" ;
            Install-Module @pltIMod ;
        } ; # IsInstalled
        $pltIMod = @{Name = $modname ; ErrorAction = 'Stop' ; verbose=$false} ;
        if ($minvers) { $pltIMod.add('MinimumVersion', $minvers) } ;
        Try { Get-Module $modname -ErrorAction Stop | out-null } Catch {
            write-verbose "Import-Module w`n$(($pltIMod|out-string).trim())" ;
            Import-Module @pltIMod ;
        } ; # IsImported

        # .dll etc loads, from connect-exchangeonline: (should be installed with the above)
        if (-not($ExchangeOnlineMgmtPath)) {
            $EOMgmtModulePath = split-path (get-module ExchangeOnlineManagement -list).Path ;
        } ;
        if (!$RestModule) { $RestModule = "Microsoft.Exchange.Management.RestApiClient.dll" } ;
        # stock uses $PSScriptRoot, which will be the verb-EXO path, not the EXOMgmt module have to dyn locate it
        if (!$RestModulePath) {
            $RestModulePath = join-path -path $EOMgmtModulePath -childpath $RestModule  ;
        } ;
        # paths to proper Module path: Name lists as: Microsoft.Exchange.Management.RestApiClient
        if (-not(get-module $RestModule.replace('.dll',''))) {
            Import-Module $RestModulePath -verbose:$false ;
        } ;
        if (!$ExoPowershellGalleryModule) { $ExoPowershellGalleryModule = "Microsoft.Exchange.Management.ExoPowershellGalleryModule.dll" } ;
        if (!$ExoPowershellGalleryModulePath) {
            $ExoPowershellGalleryModulePath = join-path -path $EOMgmtModulePath -childpath $ExoPowershellGalleryModule ;
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
                write-host -foregroundcolor gray "(closing $($existingPSSession.count) legacy EXO sessions...)" ;
                for ($index = 0; $index -lt $existingPSSession.count; $index++) {
                    $session = $existingPSSession[$index] ;
                    Remove-PSSession -session $session ;
                    Write-Host "Removed the PSSession $($session.Name) connected to $($session.ComputerName)" ;
                } ;
                $bExistingEXOGood = $false ;
            } ;
        } ;
        #if ( Get-PSSession | where-object { $_.ConfigurationName -like "Microsoft.Exchange" -and $_.Name -like "ExchangeOnlineInternalSession*" -and $_.State -like "*Opened*" -AND ($_.Availability -eq 'Available') } ) {
        # update to *not* tamper with CCMS connects
        if (!$rgxExoPsHostName) { $rgxExoPsHostName = "^(ps\.outlook\.com|outlook\.office365\.com)$" } ;
        if ( Get-PSSession | where-object { $_.ConfigurationName -like "Microsoft.Exchange" -and $_.Name -like "ExchangeOnlineInternalSession*" -and $_.State -like "*Opened*" -AND ($_.Availability -eq 'Available') -AND ($_.ComputerName -match $rgxExoPsHostName) } ) {
            if( get-command Get-xoAcceptedDomain -ea 0) {
                 #$TenOrg = get-TenantTag -Credential $Credential ;
                if(!(Get-Variable  -name "$($TenOrg)Meta").value.o365_AcceptedDomains){
                    set-Variable  -name "$($TenOrg)Meta" -value ( (Get-Variable  -name "$($TenOrg)Meta").value += @{'o365_AcceptedDomains' = (Get-xoAcceptedDomain).domainname} )
                } ;
                #if ((Get-xoAcceptedDomain).domainname.contains($Credential.username.split('@')[1].tostring())) {
                if((Get-Variable  -name "$($TenOrg)Meta").value.o365_AcceptedDomains.contains($Credential.username.split('@')[1].tostring())){
                    # validate that the connected EXO is to the $Credential tenant
                    write-verbose "(Existing EXO Authenticated & Functional:$($Credential.username.split('@')[1].tostring()))" ;
                    $bExistingEXOGood = $true ;
                # issue: found fresh bug in cxo: svcacct UPN suffix @tenantname.onmicrosoft.com, but testing against AccepteDomain, it's not in there (tho @toroco.mail.onmicrosoft.comis)
                }elseif((Get-Variable  -name "$($TenOrg)Meta").value.o365_TenantDomain -eq ($Credential.username.split('@')[1].tostring())){
                    $smsg = "(EXO Authenticated & Functional(TenDom):$($Credential.username.split('@')[1].tostring()))" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    $bExistingEXOGood = $true ;
                } else {
                    write-verbose "(Credential mismatch:disconnecting from existing EXO:$($eEXO.Identity) tenant)" ;
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
            if(!(get-module $ExoPowershellGalleryModule.replace('.dll','') )){ Import-Module $ExoPowershellGalleryModulePath -verbose:$false} ;
            $global:_EXO_ModulePath = $ExoPowershellGalleryModulePath;

            <# prior module code
            #Connect-ExchangeOnline -Credential $credO365TORSID -Prefix 'xo' -ShowBanner:$false ;
            # *** LEFT OFF HERE 5:01 PM 7/29/2020 *** not sure if it supports allowclobber, if it's actually wrapping pssession, it sure as shit does!

            $pltCXO = @{
                Prefix     = [string]$Prefix ;
                ShowBanner = [switch]$false ;
            } ;
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

            if ($MFA) {
                # -UserPrincipalName
                #$pltCXO.Add("UserPrincipalName", [string]$Credential.username);
                if ($UserPrincipalName) {
                    $pltNEXOS.Add("UserPrincipalName", [string]$UserPrincipalName);
                    write-verbose "(using cred:$([string]$UserPrincipalName))" ; 
                } elseif ($Credential -AND !$UserPrincipalName){
                    $pltNEXOS.Add("UserPrincipalName", [string]$Credential.username);
                    write-verbose "(using cred:$($credential.username))" ; 
                };
            } else {
                # just use the passed $Credential vari
                #$pltCXO.Add("Credential", [System.Management.Automation.PSCredential]$Credential);
                $pltNEXOS.Add("Credential", [System.Management.Automation.PSCredential]$Credential);
                write-verbose "(using cred:$($credential.username))" ; 
            } ;

            if ($AzureADAuthorizationEndpointUri) { $pltNEXOS.Add("AzureADAuthorizationEndpointUri", [string]$AzureADAuthorizationEndpointUri) } ;
            if ($ConnectionUri) { $pltNEXOS.Add("ConnectionUri", [string]$ConnectionUri) } ;

            #Write-Host "Connecting to EXOv2:($($credential.username.split('@')[1]))"  ;
            Write-Host "Connecting to EXOv2:($($credential.username))"  ;
            #write-verbose "`n$((get-date).ToString('HH:mm:ss')):Connect-ExchangeOnline w`n$(($pltCXO|out-string).trim())" ;
            write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):New-ExoPSSession w`n$(($pltNEXOS|out-string).trim())" ;
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

                    +[kadriTSS]::[PS]:D:\scripts$ $error[0]
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
                write-warning -verbose:$true "$((get-date).ToString('HH:mm:ss')):'Get-FormatData command is not in the expected format' EXO bug: Retrying with '&SerializationLevel=Full'ConnectionUri`n(details at https://answers.microsoft.com/en-us/msoffice/forum/all/cannot-connect-to-exchange-online-via-powershell/)" ;
                write-verbose "`n$((get-date).ToString('HH:mm:ss')):Connect-ExchangeOnline w`n$(($pltCXO|out-string).trim())" ;
                <# when this crashes, it leaves an open PSS matching below that TIES UP YOUR CONN QUOTA!
                Get-PSSession | where-object {$_.ConfigurationName -like "Microsoft.Exchange" -and $_.Name -like "ExchangeOnlineInternalSession*" -and $_.State -like "*Opened*" -AND ($_.Availability -eq 'Available')}
                #>
                TRY {
                    # cleanup the borked attempt left half-functioning
                    #Disconnect-ExchangeOnline -confirm:$false ;
                    #Connect-ExchangeOnline @pltCXO ;
                    write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):New-ExoPSSession w`n$(($pltNEXOS|out-string).trim())" ;
                        $PSSession = New-ExoPSSession @pltNEXOS ;
                        #Add-PSTitleBar $sTitleBarTag ;
                } CATCH {
                    $ErrTrapd = $_ ;
                    Write-Warning "$(get-date -format 'HH:mm:ss'): Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                    Break #STOP(debug)|EXIT(close)|Continue(move on in loop cycle) ;
                } ;
            } CATCH [System.Management.Automation.RuntimeException] {
                # see if we can trap the weird blank ConnnectionURI error
                #$pltCXO.Add('ConnectionUri', [string]'https://outlook.office365.com/powershell-liveid/') ;
                $pltNEXOS.Add('ConnectionUri', [string]'https://outlook.office365.com/powershell-liveid/') ;
                write-warning -verbose:$true "$((get-date).ToString('HH:mm:ss')):'Blank ConnectionUri EXOv2 bug: Retrying with explicit 'ConnectionUri" ;
                write-verbose "`n$((get-date).ToString('HH:mm:ss')):Connect-ExchangeOnline w`n$(($pltCXO|out-string).trim())" ;
                TRY {
                    #Disconnect-ExchangeOnline -confirm:$false ;
                    #Connect-ExchangeOnline @pltCXO ;
                    $PSSession = New-ExoPSSession @pltNEXOS ;
                    #Add-PSTitleBar $sTitleBarTag ;
                } CATCH {
                    $ErrTrapd = $_ ;
                    Write-Warning "$(get-date -format 'HH:mm:ss'): Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                    Break #STOP(debug)|EXIT(close)|Continue(move on in loop cycle) ;
                } ;
            } catch {
                Write-Warning -Message "Tried but failed to connect to EXO V2 PS module.`n`nError message:" ;
                throw $_ ;
            } ;
            if ($error.count -ne 0) {
                if ($error[0].FullyQualifiedErrorId -eq '-2144108477,PSSessionOpenFailed') {
                    write-warning "$((get-date).ToString('HH:mm:ss')):AUTH FAIL BAD PASSWORD? ABORTING TO AVOID LOCKOUT!" ;
                    throw "$((get-date).ToString('HH:mm:ss')):AUTH FAIL BAD PASSWORD? ABORTING TO AVOID LOCKOUT!" ;
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
                Set-ExoAppSettings -ShowProgress $ShowProgress.Value -PageSize $PageSize.Value -UseMultithreading $UseMultithreading.Value -TrackPerformance $TrackPerformance.Value -ExchangeEnvironmentName $ExchangeEnvironmentName -ConnectionUri $ConnectionUri -AzureADAuthorizationEndpointUri $AzureADAuthorizationEndpointUri -EnableErrorReporting $false ;

                Add-PSTitleBar $sTitleBarTag ;
            }
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
                if(!(Get-Variable  -name "$($TenOrg)Meta").value.o365_AcceptedDomains){
                    set-Variable  -name "$($TenOrg)Meta" -value ( (Get-Variable  -name "$($TenOrg)Meta").value += @{'o365_AcceptedDomains' = (Get-xoAcceptedDomain).domainname} )
                } ;
            } ;

            if((Get-Variable  -name "$($TenOrg)Meta").value.o365_AcceptedDomains.contains($Credential.username.split('@')[1].tostring())){
                # validate that the connected EXO is to the $Credential tenant
                write-verbose "(EXO Authenticated & Functional:$($Credential.username.split('@')[1].tostring())),($($Credential.username))" ;
                $bExistingEXOGood = $true ;
            # issue: found fresh bug in cxo: svcacct UPN suffix @tenantname.onmicrosoft.com, but testing against AccepteDomain, it's not in there (tho @toroco.mail.onmicrosoft.comis)
            }elseif((Get-Variable  -name "$($TenOrg)Meta").value.o365_TenantDomain -eq ($Credential.username.split('@')[1].tostring())){
                $smsg = "(EXO Authenticated & Functional(TenDom):$($Credential.username.split('@')[1].tostring()))" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                $bExistingEXOGood = $true ;
            } else {
                write-error "(Credential mismatch:disconnecting from existing EXO:$($eEXO.Identity) tenant)" ;
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

#*------v connect-EXO2old.ps1 v------
Function connect-EXO2old {
    <#
    .SYNOPSIS
    connect-EXO2old - Establish connection to Exchange Online (via EXO V2 graph-api module)
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
    # 8:34 AM 3/31/2021 added verbose suppress to all import-mods
    * 2:01 PM 11/10/2020 swap connect-exo2 to connect-exo2old (uses connect-ExchangeOnline), also ren'd CommandPrefix parm -> Prefix (matches EXOModule spec)
    * 4:41 PM 10/8/2020 implemented AcceptedDomain caching, in connect-EXO2old to match rxo2
    * 1:18 PM 8/11/2020 fixed typo in *broken *closed varis in use; updated ExoV1 conn filter, to specificly target v1 (old matched v1 & v2) ; trimmed entire rem'd MFA block ; added trailing test-EXOToken confirm
    * 12:57 PM 8/4/2020 sorted ExchangeOnlineMgmt mod issues (splatting wo using splat char), if MS hadn't completely rewritten the access, this rewrite wouldn't have been necessary in the 1st place. I'm not looking forward to the org wide rewrites to recode verb-exoNoun -> verb-xoNoun, to accomodate the breaking-change blocking -Prefix 'exo'. ; # 1:04 PM 8/4/2020 cute: now the above error's stopped occuring on the problem tenant. Can't do further testing of the workaround, unless/until it breaks again ; * 2:39 PM 8/4/2020 fixed -match "^(Session|WinRM)\d*" rgx (lacked ^, mismatched EXOv2 conns)
    * 12:20 PM 7/29/2020 rewrite/port from connect-EXO to replace import-pssession with new connect-ExchangeOnline cmdlet (supports MFA natively) - #127 # *** LEFT OFF HERE 5:01 PM 7/29/2020 *** not sure if it supports allowclobber, if it's actually wrapping pssession, it sure as shit does!
    * 11:21 AM 7/28/2020 added Credential -> AcceptedDomains Tenant validation, also testing existing conn, and skipping reconnect unless unhealthy or wrong tenant to match credential
    * 3:24 PM 7/24/2020 updated to support tenant-alignment & sub'd out showdebug for verbose
    * 7:13 AM 7/22/2020 replaced codeblock w get-TenantTag()
    * 5:12 PM 7/21/2020 added ven supp
    * 11:50 AM 5/27/2020 added alias:cxo win func
    * 8:38 AM 4/17/2020 added a new test of $global:EOLSession, to detect initial cred fail (pw chg, outofdate creds, locked out)
    * 8:45 AM 3/3/2020 public cleanup, refactored connect-EXO2old for Meta's
    * 9:52 PM 1/16/2020 cleanup
    * 10:55 AM 12/6/2019 connect-EXO2old:added suffix to TitleBar tag for other tenants, also config'd a central tab vari
    * 9:17 AM 12/4/2019 CONSISTENTLY failing to load properly in lab, on lynms6200d - wont' get-module xxxx -listinstalled, even after load, so I rewrote an exemption diverting into the locally installed $env:userprofile\documents\WindowsPowerShell\Modules\exoMFAModule\ copy.
    * 5:14 PM 11/27/2019 repl $MFA code with get-TenantMFARequirement
    * 1:07 PM 11/25/2019 added tenant-specific alias variants for connect & reconnect
    # 1:26 PM 11/19/2019 added MFA detection fr infastrings .ps1 globals, lifted from Jeremy Bradshaw (https://github.com/JeremyTBradshaw)'s Connect-Exchange()
    # 10:35 AM 6/20/2019 added $pltiSess splat dump to the import-pssession cmd block; hard-typed the $Credential [System.Management.Automation.PSCredential]
    # 8:22 AM 11/20/2017 spliced in retry loop into reconnect-EXO2old as well, to avoid using any state testing in scripts, localize it 1x here.
    # 1:49 PM 11/2/2017 coded around non-profile gaps from not having get-admincred() - added the prompt code in to fake it
    # 12:26 PM 9/11/2017 debugged retry - catch doesn't fire properly on new-Pssession, have to test the $error state, to detect auth fails (assuming the bad pw error# is specific). $error test is currently targeting specific error returned on a bad password. Added retry, for when connection won't hold and fails breaks - need to watch out that bad pw doesn't lock out the acct!
    # 12:50 PM 6/2/2017 expanded pshelp, added Examples, cleaned up deadwood
    * # 12:10 PM 3/15/2017 connect-EXO2old typo, disable prefix auto spec, unless actually blanked. switch ProxyEnabled to non-Mandatory.
    .DESCRIPTION
    connect-EXO2old - Establish PSS to EXO V2 Modern Auth
    .PARAMETER  Prefix
    [verb]-PREFIX[command] PREFIX string for clearly marking cmdlets sourced in this connection [-Prefix tag]
    .PARAMETER  Credential
    Credential to use for this connection [-credential 's-todd.kadrie@toro.com']
    .PARAMETER
    ConnectionUri
    Connection Uri for the Remote PowerShell endpoint [-ConnectionUri 'https://outlook.office365.com/powershell-liveid/']
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    None. Returns no objects or output.
    .EXAMPLE
    connect-EXO2old
    Connect using defaults, and leverage any pre-set $global:credo365TORSID variable
    .EXAMPLE
    connect-EXO2old -Prefix exo -credential (Get-Credential -credential s-todd.kadrie@torolab.com)  ;
    Connect an explicit credential, and use 'exolab' as the cmdlet prefix
    .EXAMPLE
    $cred = get-credential -credential $o365_Torolab_SIDUpn ;
    connect-EXO2old -credential $cred ;
    Pass in a prefab credential object (useful for auto-shifting to MFA - the function will autoresolve MFA reqs based on the cred domain)
    .LINK
    #>
    [CmdletBinding()]
    #[Alias('cxo2')]
    Param(
        [Parameter(HelpMessage = "[verb]-PREFIX[command] PREFIX string for clearly marking cmdlets sourced in this connection [-Prefix tag]")]
        [string]$Prefix = 'xo',
        [Parameter(HelpMessage = "Credential to use for this connection [-credential [credential obj variable]")]
        [System.Management.Automation.PSCredential]$Credential = $global:credo365TORSID,
        [Parameter(HelpMessage = "Connection Uri for the Remote PowerShell endpoint [-ConnectionUri 'https://outlook.office365.com/powershell-liveid/']")]
        [string] $ConnectionUri = '',
        [Parameter(HelpMessage = "Debugging Flag [-showDebug]")]
        [switch] $showDebug
    ) ;
    BEGIN {
        $verbose = ($VerbosePreference -eq "Continue") ;
        if (!$rgxExoPsHostName) { $rgxExoPsHostName = "^(ps\.outlook\.com|outlook\.office365\.com)$" } ;
        $MFA = get-TenantMFARequirement -Credential $Credential ;

        # disable prefix spec, unless actually blanked (e.g. centrally spec'd in profile).
        if (!$Prefix) {
            $Prefix = 'xo' ; # 4:31 PM 7/29/2020 MS has RESERVED use of the 'exo' prefix [facepalm]
            write-verbose -verbose:$true  "(asserting Prefix:$($Prefix)" ;
        } ;

        $sTitleBarTag = "EXO2" ;
        $TenOrg = get-TenantTag -Credential $Credential ;
        if ($TenOrg -ne 'TOR') {
            # explicitly leave this tenant (default) untagged
            $sTitleBarTag += $TenOrg ;
        } ;
    } ; # BEG-E
    PROCESS {
        $bExistingEXOGood = $false ;

        # admin/SID module auto-install code (myBoxes UID split-perm CU, all else t AllUsers)
        $modname = 'ExchangeOnlineManagement' ;
        $minvers = '1.0.1' ; 
        Try {Get-Module -name $modname -listavailable -ErrorAction Stop | out-null } Catch {
            $pltInMod=[ordered]@{Name=$modname} ; 
            if( $env:COMPUTERNAME -match $rgxMyBoxUID ){$pltInMod.add('scope','CurrentUser')} else {$pltInMod.add('scope','AllUsers')} ;
            write-host -foregroundcolor YELLOW "$((get-date).ToString('HH:mm:ss')):Install-Module w scope:$($pltInMod.scope)`n$(($pltInMod|out-string).trim())" ; 
            Install-Module @pltIMod ; 
        } ; # IsInstalled
        $pltIMod = @{Name = $modname ; ErrorAction = 'Stop' ; } ;
        if($minvers){$pltIMod.add('MinimumVersion',$minvers) } ; 
        Try { Get-Module $modname -ErrorAction Stop | out-null } Catch {
            write-verbose "Import-Module w`n$(($pltIMod|out-string).trim())" ; 
            Import-Module @pltIMod ; 
        } ; # IsImported

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
        #>
        #if( Get-PSSession|Where-Object{($_.ComputerName -match $rgxExoPsHostName) -AND ($_.State -eq 'Opened') -AND ($_.Availability -eq 'Available')}){
        # Get-PSSession | fl ConfigurationName,name,state,availability
        if ( $existingPSSession = Get-PSSession | where-object { $_.ConfigurationName -like "Microsoft.Exchange" -and $_.Name -match "^(Session|WinRM)\d*" } ) {
            # ignore state & Avail, close the conflicting legacy conn's
            if ($existingPSSession.count -gt 0) {
                write-host -foregroundcolor gray "(closing $($existingPSSession.count) legacy EXO sessions...)" ;
                for ($index = 0; $index -lt $existingPSSession.count; $index++) {
                    $session = $existingPSSession[$index] ;
                    Remove-PSSession -session $session ;
                    Write-Host "Removed the PSSession $($session.Name) connected to $($session.ComputerName)" ;
                } ;
                $bExistingEXOGood = $false ;
            } ;
        } ;
        if ( Get-PSSession | where-object { $_.ConfigurationName -like "Microsoft.Exchange" -and $_.Name -like "ExchangeOnlineInternalSession*" -and $_.State -like "*Opened*" -AND ($_.Availability -eq 'Available') } ) {
            # swap in non-looping
            if( get-command Get-xoAcceptedDomain -ea 0) {
                 #$TenOrg = get-TenantTag -Credential $Credential ;
                if(!(Get-Variable  -name "$($TenOrg)Meta").value.o365_AcceptedDomains){
                    set-Variable  -name "$($TenOrg)Meta" -value ( (Get-Variable  -name "$($TenOrg)Meta").value += @{'o365_AcceptedDomains' = (Get-xoAcceptedDomain).domainname} )
                } ;
                #if ((Get-xoAcceptedDomain).domainname.contains($Credential.username.split('@')[1].tostring())) {
                if((Get-Variable  -name "$($TenOrg)Meta").value.o365_AcceptedDomains.contains($Credential.username.split('@')[1].tostring())){
                    # validate that the connected EXO is to the $Credential tenant
                    write-verbose "(Existing EXO Authenticated & Functional:$($Credential.username.split('@')[1].tostring()))" ;
                    $bExistingEXOGood = $true ;
                # issue: found fresh bug in cxo: svcacct UPN suffix @tenantname.onmicrosoft.com, but testing against AccepteDomain, it's not in there (tho @toroco.mail.onmicrosoft.comis)
                }elseif((Get-Variable  -name "$($TenOrg)Meta").value.o365_TenantDomain -eq ($Credential.username.split('@')[1].tostring())){
                    $smsg = "(EXO Authenticated & Functional(TenDom):$($Credential.username.split('@')[1].tostring()))" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    $bExistingEXOGood = $true ;
                } else {
                    write-verbose "(Credential mismatch:disconnecting from existing EXO:$($eEXO.Identity) tenant)" ;
                    Disconnect-EXO2old ;
                    $bExistingEXOGood = $false ;
                } ;
            } else {
                # capture outlier: shows a session wo the test cmdlet, force reset
                Disconnect-EXO2old ;
                $bExistingEXOGood = $false ;
            } ;
        } ;

        if ($bExistingEXOGood -eq $false) {

            #Connect-ExchangeOnline -Credential $credO365TORSID -Prefix 'xo' -ShowBanner:$false ;
            # *** LEFT OFF HERE 5:01 PM 7/29/2020 *** not sure if it supports allowclobber, if it's actually wrapping pssession, it sure as shit does!

            $pltCXO = @{
                Prefix     = [string]$Prefix ;
                ShowBanner = [switch]$false ;
            } ;

            if ($MFA) {
                # -UserPrincipalName
                $pltCXO.Add("UserPrincipalName", [string]$Credential.username);
            } else {
                # just use the passed $Credential vari
                $pltCXO.Add("Credential", [System.Management.Automation.PSCredential]$Credential);
            } ;

            #Write-Host "Connecting to EXOv2:($($credential.username.split('@')[1]))"  ;
            Write-Host "Connecting to EXOv2:($($credential.username))"  ;
            write-verbose "`n$((get-date).ToString('HH:mm:ss')):Connect-ExchangeOnline w`n$(($pltCXO|out-string).trim())" ;
            Try {
                #$global:ExoPSSession = New-PSSession @pltCXO ;
                # looks like connect-exchangonline does create a global: $global:_EXO_PreviousModuleName on successful connect 
                # - but haven't spotted it in debugging tho', so have to gcm for 1st cmdlt in the module to confirm connected, and then get-xoacceptedomain, to verify connected to desired tenant
                #$global:EOLSession = New-PSSession @pltCXO ;
                Connect-ExchangeOnline @pltCXO ;
                Add-PSTitleBar $sTitleBarTag ;
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

                    +[kadriTSS]::[PS]:D:\scripts$ $error[0]
                    Import-PSSession : Data returned by the remote Get-FormatData command is not in the expected format.
                    At C:\Program Files\WindowsPowerShell\Modules\ExchangeOnlineManagement\1.0.1\ExchangeOnlineManagement.psm1:454 char:40
                    + ... oduleInfo = Import-PSSession $PSSession -AllowClobber -DisableNameChe ...
                    +                 ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                    + CategoryInfo          : InvalidResult: (:) [Import-PSSession], ArgumentException
                    + FullyQualifiedErrorId : ErrorMalformedDataFromRemoteCommand,Microsoft.PowerShell.Commands.ImportPSSessionCommand

                    Should be trappable, even external function

                    # 1:04 PM 8/4/2020 cute: now the above error's stopped occuring on the problem tenant. Can't do further testing of the workaround, unless/until it breaks again
                #>
                $pltCXO.Add('ConnectionUri', [string]'https://outlook.office365.com/powershell-liveid?SerializationLevel=Full') ;
                write-warning -verbose:$true "$((get-date).ToString('HH:mm:ss')):'Get-FormatData command is not in the expected format' EXO bug: Retrying with '&SerializationLevel=Full'ConnectionUri`n(details at https://answers.microsoft.com/en-us/msoffice/forum/all/cannot-connect-to-exchange-online-via-powershell/)" ;
                write-verbose "`n$((get-date).ToString('HH:mm:ss')):Connect-ExchangeOnline w`n$(($pltCXO|out-string).trim())" ;
                <# when this crashes, it leaves an open PSS matching below that TIES UP YOUR CONN QUOTA!
                Get-PSSession | where-object {$_.ConfigurationName -like "Microsoft.Exchange" -and $_.Name -like "ExchangeOnlineInternalSession*" -and $_.State -like "*Opened*" -AND ($_.Availability -eq 'Available')}
                #>
                TRY {
                    # cleanup the borked attempt left half-functioning
                    Disconnect-ExchangeOnline -confirm:$false ;
                    Connect-ExchangeOnline @pltCXO ;
                    Add-PSTitleBar $sTitleBarTag ;
                } CATCH {
                    $ErrTrapd = $_ ;
                    Write-Warning "$(get-date -format 'HH:mm:ss'): Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                    Break #STOP(debug)|EXIT(close)|Continue(move on in loop cycle) ;
                } ;
            } CATCH [System.Management.Automation.RuntimeException] {
                # see if we can trap the weird blank ConnnectionURI error
                $pltCXO.Add('ConnectionUri', [string]'https://outlook.office365.com/powershell-liveid/') ;
                write-warning -verbose:$true "$((get-date).ToString('HH:mm:ss')):'Blank ConnectionUri EXOv2 bug: Retrying with explicit 'ConnectionUri" ;
                write-verbose "`n$((get-date).ToString('HH:mm:ss')):Connect-ExchangeOnline w`n$(($pltCXO|out-string).trim())" ;
                TRY {
                    #Disconnect-ExchangeOnline -confirm:$false ;
                    Connect-ExchangeOnline @pltCXO ;
                    Add-PSTitleBar $sTitleBarTag ;
                } CATCH {
                    $ErrTrapd = $_ ;
                    Write-Warning "$(get-date -format 'HH:mm:ss'): Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                    Break #STOP(debug)|EXIT(close)|Continue(move on in loop cycle) ;
                } ;
            } catch {
                Write-Warning -Message "Tried but failed to connect to EXO V2 PS module.`n`nError message:" ;
                throw $_ ;
            } ;
            if ($error.count -ne 0) {
                if ($error[0].FullyQualifiedErrorId -eq '-2144108477,PSSessionOpenFailed') {
                    write-warning "$((get-date).ToString('HH:mm:ss')):AUTH FAIL BAD PASSWORD? ABORTING TO AVOID LOCKOUT!" ;
                    throw "$((get-date).ToString('HH:mm:ss')):AUTH FAIL BAD PASSWORD? ABORTING TO AVOID LOCKOUT!" ;
                    Break ;
                } ;
            } ;

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
                if(!(Get-Variable  -name "$($TenOrg)Meta").value.o365_AcceptedDomains){
                    set-Variable  -name "$($TenOrg)Meta" -value ( (Get-Variable  -name "$($TenOrg)Meta").value += @{'o365_AcceptedDomains' = (Get-xoAcceptedDomain).domainname} )
                } ;
            } ; 
            <# old loop code
            $credDom = ($Credential.username.split("@"))[1] ;
            $Metas=(get-variable *meta|Where-Object{$_.name -match '^\w{3}Meta$'}) ;
            foreach ($Meta in $Metas){
                if( ($credDom -eq $Meta.value.legacyDomain) -OR ($credDom -eq $Meta.value.o365_TenantDomain) -OR ($credDom -eq $Meta.value.o365_OPDomain)){
                    if(!$Meta.value.o365_AcceptedDomains){
                        set-variable -Name $meta.name -Value ((get-variable -name $meta.name).value  += @{'o365_AcceptedDomains' = (Get-xoAcceptedDomain).domainname} )
                    } ; 
                    break ;
                } ;
            } ;
            #if ((Get-xoAcceptedDomain).domainname.contains($Credential.username.split('@')[1].tostring())) {
            #if($Meta.value.o365_AcceptedDomains.contains($Credential.username.split('@')[1].tostring())){
            #>
            if((Get-Variable  -name "$($TenOrg)Meta").value.o365_AcceptedDomains.contains($Credential.username.split('@')[1].tostring())){
                # validate that the connected EXO is to the $Credential tenant
                write-verbose "(EXO Authenticated & Functional:$($Credential.username.split('@')[1].tostring()))" ;
                $bExistingEXOGood = $true ;
            # issue: found fresh bug in cxo: svcacct UPN suffix @tenantname.onmicrosoft.com, but testing against AccepteDomain, it's not in there (tho @toroco.mail.onmicrosoft.comis)
            }elseif((Get-Variable  -name "$($TenOrg)Meta").value.o365_TenantDomain -eq ($Credential.username.split('@')[1].tostring())){
                $smsg = "(EXO Authenticated & Functional(TenDom):$($Credential.username.split('@')[1].tostring()))" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                $bExistingEXOGood = $true ;
            } else {
                write-error "(Credential mismatch:disconnecting from existing EXO:$($eEXO.Identity) tenant)" ;
                Disconnect-exo ;
                $bExistingEXOGood = $false ;
            } ;
        } ;
        $bExistingEXOGood | write-output ;
    }  # END-E
}

#*------^ connect-EXO2old.ps1 ^------

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
        # full path: C:\Users\kadritss\Documents\WindowsPowerShell\Modules\ExchangeOnlineManagement\1.0.1\Microsoft.Exchange.Management.ExoPowershellGalleryModule.dll
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
            Import-Module $RestModulePath -verbose:$false ;
        } ;

        if(!$ExoPowershellModule){$ExoPowershellModule = "Microsoft.Exchange.Management.ExoPowershellGalleryModule.dll"} ;
        if(!$ExoPowershellModulePath){
            $ExoPowershellModulePath = [System.IO.Path]::Combine($EOMgmtModulePath, $ExoPowershellModule) ;
        } ;
        # full path: C:\Users\kadritss\Documents\WindowsPowerShell\Modules\ExchangeOnlineManagement\1.0.1\Microsoft.Exchange.Management.ExoPowershellGalleryModule.dll
        # Name: Microsoft.Exchange.Management.ExoPowershellGalleryModule
        if(-not(get-module Microsoft.Exchange.Management.ExoPowershellGalleryModule)){
            Import-Module $ExoPowershellModulePath -verbose:$false ;
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

#*------v cxo2cmw.ps1 v------
function cxo2cmw {
    <#
    .SYNOPSIS
    cxo2CMW - Connect-EXO to specified Tenant
    .DESCRIPTION
    Connect-EXO2 - Establish PSS to EXO V2 Modern Auth PS
    .EXAMPLE
    cxo2cmw
    #>
    Connect-EXO -cred $credO365CMWCSID-Verbose:($VerbosePreference -eq 'Continue') ; 
}

#*------^ cxo2cmw.ps1 ^------

#*------v cxo2tol.ps1 v------
function cxo2TOL {
    <#
    .SYNOPSIS
    cxo2TOL - Connect-EXO to specified Tenant
    .DESCRIPTION
    Connect-EXO2 - Establish PSS to EXO V2 Modern Auth PS
    .EXAMPLE
    cxo2TOL
    #>
    Connect-EXO -cred $credO365TOLSID -Verbose:($VerbosePreference -eq 'Continue') ; 
}

#*------^ cxo2tol.ps1 ^------

#*------v cxo2tor.ps1 v------
function cxo2TOR {
    <#
    .SYNOPSIS
    cxo2TOR - Connect-EXO to specified Tenant
    .DESCRIPTION
    Connect-EXO2 - Establish PSS to EXO V2 Modern Auth PS
    .EXAMPLE
    cxo2TOR
    #>
    Connect-EXO -cred $credO365TORSID -Verbose:($VerbosePreference -eq 'Continue') ; 
}

#*------^ cxo2tor.ps1 ^------

#*------v cxo2ven.ps1 v------
function cxo2VEN {
    <#
    .SYNOPSIS
    cxo2VEN - Connect-EXO to specified Tenant
    .DESCRIPTION
    Connect-EXO2 - Establish PSS to EXO V2 Modern Auth PS
    .EXAMPLE
    cxo2VEN
    #>
    Connect-EXO -cred $credO365VENCSID -Verbose:($VerbosePreference -eq 'Continue') ; 
}

#*------^ cxo2ven.ps1 ^------

#*------v cxocmw.ps1 v------
function cxoCMW {
    <#
    .SYNOPSIS
    cxoCMW - Connect-EXO to specified Tenant
    .DESCRIPTION
    Connect-EXO - Establish PSS to https://ps.outlook.com/powershell/
    .EXAMPLE
    cxoCMW
    #>
    Connect-EXO -cred $credO365CMWCSID -Verbose:($VerbosePreference -eq 'Continue') ; 
}

#*------^ cxocmw.ps1 ^------

#*------v cxotol.ps1 v------
function cxoTOL {
    <#
    .SYNOPSIS
    cxoTOL - Connect-EXO to specified Tenant
    .DESCRIPTION
    Connect-EXO - Establish PSS to https://ps.outlook.com/powershell/
    .EXAMPLE
    cxoTOL
    #>
    Connect-EXO -cred $credO365TOLSID -Verbose:($VerbosePreference -eq 'Continue') ; 
}

#*------^ cxotol.ps1 ^------

#*------v cxotor.ps1 v------
function cxoTOR {
    <#
    .SYNOPSIS
    cxoTOR - Connect-EXO to specified Tenant
    .DESCRIPTION
    Connect-EXO - Establish PSS to https://ps.outlook.com/powershell/
    .EXAMPLE
    cxoTOR
    #>
    Connect-EXO -cred $credO365TORSID -Verbose:($VerbosePreference -eq 'Continue') ; 
}

#*------^ cxotor.ps1 ^------

#*------v cxoVEN.ps1 v------
function cxoVEN {
    <#
    .SYNOPSIS
    cxoVEN - Connect-EXO to specified Tenant
    .DESCRIPTION
    Connect-EXO - Establish PSS to https://ps.outlook.com/powershell/
    .EXAMPLE
    cxoVEN
    #>
    Connect-EXO -cred $credO365VENCSID -Verbose:($VerbosePreference -eq 'Continue') ; 
}

#*------^ cxoVEN.ps1 ^------

#*------v Disconnect-ExchangeOnline.ps1 v------
function Disconnect-ExchangeOnline{
    <#
    .SYNOPSIS
    Disconnect-ExchangeOnline.ps1 - localized verb-EXO vers of non-'$global:' funct from ExchangeOnlineManagement. The globals export fine, these don't and appear to need to be loaded manually
    .NOTES
    Version     : 1.0.0
    Author      : Todd Kadrie
    Website     :	http://www.toddomation.com
    Twitter     :	@tostka / http://twitter.com/tostka
    CreatedDate : 20201109-0833AM
    FileName    : Disconnect-ExchangeOnline.ps1
    License     : [none specified]
    Copyright   : [none specified]
    Github      : https://github.com/tostka/verb-exo
    Tags        : Powershell
    AddedCredit : Microsoft (edited version of published commands in the module)
    AddedWebsite:	https://docs.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2
    REVISIONS
    * 8:34 AM 11/9/2020 init
    .DESCRIPTION
    Disconnect-ExchangeOnline.ps1 - localized verb-EXO vers of non-'$global:' funct from ExchangeOnlineManagement. The globals export fine, these don't and appear to need to be loaded manually
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    None. Returns no objects or output.
    .EXAMPLE
    Disconnect-ExchangeOnline
    Stock call
    .LINK
    https://github.com/tostka/verb-EXO
    .LINK
    https://docs.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact='High')]
    param()

    process {
        if ($PSCmdlet.ShouldProcess(
            "Running this cmdlet clears all active sessions created using Connect-ExchangeOnline or Connect-IPPSSession.",
            "Press(Y/y/A/a) if you want to continue.",
            "Running this cmdlet clears all active sessions created using Connect-ExchangeOnline or Connect-IPPSSession. "))
        {

            # Keep track of error count at beginning.
            $errorCountAtStart = $global:Error.Count;

            try
            {
                # Cleanup current exchange online PSSessions
                #RemoveExistingPSSession
                RemoveExistingPSSessionTargeted

                # Import the module once more to ensure that Test-ActiveToken is present
                Import-Module $global:_EXO_ModulePath -Cmdlet Clear-ActiveToken;

                # Remove any active access token from the cache
                Clear-ActiveToken

                Write-Host "Disconnected successfully !"

                if ($global:_EXO_EnableErrorReporting -eq $true)
                {
                    if ($global:_EXO_TelemetryFilePath -eq $null)
                    {
                        $global:_EXO_TelemetryFilePath = New-EXOClientTelemetryFilePath
                    }

                    Push-EXOTelemetryRecord -TelemetryFilePath $global:_EXO_TelemetryFilePath -CommandName Disconnect-ExchangeOnline -CommandParams $PSCmdlet.MyInvocation.BoundParameters -OrganizationName  $global:_EXO_ExPSTelemetryOrganization -ScriptName $global:_EXO_ExPSTelemetryScriptName  -ScriptExecutionGuid $global:_EXO_ExPSTelemetryScriptExecutionGuid
                }
            }
            catch
            {
                # If telemetry is enabled, log errors generated from this cmdlet also. 
                if ($global:_EXO_EnableErrorReporting -eq $true)
                {
                    $errorCountAtProcessEnd = $global:Error.Count 

                    if ($global:_EXO_TelemetryFilePath -eq $null)
                    {
                        $global:_EXO_TelemetryFilePath = New-EXOClientTelemetryFilePath
                    }

                    # Log errors which are encountered during Disconnect-ExchangeOnline execution. 
                    Write-Warning("Writing Disconnect-ExchangeOnline errors to " + $global:_EXO_TelemetryFilePath)

                    Push-EXOTelemetryRecord -TelemetryFilePath $global:_EXO_TelemetryFilePath -CommandName Disconnect-ExchangeOnline -CommandParams $PSCmdlet.MyInvocation.BoundParameters -OrganizationName  $global:_EXO_ExPSTelemetryOrganization -ScriptName $global:_EXO_ExPSTelemetryScriptName  -ScriptExecutionGuid $global:_EXO_ExPSTelemetryScriptExecutionGuid -ErrorObject $global:Error -ErrorRecordsToConsider ($errorCountAtProcessEnd - $errorCountAtStart) 
                }

                throw $_
            }
        }
    }
}

#*------^ Disconnect-ExchangeOnline.ps1 ^------

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
    * 11:54 AM 3/31/2021 added verbose suppress on remove-module/session commands
    * 1:14 PM 3/1/2021 added color reset
    * 3:24 PM 7/24/2020 updated to support tenant-alignment & sub'd out showdebug for verbose
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
    $verbose = ($VerbosePreference -eq "Continue") ; 
    
    if(!$rgxExoPsHostName){$rgxExoPsHostName="^(ps\.outlook\.com|outlook\.office365\.com)$" } ;
    if($Global:EOLModule){$Global:EOLModule | Remove-Module -Force -Verbose:$false ; } ;
    if($global:EOLSession){$global:EOLSession | Remove-PSSession -Verbose:$false ; } ;
    Get-PSSession |Where-Object{$_.ComputerName -match $rgxExoPsHostName } | Remove-PSSession -Verbose:$false ;
    Disconnect-PssBroken -verbose:$($verbose) ;
    Remove-PSTitlebar 'EXO' ;
    
    [console]::ResetColor()  # reset console colorscheme
    <#
    $existingPSSession = Get-PSSession | Where-Object {$_.ConfigurationName -like "Microsoft.Exchange" -and $_.Name -like "ExchangeOnlineInternalSession*"} ;
    if ($existingPSSession.count -gt 0){
        for ($index = 0; $index -lt $existingPSSession.count; $index++) {
            $session = $existingPSSession[$index] ;
            Remove-PSSession -session $session ;
            Write-Host "Removed the PSSession $($session.Name) connected to $($session.ComputerName)" ;
        } ;
    } ;
    # Clear any left over PS tmp modules - keys off of vari set wi UpdateImplicitRemotingHandler (post import-pssession) 
    if ($global:_EXO_PreviousModuleName -ne $null){
        Remove-Module -Name $global:_EXO_PreviousModuleName -ErrorAction SilentlyContinue ;
        $global:_EXO_PreviousModuleName = $null ;
    } ;
    #>
}

#*------^ Disconnect-EXO.ps1 ^------

#*------v Disconnect-EXO2.ps1 v------
Function Disconnect-EXO2 {
    <#
    .SYNOPSIS
    Disconnect-EXO2 - Remove all the existing exchange online PSSessions
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
    REVISIONS   :
    * 11:55 AM 3/31/2021 suppress verbose on module/session cmdlets
    * 1:14 PM 3/1/2021 added color reset
    * 9:55 AM 7/30/2020 EXO v2 version, adapted from Disconnect-EXO, + some content from RemoveExistingPSSession
    .DESCRIPTION
    Disconnect-EXO2 - Remove all the existing exchange online PSSessions
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    None. Returns no objects or output.
    .EXAMPLE
    Disconnect-EXO2;
    .LINK
    #>
    [CmdletBinding()]
    [Alias('dxo2')]
    Param() 
    $verbose = ($VerbosePreference -eq "Continue") ; 
    <#
    if(!$rgxExoPsHostName){$rgxExoPsHostName="^(ps\.outlook\.com|outlook\.office365\.com)$" } ;
    if($Global:EOLModule){$Global:EOLModule | Remove-Module -Force ; } ;
    if($global:EOLSession){$global:EOLSession | Remove-PSSession ; } ;
    Get-PSSession |Where-Object{$_.ComputerName -match $rgxExoPsHostName } | Remove-PSSession ;
    Disconnect-PssBroken -verbose:$($verbose) ;
    Remove-PSTitlebar 'EXO' ;
    #>
    # confirm module present
    $modname = 'ExchangeOnlineManagement' ; 
    #Try {Get-Module $modname -listavailable -ErrorAction Stop | out-null } Catch {Install-Module $modname -scope CurrentUser ; } ;                 # installed
    Try {Get-Module $modname -ErrorAction Stop | out-null } Catch {Import-Module -Name $modname -MinimumVersion '1.0.1' -ErrorAction Stop -verbose:$false; } ; # imported
    # just alias disconnect-ExchangeOnline, it retires token etc as well as closing PSS, but biggest reason is it's got a confirm, hard-coded, needs a function to override
    
    #Disconnect-ExchangeOnline -confirm:$false ; 
    # just use the updated RemoveExistingEXOPSSession
    RemoveExistingEXOPSSession -Verbose:$false ;
    
    Disconnect-PssBroken -verbose:$false ;
    Remove-PSTitlebar 'EXO' ;
    [console]::ResetColor()  # reset console colorscheme
}

#*------^ Disconnect-EXO2.ps1 ^------

#*------v get-MailboxFolderStats.ps1 v------
function get-MailboxFolderStats {
    <#
    .SYNOPSIS
    get-MailboxFolderStats.ps1 - Perform smart get-mailboxfolderstatistics command, as appropriate to target location -Mailbox, on either Exchange on-premesis or Exchange Online.
    .NOTES
    Version     : 1.0.0.
    Author      : Todd Kadrie
    Website     : http://www.toddomation.com
    Twitter     : @tostka / http://twitter.com/tostka
    CreatedDate : 2021-03-12
    FileName    : get-MailboxFolderStats
    License     : MIT License
    Copyright   : (c) 2020 Todd Kadrie
    Github      : https://github.com/tostka/verb-exo
    Tags        : Powershell,ExchangeOnline,Mailbox,Statistics,Reporting
    REVISIONS
    * 3:28 PM 3/16/2021 added multi-tenant support
    * 1:12 PM 3/15/2021 init work was done 3/12, removed recursive-err generating #Require on the hosting verb-EXO module
    .DESCRIPTION
    get-MailboxFolderStats.ps1 - Perform smart get-mailboxfolderstatistics command, as appropriate to target location -Mailbox, on either Exchange on-premesis or Exchange Online.
    Dependancy on my verb-ex2010 Exchange onprem (and is within verb-exo EXO mod, which adds dependant EXO connection support).
    .PARAMETER TenOrg
    TenantTag values, indicating Tenant to Create DDG WITHIN[-TenOrg 'TOL']    
    .PARAMETER  Mailbox
    Mailbox identifier [samaccountname,name,emailaddr,alias]
    .PARAMETER Ticket
    Ticket # [-Ticket nnnnn]
    .PARAMETER IncludeAge
    Switch to include Oldest/Newest message per folder information[-IncludeAge]
    .PARAMETER IncludeSize
    Switch to include aggregate size of each folder [-IncludeSize]
    .PARAMETER NonEmptyOnly
    Switch to display infor for only non-zero content folders (defaults `$true)[-NonEmptyOnly]
    .INPUTS
    Accepts piped input.
    .OUTPUTS
    Outputs csv & console summary of mailbox folders content
    .EXAMPLE
    get-MailboxFolderStats -Mailbox quotes@domain.com -Ticket 99999 -includeage -verbose ;
    Perform a mailbox stats summary report query, on the specified mailbox, and include specified ticket# in output csv (which is output below .\logs\ dir of current directory at runtime).
    .EXAMPLE
    $report = get-MailboxFolderStats -Mailbox quotes@domain.com -Ticket 99999 -includeage -asobject ;
    Return an object for the summary report, rather than console dump (in addition to csv export)
    .EXAMPLE
    get-MailboxFolderStats -Mailbox quotes@domain.com -Ticket 347298 -includeage -includesize ;
    Perform a mailbox stats, and include size per folder (in KB) in output report
    .LINK
    https://github.com/tostka/verb-exo
    .LINK
    https://github.com/tostka/verb-ex2010
    #>
    #Requires -Version 3
    #Requires -Modules verb-ex2010
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$FALSE,HelpMessage="TenantTag values, indicating Tenant to Create DDG WITHIN[-TenOrg 'TOL']")]
        [ValidateNotNullOrEmpty()]
        [string]$TenOrg = ('TOR'),
        [Parameter(Position=0,Mandatory=$True,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,HelpMessage="Mailbox identifier [samaccountname,name,emailaddr,alias]")]
        [ValidateNotNullOrEmpty()][string]$Mailbox,    
        [Parameter(Mandatory=$false,HelpMessage="Ticket # [-Ticket nnnnn]")]
        #[ValidateLength(5)] # non-mandatory
        [int]$Ticket,
        [Parameter(HelpMessage="Switch to include Oldest/Newest message per folder information[-IncludeAge]")]
        [switch] $IncludeAge,
        [Parameter(HelpMessage="Switch to include aggregate size of each folder [-IncludeSize]")]
        [switch] $IncludeSize,
        [Parameter(HelpMessage="Switch to display info for only non-zero content folders (defaults `$true)[-NonEmptyOnly]")]
        [switch] $NonEmptyOnly=$true,
        [Parameter(HelpMessage="Switch to return raw object rather than formated console report(defaults `$true)[-NonEmptyOnly]")]
        [switch] $asObject
    ) ;
    BEGIN {
        $Verbose=($VerbosePreference -eq 'Continue') ;  
        # Get parameters this function was invoked with
        $PSParameters = New-Object -TypeName PSObject -Property $PSBoundParameters ;  
        $pltGMFS=@{identity= $Mailbox ;} ; 
        $propsFldr = @{Name='Folder';Expression={$_.Identity.tostring()}},@{Name="Items";Expression={$_.ItemsInFolder}} ;
        $rgxSysFldrs = '.*\\(Versions|SubstrateHolds|DiscoveryHolds|Yammer.*|Social\sActivity\sNotifications|Suggested\sContacts|Recipient\sCache|PersonMetadata|Audits|Calendar\sLogging|Purges)$' ; 
        if($IncludeAge){ 
            $pltGMFS.add('IncludeOldestAndNewestItems',$true) ; 
            $propsFldr += @{Name="OldestItem";Expression={get-date $_.OldestItemReceivedDate}},@{Name="NewestItem";Expression={$_.NewestItemReceivedDate}} ; 
        } ;
        if($IncludeSize){ 
            $pltGMFS.add('IncludeAnalysis',$true) ; 
            # w dehydrated, raw parsing is: $mbxstats.TotalItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1MB ;
            $propsFldr += @{Name="SizeMB";Expression={[math]::round($_.FolderSize.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1MB,2)}} ; 
        } ;

        $Retries = 4 ;
        $RetrySleep = 5 ;
        if(!$ThrottleMs){$ThrottleMs = 50 ;}
        $CredRole = 'CSVC' ; # role of svc to be dyn pulled from metaXXX if no -Credential spec'd, 
        if(!$rgxDriveBanChars){$rgxDriveBanChars = '[;~/\\\.:]' ; } ; # ;~/\.:, 

        $UseOP=$false ; 
        if((Get-Variable  -name "$($TenOrg)Meta").value.ExOPAccessFromToro -AND (Get-Variable  -name "$($TenOrg)Meta").value.Ex10Server){
            $UseOP = $true ; 
            $smsg = "$($TenOrg):Meta.ExOPAccessFromToro($((Get-Variable  -name "$($TenOrg)Meta").value.ExOPAccessFromToro)) -AND/OR Meta.Ex10Server($((Get-Variable  -name "$($TenOrg)Meta").value.Ex10Server)),`ENABLING use of OnPrem Ex system this pass." ; 
            if($verbose){ if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        } else { 
            $UseOP = $false ; 
            $smsg = "$($TenOrg):Meta.ExOPAccessFromToro($((Get-Variable  -name "$($TenOrg)Meta").value.ExOPAccessFromToro)) -AND/OR Meta.Ex10Server($((Get-Variable  -name "$($TenOrg)Meta").value.Ex10Server)),`nDISABLING use of OnPrem Ex system this pass." ; 
            if($verbose){ if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        } ; 

        # o365/EXO creds
        $o365Cred=$null ;
        <# Usage: Type defaults to SID, if not spec'd - Note: there must be a *logged in & configured *profile* 
        $o365Cred=get-TenantCredentials -TenOrg $TenOrg -verbose -userrole SID ;
        Returns a credential set for the $TenOrg Hybrid OnPrem Exchange Org
        .EXAMPLE
        $o365Cred=get-TenantCredentials -TenOrg $TenOrg -verbose -userrole CSVC ;
        Returns the CSVC Userrole credential for the $TenOrg Hybrid OnPrem Exchange Org
        .EXAMPLE
        $o365Cred=get-TenantCredentials -TenOrg $TenOrg -verbose -userrole B2BI ;
        Returns the B2BI Userrole credential for the $TenOrg Hybrid OnPrem Exchange Org
        #>
        #if($o365Cred=(get-TenantCredentials -TenOrg $TenOrg -verbose:$($verbose))){
        # force it to use the csvc mapping from $xxxmeta.o365_CSvcUpn, failthrough to SID spec 
        if($o365Cred=(get-TenantCredentials -TenOrg $TenOrg -UserRole 'CSVC','SID' -verbose:$($verbose))){
            # make it script scope, so we don't have to predetect & purge before using new-variable
            New-Variable -Name cred$($tenorg) -scope Script -Value $o365Cred.cred ;
            $smsg = "Resolved $($Tenorg) `$o365cred:$($o365Cred.cred.username) (assigned to `$cred$($tenorg))" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        } else {
            #-=-record a STATUS=-=-=-=-=-=-=
            $statusdelta = ";ERROR";
            $script:PassStatus += $statusdelta ;
            set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta) ;
            #-=-=-=-=-=-=-=-=
            $smsg = "Unable to resolve $($tenorg) `$o365Cred value!"
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            throw "Unable to resolve $($tenorg) `$o365Cred value!`nEXIT!"
            exit ;
        } ;
        <# CALLS ARE IN FORM: (cred$($tenorg))
        $pltRXO = @{Credential = $Credential ; verbose = $($verbose) ; }
        $pltRXO = @{
            Credential = (Get-Variable -name cred$($tenorg) ).value ;
            verbose = $($verbose) ; } ; 
        if ($script:useEXOv2) { reconnect-eXO2 @pltRXO }
        else { reconnect-EXO @pltRXO } ;
        # or with Tenant-specific cred($Tenorg) lookup
        #>

        if($UseOP){
            # do the OP creds too
            $OPCred=$null ;
            # default to the onprem svc acct
            $pltGHOpCred=@{TenOrg=$TenOrg ;userrole='ESVC' ;verbose=$($verbose)} ;
            if($OPCred=(get-HybridOPCredentials @pltGHOpCred).cred){
                # make it script scope, so we don't have to predetect & purge before using new-variable
                New-Variable -Name "cred$($tenorg)OP" -scope Script -Value $OPCred ;
                $smsg = "Resolved $($Tenorg) `$OPCred:$($OPCred.username) (assigned to `$cred$($tenorg)OP)" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } else {
                #-=-record a STATUS=-=-=-=-=-=-=
                $statusdelta = ";ERROR";
                $script:PassStatus += $statusdelta ;
                set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta) ;
                #-=-=-=-=-=-=-=-=
                $smsg = "Unable to resolve get-HybridOPCredentials -TenOrg $($TenOrg) -userrole 'ESVC' value!"
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                throw "Unable to resolve $($tenorg) `$OPCred value!`nEXIT!"
                exit ;
            } ;

            # === Exchange LEMS/REMS detect & connect code

            $smsg= "Using EXOP cred:`$cred$($tenorg)OP:$((Get-Variable -name "cred$($tenorg)OP" ).value.username)" ;  
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
        } ; 

    } ;  # BEGIN-E
    PROCESS {
        $ofile=".\$($ticket)-$($Mailbox)-folder-sizes-$(get-date -format 'yyyyMMdd-HHmmtt').csv" ;
        $error.clear() ;
        TRY {
            if(!(gcm get-recipient -ea 0)){rx10} ;
            $OpRcp=get-recipient $Mailbox ;
            switch ($OpRcp.recipienttype){
                "MailUser" {
                    write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($tmbx) IS AN EXO MBOX" ;
                    
                    if($VerbosePreference = "Continue"){
                        $VerbosePrefPrior = $VerbosePreference ;
                        $VerbosePreference = "SilentlyContinue" ;
                        $verbose = ($VerbosePreference -eq "Continue") ;
                    } ; 
                    disconnect-exo ; # pre-disconnect    
                    $pltRXO = @{
                        Credential = (Get-Variable -name cred$($tenorg) ).value ;
                        verbose = $($verbose) ; }
                    if ($script:useEXOv2) { reconnect-eXO2 @pltRXO }
                    else { reconnect-EXO @pltRXO } ;
                    if($VerbosePrefPrior -eq "Continue"){
                        $VerbosePreference = $VerbosePrefPrior ;
                        $verbose = ($VerbosePreference -eq "Continue") ;
                    } ;

                    set-alias ps1GetMbxFldrStat Get-exoMailboxFolderStatistics ; 
                } ;
                "UserMailbox" {
                    write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($tmbx) IS AN EX2010 MBOX" ;
                    
                    if($VerbosePreference = "Continue"){
                        $VerbosePrefPrior = $VerbosePreference ;
                        $VerbosePreference = "SilentlyContinue" ;
                        $verbose = ($VerbosePreference -eq "Continue") ;
                    } ; 
                    # connect OP
                    $pltRX10 = @{
                        Credential = (Get-Variable -name "cred$($tenorg)OP" ).value ;
                        verbose = $($verbose) ; } ;     
                    if($pltRX10){
                        Connect-Ex2010 @pltRX10 ;
                    } else { connect-Ex2010 ; } ;
                    # reenable VerbosePreference:Continue, if set, during mod loads 
                    if($VerbosePrefPrior -eq "Continue"){
                        $VerbosePreference = $VerbosePrefPrior ;
                        $verbose = ($VerbosePreference -eq "Continue") ;
                    } ;

                    set-alias ps1GetMbxFldrStat Get-MailboxFolderStatistics ; 
                } ;
                default {
                    throw "UNRECOGNIZED ONPREM RECIPIENTTYPE:$($OpRcp.recipienttype)" ; exit ; 
                } ; 
            } ;
            write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$((get-alias ps1GetMbxFldrStat).definition) w`n$(($pltGMFS|out-string).trim())" ; 
            $fldrs = ps1GetMbxFldrStat @pltGMFS ;
            if($NonEmptyOnly){
                write-host -foregroundcolor white "$((get-date).ToString('HH:mm:ss')):(REPORTING NON-ZERO FOLDERS ONLY)" ; $fldrs = $fldrs | ?{$_.ItemsInFolder -gt 0}
            } ; 
            $fldrs | ?{$_.identity -notmatch $rgxSysFldrs } | select $propsFldr | export-csv  -path $ofile -notype ;
            if(!$asObject){
                import-csv $ofile | ft -auto | out-default ; 
            } else { 
                write-verbose "-asObject specified, returning object to pipeline (rather than console dump)" ; 
                import-csv $ofile | write-output ; 
            } ; 
            write-host -foregroundcolor white "$((get-date).ToString('HH:mm:ss')):`n===>`$ofile:$($ofile)`n" ;
        } CATCH {
            Write-Warning "$(get-date -format 'HH:mm:ss'): Failed processing $($_.Exception.ItemName). `nError Message: $($_.Exception.Message)`nError Details: $($_)" ;
            Exit ;
        } ; 
    } ;  # PROC-E
    END {
        remove-alias ps1GetMbxFldrStat ;
    } ; 
    
}

#*------^ get-MailboxFolderStats.ps1 ^------

#*------v get-MsgTrace.ps1 v------
function get-MsgTrace {
    <#
    .SYNOPSIS
    get-MsgTrace.ps1 - Perform smart get-exoMessageTrace/MessageTrackingLog command, as appropriate to target location -Mailbox, on either Exchange on-premesis or Exchange Online.
    .NOTES
    Version     : 1.0.0.
    Author      : Todd Kadrie
    Website     : http://www.toddomation.com
    Twitter     : @tostka / http://twitter.com/tostka
    CreatedDate : 2021-03-12
    FileName    : get-MsgTrace.ps1
    License     : MIT License
    Copyright   : (c) 2020 Todd Kadrie
    Github      : https://github.com/tostka/verb-exo
    Tags        : Powershell,ExchangeOnline,Mailbox,Statistics,Reporting
    REVISIONS
    * 2:23 PM 3/16/2021 added multi-tenant support ; debugged both exOP & exo, added -ReportFail & -ReportRowsLimit params. At this point Exclusive params are only partially configured
    * 1:12 PM 3/15/2021 init work was done 3/12, removed recursive-err generating #Require on the hosting verb-EXO module
    .DESCRIPTION
    get-MsgTrace - Perform smart get-exoMessageTrace/MessageTrackingLog command, as appropriate to target location -Mailbox, on either Exchange on-premesis or Exchange Online.
    Dependancy on my verb-ex2010 Exchange onprem (and is within verb-exo EXO mod, which adds dependant EXO connection support).
    .PARAMETER TenOrg
    TenantTag values, indicating Tenant to Create DDG WITHIN[-TenOrg 'TOL']    
    .PARAMETER Recipients
    Recipient email addresses identifiers (comma-delimited)[-Recipients xxx@domain.com]
    .PARAMETER Sender
    Sender email address identifiers (EXO supports comma-delimited) [-Sender xxx@domain.com]
    .PARAMETER Subject
    "Message Subject string to be matched (post-filtered from broad query)[-Subject 'subject phrase']
    .PARAMETER Logon
    User Logon tag to be applied to output file[-Logon samaccountname]
    .PARAMETER Status
    Transport Status (EventID on-Prem)(RECEIVE|DELIVER|FAIL|SEND|RESOLVE|EXPAND|TRANSFER|DEFER) [-EventID SEND
    .PARAMETER Connectorid
    Connector identifier[-Connectorid SendConnX]
    .PARAMETER Source
    Source keyword to be used for filtering (STOREDRIVER|SMTP|DNS|ROUTING)[-Source SMTP]
    .PARAMETER MessageId
    "Target MessageId for search[-MessageId xxxxxxx]
    .PARAMETER MessageTraceId
    Target MessageId for search[-MessageTraceId xxxxxxx]
    .PARAMETER StartDate
    Start of time span to be searched[-StartDate 1/1/2021]
    .PARAMETER EndDate
    End of time span to be searched[-EndDate 1/7/2021]
    .PARAMETER Ticket
    Ticket # [-Ticket nnnnn]
    .PARAMETER useEXOP
    Switch to specify ONPREM Exch get-MessageTrackingLog trace (defaults `$false == EXO Message Search)[-useEXOP]
    .PARAMETER ReportRowsLimit
    Max number of rows to output to console when a -ReportXXX param is specified (defaults 100)[-ReportRowsLimit]
    .PARAMETER asObject
    Switch to return raw object rather than formated console report(defaults `$true)[-NonEmptyOnly]
    .INPUTS
    Accepts piped input.
    .OUTPUTS
    Outputs csv & console summary of mailbox folders content
    .EXAMPLE
    get-MsgTrace -Sender SENDER@DOMAIN.com -Ticket 99999 -days 7 -verbose ;
    Perform a default EXO trace last 7 days of traffic on specified sender, use specified Ticket number in csv file name, with verbose output
    .EXAMPLE
    $msgs = get-MsgTrace -Sender quotes@bossplow.com -Ticket 347298 -days 7 -asobject -verbose ;
    Above EXO MessageTrace returning an object for further postfiltering.
    .EXAMPLE
    get-msgtrace -sender monitoring@toro.com -useEXOP -ticket 99999 -d 1 -verbose ; 
    Run an ONPREM get-MessageTrackingLog search
    .EXAMPLE 
    $msgs = get-msgtrace -sender monitoring@toro.com -useEXOP -ticket 99999 -start (get-date).addhours(-1) -verbose -ReportFail; 
    Run an ONPREM get-MessageTrackingLog search, with specific -Start time (End will be asserted), with detailed dump of (first 100) EventID 'Fail' items
    .LINK
    https://github.com/tostka/verb-exo
    .LINK
    https://github.com/tostka/verb-ex2010
    #>
    #Requires -Version 3
    #Requires -Modules verb-ex2010
    # VALIDATORS: [ValidateNotNull()][ValidateNotNullOrEmpty()][ValidateLength(24,25)][ValidateLength(5)][ValidatePattern("(lyn|bcc|spb|adl)ms6(4|5)(0|1).(china|global)\.ad\.toro\.com")][ValidateSet("USEA","GBMK","AUSYD")][ValidateScript({Test-Path $_ -PathType 'Container'})][ValidateScript({Test-Path $_})][ValidateRange(21,65)][ValidateCount(1,3)]
    [CmdletBinding(DefaultParameterSetName='SendRec')]
    <# $isplt=@{  ticket="347298" ;  uid="wilinaj";  days=7 ;  Sender="quotes@bossplow.com" ;  Recipients="" ;  MessageSubject="" ;  EventID='' ;  Connectorid="" ;  Source="" ;} ; 
    #>
    Param(
        [Parameter(Mandatory=$FALSE,HelpMessage="TenantTag values, indicating Tenant to Create DDG WITHIN[-TenOrg 'TOL']")]
        [ValidateNotNullOrEmpty()]
        [string]$TenOrg = ('TOR'),
        [Parameter(ParameterSetName='SendRec',HelpMessage="Recipient email addresses identifiers (comma-delimited)[-Recipients xxx@domain.com]")]
        [string]$Recipients,    
        [Parameter(ParameterSetName='SendRec',HelpMessage="Sender email address identifier (EXO supports comma-delimited)")]
        [string]$Sender, 
        [Parameter(HelpMessage="Message Subject string to be matched (post-filtered from broad query)[-Subject 'subject phrase']")]
        [string]$Subject,
        [Parameter(HelpMessage="User Logon tag to be applied to output file[-Logon samaccountname]")]
        [string]$Logon,
        [Parameter(HelpMessage="Transport Status (EventID on-Prem)(RECEIVE|DELIVER|FAIL|SEND|RESOLVE|EXPAND|TRANSFER|DEFER) [-EventID SEND")]
        [ValidateSet("RECEIVE","DELIVER","FAIL","SEND","RESOLVE","EXPAND","TRANSFER","DEFER")]
        [string]$Status,
        [Parameter(HelpMessage="Connector identifier[-Connectorid SendConnX]")]
        [string]$Connectorid,
        [Parameter(HelpMessage="Source keyword to be used for filtering (STOREDRIVER|SMTP|DNS|ROUTING)[-Source SMTP]")]
        [ValidateSet("STOREDRIVER","SMTP","DNS","ROUTING")]
        [string]$Source,
        [Parameter(ParameterSetName='MsgID',HelpMessage="Target MessageId for search[-MessageId xxxxxxx]")]
        [string]$MessageId, 
        [Parameter(ParameterSetName='MsgTrcID',HelpMessage="Target MessageId for search[-MessageTraceId xxxxxxx]")]
        [string]$MessageTraceId,
        [Parameter(HelpMessage="Start of time span to be searched[-StartDate 1/1/2021]")]
        [string]$StartDate,
        [Parameter(HelpMessage="End of time span to be searched[-EndDate 1/7/2021]")]
        [string]$EndDate,
        [Parameter(HelpMessage="Days back to search[-Days 7]")]
        [int]$Days,
        [Parameter(Mandatory=$false,HelpMessage="Ticket # [-Ticket nnnnn]")]
        #[ValidateLength(5)] # non-mandatory
        [int]$Ticket,
        [Parameter(HelpMessage="Switch to specify ONPREM Exch get-MessageTrackingLog trace (defaults `$false == EXO Message Search)[-useEXOP]")]
        [switch] $useEXOP=$false,
        [Parameter(HelpMessage="Switch to return raw object rather than formated console report(defaults `$true)[-NonEmptyOnly]")]
        [switch] $asObject,
        [Parameter(HelpMessage="Switch to return detailed analysis of FAIL items[-ReportFail]")]
        [switch] $ReportFail,
        [Parameter(HelpMessage="Max number of rows to output to console when a -ReportXXX param is specified (defaults 100)[-ReportRowsLimit]")]
        [int]$ReportRowsLimit = 100  
    ) ;
    BEGIN {
        $Verbose=($VerbosePreference -eq 'Continue') ;  
        # Get parameters this function was invoked with
        $PSParameters = New-Object -TypeName PSObject -Property $PSBoundParameters ;
        $propsFldr = @{Name='Folder';Expression={$_.Identity.tostring()}},@{Name="Items";Expression={$_.ItemsInFolder}} ;
        $propsMsgEx10 = 'Timestamp',@{N='TimestampLocal';E={$_.Timestamp.ToLocalTime()}},'Source','EventId','RelatedRecipientAddress','Sender',@{N='Recipients';E={$_.Recipients}},"RecipientCount",@{N='RecipientStatus';E={$_.RecipientStatus}},"MessageSubject","TotalBytes",@{N='Reference';E={$_.Reference}},'MessageLatency','MessageLatencyType','InternalMessageId','MessageId','ReturnPath','ClientIp','ClientHostname','ServerIp','ServerHostname','ConnectorId','SourceContext','MessageInfo',@{N='EventData';E={$_.EventData}} ;
        $propsMsgEXO = @{N='ReceivedLocal';E={$_.Received.ToLocalTime()}},'SenderAddress','RecipientAddress','Subject','Status','ToIP','FromIP','Size','MessageId','MessageTraceId','Index' ;
        
        # pull settings per Tenant fr Meta
        $Meta = gv -name "$($TenOrg)Meta" ; 
        <# pull value fr meta
        if($Meta -is [system.array]){ throw "Unable to resolve unique `$xxxMeta! from `$TenOrg:$($TenOrg)" ; break} ; 
        if(!$Meta.value.DefaultObjectOwner){throw "Unable to resolve $($Meta.Name).value.DefaultObjectOwner from `$TenOrg:$($TenOrg)" ; break} 
        else { $ManagedBy=$Meta.value.DefaultObjectOwner} ;  ;
        #>

        $Retries = 4 ;
        $RetrySleep = 5 ;
        if(!$ThrottleMs){$ThrottleMs = 50 ;}
        $CredRole = 'CSVC' ; # role of svc to be dyn pulled from metaXXX if no -Credential spec'd, 
        if(!$rgxDriveBanChars){$rgxDriveBanChars = '[;~/\\\.:]' ; } ; # ;~/\.:, 
        
        if($useEXOP){
            $UseOP=$false ; 
            if((Get-Variable  -name "$($TenOrg)Meta").value.ExOPAccessFromToro -AND (Get-Variable  -name "$($TenOrg)Meta").value.Ex10Server){
                $UseOP = $true ; 
                $smsg = "$($TenOrg):Meta.ExOPAccessFromToro($((Get-Variable  -name "$($TenOrg)Meta").value.ExOPAccessFromToro)) -AND/OR Meta.Ex10Server($((Get-Variable  -name "$($TenOrg)Meta").value.Ex10Server)),`ENABLING use of OnPrem Ex system this pass." ; 
                if($verbose){ if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            } else { 
                $UseOP = $false ; 
                $smsg = "$($TenOrg):Meta.ExOPAccessFromToro($((Get-Variable  -name "$($TenOrg)Meta").value.ExOPAccessFromToro)) -AND/OR Meta.Ex10Server($((Get-Variable  -name "$($TenOrg)Meta").value.Ex10Server)),`nDISABLING use of OnPrem Ex system this pass." ; 
                if($verbose){ if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            } ; 
        } else { 
            # o365/EXO creds
            $o365Cred=$null ;
            <# Usage: Type defaults to SID, if not spec'd - Note: there must be a *logged in & configured *profile* 
            $o365Cred=get-TenantCredentials -TenOrg $TenOrg -verbose -userrole SID ;
            Returns a credential set for the $TenOrg Hybrid OnPrem Exchange Org
            .EXAMPLE
            $o365Cred=get-TenantCredentials -TenOrg $TenOrg -verbose -userrole CSVC ;
            Returns the CSVC Userrole credential for the $TenOrg Hybrid OnPrem Exchange Org
            .EXAMPLE
            $o365Cred=get-TenantCredentials -TenOrg $TenOrg -verbose -userrole B2BI ;
            Returns the B2BI Userrole credential for the $TenOrg Hybrid OnPrem Exchange Org
            #>
            #if($o365Cred=(get-TenantCredentials -TenOrg $TenOrg -verbose:$($verbose))){
            # force it to use the csvc mapping from $xxxmeta.o365_CSvcUpn, failthrough to SID spec 
            if($o365Cred=(get-TenantCredentials -TenOrg $TenOrg -UserRole 'CSVC','SID' -verbose:$($verbose))){
                # make it script scope, so we don't have to predetect & purge before using new-variable
                New-Variable -Name cred$($tenorg) -scope Script -Value $o365Cred.cred ;
                $smsg = "Resolved $($Tenorg) `$o365cred:$($o365Cred.cred.username) (assigned to `$cred$($tenorg))" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } else {
                #-=-record a STATUS=-=-=-=-=-=-=
                $statusdelta = ";ERROR";
                $script:PassStatus += $statusdelta ;
                set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta) ;
                #-=-=-=-=-=-=-=-=
                $smsg = "Unable to resolve $($tenorg) `$o365Cred value!"
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                throw "Unable to resolve $($tenorg) `$o365Cred value!`nEXIT!"
                exit ;
            } ;
            <# CALLS ARE IN FORM: (cred$($tenorg))
            $pltRXO = @{Credential = $Credential ; verbose = $($verbose) ; }
            $pltRXO = @{
                Credential = (Get-Variable -name cred$($tenorg) ).value ;
                verbose = $($verbose) ; } ; 
            if ($script:useEXOv2) { reconnect-eXO2 @pltRXO }
            else { reconnect-EXO @pltRXO } ;
            # or with Tenant-specific cred($Tenorg) lookup
            #>
        } ; 

        if($UseOP){
            # do the OP creds too
            $OPCred=$null ;
            # default to the onprem svc acct
            $pltGHOpCred=@{TenOrg=$TenOrg ;userrole='ESVC' ;verbose=$($verbose)} ;
            if($OPCred=(get-HybridOPCredentials @pltGHOpCred).cred){
                # make it script scope, so we don't have to predetect & purge before using new-variable
                New-Variable -Name "cred$($tenorg)OP" -scope Script -Value $OPCred ;
                $smsg = "Resolved $($Tenorg) `$OPCred:$($OPCred.username) (assigned to `$cred$($tenorg)OP)" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } else {
                #-=-record a STATUS=-=-=-=-=-=-=
                $statusdelta = ";ERROR";
                $script:PassStatus += $statusdelta ;
                set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta) ;
                #-=-=-=-=-=-=-=-=
                $smsg = "Unable to resolve get-HybridOPCredentials -TenOrg $($TenOrg) -userrole 'ESVC' value!"
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                throw "Unable to resolve $($tenorg) `$OPCred value!`nEXIT!"
                exit ;
            } ;

            # === Exchange LEMS/REMS detect & connect code

            $smsg= "Using EXOP cred:`$cred$($tenorg)OP:$((Get-Variable -name "cred$($tenorg)OP" ).value.username)" ;  
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
        } ; 
        
    } ;  # BEGIN-E
    PROCESS {
        #$ofile=".\$($ticket)-$($Mailbox)-folder-sizes-$(get-date -format 'yyyyMMdd-HHmmtt').csv" ;
        $error.clear() ;
    
        switch ($useEXOP){
            $false {

                write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):PERFORMING AN EXO MSGTRACE" ;
                if($VerbosePreference = "Continue"){
                    $VerbosePrefPrior = $VerbosePreference ;
                    $VerbosePreference = "SilentlyContinue" ;
                    $verbose = ($VerbosePreference -eq "Continue") ;
                } ; 
                disconnect-exo ; # pre-disconnect    
                $pltRXO = @{
                    Credential = (Get-Variable -name cred$($tenorg) ).value ;
                    verbose = $($verbose) ; }
                if ($script:useEXOv2) { reconnect-eXO2 @pltRXO }
                else { reconnect-EXO @pltRXO } ;
                if($VerbosePrefPrior -eq "Continue"){
                    $VerbosePreference = $VerbosePrefPrior ;
                    $verbose = ($VerbosePreference -eq "Continue") ;
                } ;

                # recycle $pltRXO for the AAD connection
                connect-AAD @pltRXO ;

                set-alias ps1GetMsgTrace Get-exoMessageTrace  ; 
                $props = $propsMsgEXO ; 
                $msgtrk=[ordered]@{
                    PageSize=1000 ;
                    Page=$null ;
                    StartDate=$null ;
                    EndDate=$null ;
                } ;
                if($Days -AND -not($StartDate -AND $EndDate)){
                    $msgtrk.StartDate=(get-date ([datetime]::Now)).adddays(-1*$days);
                    $msgtrk.EndDate=(get-date) ;
                } ;
                if($StartDate -and !($days)){
                    $msgtrk.StartDate=$(get-date $StartDate)
                } ;
                if($EndDate -and !($days)){
                    $msgtrk.EndDate=$(get-date $EndDate)
                } elseif($StartDate -and !($EndDate)){
                    write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):
    (StartDate w *NO* Enddate, asserting currenttime)" ;
                    $msgtrk.EndDate=(get-date) ;
                } ;
                
                $error.clear() ;
                TRY {
                    #Connect-AAD ;
                    $tendoms=Get-AzureADDomain ;
                } CATCH {
                    $ErrTrapt=$Error[0] ;
                    Write-Warning "$(get-date -format 'HH:mm:ss'): Failed processing $($_.Exception.ItemName). `nError Message: $($_.Exception.Message)`nError Details: $($_)" ;
                    $smsg = "FULL ERROR TRAPPED (EXPLICIT CATCH BLOCK WOULD LOOK LIKE): } catch[$($ErrTrpd.Exception.GetType().FullName)]{" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    Exit #Opts: STOP(debug)|EXIT(close)|CONTINUE(move on in loop cycle)|BREAK(exit loop iteration)|THROW $_/'CustomMsg'(end script with Err output)
                } ; 
            
                $Ten = ($tendoms |?{$_.name -like '*.mail.onmicrosoft.com'}).name.split('.')[0] ;
                $ofile ="$($ticket)-$($Ten)-$($Logon)-EXOMsgTrk" ;
                if($Sender){
                    if($Sender -match '\*'){
                        "(wild-card Sender detected)" ;
                        $msgtrk.add("SenderAddress",$Sender) ;
                    } else {
                        $msgtrk.add("SenderAddress",$Sender) ;
                    } ;
                    $ofile+=",From-$($Sender.replace("*","ANY"))" ;
                } ;
                if($Recipients){
                    if($Recipients -match '\*'){        "(wild-card Recipient detected)" ;
                        $msgtrk.add("RecipientAddress",$Recipients) ;
                    } else {
                            $msgtrk.add("RecipientAddress",$Recipients) ;
                    } ;
                    $ofile+=",To-$($Recipients.replace("*","ANY"))" ;
                } ;
                if($MessageId){
                    $msgtrk.add("MessageId",$MessageId) ;
                    $ofile+=",MsgId-$($MessageId.replace('<','').replace('>',''))" ;
                } ;
                if($MessageTraceId){
                    $msgtrk.add("MessageTraceId",$MessageTraceId) ;
                    $ofile+=",MsgId-$($MessageTraceId.replace('<','').replace('>',''))" ;
                } ;
                if($Subject){    $ofile+=",Subj-$($Subject.substring(0,[System.Math]::Min(10,$Subject.Length)))..." ;
                } ;
                if($Status){
                    $msgtrk.add("Status",$Status)  ;
                    $ofile+=",Status-$($Status)" ;
                } ;
                if($days){$ofile+= "-$($days)d-" } ;
                if($StartDate){$ofile+= "-$(get-date $StartDate -format 'yyyyMMdd-HHmmtt')-" } ;
                if($EndDate){$ofile+= "$(get-date $EndDate -format 'yyyyMMdd-HHmmtt')" } ;
                
                write-verbose -verbose:$true  "$((get-date).ToString('HH:mm:ss')):Running MsgTrk:$($Ten)" ;
    $(($msgtrk|out-string).trim()|out-default) ;
  
                TRY {
                    $Page = 1  ;
                    $Msgs=$null ;
                    do {
                        write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):Collecting - Page $($Page)..."  ;
                        $msgtrk.Page=$Page ;
                        $PageMsgs = ps1GetMsgTrace @msgtrk |  ?{$_.SenderAddress -notlike '*micro*' -or $_.SenderAddress -notlike '*root*' }  ;
                        $Page++  ;
                        $Msgs += @($PageMsgs)  ;
                    } until ($PageMsgs -eq $null) ;
                } CATCH {
                    Write-Warning "$(get-date -format 'HH:mm:ss'): Failed processing $($_.Exception.ItemName). `nError Message: $($_.Exception.Message)`nError Details: $($_)" ;
                    Exit ;
                } ; 
                $Msgs=$Msgs| Sort Received ;
                write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):==Msgs Returned:$(($Msgs|measure).count)" ;
                write-host -foregroundcolor white "$((get-date).ToString('HH:mm:ss')):Raw matches:$(($Msgs|measure).Count)" ;
                if($Subject){
                    write-host -foregroundcolor gray "$((get-date).ToString('HH:mm:ss')):Post-Filtering on Subject:$($Subject)" ;
                    $Msgs = $Msgs | ?{$_.Subject -like $Subject} ;
                    $ofile+="-Subj-$($Subject.replace("*"," ").replace("\"," "))" ;
                    write-host -foregroundcolor gray "$((get-date).ToString('HH:mm:ss')):Post Subj filter matches:$(($Msgs|measure).Count)" ;
                } ;
                $ofile+= "-run$(get-date -format 'yyyyMMdd-HHmmtt').csv" ;
                $ofile=[RegEx]::Replace($ofile, "[{0}]" -f ([RegEx]::Escape(-join [System.IO.Path]::GetInvalidFileNameChars())), '') ;
                $ofile=".\logs\$($ofile)" ;
                if($Msgs){
                    $Msgs | select $props | export-csv -notype -path $ofile  ;
                    write-host -foregroundcolor yellow "Status Distrib:" ;
                    write-host -foregroundcolor yellow "$((get-date).ToString('HH:mm:ss')):`n#*------v MOST RECENT MATCH v------" ;
                    write-host -foregroundcolor white "$(($msgs[-1]| format-list ReceivedLocal,StatusSenderAddress,RecipientAddress,Subject|out-string).trim())";
                    write-host -foregroundcolor yellow "$((get-date).ToString('HH:mm:ss')):`n#*------^ MOST RECENT MATCH ^------" ;
                    write-host -foregroundcolor yellow "$((get-date).ToString('HH:mm:ss')):`n#*------v Status DISTRIB v------" ;
                    "$(($Msgs | select -expand Status | group | sort count,count -desc | select count,name |out-string).trim())";
                    write-host -foregroundcolor yellow "$((get-date).ToString('HH:mm:ss')):`n#*------^ Status DISTRIB ^------" ;
                    if(test-path -path $ofile){
                            write-host -foregroundcolor white "$((get-date).ToString('HH:mm:ss')):(log file confirmed)" ;
                            Resolve-Path -Path $ofile | select -expand Path | out-clipboard ;
                            write-host -foregroundcolor yellow "$((get-date).ToString('HH:mm:ss')):$($Msgs.count) matches output to:`n'$($ofile)'`n(copied to CB)" ;
                    } else { "MISSING LOG FILE!" } ;

                    if($ReportFail){
                        $sBnr3="`n#*------v Status:FAIL Traffic (up to 1st $($ReportRowsLimit)) v------" ; 
                        write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($sBnr3)" ;
                        write-host -foregroundcolor cyan "$(($MSGS|?{$_.Status -eq 'FAIL'} | select -first $($ReportRowsLimit) | fl recipients,recipientstatus,ServerHostname|out-string).trim())" ; 
                        write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($sBnr3.replace('~v','~^').replace('v~','^~'))`n" ;
                    } ; 
                    
                    if($asObject){
                        $Msgs | write-output ; 
                    } ; 
                } else {
                    write-host -foregroundcolor yellow "$((get-date).ToString('HH:mm:ss')):NO MATCHES FOUND from::`n$(($msgtrk|out-string).trim()|out-default)`n(with any relevant ConnectorId postfilter)" ;
                } ;
            } ; # end EXO switchblock

            $true {
                write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):PERFORMING AN ONPREM MSGTRACK" ;
                if($VerbosePreference = "Continue"){
                    $VerbosePrefPrior = $VerbosePreference ;
                    $VerbosePreference = "SilentlyContinue" ;
                    $verbose = ($VerbosePreference -eq "Continue") ;
                } ; 
                # connect OP
                $pltRX10 = @{
                    Credential = (Get-Variable -name "cred$($tenorg)OP" ).value ;
                    verbose = $($verbose) ; } ;     
                if($pltRX10){
                    Connect-Ex2010 @pltRX10 ;
                } else { connect-Ex2010 ; } ;

                # reenable VerbosePreference:Continue, if set, during mod loads 
                if($VerbosePrefPrior -eq "Continue"){
                    $VerbosePreference = $VerbosePrefPrior ;
                    $verbose = ($VerbosePreference -eq "Continue") ;
                } ;

                set-alias ps1GetMsgTrace get-messagetrackinglog  ; 
                $props = $propsMsgEx10 ; 
                $msgtrk=@{
                    Start=(get-date ([datetime]::Now)).adddays(-1*$days) ;
                    End=(get-date) ;
                    resultsize="UNLIMITED" ;
                } ;
                # Page=$null ;
                $msgtrk=[ordered]@{
                    resultsize="UNLIMITED" ;
                    Start=$null ;
                    End=$null ;
                } ;
                if($Days -AND -not($StartDate -AND $EndDate)){
                    $msgtrk.Start=(get-date ([datetime]::Now)).adddays(-1*$days);
                    $msgtrk.End=(get-date) ;
                } ;
                if($StartDate -and !($days)){
                    $msgtrk.Start=$(get-date $StartDate)
                } ;
                if($EndDate -and !($days)){
                    $msgtrk.End=$(get-date $EndDate)
                } elseif($StartDate -and !($EndDate)){
                    write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):
    (StartDate w *NO* End, asserting currenttime)" ;
                    $msgtrk.End=(get-date) ;
                } ;
                TRY {
                    $Site=[System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::GetComputerSite().Name ;
                    # "$($ticket)-$($uid)-$($Site.substring(0,3))-MsgTrk" ;
                    $ofile ="$($ticket)-$($Site.substring(0,3))-OPMsgTrk" ;
                    if($Sender){$msgtrk.add("Sender",$Sender) ;
                        $ofile+=",From-$($Sender)" ;
                        } ;
                    if($Recipients){$msgtrk.add("Recipients",$Recipients) ;
                        $ofile+=",To-$($Recipients)" ;
                    } ;
                    if($Subject){$msgtrk.add("MessageSubject",$Subject)  ;
                        $ofile+=",Subj-$($Subject.substring(0,[System.Math]::Min(10,$Subject.Length)))..." ;
                    } ;
                    if($EventID){$msgtrk.add("EventID",$Status)  ;
                        $ofile+=",Evt-$($Status)" ;
                    } ;
                    
                    write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$((get-alias ps1GetMsgTrace).ResolvedCommandName) w`n$(($msgtrk|out-string).trim())" ; 
                    $Srvrs=(Get-ExchangeServer | where { $_.isHubTransportServer -eq $true -and $_.Site -match ".*\/$($Site)$"} | select -expand Name) ;
                    #$Msgs=($Srvrs| get-messagetrackinglog @msgtrk) | sort Timestamp ;
                    $Msgs =@() ; # 
                    # loop the servers, to provide a status output
                    foreach($Srvr in $Srvrs){
                        write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):Tracking $($Srvr) server..." ; 
                        $sMsgs = ($Srvr| get-messagetrackinglog @msgtrk) ;
                        write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):(($Srvr):$(($sMsgs|measure).count) matched msgs)" ; 
                        $Msgs+=$sMsgs ; 
                        $sMsgs = $null ; 
                    } ; 
                    #$Msgs = $Msgs |  sort Timestamp ;
                    $Msgs=$Msgs| Sort Timestamp ;
                    write-verbose -verbose:$true  "$((get-date).ToString('HH:mm:ss')):Raw matches:$(($Msgs|measure).Count)" ;
                    if($Connectorid){
                        write-verbose -verbose:$true  "$((get-date).ToString('HH:mm:ss')):Filtering on Conn:$($Connectorid)" ;
                        $Msgs = $Msgs | ?{$_.connectorid -like $Connectorid} ;
                        $ofile+="-conn-$($Connectorid.replace("*"," ").replace("\"," "))" ;
                        write-verbose -verbose:$true  "$((get-date).ToString('HH:mm:ss')):Post Conn filter matches:$(($Msgs|measure).Count)" ;
                    } ;
                    if($Source){
                        write-verbose -verbose:$true  "$((get-date).ToString('HH:mm:ss')):Filtering on Source:$($Source)" ;
                        $Msgs = $Msgs | ?{$_.Source -like $Source} ;
                        write-verbose -verbose:$true  "$((get-date).ToString('HH:mm:ss')):Post Src filter matches:$(($Msgs|measure).Count)" ;
                        $ofile+="-src-$($Source)" ;
                    } ;
                    if($Days){$ofile+= "-$($days)d-run$(get-date -format 'yyyyMMdd-HHmmtt').csv" ;} 
                    else {
                        $ofile+= "-$(get-date $msgtrk.Start -format 'yyyyMMdd-HHmmtt')-$(get-date $msgtrk.End -format 'yyyyMMdd-HHmmtt')-run$(get-date -format 'yyyyMMdd-HHmmtt').csv" ;
                    } ;  
                    $ofile=[RegEx]::Replace($ofile, "[{0}]" -f ([RegEx]::Escape(-join [System.IO.Path]::GetInvalidFileNameChars())), '') ;
                    $ofile=".\logs\$($ofile)" ;
                    
                    if($Msgs){
                        $Msgs | SELECT $props| EXPORT-CSV -notype -path $ofile ;
                        write-host -foregroundcolor white "$((get-date).ToString('HH:mm:ss')):`n#*------v MOST RECENT MATCH v------" ;
                        write-host -foregroundcolor cyan "$(((($msgs[-1]| format-list Timestamp,EventId,Sender,Recipients,MessageSubject|out-string).trim())|out-string).trim())" ; 
                        write-host -foregroundcolor white "$((get-date).ToString('HH:mm:ss')):`n#*------^ MOST RECENT MATCH ^------" ;
                        write-host -foregroundcolor white "$((get-date).ToString('HH:mm:ss')):`n#*------v EVENTID DISTRIB v------" ;
                        write-host -foregroundcolor cyan "$(($Msgs | select -expand EventId | group | sort count,count -desc | select count,name |out-string).trim())" ; 
                        write-host -fore gray "(SEND=SMTP SEND,TRANSFER=Routing,RESOLVE=Recipient conversion)" ;
                        write-host -foregroundcolor white "$((get-date).ToString('HH:mm:ss')):`n#*------^ EVENTID DISTRIB ^------" ;
                        if(test-path -path $ofile){
                            write-host -foregroundcolor white "$((get-date).ToString('HH:mm:ss')):(log file confirmed)" ;
                            Resolve-Path -Path $ofile | select -expand Path | out-clipboard ;
                            write-host -foregroundcolor white "$((get-date).ToString('HH:mm:ss')):$($Msgs.count) matches output to:`n'$($ofile)'`n(copied to CB)" ;
                        } else { "MISSING LOG FILE!" } ;
                        
                        if($ReportFail){
                            $sBnr3="`n#*~~~~~~v -ReportFail specified: Status:FAIL Traffic (up to 1st $($ReportRowsLimit)): v~~~~~~" ; 
                            write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($sBnr3)" ;
                            write-host -foregroundcolor cyan "$(((($MSGS|?{$_.eventid -eq 'fail'} | select -first $($ReportRowsLimit) | fl recipients,recipientstatus,ServerHostname|out-string).trim())|out-string).trim())" ; 
                            write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($sBnr3.replace('~v','~^').replace('v~','^~'))`n" ;
                        } ; 

                        if($asObject){
                            $Msgs | SELECT $props | write-output ; 
                        } ; 
                    } else {    write-host -foregroundcolor white "$((get-date).ToString('HH:mm:ss')):NO MATCHES FOUND from::`n$(($msgtrk|out-string).trim()|out-default)`n(with any relevant ConnectorId postfilter)" ;
                    } ;
                } CATCH {
                    Write-Warning "$(get-date -format 'HH:mm:ss'): Failed processing $($_.Exception.ItemName). `nError Message: $($_.Exception.Message)`nError Details: $($_)" ;
                    Exit ;
                } ; 
            } ;
            default {
                throw "UNRECOGNIZED useEXOP value)" ; exit ; 
            } ; 
        } ; # SWITCH-E
        
    } ;  # PROC-E
    END {
        remove-alias ps1GetMsgTrace ;
    } ; 
}

#*------^ get-MsgTrace.ps1 ^------

#*------v Get-OrgNameFromUPN.ps1 v------
function Get-OrgNameFromUPN{
    <#
    .SYNOPSIS
    Get-OrgNameFromUPN.ps1 - Extract organization name from UserPrincipalName ; localized verb-EXO vers of non-'$global:' helper funct from ExchangeOnlineManagement. The globals export fine, these don't and appear to need to be loaded manually
    .NOTES
    Version     : 1.0.0
    Author      : Todd Kadrie
    Website     :	http://www.toddomation.com
    Twitter     :	@tostka / http://twitter.com/tostka
    CreatedDate : 20201109-0833AM
    FileName    : Get-OrgNameFromUPN.ps1
    License     : [none specified]
    Copyright   : [none specified]
    Github      : https://github.com/tostka/verb-exo
    Tags        : Powershell
    AddedCredit : Microsoft (edited version of published commands in the module)
    AddedWebsite:	https://docs.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2
    REVISIONS
    * 8:34 AM 11/9/2020 init
    .DESCRIPTION
    Get-OrgNameFromUPN.ps1 - Extract organization name from UserPrincipalName ; localized verb-EXO vers of non-'$global:' helper funct from ExchangeOnlineManagement. The globals export fine, these don't and appear to need to be loaded manually

    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    None. Returns no objects or output.
    .EXAMPLE
    Get-OrgNameFromUPN
    Stock call
    .LINK
    https://github.com/tostka/verb-EXO
    .LINK
    https://docs.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2
    #>
    [CmdletBinding()]
    param([string] $UPN)
    $fields = $UPN -split '@'
    return $fields[-1]
}

#*------^ Get-OrgNameFromUPN.ps1 ^------

#*------v Invoke-EXOOnlineConnection.ps1 v------
function Invoke-ExoOnlineConnection{
    <#
    .SYNOPSIS
    Invoke-ExoOnlineConnection.ps1 - EXO non-ending MFA session, that renews it self ; once you connect to EXO with this it will stay open
    .NOTES
    Version     : 1.0.0
    Author      : Todd Kadrie
    Website     : http://www.toddomation.com
    Twitter     : @tostka / http://twitter.com/tostka
    CreatedDate : 2020-11-10
    FileName    : Invoke-ExoOnlineConnection.ps1
    License     : [none specified]
    Copyright   : [none specified]
    Github      : https://github.com/tostka/verb-XXX
    Tags        : Powershell
    AddedCredit : Mahmoud Badran
    AddedWebsite: https://techcommunity.microsoft.com/t5/exchange/60-minutes-timeout-on-mfa-session/m-p/559224
    REVISIONS
    .DESCRIPTION
    Invoke-ExoOnlineConnection.ps1 - EXO non-ending MFA session, that renews it self ; once you connect to EXO with this it will stay open
    normally came as a .ps1 with a local function. Haven't tested, looks like it should work, trick is to preregister the timer/check interval outside of the function, prior to call.
    .PARAMETER  Checktimer
    Switch to trigger a timercheck. [-Checktimer]
    PARAMETERRepairPSSession
    Switch to trigger a session repair. [-RepairPSSession]
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    None. Returns no objects or output
    .EXAMPLE
    ## Create an Timer instance to trackand recheck status
    $timer = New-Object Timers.Timer
    ## Now setup the Timer instance to fire events
    $timer.Interval = 600000
    $timer.AutoReset = $true  # enable the event again after its been fired
    $timer.Enabled = $true
    ## register your event
    ## $args[0] Timer object
    ## $args[1] Elapsed event properties
    Register-ObjectEvent -InputObject $timer -EventName Elapsed -SourceIdentifier Repair  -Action {Invoke-ExoOnlineConnection -Checktimer}
    .EXAMPLE
    .LINK
    https://github.com/tostka/verb-EXO
    .LINK
    https://techcommunity.microsoft.com/t5/exchange/60-minutes-timeout-on-mfa-session/m-p/559224
    #>
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$false,HelpMessage = "Switch to trigger a timercheck. [-Checktimer]")]
        [switch]$Checktimer,
        [Parameter(mandatory=$false, valuefrompipeline=$false,HelpMessage = "Switch to trigger a session repair. [-RepairPSSession]")]
        [switch]$RepairPSSession,
        [Parameter(HelpMessage = "Credential to use for this connection [-credential [credential obj variable]")]
        [System.Management.Automation.PSCredential]$Credential = $global:credo365TORSID
    )
    BEGIN{
        if(!$Global:ErrorActionPreference){$Global:ErrorActionPreference = "Stop"} ; 
        if(!$Global:VerbosePreference){$Global:VerbosePreference = "Continue"} ; 
        #if(!$office365UserPrincipalName){$office365UserPrincipalName = "ADMIN@o365.com" } ; 
        if(!$PSExoPowershellModuleRoot){$PSExoPowershellModuleRoot = (Get-ChildItem -Path $($env:LOCALAPPDATA+"\Apps\2.0\") -Filter CreateExoPSSession.ps1 -Recurse -ErrorAction SilentlyContinue -Force | Select -Last 1).DirectoryName } ; 
        if(!$ExoPowershellModule){$ExoPowershellModule = "Microsoft.Exchange.Management.ExoPowershellModule.dll"} ; 
        if(!$ExoPowershellModulePath){$ExoPowershellModulePath = [System.IO.Path]::Combine($PSExoPowershellModuleRoot, $ExoPowershellModule)} ; 
        if(!(get-module $ExoPowershellModule.replace('.dll','') )){Import-Module $ExoPowershellModulePath -verbose:$false } ; 
    }
    PROCESS{
        #determine if  PsSession is loaded in memory
        $ExosessionInfo = Get-PsSession
        #calculate session time style: $global:_EXO_ExchangeEnvironmentName = $ExchangeEnvironmentName;
        # MS uses these global name
        if ($global:_EXO_ExosessionStartTime){
             $global:_EXO_ExosessionTotalTime = ((Get-Date) - $global:_EXO_ExosessionStartTime)
        }
        #need to loop through each session a user might have opened previously
        foreach ($ExosessionItem in $ExosessionInfo){
            #check session timer to know if we need to break the connection in advance of a timeout. Break and make new after 40 minutes.
            if ($ExosessionItem.ComputerName.Contains("outlook.office365.com") -and $ExosessionItem.State -eq "Opened" -and $global:_EXO_ExosessionTotalTime.TotalSeconds -ge "2400"){
                Write-Verbose -Message "The PowerShell session has been running for $($global:_EXO_ExosessionTotalTime.TotalMinutes) minutes. We need to shut it down and create a new session due to the access token expiration at 60 minutes."
                $ExosessionItem | Remove-PSSession
                Start-Sleep -Seconds 3
                $strSessionFound = $false
                $global:_EXO_ExosessionTotalTime = $null #reset the timer
            } else { Write-Verbose -Message "The PowerShell session has been running for $($global:_EXO_ExosessionTotalTime.TotalMinutes) minutes.)"}
            #Force repair PSSession
            if ($ExosessionItem.ComputerName.Contains("outlook.office365.com") -and $RepairPSSession){
                Write-Verbose -Message "Attempting to repair broken PowerShell session to Exchange Online using cached credential."
                $ExosessionItem | Remove-PSSession
                Start-Sleep -Seconds 3
                $strSessionFound = $false
                $global:_EXO_ExosessionTotalTime = $null
            }elseif ($ExosessionItem.ComputerName.Contains("outlook.office365.com") -and $ExosessionItem.State -eq "Opened"){
                $strSessionFound = $true
            }
        }
        if (!$strSessionFound){
            Write-Verbose -Message "Creating new Exchange Online PowerShell session..."
            try{
                $pltNEXOS = @{
                    ExchangeEnvironmentName         = $ExchangeEnvironmentName ;
                    ConnectionUri                   = "https://outlook.office365.com/powershell-liveid/" ;
                    #AzureADAuthorizationEndpointUri = $AzureADAuthorizationEndpointUri ;
                    UserPrincipalName               = $Credential.username ;
                    PSSessionOption                 = $PSSessionOption ;
                    #Credential                      = $Credential ;
                    BypassMailboxAnchoring          = $($BypassMailboxAnchoring) ;
                    #ShowProgress                    = $($showProgress) # isn't a param of new-exopssessoin, is used with set-exo
                    #DelegatedOrg                    = $DelegatedOrganization ;
                    ErrorAction                      = 'SilentlyContinue' ; 
                    ErrorVariable                    = $newOnlineSessionError ; 
                }
                #$ExoSession  = New-ExoPSSession -UserPrincipalName $Credential.username -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -ErrorAction SilentlyContinue -ErrorVariable $newOnlineSessionError
                write-verbose "New-ExoPSSession w`n$(($pltNEXOS|out-string).trim())" ; 
                $ExoSession  = New-ExoPSSession @pltNEXOS ; 
            }catch{
                Write-Verbose -Message "Throw error..."
                throw;
            } finally {
                if ($newOnlineSessionError) {
                 Write-Verbose -Message "Final error..."
                    throw $newOnlineSessionError
                }
            }
            Write-Verbose -Message "Importing remote PowerShell session..."
            $global:_EXO_ExosessionStartTime = (Get-Date)
            #Import-PSSession $ExoSession -AllowClobber | Out-Null
            Import-PSSession $ExoSession -AllowClobber -DisableNameChecking
        } ;
    } ;
    END{} ;
}

#*------^ Invoke-EXOOnlineConnection.ps1 ^------

#*------v Print-Details.ps1 v------
function Print-Details{
    <#
    .SYNOPSIS
    Print-Details.ps1 - localized verb-EXO vers of non-'$global:' helper funct from ExchangeOnlineManagement. The globals export fine, these don't and appear to need to be loaded manually
    .NOTES
    Version     : 1.0.0
    Author      : Todd Kadrie
    Website     :	http://www.toddomation.com
    Twitter     :	@tostka / http://twitter.com/tostka
    CreatedDate : 20201109-0833AM
    FileName    : Print-Details.ps1
    License     : [none specified]
    Copyright   : [none specified]
    Github      : https://github.com/tostka/verb-exo
    Tags        : Powershell
    AddedCredit : Microsoft (edited version of published commands in the module)
    AddedWebsite:	https://docs.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2
    REVISIONS
    * 8:34 AM 11/9/2020 init
    .DESCRIPTION
    Print-Details.ps1 - localized verb-EXO vers of non-'$global:' helper funct from ExchangeOnlineManagement. The globals export fine, these don't and appear to need to be loaded manually
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    None. Returns no objects or output.
    .EXAMPLE
    Print-Details
    Stock call
    .LINK
    https://github.com/tostka/verb-EXO
    .LINK
    https://docs.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2
    #>
    [CmdletBinding()]
    param()
    Write-Host -ForegroundColor Yellow ""
    Write-Host -ForegroundColor Yellow "----------------------------------------------------------------------------"
    Write-Host -ForegroundColor Yellow "The module allows access to all existing remote PowerShell (V1) cmdlets in addition to the 9 new, faster, and more reliable cmdlets."
    Write-Host -ForegroundColor Yellow ""
    Write-Host -ForegroundColor Yellow "|--------------------------------------------------------------------------|"
    Write-Host -ForegroundColor Yellow "|    Old Cmdlets                    |    New/Reliable/Faster Cmdlets       |"
    Write-Host -ForegroundColor Yellow "|--------------------------------------------------------------------------|"
    Write-Host -ForegroundColor Yellow "|    Get-CASMailbox                 |    Get-EXOCASMailbox                 |"
    Write-Host -ForegroundColor Yellow "|    Get-Mailbox                    |    Get-EXOMailbox                    |"
    Write-Host -ForegroundColor Yellow "|    Get-MailboxFolderPermission    |    Get-EXOMailboxFolderPermission    |"
    Write-Host -ForegroundColor Yellow "|    Get-MailboxFolderStatistics    |    Get-EXOMailboxFolderStatistics    |"
    Write-Host -ForegroundColor Yellow "|    Get-MailboxPermission          |    Get-EXOMailboxPermission          |"
    Write-Host -ForegroundColor Yellow "|    Get-MailboxStatistics          |    Get-EXOMailboxStatistics          |"
    Write-Host -ForegroundColor Yellow "|    Get-MobileDeviceStatistics     |    Get-EXOMobileDeviceStatistics     |"
    Write-Host -ForegroundColor Yellow "|    Get-Recipient                  |    Get-EXORecipient                  |"
    Write-Host -ForegroundColor Yellow "|    Get-RecipientPermission        |    Get-EXORecipientPermission        |"
    Write-Host -ForegroundColor Yellow "|--------------------------------------------------------------------------|"
    Write-Host -ForegroundColor Yellow ""
    Write-Host -ForegroundColor Yellow "To get additional information, run: Get-Help Connect-ExchangeOnline or check https://aka.ms/exops-docs"
    Write-Host -ForegroundColor Yellow ""
    Write-Host -ForegroundColor Yellow "Send your product improvement suggestions and feedback to exocmdletpreview@service.microsoft.com. For issues related to the module, contact Microsoft support. Don't use the feedback alias for problems or support issues."
    Write-Host -ForegroundColor Yellow "----------------------------------------------------------------------------"
    Write-Host -ForegroundColor Yellow ""
}

#*------^ Print-Details.ps1 ^------

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
    * 3:20 PM 3/31/2021 fixed pssess typo
    * 8:30 AM 10/22/2020 added $TenOrg, swapped looping meta resolve with 1-liner approach ; added AcceptedDom caching to the middle status test (suppress one more get-exoaccepteddomain call if possible)
    * 1:30 PM 9/21/2020 added caching of AcceptedDomain, dynamically into XXXMeta - checks for .o365_AcceptedDomains, and pops w (Get-exoAcceptedDomain).domainname when blank. 
        As it's added to the $global meta, that means it stays cached cross-session, completely eliminates need to dyn query per rxo, after the first one, that stocks the value
    * 2:39 PM 8/4/2020 fixed -match "^(Session|WinRM)\d*" rgx (lacked ^, mismatched EXOv2 conns)
    * 10:35 AM 7/28/2020 tweaked retry loop to not retry-sleep 1st attempt
    * 3:24 PM 7/24/2020 updated to support tenant-alignment & sub'd out showdebug for verbose
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
      [Parameter(HelpMessage="Credential to use for this connection [-credential [credential obj variable]")]
      [System.Management.Automation.PSCredential]$Credential = $global:credo365TORSID,
      [Parameter(HelpMessage="Debugging Flag [-showDebug]")]
      [switch] $showDebug
    ) ;
    $verbose = ($VerbosePreference -eq "Continue") ; 
    if(!$rgxExoPsHostName){$rgxExoPsHostName="^(ps\.outlook\.com|outlook\.office365\.com)$" } ;

    $TenOrg = get-TenantTag -Credential $Credential ;
    
    # if we're using EXOv1-style BasicAuth, clear incompatible existing EXOv2 PSS's
    $exov2Good = Get-PSSession | where-object {($_.ConfigurationName -like "Microsoft.Exchange") -AND (
            $_.Name -like "ExchangeOnlineInternalSession*") -AND ($_.State -like "*Opened*") -AND (
            $_.Availability -eq 'Available')} ; 
    $exov2Broken = Get-PSSession | where-object {($_.ConfigurationName -like "Microsoft.Exchange") -AND (
        $_.Name -eq "ExchangeOnlineInternalSession*") -AND ($_.State -like "*Broken*")}
    $exov2Closed = Get-PSSession | where-object {($_.ConfigurationName -like "Microsoft.Exchange") -AND (
        $_.Name -eq "ExchangeOnlineInternalSession*") -AND ($_.State -like "*Closed*")}

    if($exov2Good  ){
        write-verbose "EXOv1:Disconnecting conflicting EXOv2 connection" ; 
        DisConnect-EXO2 ; 
    } ; 
    if ($exov2Broken.count -gt 0){for ($index = 0 ;$index -lt $psBroken.count ;$index++){Remove-PSSession -session $psBroken[$index]} };
    if ($exov2Closed.count -gt 0){for ($index = 0 ;$index -lt $psClosed.count ; $index++){Remove-PSSession -session $psClosed[$index] } } ; 
    
    # fault tolerant looping exo connect, don't let it exit until a connection is present, and stable, or return error for hard time out
    $tryNo=0 ; $1F=$false ;
    Do {
        if($1F){Sleep -s 5} ;
        $tryNo++ ;
        write-host "." -NoNewLine; if($tryNo -gt 1){Start-Sleep -m (1000 * 5)} ;
        # appears MFA may not properly support passing back a session vari, so go right to strict hostname matches

        $legPSSession = Get-PSSession | where-object {$_.ConfigurationName -like "Microsoft.Exchange" -AND $_.Name -match "^(Session|WinRM)\d*" -AND ($_.State -eq 'Opened') -AND ($_.Availability -eq 'Available')}
        
        if( (Get-PSSession | where-object {$_.ConfigurationName -like "Microsoft.Exchange" -AND $_.Name -match "^(Session|WinRM)\d*" -AND (
                ($_.State -ne 'Opened') -OR ($_.Availability -ne 'Available')) }) -OR (
                -not(Get-PSSession | where-object {$_.ConfigurationName -like "Microsoft.Exchange" -AND (
                $_.Name -match "^(Session|WinRM)\d*")})) ){
            write-verbose "$((get-date).ToString('HH:mm:ss')):Reconnecting:No existing PSSESSION matching Name -match (Session|WinRM) with valid Open/Availability:$((Get-PSSession|Where-Object{$_.ComputerName -match $rgxExoPsHostName}| Format-Table -a State,Availability |out-string).trim())" ;
            Disconnect-Exo ; Disconnect-PssBroken ;Start-Sleep -Seconds 3;
            if(!$Credential){
                connect-EXO ;
            } else {
                connect-EXO -credential:$($Credential) ;
            } ;
        
        }elseif($legPSSession){
            # implement caching of accepteddoms into the XXXMeta, in the session (cut back on queries to EXO on acceptedom)
            #$TenOrg = get-TenantTag -Credential $Credential ;
            if( get-command Get-exoAcceptedDomain -ea 0) {
                if(!(Get-Variable  -name "$($TenOrg)Meta").value.o365_AcceptedDomains){
                    set-Variable  -name "$($TenOrg)Meta" -value ( (Get-Variable  -name "$($TenOrg)Meta").value += @{'o365_AcceptedDomains' = (Get-exoAcceptedDomain).domainname} )
                } ;
            } ; 
            <#
            $credDom = ($Credential.username.split("@"))[1] ;
            $Metas=(get-variable *meta|Where-Object{$_.name -match '^\w{3}Meta$'}) ;
            foreach ($Meta in $Metas){
                if( ($credDom -eq $Meta.value.legacyDomain) -OR ($credDom -eq $Meta.value.o365_TenantDomain) -OR ($credDom -eq $Meta.value.o365_OPDomain)){
                    if(!$Meta.value.o365_AcceptedDomains){
                        set-variable -Name $meta.name -Value ((get-variable -name $meta.name).value  += @{'o365_AcceptedDomains' = (Get-exoAcceptedDomain).domainname} )
                    } ; 
                    break ;
                } ;
            } ;
            #>
                #if((Get-exoAcceptedDomain).domainname.contains($Credential.username.split('@')[1].tostring())){
            # do caching & check cached value, not qry unless unpopulated (first pass in global session)
            #if($Meta.value.o365_AcceptedDomains.contains($Credential.username.split('@')[1].tostring())){
            if((Get-Variable  -name "$($TenOrg)Meta").value.o365_AcceptedDomains.contains($Credential.username.split('@')[1].tostring())){
                # validate that the connected EXO is to the $Credential tenant    
                write-verbose "(Authenticated to EXO:$($Credential.username.split('@')[1].tostring()))" ; 
            # issue: found fresh bug in cxo: svcacct UPN suffix @tenantname.onmicrosoft.com, but testing against AccepteDomain, it's not in there (tho @toroco.mail.onmicrosoft.comis)
            }elseif((Get-Variable  -name "$($TenOrg)Meta").value.o365_TenantDomain -eq ($Credential.username.split('@')[1].tostring())){
                $smsg = "(EXO Authenticated & Functional(TenDom):$($Credential.username.split('@')[1].tostring()))" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                $bExistingEXOGood = $true ;
            } else { 
                write-verbose "(NOT Authenticated to Credentialed Tenant:$($Credential.username.split('@')[1].tostring()))" ; 
                Write-Host "Authenticating to EXO:$($Credential.username.split('@')[1].tostring())..."  ;
                Disconnect-Exo ; Disconnect-PssBroken ;Start-Sleep -Seconds 3;
                if(!$Credential){
                    connect-EXO -verbose:$($verbose) ;
                } else {
                    connect-EXO -credential:$($Credential) -verbose:$($verbose) ;
                } ;
            } ; 
        } else {
            throw "FAILED EXO CONNECT!"
        } ; 
        $1F=$true ;
        if($tryNo -gt $DoRetries ){throw "RETRIED EXO CONNECT $($tryNo) TIMES, ABORTING!" } ;
    } Until ((Get-PSSession |Where-Object{$_.ComputerName -match $rgxExoPsHostName -AND $_.State -eq "Opened" -AND $_.Availability -eq "Available"}))
}

#*------^ Reconnect-EXO.ps1 ^------

#*------v Reconnect-EXO2.ps1 v------
Function Reconnect-EXO2 {
   <#
    .SYNOPSIS
    Reestablish connection to Exchange Online (via EXO V2 graph-api module)
    .NOTES
    Updated By: : Todd Kadrie
    Website:	http://toddomation.com
    Twitter:	@tostka http://twitter.com/tostka
    Based on original function Author: ExactMike Perficient, Global Knowl... (Partner)
    Website:	https://social.technet.microsoft.com/Forums/msonline/en-US/f3292898-9b8c-482a-86f0-3caccc0bd3e5/exchange-powershell-monitoring-remote-sessions?forum=onlineservicesexchange
    REVISIONS   :
    # 8:34 AM 3/31/2021 added verbose suppress to all import-mods
    * 8:30 AM 10/22/2020 added $TenOrg, swapped looping meta resolve with 1-liner approach ; added AcceptedDom caching to the middle status test (suppress one more get-xoaccepteddomain call if possible)
    * 1:30 PM 9/21/2020 added caching of AcceptedDomain, dynamically into XXXMeta - checks for .o365_AcceptedDomains, and pops w (Get-exoAcceptedDomain).domainname when blank. 
        As it's added to the $global meta, that means it stays cached cross-session, completely eliminates need to dyn query per rxo, after the first one, that stocks the value
    * 1:45 PM 8/11/2020 added trailing test-EXOToken confirm
    * 2:39 PM 8/4/2020 fixed -match "^(Session|WinRM)\d*" rgx (lacked ^, mismatched EXOv2 conns)
    * 3:55 PM 7/30/2020 rewrite/port from reconnect-EXO to replace import-pssession with new connect-ExchangeOnline cmdlet (supports MFA natively) - #127 # *** LEFT OFF HERE 5:01 PM 7/29/2020 *** not sure if it supports allowclobber, if it's actually wrapping pssession, it sure as shit does!
    * 10:35 AM 7/28/2020 tweaked retry loop to not retry-sleep 1st attempt
    * 3:24 PM 7/24/2020 updated to support tenant-alignment & sub'd out showdebug for verbose
    * 11:48 AM 5/27/2020 added func alias:rxo within the func
    * 2:38 PM 4/20/2020 added local $rgxExoPsHostName
    * 8:45 AM 3/3/2020 public cleanup
    * 9:52 PM 1/16/2020 cleanup
    * 1:07 PM 11/25/2019 added *tol/*tor/*cmw alias variants for connect & reconnect
    * 9:52 AM 11/20/2019 spliced in credential matl
    * 2:55 PM 10/11/2018 connect-exo: added pre sleep skip on tryno 1
    * 8:04 AM 11/20/2017 code in a loop in the Reconnect-EXO2, until it hits or 100% times out
    # 8:47 AM 6/2/2017 cleaned up deadwood, simplified pshelp
    * 7:58 AM 3/15/2017 ren Disconnect/Connect/Reconnect-EXO2 => Disconnect/Connect/Reconnect-EXO2, added pss, doing tweaks to put into prod use
    * 2/10/14 posted version
    .DESCRIPTION
    I use this for routine test/reconnect of EXO. His orig use was within batches, to break up and requeue chunks of commands.
    Mike's original comment: Below is one
    example of how I batch items for processing and use the
    Reconnect-EXO2 function.  I'm still experimenting with how to best
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
    Reconnect-EXO2;
    Reconnect EXO connection
    .EXAMPLE
    $batchsize = 10 ;
    $RecordCount=$mr.count #this is the array of whatever you are processing ;
    $b=0 #this is the initialization of a variable used in the do until loop below ;
    $mrs = @() ;
    do {
        Write-Progress -Activity "Getting move request statistics for all $wave move requests." -Status "Processing Records $b through $($b+$batchsize) of $RecordCount." -PercentComplete ($b/$RecordCount*100) ;
        $mrs += $mr | Select-Object -skip $b -first $batchsize | foreach-object {Reconnect-EXO2; $_ | Get-OLMoveRequestStatistics} ;
        $b=$b+$batchsize ;
        } ;
    until ($b -gt $RecordCount) ;
    
    .LINK
    https://social.technet.microsoft.com/Forums/msonline/en-US/f3292898-9b8c-482a-86f0-3caccc0bd3e5/exchange-powershell-monitoring-remote-sessions?forum=onlineservicesexchange
    #>
    [CmdletBinding()]
    [Alias('rxo2')]
    Param(
      [Parameter(HelpMessage="Use Proxy-Aware SessionOption settings [-ProxyEnabled]")]
      [boolean]$ProxyEnabled = $False,
      [Parameter(HelpMessage="Credential to use for this connection [-credential [credential obj variable]")]
      [System.Management.Automation.PSCredential]$Credential = $global:credo365TORSID,
      [Parameter(HelpMessage="Debugging Flag [-showDebug]")]
      [switch] $showDebug
    ) ;
    BEGIN{
        $verbose = ($VerbosePreference -eq "Continue") ; 
        if(!$rgxExoPsHostName){$rgxExoPsHostName="^(ps\.outlook\.com|outlook\.office365\.com)$" } ;
        $modname = 'ExchangeOnlineManagement' ; 
        #Try {Get-Module $modname -listavailable -ErrorAction Stop | out-null } Catch {Install-Module $modname -scope CurrentUser ; } ;                 # installed
        Try {Get-Module $modname -ErrorAction Stop | out-null } Catch {Import-Module -Name $modname -MinimumVersion '1.0.1' -ErrorAction Stop -verbose:$false  } ; # imported
        
        $TenOrg = get-TenantTag -Credential $Credential ;

    } ;  # BEG-E
    PROCESS{
        $bExistingEXOGood = $false ; 
        if( $legPSSession = Get-PSSession | where-object {$_.ConfigurationName -like "Microsoft.Exchange" -AND $_.Name -match "^(Session|WinRM)\d*" } ){
            # ignore state & Avail, close the conflicting legacy conn's
            Disconnect-Exo ; Disconnect-PssBroken ;Start-Sleep -Seconds 3;
            $bExistingEXOGood = $false ; 
        } ; 
        #clear invalid existing EXOv2 PSS's
        $exov2Broken = Get-PSSession | where-object {$_.ConfigurationName -like "Microsoft.Exchange" -AND (
            $_.Name -eq "ExchangeOnlineInternalSession*") -AND $_.State -like "*Broken*"}
        $exov2Closed = Get-PSSession | where-object {$_.ConfigurationName -like "Microsoft.Exchange" -AND (
            $_.Name -eq "ExchangeOnlineInternalSession*") -AND $_.State -like "*Closed*"}
        
        if ($exov2Broken.count -gt 0){for ($index = 0 ;$index -lt $exov2Broken.count ;$index++){Remove-PSSession -session $exov2Broken[$index]} };
        if ($exov2Closed.count -gt 0){for ($index = 0 ;$index -lt $exov2Closed.count ; $index++){Remove-PSSession -session $exov2Closed[$index] } } ; 
    
        # appears MFA may not properly support passing back a session vari, so go right to strict hostname matches
        $exov2Good = Get-PSSession | where-object {$_.ConfigurationName -like "Microsoft.Exchange" -AND (
            $_.Name -like "ExchangeOnlineInternalSession*") -AND $_.State -like "*Opened*" -AND (
            $_.Availability -eq 'Available')} ; 

        if($exov2Good){
            if( get-command Get-xoAcceptedDomain -ea 0) {
                # add accdom caching
                #$TenOrg = get-TenantTag -Credential $Credential ;
                if(!(Get-Variable  -name "$($TenOrg)Meta").value.o365_AcceptedDomains){
                    set-Variable  -name "$($TenOrg)Meta" -value ( (Get-Variable  -name "$($TenOrg)Meta").value += @{'o365_AcceptedDomains' = (Get-xoAcceptedDomain).domainname} )
                } ;
                #if ((Get-xoAcceptedDomain).domainname.contains($Credential.username.split('@')[1].tostring())){
                if((Get-Variable  -name "$($TenOrg)Meta").value.o365_AcceptedDomains.contains($Credential.username.split('@')[1].tostring())){
                    # validate that the connected EXO is to the $Credential tenant    
                    $smsg = "(EXO Authenticated & Functional(AccDom):$($Credential.username.split('@')[1].tostring()))" ; 
                    $bExistingEXOGood = $true ; 
                # issue: found fresh bug in cxo: svcacct UPN suffix @tenantname.onmicrosoft.com, but testing against AccepteDomain, it's not in there (tho @toroco.mail.onmicrosoft.comis)
                }elseif((Get-Variable  -name "$($TenOrg)Meta").value.o365_TenantDomain -eq ($Credential.username.split('@')[1].tostring())){
                    $smsg = "(EXO Authenticated & Functional(TenDom):$($Credential.username.split('@')[1].tostring()))" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    $bExistingEXOGood = $true ;
                } else { 
                    write-verbose "(Credential mismatch:disconnecting from existing EXO:$($eEXO.Identity) tenant)" ; 
                    DisConnect-EXO2 ; 
                    $bExistingEXOGood = $false ; 
                } ; 
            } else { 
                # capture outlier: shows a session wo the test cmdlet, force reset
                DisConnect-EXO2 ; 
                $bExistingEXOGood = $false ; 
            } ; 
        } ; 

        if($bExistingEXOGood -eq $false){
            connect-exo2 -Credential $Credential -verbose:$($verbose) ; 
        } ; 

    } ;  # PROC-E
    END {
        # if ( (get-module -name tmp_* | ForEach-Object { Get-Command -module $_.name -name 'Add-xoAvailabilityAddressSpace' -ea 0 }) -AND (test-EXOToken) ) {
        if( (Get-PSSession | where-object {$_.ConfigurationName -like "Microsoft.Exchange" -AND $_.Name -like "ExchangeOnlineInternalSession*" -AND $_.State -like "*Opened*" -AND ($_.Availability -eq 'Available')}) -AND (test-EXOToken) ){ 
            # implement caching of accepteddoms into the XXXMeta, in the session (cut back on queries to EXO on acceptedom)
            # non-looping
            
            if( get-command Get-xoAcceptedDomain -ea 0) {
                #$TenOrg = get-TenantTag -Credential $Credential ;
                if(!(Get-Variable  -name "$($TenOrg)Meta").value.o365_AcceptedDomains){
                    set-Variable  -name "$($TenOrg)Meta" -value ( (Get-Variable  -name "$($TenOrg)Meta").value += @{'o365_AcceptedDomains' = (Get-xoAcceptedDomain).domainname} )
                } ;
            } ; 
            <#
            $credDom = ($Credential.username.split("@"))[1] ;
            $Metas=(get-variable *meta|Where-Object{$_.name -match '^\w{3}Meta$'}) ;
            foreach ($Meta in $Metas){
                if( ($credDom -eq $Meta.value.legacyDomain) -OR ($credDom -eq $Meta.value.o365_TenantDomain) -OR ($credDom -eq $Meta.value.o365_OPDomain)){
                    if(!$Meta.value.o365_AcceptedDomains){
                        set-variable -Name $meta.name -Value ((get-variable -name $meta.name).value  += @{'o365_AcceptedDomains' = (Get-xoAcceptedDomain).domainname} )
                    } ; 
                    break ;
                } ;
            } ;
            #>
            #if ((Get-xoAcceptedDomain).domainname.contains($Credential.username.split('@')[1].tostring())){
            #if($Meta.value.o365_AcceptedDomains.contains($Credential.username.split('@')[1].tostring())){
            if((Get-Variable  -name "$($TenOrg)Meta").value.o365_AcceptedDomains.contains($Credential.username.split('@')[1].tostring())){
                # validate that the connected EXO is to the $Credential tenant    
                write-verbose "(EXOv2 Authenticated & Functional:$($Credential.username.split('@')[1].tostring()))" ; 
                $bExistingEXOGood = $true ; 
            # issue: found fresh bug in cxo: svcacct UPN suffix @tenantname.onmicrosoft.com, but testing against AccepteDomain, it's not in there (tho @toroco.mail.onmicrosoft.comis)
            }elseif((Get-Variable  -name "$($TenOrg)Meta").value.o365_TenantDomain -eq ($Credential.username.split('@')[1].tostring())){
                $smsg = "(EXO Authenticated & Functional(TenDom):$($Credential.username.split('@')[1].tostring()))" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                $bExistingEXOGood = $true ;
            } else { 
                write-error "(Credential mismatch:disconnecting from existing EXO:$($eEXO.Identity) tenant)" ; 
                Disconnect-exo2 ; 
                $bExistingEXOGood = $false ; 
            } ;
        } ; 
    } ; # END-E 
}

#*------^ Reconnect-EXO2.ps1 ^------

#*------v Reconnect-EXO2old.ps1 v------
Function Reconnect-EXO2old {
   <#
    .SYNOPSIS
    Reestablish connection to Exchange Online (via EXO V2 graph-api module)
    .NOTES
    Updated By: : Todd Kadrie
    Website:	http://toddomation.com
    Twitter:	@tostka http://twitter.com/tostka
    Based on original function Author: ExactMike Perficient, Global Knowl... (Partner)
    Website:	https://social.technet.microsoft.com/Forums/msonline/en-US/f3292898-9b8c-482a-86f0-3caccc0bd3e5/exchange-powershell-monitoring-remote-sessions?forum=onlineservicesexchange
    REVISIONS   :
    # 11:21 AM 3/31/2021 added TenDom test, after AccDom test ; 
    * 2:08 PM 11/10/2020 ren'd the older connect-ExchangeOnline-related version, to reconnect-exo2old, in favor of the name going on the NewEXOPSSessoin-based version.
    * 8:30 AM 10/22/2020 added $TenOrg, swapped looping meta resolve with 1-liner approach ; added AcceptedDom caching to the middle status test (suppress one more get-xoaccepteddomain call if possible)
    * 1:30 PM 9/21/2020 added caching of AcceptedDomain, dynamically into XXXMeta - checks for .o365_AcceptedDomains, and pops w (Get-exoAcceptedDomain).domainname when blank. 
        As it's added to the $global meta, that means it stays cached cross-session, completely eliminates need to dyn query per rxo, after the first one, that stocks the value
    * 1:45 PM 8/11/2020 added trailing test-EXOToken confirm
    * 2:39 PM 8/4/2020 fixed -match "^(Session|WinRM)\d*" rgx (lacked ^, mismatched EXOv2 conns)
    * 3:55 PM 7/30/2020 rewrite/port from reconnect-EXO to replace import-pssession with new connect-ExchangeOnline cmdlet (supports MFA natively) - #127 # *** LEFT OFF HERE 5:01 PM 7/29/2020 *** not sure if it supports allowclobber, if it's actually wrapping pssession, it sure as shit does!
    * 10:35 AM 7/28/2020 tweaked retry loop to not retry-sleep 1st attempt
    * 3:24 PM 7/24/2020 updated to support tenant-alignment & sub'd out showdebug for verbose
    * 11:48 AM 5/27/2020 added func alias:rxo within the func
    * 2:38 PM 4/20/2020 added local $rgxExoPsHostName
    * 8:45 AM 3/3/2020 public cleanup
    * 9:52 PM 1/16/2020 cleanup
    * 1:07 PM 11/25/2019 added *tol/*tor/*cmw alias variants for connect & reconnect
    * 9:52 AM 11/20/2019 spliced in credential matl
    * 2:55 PM 10/11/2018 connect-exo: added pre sleep skip on tryno 1
    * 8:04 AM 11/20/2017 code in a loop in the Reconnect-EXO2old, until it hits or 100% times out
    # 8:47 AM 6/2/2017 cleaned up deadwood, simplified pshelp
    * 7:58 AM 3/15/2017 ren Disconnect/Connect/Reconnect-EXO2old => Disconnect/Connect/Reconnect-EXO2old, added pss, doing tweaks to put into prod use
    * 2/10/14 posted version
    .DESCRIPTION
    I use this for routine test/reconnect of EXO. His orig use was within batches, to break up and requeue chunks of commands.
    Mike's original comment: Below is one
    example of how I batch items for processing and use the
    Reconnect-EXO2old function.  I'm still experimenting with how to best
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
    Reconnect-EXO2old;
    Reconnect EXO connection
    .EXAMPLE
    $batchsize = 10 ;
    $RecordCount=$mr.count #this is the array of whatever you are processing ;
    $b=0 #this is the initialization of a variable used in the do until loop below ;
    $mrs = @() ;
    do {
        Write-Progress -Activity "Getting move request statistics for all $wave move requests." -Status "Processing Records $b through $($b+$batchsize) of $RecordCount." -PercentComplete ($b/$RecordCount*100) ;
        $mrs += $mr | Select-Object -skip $b -first $batchsize | foreach-object {Reconnect-EXO2old; $_ | Get-OLMoveRequestStatistics} ;
        $b=$b+$batchsize ;
        } ;
    until ($b -gt $RecordCount) ;
    
    .LINK
    https://social.technet.microsoft.com/Forums/msonline/en-US/f3292898-9b8c-482a-86f0-3caccc0bd3e5/exchange-powershell-monitoring-remote-sessions?forum=onlineservicesexchange
    #>
    [CmdletBinding()]
    #[Alias('rxo2')]
    Param(
      [Parameter(HelpMessage="Use Proxy-Aware SessionOption settings [-ProxyEnabled]")]
      [boolean]$ProxyEnabled = $False,
      [Parameter(HelpMessage="Credential to use for this connection [-credential [credential obj variable]")]
      [System.Management.Automation.PSCredential]$Credential = $global:credo365TORSID,
      [Parameter(HelpMessage="Debugging Flag [-showDebug]")]
      [switch] $showDebug
    ) ;
    BEGIN{
        $verbose = ($VerbosePreference -eq "Continue") ; 
        if(!$rgxExoPsHostName){$rgxExoPsHostName="^(ps\.outlook\.com|outlook\.office365\.com)$" } ;
        $modname = 'ExchangeOnlineManagement' ; 
        #Try {Get-Module $modname -listavailable -ErrorAction Stop | out-null } Catch {Install-Module $modname -scope CurrentUser ; } ;                 # installed
        Try {Get-Module $modname -ErrorAction Stop | out-null } Catch {Import-Module -Name $modname -MinimumVersion '1.0.1' -ErrorAction Stop -verbose:$false } ; # imported
        
        $TenOrg = get-TenantTag -Credential $Credential ;

    } ;  # BEG-E
    PROCESS{
        $bExistingEXOGood = $false ; 
        if( $legPSSession = Get-PSSession | where-object {$_.ConfigurationName -like "Microsoft.Exchange" -AND $_.Name -match "^(Session|WinRM)\d*" } ){
            # ignore state & Avail, close the conflicting legacy conn's
            Disconnect-Exo ; Disconnect-PssBroken ;Start-Sleep -Seconds 3;
            $bExistingEXOGood = $false ; 
        } ; 
        #clear invalid existing EXOv2 PSS's
        $exov2Broken = Get-PSSession | where-object {$_.ConfigurationName -like "Microsoft.Exchange" -AND $_.Name -eq "ExchangeOnlineInternalSession*" -AND $_.State -like "*Broken*"}
        $exov2Closed = Get-PSSession | where-object {$_.ConfigurationName -like "Microsoft.Exchange" -AND $_.Name -eq "ExchangeOnlineInternalSession*" -AND $_.State -like "*Closed*"}
        
        if ($exov2Broken.count -gt 0){for ($index = 0 ;$index -lt $exov2Broken.count ;$index++){Remove-PSSession -session $exov2Broken[$index]} };
        if ($exov2Closed.count -gt 0){for ($index = 0 ;$index -lt $exov2Closed.count ; $index++){Remove-PSSession -session $exov2Closed[$index] } } ; 
    
        # appears MFA may not properly support passing back a session vari, so go right to strict hostname matches
        $exov2Good = Get-PSSession | where-object {$_.ConfigurationName -like "Microsoft.Exchange" -AND (
            $_.Name -like "ExchangeOnlineInternalSession*") -AND $_.State -like "*Opened*" -AND (
            $_.Availability -eq 'Available')} ; 

        if($exov2Good){
            if( get-command Get-xoAcceptedDomain -ea 0) {
                # add accdom caching
                #$TenOrg = get-TenantTag -Credential $Credential ;
                if(!(Get-Variable  -name "$($TenOrg)Meta").value.o365_AcceptedDomains){
                    set-Variable  -name "$($TenOrg)Meta" -value ( (Get-Variable  -name "$($TenOrg)Meta").value += @{'o365_AcceptedDomains' = (Get-xoAcceptedDomain).domainname} )
                } ;
                #if ((Get-xoAcceptedDomain).domainname.contains($Credential.username.split('@')[1].tostring())){
                if((Get-Variable  -name "$($TenOrg)Meta").value.o365_AcceptedDomains.contains($Credential.username.split('@')[1].tostring())){
                    # validate that the connected EXO is to the $Credential tenant    
                    $smsg = "(EXO Authenticated & Functional(AccDom):$($Credential.username.split('@')[1].tostring()))" ; 
                    $bExistingEXOGood = $true ; 
                # issue: found fresh bug in cxo: svcacct UPN suffix @tenantname.onmicrosoft.com, but testing against AccepteDomain, it's not in there (tho @toroco.mail.onmicrosoft.comis)
                }elseif((Get-Variable  -name "$($TenOrg)Meta").value.o365_TenantDomain -eq ($Credential.username.split('@')[1].tostring())){
                    $smsg = "(EXO Authenticated & Functional(TenDom):$($Credential.username.split('@')[1].tostring()))" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    $bExistingEXOGood = $true ;
                } else { 
                    write-verbose "(Credential mismatch:disconnecting from existing EXO:$($eEXO.Identity) tenant)" ; 
                    DisConnect-EXO2 ; 
                    $bExistingEXOGood = $false ; 
                } ; 
            } else { 
                # capture outlier: shows a session wo the test cmdlet, force reset
                DisConnect-EXO2 ; 
                $bExistingEXOGood = $false ; 
            } ; 
        } ; 

        if($bExistingEXOGood -eq $false){
            connect-exo2 -Credential $Credential -verbose:$($verbose) ; 
        } ; 

    } ;  # PROC-E
    END {
        # if ( (get-module -name tmp_* | ForEach-Object { Get-Command -module $_.name -name 'Add-xoAvailabilityAddressSpace' -ea 0 }) -AND (test-EXOToken) ) {
        if( (Get-PSSession | where-object {$_.ConfigurationName -like "Microsoft.Exchange" -AND $_.Name -like "ExchangeOnlineInternalSession*" -AND $_.State -like "*Opened*" -AND ($_.Availability -eq 'Available')}) -AND (test-EXOToken) ){ 
            # implement caching of accepteddoms into the XXXMeta, in the session (cut back on queries to EXO on acceptedom)
            # non-looping
            
            if( get-command Get-xoAcceptedDomain -ea 0) {
                #$TenOrg = get-TenantTag -Credential $Credential ;
                if(!(Get-Variable  -name "$($TenOrg)Meta").value.o365_AcceptedDomains){
                    set-Variable  -name "$($TenOrg)Meta" -value ( (Get-Variable  -name "$($TenOrg)Meta").value += @{'o365_AcceptedDomains' = (Get-xoAcceptedDomain).domainname} )
                } ;
            } ; 
            <#
            $credDom = ($Credential.username.split("@"))[1] ;
            $Metas=(get-variable *meta|Where-Object{$_.name -match '^\w{3}Meta$'}) ;
            foreach ($Meta in $Metas){
                if( ($credDom -eq $Meta.value.legacyDomain) -OR ($credDom -eq $Meta.value.o365_TenantDomain) -OR ($credDom -eq $Meta.value.o365_OPDomain)){
                    if(!$Meta.value.o365_AcceptedDomains){
                        set-variable -Name $meta.name -Value ((get-variable -name $meta.name).value  += @{'o365_AcceptedDomains' = (Get-xoAcceptedDomain).domainname} )
                    } ; 
                    break ;
                } ;
            } ;
            #>
            #if ((Get-xoAcceptedDomain).domainname.contains($Credential.username.split('@')[1].tostring())){
            #if($Meta.value.o365_AcceptedDomains.contains($Credential.username.split('@')[1].tostring())){
            if((Get-Variable  -name "$($TenOrg)Meta").value.o365_AcceptedDomains.contains($Credential.username.split('@')[1].tostring())){
                # validate that the connected EXO is to the $Credential tenant    
                write-verbose "(EXOv2 Authenticated & Functional:$($Credential.username.split('@')[1].tostring()))" ; 
                $bExistingEXOGood = $true ; 
            # issue: found fresh bug in cxo: svcacct UPN suffix @tenantname.onmicrosoft.com, but testing against AccepteDomain, it's not in there (tho @toroco.mail.onmicrosoft.comis)
            }elseif((Get-Variable  -name "$($TenOrg)Meta").value.o365_TenantDomain -eq ($Credential.username.split('@')[1].tostring())){
                $smsg = "(EXO Authenticated & Functional(TenDom):$($Credential.username.split('@')[1].tostring()))" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                $bExistingEXOGood = $true ;
            } else { 
                write-error "(Credential mismatch:disconnecting from existing EXO:$($eEXO.Identity) tenant)" ; 
                Disconnect-exo2 ; 
                $bExistingEXOGood = $false ; 
            } ;
        } ; 
    } ; # END-E 
}

#*------^ Reconnect-EXO2old.ps1 ^------

#*------v RemoveExistingEXOPSSession.ps1 v------
function RemoveExistingEXOPSSession() {
    <#
    .SYNOPSIS
    RemoveExistingEXOPSSession.ps1 - Tweaked version of the Exchangeonline module:RemoveExistingPSSession() to avoid purging CCMW sessions on connect. Intent is to permit *concurrent* EXO & CCMS sessions.
    .NOTES
    Version     : 1.0.0
    Author      : Todd Kadrie
    Website     :	http://www.toddomation.com
    Twitter     :	@tostka / http://twitter.com/tostka
    CreatedDate : 20201109-0833AM
    FileName    : RemoveExistingEXOPSSession.ps1
    License     : [none specified]
    Copyright   : [none specified]
    Github      : https://github.com/tostka/verb-exo
    Tags        : Powershell
    AddedCredit : Microsoft (edited version of published commands in the module)
    AddedWebsite:	https://docs.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2
    REVISIONS
    * 8:34 AM 11/9/2020 init
    .DESCRIPTION
    RemoveExistingEXOPSSession.ps1 - Tweaked version of the Exchangeonline module:RemoveExistingPSSession() to avoid purging CCMW sessions on connect. Intent is to permit *concurrent* EXO & CCMS sessions.
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    None. Returns no objects or output.
    .EXAMPLE
    RemoveExistingEXOPSSession
    Stock call
    .LINK
    https://github.com/tostka/verb-EXO
    .LINK
    https://docs.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2
    #>
    [CmdletBinding()]
    param()
    
    #$existingPSSession = Get-PSSession | Where-Object {$_.ConfigurationName -like "Microsoft.Exchange" -and $_.Name -like "ExchangeOnlineInternalSession*"}
    <# filter *ONLY* EXO sessions, exclude CCMS, they differ on ComputerName endpoint:
    #-=EXO-=-=-=-=-=-=-=
    ConfigurationName : Microsoft.Exchange
    ComputerName      : outlook.office365.com
    Name              : ExchangeOnlineInternalSession_2
    #-=CCMS-=-=-=-=-=-=-=
    ConfigurationName : Microsoft.Exchange
    ComputerName      : nam02b.ps.compliance.protection.outlook.com
    Name              : ExchangeOnlineInternalSession_1
    #-=-=-=-=-=-=-=-=
    #>
    $rgxExoPsHostName = "^(ps\.outlook\.com|outlook\.office365\.com)$"
    $existingPSSession = Get-PSSession | Where-Object {$_.ConfigurationName -like "Microsoft.Exchange" -and $_.Name -like "ExchangeOnlineInternalSession*" -AND $_.ComputerName -match$rgxExoPsHostName} ; 

        if ($existingPSSession.count -gt 0) 
        {
            for ($index = 0; $index -lt $existingPSSession.count; $index++)
            {
                $session = $existingPSSession[$index]
                Remove-PSSession -session $session

                Write-Host "Removed the PSSession $($session.Name) connected to $($session.ComputerName)"
            }
        }

        # Clear any left over PS tmp modules
        if ($global:_EXO_PreviousModuleName -ne $null)
        {
            Remove-Module -Name $global:_EXO_PreviousModuleName -ErrorAction SilentlyContinue
            $global:_EXO_PreviousModuleName = $null
        }
    }

#*------^ RemoveExistingEXOPSSession.ps1 ^------

#*------v RemoveExistingPSSessionTargeted.ps1 v------
function RemoveExistingPSSessionTargeted() {
    <#
    .SYNOPSIS
    RemoveExistingPSSessionTargeted.ps1 - Tweaked version of the Exchangeonline module:RemoveExistingPSSession() to avoid purging CCMW sessions on connect. Intent is to permit *concurrent* EXO & CCMS sessions.
    .NOTES
    Version     : 1.0.0
    Author      : Todd Kadrie
    Website     :	http://www.toddomation.com
    Twitter     :	@tostka / http://twitter.com/tostka
    CreatedDate : 20201109-0833AM
    FileName    : RemoveExistingPSSessionTargeted.ps1
    License     : [none specified]
    Copyright   : [none specified]
    Github      : https://github.com/tostka/verb-exo
    Tags        : Powershell
    AddedCredit : Microsoft (edited version of published commands in the module)
    AddedWebsite:	https://docs.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2
    REVISIONS
    * 8:34 AM 11/9/2020 init
    .DESCRIPTION
    RemoveExistingPSSessionTargeted.ps1 - Tweaked version of the Exchangeonline module:RemoveExistingPSSession() to avoid purging CCMW sessions on connect. Intent is to permit *concurrent* EXO & CCMS sessions.
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    None. Returns no objects or output.
    .EXAMPLE
    RemoveExistingPSSessionTargeted
    Stock call
    .LINK
    https://github.com/tostka/verb-EXO
    .LINK
    https://docs.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2
    #>
    [CmdletBinding()]
    param()
    
    #$existingPSSession = Get-PSSession | Where-Object {$_.ConfigurationName -like "Microsoft.Exchange" -and $_.Name -like "ExchangeOnlineInternalSession*"}
    <# filter *ONLY* EXO sessions, exclude CCMS, they differ on ComputerName endpoint:
    #-=EXO-=-=-=-=-=-=-=
    ConfigurationName : Microsoft.Exchange
    ComputerName      : outlook.office365.com
    Name              : ExchangeOnlineInternalSession_2
    #-=CCMS-=-=-=-=-=-=-=
    ConfigurationName : Microsoft.Exchange
    ComputerName      : nam02b.ps.compliance.protection.outlook.com
    Name              : ExchangeOnlineInternalSession_1
    #-=-=-=-=-=-=-=-=
    #>
    $rgxExoPsHostName = "^(ps\.outlook\.com|outlook\.office365\.com)$"
    $existingPSSession = Get-PSSession | Where-Object {$_.ConfigurationName -like "Microsoft.Exchange" -and $_.Name -like "ExchangeOnlineInternalSession*" -AND $_.ComputerName -match$rgxExoPsHostName} ; 

        if ($existingPSSession.count -gt 0) 
        {
            for ($index = 0; $index -lt $existingPSSession.count; $index++)
            {
                $session = $existingPSSession[$index]
                Remove-PSSession -session $session

                Write-Host "Removed the PSSession $($session.Name) connected to $($session.ComputerName)"
            }
        }

        # Clear any left over PS tmp modules
        if ($global:_EXO_PreviousModuleName -ne $null)
        {
            Remove-Module -Name $global:_EXO_PreviousModuleName -ErrorAction SilentlyContinue
            $global:_EXO_PreviousModuleName = $null
        }
    }

#*------^ RemoveExistingPSSessionTargeted.ps1 ^------

#*------v Remove-EXOBrokenClosed.ps1 v------
function Remove-EXOBrokenClosed(){
    <#
    .SYNOPSIS
    Remove-EXOBrokenClosed - Remove broken and closed exchange online PSSessions
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
    AddedCredit : 
    AddedWebsite:	https://social.technet.microsoft.com/Forums/msonline/en-US/f3292898-9b8c-482a-86f0-3caccc0bd3e5/exchange-powershell-monitoring-remote-sessions?forum=onlineservicesexchange
    AddedTwitter:	
    REVISIONS   :
    * 9:29 AM 7/30/2020 lifted from EXO V2 connect-exchangeonline() as RemoveBrokenOrClosedPSSession()
    * 3:24 PM 7/24/2020 updated to support tenant-alignment & sub'd out showdebug for verbose
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
    Remove-EXOBrokenClosed;
    .LINK
    https://social.technet.microsoft.com/Forums/msonline/en-US/f3292898-9b8c-482a-86f0-3caccc0bd3e5/exchange-powershell-monitoring-remote-sessions?forum=onlineservicesexchange
    #>
    [CmdletBinding()]
    [Alias('dxob')]
    $exov2Broken = Get-PSSession | where-object {$_.ConfigurationName -like "Microsoft.Exchange" -and $_.Name -eq "ExchangeOnlineInternalSession*" -and $_.State -like "*Broken*"} ;
    $exov2Closed = Get-PSSession | where-object {$_.ConfigurationName -like "Microsoft.Exchange" -and $_.Name -eq "ExchangeOnlineInternalSession*" -and $_.State -like "*Closed*"} ;
    if ($exov2Broken.count -gt 0){for ($index = 0; $index -lt $exov2Broken.count; $index++) {Remove-PSSession -session $exov2Broken[$index] } } ;
    if ($exov2Closed.count -gt 0){for ($index = 0; $index -lt $exov2Closed.count; $index++) {Remove-PSSession -session $exov2Closed[$index] } } ;
}

#*------^ Remove-EXOBrokenClosed.ps1 ^------

#*------v rxo2cmw.ps1 v------
function rxo2CMW {
    <#
    .SYNOPSIS
    rxo2CMW - Reonnect-EXO2 to specified Tenant
    .DESCRIPTION
    Reconnect-EXO2 - Re-establish PSS to EXO V2 Modern Auth
    .EXAMPLE
    rxo2CMW
    #>
    Reconnect-EXO2 -cred $credO365CMWCSID -Verbose:($VerbosePreference -eq 'Continue') ; 
}

#*------^ rxo2cmw.ps1 ^------

#*------v rxo2tol.ps1 v------
function rxo2TOL {
    <#
    .SYNOPSIS
    rxo2TOL - Reonnect-EXO to specified Tenant
    .DESCRIPTION
    #>
    Reconnect-EXO2 -cred $credO365TOLSID
}

#*------^ rxo2tol.ps1 ^------

#*------v rxo2tor.ps1 v------
function rxo2TOR {
    <#
    .SYNOPSIS
    rxo2TOR - Reonnect-EXO to specified Tenant
    .DESCRIPTION
    Reconnect-EXO2 - Re-establish PSS to EXO V2 Modern Auth
    .EXAMPLE
    rxo2TOR
    #>
    Reconnect-EXO2 -cred $credO365TORSID -Verbose:($VerbosePreference -eq 'Continue') ; 
}

#*------^ rxo2tor.ps1 ^------

#*------v rxo2ven.ps1 v------
function rxo2VEN {
    <#
    .SYNOPSIS
    rxo2VEN - Reonnect-EXO to specified Tenant
    .DESCRIPTION
    Reconnect-EXO2 - Re-establish PSS to EXO V2 Modern Auth
    #>
    Reconnect-EXO2 -cred $credO365VENCSID -Verbose:($VerbosePreference -eq 'Continue') ; 
}

#*------^ rxo2ven.ps1 ^------

#*------v rxocmw.ps1 v------
function rxoCMW {
    <#
    .SYNOPSIS
    rxoCMW - Reonnect-EXO to specified Tenant
    .DESCRIPTION
    Reconnect-EXO - Re-establish PSS to https://ps.outlook.com/powershell/
    .EXAMPLE
    rxoCMW
    #>
    Reconnect-EXO -cred $credO365CMWCSID -Verbose:($VerbosePreference -eq 'Continue') ; 
}

#*------^ rxocmw.ps1 ^------

#*------v rxotol.ps1 v------
function rxoTOL {
    <#
    .SYNOPSIS
    rxoTOL - Reonnect-EXO to specified Tenant
    .DESCRIPTION
    Reconnect-EXO - Re-establish PSS to https://ps.outlook.com/powershell/
    .EXAMPLE
    rxoTOL
    #>
    Reconnect-EXO -cred $credO365TOLSID -Verbose:($VerbosePreference -eq 'Continue') ; 
}

#*------^ rxotol.ps1 ^------

#*------v rxotor.ps1 v------
function rxoTOR {
    <#
    .SYNOPSIS
    rxoTOR - Reonnect-EXO to specified Tenant
    .DESCRIPTION
    Reconnect-EXO - Re-establish PSS to https://ps.outlook.com/powershell/
    .EXAMPLE
    rxoTOR
    #>
    Reconnect-EXO -cred $credO365TORSID -Verbose:($VerbosePreference -eq 'Continue') ; 
}

#*------^ rxotor.ps1 ^------

#*------v rxoVEN.ps1 v------
function rxoVEN {
    <#
    .SYNOPSIS
    rxoVEN - Reonnect-EXO to specified Tenant
    .DESCRIPTION
    Reconnect-EXO - Re-establish PSS to https://ps.outlook.com/powershell/
    .EXAMPLE
    rxoVEN
    #>
    Reconnect-EXO -cred $credO365VENCSID -Verbose:($VerbosePreference -eq 'Continue') ; 
}

#*------^ rxoVEN.ps1 ^------

#*------v test-EXOToken.ps1 v------
function test-EXOToken {
    <#
    .SYNOPSIS
    test-EXOToken - Retrieve and summarize EXOv2 OAuth Active Token (leverages ExchangeOnlineManagement 'Microsoft.Exchange.Management.ExoPowershellGalleryModule.dll', OAuth isn't used for EXO legacy basic-auth connections)
    .NOTES
    Version     : 1.0.0.0
    Author      : Todd Kadrie
    Website     :	http://www.toddomation.com
    Twitter     :	@tostka / http://twitter.com/tostka
    CreatedDate : 2020-08-08
    FileName    : test-EXOToken
    License     : MIT License
    Copyright   : (c) 2020 Todd Kadrie
    Github      : https://github.com/tostka/verb-aad
    Tags        : Powershell,ExchangeOnline,Exchange,RemotePowershell,Connection,MFA
    REVISIONS
    # 8:34 AM 3/31/2021 added verbose suppress to all import-mods
    * 12:21 PM 8/11/2020 added dependancy mod try/tach, and a catch on the failure error returned by the underlying test-ActiveToken cmd
    * 11:58 AM 8/9/2020 init
    .DESCRIPTION
    test-EXOToken - Retrieve and summarize EXOv2 OAuth Active Token (leverages ExchangeOnlineManagement 'Microsoft.Exchange.Management.ExoPowershellGalleryModule.dll', OAuth isn't used for EXO legacy basic-auth connections)
    Trying to find a way to verify status of token, wo any interactive EXO traffic. Lifted concept from EXOM UpdateImplicitRemotingHandler().
    Test-ActiveToken doesn't appear to normally be exposed anywhere but with explicit load of the .dll
    .OUTPUT
    System.Boolean
    .EXAMPLE
    $hasActiveToken = test-EXOToken 
    $psss=Get-PSSession | where-object {$_.ConfigurationName -like "Microsoft.Exchange" -AND $_.Name -like "ExchangeOnlineInternalSession*" } ;  
    $sessionIsOpened = $psss.Runspace.RunspaceStateInfo.State -eq 'Opened'
    if (($hasActiveToken -eq $false) -or ($sessionIsOpened -ne $true)){
        #If there is no active user token or opened session then ensure that we remove the old session
        $shouldRemoveCurrentSession = $true;
    } ; 
    Retrieve and evaluate status of EXO user token against PSSessoin status for EXOv2
    .LINK
    https://github.com/tostka/verb-aad
    #>
    #Requires -Modules ExchangeOnlineManagement
    [CmdletBinding()] 
    Param() ;
    BEGIN {$verbose = ($VerbosePreference -eq "Continue") } ;
    PROCESS {
        $hasActiveToken = $false ; 
        # Save time and pretest for *any* EXOv2 PSSession, before bothering to test (no session - even closed/broken => no OAuth token)
        $exov2 = Get-PSSession | where-object {$_.ConfigurationName -like "Microsoft.Exchange" -AND $_.Name -like "ExchangeOnlineInternalSession*"} ; 
        if($exov2){
        
            # ==load dependancy module:
            # admin/SID module auto-install code (myBoxes UID split-perm CU, all else t AllUsers)
            $modname = 'ExchangeOnlineManagement' ;
            $minvers = '1.0.1' ; 
            Try {Get-Module -name $modname -listavailable -ErrorAction Stop | out-null } Catch {
                $pltInMod=[ordered]@{Name=$modname} ; 
                if( $env:COMPUTERNAME -match $rgxMyBoxUID ){$pltInMod.add('scope','CurrentUser')} else {$pltInMod.add('scope','AllUsers')} ;
                write-host -foregroundcolor YELLOW "$((get-date).ToString('HH:mm:ss')):Install-Module w scope:$($pltInMod.scope)`n$(($pltInMod|out-string).trim())" ; 
                Install-Module @pltIMod ; 
            } ; # IsInstalled
            $pltIMod = @{Name = $modname ; ErrorAction = 'Stop' ; verbose=$false } ;
            if($minvers){$pltIMod.add('MinimumVersion',$minvers) } ; 
            Try { Get-Module $modname -ErrorAction Stop | out-null } Catch {
                write-verbose "Import-Module w`n$(($pltIMod|out-string).trim())" ; 
                Import-Module @pltIMod ; 
            } ; # IsImported
      
            $error.clear() ;
            TRY {
                #=load function module (subcomponent of dep module, pathed from same dir)
                $tmodpath = join-path -path (split-path (get-module $modname -list).path) -ChildPath 'Microsoft.Exchange.Management.ExoPowershellGalleryModule.dll' ;
                if(test-path $tmodpath){ import-module -name $tmodpath -Cmdlet Test-ActiveToken -verbose:$false }
                else { throw "Unable to locate:Microsoft.Exchange.Management.ExoPowershellGalleryModule.dll" } ;  
            } CATCH {
                Write-Warning "$(get-date -format 'HH:mm:ss'): Failed processing $($_.Exception.ItemName). `nError Message: $($_.Exception.Message)`nError Details: $($_)" ;
                Break #Opts: STOP(debug)|EXIT(close)|CONTINUE(move on in loop cycle)|BREAK(exit loop iteration)|THROW $_/'CustomMsg'(end script with Err output)
            } ; 
        
            if(gcm -name Test-ActiveToken){
                $error.clear() ;
                TRY {
                    $hasActiveToken = Test-ActiveToken ; 
                } CATCH [System.Management.Automation.RuntimeException] {
                    # reflects: test-activetoken : Object reference not set to an instance of an object.
                    write-verbose "Token not present"
                } CATCH {
                    Write-Warning "$(get-date -format 'HH:mm:ss'): Failed processing $($_.Exception.ItemName). `nError Message: $($_.Exception.Message)`nError Details: $($_)" ;
                    Break #Opts: STOP(debug)|EXIT(close)|CONTINUE(move on in loop cycle)|BREAK(exit loop iteration)|THROW $_/'CustomMsg'(end script with Err output)
                } ;  
            } else { throw "missing:gcm -name Test-ActiveToken" } 

        } else { 
            write-verbose "No Token: No existing EXOv2 PSSession (ConfigurationName -like 'Microsoft.Exchange' -AND Name -like 'ExchangeOnlineInternalSession*')" ; 
        } ; 
    } ; 
    END{ $hasActiveToken | write-output } ;
}

#*------^ test-EXOToken.ps1 ^------

#*------v Test-Uri.ps1 v------
function Test-Uri{
    <#
    .SYNOPSIS
    Test-Uri.ps1 - Validates a given Uri ; localized verb-EXO vers of non-'$global:' helper funct from ExchangeOnlineManagement. The globals export fine, these don't and appear to need to be loaded manually
    .NOTES
    Version     : 1.0.0
    Author      : Todd Kadrie
    Website     :	http://www.toddomation.com
    Twitter     :	@tostka / http://twitter.com/tostka
    CreatedDate : 20201109-0833AM
    FileName    : Test-Uri.ps1
    License     : [none specified]
    Copyright   : [none specified]
    Github      : https://github.com/tostka/verb-exo
    Tags        : Powershell
    AddedCredit : Microsoft (edited version of published commands in the module)
    AddedWebsite:	https://docs.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2
    REVISIONS
    * 8:34 AM 11/9/2020 init
    .DESCRIPTION
    Test-Uri.ps1 - localized verb-EXO vers of non-'$global:' helper funct from ExchangeOnlineManagement. The globals export fine, these don't and appear to need to be loaded manually
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    None. Returns no objects or output.
    .EXAMPLE
    Test-Uri
    Stock call
    .LINK
    https://github.com/tostka/verb-EXO
    .LINK
    https://docs.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    Param
    (
        # Uri to be validated
        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true, Position=0)]
        [string]
        $UriString
    )

    [Uri]$uri = $UriString -as [Uri]

    $uri.AbsoluteUri -ne $null -and $uri.Scheme -eq 'https'
}

#*------^ Test-Uri.ps1 ^------

#*======^ END FUNCTIONS ^======

Export-ModuleMember -Function check-EXOLegalHold,Connect-ExchangeOnlineTargetedPurge,Connect-EXO,Connect-EXO2,connect-EXO2old,Connect-EXOPSSession,connect-EXOv2RAW,Connect-IPPSSessionTargetedPurge,cxo2cmw,cxo2TOL,cxo2TOR,cxo2VEN,cxoCMW,cxoTOL,cxoTOR,cxoVEN,Disconnect-ExchangeOnline,Disconnect-EXO,Disconnect-EXO2,get-MailboxFolderStats,get-MsgTrace,Get-OrgNameFromUPN,Invoke-ExoOnlineConnection,Print-Details,Reconnect-EXO,Reconnect-EXO2,Reconnect-EXO2old,RemoveExistingEXOPSSession,RemoveExistingPSSessionTargeted,Remove-EXOBrokenClosed,rxo2CMW,rxo2TOL,rxo2TOR,rxo2VEN,rxoCMW,rxoTOL,rxoTOR,rxoVEN,test-EXOToken,Test-Uri -Alias *


# SIG # Begin signature block
# MIIELgYJKoZIhvcNAQcCoIIEHzCCBBsCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUZKQ1UivaN34k+y0noMeC1Xw6
# 9IagggI4MIICNDCCAaGgAwIBAgIQWsnStFUuSIVNR8uhNSlE6TAJBgUrDgMCHQUA
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
# CisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMCMGCSqGSIb3DQEJBDEWBBQLsDdz
# Xg3ihNw6129nWkPinwPEojANBgkqhkiG9w0BAQEFAASBgHx+cRKJNhXlclrUgsmD
# dZLzOqO/bBhphXJGzsMUM+ocR42Ho8MJnVUWbMjYubhm8TAX8cHhp7ASgXItsftq
# cEwrLzSWxPuQmqERbUZ/UvzntXeOuQPxfdQE/84qskl31BwUy2z2HHcZEiBB3NbY
# gJBYZlwzDyHr9+cVvl6Pxb8H
# SIG # End signature block
