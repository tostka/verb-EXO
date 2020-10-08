﻿# verb-exo.psm1


  <#
  .SYNOPSIS
  verb-EXO - Powershell Exchange Online generic functions module
  .NOTES
  Version     : 1.0.32.0
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

        # if we're using EXOv1-style BasicAuth, clear incompatible existing EXOv2 PSS's
        $exov2Good = Get-PSSession | where-object {$_.ConfigurationName -like "Microsoft.Exchange" -and $_.Name -like "ExchangeOnlineInternalSession*" -and $_.State -like "*Opened*" -AND ($_.Availability -eq 'Available')} ; 
        $exov2Broken = Get-PSSession | where-object {$_.ConfigurationName -like "Microsoft.Exchange" -and $_.Name -eq "ExchangeOnlineInternalSession*" -and $_.State -like "*Broken*"}
        $exov2Closed = Get-PSSession | where-object {$_.ConfigurationName -like "Microsoft.Exchange" -and $_.Name -eq "ExchangeOnlineInternalSession*" -and $_.State -like "*Closed*"}

        if($exov2Good  ){
            write-verbose "EXOv1:Disconnecting conflicting EXOv2 connection" ; 
            DisConnect-EXO2 ; 
        } ; 
        if ($exov2Broken.count -gt 0){for ($index = 0 ;$index -lt $exov2Broken.count ;$index++){Remove-PSSession -session $exov2Broken[$index]} };
        if ($exov2Closed.count -gt 0){for ($index = 0 ;$index -lt $exov2Closed.count ; $index++){Remove-PSSession -session $exov2Closed[$index] } } ; 
    
        $bExistingEXOGood = $false ; 
        # $existingPSSession = Get-PSSession | where-object { $_.ConfigurationName -like "Microsoft.Exchange" -and $_.Name -match "^(Session|WinRM)\d*" }
        #if( Get-PSSession|Where-Object{($_.ComputerName -match $rgxExoPsHostName) -AND ($_.State -eq 'Opened') -AND ($_.Availability -eq 'Available')}){
        # EXOv1 & v2 both use ComputerName -match $rgxExoPsHostName, need to use the distinctive differentiators instead
        if(Get-PSSession | where-object { $_.ConfigurationName -like "Microsoft.Exchange" -AND $_.Name -match "^(Session|WinRM)\d*" -AND $_.State -eq 'Opened' -AND $_.Availability -eq 'Available' }){
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
                        Break #STOP(debug)|EXIT(close)|Continue(move on in loop cycle) ; 
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
            #if ((Get-exoAcceptedDomain).domainname.contains($Credential.username.split('@')[1].tostring())){
            # do caching & check cached value, not qry unless unpopulated (first pass in global session)
            if($Meta.value.o365_AcceptedDomains.contains($Credential.username.split('@')[1].tostring())){
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
    # 10:35 AM 6/20/2019 added $pltPSS splat dump to the import-pssession cmd block; hard-typed the $Credential [System.Management.Automation.PSCredential]
    # 8:22 AM 11/20/2017 spliced in retry loop into reConnect-EXO2 as well, to avoid using any state testing in scripts, localize it 1x here.
    # 1:49 PM 11/2/2017 coded around non-profile gaps from not having get-admincred() - added the prompt code in to fake it
    # 12:26 PM 9/11/2017 debugged retry - catch doesn't fire properly on new-Pssession, have to test the $error state, to detect auth fails (assuming the bad pw error# is specific). $error test is currently targeting specific error returned on a bad password. Added retry, for when connection won't hold and fails breaks - need to watch out that bad pw doesn't lock out the acct!
    # 12:50 PM 6/2/2017 expanded pshelp, added Examples, cleaned up deadwood
    * # 12:10 PM 3/15/2017 Connect-EXO2 typo, disable prefix auto spec, unless actually blanked. switch ProxyEnabled to non-Mandatory.
    .DESCRIPTION
    Connect-EXO2 - Establish PSS to EXO V2 Modern Auth
    .PARAMETER  CommandPrefix
    [verb]-PREFIX[command] PREFIX string for clearly marking cmdlets sourced in this connection [-CommandPrefix tag]
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
    #>
    [CmdletBinding()]
    [Alias('cxo2')]
    Param(
        [Parameter(HelpMessage = "[verb]-PREFIX[command] PREFIX string for clearly marking cmdlets sourced in this connection [-CommandPrefix tag]")]
        [string]$CommandPrefix = 'xo',
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
        if (!$CommandPrefix) {
            $CommandPrefix = 'xo' ; # 4:31 PM 7/29/2020 MS has RESERVED use of the 'exo' prefix [facepalm]
            write-verbose -verbose:$true  "(asserting Prefix:$($CommandPrefix)" ;
        } ;

        $sTitleBarTag = "EXO" ;
        $TentantTag = get-TenantTag -Credential $Credential ;
        if ($TentantTag -ne 'TOR') {
            # explicitly leave this tenant (default) untagged
            $sTitleBarTag += $TentantTag ;
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
            if ( get-command Get-xoAcceptedDomain -ea 0) {
                if ((Get-xoAcceptedDomain).domainname.contains($Credential.username.split('@')[1].tostring())) {
                    # validate that the connected EXO is to the $Credential tenant
                    write-verbose "(Existing EXO Authenticated & Functional:$($Credential.username.split('@')[1].tostring()))" ;
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

            #Connect-ExchangeOnline -Credential $credO365TORSID -Prefix 'xo' -ShowBanner:$false ;
            # *** LEFT OFF HERE 5:01 PM 7/29/2020 *** not sure if it supports allowclobber, if it's actually wrapping pssession, it sure as shit does!

            $pltCXO = @{
                Prefix     = [string]$CommandPrefix ;
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
            # verify the exov2 cmdlets actually imported as a tmp_ module w specifid prefix & 1st cmdlet:if(get-module -name tmp_* |%{gcm -module $_.name -name 'Add-xoAvailabilityAddressSpace' -ea 0 }){'Y'}else {'N'}
            if ( (get-module -name tmp_* | ForEach-Object { Get-Command -module $_.name -name 'Add-xoAvailabilityAddressSpace' -ea 0 }) -AND (test-EXOToken) ) {
                $bExistingEXOGood = $true ;
            } else { $bExistingEXOGood = $false ; }
            # implement caching of accepteddoms into the XXXMeta, in the session (cut back on queries to EXO on acceptedom)
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
            if($Meta.value.o365_AcceptedDomains.contains($Credential.username.split('@')[1].tostring())){
                # validate that the connected EXO is to the $Credential tenant
                write-verbose "(EXO Authenticated & Functional:$($Credential.username.split('@')[1].tostring()))" ;
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

#*------^ Connect-EXO2.ps1 ^------

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
    if($Global:EOLModule){$Global:EOLModule | Remove-Module -Force ; } ;
    if($global:EOLSession){$global:EOLSession | Remove-PSSession ; } ;
    Get-PSSession |Where-Object{$_.ComputerName -match $rgxExoPsHostName } | Remove-PSSession ;
    Disconnect-PssBroken -verbose:$($verbose) ;
    Remove-PSTitlebar 'EXO' ;
    
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
    Try {Get-Module $modname -ErrorAction Stop | out-null } Catch {Import-Module -Name $modname -MinimumVersion '1.0.1' -ErrorAction Stop  } ; # imported
    # just alias disconnect-ExchangeOnline, it retires token etc as well as closing PSS, but biggest reason is it's got a confirm, hard-coded, needs a function to override
    Disconnect-ExchangeOnline -confirm:$false ; 
    Disconnect-PssBroken -verbose:$($verbose) ;
    Remove-PSTitlebar 'EXO' ;
}

#*------^ Disconnect-EXO2.ps1 ^------

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

    # if we're using EXOv1-style BasicAuth, clear incompatible existing EXOv2 PSS's
    $exov2Good = Get-PSSession | where-object {$_.ConfigurationName -like "Microsoft.Exchange" -and $_.Name -like "ExchangeOnlineInternalSession*" -and $_.State -like "*Opened*" -AND ($_.Availability -eq 'Available')} ; 
    $exov2Broken = Get-PSSession | where-object {$_.ConfigurationName -like "Microsoft.Exchange" -and $_.Name -eq "ExchangeOnlineInternalSession*" -and $_.State -like "*Broken*"}
    $exov2Closed = Get-PSSession | where-object {$_.ConfigurationName -like "Microsoft.Exchange" -and $_.Name -eq "ExchangeOnlineInternalSession*" -and $_.State -like "*Closed*"}

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

        $legPSSession = Get-PSSession | where-object {$_.ConfigurationName -like "Microsoft.Exchange" -and $_.Name -match "^(Session|WinRM)\d*" -AND ($_.State -eq 'Opened') -AND ($_.Availability -eq 'Available')}
        
        if( (Get-PSSession | where-object {$_.ConfigurationName -like "Microsoft.Exchange" -and $_.Name -match "^(Session|WinRM)\d*" -AND (($_.State -ne 'Opened') -OR ($_.Availability -ne 'Available')) }) -OR (-not(Get-PSSession | where-object {$_.ConfigurationName -like "Microsoft.Exchange" -and $_.Name -match "^(Session|WinRM)\d*"})) ){
            write-verbose "$((get-date).ToString('HH:mm:ss')):Reconnecting:No existing PSSESSION matching Name -match (Session|WinRM) with valid Open/Availability:$((Get-PSSession|Where-Object{$_.ComputerName -match $rgxExoPsHostName}| Format-Table -a State,Availability |out-string).trim())" ;
            Disconnect-Exo ; Disconnect-PssBroken ;Start-Sleep -Seconds 3;
            if(!$Credential){
                connect-EXO ;
            } else {
                connect-EXO -credential:$($Credential) ;
            } ;
        
        }elseif($legPSSession){
            # implement caching of accepteddoms into the XXXMeta, in the session (cut back on queries to EXO on acceptedom)
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
            #if((Get-exoAcceptedDomain).domainname.contains($Credential.username.split('@')[1].tostring())){
            # do caching & check cached value, not qry unless unpopulated (first pass in global session)
            if($Meta.value.o365_AcceptedDomains.contains($Credential.username.split('@')[1].tostring())){
                # validate that the connected EXO is to the $Credential tenant    
                write-verbose "(Authenticated to EXO:$($Credential.username.split('@')[1].tostring()))" ; 
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
        Try {Get-Module $modname -ErrorAction Stop | out-null } Catch {Import-Module -Name $modname -MinimumVersion '1.0.1' -ErrorAction Stop  } ; # imported

    } ;  # BEG-E
    PROCESS{
        $bExistingEXOGood = $false ; 
        if( $legPSSession = Get-PSSession | where-object {$_.ConfigurationName -like "Microsoft.Exchange" -and $_.Name -match "^(Session|WinRM)\d*" } ){
            # ignore state & Avail, close the conflicting legacy conn's
            Disconnect-Exo ; Disconnect-PssBroken ;Start-Sleep -Seconds 3;
            $bExistingEXOGood = $false ; 
        } ; 
        #clear invalid existing EXOv2 PSS's
        $exov2Broken = Get-PSSession | where-object {$_.ConfigurationName -like "Microsoft.Exchange" -and $_.Name -eq "ExchangeOnlineInternalSession*" -and $_.State -like "*Broken*"}
        $exov2Closed = Get-PSSession | where-object {$_.ConfigurationName -like "Microsoft.Exchange" -and $_.Name -eq "ExchangeOnlineInternalSession*" -and $_.State -like "*Closed*"}
        
        if ($exov2Broken.count -gt 0){for ($index = 0 ;$index -lt $exov2Broken.count ;$index++){Remove-PSSession -session $exov2Broken[$index]} };
        if ($exov2Closed.count -gt 0){for ($index = 0 ;$index -lt $exov2Closed.count ; $index++){Remove-PSSession -session $exov2Closed[$index] } } ; 
    
        # appears MFA may not properly support passing back a session vari, so go right to strict hostname matches
        $exov2Good = Get-PSSession | where-object {$_.ConfigurationName -like "Microsoft.Exchange" -and $_.Name -like "ExchangeOnlineInternalSession*" -and $_.State -like "*Opened*" -AND ($_.Availability -eq 'Available')} ; 

        if($exov2Good){
            if( get-command Get-xoAcceptedDomain -ea 0) {
                if ((Get-xoAcceptedDomain).domainname.contains($Credential.username.split('@')[1].tostring())){
                    # validate that the connected EXO is to the $Credential tenant    
                    write-verbose "(Existing EXO Authenticated & Functional:$($Credential.username.split('@')[1].tostring()))" ; 
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
        if( (Get-PSSession | where-object {$_.ConfigurationName -like "Microsoft.Exchange" -and $_.Name -like "ExchangeOnlineInternalSession*" -and $_.State -like "*Opened*" -AND ($_.Availability -eq 'Available')}) -AND (test-EXOToken) ){ 
            # implement caching of accepteddoms into the XXXMeta, in the session (cut back on queries to EXO on acceptedom)
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
            #if ((Get-xoAcceptedDomain).domainname.contains($Credential.username.split('@')[1].tostring())){
            if($Meta.value.o365_AcceptedDomains.contains($Credential.username.split('@')[1].tostring())){
                # validate that the connected EXO is to the $Credential tenant    
                write-verbose "(EXOv2 Authenticated & Functional:$($Credential.username.split('@')[1].tostring()))" ; 
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
    $psss=Get-PSSession | where-object {$_.ConfigurationName -like "Microsoft.Exchange" -and $_.Name -like "ExchangeOnlineInternalSession*" } ;  
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
        $exov2 = Get-PSSession | where-object {$_.ConfigurationName -like "Microsoft.Exchange" -and $_.Name -like "ExchangeOnlineInternalSession*"} ; 
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
            $pltIMod = @{Name = $modname ; ErrorAction = 'Stop' ; } ;
            if($minvers){$pltIMod.add('MinimumVersion',$minvers) } ; 
            Try { Get-Module $modname -ErrorAction Stop | out-null } Catch {
                write-verbose "Import-Module w`n$(($pltIMod|out-string).trim())" ; 
                Import-Module @pltIMod ; 
            } ; # IsImported
      
            $error.clear() ;
            TRY {
                #=load function module (subcomponent of dep module, pathed from same dir)
                $tmodpath = join-path -path (split-path (get-module $modname -list).path) -ChildPath 'Microsoft.Exchange.Management.ExoPowershellGalleryModule.dll' ;
                if(test-path $tmodpath){ import-module -name $tmodpath -Cmdlet Test-ActiveToken }
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

#*======^ END FUNCTIONS ^======

Export-ModuleMember -Function Connect-EXO,Connect-EXO2,cxo2cmw,cxo2TOL,cxo2TOR,cxo2VEN,cxoCMW,cxoTOL,cxoTOR,cxoVEN,Disconnect-EXO,Disconnect-EXO2,Reconnect-EXO,Reconnect-EXO2,Remove-EXOBrokenClosed,rxo2CMW,rxo2TOL,rxo2TOR,rxo2VEN,rxoCMW,rxoTOL,rxoTOR,rxoVEN,test-EXOToken -Alias *


# SIG # Begin signature block
# MIIELgYJKoZIhvcNAQcCoIIEHzCCBBsCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUaHAlcX1wjr6JIO5vbXqKwpsS
# KzigggI4MIICNDCCAaGgAwIBAgIQWsnStFUuSIVNR8uhNSlE6TAJBgUrDgMCHQUA
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
# CisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMCMGCSqGSIb3DQEJBDEWBBSDCoSm
# 6vt0NHlPPlG4hKcFsPXdUTANBgkqhkiG9w0BAQEFAASBgEBKg914S7xltapONrWD
# 2xJVRBucXcDWRa6DaqIJt7gr1SuyKd7NwyxNdJLo9yiwxIOPowcTwel/4OdWZ/XI
# TCK21Ntr//BLEznatWKFo5uxzbPhI4nz0rRfrFgjfSiUHdTvysRCrsU6Op8umLKc
# /xFJcklOyRiYM5AaSbrdCt7t
# SIG # End signature block
