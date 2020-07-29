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
    Connect-EXO2 - Establish PSS to https://ps.outlook.com/powershell/
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
    [CmdletBinding()]
    [Alias('cxo2')]
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
            $CommandPrefix = 'xo' ; # 4:31 PM 7/29/2020 MS has RESERVED use of the 'exo' prefix [facepalm]
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
            if( get-command Get-xoAcceptedDomain) {
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
        
            # Connect-ExchangeOnline -Prefix 'xo' -Credential $credo365torsid -showBanner $false 
            # *** LEFT OFF HERE 5:01 PM 7/29/2020 *** not sure if it supports allowclobber, if it's actually wrapping pssession, it sure as shit does!
            $pltcXO = @{
                AllowClobber        = $true ;
                DisableNameChecking = $true ;
                Prefix              = $CommandPrefix ;
                ErrorAction         = 'Stop' ;
            } ;

            if ($MFA) {
                

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
                } catch {
                    Write-Warning -Message "Tried but failed to import the EXO PS module.`n`nError message:" ;
                    throw $_ ;
                } ;
                <# 8:45 AM 7/29/2020 VEN tenant now throwing error:
                    WARNING: Tried but failed to import the EXO PS module.
                    Error message:
                    Import-PSSession : Data returned by the remote Get-FormatData command is not in the expected format.
                    At C:\Program Files\WindowsPowerShell\Modules\verb-exo\1.0.14\verb-EXO.psm1:370 char:52
                    + ...   $Global:EOLModule = Import-Module (Import-PSSession @pltPSS) -Globa ...
                    +                                          ~~~~~~~~~~~~~~~~~~~~~~~~
                        + CategoryInfo          : InvalidResult: (:) [Import-PSSession], ArgumentException
                        + FullyQualifiedErrorId : ErrorMalformedDataFromRemoteCommand,Microsoft.PowerShell.Commands.ImportPSSessionCommand
                #>
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
}

#*------^ Connect-EXO2.ps1 ^------