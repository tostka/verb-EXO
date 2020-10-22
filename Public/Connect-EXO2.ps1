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