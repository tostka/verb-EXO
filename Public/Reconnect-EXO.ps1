# Reconnect-EXO.ps1

#*------v Reconnect-EXO.ps1 v------
Function Reconnect-EXO {
   <#
    .SYNOPSIS
    Reconnect-EXO - Reestablish connection to Exchange Online (via EXO V2/V3 graph-api module)
    .NOTES
    Version     : 1.0.0
    Author      : Todd Kadrie
    Website     :	http://www.toddomation.com
    Twitter     :	@tostka / http://twitter.com/tostka
    CreatedDate : 2020-07-29
    FileName    :
    License     :
    Copyright   :
    Github      : https://github.com/tostka/verb-exo
    Tags        : Powershell,ExchangeOnline,Exchange,RemotePowershell,Connection,MFA
    REVISIONS   :
    * 3:11 PM 7/15/2024 needed to change CHKPREREQ to check for presence of prop, not that it had a value (which fails as $false); hadn't cleared $MetaProps = ...,'DOESNTEXIST' ; confirmed cxo working non-based
    *1:44 PM 7/9/2024 passes hybrid xo/s&c, with variant prefixes (other than hard-req that prefix cc indicates an s&c conn); 
         un-remmed cc-specific $UserRole default (steers into SID);  sub'd in silent for w-v
    * 4:12 PM 7/8/2024 passes dbg xo;  spliced over updates from cxo using test-exoconnectionTDO+resolve-AppIDToCBAFriendlyName(), tore out most old PROCESS testing & all END code; spliced over constants block from cxo as well
    * 9:55 AM 6/21/2024 add: prereq checks, and $isBased support, to devert into most basic Get-ConnectionInformation , Connect-ExchangeOnline fall back support
    * 5:18 PM 4/18/2024 spliced together hybrid of latest built and recent revs; undebugged;  been working a variant missing the 4/19/23-2/26/24 revs!
    * 2:51 PM 2/26/2024 add | sort version | select -last 1  on gmos, LF installed 3.4.0 parallel to 3.1.0 and broke auth: caused mult versions to come back and conflict with the assignement of [version] type (would require [version[]] to accom both, and then you get to code everything for mult handling)
    * 12:51 PM 5/30/2023 Updates to support either -Credential, or -UserRole + -TenOrg, to support fully portable downstream credentials: 
        - Add -UserRole & explicit -TenOrg params
        - Drive TenOrg defaulted $global:o365_TenOrgDefault, or on $env:userdomain
        - use the combo thru get-TenantCredential(), then set result to $Credential
        - if using Credential, the above are backed out via get-TenantTag() on the $credential 
        - CBA identifiers are resolve always via $uRoleReturn = resolve-UserNameToUserRole -Credential $Credential ;
    * 2:02 PM 5/25/2023 updated CBH example to reflect use of $pltRXOC std conn splat
    * 3:24 PM 5/23/2023 disabled msal.ps load ; fixed typo (-eq vs =, dumping $false into pipe)     
    * 4:24 PM 5/22/2023 add missing pswlt cmd for winrm chkline
    * 10:15 AM 5/19/2023 defer to resolve-UserNameToUserRole -Credential $Credential; assign certtag from output
    # 3:40 PM 5/10/2023 ported in block of CBA-handling code at 387
    # 1:36 PM 5/2/2023 port over cxo2 update: pltCXO2 -> $pltCXO; connect-EXO2 -> connect-EXO; Disconnect-EXO2 -> Disconnect-EXO
    # 3:18 PM 4/19/2023 under EOM310: replc $xmod.version refs with $EOMMv...
    * 11:20 AM 4/18/2023 step debugs ;  consolidate reconnect-exo2 into reconnect-exo (alias reconnect-exo2 & rxo2)
    * 2:02 PM 4/17/2023 rev: $MinNoWinRMVersion from 2.0.6 => 3.0.0.
    # * 11:02 AM 4/4/2023 reduced the ipmo and vers chk block, removed the lengthy gmo -list; and any autoinstall. Assume EOM is installed, & break if it's not
    * 3:14 pm 3/29/2023: REN'D $modname => $EOMModName
    * 11:01 AM 12/21/2022 moved $pltCXO2 def out to always occur (was only when !$bExistingEXOGood )
    * 3:59 PM 8/2/2022 got through dbugging EOM v205 SID interactive pass, working ; added -MinimumVersion & - MinNoWinRMVersion ; fully works from mybox w v206p6, cEOM connection, with functional prefix.
    * 3:30 PM 7/25/2022 tests against CBA & SID interactive creds on EOM v205, need to debug now against EOM v206p6, to accomodate PSSession-less connect & test code.
    * 3:54 PM 4/1/2022 add missing $silent param (had support, but no param)
    * 11:04 AM 3/30/2022 recode for ExchangeOnlineManagement v2.0.5, fundemental revisions, since prior v1.0.1
    # below here is orig rxo revs
    * 9:03 AM 12/14/2021 cleaned comments
    * 1:17 PM 8/17/2021 added -silent param
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
    Reconnect-EXO - Reestablish connection to Exchange Online (via EXO V2/V3 graph-api module)
    .PARAMETER  Prefix
    [verb]-PREFIX[command] PREFIX string for clearly marking cmdlets sourced in this connection [-Prefix tag]

    .PARAMETER Credential
    Credential to use for this connection [-credential [credential obj variable]
     .PARAMETER MinimumVersion
    MinimumVersion required for ExchangeOnlineManagement module (defaults to '2.0.5')[-MinimumVersion '2.0.6']
    .PARAMETER MinNoWinRMVersion
    MinimumVersion required for Non-WinRM connections (of ExchangeOnlineManagement module (defaults to '3.0.0')[-MinimumVersion '2.0.6']
    .PARAMETER ProxyEnabled
    Use Proxy-Aware SessionOption settings [-ProxyEnabled]
    .PARAMETER UserRole
    Credential Optional User Role spec for credential discovery (wo -Credential)(SID|CSID|UID|B2BI|CSVC|ESVC|LSVC|ESvcCBA|CSvcCBA|SIDCBA)[-UserRole @('SIDCBA','SID','CSVC')]
    .PARAMETER TenOrg
    Optional Tenant Tag (wo -Credential)[-TenOrg 'XYZ']
    .PARAMETER showDebug
    Debugging Flag [-showDebug]
    .PARAMETER silent
    Switch to specify suppression of all but warn/error echos.
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    None. Returns no objects or output.
    .EXAMPLE
    PS>  Reconnect-EXO;
    Reconnect EXO connection self-locating creds
    .EXAMPLE
    PS>  Reconnect-EXO -credential $cred ;
    Reconnect EXO connection with explicit [pscredential] object credential specified
    .EXAMPLE
    reconnect-exo -UserRole SIDCBA -TenOrg ABC -verbose  ; 
    Demo use of UserRole (specifying a CBA variant), AND TenOrg spec, to connect (autoresolves against preconfigured credentials in profile)
    .EXAMPLE
    PS> $pltRXOC = [ordered]@{
    PS>     Credential = $Credential ;
    PS>     verbose = $($VerbosePreference -eq "Continue")  ;
    PS>     Silent = $silent ; 
    PS> } ;
    PS> if ($script:useEXOv2 -OR $useEXOv2) { reconnect-eXO2 @pltRXOC }
    PS> else { reconnect-EXO @pltRXOC } ;    
    Splatted example leveraging prefab $pltRXOC splat, derived from local variables & $VerbosePreference value.
    .EXAMPLE
    PS>  $batchsize = 10 ;
    PS>  $RecordCount=$mr.count #this is the array of whatever you are processing ;
    PS>  $b=0 #this is the initialization of a variable used in the do until loop below ;
    PS>  $mrs = @() ;
    PS>  do {
    PS>      Write-Progress -Activity "Getting move request statistics for all $wave move requests." -Status "Processing Records $b through $($b+$batchsize) of $RecordCount." -PercentComplete ($b/$RecordCount*100) ;
    PS>      $mrs += $mr | Select-Object -skip $b -first $batchsize | foreach-object {Reconnect-EXO; $_ | Get-OLMoveRequestStatistics} ;
    PS>      $b=$b+$batchsize ;
    PS>      } ;
    PS>  until ($b -gt $RecordCount) ;
    Demo use of a Reconnect-EXO call mid-looping process to ensure connection stays available for duration of long-running process.    
    .LINK
    Github      : https://github.com/tostka/verb-exo
    #>
    [CmdletBinding()]
    [Alias('rxo','reconnect-exo2','rxo2')]
    PARAM(
        [Parameter(HelpMessage = "[verb]-PREFIX[command] PREFIX string for clearly marking cmdlets sourced in this connection [-Prefix tag]")]
            [string]$Prefix = 'xo',
        [Parameter(HelpMessage = "MinimumVersion required for ExchangeOnlineManagement module (defaults to '2.0.5')[-MinimumVersion '2.0.6']")]
            [version] $MinimumVersion = '2.0.5',
        [Parameter(HelpMessage = "MinimumVersion required for Non-WinRM connections (of ExchangeOnlineManagement module (defaults to '3.0.0')[-MinimumVersion '2.0.6']")]
            [version] $MinNoWinRMVersion = '3.0.0',
        [Parameter(HelpMessage="Use Proxy-Aware SessionOption settings [-ProxyEnabled]")]
            [boolean]$ProxyEnabled = $False,
        [Parameter(HelpMessage="Credential to use for this connection [-credential [credential obj variable]")]
            [System.Management.Automation.PSCredential]$Credential,
            # = $global:credo365TORSID, # defer to TenOrg & UserRole resolution
        [Parameter(Mandatory = $false, HelpMessage = "Credential User Role spec (SID|CSID|UID|B2BI|CSVC|ESVC|LSVC|ESvcCBA|CSvcCBA|SIDCBA)[-UserRole @('SIDCBA','SID','CSVC')]")]
            # sourced from get-admincred():#182: $targetRoles = 'SID', 'CSID', 'ESVC','CSVC','UID','ESvcCBA','CSvcCBA','SIDCBA' ; 
            #[ValidateSet("SID","CSID","UID","B2BI","CSVC","ESVC","LSVC","ESvcCBA","CSvcCBA","SIDCBA")]
            # pulling the pattern from global vari w friendly err
            [ValidateScript({
                if(-not $rgxPermittedUserRoles){$rgxPermittedUserRoles = '(SID|CSID|UID|B2BI|CSVC|ESVC|LSVC|ESvcCBA|CSvcCBA|SIDCBA)'} ;
                if(-not ($_ -match $rgxPermittedUserRoles)){throw "'$($_)' doesn't match `$rgxPermittedUserRoles:`n$($rgxPermittedUserRoles.tostring())" ; } ; 
                return $true ; 
            })]
            [string[]]$UserRole = @('SIDCBA','SID','CSVC'),
        [Parameter(Mandatory=$FALSE,HelpMessage="TenantTag value, indicating Tenants to connect to[-TenOrg 'TOL']")]
            [ValidateNotNullOrEmpty()]
            #[ValidatePattern("^\w{3}$")]
            [string]$TenOrg = $global:o365_TenOrgDefault,
        [Parameter(HelpMessage="Debugging Flag [-showDebug]")]
            [switch] $showDebug,
        [Parameter(HelpMessage="Silent output (suppress status echos)[-silent]")]
            [switch]$silent
    ) ;
    BEGIN{
        $verbose = ($VerbosePreference -eq "Continue") ; 
        
        #region CHKPREREQ ; #*------v CHKPREREQ v------
        # critical dependancy Meta variables
        $MetaNames = 'TOR','CMW','TOL' # ,'NOSUCH' ;
        # critical dependancy Meta variable properties
        $MetaProps = 'legacyDomain','o365_TenantDomain' #,'DOESNTEXIST' ;
        $isBased = $true ; $gvMiss = @() ; $ppMiss = @() ;
        foreach($met in $metanames){
            write-verbose "chk:`$$($met)Meta" ;
            if(-not (gv -name "$($met)Meta" -ea 0)){
                $isBased = $false; $gvMiss += "$($met)Meta" ;
            } ;
            foreach($mp in $MetaProps){
                write-verbose "chk:`$$($met)Meta.$($mp)" ;
                #if(-not (gv -name "$($met)Meta" -ea 0).value[$mp]){ # testing has a value, not is present as a spec!
                if(-not (gv -name "$($met)Meta" -ea 0).value.keys -contains $mp){
                    $isBased = $false; $ppMiss += "$($met)Meta.$($mp)" ;
                } ;
            } ;
        } ;
        if($gvmiss){write-warning "Missing Dependant Meta variables:`n$(($gvMiss |%{"`$$($_)" }) -join ',')" } ;
        if($ppMiss){write-warning "Missing Dependant Meta vari properties:`n$(($ppMiss |%{"`$$($_)" }) -join ',')" } ;
        if(-not $isBased){ write-warning  "missing critical dependancy profile config!" } ;
        #endregion CHKPREREQ ; #*------^ END CHKPREREQ ^------


        if($gvmiss){write-warning "Missing Dependant Meta variables:`n$(($gvMiss |%{"`$$($_)" }) -join ',')" } ; 
        if($ppMiss){write-warning "Missing Dependant Meta vari properties:`n$(($ppMiss |%{"`$$($_)" }) -join ',')" } ; 
        if(-not $isBased){ throw "missing critical dependancy profile config!" } ; 
        #endregion CHKPREREQ ; #*------^ END CHKPREREQ ^------
        if(!$rgxExoPsHostName){$rgxExoPsHostName="^(ps\.outlook\.com|outlook\.office365\.com)$" } ;
        if(-not (get-variable rgxExoPsHostName -ea 0)){$rgxExoPsHostName="^(ps\.outlook\.com|outlook\.office365\.com)$" } ;
        if(-not (get-variable rgxCertThumbprint -ea 0)){$rgxCertThumbprint = '[0-9a-fA-F]{40}' ; } ;
        if(-not (get-variable rgxCertFNameSuffix -ea 0)){$rgxCertFNameSuffix = '-([A-Z]{3})$' ; } ; 

        <# 4:45 PM 7/7/2022 workaround msal.ps bug: always ipmo it FIRST: "Get-msaltoken : The property 'Authority' cannot be found on this object. Verify that the property exists."
        # admin/SID module auto-install code (myBoxes UID split-perm CU, all else t AllUsers)
        $modname = 'MSAL.PS' ;
        $error.clear() ;
        $pltIMod = @{Name = $modname ; ErrorAction = 'Stop' ; verbose=$false} ;
        # this forces a specific rev into the ipmo!
        if ($MinimumVersion) { $pltIMod.add('MinimumVersion', $MinimumVersion.tostring()) } ;
        $error.clear() ;
        Try { Get-Module $modname -ErrorAction Stop | out-null } Catch {
            $smsg = "Import-Module w`n$(($pltIMod|out-string).trim())" ;
            if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            Import-Module @pltIMod ;
        } ; # IsImported
        #>
        
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
        #if(-not (gv rgxEXoPrunspaceConnectionInfoAppName -ea 0 )){$rgxEXoPrunspaceConnectionInfoAppName = '^/(exadmin|powershell)$'}; 
        #if(-not (gv EXoPrunspaceConnectionInfoPort -ea 0 )){$EXoPrunspaceConnectionInfoPort = '80' } ; 
        # gmo varis
        #if(-not (gv rgxExoPsHostName -ea 0 )){ $rgxExoPsHostName = "^(ps\.outlook\.com|outlook\.office365\.com)$" } ;
        #if(-not (gv rgxEXOv1gmoDescription -ea 0 )){$rgxEXOv1gmoDescription = "^Implicit\sremoting\sfor\shttps://ps\.outlook\.com/PowerShell" }; 
        #if(-not (gv EXOv1gmoprivatedataImplicitRemoting -ea 0 )){$EXOv1gmoprivatedataImplicitRemoting = $true };
        #if(-not (gv rgxEXOv2gmoDescription -ea 0 )){$rgxEXOv2gmoDescription = "^Implicit\sremoting\sfor\shttps://outlook\.office365\.com/PowerShell" }; 
        #if(-not (gv EXOv2gmoprivatedataImplicitRemoting -ea 0 )){$EXOv2gmoprivatedataImplicitRemoting = $true } ;
        #if(-not (gv rgxExoPsessionstatemoduleDescription -ea 0 )){$rgxExoPsessionstatemoduleDescription = '/(exadmin|powershell)$' };
        #if(-not (gv EXOv2StateOK -ea 0 )){$EXOv2StateOK = 'Opened'} ; 
        #if(-not (gv EXOv2AvailabilityOK -ea 0 )){$EXOv2AvailabilityOK = 'Available'} ; 
        #if(-not (gv EXOv2RunStateBad -ea 0 )){ $EXOv2RunStateBad = 'Broken'} ;
        #if(-not (gv EXOv1GmoFilter -ea 0 )){$EXOv1GmoFilter = 'tmp_*' } ; 
        if(-not (gv EXOv2GmoNoWinRMFilter -ea 0 )){$EXOv2GmoNoWinRMFilter = 'tmpEXO_*' };
        # add get-connectioninformation.ConnectionURI targeting rgxs for CCMS vs EXO
        if(-not $rgxConnectionUriEXO){$rgxConnectionUriEXO = 'https://outlook\.office365\.com'} ; 
        if(-not $rgxConnectionUriEXO){$rgxConnectionUriCCMS = 'https://ps\.compliance\.protection\.outlook\.com'} ; 
        $sTitleBarTag = @("EXO2") ;
        #*------^ END PSS & GMO VARIS ^------

        # disable prefix spec, unless actually blanked (e.g. centrally spec'd in profile).
        if (-not $Prefix) {
            $Prefix = 'xo' ; # 4:31 PM 7/29/2020 MS has RESERVED use of the 'exo' prefix [facepalm]
            #$Prefix = 'cc' ; # ccms variant
            $smsg = "(asserting Prefix:$($Prefix)" ;
            if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        } ;
        if (($Prefix) -and ($Prefix -eq 'EXO')) {
            throw "Prefix 'EXO' is a reserved Prefix, please use a different prefix."
        }
        if($Prefix -eq 'cc'){
            # build in hybrid xo & ccms support, switch on the prefix spec
            $useCCMSConn = $true ; 
        }; 
        if($useCCMSConn){
            # respec userrole
            $UserRole = @('SID') ; 
            $sTitleBarTag = @("CCMS") ;
        } ; 

         # * 11:02 AM 4/4/2023 reduced the ipmo and vers chk block, removed the lengthy gmo -list; and any autoinstall. Assume EOM is installed, & break if it's not
        #region EOMREV ; #*------v EOMREV Check v------
        # reduced the ipmo and vers chk block, removed the lengthy gmo -list; and any autoinstall. Assume EOM is installed, & break if it's not
        $pltIMod = @{Name = $EOMmodname ; ErrorAction = 'Stop' ; verbose=$false} ;
        if($xmod = Get-Module $EOMmodname -ErrorAction Stop| sort version | select -last 1 ){ } else {
            $smsg = "Import-Module w`n$(($pltIMod|out-string).trim())" ;
            if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
            Try {
                Import-Module @pltIMod | out-null ;
                $xmod = Get-Module $EOMmodname -ErrorAction Stop | sort version | select -last 1 ;
            } Catch {
                $ErrTrapd=$Error[0] ;
                $smsg = "$('*'*5)`nFailed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: `n$(($ErrTrapd|out-string).trim())`n$('-'*5)" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                $smsg = $ErrTrapd.Exception.Message ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                Break ;
            } ;
        } ; # IsImported
        if([version]$xmod.version -ge $MinNoWinRMVersion){
            $MinNoWinRMVersion = $xmod.version.tostring() ;
            $IsNoWinRM = $true ; 
        }
        [boolean]$UseConnEXO = [boolean]([version]$xmod.version -ge $MinNoWinRMVersion) ; 
        #endregion EOMREV ; #*------^ END EOMREV Check  ^------

        if(-not $isBased){
            # default to most basic rudimentary connection
            $Status = Get-ConnectionInformation -ErrorAction SilentlyContinue
            If (-not ($Status)) {
              #Connect-ExchangeOnline -SkipLoadingCmdletHelp
              Connect-ExchangeOnline -SkipLoadingCmdletHelp -ShowBanner:$false ; 
            }
        }else {

            if(-not $Credential){
                if($UserRole){
                    $smsg = "Using specified -UserRole:$( $UserRole -join ',' )" ;
                    if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                } else { $UserRole = @('SID','CSVC') } ;
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
            if($Prefix){
                if($useCCMSConn){
                    $smsg = "(adding specified  Connect-IPPSSession -Prefix:$($Prefix))" ; 
                } else { 
                    $smsg = "(adding specified Connect-ExchangeOnline -Prefix:$($Prefix))" ; 
                } ; 
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                $pltCXO.add('Prefix',$Prefix) ; 
            } ; 
            if((gcm connect-EXO).Parameters.keys -contains 'silent'){
                $pltCXO.add('Silent',$false) ;
            } ;

            # defer to resolve-UserNameToUserRole -Credential $Credential
            
            $uRoleReturn = resolve-UserNameToUserRole -Credential $Credential ; 
            if($credential.username -match $rgxCertThumbprint){
                $certTag = $uRoleReturn.TenOrg ; 
            } ; 
        }  # if-E $isBased 
    } ;  # BEG-E
    PROCESS{
        if($isBased){

            # normal automation
            $bExistingEXOGood = $false ; 
            $exov2Good = $exov3Good = $null ; 
            if( $legXPSSession = Get-PSSession | where-object {$_.ConfigurationName -like "Microsoft.Exchange" -AND $_.Name -match "^(Session|WinRM)\d*" } ){
                # ignore state & Avail, close the conflicting legacy conn's
                $smsg = "(existing legacy-EXO or Broken connections found, closing)" ; 
                if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 

                Disconnect-Exo ; Disconnect-PssBroken ;Start-Sleep -Seconds 3;
                $bExistingEXOGood = $false ; 
            } ; 
            <# 3:50 PM 4/7/2022 ExchangeOnlinemanagement has a *bug*
            [Issue using ExchangeOnlineManagement v2.0.4 module to connect to Exchange Online remote powershell (EXO) and Exchange On-Prem remote powershell (EXOP) in same powershell window - Microsoft Q&A - docs.microsoft.com/](https://docs.microsoft.com/en-us/answers/questions/451786/issue-using-exchangeonlinemanagement-v204-module-t.html?childToView=804782#answer-804782)
            It *can't* cleanly reconnect the legacy EXO cmdlet dependant implicit-remote sessions, when they time out
            if there's an *existing* Exchange Onprem implicit remote session. 

            No fix, lame workaround: close all implicit remote sessions, before reopening EMO *first*, and then reconnecting EXOnPrem

            #>
            # check for existing implicit remote EXOnPrem sessions we have to kill first (and then post-reconnect)


            #clear invalid existing EXOv2 PSS's
           # fix typo from Name -eq to -like 
           # new token prop expir support:
           #((Get-PSSession | where Name -like ExchangeOnlineInternalSession*).TokenExpiryTime -lt (get-date ))
            #$exov2Broken = Get-PSSession | where-object {$_.ConfigurationName -like "Microsoft.Exchange" -AND (
            #    $_.Name -like "ExchangeOnlineInternalSession*") -AND $_.State -like "*Broken*"}
            # add token support clause: -AND ($_.TokenExpiryTime -lt (get-date )), leaving it out, may be a usecase with Broken but not post expiration
            # just fix the non-wildcard -like's w proper -eq's 
            $exov2Broken = Get-PSSession | where-object {$_.ConfigurationName -eq "Microsoft.Exchange" -AND (
                $_.Name -like "ExchangeOnlineInternalSession*") -AND $_.State -eq "Broken" }
            $exov2Closed = Get-PSSession | where-object {$_.ConfigurationName -like "Microsoft.Exchange" -AND (
                $_.Name -like "ExchangeOnlineInternalSession*") -AND $_.State -like "*Closed*"}
        
            #if ($exov2Broken.count -gt 0){for ($index = 0 ;$index -lt $exov2Broken.count ;$index++){Remove-PSSession -session $exov2Broken[$index]} };
            #if ($exov2Closed.count -gt 0){for ($index = 0 ;$index -lt $exov2Closed.count ; $index++){Remove-PSSession -session $exov2Closed[$index] } } ; 
            # sub full Disconnect-EXO for Remove-PSSession: dxo2 includes 'Clear-ActiveToken -TokenProvider $session.TokenProvider' in addition to remove-pssession

            $pltDXO2=[ordered]@{
                verbose = $($VerbosePreference -eq "Continue") ;        
                silent = $silent ; 
            } ;
            if($Prefix){
                $pltDXO2.add('Prefix',$Prefix) ; 
            } ; 

            if ( ($exov2Broken.count -gt 0) -OR ($exov2Closed.count -gt 0)){
                $smsg = "Disconnect-EXO w`n$(($pltDXO2|out-string).trim())" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                $tIn = '4/18/2024;D:\scripts\DisConnect-EXO_func.ps1' ;
                $tdt,$tsrc =  $tIn.split(';') ;
                $tdt=[datetime]$tdt ;
                if($psise -and (get-date ).date -eq $tdt){
                    $gcm = gcm (split-path $tsrc -leaf).replace('_func.ps1','') ;
                    if( $gcm -AND $gcm.source -ne ''){
                        gci function:$((split-path $tsrc -leaf).replace('_func.ps1','')) -ea 0| remove-item -force ;    
                        ipmo -fo -verb $tsrc;
                    } else {write-host "(non-Mod vers loaded)"} ;
                } ; 

                 Disconnect-EXO @pltDXO2 ;
            };
        
            # -------------
                
            #$oRet = test-EXOv2Connection -Credential $credential -CertTag $certtag -Prefix $Prefix -verbose:$($verbose) ; 
                
            #$oRet = test-EXOv2Connection -Credential $credential -Prefix $Prefix -verbose:$($verbose) ; 
            
            $pltTXO=[ordered]@{
                    erroraction = 'STOP' ;
            } ;
            if($Prefix){
                $pltTXO.add('Prefix',$Prefix) ; 
            } ; 
            $smsg = "test-EXOConnectionTDO w`n$(($pltTXO|out-string).trim())" ; 
            if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 

            $bExistingEXOGood = $bExistingCCMSGood = $false ;
            if($oRet = test-EXOConnectionTDO @pltTXO ){
                foreach($xSess in $oRet){
                    if($null -eq $xSess.Organization -AND $xSess.TenantID){
                        $Tenantdomain = convert-TenantIdToDomainName -TenantId $xSess.TenantID ;
                        $smsg = "(coercing blank Session Org, to resolved TenantID equivelent TenantDomain)" ; 
                        if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                        $xSess.Organization = $Tenantdomain ; 
                    } ; 
                    if($xSess.isCBA){
                        $uRoleReturn = resolve-AppIDToCBAFriendlyName -AppID $xSess.AppId -verbose:$($VerbosePreference -eq "Continue")  ;
                        $certUname = $uRoleReturn.FriendlyName ;
                        $certTag = $uRoleReturn.TenOrg ;
                    } ;
                    if($xSess.isValid){
                        $smsg = "Connected to " ;
                        if($xSess.isXO){
                            $smsg += "XO EOM PS "

                            $bExistingEXOGood = $true ; 
                        }
                        elseif($xSess.isSC){
                            $smsg += "Sec & Compl PS " 
                            $bExistingCCMSGood = $true ;
                        }else{
                            $smsg += "DISCONNECTED!" ; 
                        } ; 
                        if($xSess.isCBA){
                            $smsg += "using CBA:" ;
                            $smsg += " $($certUname)" ;
                        } else{
                            $smsg += "using Account:" ;
                            $smsg += " $($xsess.UserPrincipalName)" ;
                            if($null -eq $xSess.Organization -AND $Tenantdomain){
                                $smsg += " ($($Tenantdomain.split('.')[0]))" ;
                            }elseif($xSess.Organization){
                                $smsg += " ($($xSess.Organization.split('.')[0]))" ;
                            } ; 
                        } ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success

                        
                    } else {
                        $smsg = "Not currently connected (TokenStatus:$($xSess.connection.TokenStatus))" ;
                        $smsg += "`nPreviously: "
                        if($xSess.isXO){
                            $smsg += "XO EOM PS " ; 
                            $bExistingEXOGood = $false ;
                        }
                        elseif($xSess.isSC){
                            $smsg += "Sec & Compl PS " 
                            $bExistingCCMSGood = $false ;
                        }else{
                            $smsg += "DISCONNECTED!" ;
                            $bExistingEXOGood = $bExistingCCMSGood = $false ;
                        } ;
                        if($xSess.isCBA){
                            $smsg += " using CBA:" ;
                            $smsg += " $($certUname)" ;
                        } else{
                            $smsg += "using Account:" ;
                            $smsg += " $($xsess.UserPrincipalName)" ;
                        } ;
                        if($null -eq $xSess.Organization -AND $Tenantdomain){
                            $smsg += " ($($Tenantdomain.split('.')[0]))" ;
                        }elseif($xSess.Organization){
                            $smsg += " ($($xSess.Organization.split('.')[0]))" ;
                        } else {
                            $smsg += " (neither Organization nor TenantID is populated)" ;
                        } ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                    } ;
                } ;  # loop-E
            } ; 

            if($Prefix -ne 'cc' -AND $bExistingEXOGood){
                $smsg = "existing EXO session" ; 
                if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            }elseif($Prefix -eq 'cc' -AND $bExistingCCMSGood){
                $smsg = "existing S&C session" ; 
                if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            } else { 
                $smsg = "NO VALID EXO3 "
                if($Prefix -ne 'cc'){
                    $smsg += "EXO -prefix $($prefix) "
                } else {
                    $smsg += "S&C -prefix $($prefix) "
                } ; 
                $smsg += "SESSION FOUND! (DISCONNECTING...)"

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
                #$bExistingEXOGood = $bExistingCCMSGood = $false ;

                #if($bExistingEXOGood -eq $false){
                if( ($Prefix -ne 'cc' -AND -not $bExistingEXOGood) -OR ($Prefix -eq 'cc' -AND -not $bExistingCCMSGood) ){
                
                    $smsg = "connect-EXO w $($credential.username):`n$(($pltCXO|out-string).trim())" ; 
                    if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                    #connect-EXO -Credential $Credential -verbose:$($verbose) ; 
                    connect-EXO @pltCXO ;               
                } ; 

            } ;  # if($Prefix -ne 'cc' -AND $bExistingEXOGood){ }elseif($Prefix -eq 'cc' -AND $bExistingCCMSGood){
        
        } else { 
            $smsg = "(-not:`$isBased: running most basic Get-ConnectionInformation , Connect-ExchangeOnline connectivity)" ; 
            if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            # default to most basic rudimentary connection
            $Status = Get-ConnectionInformation -ErrorAction SilentlyContinue
            If (-not ($Status)) {
              #Connect-ExchangeOnline -SkipLoadingCmdletHelp
              Connect-ExchangeOnline -SkipLoadingCmdletHelp -ShowBanner:$false ; 
            }
        } ; 
    } ;  # PROC-E
   END {
        
        if($isBased){

            $pltTXO=[ordered]@{erroraction = 'STOP' } ;
            if($Prefix){$pltTXO.add('Prefix',$Prefix) } ; 
            $smsg = "test-EXOConnectionTDO w`n$(($pltTXO|out-string).trim())" ; 
            if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 

            $bExistingEXOGood = $bExistingCCMSGood = $false ;
            if($oRet = test-EXOConnectionTDO @pltTXO ){
                foreach($xSess in $oRet){
                    if($null -eq $xSess.Organization -AND $xSess.TenantID){
                        $Tenantdomain = convert-TenantIdToDomainName -TenantId $xSess.TenantID ;
                        $smsg = "(coercing blank Session Org, to resolved TenantID equivelent TenantDomain)" ; 
                        if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                        $xSess.Organization = $Tenantdomain ; 
                    } ;
                    if($xSess.isCBA){
                        $uRoleReturn = resolve-AppIDToCBAFriendlyName -AppID $xSess.AppId -verbose:$($VerbosePreference -eq "Continue")  ;
                        $certUname = $uRoleReturn.FriendlyName ;
                        $certTag = $uRoleReturn.TenOrg ;
                    } ;
                    if($xSess.isValid){
                        $smsg = "Connected to " ;
                        if($xSess.isXO){
                            $smsg += "XO EOM PS "
                            $bExistingEXOGood = $true ;
                        }
                        elseif($xSess.isSC){
                            $smsg += "Sec & Compl PS "
                            $bExistingCCMSGood = $true ;
                        }else{
                            $smsg += "DISCONNECTED!" ;
                            $bExistingEXOGood = $bExistingCCMSGood = $false ;
                        } ;
                        if($xSess.isCBA){
                            $smsg += "using CBA:" ;
                            $smsg += " $($certUname)" ;
                        } else{
                            $smsg += "using Account:" ;
                            $smsg += " $($xsess.UserPrincipalName)" ;
                            if($null -eq $xSess.Organization -AND $Tenantdomain){
                                $smsg += " ($($Tenantdomain.split('.')[0]))" ;
                            }elseif($xSess.Organization){
                                $smsg += " ($($xSess.Organization.split('.')[0]))" ;
                            } ;
                        } ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                    } else {
                        $smsg = "Not currently connected (TokenStatus:$($xSess.connection.TokenStatus))" ;
                        $smsg += "`nPreviously: "
                        if($xSess.isXO){
                            $smsg += "XO EOM PS " ;
                            $bExistingEXOGood = $false ;
                        }
                        elseif($xSess.isSC){
                            $smsg += "Sec & Compl PS "
                            $bExistingCCMSGood = $false ;
                        }else{
                            $smsg += "DISCONNECTED!" ;
                            $bExistingEXOGood = $bExistingCCMSGood = $false ;
                        } ;
                        if($xSess.isCBA){
                            $smsg += " using CBA:" ;
                            $smsg += " $($certUname)" ;
                        } else{
                            $smsg += "using Account:" ;
                            $smsg += " $($xsess.UserPrincipalName)" ;
                        } ;
                        if($null -eq $xSess.Organization -AND $Tenantdomain){
                            $smsg += " ($($Tenantdomain.split('.')[0]))" ;
                        }elseif($xSess.Organization){
                            $smsg += " ($($xSess.Organization.split('.')[0]))" ;
                        } else {
                            $smsg += " (neither Organization nor TenantID is populated)" ;
                        } ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                    } ;

                } ;  # loop-E
            } else {
                $smsg = "No connection info returned" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                $bExistingEXOGood = $bExistingCCMSGood = $false ;
            } ; 

            if($Prefix -ne 'cc' -AND $bExistingEXOGood){
                $smsg = "existing EXO session" ;
                if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            }elseif($Prefix -eq 'cc' -AND $bExistingCCMSGood){
                $smsg = "existing S&C session" ;
                if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            } else {
                $smsg = "NO VALID EXO3 "
                if($Prefix -ne 'cc'){
                    $smsg += "EXO -prefix $($prefix) "
                } else {
                    $smsg += "S&C -prefix $($prefix) "
                } ;
                $smsg += "SESSION FOUND! (DISCONNECTING...)"
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
                #$bExistingEXOGood = $bExistingCCMSGood = $false ;
                #if($bExistingEXOGood -eq $false){
                if( ($Prefix -ne 'cc' -AND -not $bExistingEXOGood) -OR ($Prefix -eq 'cc' -AND -not $bExistingCCMSGood) ){
                    $smsg = "connect-EXO w $($credential.username):`n$(($pltCXO|out-string).trim())" ;
                    if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                    #connect-EXO -Credential $Credential -verbose:$($verbose) ;
                    connect-EXO @pltCXO ;
                } ;
            } ;  # if($Prefix -ne 'cc' -AND $bExistingEXOGood){ }elseif($Prefix -eq 'cc' -AND $bExistingCCMSGood){
            
        } else { 
            $smsg = "(-not:`$isBased: running most basic Get-ConnectionInformation , Connect-ExchangeOnline connectivity)" ; 
            #if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            #else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            
            $pltGCInfo=[ordered]@{
                Credential = $Credential ;
                verbose = $($verbose) ; 
                erroraction = 'STOP' ;
            } ;
            if($Prefix){
                $smsg = "(checking specified  -Prefix:$($Prefix))" ; 
                if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                $pltGCInfo.add('ModulePrefix',$Prefix) ; 
            } ; 
            
            
            $smsg = "get-ConnectionInformation w`n$(($pltGCInfo|out-string).trim())" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            if($cInfo = Get-ConnectionInformation @$pltGCInfo){
                $smsg = "get-ConnectionInformation w`n$(($cInfo | fl |out-string).trim())" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } else {
                $smsg = "Get-ConnectionInformation: NO CONNECTION INFORMATION RETURNED! " ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
            } ;  
        } ; 
    }  # END-E
}

#*------^ Reconnect-EXO.ps1 ^------