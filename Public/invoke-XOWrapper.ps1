# invoke-XOWrapper_func.ps1
#*------v Function invoke-XOWrapper v------
function invoke-XOWrapper  {
    <#
    .SYNOPSIS
     invoke-XOWrapper.ps1 - (alias XoW) Wraps a given ExchangeOnlineManagment module (EOM) cmdlet, with a pre-authentication test and token refresh, to work around (Hybrid Exch bug: 'GetSteppablePipeline' in v2.0.5 of the module, when any basic-authenticated Exchange Onprem session is concurrently open in the session (EOM can't differentiate the EXO session from the ExOP session).
    .NOTES
    Version     : 0.0.1
    Author      : Todd Kadrie
    Website     : http://www.toddomation.com
    Twitter     : @tostka / http://twitter.com/tostka
    FileName    : invoke-XOWrapper_func.ps1
    CreatedDate : 2022-09-16
    License     : MIT License
    Copyright   : (c) 2022 Todd Kadrie
    Github      : https://github.com/tostka/verb-XXX
    Tags        : Powershell,ExchangeOnlineManagement,Bug,Workaround
    REVISIONS
    * 3:24 PM 11/28/2022 move to verb-EXO: ren xoW.ps1 -> invoke-XOWrapper with alias spec'd
    * 2:31 PM 11/23/2022 add: MinNoWinRMVersion param, used to test UseConnEXO, to avoid testing test-exoToken pres (not going to be there -gt v206); updated CBH
         Add: $Credential & -CredentialOP to permit calls from existing scripts with preferred creds
    * 4:33 PM 10/5/2022 CBH exmpl add
    * 10:12 AM 9/16/2022init
    .DESCRIPTION
    invoke-XOWrapper.ps1 - (alias XoW) Wraps a given ExchangeOnlineManagment module (EOM) cmdlet, with a pre-authentication test and token refresh, to work around (Hybrid Exch bug: 'GetSteppablePipeline' in v2.0.5 of the module, when any basic-authenticated Exchange Onprem session is concurrently open in the session (EOM can't differentiate the EXO session from the ExOP session).
    - For ExchangeOnlineManagement v205 or less: Leverages my verb-EXO module:test-EXOToken() wrap of the EOM v205 (and before) internal Test-ActiveToken() function (which has been torn out of EOM v206pre6 forward, as it doesn't use the basic auth that creates the long-unaddressed conflict).
    - ExchangeOnlineManagement v206+ it isn't necessary so it simply wraps & returns the underlying targeted cmdlet (to provide backward compatibility on either v205 or v206+).
    - Requires/relies on connection-maintenance functions (aliased below) from my verb-EXO and verb-AAD modules.
    - For use without those modules, replace the...
      dx10 ;  dxo ;  dxo2 ;  daad ;  rxo2 ;  rx10 ;  caad ;
    ... commands with...
      get-psssession | remove-pssession ; 
      Connect-ExchangeOnline -UserPrincipalName ADMINLOGON@DOMAIN ; 
      $ExOPServer="SERVEREXOP" ; $pltNPSS=@{ConnectionURI="http://$ExOPServer/powershell";ConfigurationName='Microsoft.Exchange' ; name='ExchangeOP'} ; 
      $ExPSS = New-PSSession @pltNPSS  ;
      $ExIPSS = Import-PSSession $ExPSS -allowclobber ; 
      Connect-AzureAD -Credential ADMINLOGON@DOMAIN ; 
    ... to use strictly generic EOM, AAD & PS cmdlets for the handling.
    
    - Note: Even with the above changes my verb-EXO module is still a *hard dependancy*, as the EOM:test-ActiveToken() that validates token status *isn't an exported public function* within EOM, 
    and loading it requires manually constructed pathing - .net vs .netcore variants - and an ipmo to get it into memory (hence my test-EXOToken wrapper). 
    .PARAMETER  Command
    ExchangeOnlineManagement module cmdlet to be wrapped and run. Make the target commands *a scriptblock* - Wrap the target in curly-braces - to get multiline defs in without the need for fancy nested quoting, or invoke-expression etc. (alias 'cmd')
    .PARAMETER  Credential
    Credential object
    .PARAMETER MinNoWinRMVersion
    MinimumVersion required for Non-WinRM connections (of ExchangeOnlineManagement module (defaults to '2.0.6')[-MinimumVersion '2.0.6']
    .EXAMPLE
    PS> write-verbose ipmo the invoke-XOWrapper module file into memory ; 
    PS> ipmo -fo -verb d:\scripts\invoke-XOWrapper_func.ps1 ; 
    PS> write-verbose run a curly-wrapped scriptblock through the invoke-XOWrapper() function. 
    PS> invoke-XOWrapper {get-xorecipient *somename* | select -expand prim*} ; 
    Load the function from file, then run simple wrap of get-recipient (uses positional parameter spec, to avoid need for -cmd).
    .EXAMPLE
    PS> invoke-XOWrapper -cmd {get-xorecipient *namestring* | ?{$_.PrimarySmtpAddress -like '*@domain.com'}  | select -expand prim*} -verbose ; 
        VERBOSE: (confirm EMO load)
        VERBOSE: (check EMO version)
        VERBOSE: (test for test-exoToken())
        Azure Active Directory - Disconnected
        10:15:57:Connecting to EXOv2:(somID@domain.com)
        True
        20220916 10:16:07:Adding EMS (connecting to serverExOP.sub.domain.com)...

        ComputerName                Availability  State ConfigurationName
        ------------                ------------  ----- -----------------
        serverExOP.sub.domain.com    Available Opened Microsoft.Exchange
        ...
        10:16:13:Authenticating to AAD:toro.com, w somID@domain.com...
        10:16:13:Connect-AzureAD w
        Name                           Value
        ----                           -----
        ErrorAction                    Stop
        TenantID                       549366xxxxba08b
        AccountId                      somID@domain.com
        10:16:14:
        Account                Environment TenantId                             TenantDomain           AccountType
        -------                ----------- --------                             ------------           -----------
        somID@domain.com       AzureCloud  549366xxxxba08b                      TENANT.onmicrosoft.com User
        10:16:14:Connected to Tenant:
        TenantId                             UserId                 LoginType
        --------                             ------                 ---------
        549366xxxxba08b somID@domain.com     LiveId
        10:16:14:(Authenticated to AAD:TOR as somID@domain.com    
        somerecipient@domain.com
    Includes typical reauth output when Token has expired, with verbose output
    .EXAMPLE
    PS>  $xrcp = invoke-XOWrapper {get-xorecipient somealias }
    Demo capturing return (which is dropped into pipeline within invoke-XOWrapper) into a variable ; 
    .EXAMPLE
    PS>  $xrcp = invoke-XOWrapper {get-xorecipient somealias } -credential $pltRXO.Credential -credentialOP $pltRX10.Credential ; 
    Demo passing in specified credentials ; 
    #>
    #Requires -Modules ExchangeOnlineManagement, verb-EXO, AzureAD, verb-AAD
    [CmdletBinding()] 
    [Alias('xoW')]
    PARAM(
        [Parameter(Position=0)][Alias('cmd')]
        $command,
        [Parameter(HelpMessage = "EXO Credential to use for this connection [-credential `$credo365]")]
        [System.Management.Automation.PSCredential]$Credential = $global:credo365TORSID,
        [Parameter(HelpMessage = 'ExOP Credential object (for staged hybrid XOP reconnects)[-credentialOP `$credOP]')]
        [System.Management.Automation.PSCredential]$CredentialOP = $credTORSID,
        [Parameter(HelpMessage = "MinimumVersion required for Non-WinRM connections (of ExchangeOnlineManagement module (defaults to '2.0.6')[-MinimumVersion '2.0.6']")]
        [version] $MinNoWinRMVersion = '2.0.6'
    ) ; 
    write-verbose "(confirm EMO load)" ; 
    $tMod = 'exchangeonlinemanagement' ; 
    if(-not (get-module $tMod)){ipmo -force $tMod} ; 
    $xMod = get-module $tMod ; 
    write-verbose "(check EMO version)" ; 
    #[boolean]$UseConnEXO = [boolean]([version](get-module $modname).version -ge $MinNoWinRMVersion) ; 
    [boolean]$UseConnEXO = [boolean]([version]$xMod.version -ge $MinNoWinRMVersion) ; 
    if($UseConnEXO){
        $smsg = "$($xMod.Name) v$($xMod.Version.ToString()) is GREATER than $($MinNoWinRMVersion):this function is not needed for natively Modern Auth EXO connectivit!" ; 
        $smsg += "`n(proxying command through...)" ; 
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        #Levels:Error|Warn|Info|H1|H2|H3|Debug|Verbose|Prompt        
        #Break ; 
        # should just proxy thru wo testing
        invoke-command $command ;   
    } else {
        if($Credential){
            $pltRXO = @{
                    Credential = $Credential ;
                    #verbose = $($verbose) ;
                    Verbose = $FALSE ; Silent = $false ; } ;
        } ;
        if($CredentialOP){
            $pltRX10 = @{
                    Credential = $CredentialOP ;
                    #verbose = $($verbose) ;
                    Verbose = $FALSE ; Silent = $false ; } ;
        } ;
        write-verbose "(test for test-exoToken())" ; 
        try{get-command test-exoToken | out-null }catch{
            #dx10 ;  dxo ;  dxo2 ;  daad ;  rxo2 ;  rx10 ;  caad ;
            Disconnect-Ex2010 ;  Disconnect-EXO ;  Disconnect-EXO2 ;  Disconnect-AAD ;  
            if($pltRXO){
                Reconnect-EXO2 @pltRXO ;  
            } else {Reconnect-EXO2 } 
            if($pltRX10){
                ReConnect-Ex2010 @pltRX10  ;  
            } else {Reconnect-Ex2010 } 
            if($pltRXO){
                Connect-AAD @pltRXO ;
            } else {Connect-AAD} ; 
        } ; 
        if(-not(test-exotoken)){
            Disconnect-Ex2010 ;  Disconnect-EXO ;  Disconnect-EXO2 ;  Disconnect-AAD ;  
            if($pltRXO){
                Reconnect-EXO2 @pltRXO ;  
            } else {Reconnect-EXO2 } 
            if($pltRX10){
                ReConnect-Ex2010 @pltRX10  ;  
            } else {Reconnect-Ex2010 } 
            if($pltRXO){
                Connect-AAD @pltRXO ;
            } else {Connect-AAD} ; 
        } ;
        $smsg = "invoke-command $($command) ;" ; 
        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 

        invoke-command $command ;        
    } ; 
    
    
} ;
#*------^ END Function invoke-XOWrapper ^------