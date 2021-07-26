﻿# verb-exo.psm1


  <#
  .SYNOPSIS
  verb-EXO - Powershell Exchange Online generic functions module
  .NOTES
  Version     : 1.0.89.0
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
    * 1:23 PM 5/14/2021 init version, roughed in, completely untested (was prev a largely unmodified dupe of disconnect-exo)
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
    ##Requires -Modules ActiveDirectory,verb-Auth,verb-IO,verb-Mods,verb-Text,verb-Network,verb-AAD,verb-ADMS,verb-Ex2010,verb-logging

    [CmdletBinding()]
    PARAM(
        [Parameter(Position=0,Mandatory=$True,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,HelpMessage="EXO Mailbox identifier[-mailbox 'xxx']")]
        [ValidateNotNullOrEmpty()]$Mailbox,
        [Parameter(HelpMessage="Use EXOv2 (ExchangeOnlineManagement) over basic auth legacy connection [-useEXOv2]")]
        [switch] $useEXOv2
    ) ;
    BEGIN {
        $Verbose = ($VerbosePreference -eq 'Continue') ;
        # shifting from ps1 to a function: need updates self-name:
        ${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name ;
        #-=-=configure EXO EMS aliases to cover useEXOv2 requirements-=-=-=-=-=-=
        # have to preconnect, as it gcm's the targets
        if ($script:useEXOv2) { reconnect-eXO2 }
        else { reconnect-EXO } ;
        # in this case, we need an alias for EXO, and non-alias for EXOP
        [array]$cmdletMaps = 'ps1GetxRcp;get-exorecipient;','ps1GetxMbx;get-exomailbox;','ps1SetxMbx;Set-exoMailbox;','ps1GetxUser;get-exoUser;',
            'ps1SetxCalProc;set-exoCalendarprocessing;','ps1GetxCalProc;get-exoCalendarprocessing;','ps1GetxMbxFldrPerm;get-exoMailboxfolderpermission;',
            'ps1GetxAccDom;Get-exoAcceptedDomain;','ps1GetxRetPol;Get-exoRetentionPolicy','ps1GetxDistGrp;get-exoDistributionGroup;',
            'ps1GetxDistGrpMbr;get-exoDistributionGroupmember;','ps1GetxMsgTrc;get-exoMessageTrace;','ps1GetxMsgTrcDtl;get-exoMessageTraceDetail;',
            'ps1GetxMbxFldrStats;get-exoMailboxfolderStatistics','ps1GetxMContact;Get-exomailcontact;','ps1SetxMContact;Set-exomailcontact;',
            'ps1NewxMContact;New-exomailcontact;' ,'ps1TestxMapi;Test-exoMAPIConnectivity','ps1GetxOrgCfg;Get-exoOrganizationConfig' ;
        foreach($cmdletMap in $cmdletMaps){
            if($script:useEXOv2){
                if(!($cmdlet= Get-Command $cmdletMap.split(';')[1].replace('-exo','-xo') )){ throw "unable to gcm Alias definition!:$($cmdletMap.split(';')[1])" ; break }
                $nalias = set-alias -name ($cmdletMap.split(';')[0]) -value ($cmdlet.name) -passthru ;
                write-verbose "$($nalias.Name) -> $($nalias.ResolvedCommandName)" ;
            } else {
                if(!($cmdlet= Get-Command $cmdletMap.split(';')[1])){ throw "unable to gcm Alias definition!:$($cmdletMap.split(';')[1])" ; break }
                $nalias = set-alias -name ($cmdletMap.split(';')[0]) -value ($cmdlet.name) -passthru ;                
                write-verbose "$($nalias.Name) -> $($nalias.ResolvedCommandName)" ;
            } ;
        } ;
    
        # shifting from ps1 to a function: need updates self-name:
        ${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name ;

        #$sBnr="#*======v START PASS:$($ScriptBaseName) v======" ; 
        $sBnr="#*======v START PASS:$(${CmdletName}) v======" ; 
        $smsg= $sBnr ;   
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

        # Clear error variable
        $Error.Clear() ;
        

    } ;  # BEGIN-E
    PROCESS {
         <#
        	# chk mbx-level holds
	        Rxo ; 
	        +���[LYN-8DCZ1G2]��������[Thu 11/05/2020 14:38]���
	        +[kadriTSS]::[PS]:C:\u\w\e\scripts$ get-exomailbox Helen.Gotzian@toro.com | FL LitigationHoldEnabled,InPlaceHolds
	        LitigationHoldEnabled : False
	        InPlaceHolds          : {}
	        # expand per arti
	        +���[LYN-8DCZ1G2]��������[Thu 11/05/2020 14:39]���
	        +[kadriTSS]::[PS]:C:\u\w\e\scripts$ get-exomailbox Helen.Gotzian@toro.com  | Select-Object -ExpandProperty InPlaceHolds
	        # nothing
	 
	        # check for org hold
	        +���[LYN-8DCZ1G2]��������[Thu 11/05/2020 14:39]���
	        +[kadriTSS]::[PS]:C:\u\w\e\scripts$ Get-exoOrganizationConfig | FL InPlaceHolds
	        InPlaceHolds : {}
	        # expand spec
	        +���[LYN-8DCZ1G2]��������[Thu 11/05/2020 14:40]���
	        +[kadriTSS]::[PS]:C:\u\w\e\scripts$ Get-exoOrganizationConfig | select -expand InPlaceHolds
	        # nothing
	        # check compliancetaghold (per above)
	        +���[LYN-8DCZ1G2]��������[Thu 11/05/2020 14:44]���
	        +[kadriTSS]::[PS]:C:\u\w\e\scripts$ get-exomailbox Helen.Gotzian@toro.com  |FL ComplianceTagHoldApplied
	        ComplianceTagHoldApplied : False
	 
	        No holds above.
	 
	        # eDiscovery holds � appears to require the GUID from one of the blank values above.(can't check)
	        If had it, my run on the details would be:
	        connect-ccms ; 
	        $CaseHold = Get-ccCaseHoldPolicy <hold GUID without prefix> ; 
	        Get-ccComplianceCase $CaseHold.CaseId | FL Name ; 
	        $CaseHold | FL Name,ExchangeLocation ; 
	        Get-exoMailboxSearch -InPlaceHoldIdentity <hold GUID> | FL Name,SourceMailboxes
	        # check RetentionCompliancePolicy
	        Get-ccRetentionCompliancePolicy <hold GUID without prefix or suffix> -DistributionDetail  | FL Name,*Location
	 
	        # check compliancetaghold in mbx:
	        +���[LYN-8DCZ1G2]��������[Thu 11/05/2020 14:44]���
	        +[kadriTSS]::[PS]:C:\u\w\e\scripts$ get-exomailbox Helen.Gotzian@toro.com  |FL ComplianceTagHoldApplied
	        ComplianceTagHoldApplied : False
	 
	        Erm, did anyone *read* the following on holds in the above article?:
	        This appears to be *routine* behavior per section�
	 
		        Managing mailboxes on delay hold  - https://docs.microsoft.com/en-us/microsoft-365/compliance/identify-a-hold-on-an-exchange-online-mailbox?view=o365-worldwide#managing-mailboxes-on-delay-hold
		 
		        After any type of hold is removed from a mailbox, a delay hold is applied. This means that the actual removal of the hold is delayed for 30 days to prevent data from being permanently deleted (purged) from the mailbox. This gives admins an opportunity to search for or recover mailbox items that will be purged after a hold is removed. A delay hold is placed on a mailbox the next time the Managed Folder Assistant processes the mailbox and detects that a hold was removed. Specifically, a delay hold is applied to a mailbox when the Managed Folder Assistant sets one of the following mailbox properties to True:
		                � DelayHoldApplied: This property applies to email-related content (generated by people using Outlook and Outlook on the web) that's stored in a user's mailbox.
		                � DelayReleaseHoldApplied: This property applies to cloud-based content (generated by non-Outlook apps such as Microsoft Teams, Microsoft Forms, and Microsoft Yammer) that's stored in a user's mailbox. Cloud data generated by a Microsoft app is typically stored in a hidden folder in a user's mailbox.
		        When a delay hold is placed on the mailbox (when either of the previous properties is set to True), the mailbox is still considered to be on hold for an unlimited hold duration, as if the mailbox was on Litigation Hold. After 30 days, the delay hold expires, and Microsoft 365 will automatically attempt to remove the delay hold (by setting the DelayHoldApplied or DelayReleaseHoldApplied property to False) so that the hold is removed. After either of these properties are set to False, the corresponding items that are marked for removal are purged the next time the mailbox is processed by the Managed Folder Assistant.
		        To view the values for the DelayHoldApplied and DelayReleaseHoldApplied properties for a mailbox, run the following command in Exchange Online PowerShell.
	 
	        # checking the above:
	        +���[LYN-8DCZ1G2]��������[Thu 11/05/2020 14:49]���
	        +[kadriTSS]::[PS]:C:\u\w\e\scripts$ get-exomailbox Helen.Gotzian@toro.com  | FL *HoldApplied*
	        ComplianceTagHoldApplied : False
	        DelayHoldApplied         : True
	        DelayReleaseHoldApplied  : True
        #>
        $error.clear() ;
        TRY {
            $objReturn=[ordered]@{
                Held=$false ; 
                LitigationHoldEnabled=$null ; 
                InPlaceHolds =$null ; 
                ComplianceTagHoldApplied =$null ; 
                DelayHoldApplied =$null ; 
                DelayReleaseHoldApplied =$null ; 
                OrgInPlaceHolds =$null ; 
            } ; 
            $xmbx = ps1GetxMbx -id $Mailbox -ea STOP; 
            $xOrgCfgInPlaceHolds = ps1GetxOrgCfg -ea STOP | select -expand InPlaceHolds
            if($xmbx.LitigationHoldEnabled){
                $objReturn.Held=$true ;
                $objReturn.LitigationHoldEnabled = $xmbx.LitigationHoldEnabled;
            } ; 
            if($xmbx.ComplianceTagHoldApplied){
                $objReturn.Held=$true ;
                $objReturn.ComplianceTagHoldApplied = $xmbx.ComplianceTagHoldApplied;
            } ; 
            if($xmbx.DelayHoldApplied){
                $objReturn.Held=$true ;
                $objReturn.DelayHoldApplied = $xmbx.DelayHoldApplied;
            } ; 
            if($xmbx.DelayReleaseHoldApplied){
                $objReturn.Held=$true ;
                $objReturn.DelayReleaseHoldApplied = $xmbx.DelayReleaseHoldApplied;
            } ; 
            # checking orgs: Get-exoOrganizationConfig | FL InPlaceHolds
            # reportedly expanding InPlaceHolds will return a list of mbxs, but I can't find an example of the actual return, to try to test for it.
            if(xOrgCfgInPlaceHolds){
                $objReturn.Held=$true ;
                $objReturn.OrgInPlaceHolds = $xOrgCfgInPlaceHolds;
                $smsg = "$(${CmdletName}):detected $((get-alias ps1GetxOrgCfg).definition).OrgInPlaceHolds`nbut the function is not currently written to *expand and compare* the value contents`n(requires a code update to properly work with the sample returned)" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Warn } #Error|Warn|Debug 
                else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } ; 

        } CATCH {
            $ErrTrapd=$Error[0] ;
            $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            #-=-record a STATUSWARN=-=-=-=-=-=-=
            $statusdelta = ";WARN"; # CHANGE|INCOMPLETE|ERROR|WARN|FAIL ;
            if(gv passstatus -scope Script -ea 0){$script:PassStatus += $statusdelta } ;
            if(gv -Name PassStatus_$($tenorg) -scope Script -ea 0){set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta)} ; 
            #-=-=-=-=-=-=-=-=
            $smsg = "FULL ERROR TRAPPED (EXPLICIT CATCH BLOCK WOULD LOOK LIKE): } catch[$($ErrTrapd.Exception.GetType().FullName)]{" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            Break #Opts: STOP(debug)|EXIT(close)|CONTINUE(move on in loop cycle)|BREAK(exit loop iteration)|THROW $_/'CustomMsg'(end script with Err output)
        } ; 
    } ;  # PROC-E
    END {
        $objReturn | write-output ; 
    } ;  # END-E
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
    * 1:20 PM 7/21/2021 enabled TOR titlebar tagging with TenOrg (prompt tagging by scraping TitleBar values)
    * 11:40 AM 5/14/2021 added -ea 0 to the gv tests (suppresses not-found error when called without logging config)
    * 11:43 AM 4/2/2021 updated added wlt & recstat support, updated catch blocks
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

        $TenOrg=get-TenantTag -Credential $Credential ; 
        $sTitleBarTag = @("EXO") ;
        $sTitleBarTag += $TenOrg ;
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
                $error.clear() ;
                TRY {
                    $global:EOLSession = New-PSSession @EXOsplat ;
                } CATCH {
                    Write-Warning "$(get-date -format 'HH:mm:ss'): New-PSSession: Failed processing $($_.Exception.ItemName). `nError Message: $($_.Exception.Message)`nError Details: $($_)" ;
                    $smsg = "FULL ERROR TRAPPED (EXPLICIT CATCH BLOCK WOULD LOOK LIKE): } catch[$($Error[0].Exception.GetType().FullName)]{" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    #-=-record a STATUSERROR=-=-=-=-=-=-=
                    $statusdelta = ";ERROR"; # CHANGE|INCOMPLETE|ERROR|WARN|FAIL ;
                    if(gv passstatus -scope Script -ea 0){$script:PassStatus += $statusdelta } ;
                    if(gv -Name PassStatus_$($tenorg) -scope Script -ea 0){set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta)} ; 
                    #-=-=-=-=-=-=-=-=
                    BREAK #Opts: STOP(debug)|EXIT(close)|CONTINUE(move on in loop cycle)|BREAK(exit loop iteration)|THROW $_/'CustomMsg'(end script with Err output)
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
                write-verbose "`n$((get-date).ToString('HH:mm:ss')):-PSSession w`n$(($pltiSess|out-string).trim())`n$((get-date).ToString('HH:mm:ss')):Import-Module w`n$(($pltIMod|out-string).trim())" ;
                Try {
                    # Verbose:Continue is VERY noisey for module loads. Bracketed suppress:
                    <#
                    if($VerbosePreference = "Continue"){
                        $VerbosePrefPrior = $VerbosePreference ;
                        $VerbosePreference = "SilentlyContinue" ;
                        $verbose = ($VerbosePreference -eq "Continue") ;
                    } ; 
                    #>
                    #$Global:EOLModule = Import-Module (Import-PSSession @pltiSess) -Global -Prefix $CommandPrefix -PassThru -DisableNameChecking -Verbose:$false  ;
                    $Global:EOLModule = Import-Module (Import-PSSession @pltiSess) @pltIMod  ;
                    <# reenable VerbosePreference:Continue, if set, during mod loads 
                    if($VerbosePrefPrior -eq "Continue"){
                        $VerbosePreference = $VerbosePrefPrior ;
                        $verbose = ($VerbosePreference -eq "Continue") ;
                    } ; 
                    #>
                    Add-PSTitleBar $sTitleBarTag -verbose:$($VerbosePreference -eq "Continue");
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
                    $smsg = "Get-FormatData command is not in the expected format' EXO bug: Retrying with '&SerializationLevel=Full'ConnectionUri`n(details at https://answers.microsoft.com/en-us/msoffice/forum/all/cannot-connect-to-exchange-online-via-powershell/)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-Warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    #-=-record a STATUSWARN=-=-=-=-=-=-=
                    $statusdelta = ";WARN"; # CHANGE|INCOMPLETE|ERROR|WARN|FAIL ;
                    if(gv passstatus -scope Script -ea 0){$script:PassStatus += $statusdelta } ;
                    if(gv -Name PassStatus_$($tenorg) -scope Script -ea 0){set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta)} ; 
                    #-=-=-=-=-=-=-=-=
                    write-host -foregroundcolor white "`n$((get-date).ToString('HH:mm:ss')):New-PSSession w`n$(($EXOsplat|out-string).trim())" ;
                    $error.clear() ;
                    TRY {
                        $global:EOLSession | Remove-PSSession; ; 
                        $global:EOLSession = New-PSSession @EXOsplat ;
                    } CATCH {
                        $ErrTrapd=$Error[0] ;
                        $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                        else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        #-=-record a STATUSWARN=-=-=-=-=-=-=
                        $statusdelta = ";WARN"; # CHANGE|INCOMPLETE|ERROR|WARN|FAIL ;
                        if(gv passstatus -scope Script -ea 0){$script:PassStatus += $statusdelta } ;
                        if(gv -Name PassStatus_$($tenorg) -scope Script -ea 0){set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta)} ; 
                        #-=-=-=-=-=-=-=-=
                        $smsg = "FULL ERROR TRAPPED (EXPLICIT CATCH BLOCK WOULD LOOK LIKE): } catch[$($ErrTrapd.Exception.GetType().FullName)]{" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug 
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        Break #Opts: STOP(debug)|EXIT(close)|CONTINUE(move on in loop cycle)|BREAK(exit loop iteration)|THROW $_/'CustomMsg'(end script with Err output)
                    } ; 
                    
                    $pltiSess = [ordered]@{Session = $global:EOLSession ; Prefix = $CommandPrefix ; DisableNameChecking = $true  ; AllowClobber = $true ; ErrorAction = 'Stop' ; verbose=$false ;} ;
                    $pltIMod=@{Global=$true;PassThru=$true;DisableNameChecking=$true ; verbose=$false} ; # force verbose off, suppress spam in console
                    if($CommandPrefix){
                        $pltIMod.add('Prefix',$CommandPrefix) ;
                        $pltISess.add('Prefix',$CommandPrefix) ;
                    } ;
                    write-host -foregroundcolor white "`n$((get-date).ToString('HH:mm:ss')):Import-PSSession w`n$(($pltiSess|out-string).trim())" ;
                    $error.clear() ;
                    TRY {
                        $Global:EOLModule = Import-Module (Import-PSSession @pltiSess) @pltIMod   ;
                    } CATCH {
                        $ErrTrapd=$Error[0] ;
                        $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                        else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        #-=-record a STATUSWARN=-=-=-=-=-=-=
                        $statusdelta = ";WARN"; # CHANGE|INCOMPLETE|ERROR|WARN|FAIL ;
                        if(gv passstatus -scope Script -ea 0){$script:PassStatus += $statusdelta } ;
                        if(gv -Name PassStatus_$($tenorg) -scope Script -ea 0){set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta)} ; 
                        #-=-=-=-=-=-=-=-=
                        $smsg = "FULL ERROR TRAPPED (EXPLICIT CATCH BLOCK WOULD LOOK LIKE): } catch[$($ErrTrapd.Exception.GetType().FullName)]{" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug 
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        Break #Opts: STOP(debug)|EXIT(close)|CONTINUE(move on in loop cycle)|BREAK(exit loop iteration)|THROW $_/'CustomMsg'(end script with Err output)
                    } ; 
                    # reenable VerbosePreference:Continue, if set, during mod loads 
                    if($VerbosePrefPrior -eq "Continue"){
                        $VerbosePreference = $VerbosePrefPrior ;
                        $verbose = ($VerbosePreference -eq "Continue") ;
                    } ; 
                    Add-PSTitleBar $sTitleBarTag -verbose:$($VerbosePreference -eq "Continue");

                } CATCH {
                        $ErrTrapd=$Error[0] ;
                        $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                        else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        #-=-record a STATUSWARN=-=-=-=-=-=-=
                        $statusdelta = ";WARN"; # CHANGE|INCOMPLETE|ERROR|WARN|FAIL ;
                        if(gv passstatus -scope Script -ea 0){$script:PassStatus += $statusdelta } ;
                        if(gv -Name PassStatus_$($tenorg) -scope Script -ea 0){set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta)} ; 
                        #-=-=-=-=-=-=-=-=
                        $smsg = "FULL ERROR TRAPPED (EXPLICIT CATCH BLOCK WOULD LOOK LIKE): } catch[$($ErrTrapd.Exception.GetType().FullName)]{" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug 
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        Break #Opts: STOP(debug)|EXIT(close)|CONTINUE(move on in loop cycle)|BREAK(exit loop iteration)|THROW $_/'CustomMsg'(end script with Err output)
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
    # 1:31 PM 7/21/2021 revised Add-PSTitleBar $sTitleBarTag with TenOrg spec (for prompt designators)
    * 11:53 AM 4/2/2021 updated with rlt & recstat support, updated catch blocks
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

        $TenOrg = get-TenantTag -Credential $Credential ;
        $sTitleBarTag = @("EXO2") ;
        $sTitleBarTag += $TenOrg ;

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
                $error.clear() ;
                TRY {
                    # cleanup the borked attempt left half-functioning
                    #Disconnect-ExchangeOnline -confirm:$false ;
                    #Connect-ExchangeOnline @pltCXO ;
                    write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):New-ExoPSSession w`n$(($pltNEXOS|out-string).trim())" ;
                    $PSSession = New-ExoPSSession @pltNEXOS ;
                    #Add-PSTitleBar $sTitleBarTag ;
                } CATCH {
                    $ErrTrapd=$Error[0] ;
                    $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    #-=-record a STATUSWARN=-=-=-=-=-=-=
                    $statusdelta = ";WARN"; # CHANGE|INCOMPLETE|ERROR|WARN|FAIL ;
                    if(gv passstatus -scope Script -ea 0){$script:PassStatus += $statusdelta } ;
                    if(gv -Name PassStatus_$($tenorg) -scope Script -ea 0){set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta)} ; 
                    #-=-=-=-=-=-=-=-=
                    $smsg = "FULL ERROR TRAPPED (EXPLICIT CATCH BLOCK WOULD LOOK LIKE): } catch[$($ErrTrapd.Exception.GetType().FullName)]{" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    Break #Opts: STOP(debug)|EXIT(close)|CONTINUE(move on in loop cycle)|BREAK(exit loop iteration)|THROW $_/'CustomMsg'(end script with Err output)
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
                    $ErrTrapd=$Error[0] ;
                    $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    #-=-record a STATUSWARN=-=-=-=-=-=-=
                    $statusdelta = ";WARN"; # CHANGE|INCOMPLETE|ERROR|WARN|FAIL ;
                    if(gv passstatus -scope Script -ea 0){$script:PassStatus += $statusdelta } ;
                    if(gv -Name PassStatus_$($tenorg) -scope Script -ea 0){set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta)} ; 
                    #-=-=-=-=-=-=-=-=
                    $smsg = "FULL ERROR TRAPPED (EXPLICIT CATCH BLOCK WOULD LOOK LIKE): } catch[$($ErrTrapd.Exception.GetType().FullName)]{" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    Break #Opts: STOP(debug)|EXIT(close)|CONTINUE(move on in loop cycle)|BREAK(exit loop iteration)|THROW $_/'CustomMsg'(end script with Err output)
                } ; 
            } CATCH {
                $ErrTrapd=$Error[0] ;
                $smsg = "failed to connect to EXO V2 PS module`nFailed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                #-=-record a STATUSWARN=-=-=-=-=-=-=
                $statusdelta = ";WARN"; # CHANGE|INCOMPLETE|ERROR|WARN|FAIL ;
                if(gv passstatus -scope Script -ea 0){$script:PassStatus += $statusdelta } ;
                if(gv -Name PassStatus_$($tenorg) -scope Script -ea 0){set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta)} ; 
                #-=-=-=-=-=-=-=-=
                $smsg = "FULL ERROR TRAPPED (EXPLICIT CATCH BLOCK WOULD LOOK LIKE): } catch[$($ErrTrapd.Exception.GetType().FullName)]{" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                Break #Opts: STOP(debug)|EXIT(close)|CONTINUE(move on in loop cycle)|BREAK(exit loop iteration)|THROW $_/'CustomMsg'(end script with Err output)
            } ; 
            if ($error.count -ne 0) {
                if ($error[0].FullyQualifiedErrorId -eq '-2144108477,PSSessionOpenFailed') {
                    $smsg = "AUTH FAIL BAD PASSWORD? ABORTING TO AVOID LOCKOUT!" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    #-=-record a STATUSWARN=-=-=-=-=-=-=
                    $statusdelta = ";WARN"; # CHANGE|INCOMPLETE|ERROR|WARN|FAIL ;
                    if(gv passstatus -scope Script -ea 0){$script:PassStatus += $statusdelta } ;
                    if(gv -Name PassStatus_$($tenorg) -scope Script -ea 0){set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta)} ; 
                    #-=-=-=-=-=-=-=-=
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

                Add-PSTitleBar $sTitleBarTag -verbose:$($VerbosePreference -eq "Continue");;
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

#*------v convert-HistoricalSearchCSV.ps1 v------
function convert-HistoricalSearchCSV {
    <#
    .SYNOPSIS
    convert-HistoricalSearchCSV.ps1 - Summarize (to XML) or re-expand(to CSV), MS EXO HistoricalSearch csv output files, to permit MessageTrace-style parsing of the output for delivery patterns.
    .NOTES
    Version     : 1.0.0
    Author      : Todd Kadrie
    Website     : http://www.toddomation.com
    Twitter     : @tostka / http://twitter.com/tostka
    CreatedDate : 2021-04-23
    FileName    : convert-HistoricalSearchCSV.ps1
    License     : MIT License
    Copyright   : (c) 2021 Todd Kadrie
    Github      : https://github.com/tostka/verb-EXO
    Tags        : Powershell
    AddedCredit : REFERENCE
    AddedWebsite: URL
    AddedTwitter: URL
    REVISIONS
    * 2:54 PM 4/23/2021 wrote as freestanding .ps1, decided to flip it into func in verb-EXO
    .DESCRIPTION
    convert-HistoricalSearchCSV - Summarize (to XML) or re-expand(to CSV), MS EXO HistoricalSearch csv output files, to permit MessageTrace-style parsing of the output for delivery patterns.
    Issue is that HistoricalSearch csv files summarize a lot of detail from the normal MessageTrace .csv output, into the single Recipient_status field,
    which is a concatonated combo of every reciopient, double-hash (##) delimited with the following information per recipient
    <email address>##<status>
    And there can be a series of Status entries logged, for the single email address.

    - If ToXML is chosen, the RecipientAddress & RecipientEvents are nested as an array of CustomObjects in a field named 'RecipientStatuses'
    - If ToCsv is chosen, each transaction is unpacked back into separate 'Status' lines for each RecipientStatus (closer to the way get-MessageTrace returns records)

    The benefit of expanded CSV, over the native HS output, is you can do MessageTrace-like parsing of the output:
    $msgsx = import-csv -path path-to\file.csv ; 
    $msgsx | group status | ft -auto count,name
    Count Name
    ----- ----
      119 Receive
      117 Deliver
        2 Fail

    .PARAMETER  Files
    Array of HistoricalSearch .csv file paths[-Files c:\pathto\HistSearch.csv]
    .PARAMETER ToXML
    ToXML switch (generates nested summary XML)[-ToXML]
    .PARAMETER ToCSV
    ToCSV switch (Defaults True ; expands transactions into a logged entry per RecipientStatus)[-ToCSV]
    .PARAMETER DoDots
    Use progress dotcrawl over explicit x/y echo.
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    outputs .csv or .xml with variant [originalname]-Expanded.[ext] filename of source .csv file.
    Returns a string filepath to pipeline for each converted file.
    .EXAMPLE
    convert-HistoricalSearchCSV.ps1 -ToXML -Files "C:\usr\work\incid\123456-fname.lname@domain.com-EXOHistSrch,-60D-History,From-ANY@mssociety.org,20210222-0000AM-20210423-0919AM,run-20210423-1007AM.csv" ; 
    .EXAMPLE
    convert-HistoricalSearchCSV.ps1 -ToCSV -Files "C:\usr\work\incid\654321-MTSummary_fname.lname@domain.com-90D-History.csv" ; 
    .LINK
    https://github.com/tostka/verb-exo
    .LINK
    https://docs.microsoft.com/en-us/powershell/module/exchange/get-historicalsearch
    .LINK
    https://docs.microsoft.com/en-us/powershell/module/exchange/get-messagetrace
    #>
    #Requires -Version 3
    [CmdletBinding(DefaultParameterSetName='CSV')]
    PARAM(
        [Parameter(Position=0,Mandatory=$True,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,HelpMessage="Array of HistoricalSearch .csv file paths[-Files c:\pathto\HistSearch.csv]")]
        [ValidateNotNullOrEmpty()]$Files,
        [Parameter(ParameterSetName='XML',HelpMessage="ToXML switch (generates nested summary XML)[-ToXML]")]
        [switch] $ToXML,
        [Parameter(ParameterSetName='CSV',HelpMessage="ToCSV switch (expands transactions into a line per RecipientStatus)[-ToCSV]")]
        [switch] $ToCSV=$true,
        [Parameter(HelpMessage="Use progress dotcrawl over explicit x/y echo switch[-DoDots]")]
        [switch]$DoDots=$true 
    ) ;
    $pltXCsv = [ordered]@{
        path = $null ; 
        NoTypeInformation = $true ;
    } ; 
    foreach($file in $files){
        $error.clear() ;
        TRY {
            $ifile= gci -path $file; 
            $records = import-csv -path $ifile.fullname -Encoding Unicode ; 
        } CATCH {
            $ErrTrapd=$Error[0] ;
            $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            #-=-record a STATUSWARN=-=-=-=-=-=-=
            $statusdelta = ";WARN"; # CHANGE|INCOMPLETE|ERROR|WARN|FAIL ;
            if(gv passstatus -scope Script -ea 0){$script:PassStatus += $statusdelta } ;
            if(gv -Name PassStatus_$($tenorg) -scope Script -ea 0){set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta)} ; 
            #-=-=-=-=-=-=-=-=
            $smsg = "FULL ERROR TRAPPED (EXPLICIT CATCH BLOCK WOULD LOOK LIKE): } catch[$($ErrTrapd.Exception.GetType().FullName)]{" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            Break #Opts: STOP(debug)|EXIT(close)|CONTINUE(move on in loop cycle)|BREAK(exit loop iteration)|THROW $_/'CustomMsg'(end script with Err output)
        } ; 
        <# Recipient_status: The status of the delivery of the message to the recipient. 
        If the message was sent to multiple recipients, it will show all the recipients 
        and the corresponding status for each, in the format: <email address>##<status>.
        For example: 
        ##Receive, Send means the message was received by the service and was sent to the intended destination.
        ##Receive, Fail means the message was received by the service but delivery to the intended destination failed.
        ##Receive, Deliver means the message was received by the service and was delivered to the recipient?s mailbox.
        Multi recipients appear like:
        Fname.Lname@domain.com##Receive, Deliver;Fname.Lname@domain.com##Receive, Deliver;Fname.Lname@domain.com##Receive, Deliver;Fname.Lname@domain.com##Receive, Deliver;Fname.Lname@domain.com##Receive, Deliver;Fname.Lname@domain.com##Receive, Deliver;Fname.Lname@domain.com##Receive, Deliver;Fname.Lname@domain.com##Receive, Deliver
        #>
        $aggreg = @() ; 
        $procd = 0 ; $ttl = (($records|measure).count) ; $ino=0 ; 
        if($DoDots){write-host -foregroundcolor Red "[" -NoNewline } ; 
        foreach ($record in $records){
            $procd++ ; 
            # echo every 3rd record
            if(($procd % 3) -eq 0){
                if($DoDots){
                      $ino++ ; 
                      if(($ino % 80) -eq 0){
                        write-host "." ; $ino=0 ;
                      } else {write-host "." -NoNewLine} ;
                } else { 
                    write-host -foregroundcolor gray "$((get-date).ToString('HH:mm:ss')):($($procd)/$($ttl)):" ; 
                } ; 
            } ; 
            <# typical HistoricalSearch csv record & fields:
            origin_timestamp_utc : 2021-03-23T10:00:09.3284899Z
            sender_address       : Fname.Lname@domain.com
            recipient_status     : Fname.Lname@domain.com##Receive, Deliver;"Fname LName<fname.lname"@domain.com##Receive, Fail
            message_subject      : AW: Fwd: SOME SUBJECT 123456 22-03-2021
            total_bytes          : 49790
            message_id           : <PH0PR04MB73657A6BEBB3F89D9F4FC85A8C649@PH0PR04MB7365.namprd04.prod.outlook.com>
            network_message_id   : 81945af2-cab7-45ad-ba23-08d8ede2715d
            original_client_ip   : 123.456.789.012
            directionality       : Originating
            connector_id         : To_DefaultOpportunisticTLS
            delivery_priority    : Normal
            #>
        
            $error.clear() ;
            TRY {
                <# fields from a typical MessageTrace (emulate the same names):
                PSComputerName
                RunspaceId
                PSShowComputerName
                Organization
                MessageId
                Received
                SenderAddress
                RecipientAddress
                Subject
                Status
                ToIP
                FromIP
                Size
                MessageTraceId
                StartDate
                EndDate
                Index
                #>
                $rcpRecs = $record.recipient_status.split(';') ; # split recipients
                if($ToXML){
                    $TransSummary = [ordered]@{
                        Received=([datetime]$record.origin_timestamp_utc).ToLocalTime() ; # converting HistSearch GMT to LocalTime
                        ReceivedGMT=$record.origin_timestamp_utc ;
                        SenderAddress=$record.sender_address ;
                        RecipientAddress= $null ; 
                        Subject=$record.message_subject ;
                        Size=$record.total_bytes ;
                        MessageID=$record.message_id ;
                        OriginalClientIP=$record.original_client_ip ;
                        Directionality=$record.directionality ;
                        ConnectorID=$record.connector_id ;
                        DeliveryPriority=$record.delivery_priority ;
                    } ; 
                    $RecipientStatuses=@() ; 
                    foreach($rcpRec in $rcpRecs){
                        $statusRpt = [ordered]@{
                            RecipientAddress =  ($rcpRec -split '##')[0] ; 
                            RecipientEvents = ($rcpRec -split '##')[1] -split ', ' ; 
                        } ; 
                        $RecipientStatuses += New-Object PSObject -Property $statusRpt ; 
                    } ; 
                    $TransSummary.RecipientStatuses = $RecipientStatuses ; 
                    $aggreg += New-Object PSObject -Property $TransSummary ; 
                } elseif($ToCSV){
                    $TransSummary = [ordered]@{
                        Received=([datetime]$record.origin_timestamp_utc).ToLocalTime() ;
                        ReceivedGMT=$record.origin_timestamp_utc ;
                        SenderAddress=$record.sender_address ;
                        RecipientAddress= $null ; 
                        Subject=$record.message_subject ;
                        Size=$record.total_bytes ;
                        MessageID=$record.message_id ;
                        OriginalClientIP=$record.original_client_ip ;
                        Directionality=$record.directionality ;
                        ConnectorID=$record.connector_id ;
                        DeliveryPriority=$record.delivery_priority ;
                        Status= $null ; 
                    } ; 
                    foreach($rcpRec in $rcpRecs){
                        $TransSummary.RecipientAddress =  ($rcpRec -split '##')[0] ; 
                        foreach ($status in ($rcpRec -split '##')[1] -split ', '){
                            $TransSummary.Status = $status ;
                            $aggreg += New-Object PSObject -Property $TransSummary ; 
                        } ; 
                    } ; 
                } else { throw "neither ToCSV or ToXML specified!" } ; 
            } CATCH {
                $ErrTrapd=$Error[0] ;
                $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                Continue ; 
            } ; 
        } ; 
        if($DoDots){write-host -foregroundcolor Red "]" } ; 
        if($ToCSV){
            $pltXCsv.path = join-path -Path ($ifile.DirectoryName) -ChildPath "$($ifile.BaseName)-EXPANDED$($ifile.Extension)" ; 
            write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):export-csv w`n$(($pltXCsv |out-string).trim())" ; 
            $aggreg | export-csv @pltXCsv ;
            $pltXCsv.path | write-output ;
        } elseif ($ToXML){
            $opath = join-path -Path ($ifile.DirectoryName) -ChildPath "$($ifile.BaseName)-EXPANDED.XML" ; 
            write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):export-cliXML to`n$(($opath|out-string).trim())" ; 
            $aggreg | export-clixml -Path $opath  ;
            $opath | write-output ;
        } else { 

        } ; 
    } ;  # loop-E $files
}

#*------^ convert-HistoricalSearchCSV.ps1 ^------

#*------v cxo2cmw.ps1 v------
function cxo2cmw {
    <#
    .SYNOPSIS
    cxo2CMW - Connect-EXO to specified Tenant
    .NOTES
    REVISIONS
    * 10:16 AM 7/20/2021 reverted old typo (missing '[exo]2' in call)
    .DESCRIPTION
    Connect-EXO2 - Establish PSS to EXO V2 Modern Auth PS
    .EXAMPLE
    cxo2cmw
    #>
    Connect-EXO2 -cred $credO365CMWCSID -Verbose:($VerbosePreference -eq 'Continue') ; 
}

#*------^ cxo2cmw.ps1 ^------

#*------v cxo2tol.ps1 v------
function cxo2TOL {
    <#
    .SYNOPSIS
    cxo2TOL - Connect-EXO to specified Tenant
    .NOTES
    REVISIONS
    * 10:16 AM 7/20/2021 reverted old typo (missing '[exo]2' in call)
    .DESCRIPTION
    Connect-EXO2 - Establish PSS to EXO V2 Modern Auth PS
    .EXAMPLE
    cxo2TOL
    #>
    Connect-EXO2 -cred $credO365TOLSID -Verbose:($VerbosePreference -eq 'Continue') ;
}

#*------^ cxo2tol.ps1 ^------

#*------v cxo2tor.ps1 v------
function cxo2TOR {
    <#
    .SYNOPSIS
    cxo2TOR - Connect-EXO to specified Tenant
    .NOTES
    REVISIONS
    * 10:16 AM 7/20/2021 reverted old typo (missing '[exo]2' in call)
    .DESCRIPTION
    Connect-EXO2 - Establish PSS to EXO V2 Modern Auth PS
    .EXAMPLE
    cxo2TOR
    #>
    Connect-EXO2 -cred $credO365TORSID -Verbose:($VerbosePreference -eq 'Continue') ; 
}

#*------^ cxo2tor.ps1 ^------

#*------v cxo2ven.ps1 v------
function cxo2VEN {
    <#
    .SYNOPSIS
    cxo2VEN - Connect-EXO to specified Tenant
    .NOTES
    REVISIONS
    * 10:16 AM 7/20/2021 reverted old typo (missing '[exo]2' in call)
    .DESCRIPTION
    Connect-EXO2 - Establish PSS to EXO V2 Modern Auth PS
    .EXAMPLE
    cxo2VEN
    #>
    Connect-EXO2 -cred $credO365VENCSID -Verbose:($VerbosePreference -eq 'Continue') ; 
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
    Remove-PSTitlebar 'EXO' -verbose:$($VerbosePreference -eq "Continue");
    
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
    Remove-PSTitlebar 'EXO' -verbose:$($VerbosePreference -eq "Continue");
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
    Remove-PSTitlebar 'EXO' -verbose:$($VerbosePreference -eq "Continue");
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
    # 12:17 PM 5/14/2021 updated passstatus code to curr, and added -ea to the gv's (suppress errors when not present)
    * 11:54 AM 4/2/2021 updated wlt & recstat support, updated catch blocks
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
            if(gv passstatus -scope Script -ea 0){$script:PassStatus += $statusdelta } ;
            if(gv -Name PassStatus_$($tenorg) -scope Script -ea 0){set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta)} ; 
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
                if(gv passstatus -scope Script -ea 0){$script:PassStatus += $statusdelta } ;
                if(gv -Name PassStatus_$($tenorg) -scope Script -ea 0){set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta)} ; 
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
            $ErrTrapd=$Error[0] ;
            $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            #-=-record a STATUSWARN=-=-=-=-=-=-=
            $statusdelta = ";WARN"; # CHANGE|INCOMPLETE|ERROR|WARN|FAIL ;
            if(gv passstatus -scope Script -ea 0){$script:PassStatus += $statusdelta } ;
            if(gv -Name PassStatus_$($tenorg) -scope Script -ea 0){set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta)} ; 
            #-=-=-=-=-=-=-=-=
            $smsg = "FULL ERROR TRAPPED (EXPLICIT CATCH BLOCK WOULD LOOK LIKE): } catch[$($ErrTrapd.Exception.GetType().FullName)]{" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            Break #Opts: STOP(debug)|EXIT(close)|CONTINUE(move on in loop cycle)|BREAK(exit loop iteration)|THROW $_/'CustomMsg'(end script with Err output)
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
     * 12:17 PM 5/14/2021 updated passstatus code to curr, and added -ea to the gv's (suppress errors when not present)
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
                if(gv passstatus -scope Script -ea 0){$script:PassStatus += $statusdelta } ;
                if(gv -Name PassStatus_$($tenorg) -scope Script -ea 0){set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta)} ; 
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
                if(gv passstatus -scope Script -ea 0){$script:PassStatus += $statusdelta } ;
                if(gv -Name PassStatus_$($tenorg) -scope Script -ea 0){set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta)} ; 
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

#*------v move-MailboxToXo.ps1 v------
function move-MailboxToXo{
    <#
    .SYNOPSIS
    move-MailboxToXo.ps1 - EX Hybrid Onprem-> EXO mailbox move
    .NOTES
    Version: 1.1.13
    Author: Todd Kadrie
    Website:	http://www.toddomation.com
    Twitter:	@tostka, http://twitter.com/tostka
    REVISIONS   :
    # 10:46 AM 6/2/2021 sub'd verb-logging for v-trans
    * 1:54 PM 5/19/2021 expanded get-hybridcred to use both esvc & sid userroles
    * 11:40 AM 5/14/2021 added -ea 0 to the gv tests (suppresses not-found error when called without logging config)
    # 11:11 AM 5/7/2021 replaced verbose & bdebugs ; fixed missing logging function & added trailing echo of path.
    * 3:55 PM 5/6/2021 update logging, with move to module, it's logging into the ALlusers profile dir
    * 11:05 AM 5/5/2021 refactor into a function, add to verb-exo; ren move-ExoMailboxNow -> move-MailboxToXo
    * 9:54 AM 5/5/2021 removed comment dates; rem'd spurious -domainc param on get-exorecipient
    * 3:15 PM 5/4/2021 v1.1.10 added trailing |?{$_.length} bug workaround for get-gcfastxo.ps1
    * 1:36 PM 3/31/2021 updated rgx to incl jumpboxes; repairs to new cross-org material (subseq to maint-RetPol changes/breaks) ; For some reason nowe requires ; rewrote all connect-reconn's to use central def'd splats AcceptLargeDataLoss in the movereq (wasn't there prev, but worked) ; had shifted out of domaincontroller use (due to new use of get-addomainctontroller returning array's instead of single entries, worked around the issue, put dc's back into queries, to cut down on retry/dawdle. ; updated the re-publish script under description, now echo's requiredvers
    * 9:41 AM 3/30/2021 publishing
    * 9:23 AM 11/17/2020 lots of minor revs, made sure all
    (re)connect-(ex2010|exo)'s had creds, and supported Exov2 switch, replaced
    approp write-verbose -> pswlt blocks ; updated catches to full cred & pswlt
    support ; dbg'd whatif succeessfully. Successfully moved TestNewGenericTodd
    * 9:07 AM 10/26/2020 added EXOv2 support (param, and if/then calls) ;
    * 3:17 PM 10/23/2020 exported xo dc code to new verb-adms:get-gcfastXO() ; now
    fully XO-supporting, determins MigEndPt by cycling $xxxMeta.MepArry
    (semi-colon-delimited string of MEP name;fqdn;regex - regex matches regional
    server DB names and steers the mbx MEP to match).
    * 4:13 PM 10/21/2020 midway through multi-tenant update, left off with:still needs ADMS mount-ADForestDrives() and set-location code @ 395 (had to recode mount-admforestdrives and debug cred production code & infra-string inputs before it would work; will need to dupe to suspend variant on final completion
    * 11:48 AM 4/27/2020 updated for jumpbox/published/installed script, genericized
    * 1:46 PM 2/4/2020 #920 Remove-exoMoveRequest: doesn't like $true on force
    * 12:35 PM 11/26/2019 debugged through issues created by genericizing the domain lookups - still dependant on hard MEP choice mappings on db names - added a default mep chooser that just takes index-0 on the list. Better'n nothing but if CMW has multi-meps, it could pick the wrong one.
    * 12:23 PM 11/22/2019 partial port for x-tenant use, need to get the OPCred spec pulled out of globals too...
    * 7:30 AM 8/14/2019 `$BatchName:revise & add SID
    * 9:38 AM 5/8/2019 added -NoTEST param, to skip MEP tests (no point running 2x, after initial whatif pass)
    * 9:39 AM 1/21/2019 final form used to migrate forwarded mbxs to EXO - RetentPolicy code disabled until we have DL's stocked with population of each target policy, for exempting if they turn up Default DRM Retention Policy again.
    * 8:27 PM 1/18/2019 added -BatchName param, to append later items to an existing Batch, without auto-gening the batchname
    * 7:24 PM 1/18/2019 a LOT of updates, used for forarded mbx moves, 97 - would have used MigrationBatch, but needed to use monitor-exoxxx.ps1 to catch moverequest completion and migrate fowd on the fly
    * 3:59 PM 1/18/2019 spliced in the functions & modloads from convert-OPUserGenerics2Shared.ps1 and the write-log() from check-exolicense.ps1.
    * 10:58 AM 11/27/2018 updated the trailing batch status echo
    * 2:07 PM 11/16/2018 fix typo and incomplete Test-exoMigrationServerAvailability  code, now failure causes it to retsta all MEPs, also removed duped credential prompts
    * 2:19 PM 10/9/2018 correct help text for TargetMailboxes - it's an array, not a comma-delim list
    * 2:49 PM 9/19/2018 added -CompleteEveningOf support and self-determining 6pm/5:30($ADSiteCodeAU) Cutover targets by region/MEP
    * 2:50 PM 9/5/2018 added existingmove pre-removal validated
    * 1:57 PM 9/5/2018 added echo'd replicated-monitoring cmd to the mix, updated move test to check that EXOrecipient has the $TORMeta['o365_TenantDomainMail'] address, or abort move
    * 1:12 PM 8/29/2018 added @$TORMeta['o365_TenantDom'] email addr test and auto-fix, echo of monitoring command to console
    * 1:36 PM 8/27/2018 Added UPN identify (was throwing error: The operation couldn't be performed because 'servicedesk' matches multiple entries.). Also looking up explicit $tmbx obj to pull values from, pretest, supporess prompt for predefined $OPcred ;
    * ident code both move-EXOmailboxNow.ps1 move-EXOmailboxSuspend.ps1 & , only diff is Suspend uses: SuspendWhenReadyToComplete=$TRUE in following;
    $MvSplat=@{
        Identity=$tmbx.userprincipalname ;
        remote=$true;
        RemoteHostName="MEPFQDN.DOMAIN.com" ;
        TargetDeliveryDomain="$TORMeta['o365_TenantDomainMail']" ;
        RemoteCredential=$OPcred ;
        BadItemLimit=1000 ;
        Suspend=$false ;
        SuspendWhenReadyToComplete=$false ;
        BatchName=$Batchname ;
        whatif=$($whatif) ;
    } ;
    .DESCRIPTION
    move-MailboxToXo.ps1 - Non-Suspend Onprem-> EXO mailbox move
    .PARAMETER TargetMailboxes
    Mailbox identifiers(array)[-Targetmailboxes]
    .PARAMETER BatchFile
    CSV file of mailbox descriptors, including at least PrimarySMTPAddress field [-BatchFile c:\path-to\file.csv]
    .PARAMETER BatchName
    Hard-code MoveRequest BatchName
    .PARAMETER Suspend
    Suspend move on creation Flag [-Suspend]
    .PARAMETER NoTEST
    NoTest Flag [-NoTEST]
    .PARAMETER Credential
    Credential to use for this connection [-credential [credential obj variable]")][System.Management.Automation.PSCredential]
    .PARAMETER UserRole
    Role of account (SID|CSID|UID|B2BI|CSVC|ESvc|LSvc)[-UserRole SID]
    .PARAMETER useEXOv2
    Use EXOv2 (ExchangeOnlineManagement) over basic auth legacy connection [-useEXOv2]
    .PARAMETER ShowDebug
    Parameter to display Debugging messages [-ShowDebug switch]
    .PARAMETER whatif
    Whatif Flag (DEFAULTED TRUE!) [-whatIf]
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    None. Returns no objects or output.
    .EXAMPLE
    move-MailboxToXo.ps1 -TargetMailboxes GTC-Ventrac@toro.com -showDebug  -whatIf ;
    Perform immediate move of specified mailbox, with debug output & whtif pass
    .EXAMPLE
    move-MailboxToXo.ps1 -TargetMailboxes GTC-Ventrac@toro.com -showDebug -notest -whatIf ;
    Perform immediate move of specified mailbox, suppress MEP tests (-NoTest), showdebug output & whatif pass
    .LINK
    #>
    #Requires -Modules ActiveDirectory, ExchangeOnlineManagement, verb-ADMS, verb-Ex2010, verb-IO, verb-logging, verb-Mods, verb-Network, verb-Text, verb-logging
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$FALSE,HelpMessage="TenantTag value, indicating Tenants to connect to[-TenOrg 'TOL']")]
        [ValidateNotNullOrEmpty()]
        $TenOrg = 'TOR',
        [Parameter(Position=0,Mandatory=$False,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,HelpMessage="Mailbox identifiers(array)[-Targetmailboxes]")]
        [ValidateNotNullOrEmpty()]$TargetMailboxes,
        [Parameter(Mandatory=$false,HelpMessage="CSV file of mailbox descriptors, including at least PrimarySMTPAddress field [-BatchFile c:\path-to\file.csv]")]
        [ValidateScript({Test-Path $_})][string]$BatchFile,
        [Parameter(Position=0,HelpMessage="Hard-code MoveRequest BatchName")]
        [string]$BatchName,
        [Parameter(HelpMessage="Suspend move on creation Flag [-Suspend]")]
        [switch] $Suspend,
        [Parameter(HelpMessage="NoTest Flag [-NoTEST]")]
        [switch] $NoTEST,
        [Parameter(HelpMessage="Credential to use for cloud actions [-credential [credential obj variable]")][System.Management.Automation.PSCredential]
        $Credential,
        # = $global:$credO365TORSID,
        [Parameter(HelpMessage = "Role of account (SID|CSID|UID|B2BI|CSVC|ESvc|LSvc)[-UserRole SID]")]
        [ValidateSet('SID','CSID','UID','B2BI','CSVC')]
        [string]$UserRole='SID',
        [Parameter(HelpMessage="Use EXOv2 (ExchangeOnlineManagement) over basic auth legacy connection [-useEXOv2]")]
        [switch] $useEXOv2,
        [Parameter(HelpMessage="Unpromtped run Flag [-showDebug]")]
        [switch] $NoPrompt,
        [Parameter(HelpMessage="Debugging Flag [-showDebug]")]
        [switch] $showDebug,
        [Parameter(HelpMessage="Whatif Flag (DEFAULTED TRUE!) [-whatIf]")]
        [switch] $whatIf=$true
    ) # PARAM BLOCK END

    $verbose = ($VerbosePreference -eq "Continue") ;

    #region INIT; # ------
    #*======v SCRIPT/DOMAIN/MACHINE/INITIALIZATION-DECLARE-BOILERPLATE v======
    # SCRIPT-CONFIG MATERIAL TO SET THE UNDERLYING $DBGPREF:
    if ($Whatif) { write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):`$Whatif is TRUE (`$whatif:$($whatif))" ; };
    # If using WMI calls, push any cred into WMI:
    #if ($Credential -ne $Null) {$WmiParameters.Credential = $Credential }  ;

    # scriptname with extension
    $ScriptDir=(Split-Path -parent $MyInvocation.MyCommand.Definition) + "\" ;
    $ScriptBaseName = (Split-Path -Leaf ((&{$myInvocation}).ScriptName))  ;
    $ScriptNameNoExt = [system.io.path]::GetFilenameWithoutExtension($MyInvocation.InvocationName) ;
    $ComputerName = $env:COMPUTERNAME ;

    # build smtpfrom & to fr Meta
    $smtpFrom = (($scriptBaseName.replace(".","-")) + "@$((Get-Variable  -name "$($TenOrg)Meta").value.o365_OPDomain)") ;
    #$smtpSubj= ("Daily Rpt: "+ (Split-Path $transcript -Leaf) + " " + [System.DateTime]::Now) ;
    $smtpSubj= "Proc Rpt:"   ;
    if($whatif) {
        $smtpSubj+="WHATIF:" ;
    } else {
        $smtpSubj+="PROD:" ;
    } ;
    $smtpSubj+= "$($ScriptBaseName):$(get-date -format 'yyyyMMdd-HHmmtt')"   ;
    #$smtpTo=$TORMeta['NotificationAddr1'] ;
    #$smtpTo=$TORMeta['NotificationDlUs'] ;
    $smtpToFailThru="todd.kadrie@toro.com"
    # one bene of looping: no module dependancy, works before modloads occur
    # pull the notifc smtpto from the xxxMeta.NotificationDlUs value
    # non-looping - $TenOrg is an input param, does't need modules to work yet
    if(!$showdebug){
        if((Get-Variable  -name "$($TenOrg)Meta").value.NotificationDlUs){
            $smtpTo = (Get-Variable  -name "$($TenOrg)Meta").value.NotificationDlUs ;
        }elseif((Get-Variable  -name "$($TenOrg)Meta").value.NotificationAddr1){
            $smtpTo = (Get-Variable  -name "$($TenOrg)Meta").value.NotificationAddr1 ;
        } else {
            $smtpTo=$smtpToFailThru;
        } ;
    } else {
        # debug pass, don't send to main dl, use NotificationAddr1    if((Get-Variable  -name "$($TenOrg)Meta").value.NotificationDlUs){
        if((Get-Variable  -name "$($TenOrg)Meta").value.NotificationAddr1){
            #set-variable -Name $meta.name -Value ((get-variable -name $meta.name).value  += @{'o365_AcceptedDomains' = (Get-exoAcceptedDomain).domainname} )
            $smtpTo = (Get-Variable  -name "$($TenOrg)Meta").value.NotificationAddr1 ;
        } else {
            $smtpTo=$smtpToFailThru ;
        } ;
    }

    $sQot = [char]34 ; $sQotS = [char]39 ;
    $NoProf=[bool]([Environment]::GetCommandLineArgs() -like '-noprofile'); # if($NoProf){# do this};

    #$ProgInterval= 500 ; # write-progress wait interval in ms
    # add gui vb prompt support
    #[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null ;
    # should use Windows.Forms where possible, more stable

    switch -regex ($env:COMPUTERNAME) {
        ($rgxMyBox) { $LocalInclDir = "c:\usr\work\exch\scripts" ; }
        ($rgxProdEx2010Servers) { $LocalInclDir = "c:\scripts" ; }
        ($rgxLabEx2010Servers) { $LocalInclDir = "c:\scripts" ; }
        ($rgxProdL13Servers) { $LocalInclDir = "c:\scripts" ; }
        ($rgxLabL13Servers) { $LocalInclDir = "c:\scripts" ; }
        ($rgxAdminJumpBoxes) {
            $LocalInclDir = (split-path $profile) ;
        }
    } ;

    #configure EXO EMS aliases to cover useEXOv2 requirements
    switch ($script:useEXOv2){
        $true {
            write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):using ExoV2 cmdlets" ;
            #reconnect-eXO2 @pltRXO ;
            set-alias ps1GetXRcp get-xorecipient ;
            set-alias ps1GetXMbx get-xomailbox ;
            set-alias ps1SetXMbx Set-xoMailbox ;
            set-alias ps1GetxUser get-xoUser ;
            set-alias ps1GetXCalProc get-xoCalendarprocessing ;
            set-alias ps1GetXMbxFldrPerm get-xoMailboxfolderpermission ;
            set-alias ps1GetXAccDom Get-xoAcceptedDomain ;
            set-alias ps1GGetXRetPol Get-xoRetentionPolicy ;
            set-alias ps1GetXDistGrp get-xoDistributionGroup ;
            set-alias ps1GetXDistGrpMbr get-xoDistributionGroupmember ;
            set-alias ps1TestXMigrSrvrAvail Test-xoMigrationServerAvailability ;
            set-alias ps1GetXMovReq get-xomoverequest ;
            set-alias ps1RmvXMovReq  Remove-xoMoveRequest ;
            set-alias ps1NewXMovReq  New-xoMoveRequest ;
            set-alias ps1GetXMovReqStats Get-xoMoveRequestStatistics
        }
        $false {
            write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):using EXO cmdlets" ;
            #reconnect-exo @pltRXO
            set-alias ps1GetXRcp get-exorecipient ;
            set-alias ps1GetXMbx get-exomailbox ;
            set-alias ps1SetXMbx Set-exoMailbox ;
            set-alias ps1GetxUser get-exoUser ;
            set-alias ps1GetXCalProc get-exoCalendarprocessing  ;
            set-alias ps1GetXMbxFldrPerm get-exoMailboxfolderpermission  ;
            set-alias ps1GetXAccDom Get-exoAcceptedDomain ;
            set-alias ps1GGetXRetPol Get-exoRetentionPolicy
            set-alias ps1GetXDistGrp get-exoDistributionGroup  ;
            set-alias ps1GetXDistGrpMbr get-exoDistributionGroupmember ;
            set-alias ps1TestXMigrSrvrAvail Test-exoMigrationServerAvailability ;
            set-alias ps1GetXMovReq get-exomoverequest ;
            set-alias ps1RmvXMovReq  Remove-exoMoveRequest ;
            set-alias ps1NewXMovReq  New-exoMoveRequest ;
            set-alias ps1GetXMovReqStats Get-exoMoveRequestStatistics
        } ;
    } ;  # SWTCH-E useEXOv2

    $Retries = 4 ;
    $RetrySleep = 5 ;
    $DawdleWait = 30 ; # wait time (secs) between dawdle checks
    $DirSyncInterval = 30 ; # AADConnect dirsync interval
    if(!$ThrottleMs){$ThrottleMs = 50 ;} ;
    $CredRole = 'CSVC' ; # role of svc to be dyn pulled from metaXXX if no -Credential spec'd,

    #$LastDays = -3 ;
    #if($LastDays -gt 0){$LastDays = $LastDays * -1 } ; # flip non-negative to negative integer, updated *-1 in the usage line

    if(!$rgxDriveBanChars){$rgxDriveBanChars = '[;~/\\\.:]' ; } ; # ;~/\.:,

    #*======v FUNCTIONS v======

    #-=-=TEMP SUPPRESS VERBOSE-=-=-=-=-=-=
    # suppress VerbosePreference:Continue, if set, during mod loads (VERY NOISEY)
    if($VerbosePreference = "Continue"){
        $VerbosePrefPrior = $VerbosePreference ;
        $VerbosePreference = "SilentlyContinue" ;
        $verbose = ($VerbosePreference -eq "Continue") ;
    } ;
    <# flip to require specs over tmod loads
    #*------v  MOD LOADS  v------
    # strings are: "[tModName];[tModFile];tModCmdlet"
    $tMods = @() ;
    $tMods+="verb-Auth;C:\sc\verb-Auth\verb-Auth\verb-Auth.psm1;get-password" ;
    $tMods+="verb-logging;C:\sc\verb-logging\verb-logging\verb-logging.psm1;write-log";
    $tMods+="verb-IO;C:\sc\verb-IO\verb-IO\verb-IO.psm1;Add-PSTitleBar" ;
    $tMods+="verb-Mods;C:\sc\verb-Mods\verb-Mods\verb-Mods.psm1;check-ReqMods" ;
    $tMods+="verb-Text;C:\sc\verb-Text\verb-Text\verb-Text.psm1;Remove-StringDiacritic" ;
    #$tMods+="verb-Desktop;C:\sc\verb-Desktop\verb-Desktop\verb-Desktop.psm1;Speak-words" ;
    $tMods+="verb-dev;C:\sc\verb-dev\verb-dev\verb-dev.psm1;Get-CommentBlocks" ;
    $tMods+="verb-Network;C:\sc\verb-Network\verb-Network\verb-Network.psm1;Send-EmailNotif" ;
    $tMods+="verb-Automation.ps1;C:\sc\verb-Automation.ps1\verb-Automation.ps1\verb-Automation.ps1.psm1;Retry-Command" ;
    #$tMods+="verb-AAD;C:\sc\verb-AAD\verb-AAD\verb-AAD.psm1;Build-AADSignErrorsHash";
    $tMods+="verb-ADMS;C:\sc\verb-ADMS\verb-ADMS\verb-ADMS.psm1;load-ADMS";
    $tMods+="verb-Ex2010;C:\sc\verb-Ex2010\verb-Ex2010\verb-Ex2010.psm1;Connect-Ex2010";
    $tMods+="verb-EXO;C:\sc\verb-EXO\verb-EXO\verb-EXO.psm1;Connect-Exo";
    #$tMods+="verb-L13;C:\sc\verb-L13\verb-L13\verb-L13.psm1;Connect-L13";
    #$tMods+="verb-Teams;C:\sc\verb-Teams\verb-Teams\verb-Teams.psm1;Connect-Teams";
    #$tMods+="verb-SOL;C:\sc\verb-SOL\verb-SOL\verb-SOL.psm1;Connect-SOL" ;
    #$tMods+="verb-Azure;C:\sc\verb-Azure\verb-Azure\verb-Azure.psm1;get-AADBearToken" ;
    foreach($tMod in $tMods){
    $tModName = $tMod.split(';')[0] ; $tModFile = $tMod.split(';')[1] ; $tModCmdlet = $tMod.split(';')[2] ;
    $smsg = "( processing `$tModName:$($tModName)`t`$tModFile:$($tModFile)`t`$tModCmdlet:$($tModCmdlet) )" ;
    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
    if($tModName -eq 'verb-Network' -OR $tModName -eq 'verb-Azure'){
        #write-host "GOTCHA!" ;
    } ;
    $lVers = get-module -name $tModName -ListAvailable -ea 0 ;
    if($lVers){   $lVers=($lVers | sort version)[-1];   try {     import-module -name $tModName -RequiredVersion $lVers.Version.tostring() -force -DisableNameChecking -Verbose:$false  } catch {     write-warning "*BROKEN INSTALLED MODULE*:$($tModName)`nBACK-LOADING DCOPY@ $($tModDFile)" ;import-module -name $tModDFile -force -DisableNameChecking -verbose:$false  } ;
    } elseif (test-path $tModFile) {
        write-warning "*NO* INSTALLED MODULE*:$($tModName)`nBACK-LOADING DCOPY@ $($tModDFile)" ;
        try {import-module -name $tModDFile -force -DisableNameChecking -Verbose:$false} # force non-verbose, suppress spam
        catch {   write-error "*FAILED* TO LOAD MODULE*:$($tModName) VIA $(tModFile) !" ;   $tModFile = "$($tModName).ps1" ;   $sLoad = (join-path -path $LocalInclDir -childpath $tModFile) ;   if (Test-Path $sLoad) {       Write-Verbose -verbose ((Get-Date).ToString("HH:mm:ss") + "LOADING:" + $sLoad) ;       . $sLoad ;       if ($showdebug) { Write-Verbose -verbose "Post $sLoad" };   } else {       $sLoad = (join-path -path $backInclDir -childpath $tModFile) ;       if (Test-Path $sLoad) {           Write-Verbose -verbose ((Get-Date).ToString("HH:mm:ss") + "LOADING:" + $sLoad) ;           . $sLoad ;           if ($showdebug) { Write-Verbose -verbose "Post $sLoad" };       } else {           Write-Warning ((Get-Date).ToString("HH:mm:ss") + ":MISSING:" + $sLoad + " EXITING...") ;           exit;       } ;   } ; } ;
    } ;
    if(!(test-path function:$tModCmdlet)){
        write-warning -verbose:$true  "UNABLE TO VALIDATE PRESENCE OF $tModCmdlet`nfailing through to `$backInclDir .ps1 version" ;
        $sLoad = (join-path -path $backInclDir -childpath "$($tModName).ps1") ;
        if (Test-Path $sLoad) {     Write-Verbose -verbose:$true ((Get-Date).ToString("HH:mm:ss") + "LOADING:" + $sLoad) ;     . $sLoad ;     if ($showdebug) { Write-Verbose -verbose "Post $sLoad" };     if(!(test-path function:$tModCmdlet)){         write-warning "$((get-date).ToString('HH:mm:ss')):FAILED TO CONFIRM `$tModCmdlet:$($tModCmdlet) FOR $($tModName)" ;     } else {          write-verbose -verbose:$true  "(confirmed $tModName loaded: $tModCmdlet present)"     }
        } else {     Write-Warning ((Get-Date).ToString("HH:mm:ss") + ":MISSING:" + $sLoad + " EXITING...") ;     exit; } ;
    } else {     write-verbose -verbose:$true  "(confirmed $tModName loaded: $tModCmdlet present)" } ;
    if($tModName -eq 'verb-logging'){

            # if($PSCommandPath){   $logspec = start-Log -Path $PSCommandPath -NoTimeStamp -Tag LASTPASS -showdebug:$($showdebug) -whatif:$($whatif) ;
#             } else {    $logspec = start-Log -Path ($MyInvocation.MyCommand.Definition) -showdebug:$($showdebug) -whatif:$($whatif) ; } ;
#             if($logspec){
#                 $stopResults = try {Stop-transcript -ErrorAction stop} catch {} ;
#                 $logging=$logspec.logging ;
#                 $logfile=$logspec.logfile ;
#                 $transcript=$logspec.transcript ;
#                 #Configure default logging from parent script name
#                 #start-transcript -Path $transcript ;
#             } else {throw "Unable to configure logging!" } ;


    } ;
    } ;  # loop-E
    #*------^ END MOD LOADS ^------
    #>
    #-=-=-=-=RE-ENABLE PRIOR VERBOSE-=-=-=-=
    # reenable VerbosePreference:Continue, if set, during mod loads
    if($VerbosePrefPrior -eq "Continue"){
        $VerbosePreference = $VerbosePrefPrior ;
        $verbose = ($VerbosePreference -eq "Continue") ;
    } ;
    #-=-=-=-=-=-=-=-=

    #*------v Function check-ReqMods  v------
    function check-ReqMods ($reqMods){    $bValidMods=$true ;    $reqMods | foreach-object {        if( !(test-path function:$_ ) ) {          write-error "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Missing $($_) function." ;          $bValidMods=$false ;        }    } ;    write-output $bValidMods ;} ;
    #*------^ END Function check-ReqMods  ^------

    #*======^ END FUNCTIONS ^======

    #*======v SUB MAIN v======

    # email trigger vari, it will be semi-delimd list of mail-triggering events
    $script:PassStatus = $null ;

    # check for $TenOrg & credential alignment
    # with credential un-defaulted, no need to compare $TenOrg & credential
    $tvari = "PassStatus_$($tenorg)" ; if(get-Variable -Name $tvari -scope Script -ea 0){Remove-Variable -Name $tvari -scope Script}
    New-Variable -Name PassStatus_$($tenorg) -scope Script -Value $null ;

    $reqMods+="Add-PSTitleBar;Remove-PSTitleBar".split(";") ;
    #Disconnect-EMSR (variant name in some ps1's for Disconnect-Ex2010)
    #$reqMods+="Reconnect-CCMS;Connect-CCMS;Disconnect-CCMS".split(";") ;
    #$reqMods+="Reconnect-SOL;Connect-SOL;Disconnect-SOL".split(";") ;
    $reqMods+="Test-TranscriptionSupported;Test-Transcribing;Stop-TranscriptLog;Start-IseTranscript;Start-TranscriptLog;get-ArchivePath;Archive-Log;Start-TranscriptLog".split(";") ;
    # add verb-automation content
    $reqMods+="retry-command".split(";") ;
    # lab, fails wo
    $reqMods+="Load-EMSSnap" ;
    # remove dupes
    $reqMods=$reqMods| select -Unique ;

    # detect profile installs (installed mod or script), and redir to stock location
            $dPref = 'd','c' ; foreach($budrv in $dpref){ if(test-path -path "$($budrv):\scripts" -ea 0 ){ break ;  } ;  } ;
            [regex]$rgxScriptsModsAllUsersScope="^$([regex]::escape([environment]::getfolderpath('ProgramFiles')))\\((Windows)*)PowerShell\\(Scripts|Modules)" ;
            [regex]$rgxScriptsModsCurrUserScope="^$([regex]::escape([environment]::getfolderpath('Mydocuments')))\\((Windows)*)PowerShell\\(Scripts|Modules)" ;
            # -Tag "($TenOrg)-LASTPASS" 
            # Tag=$lTag 
            $pltSLog = [ordered]@{ NoTimeStamp=$false ;  showdebug=$($showdebug) ;whatif=$($whatif) ;} ;
            if($PSCommandPath){
                if(($PSCommandPath -match $rgxScriptsModsAllUsersScope) -OR ($PSCommandPath -match $rgxScriptsModsCurrUserScope) ){
                    # AllUsers or CU installed script, divert into [$budrv]:\scripts (don't write logs into allusers context folder)
                    if($PSCommandPath -match '\.ps(d|m)1$'){
                        # module function: use the ${CmdletName} for childpath
                        $pltSLog.Path= (join-path -Path "$($budrv):\scripts" -ChildPath "$(${CmdletName}).ps1" )  ;
                    } else { 
                        $pltSLog.Path=(join-path -Path "$($budrv):\scripts" -ChildPath (split-path $PSCommandPath -leaf)) ;
                    } ; 
                }else {
                    $pltSLog.Path=$PSCommandPath ;
                } ;
            } else {
                if( ($MyInvocation.MyCommand.Definition -match $rgxScriptsModsAllUsersScope) -OR ($MyInvocation.MyCommand.Definition -match $rgxScriptsModsCurrUserScope) ){
                    $pltSLog.Path=(join-path -Path "$($budrv):\scripts" -ChildPath (split-path $PSCommandPath -leaf)) ;
                } else {
                    $pltSLog.Path=$MyInvocation.MyCommand.Definition ;
                } ;
            } ;
            $smsg = "start-Log w`n$(($pltSLog|out-string).trim())" ;
            if($verbose){ if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            $logspec = start-Log @pltSLog 

    # reloc batch calc, to include the substring in the log/transcript
    if(!$BatchName){
        $BatchName = "ExoMoves-$($env:USERNAME)" ;
        # include the 1st TargetMailbox fr the param, use first 12 chars (or less)
        $BatchName += '-' + @($TargetMailboxes)[0].tostring().substring(0,[System.Math]::Min(12,@($TargetMailboxes)[0].length)) ;
        if( (@($TargetMailboxes)|measure).count -gt 1 ){
            # append an ellipses to indicate multiple mbxs moved
            $BatchName += '...' ;
        } ;
        $BatchName += "-$(get-date -format 'yyyyMMdd-HHmmtt')" ;
        $smsg= "Using Dynamic BatchName:$($BatchName)" ;
    } else {
        $smsg= "Using -BatchName:$($BatchName)" ;
    } ;
    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

    if($logspec){
        $stopResults = try {Stop-transcript -ErrorAction stop} catch {} ;
        $logging=$logspec.logging ;
        $logfile=$logspec.logfile ;
        $transcript=$logspec.transcript ;
        #Configure default logging from parent script name
        # logfile                        C:\usr\work\o365\scripts\logs\move-MailboxToXo-(TOR)-LASTPASS-LOG-BATCH-WHATIF-log.txt
        # transcript                     C:\usr\work\o365\scripts\logs\move-MailboxToXo-(TOR)-LASTPASS-Transcript-BATCH-WHATIF-trans-log.txt
        $logfile = $logfile.replace('-LASTPASS','').replace('BATCH',(Remove-InvalidFileNameChars -name $BatchName )) ;
        $transcript = $transcript.replace('-LASTPASS','').replace('BATCH',(Remove-InvalidFileNameChars -name $BatchName )) ;
        if(Test-TranscriptionSupported){start-transcript -Path $transcript }
        else { write-warning "$($host.name) v$($host.version.major) does *not* support Transcription!" } ;
    } else {throw "Unable to configure logging!" } ;

    $smsg= "#*======v START PASS:$($ScriptBaseName) v======" ;
    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

    # seeing Curly & Cheech turn up in EX10 queries, pre-purge *any* AD psdrive
    if($existingADPSDrives = get-psdrive -PSProvider ActiveDirectory -ea 0){
        $smsg = "Purging *existing* AD PSDrives found:$(($existingADPSDrives| ft -auto name,provider,root,globalcatalog|out-string).trim())" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        $error.clear() ;
        TRY {
            $existingADPSDrives | remove-psdrive -Verbose:$($verbose) # -WhatIf:$($whatif) ;
        } CATCH {
            Write-Warning "$(get-date -format 'HH:mm:ss'): Failed processing $($_.Exception.ItemName). `nError Message: $($_.Exception.Message)`nError Details: $($_)" ;
            $smsg = "FULL ERROR TRAPPED (EXPLICIT CATCH BLOCK WOULD LOOK LIKE): } catch[$($Error[0].Exception.GetType().FullName)]{" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            BREAK ;
        } ;

    } ;
    # also purge the $global:ADPsDriveNames or $script:ADPsDriveNames
    if(gv -name ADPsDriveNames -scope global -ea 0){
        $error.clear() ;
        TRY {
            Remove-Variable -name ADPsDriveNames -Scope Global ;
        } CATCH {
            Write-Warning "$(get-date -format 'HH:mm:ss'): Failed processing $($_.Exception.ItemName). `nError Message: $($_.Exception.Message)`nError Details: $($_)" ;
            $smsg = "FULL ERROR TRAPPED (EXPLICIT CATCH BLOCK WOULD LOOK LIKE): } catch[$($Error[0].Exception.GetType().FullName)]{" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            BREAK ;
        } ;
    } ;
    if(gv -name ADPsDriveNames -scope script -ea 0){
        $error.clear() ;
        TRY {
            Remove-Variable -name ADPsDriveNames -Scope script ;
        } CATCH {
            Write-Warning "$(get-date -format 'HH:mm:ss'): Failed processing $($_.Exception.ItemName). `nError Message: $($_.Exception.Message)`nError Details: $($_)" ;
            $smsg = "FULL ERROR TRAPPED (EXPLICIT CATCH BLOCK WOULD LOOK LIKE): } catch[$($Error[0].Exception.GetType().FullName)]{" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            BREAK ;
        } ;
    } ;

    # $XXXMeta.ExOPAccessFromToro & Ex10Server
    # steer all onprem code
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

    # TEST
    if( !(check-ReqMods $reqMods) ) {write-error "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Missing function. EXITING." ; exit ;}  ;

    $useEXO = $true ; # non-dyn setting, drives variant EXO reconnect & query code
    if($useEXO){
        #*------v GENERIC EXO CREDS & SVC CONN BP v------
        # o365/EXO creds
        <### Usage: Type defaults to SID, if not spec'd - Note: there must be a *logged in & configured *profile*
        $o365Cred=get-TenantCredentials -TenOrg $TenOrg -verbose -userrole SID ;
        Returns a credential set for the $TenOrg Hybrid OnPrem Exchange Org
        .EXAMPLE
        $o365Cred=get-TenantCredentials -TenOrg $TenOrg -verbose -userrole CSVC ;
        Returns the CSVC Userrole credential for the $TenOrg Hybrid OnPrem Exchange Org
        .EXAMPLE
        $o365Cred=get-TenantCredentials -TenOrg $TenOrg -verbose -userrole B2BI ;
        Returns the B2BI Userrole credential for the $TenOrg Hybrid OnPrem Exchange Org
        ###>
        $o365Cred=$null ;
        <# $TenOrg is a mandetory param in this script, skip dyn resolution
        switch -regex ($env:USERDOMAIN){
            "(TORO|CMW)" {$TenOrg = $env:USERDOMAIN.substring(0,3).toupper() } ;
            "TORO-LAB" {$TenOrg = 'TOL' }
            default {
                throw "UNRECOGNIZED `$env:USERDOMAIN!:$($env:USERDOMAIN)" ;
                exit ;
            } ;
        } ;
        #>
        if($o365Cred=(get-TenantCredentials -TenOrg $TenOrg -UserRole 'CSVC','SID' -verbose:$($verbose))){
            # make it script scope, so we don't have to predetect & purge before using new-variable - except now it does [headcratch]
            $tvari = "cred$($tenorg)" ; if(get-Variable -Name $tvari -scope Script){Remove-Variable -Name $tvari -scope Script}
            New-Variable -Name cred$($tenorg) -scope Script -Value $o365Cred.cred ;
            $smsg = "Resolved $($Tenorg) `$o365cred:$($o365Cred.cred.username) (assigned to `$cred$($tenorg))" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        } else {
            $statusdelta = ";ERROR";
            $script:PassStatus += $statusdelta ;
            set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta) ;
            $smsg = "Unable to resolve $($tenorg) `$o365Cred value!"
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            throw "Unable to resolve $($tenorg) `$o365Cred value!`nEXIT!"
            exit ;
        } ;
        <### CALLS ARE IN FORM: (cred$($tenorg))
        $pltRXO = @{Credential = $Credential ; verbose = $($verbose) ; }
        $pltRXO = @{
            Credential = (Get-Variable -name cred$($tenorg) ).value ;
            verbose = $($verbose) ; } ;
        if ($script:useEXOv2) { reconnect-eXO2 @pltRXO }
        else { reconnect-EXO @pltRXO } ;
        # or with Tenant-specific cred($Tenorg) lookup
        #$pltRXO creds & .username can also be used for AzureAD connections
        Connect-AAD @pltRXO ;
        ###>
        # configure splat for connections: (see above useage)
        $pltRXO = @{
            Credential = (Get-Variable -name cred$($tenorg) ).value ;
            verbose = $($verbose) ; } ;
        #*------^ END GENERIC EXO CREDS & SVC CONN BP ^------
    } # if-E $useEXO

    if($UseOP){
        #*------v GENERIC EXOP CREDS & SRVR CONN BP v------
        # do the OP creds too
        $OPCred=$null ;
        # default to the onprem svc acct
        $pltGHOpCred=@{TenOrg=$TenOrg ;userrole='ESVC','SID'; verbose=$($verbose)} ;
        if($OPCred=(get-HybridOPCredentials @pltGHOpCred).cred){
            # make it script scope, so we don't have to predetect & purge before using new-variable
            $tvari = "cred$($tenorg)OP" ; if(get-Variable -Name $tvari -scope Script){Remove-Variable -Name $tvari -scope Script} ;
            New-Variable -Name "cred$($tenorg)OP" -scope Script -Value $OPCred ;
            $smsg = "Resolved $($Tenorg) `$OPCred:$($OPCred.username) (assigned to `$cred$($tenorg)OP)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        } else {
            $statusdelta = ";ERROR";
            $script:PassStatus += $statusdelta ;
            set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta) ;
            $smsg = "Unable to resolve get-HybridOPCredentials -TenOrg $($TenOrg) -userrole 'ESVC' value!"
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            throw "Unable to resolve $($tenorg) `$OPCred value!`nEXIT!"
            exit ;
        } ;
        $smsg= "Using EXOP cred:`$cred$($tenorg)OP:$((Get-Variable -name "cred$($tenorg)OP" ).value.username)" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        <# CALLS ARE IN FORM: (cred$($tenorg))
        $pltRX10 = @{
            Credential = (Get-Variable -name "cred$($tenorg)OP" ).value ;
            verbose = $($verbose) ; }
        Reconnect-Ex2010 @pltRX10 ; # local org conns
        #$pltRx10 creds & .username can also be used for local ADMS connections
        #>
        $pltRX10 = @{
            Credential = (Get-Variable -name "cred$($tenorg)OP" ).value ;
            verbose = $($verbose) ; } ;
        # TEST
        if( !(check-ReqMods $reqMods) ) {write-error "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Missing function. EXITING." ; exit ;}  ;
        # defer cx10/rx10, until just before get-recipients qry
        #*------^ END GENERIC EXOP CREDS & SRVR CONN BP ^------
        # connect to ExOP X10
        <#
        if($pltRX10){
            ReConnect-Ex2010 @pltRX10 ;
        } else { Reconnect-Ex2010 ; } ;
        #>
    } ;  # if-E $useEXOP


    $smsg= "Using EXOP cred:$($pltRXO.Credential.username)" ;
    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

    $smsg= "Using OPCred cred:$($pltRX10.Credential.username)" ;
    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

    if($VerbosePreference = "Continue"){
        $VerbosePrefPrior = $VerbosePreference ;
        $VerbosePreference = "SilentlyContinue" ;
        $verbose = ($VerbosePreference -eq "Continue") ;
    } ;
    if($pltRX10){ReConnect-Ex2010 @pltRX10 }
    else { Reconnect-Ex2010 ; } ;
    if($VerbosePrefPrior -eq "Continue"){
        $VerbosePreference = $VerbosePrefPrior ;
        $verbose = ($VerbosePreference -eq "Continue") ;
    } ;

    # load ADMS
    $reqMods+="load-ADMS".split(";") ;
    if( !(check-ReqMods $reqMods) ) {write-error "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Missing function. EXITING." ; exit ;}  ;
    $smsg = "(loading ADMS...)" ;
    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

    load-ADMS ;

    # multi-org AD
    <#still needs ADMS mount-ADForestDrives() and set-location code @ 395 (had to recode mount-admforestdrives and debug cred production code & infra-string inputs before it would work; will need to dupe to suspend variant on final completion
    #>

    if(!$global:ADPsDriveNames){
        $smsg = "(connecting X-Org AD PSDrives)" ;
        if($verbose){ if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
        $global:ADPsDriveNames = mount-ADForestDrives -verbose:$($verbose) ;
    } ;

    # EXO connection


    #$reqMods+="connect-exo;Reconnect-exo;Disconnect-exo".split(";") ;
    if( !(check-ReqMods $reqMods) ) {write-error "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Missing function. EXITING." ; exit ;}  ;
    $smsg = "(loading EXO...)" ;
    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
    #reconnect-exo -credential $pltRXO.Credential ;
    if ($script:useEXOv2) { reconnect-eXO2 @pltRXO }
    else { reconnect-EXO @pltRXO } ;


    <# RLMS connection
    $reqMods+="Get-LyncServerInSite;load-LMS;Disconnect-LMSR".split(";") ;
    if( !(check-ReqMods $reqMods) ) {write-error "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Missing function. EXITING." ; exit ;}  ;
    write-verbose -verbose:$true  "(loading LMS...)" ;
    Reconnect-L13 ;
    #>

    # forestdom is used to build the expected Tenant onmicrosoft address filter
    #$script:forestdom=((get-adforest | select -expand upnsuffixes) |?{$_ -eq (Get-Variable  -name "$($TenOrg)Meta").value.o365_OPDomain}) ;
    #-=-get-gcfastXO use to pull a root domain (or enable-exforestview for OPEX)=-=-=-=-=-
    if($UseOP){
        # suppress VerbosePreference:Continue, if set, during mod loads (VERY NOISEY)
        if($VerbosePreference -eq "Continue"){
            $VerbosePrefPrior = $VerbosePreference ;
            $VerbosePreference = "SilentlyContinue" ;
            $verbose = ($VerbosePreference -eq "Continue") ;
        } ;
        # we need $domaincontroller pop'd whether useCached or not
        # connect to ExOP X10 4:02 PM 3/23/2021 this SHOULDN'T BE A RECONNECT! should be a cold connect, *with* a pre-disconnect! I just watched it skid from TOR to CMW wo dropping connect!
        if($pltRX10){
            #Disconnect-Ex2010 -verbose:$($verbose) ;
            #get-pssession | remove-pssession ;
            # if it's not a multi-tenant process, don't pre-disconnect
            $smsg = "reconnect-Ex2010 w`n$(($pltRX10|out-string).trim())" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            ReConnect-Ex2010 @pltRX10 ;
        } else { connect-Ex2010 ; } ;
        if($VerbosePrefPrior -eq "Continue"){
            $VerbosePreference = $VerbosePrefPrior ;
            $verbose = ($VerbosePreference -eq "Continue") ;
        } ;
        # pre-clear dc, before querying
        $domaincontroller = $null ;
        # we don't know which subdoms may be in play
        pushd ; # cache pwd
        if( $tPsd = "$((Get-Variable  -name "$($TenOrg)Meta").value.ADForestName -replace $rgxDriveBanChars):" ){
            if(test-path $tPsd){
                $error.clear() ;
                TRY {
                    set-location -Path $tPsd -ea STOP ;
                    $objForest = get-adforest ;
                    $doms = @($objForest.Domains) ; # ad mod get-adforest vers
                    # do simple detect 2 doms (parent & child), use child (non-parent dom):
                    if(($doms|?{$_ -ne $objforest.name}|measure).count -eq 1){
                        $subdom = $doms|?{$_ -ne $objforest.name} ;
                        $domaincontroller = get-gcfastxo -TenOrg $TenOrg -Subdomain $subdom -verbose:$($verbose) |?{$_.length};
                        $smsg = "get-gcfastxo:returned $($domaincontroller)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } else {
                        # as this is just EX, and not AD search, open the forestview up - all Ex OP qrys will search entire forest
                        enable-forestview
                        $domaincontroller = $null ;
                    } ;
                    $script:forestdom=(($objForest | select -expand upnsuffixes) |?{$_ -eq (Get-Variable  -name "$($TenOrg)Meta").value.o365_OPDomain}) ;
                } CATCH {
                    $ErrTrapd=$Error[0] ;
                    $smsg= "Failed to exec cmd because: $($ErrTrapd)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    $smsg = "FULL ERROR TRAPPED (EXPLICIT CATCH BLOCK WOULD LOOK LIKE): } catch[$($ErrTrapd.Exception.GetType().FullName)]{" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    popd ; # restore dir
                    Exit #Opts: STOP(debug)|EXIT(close)|CONTINUE(move on in loop cycle)|BREAK(exit loop iteration)|THROW $_/'CustomMsg'(end script with Err output)
                } ;
            } else {
                $smsg = "UNABLE TO FIND *MOUNTED* AD PSDRIVE $($Tpsd) FROM `$$($TENorg)Meta!" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug
                else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                Exit #Opts: STOP(debug)|EXIT(close)|CONTINUE(move on in loop cycle)|BREAK(exit loop iteration)|THROW $_/'CustomMsg'(end script with Err output)
            } ;
        } else {
            $smsg = "UNABLE TO RESOLVE PROPER AD PSDRIVE FROM `$$($TENorg)Meta!" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug
            else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            Exit #Opts: STOP(debug)|EXIT(close)|CONTINUE(move on in loop cycle)|BREAK(exit loop iteration)|THROW $_/'CustomMsg'(end script with Err output)
        } ;
        popd ; # cd to prior dir
    } ; # $useOP
    #-=-=-=-=-=-=-

    # move the domaincontroller lookup code down into the per-user, user's subdomain is required to run get-aduser or Exch cmdlets to find the object.

    #-=-NON-SUSPEND, COMPLETE 1ST PASS:SuspendWhenReadyToComplete=$false, launch Suspend=$false=-=-=-=-=-=-=

    # fall back manual prompt creds
    <#
    if($TORMeta['o365_SIDUpn'] -AND $TORMeta['logon_SID']){
        $EXOMoveID=$TORMeta['o365_SIDUpn'] # EXO UPN-based admin sid
        $OPMoveID=$TORMeta['logon_SID'] ; # # ONPrem EX legacy fmt admin sid
    } ;
    #>
    # tenant & hyb creds already gotten above: EXO:$pltRXO.Credential & Ex10/AD:$pltRX10.Credential

    if(!$ThrottleMs){$ThrottleMs = 500 } ;#amt to throttle per pass, to *try* to stay out of throttling

    write-host "`a`n" ;
    write-host -foregroundcolor green "DID YOU PRECONVERT SECURITY GRPS & DUMP ACLS!?" ;
    write-host "`a`n" ;
    if($NoPrompt){
        $smsg = "(-NoPrompt: skipping interactive)" ;
        if($verbose){ if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
        $bRet = 'YYY'
    }else{
        $bRet=Read-Host "Enter YYY to continue. Anything else will exit"
    } ;
    if ($bRet.ToUpper() -eq "YYY") {
        Write-host "Moving on"
    } else {
        Write-Host "Invalid response. Exiting"
        # exit <asserted exit error #>
        exit 1
    } # if-block end

    <# timezone standards:
    $cstzone = [System.TimeZoneInfo]::FindSystemTimeZoneById("Central Standard Time") ;
    $AUCSTzone = [System.TimeZoneInfo]::FindSystemTimeZoneById("Cen. Australia Standard Time") ;
    $GMTzone = [System.TimeZoneInfo]::FindSystemTimeZoneById("GMT Standard Time") ;
    #>

    <#
    if($global:credo365TORSID){$exocred = $global:credo365TORSID}
    else {$exocred = Get-Credential -credential $EXOMoveID } ;

    if($global:credTORSID){$pltRX10.Credential = $global:credTORSID}
    else {$pltRX10.Credential = Get-Credential -credential $OPMoveID} ;
    #>
    if($pltRX10){
        ReConnect-Ex2010 @pltRX10 ;
    } else { Reconnect-Ex2010 ; } ;


    if($BatchFile){
        $smsg= "Using -BatchName:$($BatchName)" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        $TargetMailboxes=import-csv -path $BatchFile | select -expand PrimarySMTPAddress;
    } elseif ($Targetmailboxes){
        # defer to parameter version
    }else {
        $smsg= "MISSING `$BATCHFILE, ABORTING!" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        _cleanup
    } ;  ;

    # moved batchname calc up to logging area (to include in log)

    #if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
    #else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

    $ttl=($TargetMailboxes|measure).count ;
    $Procd=0 ;

    foreach($tMbxId in $TargetMailboxes){
        $Procd++ ;
        $sBnr="#*======v `$tmbx:($($Procd)/$($ttl)):$($tMbxId) v======" ;
        $smsg="$($sBnr)" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

        if ($script:useEXOv2) { reconnect-eXO2 @pltRXO }
        else { reconnect-EXO @pltRXO } ;

        # use new get-GCFastXO cross-org dc finder
        #$domaincontroller = get-GCFastXO -TenOrg $TenOrg -ADObject $tMbxId -verbose:$($verbose) ;
        #-=-=use new get-GCFastXO cross-org dc finder against a TenOrg and -ADObject-=-=-=-=-=-=
        $domaincontroller = $null ; # pre-clear, ensure no xo carryover
        if($tMbxId){
            # the get-addomaincontroller is returning an array; use the first item (second is blank)
            $domaincontroller = get-GCFastXO -TenOrg $TenOrg -ADObject $tMbxId -verbose:$($verbose) |?{$_.length} ;
        } else {throw "unpopulated `$TargetMailboxes parameter, unable to resolve a matching OR OP_ExADRoot property" ; } ;
        #-=-=-=-=-=-=-=-=
        # issue is that 2 objects are coming back: first is null, 2nd is the dc spec
        $Exit = 0 ;
        Do {
            Try {
                if(!(get-AdServerSettings).ViewEntireForest){ enable-ForestView } ;
                $ombx=get-mailbox -id $tMbxId -domaincontroller $domaincontroller ;
                # dc issues, drop it, and use dawdles
                #$ombx=get-mailbox -id $tMbxId  ;

                $Exit = $Retries ;
            } Catch {
                $smsg = "Failed to exec cmd because: $($Error[0])" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                Start-Sleep -Seconds $RetrySleep ;
                    if($pltRX10){
                        ReConnect-Ex2010 @pltRX10 ;
                    } else { Reconnect-Ex2010 ; } ;
                $Exit ++ ;
                $smsg = "Try #: $Exit" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" }
                If ($Exit -eq $Retries) {
                    $smsg = "Unable to exec cmd!"
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                } ;
            }  ;
        } Until ($Exit -eq $Retries) ;

        $rgxTenDomAddr = 'smtp:.*@\w*.mail.onmicrosoft.com' ;
        $rgxTenDomEAP = '^smtp:.*(\@\w*\.mail\.onmicrosoft.com)$'
        #$TenantDomainMail=$TORMeta['o365_TenantDomainMail'] ;
        $TenantDomainMail=(Get-Variable  -name "$($TenOrg)Meta").value.o365_TenantDomainMail ;

        # dynamically pull the onmicrosoft address from lowest priority policy
        $error.clear() ;
        TRY {
            ((get-emailaddresspolicy | sort priority -desc)[0] | select -expand EnabledEmailAddressTemplates|?{$_ -match $rgxTenDomEAP }) | out-null ;
            $OnMicrosoftAddrDomain=$matches[1] ;
        } CATCH {
            $smsg = "Failed processing $($_.Exception.ItemName). `nError Message: $($_.Exception.Message)`nError Details: $($_)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            Exit #Opts: STOP(debug)|EXIT(close)|CONTINUE(move on in loop cycle)|BREAK(exit loop iteration)|THROW $_/'CustomMsg'(end script with Err output)
            # -Variable -Name PassStatus_$($tenorg) -scope Script -Value $null ;
        } ;

        if($ombx){

            if($tenantCoAddr=$ombx | select -expand emailaddresses | ?{$_ -match $rgxTenDomAddr}){
                $smsg= "mbx HAS matching address:$($tenantCoAddr)" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } else {
                # moved the $OnMicrosoftAddrDomain lookup outside of the if/then
                if($OnMicrosoftAddrDomain){
                    $smsg= "mbx MISSING @TENANTDOM.mail.onmicrosoft.com:$($addr.PrimarySMTPAddress)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN} #Error|Warn|Debug
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    # precheck for addrs
                    $smsg= "SMTP Addrs:`n$(($ombx | Select -Expand EmailAddresses | ? {$_ -like "smtp:*"}|out-string).trim())" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    $ombx | Select -Expand EmailAddresses | ? {$_ -like "smtp:*"} ;
                    $smsg= "EAP Settings:`n$(($ombx|fl EmailAddressPolicyEnabled,CustomAttribute5|out-string).trim())" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    $newOnMSAddr="$($ombx.alias)$($OnMicrosoftAddrDomain)" ;
                    $spltSetmailbox=@{
                        identity=$ombx.samaccountname ;
                        #domaincontroller= $domaincontroller ;
                        EmailAddresses= @{add="$($newOnMSAddr)"} ;
                        whatif=$($whatif) ;
                    } ;

                    $smsg= "Set-Mailbox w`n$(($spltSetmailbox|out-string).trim())" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                    $smsg= "emailaddresses (expanded):`n$(($spltsetmailbox.emailaddresses |out-string).trim())" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                    $Exit = 0 ;
                    Do {
                        Try {
                            Set-Mailbox @spltSetmailbox ;
                            $Exit = $Retries ;
                        } Catch {
                            $smsg = "Failed to exec cmd because: $($Error[0])" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            Start-Sleep -Seconds $RetrySleep ;
                            if($pltRX10){
                                ReConnect-Ex2010 @pltRX10 ;
                            } else { Reconnect-Ex2010 ; } ;
                            $Exit ++ ;
                            $smsg = "Try #: $Exit" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" }
                            If ($Exit -eq $Retries) {
                                $smsg = "Unable to exec cmd!"
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            } ;
                        }  ;
                    } Until ($Exit -eq $Retries) ;

                    $1F=$false ;Do {
                        if($1F){Sleep -s 5} ;  write-host "." -NoNewLine ; $1F=$true ;
                        Try {
                            $ombx=get-Mailbox $ombx.samaccountname
                        } CATCH {
                            Write-Warning "$(get-date -format 'HH:mm:ss'): Failed processing $($_.Exception.ItemName). `nError Message: $($_.Exception.Message)`nError Details: $($_)" ;
                            $smsg = "FULL ERROR TRAPPED (EXPLICIT CATCH BLOCK WOULD LOOK LIKE): } catch[$($Error[0].Exception.GetType().FullName)]{" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            Break #Opts: STOP(debug)|EXIT(close)|CONTINUE(move on in loop cycle)|BREAK(exit loop iteration)|THROW $_/'CustomMsg'(end script with Err output)
                        } ;
                    } Until ($ombx.emailaddresses -like '*.onmicrosoft.com') ;

                    $smsg= "`nUpdated Addrs:`n$(( $ombx.EmailAddresses |?{$_ -match 'smtp:.*'}|out-string).trim())" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                    write-host -foregroundcolor yellow "$((get-date).ToString('HH:mm:ss')):MONITORING COMMAND FOR ABOVE:" ;
                    "Do {write-host '.' -NoNewLine;Start-Sleep -m (1000 * 60)} Until (Get-exoRecipient $($ombx.userprincipalname) -ea 0| select -expand emailaddresses|?{$_ -like '*@$($TenantDomainMail)'}) ; write-host '``a' "

                } else{
                    $smsg= ":mbx MISSING @$($TenantDomainMail):$($addr.PrimarySMTPAddress)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug
                    else{ write-host -foregroundcolor RED "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                } ;

                if($tenantCoAddr=$ombx | select -expand emailaddresses | ?{$_ -match [regex]'smtp:.*@$($TenantDomainMail)'}){
                    $smsg= "mbx HAS matching address:$($tenantCoAddr)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                } else {
                    $smsg= "mbx MISSING @$($TenantDomainMail):$($addr.PrimarySMTPAddress)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug
                    else{ write-host -foregroundcolor RED "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                } ;

            } ;

            $Exit = 0 ;
            Do {
                Try {
                    #$exorcp=Get-exoRecipient $($ombx.userprincipalname) -ea 0| select -expand emailaddresses|?{$_ -match $rgxTenDomAddr} ;
                    # go back to using static dc
                    $exorcp=Get-exoRecipient $($ombx.userprincipalname) ;
                    $Exit = $Retries ;
                } Catch {
                    $smsg = "Failed to exec cmd because: $($Error[0])" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    Start-Sleep -Seconds $RetrySleep ;
                    if($pltRX10){
                        ReConnect-Ex2010 @pltRX10 ;
                    } else { Reconnect-Ex2010 ; } ;
                    $Exit ++ ;
                    $smsg = "Try #: $Exit" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" }
                    If ($Exit -eq $Retries) {
                        $smsg = "Unable to exec cmd!"
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ;
                }  ;
            } Until ($Exit -eq $Retries) ;

            if($exorcp){
                # use the cached obj
                #$RemoteHostName = ($MigEnd|?{$_.remoteserver -like 'my*'})[0].remoteserver ;
                $MvSplat=[ordered]@{
                    Identity=$ombx.userprincipalname ;
                    remote=$true;
                    #RemoteHostName=$RemoteHostName ;
                    RemoteHostName=$null ;
                    TargetDeliveryDomain= $OnMicrosoftAddrDomain.replace('@','') # @$($TenantDomainMail)
                    RemoteCredential=$pltRX10.Credential ;
                    BadItemLimit=1000 ;
                    AcceptLargeDataLoss=$true ;
                    Suspend=$false ;
                    SuspendWhenReadyToComplete=$false ;
                    BatchName=$Batchname ;
                    whatif=$($whatif) ;
                } ;

                #
                $tMEPID=$null ;
                if((Get-Variable  -name "$($TenOrg)Meta").value.MEPArray){
                        #set-variable -Name $meta.name -Value ((get-variable -name $meta.name).value  += @{'o365_AcceptedDomains' = (Get-exoAcceptedDomain).domainname} )
                        #$smtpTo = (Get-Variable  -name "$($TenOrg)Meta").value.NotificationAddr1 ;
                        # loop the MepArray members, till you find the one the mbx.db matches on, then use it's MEPID
                        foreach($MA in (Get-Variable  -name "$($TenOrg)Meta").value.meparray){
                            #$ombx.Database
                            if($ombx.Database -match $MA.split(';')[2]){
                                $tMEPID = $MA.split(';')[1] ;
                                break ;
                            } ;
                        } ;
                } else {
                    $smsg = "UNABLE TO RESOLVE `$($TenOrg)Meta.MEPArray" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug
                    else{ write-host -foregroundcolor RED "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    throw $smsg ;
                } ;
                if($tMEPID){$MvSplat.RemoteHostName= $tMepID}
                else {
                    $smsg = "`$tMEPID undefined!" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug
                    else{ write-host -foregroundcolor RED "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    throw $smsg ;
                } ;

                if(!$NoTEST){
                    $tMEPID |foreach-object {
                        $error.clear() ;
                        TRY {
                            $smsg= "Testing OnPrem Admin account $($pltRX10.Credential.username) against $($_)" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            if ($script:useEXOv2) { reconnect-eXO2 @pltRXO }
                            else { reconnect-EXO @pltRXO } ;
                            $oTest= ps1TestXMigrSrvrAvail -ExchangeRemoteMove -RemoteServer $_ -Credentials $pltRX10.Credential ;
                        } CATCH {
                            $smsg= "Failed processing $($_.Exception.ItemName). `nError Message: $($_.Exception.Message)`nError Details: $($_)" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug
                            else{ write-host -foregroundcolor RED "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            $bMEPFail=$true ;
                            #Exit #Opts: STOP(debug)|EXIT(close)|Continue(move on in loop cycle)
                            Continue
                        } # try/catch-E ;
                        $smsg= $oTest ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ;
                } else {
                    write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):-NoTEST param detected: Skipping MEP tests" ;
                } ;

                <# version that keys suspend off of explicit script name
                switch -regex ($ScriptBaseName){
                    "^move-EXOmailboxSuspend\.ps1$" {
                        $smsg= "Configuring:SuspendWhenReadyToComplete=`$TRUE" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } ; #Error|Warn
                        $MvSplat.SuspendWhenReadyToComplete=$true ;
                    }
                    "^move-MailboxToXo\.ps1$" {
                        $smsg= "Configuring:SuspendWhenReadyToComplete=`$FALSE" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } ; #Error|Warn
                        $MvSplatSuspendWhenReadyToComplete=$false ;
                    }
                    default {throw "Unrecognized FILENAME:$($ScriptNameNoExt)"}
                } ;
                #>
                # switch to explicit -Suspend param (for move into verb-exo module): $Suspend
                if($Suspend){
                    $smsg= "-Suspend specified: Configuring:SuspendWhenReadyToComplete=`$TRUE" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                    else{ write-host -foregroundcolor YELLOW "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    $MvSplat.SuspendWhenReadyToComplete=$true ;
                } else {
                    $smsg= "Configuring:SuspendWhenReadyToComplete=`$FALSE" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    $MvSplat.SuspendWhenReadyToComplete=$false ;
                } ;

                $Exit = 0 ;
                Do {
                    Try {
                        $existMove=ps1GetXMovReq -Identity $mvsplat.identity -ea 0 ;
                        $Exit = $Retries ;
                    } Catch {
                        $smsg = "Failed to exec cmd because: $($Error[0])" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        Start-Sleep -Seconds $RetrySleep ;
                        if($pltRX10){
                            ReConnect-Ex2010 @pltRX10 ;
                        } else { Reconnect-Ex2010 ; } ;
                        $Exit ++ ;
                        $smsg = "Try #: $Exit" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" }
                        If ($Exit -eq $Retries) {
                            $smsg = "Unable to exec cmd!"
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        } ;
                    }  ;
                } Until ($Exit -eq $Retries) ;

                if($existMove){
                    $smsg= "==Removing ExistMove:$($existMove.alias)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                    $Exit = 0 ;
                    Do {
                        Try {
                            # add force confirm, it prompted to purge some priors
                            ps1RmvXMovReq $existMove.alias -force:$true -confirm:$false -whatif:$($whatif) ;
                            $Exit = $Retries ;
                        } Catch {
                            $smsg = "Failed to exec cmd because: $($Error[0])" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            Start-Sleep -Seconds $RetrySleep ;
                            if ($script:useEXOv2) { reconnect-eXO2 @pltRXO }
                            else { reconnect-EXO @pltRXO } ;
                            $Exit ++ ;
                            $smsg = "Try #: $Exit" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" }
                            If ($Exit -eq $Retries) {
                                $smsg = "Unable to exec cmd!"
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            } ;
                    }  ;

                    } Until ($Exit -eq $Retries) ;


                    "(waiting for move to purge)" ;
                    Do {write-host "." -NoNewLine;Start-Sleep -m (1000 * 10)} Until (!(ps1GetXMovReq -Identity $mvsplat.identity -ea 0))
                } ;

                $smsg= "===$($ombx.UserPrincipalName):$((get-alias ps1NewXMovReq).definition) w`n$(($mvSplat|out-string).trim())" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                $Exit = 0 ;
                Do {
                    Try {
                        $mvResult = ps1NewXMovReq @MvSplat;
                        <# if you don't cap output, it drops into the pipeline:
                        DisplayName           StatusDetail        TotalMailboxSize TotalArchiveSize PercentComplete
                        -----------           ------------        ---------------- ---------------- ---------------
                        Stg-Consumer Warranty WaitingForJobPickup 0 B (0 bytes)                     0
                        Stg-Dlradmin          WaitingForJobPickup 0 B (0 bytes)                     0
                        #>
                        $statusdelta = ";CHANGE";
                        $script:PassStatus += $statusdelta ;
                        set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta) ;
                        $smsg = "Move initiated:$($MvSplat.identity):`n$(($mvResult | ft -auto DisplayName,StatusDetail,TotalMailboxSize,TotalArchiveSize,PercentComplete|out-string).trim())" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                        $Exit = $Retries ;
                    } Catch {
                        $smsg = "Failed to exec cmd because: $($Error[0])" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        Start-Sleep -Seconds $RetrySleep ;
                        if ($script:useEXOv2) { reconnect-eXO2 @pltRXO }
                        else { reconnect-EXO @pltRXO } ;
                        $Exit ++ ;
                        $smsg = "Try #: $Exit" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" }
                        If ($Exit -eq $Retries) {
                            $smsg = "Unable to exec cmd!"
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        } ;
                    }  ;
                } Until ($Exit -eq $Retries) ;

            } else {
                $smsg= "===$($ombx.userprinciplname) missing $((Get-Variable  -name "$($TenOrg)Meta").value.o365_TenantDom) address at EXO END.`nSKIPPING EXO MOVE!" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug
                else{ write-host -foregroundcolor RED "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } ;
        } else {
                $smsg= "===$($tMbxId):NOT FOUND, SKIPPING!:" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug
                else{ write-host -foregroundcolor RED "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        };

        $smsg= "$($sBnr.replace('=v','=^').replace('v=','^='))" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

        # add an EXO delay to avoid issues
        Start-Sleep -Milliseconds $ThrottleMs ;
    } ;  # loop-E Mailboxes

    if(!$whatif){
        $smsg= "CLOUD MIGRATION STATUS:`n$((ps1GetXMovReq -BatchName $BatchName | ps1GetXMovReqStats | FL DisplayName,status,percentcomplete,itemstransferred,BadItemsEncountered|out-string).trim())`n" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        $smsg= "`nContinue monitoring with:`n$((get-alias ps1GetXMovReq).definition) -BatchName $($BatchName) | $((get-alias ps1GetXMovReqStats).definition) | fl DisplayName,status,percentcomplete,itemstransferred,BadItemsEncountered`n" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
    } ;
    

    # return the passstatus to the pipeline
    $script:PassStatus | write-output

    if((get-AdServerSettings).ViewEntireForest){ disable-ForestView } ;

    if($host.Name -eq "Windows PowerShell ISE Host" -and $host.version.major -lt 5){
        # 11:51 AM 9/22/2020 isev5 supports transcript, anything prior has to fake it
        # 8:46 AM 3/11/2015 shift the logfilename gen out here, so that we can arch it
        #$Logname= (join-path -path (join-path -path $scriptDir -childpath "logs") -childpath ($scriptNameNoExt + "-" + (get-date -uformat "%Y%m%d-%H%M" ) + "-ISEtrans.log")) ;
        # 2:16 PM 4/27/2015 shift to static timestamp $timeStampNow
        #$Logname= (join-path -path (join-path -path $scriptDir -childpath "logs") -childpath ($scriptNameNoExt + "-" + $timeStampNow + "-ISEtrans.log")) ;
        # 2:02 PM 9/21/2018 missing $timestampnow, hardcode
        #$Logname=(join-path -path (join-path -path $scriptDir -childpath "logs") -childpath ($scriptNameNoExt + "-" + (get-date -format 'yyyyMMdd-HHmmtt') + "-ISEtrans.log")) ;
        # maintain-ExoUsrMbxFreeBusyDetails-TOR-ForceAll-Transcript-BATCH-EXEC-20200921-1539PM-trans-log.txt
        $Logname=$transcript.replace('-trans-log.txt','-ISEtrans-log.txt') ;
        write-host "`$Logname: $Logname";
        Start-iseTranscript -logname $Logname  -Verbose:($VerbosePreference -eq 'Continue') ;
        #Archive-Log $Logname ;
        # 1:23 PM 4/23/2015 standardize processing file so that we can send a link to open the transcript for review
        $transcript = $Logname ;
        if($host.version.Major -ge 5){ stop-transcript  -Verbose:($VerbosePreference -eq 'Continue')} # ISE in psV5 actually supports transcription. If you don't stop it, it just keeps rolling
    } else {
        write-verbose "$((get-date).ToString('HH:mm:ss')):Stop Transcript" ;
        Stop-TranscriptLog -Transcript $transcript -verbose:$($VerbosePreference -eq "Continue") ;
    } # if-E
    
    # prod is still showing a running unstopped transcript, kill it again
    $stopResults = try {Stop-transcript -ErrorAction stop} catch {} ;
    
    # also trailing echo the log:
    $smsg = "`$logging:`$true:written to:`n$($logfile)" ; 
    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
    $smsg = "$((get-date).ToString('HH:mm:ss')):$($sBnr.replace('=v','=^').replace('v=','^='))" ;
    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
    #*------^ END SUB MAIN ^------
}

#*------^ move-MailboxToXo.ps1 ^------

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

#*------v resolve-Name.ps1 v------
Function resolve-Name {
    <#
    .SYNOPSIS
    resolve-Name.ps1 - Port 7nlu to a verb-EXO function. Resolves a displayname into Exchange Online/Exchange Onprem mailbox/MsolUser/AzureADUser/ADUser info, and licensing status. Detect's cross-org hybrid AD objects as well. 
    .NOTES
    Version     : 1.0.0
    Author      : Todd Kadrie
    Website     : http://www.toddomation.com
    Twitter     : @tostka / http://twitter.com/tostka
    CreatedDate : 2021-06-09
    FileName    : resolve-Name.ps1
    License     : MIT License
    Copyright   : (c) 2020 Todd Kadrie
    Github      : https://github.com/tostka/verb-EXO
    Tags        : Powershell,ExchangeOnline,Exchange,MsolUser,AzureADUser,ADUser
    REVISIONS
    * 1:17 PM 6/10/2021 added missing $exMProps add lic grp memberof check for aadu, for x-hyb users; add missing $rgxLicGrp, as $rgxLicGrpDN & $rgxLicGrpDName (aduser & aaduser respectively); pulled datestamps on echo's, simplified echo's (removed "$($smsg)")
    * 4:00 PM 6/9/2021 added alias 'nlu' (7nlu is still ahk macro) ; fixed typo; expanded echo for $lic;flipped -displayname to -identifier, and handle smtpaddr|alias|displayname lookups ; init; 
    .DESCRIPTION
    resolve-Name.ps1 - Port 7nlu to a verb-EXO function. Resolves a mailbox user Identifier into Exchange Online/Exchange Onprem mailbox/MsolUser/AzureADUser info, and licensing status. Detect's cross-org hybrid AD objects as well. 
    .PARAMETER TenOrg
    Tenant Org designator (defaults to TOR)
    .PARAMETER Identifier
    User Displayname|UPN|alias to be resolved[-Identifier 'Some Username'
    .PARAMETER Ticket
    Ticket # [-Ticket nnnnn]
    .PARAMETER useEXOv2
    Use EXOv2 (ExchangeOnlineManagement) over basic auth legacy connection [-useEXOv2]
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    .EXAMPLE
    resolve-Name -Identifier 'Some User'
    Command-line resolve displayname to summary details.
    .EXAMPLE
    resolve-Name -Identifier 'Some.User@domain.com'
    Command-line resolve email address to summary details.
    .EXAMPLE
    resolve-Name -Identifier 'alias'
    Command-line resolve mail alias value to summary details.
    .EXAMPLE
    resolve-Name
    Where no -Identifier is specified, defaults to checking clipboard for a Identifier equivelent.
    .LINK
    https://github.com/tostka/verb-EXO
    #>
    #Requires -Modules ActiveDirectory,AzureAD,MSOnline,verb-Auth,verb-IO,verb-Mods,verb-Text,verb-AAD,verb-ADMS,verb-Ex2010,verb-logging
    [CmdletBinding()]
    [Alias('nlu')]
    PARAM(
        [Parameter(Mandatory=$FALSE,HelpMessage="TenantTag value, indicating Tenants to connect to[-TenOrg 'TOR']")]
        $TenOrg = 'TOR',
        [Parameter(Position=0,Mandatory=$False,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,HelpMessage="User Identifier to be resolved[-Identifier 'Some Username'")]        
        $Identifier,
        [Parameter(HelpMessage="Use EXOv2 (ExchangeOnlineManagement) over basic auth legacy connection [-useEXOv2]")]
        [switch] $useEXOv2
    ) ;
    BEGIN {
        $Verbose = ($VerbosePreference -eq 'Continue') ;
        
        #$propsXmbx = 'UserPrincipalName','Alias','ExchangeGuid','Database','ExternalDirectoryObjectId','RemoteRecipientType'
        #$propsOPmbx = 'UserPrincipalName','SamAccountName','RecipientType','RecipientTypeDetails' ; 
        $exMProps='samaccountname','alias','windowsemailaddress','DistinguishedName''RecipientType','RecipientTypeDetails' ;

        #$adprops = "samaccountname", "msExchRemoteRecipientType", "msExchRecipientDisplayType", "msExchRecipientTypeDetails", "userprincipalname" ;
        $adprops = "samaccountname","UserPrincipalName","memberof","msExchMailboxGuid","msexchrecipientdisplaytype","msExchRecipientTypeDetails","msExchRemoteRecipientType"
        
        [regex]$rgxDname = "^[\w'\-,.][^0-9_!?????/\\+=@#$%?&*(){}|~<>;:[\]]{2,}$"
        # below doesn't encode cleanly, mainly black diamonds - better w alt font (non-lucida console)
        #"^[a-zA-Z??????acce????ei????ln??????????uu??zz??c????????ACCEE????????ILN??????????UU??ZZ????C???? ,.'-]+$"
        [regex]$rgxEmailAddr = "^([0-9a-zA-Z]+[-._+&'])*[0-9a-zA-Z]+@([-0-9a-zA-Z]+[.])+[a-zA-Z]{2,63}$"
        [regex]$rgxCMWDomain = 'DC=cmw,DC=internal$' ;
        [regex]$rgxExAlias = "^[0-9a-zA-Z-._+&]{1,64}$" ;
        # used for adu.memberof
        [regex]$rgxLicGrpDN = "^CN=ENT-APP-Office365-.*-DL,((OU=Enterprise\sApplications,)*)OU=ENTERPRISE,DC=global,DC=ad,DC=toro((lab)*),DC=com$" ;  ; 
        # used for taadu memberof
        [regex]$rgxLicGrpDName = "^ENT-APP-Office365-.*-DL((\s)*)$" ;
        #"^ENT-APP-Office365-.*-DL$" ;  
        # cute, we've got cmw AAD grps with trailing spaces: 'ENT-APP-Office365-CMWUsers-E3-DL ', pull trailing $

        if(!$Identifier -AND (gcm get-clipboard) -AND (get-clipboard)){
            $Identifier = get-clipboard ;
            #$cb = get-clipboard ; 
        } elseif($Identifier){


        } else {
            write-warning "No Identifier specified, and clipboard did not match 'Identifier' content" ; 
            Break ;
        } ; 

        <#[regex]$rgxDname = "^[\w'\-,.][^0-9_!?????/\\+=@#$%?&*(){}|~<>;:[\]]{2,}$"
        # below doesn't encode cleanly, mainly black diamonds
        #"^[a-zA-Z??????acce????ei????ln??????????uu??zz??c????????ACCEE????????ILN??????????UU??ZZ????C???? ,.'-]+$"
        [regex]$rgxEmailAddr = "^([0-9a-zA-Z]+[-._+&'])*[0-9a-zA-Z]+@([-0-9a-zA-Z]+[.])+[a-zA-Z]{2,63}$"
        #>
        $IdentifierType = $null ; 
        switch -regex ($Identifier){
            $rgxExAlias {
                write-verbose "(`$Identifier appears to be an Alias)" ;
                $IdentifierType = "Alias" ;
                #$Displayname = $Identifier.split('@')[0].replace('.',' ')
                #$nameparts = $Identifier.split('@')[0].replace('.',' ').split(' ')
                $nameparts = $Identifier.split(' ')
                break;
            } ;
            $rgxDname {
                write-verbose "(`$Identifier appears to be a DisplayName)" ;
                $IdentifierType = "DisplayName" ;
                $nameparts = $Identifier.split(' ')
                break;
            }
            $rgxEmailAddr {
                write-verbose "(`$Identifier appears to be an SmtpAddress)" ;
                $IdentifierType = "SmtpAddress" ;
                #$Displayname = $Identifier.split('@')[0].replace('.',' ')
                $nameparts = $Identifier.split('@')[0].replace('.',' ').split(' ')
                break;
            } ;
            default {
                write-warning "Unable to resolve -Identifier ($($Identifier)) into a proper DisplayName|EmailAddress|Alias string" ;
                $IdentifierType = $null ;
                break ;
            }
        } ;
        #if($Identifier -match $rgxDname){
        #        $nameparts = $Identifier.split(' ')
        switch (($nameparts|measure).count){
            "1" {
                # it's an alias
                #Identifier = vString 
                $fname = "" 
                $lname = $nameparts
            }
            "2" {
                <#/*
                RegExMatch(vString, "^\w*\s\w*$", displayname)
                RegExMatch(vString, "\w*(?=[\s])", fname)
                RegExMatch(vString, "(?<=\s)\w*$", lname)
                */
                #>
                #displayname = vString 
                $fname = $nameparts[0] ;
                $lname = $nameparts[1] ;
            }
            default{
                # assume the last 2/* are the last name ( concat no space for searches).
                #displayname = vString 
                $fname = $nameparts[0] ; 
                $lname = $nameparts[1..[int]($nameparts.getupperbound(0))] -join ' ' ;
            }
        } ;
        #} ; 
        
        $sBnr="===v Input (& splits): '$($Identifier)' | '$($fname)' | '$($lname)' v===" ;
        write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($sBnr)" ;
        #-=-=configure EXO EMS aliases to cover useEXOv2 requirements-=-=-=-=-=-=
        # have to preconnect, as it gcm's the targets
        if ($script:useEXOv2) { reconnect-eXO2 -verbose:$($verbose)}
        else { reconnect-EXO -verbose:$($verbose)} ;
        # in this case, we need an alias for EXO, and non-alias for EXOP
        [array]$cmdletMaps = 'ps1GetxRcp;get-exorecipient;','ps1GetxMbx;get-exomailbox;','ps1GetxUser;get-exoUser;'
        foreach($cmdletMap in $cmdletMaps){
            if($script:useEXOv2){
                if(!($cmdlet= Get-Command $cmdletMap.split(';')[1].replace('-exo','-xo') )){ throw "unable to gcm Alias definition!:$($cmdletMap.split(';')[1])" ; break }
                $nalias = set-alias -name ($cmdletMap.split(';')[0]) -value ($cmdlet.name) -passthru ;
                write-verbose "$($nalias.Name) -> $($nalias.ResolvedCommandName)" ;
            } else {
                if(!($cmdlet= Get-Command $cmdletMap.split(';')[1])){ throw "unable to gcm Alias definition!:$($cmdletMap.split(';')[1])" ; break }
                $nalias = set-alias -name ($cmdletMap.split(';')[0]) -value ($cmdlet.name) -passthru ;                
                write-verbose "$($nalias.Name) -> $($nalias.ResolvedCommandName)" ;
            } ;
        } ;
    
        # shifting from ps1 to a function: need updates self-name:
        ${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name ;

        #$sBnr="#*======v START PASS:$($ScriptBaseName) v======" ; 
        <#$sBnr="#*======v START PASS:$(${CmdletName}) v======" ; 
        $smsg= $sBnr ;   
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
        else{ write-host -foregroundcolor green $smsg } ;
        #>
        
        # steer all onprem code on $XXXMeta.ExOPAccessFromToro & Ex10Server values
        $UseOP=$true ; 

        $useEXO = $true ; # non-dyn setting, drives variant EXO reconnect & query code
        if($useEXO){
            #*------v GENERIC EXO CREDS & SVC CONN BP v------
            # o365/EXO creds
            <### Usage: Type defaults to SID, if not spec'd - Note: there must be a *logged in & configured *profile* 
            $o365Cred=get-TenantCredentials -TenOrg $TenOrg -verbose -userrole SID ;
            Returns a credential set for the $TenOrg Hybrid OnPrem Exchange Org
            .EXAMPLE
            $o365Cred=get-TenantCredentials -TenOrg $TenOrg -verbose -userrole CSVC ;
            Returns the CSVC Userrole credential for the $TenOrg Hybrid OnPrem Exchange Org
            .EXAMPLE
            $o365Cred=get-TenantCredentials -TenOrg $TenOrg -verbose -userrole B2BI ;
            Returns the B2BI Userrole credential for the $TenOrg Hybrid OnPrem Exchange Org
            ###>
            $o365Cred=$null ;
            <# $TenOrg is a mandetory param in this script, skip dyn resolution
            switch -regex ($env:USERDOMAIN){
                "(TORO|CMW)" {$TenOrg = $env:USERDOMAIN.substring(0,3).toupper() } ;
                "TORO-LAB" {$TenOrg = 'TOL' }
                default {
                    throw "UNRECOGNIZED `$env:USERDOMAIN!:$($env:USERDOMAIN)" ; 
                    Break ; 
                } ;
            } ; 
            #>
            if($o365Cred=(get-TenantCredentials -TenOrg $TenOrg -UserRole 'CSVC','SID' -verbose:$($verbose))){
                # make it script scope, so we don't have to predetect & purge before using new-variable
                if(get-Variable -Name cred$($tenorg) -scope Script -ea 0){remove-variable -Name cred$($tenorg) -scope Script} ; 
                New-Variable -Name cred$($tenorg) -scope Script -Value $o365Cred.cred ;
                $smsg = "Resolved $($Tenorg) `$o365cred:$($o365Cred.cred.username) (assigned to `$cred$($tenorg))" ; 
                write-verbose $smsg  ;
            } else {
                $statusdelta = ";ERROR";
                $script:PassStatus += $statusdelta ;
                set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta) ;
                $smsg = "Unable to resolve $($tenorg) `$o365Cred value!"
                write-verbose $smsg  ;
                throw "Unable to resolve $($tenorg) `$o365Cred value!`nEXIT!"
                Break ;
            } ;
            <### CALLS ARE IN FORM: (cred$($tenorg))
            $pltRXO = @{Credential = $Credential ; verbose = $($verbose) ; }
            $pltRXO = @{
                Credential = (Get-Variable -name cred$($tenorg) ).value ;
                verbose = $($verbose) ; } ; 
            if ($script:useEXOv2) { reconnect-eXO2 @pltRXO }
            else { reconnect-EXO @pltRXO } ;
            # or with Tenant-specific cred($Tenorg) lookup
            #$pltRXO creds & .username can also be used for AzureAD connections 
            Connect-AAD @pltRXO ; 
            ###>
            # configure splat for connections: (see above useage)
            $pltRXO = @{
                Credential = (Get-Variable -name cred$($tenorg) ).value ;
                verbose = $($verbose) ; } ; 
            #
            #*------^ END GENERIC EXO CREDS & SVC CONN BP ^------
        } # if-E $useEXO

        if($UseOP){
            #*------v GENERIC EXOP CREDS & SRVR CONN BP v------
            # do the OP creds too
            $OPCred=$null ;
            # default to the onprem svc acct
            $pltGHOpCred=@{TenOrg=$TenOrg ;userrole='ESVC','SID'; verbose=$($verbose)} ;
            if($OPCred=(get-HybridOPCredentials @pltGHOpCred).cred){
                # make it script scope, so we don't have to predetect & purge before using new-variable
                if(get-Variable -Name "cred$($tenorg)OP" -scope Script -ea 0){remove-variable -Name "cred$($tenorg)OP" -scope Script} ; 
                New-Variable -Name "cred$($tenorg)OP" -scope Script -Value $OPCred ;
                $smsg = "Resolved $($Tenorg) `$OPCred:$($OPCred.username) (assigned to `$cred$($tenorg)OP)" ; 
                write-verbose $smsg  ;
            } else {
                $smsg = "Unable to resolve get-HybridOPCredentials -TenOrg $($TenOrg) -userrole 'ESVC' value!"
                write-verbose $smsg  ;
                throw "Unable to resolve $($tenorg) `$OPCred value!`nEXIT!"
                Break ;
            } ;
            $smsg= "Using EXOP cred:`$cred$($tenorg)OP:$((Get-Variable -name "cred$($tenorg)OP" ).value.username)" ;  
            write-verbose $smsg  ; 
            <# CALLS ARE IN FORM: (cred$($tenorg))
            $pltRX10 = @{
                Credential = (Get-Variable -name "cred$($tenorg)OP" ).value ;
                verbose = $($verbose) ; }
            ReConnect-Ex2010XO @pltRX10 ; # cross-prem conns
            Reconnect-Ex2010 @pltRX10 ; # local org conns
            #$pltRx10 creds & .username can also be used for local ADMS connections
            #>
            $pltRX10 = @{
                Credential = (Get-Variable -name "cred$($tenorg)OP" ).value ;
                verbose = $($verbose) ; } ;     
            # TEST
        
            # defer cx10/rx10, until just before get-recipients qry
            #*------^ END GENERIC EXOP CREDS & SRVR CONN BP ^------
            # connect to ExOP X10
            if($pltRX10){
                #ReConnect-Ex2010XO @pltRX10 ;
                ReConnect-Ex2010 @pltRX10 ;
            } else { Reconnect-Ex2010 ; } ; 
        } ;  # if-E $useEXOP

        <# already confirmed in modloads
        # load ADMS
        $reqMods += "load-ADMS".split(";") ;
        if ( !(check-ReqMods $reqMods) ) { write-error "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Missing function. EXITING." ; Break ; }  ;
        #>
        write-verbose "(loading ADMS...)" ;
        # 2:12 PM 6/9/2021 load-ADMS is returning boolean, capture it
        $bRet = load-ADMS -verbose:$($verbose) ;

        if($UseOP){
            # resolve $domaincontroller dynamic, cross-org
            # setup ADMS PSDrives per tenant 
            if(!$global:ADPsDriveNames){
                $smsg = "(connecting X-Org AD PSDrives)" ;
                write-verbose $smsg  ;
                $global:ADPsDriveNames = mount-ADForestDrives -verbose:$($verbose) ;
            } ; 
            if(($global:ADPsDriveNames|measure).count){
                $useEXOforGroups = $false ; 
                $smsg = "Confirming ADMS PSDrives:`n$(($global:ADPsDriveNames.Name|%{get-psdrive -Name $_ -PSProvider ActiveDirectory} | ft -auto Name,Root,Provider|out-string).trim())" ;
                write-verbose $smsg  ;
                # returned object
                #         $ADPsDriveNames
                #         UserName                Status Name        
                #         --------                ------ ----        
                #         DOM\Samacctname   True  [forestname wo punc] 
                #         DOM\Samacctname   True  [forestname wo punc]
                #         DOM\Samacctname   True  [forestname wo punc]
        
            } else { 
                #-=-record a STATUS=-=-=-=-=-=-=
                $statusdelta = ";ERROR";
                $script:PassStatus += $statusdelta ;
                set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta) ;
                #-=-=-=-=-=-=-=-=
                $smsg = "Unable to detect POPULATED `$global:ADPsDriveNames!`n(should have multiple values, resolved to $()"
                write-warning $smsg  ;
                throw "Unable to resolve $($tenorg) `$o365Cred value!`nEXIT!"
                Break ;
            } ; 
        } ; 
        #if (!$domaincontroller) { $domaincontroller = get-gcfast } ;
        #if(!$domaincontroller){ if(test-path function:get-gcfast){$domaincontroller=get-gcfast} else { throw "no get-gcfast()!" } ;} else {"(existing `$domaincontroller:$($domaincontroller))"} ;
        # use new get-GCFastXO cross-org dc finde
        # default to Op_ExADRoot forest from $TenOrg Meta
        $domaincontroller = get-GCFastXO -TenOrg $TenOrg -subdomain ((gv -name "$($TenOrg)Meta").value['OP_ExADRoot']) -verbose:$($verbose) |?{$_.length};


        # MSOL CONNECTION
        $reqMods += "connect-msol".split(";") ;
        if ( !(check-ReqMods $reqMods) ) { write-error "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Missing function. EXITING." ; Break ; }  ;
        write-verbose "(loading AAD...)" ;
        #connect-msol ;
        connect-msol @pltRXO ; 
        #

        # AZUREAD CONNECTION
        $reqMods += "Connect-AAD".split(";") ;
        if ( !(check-ReqMods $reqMods) ) { write-error "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Missing function. EXITING." ; Break ; }  ;
        write-verbose "(loading AAD...)" ;
        #connect-msol ;
        Connect-AAD @pltRXO ; 
        #


        #
        <# EXO connection
        $pltRXO = @{
            Credential = (Get-Variable -name cred$($tenorg) ).value ;
            verbose = $($verbose) ; } ; 
        if($VerbosePreference = "Continue"){
            $VerbosePrefPrior = $VerbosePreference ;
            $VerbosePreference = "SilentlyContinue" ;
            $verbose = ($VerbosePreference -eq "Continue") ;
        } ; 
        #disconnect-exo ;
        if ($script:useEXOv2) { reconnect-eXO2 @pltRXO }
        else { reconnect-EXO @pltRXO } ;
        # reenable VerbosePreference:Continue, if set, during mod loads 
        if($VerbosePrefPrior -eq "Continue"){
            $VerbosePreference = $VerbosePrefPrior ;
            $verbose = ($VerbosePreference -eq "Continue") ;
        } ;
        #>

        
        # 3:00 PM 9/12/2018 shift this to 1x in the script ; - this would need to be customized per tenant, not used (would normally be for forcing UPNs, but CMW uses brand UPN doms)
        #$script:forestdom = ((get-adforest | select -expand upnsuffixes) | ? { $_ -eq 'toro.com' }) ;

        # Clear error variable
        $Error.Clear() ;
        

    } ;  # BEGIN-E
    PROCESS {
        <#$IdentifierType = "DisplayName" ;
        $IdentifierType = "SmtpAddress" ;
        $IdentifierType = "Alias" ;
        #>
        $pltGetxUser=[ordered]@{
            ErrorAction = 'STOP' ;
        } ;
        switch -regex($IdentifierType){
            '(Alias|SmtpAddress)'{
                $pltGetxUser.add('Identity',$Identifier) ;
            }
            'DisplayName'{
                $fltr = "displayname -like '$Identifier'" ; 
                $pltGetxUser.add('filter',$fltr) ;
            }
            default {
                write-warning "Unable to resolve `$IdentifierType ($($IdentifierType)) into a recognized value" ;
                break ;
            }
        } ;

        write-verbose "$((get-alias ps1GetxUser).definition) w`n$(($pltGetxUser|out-string).trim())" ;         
        #rxo ; cmsol ; caad ; rx10 ;
        $error.clear() ;
        TRY {
            $txUser =ps1GetxUser @pltGetxUser ;
            if($msolu = get-msoluser -user $txUser.UserPrincipalName |?{$_.islicensed}){
            #if($msolu = get-msoluser -user $txUser.UserPrincipalName ){
                $tAADu = get-AzureAdUser -objectID $msolu.UserPrincipalName |?{($_.provisionedplans.service -eq 'exchange')} ;
                if($taadu.extensionproperty.onPremisesDistinguishedName -match $rgxCMWDomain){
                    $bCmwAD=$true ;
                    write-host -fo yellow "ADUser is onprem CMW hybrid!:`n$($taadu.extensionproperty.onPremisesDistinguishedName)" ; 
                } elseif($taadu.DirSyncEnabled -AND $taadu.ImmutableId) {
                    #$tadu = get-aduser -filter {UserPrincipalName -eq $txUser.UserPrincipalName }
                    # no use the converted immutableid
                    $guid=New-Object -TypeName guid (,[System.Convert]::FromBase64String($taadu.ImmutableId)) ;
                    $tadu = get-aduser -Identity $guid.guid ; 
                };
            } else { 
                write-warning "No matching licensed MSolu:(get-msoluser -user $txUser.UserPrincipalName)" ; 
            } ; 
        } CATCH {
            $ErrTrapd=$Error[0] ;
            $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
            write-warning $smsg ;
        } ; 
        switch ($txUser.Recipienttype){
            'UserMailbox'{
                $error.clear() ;
                TRY {$xmbx = ps1GetxMbx -id $txUser.UserPrincipalName -ea stop }
                CATCH {
                    $ErrTrapd=$Error[0] ;
                    $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                    write-warning $smsg ;
                } ; 
            }
            'MailUser'{
                $error.clear() ;
                TRY {$opmbx = get-mailbox -id $txUser.UserPrincipalName -ea stop }
                CATCH {
                    $ErrTrapd=$Error[0] ;
                    $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                    write-warning $smsg ;
                } ; 
            } ;
            default {write-warning "non-mailbox/mailuser object"} 
        } ; 
        if($txUser){
            if($tadu){"=get-ADUser:>`n$(($tadu |fl samaccountn*,userpr*,msRTCSIP-PrimaryU*,msRTCSIP-L*,msRTCSIP-Usere*,tit*|out-string).trim())" 
            } else {
                write-host "=get-ADUser:>(Non-local AD user)`n=$((get-alias ps1GetxUser).definition):`n$(($txUser|fl userpr*,tit*,Offi*,Compa*|out-string).trim())" 
            } ;
            if($xmbx){"=get-Xmbx>:`n$(($xmbx| fl ($exMProps |?{$_ -notmatch '(samaccountname|DistinguishedName)'})|out-string).trim())" } ;
            if($opmbx){"=get-OPmbx>:`n$(($opmbx| fl $exMProps |out-string).trim())" };
            if($msolu){
                write-host "$(($msolu|fl @{Name='HasLic';Expression={$_.IsLicensed }},@{Name='LicIssue';Expression={$_.LicenseReconciliationNeeded }}|out-string).trim())" ; 
            "Licenses Assigned:`n$((($msolu.licenses.AccountSkuId) -join ";" | out-string).trim())" ;
                if(!($bCmwAD)){
                    if($LicGrp = $tadu.memberof -match $rgxLicGrpDN){
                        write-host "LicGrp(AD):$(($LicGrp|out-string).trim())" ; 
                    } else { 
                        write-host "LicGrp(AD):(no ADUser.memberof matched pattern:`n$($rgxLicGrpDN.tostring())" ; 
                    } ; 
                } else {
                    write-host -fo yellow  "Unable to expand ADU, user is hybrid AD from CMW.internal domain`nproxying AzureADUser memberof" ; 
                    if($taadu){
                        $mbrof = $taadu | Get-AzureADUserMembership | select DisplayName,DirSyncEnabled,MailEnabled,SecurityEnabled,Mail,objectid ;
                        if($LicGrp = $mbrof.displayname -match $rgxLicGrpDName){
                            write-host "LicGrp(AAD):$(($LicGrp|out-string).trim())" ; 
                        } else { 
                            write-host "LicGrp(AAD):(no ADUser.memberof matched pattern:`n$($rgxLicGrpDName.tostring())" ; 
                        } ; 
                    } else { 
                        write-warning "(unpopulated AzureADUser: skipping memberof)" ; 
                    }
                } ; 
            }else {
                write-warning "Unable to find matching MsolU for $Identifier" ; 
            } ; 
        } ; 
        
    } ;  # PROC-E
    END {
        # =========== wrap up Tenant connections
        <# suppress VerbosePreference:Continue, if set, during mod loads (VERY NOISEY)
        if($VerbosePreference = "Continue"){
            $VerbosePrefPrior = $VerbosePreference ;
            $VerbosePreference = "SilentlyContinue" ;
            $verbose = ($VerbosePreference -eq "Continue") ;
        } ;
        if($script:useEXOv2){
            disconnect-exo2 -verbose:$($verbose) ;
        } else {
            disconnect-exo -verbose:$($verbose) ;
        } ;
        # aad mod *does* support disconnect (msol doesen't!)
        #Disconnect-AzureAD -verbose:$($verbose) ;
        # reenable VerbosePreference:Continue, if set, during mod loads
        if($VerbosePrefPrior -eq "Continue"){
            $VerbosePreference = $VerbosePrefPrior ;
            $verbose = ($VerbosePreference -eq "Continue") ;
        } ;
        #>
        # clear the script aliases
        write-verbose "clearing ps1* aliases in Script scope" ; 
        get-alias -scope Script |Where-Object{$_.name -match '^ps1.*'} | ForEach-Object{Remove-Alias -alias $_.name} ;

        write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($sBnr.replace('=v','=^').replace('v=','^='))" ;
        
        write-verbose "(explicit EXIT...)" ;
        Break ;


    } ;  # END-E
}

#*------^ resolve-Name.ps1 ^------

#*------v resolve-user.ps1 v------
function resolve-user {
    <#
    .SYNOPSIS
    resolve-user.ps1 - Resolve specified array of -users (displayname, emailaddress, samaccountname) to mail asset, lic & ticket descriptors
    .NOTES
    Version     : 1.0.0.
    Author      : Todd Kadrie
    Website     : http://www.toddomation.com
    Twitter     : @tostka / http://twitter.com/tostka
    CreatedDate : 2021-07-13
    FileName    : resolve-user.ps1
    License     : MIT License
    Copyright   : (c) 2021 Todd Kadrie
    Github      : https://github.com/tostka/verb-XXX
    Tags        : Powershell
    AddedCredit : REFERENCE
    AddedWebsite: URL
    AddedTwitter: URL
    REVISIONS
    * 12:55 PM 7/19/2021 added guest & exo-mailcontact support (resolving missing ext-federated addresses), retolled logic down to grcp & gxrcp to drive balance of tests.
    * 12:05 PM 7/14/2021 rem'd requires: verb-exo  rem'd requires version 5 (gen'ing 'version' is specified more than once.); rem'd the $rgxSamAcctName, gen's parsing errors compiling into mod ;  added alias 'ulu'; added mailcontact excl on init grcp, to force those to exombx qry ; init vers
    .DESCRIPTION
    .PARAMETER  users
    Array of user descriptors: displayname, emailaddress, UPN, samaccountname (checks clipboard where unspecified)
    .PARAMETER useEXOv2
    Use EXOv2 (ExchangeOnlineManagement) over basic auth legacy connection [-useEXOv2]
    .INPUTS
    None. Does not accepted piped input.(.NET types, can add description)
    .OUTPUTS
    Returns report to pipeline
    .EXAMPLE
    PS> .\resolve-user.ps1
    Default, attempts to parse a user descriptor from clipboard
    .EXAMPLE
    PS> .\resolve-user.ps1 -users 'John Public' 
    Process user displayname
    .EXAMPLE
    PS> .\resolve-user.ps1 -users 'Todd.Kadrie@toro.com','Stacy Sotelo','lynctest1','lforsythe@charlesmachine.works','confroom-b3-trenchingrm@toroco.onmicrosoft.com' -verbose ;
    Process an array of descriptors
    .LINK
    https://github.com/tostka/verb-exo
    .LINK
    #>
    ###Requires -Version 5
    #Requires -Modules ActiveDirectory, MSOnline, AzureAD, ExchangeOnlineManagement, verb-AAD, verb-ADMS, verb-Ex2010
    #Requires -RunasAdministrator
    # VALIDATORS: [ValidateNotNull()][ValidateNotNullOrEmpty()][ValidateLength(24,25)][ValidateLength(5)][ValidatePattern("(lyn|bcc|spb|adl)ms6(4|5)(0|1).(china|global)\.ad\.toro\.com")][ValidateSet("USEA","GBMK","AUSYD")][ValidateScript({Test-Path $_ -PathType 'Container'})][ValidateScript({Test-Path $_})][ValidateRange(21,65)][ValidateCount(1,3)]
    ## [OutputType('bool')] # optional specified output type
    [CmdletBinding()]
    [Alias('ulu')]
    PARAM(
        [Parameter(Position=0,Mandatory=$False,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,HelpMessage="Array of user descriptors: displayname, emailaddress, UPN, samaccountname (checks clipboard where unspecified)")]
        #[ValidateNotNullOrEmpty()]
        #[Alias('ALIAS1', 'ALIAS2')]
        [array]$users,
        [Parameter(HelpMessage="Use EXOv2 (ExchangeOnlineManagement) over basic auth legacy connection [-useEXOv2]")]
        [switch] $useEXOv2
    ) ;
    BEGIN{
        $rgxEmailAddr = "^([0-9a-zA-Z]+[-._+&'])*[0-9a-zA-Z]+@([-0-9a-zA-Z]+[.])+[a-zA-Z]{2,63}$" ; 
        $rgxDName = "^([a-zA-Z]{2,}\s[a-zA-Z]{1,}'?-?[a-zA-Z]{2,}\s?([a-zA-Z]{1,})?)" ; 
        $rgxSamAcctNameTOR = "^\w{2,20}$" ; # up to 20k, the limit prior to win2k
        #$rgxSamAcctName = "^[^\/\\\[\]:;|=,+?<>@�]+$" # no char limit ;

        if(!$users){
            $users= (get-clipboard).trim().replace("'",'').replace('"','') ; 
            if($users){
                write-verbose "No -users specified, detected value on clipboard:`n$($users)" ; 
            } else { 
                write-warning "No -users specified, nothing matching dname, emailaddress or samaccountname, found on clipboard. EXITING!" ; 
                Break ; 
            } ; 
        } else {
            write-verbose "($(($users|measure).count)) user(s) specified:`n'$($users -join "','")'" ;         
        } ; 

        rx10 -Verbose:$false ; rxo  -Verbose:$false ; cmsol  -Verbose:$false ;

    } 
    PROCESS{
        #$dname= 'Todd Kadrie' ; 
        #$dname = 'Stacy Sotelo'
        $ttl = ($users|measure).count ; $Procd=0 ;
        [array]$Rpt =@() ; 
        foreach ($usr in $users){
            $fname = $lname = $dname = $tRcp = $tMbx = $txRMbx = $tADU = $txMbx = $txU = $xMmbrOf = $mu = $licgrp = $null ; 
            $isEml=$isDname=$isSamAcct=$false ; 
            $procd++ ; 
        
            switch -regex ($usr){
                $rgxEmailAddr {
                    $fname,$lname = $usr.split('@')[0].split('.') ; 
                    $dname = $usr ;
                    write-verbose "(detected user ($($usr)) as EmailAddr)" ; 
                    $isEml = $true ;
                } 
                $rgxDName {
                    $fname,$lname = $usr.split(' ') ;
                    $dname = $usr ; 
                    write-verbose "(detected user ($($usr)) as DisplayName)" ; 
                    $isDname = $true ;
                } 
                $rgxSamAcctNameTOR {
                    $lname = $usr ; 
                    write-verbose "(detected user ($($usr)) as SamAccountName)" ; 
                    $isSamAcct  = $true ;
                } 
                default {
                    write-warning "$((get-date).ToString('HH:mm:ss')):No -user specified, nothing matching dname, emailaddress or samaccountname, found on clipboard. EXITING!" ; 
                    #Break ; 
                } ; 
            } ; 

            $sBnr="===v ($($Procd)/$($ttl)):Input: '$($usr)' | '$($fname)' | '$($lname)' v===" ;
            if($isEml){$sBnr+="(EML)"}
            elseif($isDname){$sBnr+="(DNAM)"}
            elseif($isSamAcct){$sBnr+="(SAM)"}
            write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($sBnr)" ;
        
            $xMProps="samaccountname","windowsemailaddress","DistinguishedName","Office","RecipientTypeDetails" ;
            $lProps = @{Name='HasLic'; Expression={$_.IsLicensed }},@{Name='LicIssue'; Expression={$_.LicenseReconciliationNeeded }} ;
            $rgxOPLic = '^CN\=ENT\-APP\-Office365\-(EXOK|F1|MF1)-DL$' ; 
            $rgxXLic = '^CN\=ENT\-APP\-Office365\-(EXOK|F1|MF1)-DL$' ; 
            write-host -foreground yellow "get-Rmbx/xMbx: " -nonewline;

            # $isEml=$isDname=$isSamAcct=$false ; 
            $pltgM=[ordered]@{} ; 
            if($isEml -OR $isSamAcct){
                $pltgM.add('identity',$usr) ;
            } ; 
            if($isDname){
                $fltr = "displayname -like '$dname'" ; 
                $pltgM.add('filter',$fltr) ;
            } ; 

            $error.clear() ;
            
            $tRcp=get-recipient @pltgM -ea 0 |?{$_.recipienttypedetails -ne 'MailContact'}
            #$txMbx=get-exomailbox @pltgM -ea 0 ;
            $txRcp=get-exorecipient @pltgM -ea 0 ;

            write-verbose "get-[exo]Recipient w`n$(($pltgM|out-string).trim())" ; 
            #write-verbose "get-recipient w`n$(($pltgM|out-string).trim())" ; 
            # exclude contacts, they don't represent real onprem mbx assoc, and we need to refer those to EXO mbx qry anyway.
            if($tRcp){
                $error.clear() ;
                TRY {
                    switch -regex ($tRcp.recipienttype){
                        "UserMailbox" {
                            $tMbx=get-mailbox $tRcp.identity ;
                            $Rpt += $tMbx.primarysmtpaddress ; 
                        } 
                        "MailUser" {
                            $txRMbx=get-remotemailbox $tRcp.identity  ;
                            $Rpt += $txRMbx.primarysmtpaddress ;  
                        } ;
                        default {
                            write-warning "$((get-date).ToString('HH:mm:ss')):Unsupported RecipientType:($tRcp.recipienttype). EXITING!" ; 
                            Break ; 
                        }
                    }
                    $pltGadu=[ordered]@{Identity = $null ; Properties='*' ;errorAction='SilentlyContinue'} ;
                    if($txRMbx ){
                        $pltGadu.identity = $txRmbx.samaccountname;
                    }elseif($tMbx){
                        $pltGadu.identity = $tmbx.samaccountname ;
                    } ; 
                    if($pltGadu.identity){
                        $tADU =Get-ADUser @pltGadu ;
                        $smsg = "(TOR USER, fed:ad.toro.com)" ; 
                        write-host -Fore yellow $smsg ; 
                        if($txRMbx){$smsg = "$(($txRMbx |fl $xMProps|out-string).trim())`n-Title:$($tADU.Title)" } ; 
                        if($tMbx){$smsg =  "$(($tMbx |fl $xMProps|out-string).trim())`n-Title:$($tADU.Title)" } ; 
                        write-host $smsg ;
                    } ; 
                } CATCH {
                    $ErrTrapd=$Error[0] ;
                    $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    Continue ; 
                } ; 
            #}elseif($txMbx){
            }elseif($txRcp){
                foreach($txR in $txRcp){
                    TRY {
                        switch -regex ($txR.recipienttypedetails){
                            "UserMailbox" {
                                $txMbx=get-exomailbox @pltgM -ea 0 ;
                                $Rpt += $txMbx.primarysmtpaddress ; 
                                break ; 
                            } 
                            "GuestMailUser" {
                                #$txRMbx=get-remotemailbox $txR.identity  ;
                                caad -verbose:$false ; 
                                $txU = $txR | get-exouser ;
                                $txGuest = get-AzureAdUser  -objectid $txU.userPrincipalName ;
                                #$Rpt += $txRMbx.primarysmtpaddress ;  
                                write-host "$($txR.ExternalEmailAddress): matches a Guest object with UPN:$($txU.userPrincipalName)" ; 
                                if($txGuest.EmailAddresses -eq $null){
                                    write-warning "Guest appears to have damage from conficting replicated onprem MailContact, as it's EmailAddresses property is *blank*" ; 
                                } ; 
                                break ; 
                            } ;
                            "MailContact" {
                                #$txRMbx=get-remotemailbox $txR.identity  ;
                                #$Rpt += $txRMbx.primarysmtpaddress ;  
                                write-host "$($txR.primarysmtpaddress): matches an EXO MailContact with external Email: $($txR.primarysmtpaddress)" ; 
                                break ; 
                            } ;
                            "MailUniversalSecurityGroup" {
                                #$txRMbx=get-remotemailbox $txR.identity  ;
                                #$Rpt += $txRMbx.primarysmtpaddress ;  
                                write-host "$($txR.primarysmtpaddress): matches an EXO MailUniversalSecurityGroup with Dname: $($txR.displayname)" ; 
                                break ; 
                            } ;
                            default {
                                write-warning "$((get-date).ToString('HH:mm:ss')):Unsupported RecipientType:($tRcp.recipienttype). EXITING!" ; 
                                Break ; 
                            }
                        }
                    } CATCH {
                        $ErrTrapd=$Error[0] ;
                        $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                        else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        Continue ; 
                    } ; 
                }  # loop-E $txR
                # contacts and guests won't drop with $txRmbx or $tmbx populated
                TRY {
                    $pltGadu=[ordered]@{Identity = $null ; Properties='*' ;errorAction='SilentlyContinue'} ;
                    if($txRMbx ){
                        $pltGadu.identity = $txRmbx.samaccountname;
                    }elseif($tMbx){
                        $pltGadu.identity = $tmbx.samaccountname ;
                    } ; 
                    if($pltGadu.identity){
                        $tADU =Get-ADUser @pltGadu ;
                        $smsg = "(TOR USER, fed:ad.toro.com)" ; 
                        write-host -Fore yellow $smsg ; 
                        if($txRMbx){$smsg = "$(($txRMbx |fl $xMProps|out-string).trim())`n-Title:$($tADU.Title)" } ; 
                        if($tMbx){$smsg =  "$(($tMbx |fl $xMProps|out-string).trim())`n-Title:$($tADU.Title)" } ; 
                        write-host $smsg ;
                    } ; 
                } CATCH {
                    $ErrTrapd=$Error[0] ;
                    $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    Continue ; 
                } ; 

                $Rpt += $txMbx.primarysmtpaddress ; 
                if($txMbx.isdirsynced){$smsg="(CMW USER, fed:cmw.internal)"}
                else{$smsg="(CLOUD-1ST ACCT, unfederated)"} ;
                write-host -Fore yellow $smsg ; 
                # skip user lookup if guest already pulled it 
                if(!$txU){
                    $txU = get-exouser -id $txmbx.UserPrincipalName ; 
                } 
                if($txMbx){
                write-host "=get-xMbx:>`n$(($txMbx |fl ($xMprops |?{$_ -notmatch '(sam.*|dist.*)'})|out-string).trim())`n-Title:$($txU.Title)";
                }elseif($txGuest){
                    write-host "=get-AADU:>`n$(($txGuest |fl userp*,PhysicalDeliveryOfficeName,JobTitle|out-string).trim())"
                } ; 
                TRY {
                    $xMmbrOf = Get-exoRecipient -Filter "Members -eq '$($txU.DistinguishedName)'" -RecipientTypeDetails GroupMailbox,MailUniversalDistributionGroup,MailUniversalSecurityGroup ;
                } CATCH {
                    $ErrTrapd=$Error[0] ;
                    $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    Continue ; 
                } ; 
            } else { write-warning "(no matching EXOP or EXO recipient object:$($usr))"   } ; # don't break, doesn't continue loop

            $pltgMU=@{UserPrincipalName=$null} ; 
            if($tADU){$pltgMU.UserPrincipalName = $tADU.UserPrincipalName } 
            elseif($txMbx){$pltgMU.UserPrincipalName = $txMbx.UserPrincipalName }
            elseif($txGuest){$pltgMU.UserPrincipalName = $txGuest.userprincipalname } 
            else{} ; 
            
            if($pltgMU.UserPrincipalName){
                write-host -foregroundcolor yellow "=get-msoluser $($pltgMU.UserPrincipalName):(licences)>:" ;
                TRY{
                    $mu=get-msoluser @pltgMU ;
                } CATCH {
                    $ErrTrapd=$Error[0] ;
                    $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    Continue ; 
                } ; 
                $smsg = "$(($mu|fl $lProps|out-string).trim())`n" ;
                $smsg += "Licenses Assigned:$(($mu.licenses.AccountSkuId -join '; '|out-string).trim())" ; 
                write-host $smsg ; 
                if($tadu){$licgrp = $tadu.memberof |?{$_ -match $rgxOPLic }}
                elseif($xMmbrOf){$licgrp = $xMmbrOf.Name |?{$_ -match $rgxXLic}}
                if(!($licgrp) -AND ($mu.licenses.AccountSkuId -contains 'toroco:ENTERPRISEPACK')){$licgrp = '(direct-assigned E3)'} ; 
                if($licgrp){$smsg = "LicGrp:$($licgrp)"}
                else{$smsg = "LicGrp:(unresolved, direct-assigned other?)" } ; 
                write-host $smsg ; 
            } ; 
            write-host -foregroundcolor green $sBnr.replace('=v','=^').replace('v=','^=') ; 
        } ; 
    }
    END{
        $Rpt -join ',' | out-clipboard ; 
        write-host "(output copied to clipboard)"
        $Rpt -join ',' | write-output ;
     }
 }

#*------^ resolve-user.ps1 ^------

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
    Reconnect-EXO2 -cred $credO365TOLSID -Verbose:($VerbosePreference -eq 'Continue')
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

#*------v test-ExoPSession.ps1 v------
Function test-ExoPSession {
  <#
    .SYNOPSIS
    test-ExoPSession - Does a *simple* - NO-ORG REVIEW - validation of functional PSSession with: ConfigurationName:'Microsoft.Exchange' -AND Name match  '^(Exchange2010|Session\sfor\simplicit\sremoting\smodule\sat\s.*)' -AND State:'Opened' -AND Availability:'Available' -AND can gcm -name 'Add-ADPermission'
    .NOTES
    Author: Todd Kadrie
    Website:	http://toddomation.com
    Twitter     :	@tostka / http://twitter.com/tostka
    Version     : 1.0.0
    CreatedDate : 2021-04-15
    FileName    : test-ExoPSession()
    License     : MIT License
    Copyright   : (c) 2020 Todd Kadrie
    Github      : https://github.com/tostka
    Tags        : Powershell,Exchange,Exchange-2013,Exchange-2016
    REVISIONS   :
    * 11:40 AM 5/14/2021 added -ea 0 to the gv tests (suppresses not-found error when called without logging config)
    * 10:38 AM 5/3/2021 init vers
    .DESCRIPTION
    test-ExoPSession - Does a *simple* - NO-ORG REVIEW - validation of functional EXO PSSession with: ConfigurationName:'Microsoft.Exchange' -AND Name match (ExchangeOnlineInternalSession| "^(Session|WinRM)\d*) -AND State:'Opened' -AND Availability:'Available' -AND can gcm -name 'Add-*ATPEvaluation'.
    This does *NO* validation that any specific EXOnPrem org is attached! It just validates that an existing PSSession *exists* that *generically* matches a Remote Exchange Mgmt Shell connection in a usable state. Use case is scripts/functions that *assume* you've already pre-established a suitable connection, and just need to pre-test that *any* PSS is already open, before attempting commands. 
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    System.Management.Automation.Runspaces.PSSession. Returns the functional PSSession object(s)
    .EXAMPLE
    PS> if(test-ExoPSession){'OK'} else { 'NOGO!'}  ;
    .LINK
    https://github.com/tostka/verb-Exo/
    #>
    [CmdletBinding()]
    Param()  ;
    BEGIN{
        $verbose = ($VerbosePreference -eq "Continue") ;
        if(!$rgxExoPsHostName){$rgxExoPsHostName="^(ps\.outlook\.com|outlook\.office365\.com)$" } ;
        $testCommand = 'Add-*ATPEvaluation' ; 
        $propsREMS = 'Id','Name','ComputerName','ComputerType','State','ConfigurationName','Availability' ; 
    } ;  # BEG-E
    PROCESS{
        $error.clear() ;
        TRY {
            $exov2Good = Get-PSSession | where-object {($_.ConfigurationName -like "Microsoft.Exchange") -AND (
            $_.Name -like "ExchangeOnlineInternalSession*") -AND ($_.ComputerName -match $rgxExoPsHostName) -AND ($_.State -like "*Opened*") -AND (
            $_.Availability -eq 'Available')} ; 
            $exov1Good = (Get-PSSession | where-object {$_.ConfigurationName -like "Microsoft.Exchange" -AND $_.Name -match "^(Session|WinRM)\d*" -AND ($_.ComputerName -match $rgxExoPsHostName) -AND (
                ($_.State -eq 'Opened') -AND ($_.Availability -eq 'Available')) }) ;
            if( $exov2Good -OR $exov1Good ){
                $REMSexo=@() ; 
                $REMSexo = $exov2Good ; 
                $REMSexo += $exov1Good ; 
                $smsg = "valid EXO EMS PSSession found:`n$(($REMSexo|ft -a $propsREMS |out-string).trim())" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-VERBOSE "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                # test agnostic of prefix variant
                if($tmod = (get-command $testCommand ).source){
                    $smsg = "(confirmed PSSession open/available, with $($testCommand) available)" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    $REMSexo | write-output ; ;
                } else { 
                    throw "NO FUNCTIONAL PSSESSION FOUND!" ; 
                } ; 
            } else {
                throw "No existing open/available EXO Remote Exchange Management Shell found!"
            } ;
        } CATCH {
            $ErrTrapd = $_ ;
            write-warning "$(get-date -format 'HH:mm:ss'): Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
            #-=-record a STATUSERROR=-=-=-=-=-=-=
            $statusdelta = ";ERROR"; # CHANGE|INCOMPLETE|ERROR|WARN|FAIL ;
            if(gv passstatus -scope Script -ea 0 ){$script:PassStatus += $statusdelta } ;
            if(gv -Name PassStatus_$($tenorg) -scope Script  -ea 0 ){set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta)} ;
            #-=-=-=-=-=-=-=-=
        } ;
        
    } ;  # PROC-E
    END {}
}

#*------^ test-ExoPSession.ps1 ^------

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
    * 11:40 AM 5/14/2021 added -ea 0 to the gv tests (suppresses not-found error when called without logging config)
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
                $ErrTrapd=$Error[0] ;
                $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                #-=-record a STATUSWARN=-=-=-=-=-=-=
                $statusdelta = ";WARN"; # CHANGE|INCOMPLETE|ERROR|WARN|FAIL ;
                if(gv passstatus -scope Script -ea 0){$script:PassStatus += $statusdelta } ;
                if(gv -Name PassStatus_$($tenorg) -scope Script -ea 0){set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta)} ; 
                #-=-=-=-=-=-=-=-=
                $smsg = "FULL ERROR TRAPPED (EXPLICIT CATCH BLOCK WOULD LOOK LIKE): } catch[$($ErrTrapd.Exception.GetType().FullName)]{" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
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
                    $ErrTrapd=$Error[0] ;
                    $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    #-=-record a STATUSWARN=-=-=-=-=-=-=
                    $statusdelta = ";WARN"; # CHANGE|INCOMPLETE|ERROR|WARN|FAIL ;
                    if(gv passstatus -scope Script -ea 0){$script:PassStatus += $statusdelta } ;
                    if(gv -Name PassStatus_$($tenorg) -scope Script -ea 0){set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta)} ; 
                    #-=-=-=-=-=-=-=-=
                    $smsg = "FULL ERROR TRAPPED (EXPLICIT CATCH BLOCK WOULD LOOK LIKE): } catch[$($ErrTrapd.Exception.GetType().FullName)]{" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level ERROR } #Error|Warn|Debug 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
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

#*------v test-xoMailbox.ps1 v------
Function test-xoMailbox {
    <#
    .SYNOPSIS
    test-xoMailbox.ps1 - Run quick mailbox function validation and metrics gathering, to streamline, 'Is it working?' questions. 
    .NOTES
    Version     : 10.0
    Author      : Todd Kadrie
    Website     : http://www.toddomation.com
    Twitter     : @tostka / http://twitter.com/tostka
    CreatedDate : 2021-04-21
    FileName    : test-xoMailbox.ps1
    License     : MIT License
    Copyright   : (c) 2020 Todd Kadrie
    Github      : https://github.com/tostka/verb-XXX
    Tags        : Powershell,ExchangeOnline,Exchange,Resource,MessageTrace
    REVISIONS
    * 1:51 PM 5/19/2021 expanded $pltGHOpCred= to include 'ESVC','SID'; verbose=$($verbose)} ;
    # 12:27 PM 5/11/2021 updated to cover Room|Shared|Equipment mbx types, along with UserMailbox
    # 3:47 PM 5/6/2021 recoded start-log switching, was writing logs into AllUsers profile dir ; swapped out 'Fail' msgtrace expansion, and have it do all non-Delivery statuses, up to the $MsgTraceNonDeliverDetailsLimit = 10; tested, appears functional
    # 3:15 PM 5/4/2021 added trailing |?{$_.length} bug workaround for get-gcfastxo.ps1
    * 4:20 PM 4/29/2021 debugged, looks functional - could benefit from moving the msgtrk summary down into the output block, but [shrug]
    * 7:56 AM 4/28/2021 init
    .DESCRIPTION
    test-xoMailbox.ps1 - Run quick mailbox function validation and metrics gathering, to streamline, 'Is it working?' questions. 
    .PARAMETER Mailboxes
    Array of Mailbox email addresses [-Mailboxes mbx1@domain.com]
    .PARAMETER Ticket
    Ticket # [-Ticket nnnnn]
    .PARAMETER useEXOv2
    Use EXOv2 (ExchangeOnlineManagement) over basic auth legacy connection [-useEXOv2]
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    .EXAMPLE
    test-xoMailbox.ps1 -TenOrg TOR -Mailboxes 'Fname.LName@domain.com','FName2.Lname2@domain.com' -Ticket 610706 -verbose ;
    .EXAMPLE
    .\test-xoMailbox.ps1
    .LINK
    https://github.com/tostka/verb-EXO
    #>
    #Requires -Modules ActiveDirectory,verb-Auth,verb-IO,verb-Mods,verb-Text,verb-Network,verb-AAD,verb-ADMS,verb-Ex2010,verb-logging
    [CmdletBinding()]
    PARAM(
        [Parameter(Mandatory=$FALSE,HelpMessage="TenantTag value, indicating Tenants to connect to[-TenOrg 'TOL']")]
        [ValidateNotNullOrEmpty()]
        $TenOrg = 'TOR',
        [Parameter(Position=0,Mandatory=$True,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,HelpMessage="Array of EXO Mailbox identifiers to be processed [-mailboxes 'xxx','yyy']")]
        [ValidateNotNullOrEmpty()]$Mailboxes,
        [Parameter(Mandatory = $True, HelpMessage = "Ticket # [-Ticket nnnnn]")]
        [array]$Tickets,
        [Parameter(Position=0,Mandatory=$False,HelpMessage="Specific ResourceDelegate address to be confirmed for detailed delivery [emailaddr]")]
        [ValidateNotNullOrEmpty()][string]$TargetDelegate,
        [Parameter(Mandatory=$false,HelpMessage="Days back to search (defaults to 10)[-Days 10]")]
        [int]$Days=10,
        [Parameter(HelpMessage="Use EXOv2 (ExchangeOnlineManagement) over basic auth legacy connection [-useEXOv2]")]
        [switch] $useEXOv2
    ) ;
    BEGIN {
        $Verbose = ($VerbosePreference -eq 'Continue') ;
        #*======v SCRIPT/DOMAIN/MACHINE/INITIALIZATION-DECLARE-BOILERPLATE v======
        # SCRIPT-CONFIG MATERIAL TO SET THE UNDERLYING $DBGPREF:
        if ($ShowDebug) { $DebugPreference = "Continue" ; write-debug "(`$showDebug:$showDebug ;`$DebugPreference:$DebugPreference)" ; };
        if ($Whatif) { write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):`$Whatif is TRUE (`$whatif:$($whatif))" ; };
        # If using WMI calls, push any cred into WMI:
        #if ($Credential -ne $Null) {$WmiParameters.Credential = $Credential }  ;

        <# scriptname with extension
        $ScriptDir = (Split-Path -parent $MyInvocation.MyCommand.Definition) + "\" ;
        $ScriptBaseName = (Split-Path -Leaf ((& { $myInvocation }).ScriptName))  ;
        $ScriptNameNoExt = [system.io.path]::GetFilenameWithoutExtension($MyInvocation.InvocationName) ;
        #>
        # 11:24 AM 6/21/2019 UPDATED HYBRID VERS
        if($showDebug){
            write-host -foregroundcolor green "`SHOWDEBUG: `$PSScriptRoot:$($PSScriptRoot)`n`$PSCommandPath:$($PSCommandPath)" ;
        } ;
        if ($psISE){
                $ScriptDir = Split-Path -Path $psISE.CurrentFile.FullPath ;
                $ScriptBaseName = split-path -leaf $psise.currentfile.fullpath ;
                $ScriptNameNoExt = [system.io.path]::GetFilenameWithoutExtension($psise.currentfile.fullpath) ;
                        $PSScriptRoot = $ScriptDir ;
                if($PSScriptRoot -ne $ScriptDir){ write-warning "UNABLE TO UPDATE BLANK `$PSScriptRoot TO CURRENT `$ScriptDir!"} ;
                $PSCommandPath = $psise.currentfile.fullpath ;
                if($PSCommandPath -ne $psise.currentfile.fullpath){ write-warning "UNABLE TO UPDATE BLANK `$PSCommandPath TO CURRENT `$psise.currentfile.fullpath!"} ;
        } else {
            if($host.version.major -lt 3){
                $PSScriptRoot = Split-Path $MyInvocation.MyCommand.Path -Parent ;
                $PSCommandPath = $myInvocation.ScriptName ;
                $ScriptBaseName = (Split-Path -Leaf ((&{$myInvocation}).ScriptName))  ;
                $ScriptNameNoExt = [system.io.path]::GetFilenameWithoutExtension($MyInvocation.InvocationName) ;
            } elseif($PSScriptRoot) {
                $ScriptDir = $PSScriptRoot ;
                if($PSCommandPath){
                    $ScriptBaseName = split-path -leaf $PSCommandPath ;
                    $ScriptNameNoExt = [system.io.path]::GetFilenameWithoutExtension($PSCommandPath) ;
                } else {
                    $PSCommandPath = $myInvocation.ScriptName ;
                    $ScriptBaseName = (Split-Path -Leaf ((&{$myInvocation}).ScriptName))  ;
                    $ScriptNameNoExt = [system.io.path]::GetFilenameWithoutExtension($MyInvocation.InvocationName) ;
                } ;
            } else {
                if($MyInvocation.MyCommand.Path) {
                    $PSScriptRoot = Split-Path $MyInvocation.MyCommand.Path -Parent ;
                    $ScriptBaseName = (Split-Path -Leaf ((&{$myInvocation}).ScriptName))  ;
                    $ScriptNameNoExt = [system.io.path]::GetFilenameWithoutExtension($MyInvocation.InvocationName) ;
                } else {
                    throw "UNABLE TO POPULATE SCRIPT PATH, EVEN `$MyInvocation IS BLANK!" ;
                } ;
            } ;
        } ;
        if($showDebug){
            write-host -foregroundcolor green "`SHOWDEBUG: `$ScriptDir:$($ScriptDir)`n`$ScriptBaseName:$($ScriptBaseName)`n`$ScriptNameNoExt:$($ScriptNameNoExt)`n`$PSScriptRoot:$($PSScriptRoot)`n`$PSCommandPath:$($PSCommandPath)" ;
        } ;
    
        $ComputerName = $env:COMPUTERNAME ;
        $sQot = [char]34 ; $sQotS = [char]39 ;
        $NoProf = [bool]([Environment]::GetCommandLineArgs() -like '-noprofile'); # if($NoProf){# do this};
        $MyBox = "LYN-3V6KSY1", "TIN-BYTEIII", "TIN-BOX", "TINSTOY", "LYN-8DCZ1G2" ;
        $DomainWork = "TORO";
        $DomHome = "REDBANK";
        $DomLab = "TORO-LAB";
        #$ProgInterval= 500 ; # write-progress wait interval in ms
        # 12:23 PM 2/20/2015 add gui vb prompt support
        #[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null ;
        # 11:00 AM 3/19/2015 should use Windows.Forms where possible, more stable

        switch -regex ($env:COMPUTERNAME) {
            ($rgxMyBox) { $LocalInclDir = "c:\usr\work\exch\scripts" ; }
            ($rgxProdEx2010Servers) { $LocalInclDir = "c:\scripts" ; }
            ($rgxLabEx2010Servers) { $LocalInclDir = "c:\scripts" ; }
            ($rgxProdL13Servers) { $LocalInclDir = "c:\scripts" ; }
            ($rgxLabL13Servers) { $LocalInclDir = "c:\scripts" ; }
            ($rgxAdminJumpBoxes) { $LocalInclDir = (split-path $profile) ; }
        } ;

        $Retries = 4 ;
        $RetrySleep = 5 ;
        $DawdleWait = 30 ; # wait time (secs) between dawdle checks
        $DirSyncInterval = 30 ; # AADConnect dirsync interval 

        # email trigger vari, it will be semi-delimd list of mail-triggering events
        $PassStatus = $null ;


        $bMaintainEquipementLists = $false ; 
        $MsgTraceNonDeliverDetailsLimit = 10 ; # number of non status:Delivered messages to check/dump
    
        $ModMsgWindowMins = 5 ;
        if(!$Retries){$Retries = 4 } ;
        if(!$RetrySleep){$RetrySleep = 5} ;

         #$rgxOnMSAddr='.*.mail\.onmicrosoft\.com' ;
        # cover either mail. or not
        $rgxOnMSAddr = '.*@\w*((\.mail)*)\.onmicrosoft.com'
        $rgxExoSysFolders = '.*\\(Recipient\sCache|PersonMetadata|Audits|Calendar\sLogging|Purges)$'
            #'.*\\(Recipient\sCache|PersonMetadata|Audits|Calendar\sLogging|Purges|GAL\sContacts|Yammer\sRoot|Recoverable\sItems|Deletions|Versions)'
        #$rgxExcl = '.*\\(Recipient\sCache|PersonMetadata|Audits|Calendar\sLogging|Purges)$' ;

        $rgxSID = "^S-\d-\d+-(\d+-){1,14}\d+$" ;
        $rgxEntLicGrps = "CN=ENT-APP-Office365-.*-DL,OU=ENTERPRISE,DC=global,DC=ad,DC=toro((lab)*),DC=com" ;

        $propsXmbx = 'UserPrincipalName','Alias','ExchangeGuid','Database','ExternalDirectoryObjectId','RemoteRecipientType'
        $propsOPmbx = 'UserPrincipalName','SamAccountName','RecipientType','RecipientTypeDetails' ; 
        $propsRmbx = 'UserPrincipalName','ExchangeGuid','RemoteRoutingAddress','RemoteRecipientType' ;
        #$propNames = "SamAccountName","UserPrincipalName","name","mailNickname","msExchHomeServerName","mail","msRTCSIP-UserEnabled","msRTCSIP-PrimaryUserAddress","msRTCSIP-Line","DistinguishedName","Description","info","Enabled","LastLogonDate","userAccountControl","manager","whenChanged","whenCreated","msExchRecipientDisplayType","msExchRecipientTypeDetails","City","Company","Country","countryCode","Office","Department","Division","EmailAddress","employeeType","State","StreetAddress","surname","givenname","telephoneNumber" ;
        #$propNames = "SamAccountName","UserPrincipalName","name","mailNickname","msExchHomeServerName","mail","msRTCSIP-UserEnabled","msRTCSIP-PrimaryUserAddress","msRTCSIP-Line","DistinguishedName","Description","info","Enabled","LastLogonDate","userAccountControl","manager","whenChanged","whenCreated","msExchRecipientDisplayType","msExchRecipientTypeDetails","City","Company","Country","countryCode","Office","Department","Division","EmailAddress","employeeType","State","StreetAddress","surname","givenname","telephoneNumber" ;
        #$adprops = "samaccountname", "msExchRemoteRecipientType", "msExchRecipientDisplayType", "msExchRecipientTypeDetails", "userprincipalname" ;
        $adprops = "samaccountname","UserPrincipalName","memberof","msExchMailboxGuid","msexchrecipientdisplaytype","msExchRecipientTypeDetails","msExchRemoteRecipientType"
        $propsmbxfldrs = @{Name='Folder'; Expression={$_.Identity.tostring()}},@{Name='Items'; Expression={$_.ItemsInFolder}}, @{n='SizeMB'; e={($_.FolderSize.tostring().split('(')[1].split(' ')[0].replace(',','')/1MB).ToString('0.000')}}, @{Name='OldestItem'; Expression={get-date $_.OldestItemReceivedDate -f 'yyyy/MM/dd'}},@{Name='NewestItem'; Expression={get-date $_.NewestItemReceivedDate -f 'yyyy/MM/dd'}} ;
        $propsMsgTrc = @{N='ReceivedLocal';E={$_.Received.ToLocalTime()}},'SenderAddress','RecipientAddress','Subject','Status','ToIP','FromIP','Size','MessageId','MessageTraceId','Index' ; 
        $propsMsgTrcDtl = 'Date','Event','Action','Detail','Data' ;  
        $msgprops = 'Received','SenderAddress','RecipientAddress','Subject','Status' # ,'MessageId','MessageTraceId' ;
        $mtdprops = 'Date','Event','Detail','data' ;
    
        #*======v HELPER FUNCTIONS v======

        #-------v Function _cleanup v-------
        function _cleanup {
            # clear all objects and exit
            # Clear-item doesn't seem to work as a variable release

            # 12:58 PM 7/23/2019 spliced in chunks from same in maintain-restrictedothracctsmbxs.ps1
            # 2:02 PM 9/21/2018 missing $timestampnow, hardcode
            # 8:45 AM 10/13/2015 reset $DebugPreference to default SilentlyContinue, if on
            # # 8:46 AM 3/11/2015 at some time from then to 1:06 PM 3/26/2015 added ISE Transcript
            # 8:39 AM 12/10/2014 shifted to stop-transcriptLog function
            # 7:43 AM 1/24/2014 always stop the running transcript before exiting

            write-verbose "_cleanup" ; 
            <# transcript/log are handled in the mbx loop
            #stop-transcript
            # 11:16 AM 1/14/2015 aha! does this return a value!??
            if ($host.Name -eq "Windows PowerShell ISE Host") {
                # 8:46 AM 3/11/2015 shift the logfilename gen out here, so that we can arch it
                #$Logname= (join-path -path (join-path -path $scriptDir -childpath "logs") -childpath ($scriptNameNoExt + "-" + (get-date -uformat "%Y%m%d-%H%M" ) + "-ISEtrans.log")) ;
                # 2:16 PM 4/27/2015 shift to static timestamp $timeStampNow
                #$Logname= (join-path -path (join-path -path $scriptDir -childpath "logs") -childpath ($scriptNameNoExt + "-" + $timeStampNow + "-ISEtrans.log")) ;
                # 2:02 PM 9/21/2018 missing $timestampnow, hardcode
                $Logname = (join-path -path (join-path -path $scriptDir -childpath "logs") -childpath ($scriptNameNoExt + "-" + (get-date -format 'yyyyMMdd-HHmmtt') + "-ISEtrans.log")) ;
                write-host "`$Logname: $Logname";
                Start-iseTranscript -logname $Logname ;
                #Archive-Log $Logname ;
                # 1:23 PM 4/23/2015 standardize processing file so that we can send a link to open the transcript for review
                $transcript = $Logname
            } else {
                write-verbose "$(get-timestamp):Stop Transcript" ;
                Stop-TranscriptLog ;
                #write-verbose "$(get-timestamp):Archive Transcript" ;
                #Archive-Log $transcript ;
            } # if-E
            
            # also echo the log:
            if ($logging) { 
                # Write-Log -LogContent $smsg -Path $logfile
                $smsg = "`$logging:`$true:written to:`n$($logfile)" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } ; 
            #>

            # need to email transcript before archiving it
            
            #$smtpSubj= "Proc Rpt:$($ScriptBaseName):$(get-date -format 'yyyyMMdd-HHmmtt')"   ;

            #Load as an attachment into the body text:
            #$body = (Get-Content "path-to-file\file.html" ) | converto-html ;
            #$SmtpBody += ("Pass Completed "+ [System.DateTime]::Now + "`nResults Attached: " +$transcript) ;
            # 4:07 PM 10/11/2018 giant transcript, no send
            #$SmtpBody += "Pass Completed $([System.DateTime]::Now)`nResults Attached:($transcript)" ;
            $SmtpBody += "Pass Completed $([System.DateTime]::Now)`nTranscript:($transcript)" ;
            # 12:55 PM 2/13/2019 append the $PassStatus in for reference
            if($PassStatus ){
                $SmtpBody += "`n`$PassStatus triggers:: $($PassStatus)`n" ;
            } ;
            $SmtpBody += "`n$('-'*50)" ;
            #$SmtpBody += (gc $outtransfile | ConvertTo-Html) ;
            # name $attachment for the actual $SmtpAttachment expected by Send-EmailNotif
            #$SmtpAttachment=$transcript ;
            # 1:33 PM 4/28/2017 test for ERROR|CHANGE - actually non-blank, only gets appended to with one or the other
            if($PassStatus ){
                Send-EmailNotif ;
            } else {
                write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):No Email Report: `$Passstatus is `$null ; " ;
            } ;

            
            #$smsg= "#*======^ END PASS:$($ScriptBaseName) ^======" ;
            #this is now a function, use: ${CmdletName}
            $smsg= "#*======^ END PASS:$(${CmdletName}) ^======" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn

            Break
        } #*------^ END Function _cleanup ^------


        #*======^ END HELPER FUNCTIONS ^======

        #*======v SUB MAIN v======

        # email trigger vari, it will be semi-delimd list of mail-triggering events
        $script:PassStatus = $null ;
        $Alltranscripts = @() ;

        # defer transcript until mbx loop
    
        #-=-=configure EXO EMS aliases to cover useEXOv2 requirements-=-=-=-=-=-=
        # have to preconnect, as it gcm's the targets
        if ($script:useEXOv2) { reconnect-eXO2 }
        else { reconnect-EXO } ;
        # in this case, we need an alias for EXO, and non-alias for EXOP
        [array]$cmdletMaps = 'ps1GetxRcp;get-exorecipient;','ps1GetxMbx;get-exomailbox;','ps1SetxMbx;Set-exoMailbox;','ps1GetxUser;get-exoUser;',
            'ps1SetxCalProc;set-exoCalendarprocessing;','ps1GetxCalProc;get-exoCalendarprocessing;','ps1GetxMbxFldrPerm;get-exoMailboxfolderpermission;',
            'ps1GetxAccDom;Get-exoAcceptedDomain;','ps1GetXRetPol;Get-exoRetentionPolicy','ps1GetxDistGrp;get-exoDistributionGroup;',
            'ps1GetxDistGrpMbr;get-exoDistributionGroupmember;','ps1GetxMsgTrc;get-exoMessageTrace;','ps1GetxMsgTrcDtl;get-exoMessageTraceDetail;',
            'ps1GetxMbxFldrStats;get-exoMailboxfolderStatistics','ps1GetxMContact;Get-exomailcontact;','ps1SetxMContact;Set-exomailcontact;',
            'ps1NewxMContact;New-exomailcontact;' ,'ps1TestxMapi;Test-exoMAPIConnectivity' ;
        foreach($cmdletMap in $cmdletMaps){
            if($script:useEXOv2){
                if(!($cmdlet= Get-Command $cmdletMap.split(';')[1].replace('-exo','-xo') )){ throw "unable to gcm Alias definition!:$($cmdletMap.split(';')[1])" ; break }
                $nalias = set-alias -name ($cmdletMap.split(';')[0]) -value ($cmdlet.name) -passthru ;
                write-verbose "$($nalias.Name) -> $($nalias.ResolvedCommandName)" ;
            } else {
                if(!($cmdlet= Get-Command $cmdletMap.split(';')[1])){ throw "unable to gcm Alias definition!:$($cmdletMap.split(';')[1])" ; break }
                $nalias = set-alias -name ($cmdletMap.split(';')[0]) -value ($cmdlet.name) -passthru ;                
                write-verbose "$($nalias.Name) -> $($nalias.ResolvedCommandName)" ;
            } ;
        } ;
    
        # shifting from ps1 to a function: need updates self-name:
        ${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name ;

        #$sBnr="#*======v START PASS:$($ScriptBaseName) v======" ; 
        $sBnr="#*======v START PASS:$(${CmdletName}) v======" ; 
        $smsg= $sBnr ;   
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;


        # steer all onprem code on $XXXMeta.ExOPAccessFromToro & Ex10Server values
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

        $useEXO = $true ; # non-dyn setting, drives variant EXO reconnect & query code
        if($useEXO){
            #*------v GENERIC EXO CREDS & SVC CONN BP v------
            # o365/EXO creds
            <### Usage: Type defaults to SID, if not spec'd - Note: there must be a *logged in & configured *profile* 
            $o365Cred=get-TenantCredentials -TenOrg $TenOrg -verbose -userrole SID ;
            Returns a credential set for the $TenOrg Hybrid OnPrem Exchange Org
            .EXAMPLE
            $o365Cred=get-TenantCredentials -TenOrg $TenOrg -verbose -userrole CSVC ;
            Returns the CSVC Userrole credential for the $TenOrg Hybrid OnPrem Exchange Org
            .EXAMPLE
            $o365Cred=get-TenantCredentials -TenOrg $TenOrg -verbose -userrole B2BI ;
            Returns the B2BI Userrole credential for the $TenOrg Hybrid OnPrem Exchange Org
            ###>
            $o365Cred=$null ;
            <# $TenOrg is a mandetory param in this script, skip dyn resolution
            switch -regex ($env:USERDOMAIN){
                "(TORO|CMW)" {$TenOrg = $env:USERDOMAIN.substring(0,3).toupper() } ;
                "TORO-LAB" {$TenOrg = 'TOL' }
                default {
                    throw "UNRECOGNIZED `$env:USERDOMAIN!:$($env:USERDOMAIN)" ; 
                    Break ; 
                } ;
            } ; 
            #>
            if($o365Cred=(get-TenantCredentials -TenOrg $TenOrg -UserRole 'CSVC','SID' -verbose:$($verbose))){
                # make it script scope, so we don't have to predetect & purge before using new-variable
                New-Variable -Name cred$($tenorg) -scope Script -Value $o365Cred.cred ;
                $smsg = "Resolved $($Tenorg) `$o365cred:$($o365Cred.cred.username) (assigned to `$cred$($tenorg))" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } else {
                $statusdelta = ";ERROR";
                $script:PassStatus += $statusdelta ;
                set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta) ;
                $smsg = "Unable to resolve $($tenorg) `$o365Cred value!"
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                throw "Unable to resolve $($tenorg) `$o365Cred value!`nEXIT!"
                Break ;
            } ;
            <### CALLS ARE IN FORM: (cred$($tenorg))
            $pltRXO = @{Credential = $Credential ; verbose = $($verbose) ; }
            $pltRXO = @{
                Credential = (Get-Variable -name cred$($tenorg) ).value ;
                verbose = $($verbose) ; } ; 
            if ($script:useEXOv2) { reconnect-eXO2 @pltRXO }
            else { reconnect-EXO @pltRXO } ;
            # or with Tenant-specific cred($Tenorg) lookup
            #$pltRXO creds & .username can also be used for AzureAD connections 
            Connect-AAD @pltRXO ; 
            ###>
            # configure splat for connections: (see above useage)
            $pltRXO = @{
                Credential = (Get-Variable -name cred$($tenorg) ).value ;
                verbose = $($verbose) ; } ; 
            #*------^ END GENERIC EXO CREDS & SVC CONN BP ^------
        } # if-E $useEXO

        if($UseOP){
            #*------v GENERIC EXOP CREDS & SRVR CONN BP v------
            # do the OP creds too
            $OPCred=$null ;
            # default to the onprem svc acct
            $pltGHOpCred=@{TenOrg=$TenOrg ;userrole='ESVC','SID'; verbose=$($verbose)} ;
            if($OPCred=(get-HybridOPCredentials @pltGHOpCred).cred){
                # make it script scope, so we don't have to predetect & purge before using new-variable
                New-Variable -Name "cred$($tenorg)OP" -scope Script -Value $OPCred ;
                $smsg = "Resolved $($Tenorg) `$OPCred:$($OPCred.username) (assigned to `$cred$($tenorg)OP)" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } else {
                $statusdelta = ";ERROR";
                $script:PassStatus += $statusdelta ;
                set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta) ;
                $smsg = "Unable to resolve get-HybridOPCredentials -TenOrg $($TenOrg) -userrole 'ESVC' value!"
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                throw "Unable to resolve $($tenorg) `$OPCred value!`nEXIT!"
                Break ;
            } ;
            $smsg= "Using EXOP cred:`$cred$($tenorg)OP:$((Get-Variable -name "cred$($tenorg)OP" ).value.username)" ;  
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
            <# CALLS ARE IN FORM: (cred$($tenorg))
            $pltRX10 = @{
                Credential = (Get-Variable -name "cred$($tenorg)OP" ).value ;
                verbose = $($verbose) ; }
            ReConnect-Ex2010XO @pltRX10 ; # cross-prem conns
            Reconnect-Ex2010 @pltRX10 ; # local org conns
            #$pltRx10 creds & .username can also be used for local ADMS connections
            #>
            $pltRX10 = @{
                Credential = (Get-Variable -name "cred$($tenorg)OP" ).value ;
                verbose = $($verbose) ; } ;     
            # TEST
        
            # defer cx10/rx10, until just before get-recipients qry
            #*------^ END GENERIC EXOP CREDS & SRVR CONN BP ^------
            # connect to ExOP X10
            if($pltRX10){
                #ReConnect-Ex2010XO @pltRX10 ;
                ReConnect-Ex2010 @pltRX10 ;
            } else { Reconnect-Ex2010 ; } ; 
        } ;  # if-E $useEXOP

        <# already confirmed in modloads
        # load ADMS
        $reqMods += "load-ADMS".split(";") ;
        if ( !(check-ReqMods $reqMods) ) { write-error "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Missing function. EXITING." ; Break ; }  ;
        #>
        write-host -foregroundcolor gray  "(loading ADMS...)" ;
        #write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):MSG" ; 

        load-ADMS ;

        if($UseOP){
            # resolve $domaincontroller dynamic, cross-org
            # setup ADMS PSDrives per tenant 
            if(!$global:ADPsDriveNames){
                $smsg = "(connecting X-Org AD PSDrives)" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                $global:ADPsDriveNames = mount-ADForestDrives -verbose:$($verbose) ;
            } ; 
            if(($global:ADPsDriveNames|measure).count){
                $useEXOforGroups = $false ; 
                $smsg = "Confirming ADMS PSDrives:`n$(($global:ADPsDriveNames.Name|%{get-psdrive -Name $_ -PSProvider ActiveDirectory} | ft -auto Name,Root,Provider|out-string).trim())" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                # returned object
                #         $ADPsDriveNames
                #         UserName                Status Name        
                #         --------                ------ ----        
                #         DOM\Samacctname   True  [forestname wo punc] 
                #         DOM\Samacctname   True  [forestname wo punc]
                #         DOM\Samacctname   True  [forestname wo punc]
        
            } else { 
                #-=-record a STATUS=-=-=-=-=-=-=
                $statusdelta = ";ERROR";
                $script:PassStatus += $statusdelta ;
                set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta) ;
                #-=-=-=-=-=-=-=-=
                $smsg = "Unable to detect POPULATED `$global:ADPsDriveNames!`n(should have multiple values, resolved to $()"
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                throw "Unable to resolve $($tenorg) `$o365Cred value!`nEXIT!"
                Break ;
            } ; 
        } ; 
        #if (!$domaincontroller) { $domaincontroller = get-gcfast } ;
        #if(!$domaincontroller){ if(test-path function:get-gcfast){$domaincontroller=get-gcfast} else { throw "no get-gcfast()!" } ;} else {"(existing `$domaincontroller:$($domaincontroller))"} ;
        # use new get-GCFastXO cross-org dc finde
        # default to Op_ExADRoot forest from $TenOrg Meta
        $domaincontroller = get-GCFastXO -TenOrg $TenOrg -subdomain ((gv -name "$($TenOrg)Meta").value['OP_ExADRoot']) -verbose:$($verbose) |?{$_.length};


        <# MSOL CONNECTION
        $reqMods += "connect-msol".split(";") ;
        if ( !(check-ReqMods $reqMods) ) { write-error "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Missing function. EXITING." ; Break ; }  ;
        write-host -foregroundcolor gray  "(loading AAD...)" ;
        #connect-msol ;
        connect-msol @pltRXO ; 
        #>

        # AZUREAD CONNECTION
        $reqMods += "Connect-AAD".split(";") ;
        if ( !(check-ReqMods $reqMods) ) { write-error "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Missing function. EXITING." ; Break ; }  ;
        write-host -foregroundcolor gray  "(loading AAD...)" ;
        #connect-msol ;
        Connect-AAD @pltRXO ; 
        #


        #
        # EXO connection
        $pltRXO = @{
            Credential = (Get-Variable -name cred$($tenorg) ).value ;
            verbose = $($verbose) ; } ; 
        if($VerbosePreference = "Continue"){
            $VerbosePrefPrior = $VerbosePreference ;
            $VerbosePreference = "SilentlyContinue" ;
            $verbose = ($VerbosePreference -eq "Continue") ;
        } ; 
        disconnect-exo ;
        if ($script:useEXOv2) { reconnect-eXO2 @pltRXO }
        else { reconnect-EXO @pltRXO } ;
        # reenable VerbosePreference:Continue, if set, during mod loads 
        if($VerbosePrefPrior -eq "Continue"){
            $VerbosePreference = $VerbosePrefPrior ;
            $verbose = ($VerbosePreference -eq "Continue") ;
        } ;
        #

        
        # 3:00 PM 9/12/2018 shift this to 1x in the script ; - this would need to be customized per tenant, not used (would normally be for forcing UPNs, but CMW uses brand UPN doms)
        #$script:forestdom = ((get-adforest | select -expand upnsuffixes) | ? { $_ -eq 'toro.com' }) ;

        # Clear error variable
        $Error.Clear() ;
        

    } ;  # BEGIN-E
    PROCESS {
        $rMailboxes = @() ;
        $Error.Clear() ;
        $SearchesRun = 0 ;
        $smsg = "Net:$(($Mailboxes|measure).count) mailboxes" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        $ttl = ($Mailboxes | measure).count ;
        $Procd = 0 ;

        # 9:20 AM 2/25/2019 Tickets will be an array of nnn's to match the mbxs, so use $Procd-1 as the index for tick# in the array
        if((($Mailboxes|measure).count -gt 1) -AND (($Tickets|measure).count -eq 1)){
            # Mult mbxs with single Ticket, expand ticket count to match # of mbxes
            foreach($mbx in $Mailboxes){
                $Tickets+=$Tickets[0] ; 
            } ; 
        } ; 
        foreach($Mailbox in $Mailboxes){
            $Procd++ ;
        
            if ($Ticket = $Tickets[($Procd - 1)]) {
                $sBnr = "#*======v `$Ticket:$($ticket):`$Mailbox:($($Procd)/$($ttl)):$($Mailbox) v======" ;
                $ofileroot = ".\logs\$($ticket)-" ;
                $lTag = "$($ticket)-$($Mailbox)-" ; 
            }else {
                $sBnr = "#*======v `$Ticket:(not spec'd):`$Mailbox:($($Procd)/$($ttl)):$($Mailbox) v======" ;
                $ofileroot = ".\logs\" ;
                $lTag = "$($Mailbox)-" ; 
            } ;

            # detect profile installs (installed mod or script), and redir to stock location
            $dPref = 'd','c' ; foreach($budrv in $dpref){ if(test-path -path "$($budrv):\scripts" -ea 0 ){ break ;  } ;  } ;
            [regex]$rgxScriptsModsAllUsersScope="^$([regex]::escape([environment]::getfolderpath('ProgramFiles')))\\((Windows)*)PowerShell\\(Scripts|Modules)" ;
            [regex]$rgxScriptsModsCurrUserScope="^$([regex]::escape([environment]::getfolderpath('Mydocuments')))\\((Windows)*)PowerShell\\(Scripts|Modules)" ;
            # -Tag "($TenOrg)-LASTPASS" 
            $pltSLog = [ordered]@{ NoTimeStamp=$false ; Tag=$lTag  ; showdebug=$($showdebug) ;whatif=$($whatif) ;} ;
            if($PSCommandPath){
                if(($PSCommandPath -match $rgxScriptsModsAllUsersScope) -OR ($PSCommandPath -match $rgxScriptsModsCurrUserScope) ){
                    # AllUsers or CU installed script, divert into [$budrv]:\scripts (don't write logs into allusers context folder)
                    if($PSCommandPath -match '\.ps(d|m)1$'){
                        # module function: use the ${CmdletName} for childpath
                        $pltSLog.Path= (join-path -Path "$($budrv):\scripts" -ChildPath "$(${CmdletName}).ps1" )  ;
                    } else { 
                        $pltSLog.Path=(join-path -Path "$($budrv):\scripts" -ChildPath (split-path $PSCommandPath -leaf)) ;
                    } ; 
                }else {
                    $pltSLog.Path=$PSCommandPath ;
                } ;
            } else {
                if( ($MyInvocation.MyCommand.Definition -match $rgxScriptsModsAllUsersScope) -OR ($MyInvocation.MyCommand.Definition -match $rgxScriptsModsCurrUserScope) ){
                    $pltSLog.Path=(join-path -Path "$($budrv):\scripts" -ChildPath (split-path $PSCommandPath -leaf)) ;
                } else {
                    $pltSLog.Path=$MyInvocation.MyCommand.Definition ;
                } ;
            } ;
            $smsg = "start-Log w`n$(($pltSLog|out-string).trim())" ;
            if($verbose){ if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            $logspec = start-Log @pltSLog ;

            if($logspec){
                $stopResults = try {Stop-transcript -ErrorAction stop} catch {} ;
                $logging=$logspec.logging ;
                $logfile=$logspec.logfile ;
                $transcript=$logspec.transcript ;
                #Configure default logging from parent script name
                # logfile                        C:\usr\work\o365\scripts\logs\move-MailboxToXo-(TOR)-LASTPASS-LOG-BATCH-WHATIF-log.txt
                # transcript                     C:\usr\work\o365\scripts\logs\move-MailboxToXo-(TOR)-LASTPASS-Transcript-BATCH-WHATIF-trans-log.txt
                #$logfile = $logfile.replace('-LASTPASS','').replace('BATCH',(Remove-InvalidFileNameChars -name $BatchName )) ;
                $logfile = $logfile.replace('-LASTPASS','').replace('BATCH','') ;
                #$transcript = $transcript.replace('-LASTPASS','').replace('BATCH',(Remove-InvalidFileNameChars -name $BatchName )) ;
                $transcript = $transcript.replace('-LASTPASS','').replace('BATCH','') ;
                if(Test-TranscriptionSupported){start-transcript -Path $transcript }
                else { write-warning "$($host.name) v$($host.version.major) does *not* support Transcription!" } ;
            } else {throw "Unable to configure logging!" } ;
    
            $smsg = "$($sBnr)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

            $sBnr="#*======v MAILBOX:$($Mailbox) v======" ;
            $smsg = "$((get-date).ToString('HH:mm:ss')):$($sBnr)" ;
            $pltXcsv = [ordered]@{Path=$null; NoTypeInformation=$true ; } ;
            $error.clear() ;

            $rcp = $null ; $exorcp = $null ;
            $exombx = $null ; $opombx = $null ;  $rmbx = $null ;
            $targetUPN = $NULL ;
            $adu = $null ; $msolu = $null ;
            $bMissingMbx = $false ;
            $exolicdetails = $null ;
            $obj = $null ;
            $grantgroup = $null ;
            $tuser = $null ;
            $UPN = $null ;
            $MailboxFoldersExo = $null ;

            $Exit = 0 ;

            if($pltRXO){
                if ($script:useEXOv2) { reconnect-eXO2 @pltRXO }
                else { reconnect-EXO @pltRXO } ;
            } else {reconnect-exo ;} ; 
            #Reconnect-Ex2010 ;
            if($pltRX10){
                #ReConnect-Ex2010XO @pltRX10 ;
                ReConnect-Ex2010 @pltRX10 ;
            } else { Reconnect-Ex2010 ; } ; 

            $Exit = 0 ; $error.clear() ;
            Do {
                Try {
                    $rcp = get-recipient -id $Mailbox -domaincontroller $domaincontroller -ea stop ;
                
                    $smsg= "(successful get-recipient -id $($Mailbox))" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else { write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    $Exit = $Retries ;
                }
                Catch {
                    $errTrpd=$_ ; 
                    if( $errtrpd -match "\scouldn't\sbe\sfound\son\s" ){
                        $smsg = "(no EXOP recicpient found)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                        else { write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        $Exit = $Retries ;
                    } else { 
                        $smsg=": Error Details: $($errTrpd)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                        Start-Sleep -Seconds $RetrySleep ;
                        if($pltRX10){
                            #ReConnect-Ex2010XO @pltRX10 ;
                            ReConnect-Ex2010 @pltRX10 ;
                        } else { ReConnect-Ex2010  } ; 
                        $Exit ++ ;
                        Write-Verbose "Failed to exec cmd because: $($Error[0])" ;
                        Write-Verbose "Try #: $Exit" ;                    
                    } ; 
                    If ($Exit -eq $Retries) { Write-Warning "Unable to exec cmd!" } ;
                }  ;
            } Until ($Exit -eq $Retries) ;

            $Exit = 0 ; $error.clear() ;
            Do {
                Try {
                    $pltgxrcp=[ordered]@{identity=$Mailbox ; erroraction='stop'; } ; 

                    $smsg= "$((get-alias ps1GetxRcp).definition) w`n$(($pltgxrcp|out-string).trim())" ; 
                    if($verbose){ if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 

                    $exorcp = ps1GetxRcp @pltgxrcp ;
                
                    $smsg= "(successful get-exorecipient -id $($Mailbox))" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                
                    $Exit = $Retries ;
                }
                Catch {
                    # get fails go here with: The operation couldn't be performed because object 'Anthony.Magana@toro.com' couldn't be found on 'CY4PR04A008DC10.NAMPR04A008.PROD.OUTLOOK.COM'.
                    $errTrpd=$_ ; 
                    if( $errtrpd -match "\scouldn't\sbe\sfound\son\s" ){
                        $smsg = "(no EXO recipient found)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                        else { write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        $Exit = $Retries ;
                    } else { 
                        $smsg=": Error Details: $($errTrpd)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                        Start-Sleep -Seconds $RetrySleep ;
                        if($pltRXO){
                            if ($script:useEXOv2) { reconnect-eXO2 @pltRXO }
                            else { reconnect-EXO @pltRXO } ;
                        } else {reconnect-exo ;} ; 
                        $Exit ++ ;
                        Write-Verbose "Failed to exec cmd because: $($Error[0])" ;
                        Write-Verbose "Try #: $Exit" ;                    
                    } ; 
                    If ($Exit -eq $Retries) { Write-Warning "Unable to exec cmd!" } ;
                }  ;
            } Until ($Exit -eq $Retries) ;


            $Exit = 0 ; $error.clear() ;
            Do {
                Try {
                    $pltgxmbx=[ordered]@{identity=$Mailbox ; erroraction='stop'; } ; 
                    $smsg= "$((get-alias ps1GetxMbx).definition) w`n$(($pltgxmbx|out-string).trim())" ; 
                    if($verbose){ if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 

                
                    $exombx = ps1GetxMbx @pltgxmbx ; 

                    $smsg= "(successful get-exomailbox -id $($Mailbox))" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;


                    if ($exombx) {
                        $pltGetxMbxFldrStats=[ordered]@{Identity=$exombx.identity ;IncludeOldestAndNewestItems=$true; erroraction='stop'; } ; 
                        #$smsg= "(collecting get-exomailboxfolderstatistics...)" ;
                        $smsg= "$((get-alias ps1GetxMbxFldrStats).definition) w`n$(($pltGetxMbxFldrStats|out-string).trim())" ; 
                        if($verbose){ if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                        $exofldrs = ps1GetxMbxFldrStats @pltGetxMbxFldrStats | 
                            ?{($_.ItemsInFolder -gt 0 ) -AND ($_.identity -notmatch $rgxExoSysFolders)} ;
                        #$fldrs = get-exomailboxfolderstatistics -id KADRITS -IncludeOldestAndNewestItems ;
                        #$fldrs = $fldrs |?{($_.ItemsInFolder -gt 0 ) -AND ($_.identity -notmatch $rgxExoSysFolders)} ;
                        #$fldrs | ft -auto $propsmbxfldrs ;

                        # do a 7day msgtrc
                        #-=-=-=-=-=-=-=-=
                        $isplt=@{    
                            ticket=$ticket ;
                            uid=$Mailbox;
                            days=$Days ;
                            StartDate='' ;
                            EndDate='' ;
                            Sender="" ;
                            Recipients=$exombx.PrimarySmtpAddress ;
                            MessageSubject="" ;
                            Status='' ;
                            MessageTraceId='' ;
                            MessageId='' ;
                        }  ;
                        $msgtrk=@{ PageSize=1000 ; Page=$null ; StartDate=$null ; EndDate=$null ; } ;

                        if($isplt.days){  
                            $msgtrk.StartDate=(get-date ([datetime]::Now)).adddays(-1*$isplt.days);
                            $msgtrk.EndDate=(get-date) ;
                        } ;
                        if($isplt.StartDate -and !($isplt.days)){$msgtrk.StartDate=$(get-date $isplt.StartDate)} ;
                        if($isplt.EndDate -and !($isplt.days)){$msgtrkEndDate=$(get-date $isplt.EndDate)} 
                        elseif($isplt.StartDate -and !($isplt.EndDate)){
                            $smsg = '(StartDate w *NO* Enddate, asserting currenttime)' ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            $msgtrkEndDate=(get-date) ;
                        } ;

                        TRY{$tendoms=Get-AzureADDomain }CATCH{
                            #write-warning "NOT AAD CONNECTED!" ;BREAK ;
                            $smsg = "Not AAD connected, reconnect..." ; 
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            Connect-AAD @pltRXO ; 
                        } ;

                        $Ten = ($tendoms |?{$_.name -like '*.mail.onmicrosoft.com'}).name.split('.')[0] ;
                        $ofile ="$($isplt.ticket)-$($Ten)-$($isplt.uid)-EXOMsgTrk" ;
                        if($isplt.Sender){
                            if($isplt.Sender -match '\*'){
                                $smsg = "(wild-card Sender detected)" ;
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                $msgtrk.add("SenderAddress",$isplt.Sender) ;
                            } else {
                                $msgtrk.add("SenderAddress",$isplt.Sender) ;
                            } ;
                            $ofile+=",From-$($isplt.Sender.replace("*","ANY"))" ;
                        } ;
                        if($isplt.Recipients){
                            if($isplt.Recipients -match '\*'){
                                $smsg = "(wild-card Recipient detected)" ;
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                $msgtrk.add("RecipientAddress",$isplt.Recipients) ;
                            } else {
                                $msgtrk.add("RecipientAddress",$isplt.Recipients) ;
                            } ;
                            $ofile+=",To-$($isplt.Recipients.replace("*","ANY"))" ;
                        } ;

                        if($isplt.MessageId){
                            $msgtrk.add("MessageId",$isplt.MessageId) ;
                            $ofile+=",MsgId-$($isplt.MessageId.replace('<','').replace('>',''))" ;
                        } ;
                        if($isplt.MessageTraceId){
                            $msgtrk.add("MessageTraceId",$isplt.MessageTraceId) ;
                            $ofile+=",MsgId-$($isplt.MessageTraceId.replace('<','').replace('>',''))" ;
                        } ;
                        if($isplt.MessageSubject){
                            $ofile+=",Subj-$($isplt.MessageSubject.substring(0,[System.Math]::Min(10,$isplt.MessageSubject.Length)))..." ;
                        } ;
                        if($isplt.Status){
                            $msgtrk.add("Status",$isplt.Status)  ;
                            $ofile+=",Status-$($isplt.Status)" ;
                        } ;
                        if($isplt.days){$ofile+= "-$($isplt.days)d-" } ;
                        if($isplt.StartDate){$ofile+= "-$(get-date $isplt.StartDate -format 'yyyyMMdd-HHmmtt')-" } ;
                        if($isplt.EndDate){$ofile+= "$(get-date $isplt.EndDate -format 'yyyyMMdd-HHmmtt')" } ;

                        $smsg = "Running MsgTrk:$($Ten)" ;
                        $smsg += "`n$(($msgtrk|out-string).trim()|out-default)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                        $Page = 1  ;
                        $Msgs=$null ;
                        do {
                            $smsg = "Collecting - Page $($Page)..."  ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                            $msgtrk.Page=$Page ;
                            $PageMsgs = ps1GetxMsgTrc @msgtrk |  ?{$_.SenderAddress -notlike '*micro*' -or $_.SenderAddress -notlike '*root*' }  ;
                            $Page++  ;
                            $Msgs += @($PageMsgs)  ;
                        } until ($PageMsgs -eq $null) ;
                        $Msgs=$Msgs| Sort Received ;
                    
                        $smsg = "==Msgs Returned:$(($Msgs|measure).count)`nRaw matches:$(($Msgs|measure).Count)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                        if($isplt.MessageSubject){
                            $smsg = "Post-Filtering on Subject:$($isplt.MessageSubject)" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                            $Msgs = $Msgs | ?{$_.Subject -like $isplt.MessageSubject} ;
                            $ofile+="-Subj-$($isplt.MessageSubject.replace("*"," ").replace("\"," "))" ;
                            $smsg = "Post Subj filter matches:$(($Msgs|measure).Count)" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                        } ;
                        $ofile+= "-run$(get-date -format 'yyyyMMdd-HHmmtt').csv" ;
                        #$ofile=[RegEx]::Replace($ofile, "[{0}]" -f ([RegEx]::Escape(-join [System.IO.Path]::GetInvalidFileNameChars())), '') ;
                        $ofile = Remove-InvalidFileNameChars -name $ofile ; 
                        $ofile=".\logs\$($ofile)" ;
                     
                        if($Msgs){
                            $Msgs | select  | export-csv -notype -path $ofile  ;
                            "Status Distrib:" ;
                            $smsg = "`n#*------v MOST RECENT MATCH v------" ;
                            $smsg += "`n$(($msgs[-1]| format-list ReceivedLocal,StatusSenderAddress,RecipientAddress,Subject|out-string).trim())";
                            $smsg += "`n#*------^ MOST RECENT MATCH ^------" ;
                            $smsg += "`n#*------v Status DISTRIB v------" ;
                            $smsg += "`n$(($Msgs | select -expand Status | group | sort count,count -desc | select count,name |out-string).trim())";
                            $smsg += "`n#*------^ Status DISTRIB ^------" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            <# Count Name                      Group
                            ----- ----                      -----
                              448 Delivered                 {@{PSComputerName=ps.outlook.com; RunspaceId=25f3aa28-9437-4e30-aa8f-8d83d2d2fc5a; PSShowComputerName=False; Organization=toroco.onmicrosoft.com; MessageId=<BN8PR12MB33158F1D07B3862078D758E6EB499@BN8PR12MB3315.namprd12.prod.o...
                                1 Failed                    {@{PSComputerName=ps.outlook.com; RunspaceId=25f3aa28-9437-4e30-aa8f-8d83d2d2fc5a; PSShowComputerName=False; Organization=toroco.onmicrosoft.com; MessageId=<5bc2055fed4d43d3827cf7f61d37a4c9@CH2PR04MB7062.namprd04.prod.outlook...
                                1 Quarantined               {@{PSComputerName=ps.outlook.com; RunspaceId=25f3aa28-9437-4e30-aa8f-8d83d2d2fc5a; PSShowComputerName=False; Organization=toroco.onmicrosoft.com; MessageId=<threatsim-5f0bc0101d-c200b2590d@app.emaildistro.com>; Received=4/28/...
                                1 FilteredAsSpam            {@{PSComputerName=ps.outlook.com; RunspaceId=25f3aa28-9437-4e30-aa8f-8d83d2d2fc5a; PSShowComputerName=False; Organization=toroco.onmicrosoft.com; MessageId=<SA0PR01MB61858E3C6111672E081373C1E45F9@SA0PR01MB6185.prod.exchangela...
                            #>
                            $nonDelivStats = $msgs | ?{$_.status -ne 'Delivered'} | group status | select-object -expand name ; 
                            
                            foreach ($status in $nonDelivStats){
                                $smsg = "Enumerating Status:$($status) messages" ; 
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                if($StatMsgs  = $msgs|?{$_.status -eq $status}){
                                    if(($StatMsgs |measure).count -gt 10){
                                        $smsg = "(over $($MsgTraceNonDeliverDetailsLimit) $($status) msgs: processing a sample of last $($MsgTraceNonDeliverDetailsLimit)...)" ; 
                                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                        $StatMsgs = $StatMsgs | select-object -Last $MsgTraceNonDeliverDetailsLimit ;
                                    } ; 
                                    $statTtl = ($StatMsgs |measure).count ; $fProcd=0 ; 
                                    $smsg = "$($statTtl) Status:'$($status)' messages returned. Expanding detailed processing..." ; 
                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                        foreach($fmsg in $StatMsgs ){
                                            $fProcd++ ; 
                                            $sBnrS="`n#*------v PROCESSING $($status)#$($fProcd)/$($statTtl): v------" ; 
                                            $smsg= "$($sBnrS)" ;
                                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                                            $pltGetxMsgTrcDtl = [ordered]@{MessageTraceId=$fmsg.MessageTraceId ;RecipientAddress=$fmsg.RecipientAddress} ; 
                                            $smsg= "$((get-alias ps1GetxMsgTrcDtl).definition) w`n$(($pltGetxMsgTrcDtl|out-string).trim())" ; 
                                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                
                                            # this is nested in a try already
                                            $MsgTrkDtl = ps1GetxMsgTrcDtl @pltGetxMsgTrcDtl ; 

                                            if($MsgTrkDtl){
                                                $smsg = "`n-----`n$(( $MsgTrkDtl | fl $propsMsgTrcDtl|out-string).trim())`n-----`n" ; 
                                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                                                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                            } else { 
                                                $smsg = "(no matching MessageTraceDetail was returned by MS for MsgTrcID:$($pltGetxMsgTrcDtl.MessageTraceId), aged out of availability already)" ; 
                                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                                                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                                $smsg = "$($status)#$($fProcd)/$($statTtl):DETAILS`n----`n$(($fmsg|fl | out-string).trim())`n----`n" ; 
                                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                                                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                            } ; 

                                            $smsg= "$($sBnrS.replace('-v','-^').replace('v-','^-'))" ;
                                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                        } ; 
                                
                                }
                            } ; #nonDelivStats
                            if(test-path -path $ofile){
                                $smsg = "(log file confirmed)" ;
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                Resolve-Path -Path $ofile | select -expand Path | out-clipboard ;
                                $smsg = "$($Msgs.count) matches output to:`n'$($ofile)'`n(copied to CB)" ;
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                            } else { 
                                $smsg = "MISSING LOG FILE!"  ; 
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            } ;
                        } else {
                            $smsg = "NO MATCHES FOUND from::`n$(($msgtrk|out-string).trim()|out-default)`n(with any relevant ConnectorId postfilter)" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        } ;

                        #-=-=-=-=-=-=-=-=


                    } else {
                        $smsg = "(no EXO mailbox found, skipping $((get-alias ps1GetxMbxFldrStats).definition) )" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                        else { write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ;
                    $Exit = $Retries ;
                }
                Catch {
                    # get fails go here with: The operation couldn't be performed because object 'Anthony.Magana@toro.com' couldn't be found on 'CY4PR04A008DC10.NAMPR04A008.PROD.OUTLOOK.COM'.
                    $errTrpd=$_ ; 
                    if( $errtrpd -match "\scouldn't\sbe\sfound\son\s" ){
                        $smsg = "(no EXO mailbox found)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                        else { write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        $Exit = $Retries ;
                    } else { 
                        $smsg=": Error Details: $($errTrpd)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                        Start-Sleep -Seconds $RetrySleep ;
                        if($pltRXO){
                            if ($script:useEXOv2) { reconnect-eXO2 @pltRXO }
                            else { reconnect-EXO @pltRXO } ;
                        } else {reconnect-exo ;} ; 
                        $Exit ++ ;
                        Write-Verbose "Failed to exec cmd because: $($Error[0])" ;
                        Write-Verbose "Try #: $Exit" ;                    
                    } ; 
                    If ($Exit -eq $Retries) { Write-Warning "Unable to exec cmd!" } ;
                }  ;
            } Until ($Exit -eq $Retries) ;

            $Exit = 0 ; $error.clear() ;
            Do {
                Try {
                    $pltgRmbx=@{identity=$Mailbox ; domaincontroller=$domaincontroller ; erroraction='stop'; } ; 
                    $smsg= "get-RemoteMailbox w`n$(($pltgRmbx|out-string).trim())" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                    $rmbx = get-remotemailbox @pltgRmbx ;

                    if($rmbx){
                        $smsg= "(successful get-remotemailbox  -id $($Mailbox))" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ;
                    $Exit = $Retries ;
                }
                Catch {
                    # get fails go here with: The operation couldn't be performed because object 'blahblah' couldn't be found on 'BCCMS8100.global.ad.toro.com'.
                    $errTrpd=$_ ; 
                    if( $errtrpd -match "\scouldn't\sbe\sfound\son\s" ){
                        $smsg = "(no EXOP remotemailbox found)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                        else { write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        $Exit = $Retries ;
                    } else { 
                        $smsg=": Error Details: $($errTrpd)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                        Start-Sleep -Seconds $RetrySleep ;
                        if($pltRX10){
                            #ReConnect-Ex2010XO @pltRX10 ;
                            ReConnect-Ex2010 @pltRX10 ;
                        } else { ReConnect-Ex2010  } ; 
                        $Exit ++ ;
                        Write-Verbose "Failed to exec cmd because: $($Error[0])" ;
                        Write-Verbose "Try #: $Exit" ;                    
                    } ; 
                    If ($Exit -eq $Retries) { Write-Warning "Unable to exec cmd!" } ;
                }  ;
            } Until ($Exit -eq $Retries) ;


            if ($exombx ) {
                $targetUPN = $exombx.userprincipalname ;
            }
            elseif ($rmbx) {
                $targetUPN = $rmbx.userprincipalname ;
            }
            else {
                throw "Unable to locate either an EXO mailbox or an on-prem RemoteMailbox. ABORTING!"
                Break ;
            } ;


            $Exit = 0 ; $error.clear() ;
            Do {
                Try {
                    $pltgMbx=[ordered]@{identity=$targetUPN ;domaincontroller=$domaincontroller ;erroraction='SilentlyContinue';  };
                    $smsg = "get-mailbox w`n$(($pltgMbx|out-string).trim())" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                    $opmbx = get-mailbox @pltgMbx ; 
                    if($opmbx){
                        $smsg= "(successful get-mailbox -id $($targetUPN))" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ;
                    $Exit = $Retries ;

                }
                Catch {
                    # get fails go here with: The operation couldn't be performed because object 'blahblah' couldn't be found on 'BCCMS8100.global.ad.toro.com'.
                    $errTrpd=$_ ; 
                    if( $errtrpd -match "\scouldn't\sbe\sfound\son\s" ){
                        $smsg = "(no EXOP mailbox found)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                        else { write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        $Exit = $Retries ;
                    } else { 
                        $smsg=": Error Details: $($errTrpd)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                        Start-Sleep -Seconds $RetrySleep ;
                        if($pltRX10){
                            #ReConnect-Ex2010XO @pltRX10 ;
                            ReConnect-Ex2010 @pltRX10 ;
                        } else { ReConnect-Ex2010  } ; 
                        $Exit ++ ;
                        Write-Verbose "Failed to exec cmd because: $($Error[0])" ;
                        Write-Verbose "Try #: $Exit" ;                    
                    } ; 
                    If ($Exit -eq $Retries) { Write-Warning "Unable to exec cmd!" } ;
                }  ;
            } Until ($Exit -eq $Retries) ;


            $Exit = 0 ; $error.clear() ;
            Do {
                Try {
                    connect-msol @pltRXO ;
                    $pltgMSOLU = [ordered]@{userprincipalname=$targetUPN  ;erroraction='SilentlyContinue';} ; 
                    $smsg = "get-msoluser w`n$(($pltgMSOLU|out-string).trim())" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                    $msolu = get-msoluser @pltgMSOLU ;
                    if($msolu){
                        $smsg= "(successful get-msoluser -userprincipalname $($targetUPN))" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ;
                    $Exit = $Retries ;
                }
                Catch {
                    # get fails go here with: get-msoluser : User Not Found.  User: blah@toro.com.
                    $errTrpd=$_ ; 
                    if( $errtrpd -match "User\sNot\sFound" ){
                        $smsg = "(no EXOP mailbox found)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                        else { write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        $Exit = $Retries ;
                    } else { 
                        $smsg=": Error Details: $($errTrpd)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                        Start-Sleep -Seconds $RetrySleep ;
                        if($pltRXO){
                            if ($script:useEXOv2) { reconnect-eXO2 @pltRXO }
                            else { reconnect-EXO @pltRXO } ;
                        } else {reconnect-exo ;} ; 
                        $Exit ++ ;
                        Write-Verbose "Failed to exec cmd because: $($Error[0])" ;
                        Write-Verbose "Try #: $Exit" ;                    
                    } ; 
                    If ($Exit -eq $Retries) { Write-Warning "Unable to exec cmd!" } ;
                }  ;
            } Until ($Exit -eq $Retries) ;

            if ($msolu) {
                connect-msol @pltRXO;
               # to find group, check user's membership
                # do the full lookup first

                if ($msolu.IsLicensed -AND !($msolu.LicenseReconciliationNeeded)) {
                    $smsg = "USER HAS *NO* LICENSING ISSUES:" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else { write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                }
                else {
                    $smsg = "USER *HAS* LICENSING ISSUES:" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else { write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                } ;

                $smsg = "Get-MsolUser:`n$(($msolu | select userprin*,*Error*,*status*,softdel*,lic*,islic*|out-string).trim())`n" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else { write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                if ($null -eq $msolu.SoftDeleteTimestamp) {
                    $smsg = "$($msol.userprincipalname) has a BLANK SoftDeleteTimestamp`n=>If mailbox missing it would indicate the user wasn't properly de-licensed (or would have fallen into dumpster at >30d).`n That scenario would reflect a replic break (AAD sync loss wo proper update)`n" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else { write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                } ;

                if ($EXOLicDetails = get-MsolUserLicenseDetails -UPNs $targetUPN -showdebug:$($showdebug) ) {
                    $smsg = "Returned License Details`n$(($EXOLicDetails|out-string).trim())" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } ; #Error|Warn
                }
                else {
                    $smsg = "UNABLE TO RETURN License Details FOR `n$(($targetUPN|out-string).trim())" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } ; #Error|Warn
                };

                # 2:12 PM 1/27/2020 rough in Pull-AADSignInReports.ps1 syntax
                $AADSigninSyntax=@"

---------------------------------------------------------------------------------
If you would like to retrieve user AAD Sign On Reports, use the following command:

A. Query the events into json outputs:

.\Pull-AADSignInReports.ps1 -UPNs "$($targetUPN)" -ticket "$($Ticket)" -StartDate (Get-Date).AddDays(-30) -showdebug ;

B. Process & Profile the output .json file from the above:
.\profile-AAD-Signons.ps1 -Files PATH-TO-FILE.json ;

---------------------------------------------------------------------------------

"@ ;

            }
            else {
                $smsg = "$($targetUPN):NO MATCHING MSOLUSER RETURNED!" ; ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } ;

        
            $Exit = 0 ; $error.clear() ;
            Do {
                Try {
                    $aduFilter = { UserPrincipalName -eq $targetUPN }  ;
                    $pltgADU=[ordered]@{filter=$aduFilter;Properties=$adprops ;server=$domaincontroller  ;erroraction='SilentlyContinue'; } ; 
                    $smsg = "get-aduser w`n$(($pltgADU|out-string).trim())" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                    #$adu = get-aduser -filter { UserPrincipalName -eq $targetUPN } -Properties $adprops -server $domaincontroller -ea 0  ; 
                    $adu = get-aduser @pltgADU  ; 
                    if($adu){
                        $smsg= "(successful get-aduser -filter { UserPrincipalName -eq $($targetUPN) }" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ;
                    $Exit = $Retries ;
                }
                Catch {
                    $errTrpd=$_ ;
                    $smsg = "Failed processing $($errTrpd.Exception.ItemName). `nError Message: $($errTrpd.Exception.Message)`nError Details: $($errTrpd)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Error } #Error|Warn|Debug
                    else { write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    Start-Sleep -Seconds $RetrySleep ;
                    $Exit ++ ;
                    $smsg = "Failed to exec cmd because: $($Error[0])`nTry #: $($Exit)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Error } #Error|Warn|Debug
                    else { write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    If ($Exit -eq $Retries) { Write-Warning "Unable to exec cmd!" } ;
                }  ;
            } Until ($Exit -eq $Retries) ;

            # DETERMINE LIC GRANT GROUP
            if ($GrantGroupDN = $adu | select -expand memberof | ? { $_ -match $rgxEntLicGrps } ) {
                Try {
                    # 8:47 AM 3/1/2019 #626:expand the name, (was coming out a hash)
                    $GrantGroup = get-adgroup -id $grantgroupDN | select -expand name ;
                }
                catch {
                    $grantgroup = "(NOT RESOLVABLE)"
                } ;
            }
            else {
                $grantgroup = "(NOT RESOLVABLE)"
            } ;

            # TRIAGE RESULTS
            if ($exombx -AND $opmbx) {
                # SPLIT BRAIN
                $smsg = "USER HAS SPLIT-BRAIN MAILBOX (MBX IN EXO & EXOP)!" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                else { write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

            }elseif (($rcp.RecipientTypeDetails -eq 'RemoteUserMailbox') -AND ($exorcp.RecipientTypeDetails -eq 'MailUser') ) {
                # NO BRAIN
                $smsg = "USER HAS NO-BRAIN MAILBOX (*NO* MBX IN *EITHER* EXO & EXOP)!" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                else { write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

             }
            elseif ( ($rcp.RecipientTypeDetails -match '(User|Shared|Room|Equipment)Mailbox') -AND ($exorcp.RecipientTypeDetails -eq 'MailUser') ) {
                # ON PREM MAILBOX
                #$mapiTest = Test-MAPIConnectivity -id $Mailbox ;
                $pltTMapi=[ordered]@{identity=$Mailbox;domaincontroller=$domaincontroller  ;erroraction='SilentlyContinue'; } ; 
                $smsg = "Test-MAPIConnectivity w`n$(($pltTMapi|out-string).trim())" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                Try {
                    $mapiTest = Test-MAPIConnectivity @pltTMapi  ; 
                    if ($mapiTest.Result -eq 'Success') {
                        $smsg= "(successful Outlook Mbx MAPI validate: Test-MAPIConnectivity -id $($Mailbox) }" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ;
                    $Exit = $Retries ;
                } Catch {
                    $errTrpd=$_ ;
                    $smsg = "Failed processing $($errTrpd.Exception.ItemName). `nError Message: $($errTrpd.Exception.Message)`nError Details: $($errTrpd)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Error } #Error|Warn|Debug
                    else { write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    Start-Sleep -Seconds $RetrySleep ;
                    $Exit ++ ;
                    $smsg = "Failed to exec cmd because: $($Error[0])`nTry #: $($Exit)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Error } #Error|Warn|Debug
                    else { write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    If ($Exit -eq $Retries) { Write-Warning "Unable to exec cmd!" } ;
                }  ;


                $FixText = @"
========================================
Problem User:`t "$($adu.UserPrincipalName)"
Has an *ONPREM* mailbox:Type:$($exombx.recipienttypedetails) hosted in $($opmbx.database)
"@ ;
                if ($mapiTest.Result -eq 'Success') {
                    $FixText2 = @"
*with *NO* detected issues*
Mailbox Outlook connectivity tested and validated functional...
"@ ;
                } else {
                    $FixText2 = @"
*FAILED Test-MAPIConnectivity*
Mailbox Outlook connectivity test failed to connect!...
"@ ;
                }

                $FixText3 = @"

$(($mapiTest|out-string).trim())

The user's o365 LICENSESKUs:                     "$($exolicdetails.LicAccountSkuID)"
With DisplayNames:                               "$($exolicdetails.LicenseFriendlyName)"
The user's o365 Licensing group appears to be:  "$($grantgroup)"

OnPrem RecipientTypeDetails:`t "$($rcp.RecipientTypeDetails)"
OnPrem WhenCreated:`t "$($rcp.WhenCreated)"
EXO RecipientTypeDetails:`t "$($exorcp.RecipientTypeDetails)"
EXO WhenMailboxCreated:`t "$($exombx.WhenMailboxCreated)"
"@

                $smsg = "$($FixText)`n$($FixText2)`n$($FixText3)"  ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else { write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                #output the AAD Signon Profile info: $AADSigninSyntax
                $smsg = "$($AADSigninSyntax)"  ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else { write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

            }elseif ( ($rcp.RecipientTypeDetails -match 'Remote(User|Shared|Room|Equipment)Mailbox' ) -AND ($exorcp.RecipientTypeDetails -match '(User|Shared|Room|Equipment)Mailbox') ) {
                # EXO MAILBOX
                $pltTxmc=@{identity=$Mailbox ;erroraction='SilentlyContinue'; } ;
                Try {
                    $mapiTest = ps1TestXMapi @pltTxmc ; 
                    if ($mapiTest.Result -eq 'Success') {
                        $smsg= "(successful Outlook Mbx MAPI validate: Test-MAPIConnectivity -id $($Mailbox) }" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ;
                    $Exit = $Retries ;
                } Catch {
                    $errTrpd=$_ ;
                    $smsg = "Failed processing $($errTrpd.Exception.ItemName). `nError Message: $($errTrpd.Exception.Message)`nError Details: $($errTrpd)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Error } #Error|Warn|Debug
                    else { write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    Start-Sleep -Seconds $RetrySleep ;
                    $Exit ++ ;
                    $smsg = "Failed to exec cmd because: $($Error[0])`nTry #: $($Exit)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Error } #Error|Warn|Debug
                    else { write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    If ($Exit -eq $Retries) { Write-Warning "Unable to exec cmd!" } ;
                }  ;
                $FixText = @"
========================================
User:`t "$($adu.UserPrincipalName)"
Has an *EXO* mailbox:Type:$($exombx.recipienttypedetails) in db:$($exombx.database)
"@ ;
                if ($mapiTest.Result -eq 'Success') {
                $FixText2 = @"
*with *NO* detected issues*
Mailbox Outlook connectivity tested and validated functional...
"@ ;
                } else {
                    $FixText2 = @"
*FAILED Test-exoMAPIConnectivity*
Mailbox Outlook connectivity test failed to connect!...
"@ ;
                } ; 

                $FixText3 = @"

$(($mapiTest|out-string).trim())

The user's o365 LICENSESKUs:                     "$($exolicdetails.LicAccountSkuID)"
With DisplayNames:                               "$($exolicdetails.LicenseFriendlyName)"
The user's o365 Licensing group appears to be:  "$($grantgroup)"
$(if($rcp.recipienttypedetails -ne 'RemoteUserMailbox'){
    '(non-usermailbox, *non*-licensed is typical status)'
})

OnPrem RecipientTypeDetails:`t "$($rcp.RecipientTypeDetails)"
OnPrem WhenCreated:`t "$($rcp.WhenCreated)"
EXO RecipientTypeDetails:`t "$($exorcp.RecipientTypeDetails)"
EXO WhenMailboxCreated:`t "$($exombx.WhenMailboxCreated)"

Mailbox content:
$(($exofldrs | ft -auto $propsmbxfldrs|out-string).trim())

"@


                $smsg = "$($FixText)`n$($FixText2)`n$($FixText3)"  ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else { write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                #output the AAD Signon Profile info: $AADSigninSyntax
                $smsg = "$($AADSigninSyntax)"  ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else { write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

            } else {
                # UNRECOGNIZED RECIPIENTTYPE COMBO
                # ($rcp.RecipientTypeDetails -eq 'RemoteUserMailbox'), ($exorcp.RecipientTypeDetails -eq 'UserMailbox')
                $smsg= "UNRECOGNIZED RECIPIENTTYPE COMBO:`nOPREM:RecipientTypeDetails:$($rcp.RecipientTypeDetails)`nEXO:RecipientTypeDetails:$($exorcp.RecipientTypeDetails)" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Warn } #Error|Warn|Debug
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

            } ; 

            
            if($host.Name -eq "Windows PowerShell ISE Host" -and $host.version.major -lt 5){
                # 11:51 AM 9/22/2020 isev5 supports transcript, anything prior has to fake it
                # 8:46 AM 3/11/2015 shift the logfilename gen out here, so that we can arch it
                #$Logname= (join-path -path (join-path -path $scriptDir -childpath "logs") -childpath ($scriptNameNoExt + "-" + (get-date -uformat "%Y%m%d-%H%M" ) + "-ISEtrans.log")) ;
                # 2:16 PM 4/27/2015 shift to static timestamp $timeStampNow
                #$Logname= (join-path -path (join-path -path $scriptDir -childpath "logs") -childpath ($scriptNameNoExt + "-" + $timeStampNow + "-ISEtrans.log")) ;
                # 2:02 PM 9/21/2018 missing $timestampnow, hardcode
                #$Logname=(join-path -path (join-path -path $scriptDir -childpath "logs") -childpath ($scriptNameNoExt + "-" + (get-date -format 'yyyyMMdd-HHmmtt') + "-ISEtrans.log")) ;
                # maintain-ExoUsrMbxFreeBusyDetails-TOR-ForceAll-Transcript-BATCH-EXEC-20200921-1539PM-trans-log.txt
                $Logname=$transcript.replace('-trans-log.txt','-ISEtrans-log.txt') ;
                write-host "`$Logname: $Logname";
                Start-iseTranscript -logname $Logname  -Verbose:($VerbosePreference -eq 'Continue') ;
                #Archive-Log $Logname ;
                # 1:23 PM 4/23/2015 standardize processing file so that we can send a link to open the transcript for review
                $transcript = $Logname ;
                if($host.version.Major -ge 5){ stop-transcript  -Verbose:($VerbosePreference -eq 'Continue')} # ISE in psV5 actually supports transcription. If you don't stop it, it just keeps rolling
            } else {
                write-verbose "$((get-date).ToString('HH:mm:ss')):Stop Transcript" ;
                Stop-TranscriptLog -Transcript $transcript -verbose:$($VerbosePreference -eq "Continue") ;
            } # if-E
            
            # also trailing echo the log:
            $smsg = "`$logging:`$true:written to:`n$($logfile)" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            $smsg = "$((get-date).ToString('HH:mm:ss')):$($sBnr.replace('=v','=^').replace('v=','^='))" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            
            $logging = $false ;  # reset logging for next pass

        } ; # loop-E $mailboxes
    } ;  # PROC-E
    END {
        # =========== wrap up Tenant connections
        # suppress VerbosePreference:Continue, if set, during mod loads (VERY NOISEY)
        if($VerbosePreference = "Continue"){
            $VerbosePrefPrior = $VerbosePreference ;
            $VerbosePreference = "SilentlyContinue" ;
            $verbose = ($VerbosePreference -eq "Continue") ;
        } ;
        if($script:useEXOv2){
            disconnect-exo2 -verbose:$($verbose) ;
        } else {
            disconnect-exo -verbose:$($verbose) ;
        } ;
        # aad mod *does* support disconnect (msol doesen't!)
        #Disconnect-AzureAD -verbose:$($verbose) ;
        # reenable VerbosePreference:Continue, if set, during mod loads
        if($VerbosePrefPrior -eq "Continue"){
            $VerbosePreference = $VerbosePrefPrior ;
            $verbose = ($VerbosePreference -eq "Continue") ;
        } ;

        $pltCleanup=@{
            LogPath=$tmpcopy
            summarizeStatus=$false ;
            NoTranscriptStop=$true ;
            showDebug=$($showDebug) ;
            whatif=$($whatif) ;
            Verbose = ($VerbosePreference -eq 'Continue') ;
        } ;
        $smsg = "_cleanup():w`n$(($pltCleanup|out-string).trim())" ;
        if($verbose){ if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
        #_cleanup -LogPath $tmpcopy ;
        _cleanup @pltCleanup ;
        # prod is still showing a running unstopped transcript, kill it again
        $stopResults = try {Stop-transcript -ErrorAction stop} catch {} ;

        $smsg = "$($sBnrS.replace('-v','-^').replace('v-','^-'))" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

        # clear the script aliases
        write-verbose "clearing ps1* aliases in Script scope" ; 
        get-alias -scope Script |Where-Object{$_.name -match '^ps1.*'} | ForEach-Object{Remove-Alias -alias $_.name} ;

        write-verbose "(explicit EXIT...)" ;
        Break ;

    } ;  # END-E
}

#*------^ test-xoMailbox.ps1 ^------

#*======^ END FUNCTIONS ^======

Export-ModuleMember -Function check-EXOLegalHold,Connect-ExchangeOnlineTargetedPurge,Connect-EXO,Connect-EXO2,connect-EXO2old,Connect-EXOPSSession,connect-EXOv2RAW,Connect-IPPSSessionTargetedPurge,convert-HistoricalSearchCSV,cxo2cmw,cxo2TOL,cxo2TOR,cxo2VEN,cxoCMW,cxoTOL,cxoTOR,cxoVEN,Disconnect-ExchangeOnline,Disconnect-EXO,Disconnect-EXO2,get-MailboxFolderStats,get-MsgTrace,Get-OrgNameFromUPN,Invoke-ExoOnlineConnection,move-MailboxToXo,check-ReqMods,Print-Details,Reconnect-EXO,Reconnect-EXO2,Reconnect-EXO2old,RemoveExistingEXOPSSession,RemoveExistingPSSessionTargeted,Remove-EXOBrokenClosed,resolve-Name,resolve-user,rxo2CMW,rxo2TOL,rxo2TOR,rxo2VEN,rxoCMW,rxoTOL,rxoTOR,rxoVEN,test-ExoPSession,test-EXOToken,Test-Uri,test-xoMailbox,_cleanup -Alias *


# SIG # Begin signature block
# MIIELgYJKoZIhvcNAQcCoIIEHzCCBBsCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUX+zEg/H5CwLNWxAHQ7F7dIkh
# GdKgggI4MIICNDCCAaGgAwIBAgIQWsnStFUuSIVNR8uhNSlE6TAJBgUrDgMCHQUA
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
# CisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMCMGCSqGSIb3DQEJBDEWBBQxg7X6
# Ls48G37amohM4QGif7/UiTANBgkqhkiG9w0BAQEFAASBgIz50Ierw+3wfQqcvLJ8
# OG9prTLZ6dCpL/5htbSMHHjS1amTArA8yeoRvwGXcIehM5Fgl5EMco7Ilu9XU3PV
# kf/R/xwumqvEhoFQSR2Syw6BTZchlX6QsYRyKGcRWYDhf/cwUmwS2jU4GMkzez2h
# TJ77CV57as5aqrZbk+P223v7
# SIG # End signature block
