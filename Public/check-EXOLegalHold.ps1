﻿#*------v check-EXOLegalHold.ps1 v------
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
	        ╔░▒▓[LYN-8DCZ1G2]▓▒░░░░▒▓[Thu 11/05/2020 14:38]▓▒░
	        ╚[kadriTSS]::[PS]:C:\u\w\e\scripts$ get-exomailbox Helen.Gotzian@toro.com | FL LitigationHoldEnabled,InPlaceHolds
	        LitigationHoldEnabled : False
	        InPlaceHolds          : {}
	        # expand per arti
	        ╔░▒▓[LYN-8DCZ1G2]▓▒░░░░▒▓[Thu 11/05/2020 14:39]▓▒░
	        ╚[kadriTSS]::[PS]:C:\u\w\e\scripts$ get-exomailbox Helen.Gotzian@toro.com  | Select-Object -ExpandProperty InPlaceHolds
	        # nothing
	 
	        # check for org hold
	        ╔░▒▓[LYN-8DCZ1G2]▓▒░░░░▒▓[Thu 11/05/2020 14:39]▓▒░
	        ╚[kadriTSS]::[PS]:C:\u\w\e\scripts$ Get-exoOrganizationConfig | FL InPlaceHolds
	        InPlaceHolds : {}
	        # expand spec
	        ╔░▒▓[LYN-8DCZ1G2]▓▒░░░░▒▓[Thu 11/05/2020 14:40]▓▒░
	        ╚[kadriTSS]::[PS]:C:\u\w\e\scripts$ Get-exoOrganizationConfig | select -expand InPlaceHolds
	        # nothing
	        # check compliancetaghold (per above)
	        ╔░▒▓[LYN-8DCZ1G2]▓▒░░░░▒▓[Thu 11/05/2020 14:44]▓▒░
	        ╚[kadriTSS]::[PS]:C:\u\w\e\scripts$ get-exomailbox Helen.Gotzian@toro.com  |FL ComplianceTagHoldApplied
	        ComplianceTagHoldApplied : False
	 
	        No holds above.
	 
	        # eDiscovery holds – appears to require the GUID from one of the blank values above.(can't check)
	        If had it, my run on the details would be:
	        connect-ccms ; 
	        $CaseHold = Get-ccCaseHoldPolicy <hold GUID without prefix> ; 
	        Get-ccComplianceCase $CaseHold.CaseId | FL Name ; 
	        $CaseHold | FL Name,ExchangeLocation ; 
	        Get-exoMailboxSearch -InPlaceHoldIdentity <hold GUID> | FL Name,SourceMailboxes
	        # check RetentionCompliancePolicy
	        Get-ccRetentionCompliancePolicy <hold GUID without prefix or suffix> -DistributionDetail  | FL Name,*Location
	 
	        # check compliancetaghold in mbx:
	        ╔░▒▓[LYN-8DCZ1G2]▓▒░░░░▒▓[Thu 11/05/2020 14:44]▓▒░
	        ╚[kadriTSS]::[PS]:C:\u\w\e\scripts$ get-exomailbox Helen.Gotzian@toro.com  |FL ComplianceTagHoldApplied
	        ComplianceTagHoldApplied : False
	 
	        Erm, did anyone *read* the following on holds in the above article?:
	        This appears to be *routine* behavior per section…
	 
		        Managing mailboxes on delay hold  - https://docs.microsoft.com/en-us/microsoft-365/compliance/identify-a-hold-on-an-exchange-online-mailbox?view=o365-worldwide#managing-mailboxes-on-delay-hold
		 
		        After any type of hold is removed from a mailbox, a delay hold is applied. This means that the actual removal of the hold is delayed for 30 days to prevent data from being permanently deleted (purged) from the mailbox. This gives admins an opportunity to search for or recover mailbox items that will be purged after a hold is removed. A delay hold is placed on a mailbox the next time the Managed Folder Assistant processes the mailbox and detects that a hold was removed. Specifically, a delay hold is applied to a mailbox when the Managed Folder Assistant sets one of the following mailbox properties to True:
		                · DelayHoldApplied: This property applies to email-related content (generated by people using Outlook and Outlook on the web) that's stored in a user's mailbox.
		                · DelayReleaseHoldApplied: This property applies to cloud-based content (generated by non-Outlook apps such as Microsoft Teams, Microsoft Forms, and Microsoft Yammer) that's stored in a user's mailbox. Cloud data generated by a Microsoft app is typically stored in a hidden folder in a user's mailbox.
		        When a delay hold is placed on the mailbox (when either of the previous properties is set to True), the mailbox is still considered to be on hold for an unlimited hold duration, as if the mailbox was on Litigation Hold. After 30 days, the delay hold expires, and Microsoft 365 will automatically attempt to remove the delay hold (by setting the DelayHoldApplied or DelayReleaseHoldApplied property to False) so that the hold is removed. After either of these properties are set to False, the corresponding items that are marked for removal are purged the next time the mailbox is processed by the Managed Folder Assistant.
		        To view the values for the DelayHoldApplied and DelayReleaseHoldApplied properties for a mailbox, run the following command in Exchange Online PowerShell.
	 
	        # checking the above:
	        ╔░▒▓[LYN-8DCZ1G2]▓▒░░░░▒▓[Thu 11/05/2020 14:49]▓▒░
	        ╚[kadriTSS]::[PS]:C:\u\w\e\scripts$ get-exomailbox Helen.Gotzian@toro.com  | FL *HoldApplied*
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