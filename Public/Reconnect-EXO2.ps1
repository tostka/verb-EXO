﻿#*------v Reconnect-EXO2.ps1 v------
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
    [Alias('rxo')]
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
        if( $legPSSession = Get-PSSession | where-object {$_.ConfigurationName -like "Microsoft.Exchange" -and $_.Name -match "(Session|WinRM)\d*" } ){
            # ignore state & Avail, close the conflicting legacy conn's
            Disconnect-Exo ; Disconnect-PssBroken ;Start-Sleep -Seconds 3;
            $bExistingEXOGood = $false ; 
        } ; 
        # if we're using EXOv1-style BasicAuth, clear incompatible existing EXOv2 PSS's
        $exov2Broken = Get-PSSession | where-object {$_.ConfigurationName -like "Microsoft.Exchange" -and $_.Name -eq "ExchangeOnlineInternalSession*" -and $_.State -like "*Broken*"}
        $exov2Closed = Get-PSSession | where-object {$_.ConfigurationName -like "Microsoft.Exchange" -and $_.Name -eq "ExchangeOnlineInternalSession*" -and $_.State -like "*Closed*"}
        
        if ($exov2Broken.count -gt 0){for ($index = 0 ;$index -lt $psBroken.count ;$index++){Remove-PSSession -session $psBroken[$index]} };
        if ($exov2Closed.count -gt 0){for ($index = 0 ;$index -lt $psClosed.count ; $index++){Remove-PSSession -session $psClosed[$index] } } ; 
    
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
        if(Get-PSSession | where-object {$_.ConfigurationName -like "Microsoft.Exchange" -and $_.Name -like "ExchangeOnlineInternalSession*" -and $_.State -like "*Opened*" -AND ($_.Availability -eq 'Available')} ){ 
            if ((Get-xoAcceptedDomain).domainname.contains($Credential.username.split('@')[1].tostring())){
                # validate that the connected EXO is to the $Credential tenant    
                write-verbose "(EXOv2 Authenticated & Functional:$($Credential.username.split('@')[1].tostring()))" ; 
                $bExistingEXOGood = $true ; 
            } else { 
                write-error "(Credential mismatch:disconnecting from existing EXO:$($eEXO.Identity) tenant)" ; 
                Disconnect-exo2 ; 
                $bExistingEXOGood = $false ; 
            } ;
        } ; 
    }  # END-E 
}
#*------^ Reconnect-EXO2.ps1 ^------