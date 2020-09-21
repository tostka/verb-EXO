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