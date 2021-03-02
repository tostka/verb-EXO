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
    if($Global:EOLModule){$Global:EOLModule | Remove-Module -Force ; } ;
    if($global:EOLSession){$global:EOLSession | Remove-PSSession ; } ;
    Get-PSSession |Where-Object{$_.ComputerName -match $rgxExoPsHostName } | Remove-PSSession ;
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
