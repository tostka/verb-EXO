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
    
    #Disconnect-ExchangeOnline -confirm:$false ; 
    # just use the updated RemoveExistingEXOPSSession
    RemoveExistingEXOPSSession
    
    Disconnect-PssBroken -verbose:$($verbose) ;
    Remove-PSTitlebar 'EXO' ;
}

#*------^ Disconnect-EXO2.ps1 ^------