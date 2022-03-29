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
    * 3:03 PM 3/29/2022 rewrote to reflect current specs in v2.0.5 of ExchangeOnlineManagement:Disconnect-ExchangeOnlineManagement cmds
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
    
    # lifted from connect-EXO2 
    # .dll etc loads, from connect-exchangeonline: (should be installed with the above)
    if (-not($ExchangeOnlineMgmtPath)) {
        $EOMgmtModulePath = split-path (get-module ExchangeOnlineManagement -list).Path ;
        if($IsCoreCLR){
            $EOMgmtModulePath = resolve-path -Path $EOMgmtModulePath\netcore ;
            $smsg = "(.netcore path in use:" ; 
        } else { 
            $EOMgmtModulePath = resolve-path -Path $EOMgmtModulePath\netFramework
            $smsg = "(.netnetFramework path in use:" ;                 
        } ; 
        $smsg += "$($EOMgmtModulePath))" ; 
        if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
    } ;
    #the verb-*token functions are in here    
    if (-not $ExoPowershellGalleryModule) { $ExoPowershellGalleryModule = "Microsoft.Exchange.Management.ExoPowershellGalleryModule.dll" } ;
    if (-not $ExoPowershellGalleryModulePath) {
        $ExoPowershellGalleryModulePath = join-path -path $EOMgmtModulePath -childpath $ExoPowershellGalleryModule ;
        if(-not (test-path $ExoPowershellGalleryModulePath)){
            $smsg = "UNABLE TO test-path `$ExoPowershellGalleryModulePath!:`n$($ExoPowershellGalleryModulePathz)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN }
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            THROW $smsg
            BREAK ;
        } ;
    } ;
    # full path: C:\Users\USER\Documents\WindowsPowerShell\Modules\ExchangeOnlineManagement\1.0.1\Microsoft.Exchange.Management.ExoPowershellGalleryModule.dll
    # Name: Microsoft.Exchange.Management.ExoPowershellGalleryModule
    if (-not(get-module $ExoPowershellGalleryModule.replace('.dll','') )) {
        Import-Module $ExoPowershellGalleryModulePath -Verbose:$false ;
    } ;    
    
    # confirm module present
    $modname = 'ExchangeOnlineManagement' ; 
    #Try {Get-Module $modname -listavailable -ErrorAction Stop | out-null } Catch {Install-Module $modname -scope CurrentUser ; } ;                 # installed
    Try {Get-Module $modname -ErrorAction Stop | out-null } Catch {Import-Module -Name $modname -MinimumVersion '1.0.1' -ErrorAction Stop -verbose:$false; } ; # imported
    
    $existingPSSession = Get-PSSession | Where-Object {$_.ConfigurationName -like "Microsoft.Exchange" -and $_.Name -like "ExchangeOnlineInternalSession*"} ;
    if ($existingPSSession.count -gt 0) {
        for ($index = 0; $index -lt $existingPSSession.count; $index++){
            $session = $existingPSSession[$index]
            Remove-PSSession -session $session
            Write-Host "Removed the PSSession $($session.Name) connected to $($session.ComputerName)"
            # Remove any active access token from the cache
            Clear-ActiveToken -TokenProvider $session.TokenProvider
            # Remove any previous modules loaded because of the current PSSession
            if ($session.PreviousModuleName -ne $null){
                if ((Get-Module $session.PreviousModuleName).Count -ne 0){
                    Remove-Module -Name $session.PreviousModuleName -ErrorAction SilentlyContinue
                }
                $session.PreviousModuleName = $null
            } ; 
            # Remove any leaked module in case of removal of broken session object
            if ($session.CurrentModuleName -ne $null){
                if ((Get-Module $session.CurrentModuleName).Count -ne 0){
                    Remove-Module -Name $session.CurrentModuleName -ErrorAction SilentlyContinue ; 
                } ;  
            }  ; 
        } ;  # loop-E
    } ; # if-E $existingPSSession.count -gt 0
    
    # just alias disconnect-ExchangeOnline, it retires token etc as well as closing PSS, but biggest reason is it's got a confirm, hard-coded, needs a function to override
    
    #Disconnect-ExchangeOnline -confirm:$false ; 
    # just use the updated RemoveExistingEXOPSSession
    #RemoveExistingEXOPSSession -Verbose:$false ;
    # v2.0.5 3:01 PM 3/29/2022 no longer exists
    
    Disconnect-PssBroken -verbose:$false ;
    Remove-PSTitlebar 'EXO2' #-verbose:$($VerbosePreference -eq "Continue");
    #[console]::ResetColor()  # reset console colorscheme
} ; 

#*------^ Disconnect-EXO2.ps1 ^------