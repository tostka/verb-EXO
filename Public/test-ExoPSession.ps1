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