# test-EXODomainDNSRecordsTDO.ps1

<#
.SYNOPSIS
test-EXODomainDNSRecordsTDO - Validate that required Office 365 Exchange Online DNS DNS TXT record exist for specified DomainName
.NOTES
Version     : 0.0.
Author      : Todd Kadrie
Website     : http://www.toddomation.com
Twitter     : @tostka / http://twitter.com/tostka
CreatedDate : 2024-06-07
FileName    : test-EXODomainDNSRecordsTDO
License     : MIT License
Copyright   : (c) 2024 Todd Kadrie
Github      : https://github.com/tostka/verb-AAD
Tags        : Powershell,AzureAD,Authentication,Test
AddedCredit : 
AddedWebsite: 
AddedTwitter: 
REVISIONS
* 3:16 PM 7/11/2024 cleaned up CBH params etc in: test-ExoDnsRecordTDO
* 3:25 PM 6/18/2024 round out to fscript
* 11:07 AM 6/12/2024 init
.DESCRIPTION
test-EXODomainDNSRecordsTDO - Validate that required Office 365 Exchange Online DNS DNS TXT record exist for specified DomainName
Leverages generic sub function test-ExoDnsRecordTDO to run the query & validation

.PARAMETER Ticket
Ticket #[-Ticket 123456]
.PARAMETER DomainName
DomainName to be confirmed[-DomainName somdeomain.tld]
.PARAMETER DomainOwnerValidationString
Domain Ownership Validation String[-DomainOwnerValidationString 'MS=msnnnnnnnn']
.PARAMETER SpfModelDomain
DomainName from which to obtain model SPF string for comparison[-SpfModelDomain somdeomain.tld]
.INPUTS
Does not accept piped input
.OUTPUTS
None (records transcript file)
.EXAMPLE
.\test-EXODomainDNSRecordsTDO.ps1 -Ticket 835841 -DomainName myturf.com -DomainOwnerValidationString 'MS=ms60604949' -Verbose
demo typical pass
.EXAMPLE
PS> test-EXODomainDNSRecordsTDO.ps1 -Ticket 123456 -Domainname somedomain.tld
Typical pass
.LINK
https://bitbucket.org/tostka/powershell/
#>  
##Requires -Modules AzureAD, verb-AAD
[CmdletBinding()]
## PSV3+ whatif support:[CmdletBinding(SupportsShouldProcess)]
###[Alias('Alias','Alias2')]
PARAM(
    [Parameter(Mandatory=$true,HelpMessage="Ticket #[-Ticket 123456]")]
        [string]$Ticket,
    [Parameter(Mandatory=$true,HelpMessage="DomainName to be confirmed[-DomainName somdeomain.tld]")]
        [string]$DomainName,
    [Parameter(Mandatory=$true,HelpMessage="Domain Ownership Validation String[-DomainOwnerValidationString 'MS=msnnnnnnnn']")]
        [string]$DomainOwnerValidationString,
    [Parameter(Mandatory=$false,HelpMessage="DomainName from which to obtain model SPF string for comparison[-SpfModelDomain somdeomain.tld]")]
        [string]$SpfModelDomain = 'myturf.com'
);
#region CONSTANTS_AND_ENVIRO #*======v CONSTANTS_AND_ENVIRO v======
# Debugger:proxy automatic variables that aren't directly accessible when debugging (must be assigned and read back from another vari) ; 
$rPSCmdlet = $PSCmdlet ; 
$rPSScriptRoot = $PSScriptRoot ; 
$rPSCommandPath = $PSCommandPath ; 
$rMyInvocation = $MyInvocation ; 
$rPSBoundParameters = $PSBoundParameters ; 
[array]$score = @() ; 
if($rPSCmdlet.MyInvocation.InvocationName){
    if($rPSCmdlet.MyInvocation.InvocationName -match '\.ps1$'){
        $score+= 'ExternalScript' 
    }elseif($rPSCmdlet.MyInvocation.InvocationName  -match '^\.'){
        write-warning "dot-sourced invocation detected!:$($rPSCmdlet.MyInvocation.InvocationName)`n(will be unable to leverage script path etc from MyInvocation objects)" ; 
        # dot sourcing is implicit scripot exec
        $score+= 'ExternalScript' ; 
    } else {$score+= 'Function' };
} ; 
if($rPSCmdlet.CommandRuntime){
    if($rPSCmdlet.CommandRuntime.tostring() -match '\.ps1$'){$score+= 'ExternalScript' } else {$score+= 'Function' }
} ; 
$score+= $rMyInvocation.MyCommand.commandtype.tostring() ; 
$grpSrc = $score | group-object -NoElement | sort count ;
if( ($grpSrc |  measure | select -expand count) -gt 1){
    write-warning  "$score mixed results:$(($grpSrc| ft -a count,name | out-string).trim())" ;
    if($grpSrc[-1].count -eq $grpSrc[-2].count){
        write-warning "Deadlocked non-majority results!" ;
    } else {
        $runSource = $grpSrc | select -last 1 | select -expand name ;
    } ;
} else {
    write-verbose "consistent results" ;
    $runSource = $grpSrc | select -last 1 | select -expand name ;
};
write-host "Calculated `$runSource:$($runSource)" ;
'score','grpSrc' | get-variable | remove-variable ; # cleanup temp varis

# function self-name (equiv to script's: $MyInvocation.MyCommand.Path) ;
${CmdletName} = $rPSCmdlet.MyInvocation.MyCommand.Name ;
$PSParameters = New-Object -TypeName PSObject -Property $rPSBoundParameters ;
write-verbose "`$rPSBoundParameters:`n$(($rPSBoundParameters|out-string).trim())" ;
$Verbose = ($VerbosePreference -eq 'Continue') ; 
# pre psv2, no $rPSBoundParameters autovari to check, so back them out:
if($rPSCmdlet.MyInvocation.InvocationName){
    if($rPSCmdlet.MyInvocation.InvocationName  -match '^\.'){
        $smsg = "detected dot-sourced invocation: Skipping `$PSCmdlet.MyInvocation.InvocationName-tied cmds..." ; 
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
    } else { 
        write-verbose 'Collect all non-default Params (works back to psv2 w CmdletBinding)'
        $ParamsNonDefault = (Get-Command $rPSCmdlet.MyInvocation.InvocationName).parameters | Select-Object -expand keys | Where-Object{$_ -notmatch '(Verbose|Debug|ErrorAction|WarningAction|ErrorVariable|WarningVariable|OutVariable|OutBuffer)'} ;
    } ; 
} else { 
    $smsg = "(blank `$rPSCmdlet.MyInvocation.InvocationName, skipping Parameters collection)" ; 
    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
} ; 
#region ENVIRO_DISCOVER ; #*------v ENVIRO_DISCOVER v------
<#
# Debugger:proxy automatic variables that aren't directly accessible when debugging ; 
$rPSScriptRoot = $PSScriptRoot ; 
$rPSCommandPath = $PSCommandPath ; 
$rMyInvocation = $MyInvocation ; 
$rPSBoundParameters = $PSBoundParameters ; 
#>
$ScriptDir = $scriptName = '' ;     
if($ScriptDir -eq '' -AND ( (get-variable -name rPSScriptRoot -ea 0) -AND (get-variable -name rPSScriptRoot).value.length)){
    $ScriptDir = $rPSScriptRoot
} ; # populated rPSScriptRoot
if( (get-variable -name rPSCommandPath -ea 0) -AND (get-variable -name rPSCommandPath).value.length){
    $ScriptName = $rPSCommandPath
} ; # populated rPSCommandPath
if($ScriptDir -eq '' -AND $runSource -eq 'ExternalScript'){$ScriptDir = (Split-Path -Path $rMyInvocation.MyCommand.Source -Parent)} # Running from File
# when $runSource:'Function', $rMyInvocation.MyCommand.Source is empty,but on functions also tends to pre-hit from the rPSCommandPath entFile.FullPath ;
if( $scriptname -match '\.psm1$' -AND $runSource -eq 'Function'){
    write-host "MODULE-HOMED FUNCTION:Use `$CmdletName to reference the running function name for transcripts etc (under a .psm1 `$ScriptName will reflect the .psm1 file  fullname)"
    if(-not $CmdletName){write-warning "MODULE-HOMED FUNCTION with BLANK `$CmdletNam:$($CmdletNam)" } ;
} # Running from .psm1 module
if($ScriptDir -eq '' -AND (Test-Path variable:psEditor)) {
    write-verbose "Running from VSCode|VS" ; 
    $ScriptDir = (Split-Path -Path $psEditor.GetEditorContext().CurrentFile.Path -Parent) ; 
        if($ScriptName -eq ''){$ScriptName = $psEditor.GetEditorContext().CurrentFile.Path }; 
} ;
if ($ScriptDir -eq '' -AND $host.version.major -lt 3 -AND $rMyInvocation.MyCommand.Path.length -gt 0){
    $ScriptDir = $rMyInvocation.MyCommand.Path ; 
    write-verbose "(backrev emulating `$rPSScriptRoot, `$rPSCommandPath)"
    $ScriptName = split-path $rMyInvocation.MyCommand.Path -leaf ;
    $rPSScriptRoot = Split-Path $ScriptName -Parent ;
    $rPSCommandPath = $ScriptName ;
} ;
if ($ScriptDir -eq '' -AND $rMyInvocation.MyCommand.Path.length){
    if($ScriptName -eq ''){$ScriptName = $rMyInvocation.MyCommand.Path} ;
    $ScriptDir = $rPSScriptRoot = Split-Path $rMyInvocation.MyCommand.Path -Parent ;
}
if ($ScriptDir -eq ''){throw "UNABLE TO POPULATE SCRIPT PATH, EVEN `$rMyInvocation IS BLANK!" } ;
if($ScriptName){
    if(-not $ScriptDir ){$ScriptDir = Split-Path -Parent $ScriptName} ; 
    $ScriptBaseName = split-path -leaf $ScriptName ;
    $ScriptNameNoExt = [system.io.path]::GetFilenameWithoutExtension($ScriptName) ;
} ; 
# blank $cmdlet name comming through, patch it for Scripts:
if(-not $CmdletName -AND $ScriptBaseName){
    $CmdletName = $ScriptBaseName
}
# last ditch patch the values in if you've got a $ScriptName
if($rPSScriptRoot.Length -ne 0){}else{ 
    if($ScriptName){$rPSScriptRoot = Split-Path $ScriptName -Parent }
    else{ throw "Unpopulated, `$rPSScriptRoot, and no populated `$ScriptName from which to emulate the value!" } ; 
} ; 
if($rPSCommandPath.Length -ne 0){}else{ 
    if($ScriptName){$rPSCommandPath = $ScriptName }
    else{ throw "Unpopulated, `$rPSCommandPath, and no populated `$ScriptName from which to emulate the value!" } ; 
} ; 
if(-not ($ScriptDir -AND $ScriptBaseName -AND $ScriptNameNoExt  -AND $rPSScriptRoot  -AND $rPSCommandPath )){ 
    throw "Invalid Invocation. Blank `$ScriptDir/`$ScriptBaseName/`ScriptNameNoExt" ; 
    BREAK ; 
} ; 
# echo results dyn aligned:
$tv = 'runSource','CmdletName','ScriptName','ScriptBaseName','ScriptNameNoExt','ScriptDir','PSScriptRoot','PSCommandPath','rPSScriptRoot','rPSCommandPath' ; 
$tvmx = ($tv| Measure-Object -Maximum -Property Length).Maximum * -1 ; 
$tv | get-variable | %{  write-verbose  ("`${0,$tvmx} : {1}" -f $_.name,$_.value) } ; 
'tv','tvmx'|get-variable | remove-variable ; # cleanup temp varis
        
#endregion ENVIRO_DISCOVER ; #*------^ END ENVIRO_DISCOVER ^------
# local Constants:

#endregion CONSTANTS_AND_ENVIRO ; #*------^ END CONSTANTS_AND_ENVIRO ^------

#region BANNER ; #*------v BANNER v------
$sBnr="#*======v $(${CmdletName}): v======" ;
$smsg = $sBnr ;
if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level H1 } #Error|Warn|Debug
else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
#endregion BANNER ; #*------^ END BANNER ^------
if($tormeta.OP_ExEgressSubnets){
    $TTCEggr = " ip4:$($tormeta.OP_ExEgressSubnets)" ; 
} else { 
    $smsg = "NO CONFIGURED `$tormeta.OP_ExEgressSubnets! Unable to populate `$TTCEggr!" ; 
    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
    throw $smsg ; 
    break ; 
} ; 
if($cmwmeta.OP_ExEgressSubnets){
    $CMWEggr = " ip4:$($tormeta.OP_ExEgressSubnets)" ; 
} else { 
    $smsg = "NO CONFIGURED `$tormeta.OP_ExEgressSubnets! Unable to populate `$TTCEggr!" ; 
    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
    throw $smsg ; 
    break ; 
} ; 

#*======v FUNCTIONS v======

#*------v Function test-ExoDnsRecordTDO v------
if(-not (get-command test-ExoDnsRecordTDO -ea 0)){
    function test-ExoDnsRecordTDO{
        <#
        .SYNOPSIS
        test-ExoDnsRecordTDO - Boilerplate wrapper for Resolve-DNSName, that runs tests and validates proper returns, against specified testFalue,  Resolve & Validate Mail-related DNS Records (MX, TXT Domain Verific, SPF & DKIM; CNAME autodiscover record) 
        .NOTES
        Version     : 0.0.1
        Author      : Todd Kadrie
        Website     : http://www.toddomation.com
        Twitter     : @tostka / http://twitter.com/tostka
        CreatedDate : 2024-06-07
        FileName    : test-ExoDnsRecordTDO
        License     : MIT License
        Copyright   : (c) 2024 Todd Kadrie
        Github      : https://github.com/tostka/verb-AAD
        Tags        : Powershell,AzureAD,Authentication,Test
        AddedCredit : 
        AddedWebsite: 
        AddedTwitter: 
        REVISIONS
        * 3:16 PM 7/11/2024 cleaned up CBH params etc
        * 3:40 PM 6/18/2024 ren $fltr -> $filter ; $tvalue -> $testValue ; round out into full function; pull sources from Metas; shift into param intputs
        * 6:16 PM 6/12/2024 init
        .DESCRIPTION
        test-ExoDnsRecordTDO - Resolve & Validate Mail-related DNS Records (MX, TXT Domain Verific, SPF & DKIM; CNAME autodiscover record) 
        .PARAMETER Name
        DNS Name[-Name host.domain.tld]
        .PARAMETER Type
        DNS Type[-Type TXT]
        .PARAMETER filter
        TXT type, String property post-filter[-filter '^v=DKIM1']
        .PARAMETER testValue
        Validating value string[-testValue TXT]
        .PARAMETER Server
        DNS Server[-Server 8.8.8.8]
        .INPUTS
        Does not accept piped input
        .OUTPUTS
        System.Object summary
        .EXAMPLE
        write-verbose 'Domain Ownership TXT Validator record test';
        $DomainName = 'somedomain.tld' ; 
        $ret  = test-ExoDnsRecordTDO -Name $domainname -Type TXT -filter '^MS='  -testValue 'MS=msnnnnnnnn' ; 
        if($ret.Validated -eq $true){write-host "Valid MX Record:`n$(($ret.DNSObject|out-string).trim())" } else { write-warning "Failed Validation" } ; 
        Typical pass
        .EXAMPLE
        write-verbose 'mx record test';
        $DomainName = 'somedomain.tld' ; 
        $ret  = test-ExoDnsRecordTDO -Name $domainname -Type MX -testValue 'myturf-com.mail.protection.outlook.com'
        if($ret.Validated -eq $true){write-host "Valid MX Record:`n$(($ret.DNSObject|out-string).trim())" } else { write-warning "Failed Validation" } ; 
        Typical pass
        .EXAMPLE
        write-verbose 'CNAME autodiscover record test';
        $DomainName = 'somedomain.tld' ; 
        $ret  = test-ExoDnsRecordTDO -Name "autodiscover.$($DomainName)" -Type CNAME -testValue 'autodiscover.outlook.com' ; 
        if($ret.Validated -eq $true){write-host "Valid CNAME Record:`n$(($ret.DNSObject|out-string).trim())" } else { write-warning "Failed Validation" } ; 
        Typical pass
        .EXAMPLE
        write-verbose 'TXT SPF DNS record test';
        $DomainName = 'somedomain.tld' ; 
        $ret  = test-ExoDnsRecordTDO -Name $domainname -Type TXT -filter '^v=spf1' -testValue "v=spf1 ip4:148.163.146.158 ip4:148.163.142.153 ip4:170.92.0.0/16 ip4:205.142.232.90 include:spf.protection.outlook.com ~all" ; 
        if($ret.Validated -eq $true){write-host "Valid TXT Record:`n$(($ret.DNSObject|out-string).trim())" } else { write-warning "Failed Validation" } ; 
        Typical pass
        .EXAMPLE
        write-verbose 'CNAME DKIM selector records test (with referred TXT key NameHost resolution)';
        $DomainName = 'somedomain.tld' ; 
        foreach($sel in @('selector1','selector2')){
            $pltTDN=[ordered]@{
                Name = "$($sel)._domainkey.$($DomainName)" ; 
                Type = 'CNAME' ; 
                filter = ''
                testValue = "$($sel)-$($domainname.replace('.','-'))._domainkey.toroco.onmicrosoft.com" ; 
            } ;
            $smsg = "test-ExoDnsRecordTDO w`n$(($pltTDN|out-string).trim())" ; 
            write-host -foregroundcolor gray "$((get-date).ToString('HH:mm:ss')):$($smsg)"  ;
            $ret  = test-ExoDnsRecordTDO @pltTDN ; 
            if($ret.Validated -eq $true){write-host "Valid CNAME DKIM Selector Record:`n$(($ret.DNSObject|out-string).trim())" } else { write-warning "Failed Validation" } ; 
            write-host "--> Re-resolve expanded CNAME target NameHost" 
            $pltTDN=[ordered]@{
                Name = $ret.dnsobject.namehost ; 
                Type = 'TXT' ; 
                filter = '^v=DKIM1' ; # set filter for string match post filter
                testValue = "" ; 
            } ;
            $smsg = "test-ExoDnsRecordTDO w`n$(($pltTDN|out-string).trim())" ; 
            write-host -foregroundcolor gray "$((get-date).ToString('HH:mm:ss')):$($smsg)"  ;
            $ret  = test-ExoDnsRecordTDO @pltTDN ; 
            if($ret.Validated -eq $true){write-host "Valid TXT target DKIM key Record:`n$(($ret.DNSObject|out-string).trim())" } else { write-warning "Failed Validation" } ; 
        } ; 
        Typical pass
        .EXAMPLE
        write-verbose 'CNAME DKIM selector records test (with referred TXT key NameHost resolution)';
        $DomainName = 'somedomain.tld' ; 
        foreach($sel in @('dkim1','dkim2')){
            $pltTDN=[ordered]@{
                Name = "$($sel)._domainkey.$($DomainName)" ; 
                Type = 'CNAME' ; 
                filter = ''
                testValue = "$($sel)-$($domainname.replace('.','-'))._domainkey.toroco.onmicrosoft.com" ; 
            } ;
            $smsg = "test-ExoDnsRecordTDO w`n$(($pltTDN|out-string).trim())" ; 
            write-host -foregroundcolor gray "$((get-date).ToString('HH:mm:ss')):$($smsg)"  ;
            $ret  = test-ExoDnsRecordTDO @pltTDN ; 
            if($ret.Validated -eq $true){write-host "Valid CNAME DKIM Selector Record:`n$(($ret.DNSObject|out-string).trim())" } else { write-warning "Failed Validation" } ; 
            write-host "--> Re-resolve expanded CNAME target NameHost" 
            $pltTDN=[ordered]@{
                Name = $ret.dnsobject.namehost ; 
                Type = 'TXT' ; 
                filter = '^v=DKIM1' ; # set filter for string match post filter
                testValue = "" ; 
            } ;
            $smsg = "test-ExoDnsRecordTDO w`n$(($pltTDN|out-string).trim())" ; 
            write-host -foregroundcolor gray "$((get-date).ToString('HH:mm:ss')):$($smsg)"  ;
            $ret  = test-ExoDnsRecordTDO @pltTDN ; 
            if($ret.Validated -eq $true){write-host "Valid TXT target DKIM key Record:`n$(($ret.DNSObject|out-string).trim())" } else { write-warning "Failed Validation" } ; 
        } ; 
        Typical pass using non-o365 custom selictor specs
    .EXAMPLE
    $DomainName = 'somedomain.tld' ;
    $sBnr3="`n#*~~~~~~v CHECK:$($DomainName): Domain Ownership Verification 'TXT' DNS record v~~~~~~" ;
    write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($sBnr3)" ;
    $pltTDN=[ordered]@{
        Name = $DomainName ;
        Type = 'TXT' ;
        filter = '^MS=' ;
        testValue = MS=msnnnnnnnn
    } ;
    $smsg = "test-ExoDnsRecordTDO w`n$(($pltTDN|out-string).trim())" ;
    write-host -foregroundcolor gray "$((get-date).ToString('HH:mm:ss')):$($smsg)"  ;
    $ret  = test-ExoDnsRecordTDO @pltTDN ;
    if($ret.Validated -eq $true){write-host "Valid TXT DomAuth Record:`n$(($ret.DNSObject|out-string).trim())" } else { write-warning "Failed Validation" } ;
    write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($sBnr3.replace('~v','~^').replace('v~','^~'))`n" ;
    $sBnr3="`n#*~~~~~~v CHECK:$($DomainName): Domain 'MX' DNS record v~~~~~~" ;
    write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($sBnr3)" ;
    $pltTDN=[ordered]@{
        Name = $DomainName ;
        Type =  'MX' ;
        filter = '' ; # set filter for string match post filter
        testValue = 'myturf-com.mail.protection.outlook.com'
    } ;
    $smsg = "test-ExoDnsRecordTDO w`n$(($pltTDN|out-string).trim())" ;
    write-host -foregroundcolor gray "$((get-date).ToString('HH:mm:ss')):$($smsg)"  ;
    $ret  = test-ExoDnsRecordTDO @pltTDN ;
    if($ret.Validated -eq $true){write-host "Valid MX Record:`n$(($ret.DNSObject|out-string).trim())" } else { write-warning "Failed Validation" } ;
    write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($sBnr3.replace('~v','~^').replace('v~','^~'))`n" ;
    $sBnr3="`n#*~~~~~~v CHECK:$($DomainName): Autodiscover CNAME DNS record v~~~~~~" ;
    write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($sBnr3)" ;
    $pltTDN=[ordered]@{
        Name = "autodiscover.$($DomainName)" ;
        Type = 'CNAME' ;
        filter = '' ; # set filter for string match post filter
        testValue = 'autodiscover.outlook.com' ;
    } ;
    $smsg = "test-ExoDnsRecordTDO w`n$(($pltTDN|out-string).trim())" ;
    write-host -foregroundcolor gray "$((get-date).ToString('HH:mm:ss')):$($smsg)"  ;
    $ret  = test-ExoDnsRecordTDO @pltTDN ;
    if($ret.Validated -eq $true){write-host "Valid CNAME Autodiscover Record:`n$(($ret.DNSObject|out-string).trim())" } else { write-warning "Failed Validation" } ;
    write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($sBnr3.replace('~v','~^').replace('v~','^~'))`n" ;
    $sBnr3="`n#*~~~~~~v CHECK:$($DomainName): Domain SPF 'TXT' DNS record v~~~~~~" ;
    write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($sBnr3)" ;
    $pltTDN=[ordered]@{
        Name = $DomainName ;
        Type = 'TXT' ;
        filter = '^v=spf1' ; # set filter for string match post filter
        testValue = "v=spf1 ip4:148.163.146.158 ip4:148.163.142.153 ip4:170.92.0.0/16 ip4:205.142.232.90 include:spf.protection.outlook.com ~all" ;
    } ;
    $smsg = "test-ExoDnsRecordTDO w`n$(($pltTDN|out-string).trim())" ;
    write-host -foregroundcolor gray "$((get-date).ToString('HH:mm:ss')):$($smsg)"  ;
    $ret  = test-ExoDnsRecordTDO @pltTDN ;
    if($ret.Validated -eq $true){write-host "Valid TXT SPF Record:`n$(($ret.DNSObject|out-string).trim())" } else { write-warning "Failed Validation" } ;
    write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($sBnr3.replace('~v','~^').replace('v~','^~'))`n" ;
    $sBnr3="`n#*~~~~~~v CHECK:$($DomainName):'CNAME' DKIM DNS records v~~~~~~" ;
    write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($sBnr3)" ;
    foreach($sel in @('selector1','selector2')){
        $pltTDN=[ordered]@{
            Name = "$($sel)._domainkey.$($DomainName)" ;
            Type = 'CNAME' ;
            filter = ''
            testValue = "$($sel)-$($domainname.replace('.','-'))._domainkey.toroco.onmicrosoft.com" ;
        } ;
        $smsg = "test-ExoDnsRecordTDO w`n$(($pltTDN|out-string).trim())" ;
        write-host -foregroundcolor gray "$((get-date).ToString('HH:mm:ss')):$($smsg)"  ;
        $ret  = test-ExoDnsRecordTDO @pltTDN ;
        if($ret.Validated -eq $true){write-host "Valid CNAME DKIM Selector Record:`n$(($ret.DNSObject|out-string).trim())" } else { write-warning "Failed Validation" } ;
        write-host "--> Re-resolve expanded CNAME target NameHost"
        $pltTDN=[ordered]@{
            Name = $ret.dnsobject.namehost ;
            Type = 'TXT' ;
            filter = '^v=DKIM1' ; # set filter for string match post filter
            testValue = "" ;
        } ;
        $smsg = "test-ExoDnsRecordTDO w`n$(($pltTDN|out-string).trim())" ;
        write-host -foregroundcolor gray "$((get-date).ToString('HH:mm:ss')):$($smsg)"  ;
        $ret  = test-ExoDnsRecordTDO @pltTDN ;
        if($ret.Validated -eq $true){write-host "Valid TXT target DKIM key Record:`n$(($ret.DNSObject|out-string).trim())" } else { write-warning "Failed Validation" } ;
    } ;
    write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($sBnr3.replace('~v','~^').replace('v~','^~'))`n" ;
    Big all inclusive test of the range of DNS records required for O365 EXO domains.
        .LINK
        https://bitbucket.org/tostka/powershell/
        #>  
        [CmdletBinding()]
        Param(
            [Parameter(Mandatory=$True,HelpMessage="DNS Name[-Name host.domain.tld]")]
            [string]$Name,
            [Parameter(Mandatory=$True,HelpMessage="DNS Type[-Type TXT]")]
                [ValidateSet('MX','TXT','CNAME')]
                [string]$Type,
            [Parameter(HelpMessage="TXT type, String property post-filter[-filter '^v=DKIM1']")]
                [string]$filter,
            [Parameter(HelpMessage="Validating value string[-testValue TXT]")]
                [string]$testValue,
            [Parameter(Mandatory=$false,HelpMessage="DNS Server[-Server 8.8.8.8]")]
                [string]$Server = '1.1.1.1'
        ) ; 
        $pltRvDN=[ordered]@{
            Name = $Name ; 
            Server = $Server ; 
            Type = $type ; 
            erroraction = 'STOP' ;
        } ;
        $oReturn = [ordered]@{
            DNSObject = $null ; 
            Type = $Type ; 
            Validated = $false ; 
        } ; 
        if($testValue){$oReturn.add('testValue',$testValue)} ; 
        if($filter){$oReturn.add('filter',$filter)} ; 
        $smsg = "resolve-DNSName w`n$(($pltRvDN|out-string).trim())" ; 
        write-host -foregroundcolor gray "$((get-date).ToString('HH:mm:ss')):$($smsg)" ;
        TRY{
            if($rec = resolve-DNSName @pltRvDN ){
                switch($pltRvDN.Type){
                    'MX' {
                        if($rec.NameExchange -eq $testValue ){
                            write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):==> NameExchange value matches spec: $($testValue)" ; 
                            $oReturn.DNSObject = $rec ; 
                            $oReturn.Validated = $true ;
                            [pscustomobject]$oReturn | write-output  ; 
                        } else{
                            write-warning "$((get-date).ToString('HH:mm:ss')):String value DOES NOT MATCH MS specified validator!: $($testValue)" ; 
                            $oReturn.DNSObject = $rec ; 
                            $oReturn.Validated = $false ;
                            [pscustomobject]$oReturn | write-output  ; 
                        }; 
                    } ; 
                    'TXT' {
                        $oReturn.add('strings',$null)
                        if($filter -AND ($rec| ? strings -match $filter)){
                            $rec = $rec| ? strings -match $filter ; 
                        } ; 
                        write-host -foregroundcolor green "`n$(($rec | ft -a |out-string).trim())" ; 
                        if($testValue -AND (($rec| select -expand strings) -eq $testValue)){
                            write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):==> String value matches MS specified validator: $($testValue)" ; 
                            $oReturn.DNSObject = $rec ; 
                            $oReturn.Strings = ($rec| select -expand strings) ; 
                            $oReturn.Validated = $true ;
                            [pscustomobject]$oReturn | write-output  ; 
                        }elseif($testValue -AND (($rec| select -expand strings) -ne $testValue)){
                            write-warning "$((get-date).ToString('HH:mm:ss')):String value DOES NOT MATCH MS specified validator!: $($testValue)" ; 
                            $oReturn.DNSObject = $rec ; 
                            $oReturn.Strings = ($rec| select -expand strings) ; 
                            $oReturn.Validated = $false ;
                            [pscustomobject]$oReturn | write-output  ; 
                        }elseif($filter -AND $rec){
                            write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):No comparison `$testValue, but matched `$filter:$($filter): ==> String value matches specification" ; 
                            $oReturn.DNSObject = $rec ; 
                            $oReturn.Strings = ($rec| select -expand strings) ; 
                            $oReturn.Validated = $true ;
                            [pscustomobject]$oReturn | write-output  ; 
                        } else{
                            write-warning "$((get-date).ToString('HH:mm:ss')):String value DOES NOT MATCH MS specified validator!: $($testValue)" ; 
                            $oReturn.DNSObject = $rec ; 
                            $oReturn.Strings = ($rec| select -expand strings) ; 
                            $oReturn.Validated = $false ;
                            [pscustomobject]$oReturn | write-output  ; 
                        }; 
                    } 
                    'CNAME' {
                        if($rec.NameHost -eq $testValue ){
                            write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):NameHost value matches spec: $($testValue)" ; 
                            $oReturn.DNSObject = $rec ; 
                            $oReturn.Validated = $true ;
                            [pscustomobject]$oReturn | write-output  ;
                        } else{
                            write-warning "$((get-date).ToString('HH:mm:ss')):value DOES NOT MATCH specification!: $($testValue)" ; 
                            $oReturn.DNSObject = $rec ; 
                            $oReturn.Validated = $false ;
                            [pscustomobject]$oReturn | write-output  ;
                        }; 
                    }
                    default {
                        $smsg = "unrecognized Type: $($type) !" ; 
                        write-warning $smsg ; 
                        $oReturn.DNSObject = $rec ; 
                        $oReturn.Validated = $false ;
                        [pscustomobject]$oReturn | write-output  ;
                        throw $smsg ; 
                    } ;
                } 
            }else{
                write-warning "Unable to resolve-DNSName w`n$(($pltRvDN|out-string).trim())!" ; 
            } ;
        } CATCH {
            $ErrTrapd=$Error[0] ;
            $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
            write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" ;
            $oReturn.DNSObject = $rec ; 
            $oReturn.Validated = $false ;
            [pscustomobject]$oReturn | write-output  ;
            #throw $smsg ;
            #Continue
        } ; 
    } ; 
} ; 
#*------^ END Function test-ExoDnsRecordTDO ^------




#*======^ END FUNCTIONS ^======
$transcript = ".\logs\$($Ticket)-$($DomainName)-test-EXODomainDNSRecordsTDO-$(get-date -format 'yyyyMMdd-HHmmtt')-trans-log.txt" ; 
$stopResults = try {Stop-transcript -ErrorAction stop} catch {} ;
if($stopResults){
    $smsg = "Stop-transcript:$($stopResults)" ; 
    write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)"  ; 
} ; 
$startResults = start-Transcript -path $transcript ;
if($startResults){
    $smsg = "start-transcript:$($startResults)" ; 
    write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" ;
} ; 

$sBnr3="`n#*~~~~~~v CHECK:$($DomainName): Domain Ownership Verification 'TXT' DNS record v~~~~~~" ;
write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($sBnr3)" ;
$pltTDN=[ordered]@{
    Name = $DomainName ;
    Type = 'TXT' ;
    filter = '^MS=' ;
    testValue = $DomainOwnerValidationString  ; 
} ;
$smsg = "test-ExoDnsRecordTDO w`n$(($pltTDN|out-string).trim())" ;
write-host -foregroundcolor gray "$((get-date).ToString('HH:mm:ss')):$($smsg)"  ;
$ret  = test-ExoDnsRecordTDO @pltTDN ;
if($ret.Validated -eq $true){write-host "Valid TXT DomAuth Record:`n$(($ret.DNSObject|out-string).trim())" } else { write-warning "Failed Validation" } ;
write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($sBnr3.replace('~v','~^').replace('v~','^~'))`n" ;

$sBnr3="`n#*~~~~~~v CHECK:$($DomainName): Domain 'MX' DNS record v~~~~~~" ;
write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($sBnr3)" ;
$pltTDN=[ordered]@{
    Name = $DomainName ;
    Type =  'MX' ;
    filter = '' ; # set filter for string match post filter
    testValue = "$($domainname.replace('.','-')).mail.protection.outlook.com"
} ;
$smsg = "test-ExoDnsRecordTDO w`n$(($pltTDN|out-string).trim())" ;
write-host -foregroundcolor gray "$((get-date).ToString('HH:mm:ss')):$($smsg)"  ;
$ret  = test-ExoDnsRecordTDO @pltTDN ;
if($ret.Validated -eq $true){write-host "Valid MX Record:`n$(($ret.DNSObject|out-string).trim())" } else { write-warning "Failed Validation" } ;
write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($sBnr3.replace('~v','~^').replace('v~','^~'))`n" ;

$sBnr3="`n#*~~~~~~v CHECK:$($DomainName): Autodiscover CNAME DNS record v~~~~~~" ;
write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($sBnr3)" ;
$pltTDN=[ordered]@{
    Name = "autodiscover.$($DomainName)" ;
    Type = 'CNAME' ;
    filter = '' ; # set filter for string match post filter
    testValue = 'autodiscover.outlook.com' ;
} ;
$smsg = "test-ExoDnsRecordTDO w`n$(($pltTDN|out-string).trim())" ;
write-host -foregroundcolor gray "$((get-date).ToString('HH:mm:ss')):$($smsg)"  ;
$ret  = test-ExoDnsRecordTDO @pltTDN ;
if($ret.Validated -eq $true){write-host "Valid CNAME Autodiscover Record:`n$(($ret.DNSObject|out-string).trim())" } else { write-warning "Failed Validation" } ;
write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($sBnr3.replace('~v','~^').replace('v~','^~'))`n" ;

$sBnr3="`n#*~~~~~~v CHECK:$($DomainName): Domain SPF 'TXT' DNS record v~~~~~~" ;
write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($sBnr3)" ;
write-host "Basing SPF Strings test on $($SpfModelDomain) core SPF" 
if($res = resolve-dnsname -Name $SpfModelDomain  -Type TXT -Server 1.1.1.1 |? strings -match 'v=spf1'){
    $spf = $res.strings ;
    write-verbose "sub out quotes; split on spaces; join the 1st & ip4 records" ;
    #$spfTXT = (($spf -replace '"','' ) -split ' ' | ?{$_ -match '(v=spf1|ip4)'} ) -join ' ' ;
    $spfElements = ($spf -replace '"','' ) -split ' ' ; 
    $spfTXT = @() ; 
    write-verbose "include version spec & ip4|ip6 records" ; 
    $spfTXT += $spfElements | ?{$_ -match '(v=spf1|ip4|ip6)'} ; 
    write-verbose "include the microsoft SPF"
    $spfTXT += $spfElements | ?{$_ -match 'spf\.protection\.outlook\.com'} 
    write-verbose "include the all handling preference"
    $spfTXT += $spfElements | ?{$_ -match '(-|~)all'}
    write-verbose "join the elements with spaces & collapse to a string" ; 
    [string]$spfTXT = $spfTXT -join ' ' ; 
    # issues we need to include CMW & TTC, and they aren't on any single record anymore, so qry them out of $xxxmeta
    <# force the CMW egress in if the spec is on cidr range
    #$TTCEggr 
    #$CMWEggr 
    # can join the meta's: (@($tormeta.OP_ExEgressIPs + $cmwmeta.OP_ExEgressIPs) |  select -unique  | %{ " ip4:$($_)" }) -join ""  ; 
    # if the model includes TTCEggr ip4, but doesn't include CMWEggr ip4, append it
    if( ($spfTXT -match [regex]::Escape($TTCEggr)) -AND ($spfTXT -notmatch [regex]::Escape($CMWEggr)) ){
        $spfTXT += $CMWEggr ; 
    } ; 
    write-verbose "include:default EOP" ; 
    $spfTXT += " $(($spf -replace '"','' ) -split ' ' |?{$_ -match 'include:spf\.protection\.outlook\.com'})" ; 
    write-verbose "add trailing all directive" ; 
    $spfTXT +=  " $(($spf -replace '"','' ) -split ' ' | select -last 1)" ; 
    #>
    <#
$SPFRecord = @"
#-=-=-=-=-=-=-=-=
ZoneName : $($DomainName)
Hostname : @
Type     : TXT
Value    : "$($spfTXT)"
TTL      : 3600
#-=-=-=-=-=-=-=-=
"@ ; 
#>
    #$SpfModelDomain = 'myturf.com' ; 
    #$DomainName = 'bossplow.com' ; 
    write-host "Separately validating key elements in $($SpfModelDomain) model spf are present (in case of ordering issues):" ; 
    $modelspf = resolve-dnsname -server 1.1.1.1 -type txt -name myturf.com | ? strings -match 'spf' ; 
    $spfitems = $modelspf.strings -split ' ' ; 
    $tspfrec = resolve-dnsname -server 1.1.1.1 -type txt -name $domainname | ? strings -match 'spf' ; 
    $cspf = $tspfrec.strings -split ' ' ; 
    foreach($item in $spfitems){
        if($cspf | ?{$_ -match ([regex]::escape($item))} ){
            write-host "$item present" ; 
        }else{write-warning "$item missing"}  ; 
    } ;     
} else { 
    $smsg = "UNABLE TO: resolve-dnsname -Name $($SpfModelDomain) -Type TXT -Server 1.1.1.1 !" ;
    write-warning $smsg ; 
    throw $smsg ; 
    break ; 
} ; 
$pltTDN=[ordered]@{
    Name = $DomainName ;
    Type = 'TXT' ;
    filter = '^v=spf1' ; # set filter for string match post filter
    testValue = $spfTXT ; 
    #$TXTSpfValue  ;
} ;
$smsg = "test-ExoDnsRecordTDO w`n$(($pltTDN|out-string).trim())" ;
write-host -foregroundcolor gray "$((get-date).ToString('HH:mm:ss')):$($smsg)"  ;
$ret  = test-ExoDnsRecordTDO @pltTDN ;
if($ret.Validated -eq $true){write-host "Valid TXT SPF Record:`n$(($ret.DNSObject|out-string).trim())" } else { write-warning "Failed Validation" } ;
write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($sBnr3.replace('~v','~^').replace('v~','^~'))`n" ;
$sBnr3="`n#*~~~~~~v CHECK:$($DomainName):'CNAME' DKIM DNS records v~~~~~~" ;
write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($sBnr3)" ;
foreach($sel in @('selector1','selector2')){
    $pltTDN=[ordered]@{
        Name = "$($sel)._domainkey.$($DomainName)" ;
        Type = 'CNAME' ;
        filter = ''
        testValue = "$($sel)-$($domainname.replace('.','-'))._domainkey.toroco.onmicrosoft.com" ;
    } ;
    $smsg = "test-ExoDnsRecordTDO w`n$(($pltTDN|out-string).trim())" ;
    write-host -foregroundcolor gray "$((get-date).ToString('HH:mm:ss')):$($smsg)"  ;
    $ret  = test-ExoDnsRecordTDO @pltTDN ;
    if($ret.Validated -eq $true){write-host "Valid CNAME DKIM Selector Record:`n$(($ret.DNSObject|out-string).trim())" } else { write-warning "Failed Validation" } ;
    write-host "--> Re-resolve expanded CNAME target NameHost"
    $pltTDN=[ordered]@{
        Name = $ret.dnsobject.namehost ;
        Type = 'TXT' ;
        filter = '^v=DKIM1' ; # set filter for string match post filter
        testValue = "" ;
    } ;
    $smsg = "test-ExoDnsRecordTDO w`n$(($pltTDN|out-string).trim())" ;
    write-host -foregroundcolor gray "$((get-date).ToString('HH:mm:ss')):$($smsg)"  ;
    $ret  = test-ExoDnsRecordTDO @pltTDN ;
    if($ret.Validated -eq $true){write-host "Valid TXT target DKIM key Record:`n$(($ret.DNSObject|out-string).trim())" } else {
        $smsg = "Failed Validation target DKIM key TXT record" 
        $smsg += "`n$($pltTDN.name)" ; 
        write-warning $smsg ; 
    } ;
} ;
$smsg = "$($sBnr.replace('=v','=^').replace('v=','^='))" ;
if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level H1 } #Error|Warn|Debug
else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
$stopResults = try {Stop-transcript -ErrorAction stop} catch {} ;
if($stopResults){
    $smsg = "Stop-transcript:$($stopResults)" ; 
    # Opt:verbose
    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
    # # Opt:pswlt
    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
    #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
} ; 