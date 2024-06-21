# test-ExoDnsRecordTDO_func.ps1

#*------v Function test-ExoDnsRecordTDO v------
function test-ExoDnsRecordTDO{
    <#
    .SYNOPSIS
    test-ExoDnsRecordTDO - Resolve & Validate Mail-related DNS Records (MX, TXT Domain Verific & SPF & DKIM; CNAME) 
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
    * 6:16 PM 6/12/2024 init
    .DESCRIPTION
    test-ExoDnsRecordTDO - Resolve & Validate Mail-related DNS Records (MX, TXT Domain Verific & SPF & DKIM; CNAME) 
    .PARAMETER Name
    DNS Name[-Type fqdn.somedomain.tld]
    .PARAMETER Type
    DNS Type (MX|CNAME|TXT)[-Type TXT]
    .PARAMETER fltr
    String Type post-filter[-fltr TXT]
    .PARAMETER tvalue
    Validating value string[-tvalue TXT]
    .INPUTS
    Does not accept piped input
    .OUTPUTS
    None (records transcript file)
    .EXAMPLE
    write-verbose 'Domain Ownership TXT Validator record test';
    $DomainName = 'myturf.com' ; 
    $ret  = test-ExoDnsRecordTDO -Name $domainname -Type TXT -fltr '^MS='  -tvalue 'MS=ms60604949' ; 
    if($ret.Validated -eq $true){write-host "Valid MX Record:`n$(($ret.DNSObject|out-string).trim())" } else { write-warning "Failed Validation" } ; 
    Typical pass
    .EXAMPLE
    write-verbose 'mx record test';
    $DomainName = 'myturf.com' ; 
    $ret  = test-ExoDnsRecordTDO -Name $domainname -Type MX -tvalue 'myturf-com.mail.protection.outlook.com'
    if($ret.Validated -eq $true){write-host "Valid MX Record:`n$(($ret.DNSObject|out-string).trim())" } else { write-warning "Failed Validation" } ; 
    Typical pass
    .EXAMPLE
    write-verbose 'CNAME autodiscover record test';
    $DomainName = 'myturf.com' ; 
    $ret  = test-ExoDnsRecordTDO -Name "autodiscover.$($DomainName)" -Type CNAME -tvalue 'autodiscover.outlook.com' ; 
    if($ret.Validated -eq $true){write-host "Valid CNAME Record:`n$(($ret.DNSObject|out-string).trim())" } else { write-warning "Failed Validation" } ; 
    Typical pass
    .EXAMPLE
    write-verbose 'TXT SPF DNS record test';
    $DomainName = 'myturf.com' ; 
    $ret  = test-ExoDnsRecordTDO -Name $domainname -Type TXT -fltr '^v=spf1' -tvalue "v=spf1 ip4:148.163.146.158 ip4:148.163.142.153 ip4:170.92.0.0/16 ip4:205.142.232.90 include:spf.protection.outlook.com ~all" ; 
    if($ret.Validated -eq $true){write-host "Valid TXT Record:`n$(($ret.DNSObject|out-string).trim())" } else { write-warning "Failed Validation" } ; 
    Typical pass
    .EXAMPLE
    write-verbose 'CNAME DKIM selector records test (with referred TXT key NameHost resolution)';
    $DomainName = 'myturf.com' ; 
    foreach($sel in @('selector1','selector2')){
        $pltTDN=[ordered]@{
            Name = "$($sel)._domainkey.$($DomainName)" ; 
            Type = 'CNAME' ; 
            fltr = ''
            tvalue = "$($sel)-$($domainname.replace('.','-'))._domainkey.toroco.onmicrosoft.com" ; 
        } ;
        $smsg = "test-ExoDnsRecordTDO w`n$(($pltTDN|out-string).trim())" ; 
        write-host -foregroundcolor gray "$((get-date).ToString('HH:mm:ss')):$($smsg)"  ;
        $ret  = test-ExoDnsRecordTDO @pltTDN ; 
        if($ret.Validated -eq $true){write-host "Valid CNAME DKIM Selector Record:`n$(($ret.DNSObject|out-string).trim())" } else { write-warning "Failed Validation" } ; 
        write-host "--> Re-resolve expanded CNAME target NameHost" 
        $pltTDN=[ordered]@{
            Name = $ret.dnsobject.namehost ; 
            Type = 'TXT' ; 
            fltr = '^v=DKIM1' ; # set filter for string match post filter
            tvalue = "" ; 
        } ;
        $smsg = "test-ExoDnsRecordTDO w`n$(($pltTDN|out-string).trim())" ; 
        write-host -foregroundcolor gray "$((get-date).ToString('HH:mm:ss')):$($smsg)"  ;
        $ret  = test-ExoDnsRecordTDO @pltTDN ; 
        if($ret.Validated -eq $true){write-host "Valid TXT target DKIM key Record:`n$(($ret.DNSObject|out-string).trim())" } else { write-warning "Failed Validation" } ; 
    } ; 
    Typical pass
    .EXAMPLE
    write-verbose 'CNAME DKIM selector records test (with referred TXT key NameHost resolution)';
    $DomainName = 'myturf.com' ; 
    foreach($sel in @('dkim1','dkim2')){
        $pltTDN=[ordered]@{
            Name = "$($sel)._domainkey.$($DomainName)" ; 
            Type = 'CNAME' ; 
            fltr = ''
            tvalue = "$($sel)-$($domainname.replace('.','-'))._domainkey.toroco.onmicrosoft.com" ; 
        } ;
        $smsg = "test-ExoDnsRecordTDO w`n$(($pltTDN|out-string).trim())" ; 
        write-host -foregroundcolor gray "$((get-date).ToString('HH:mm:ss')):$($smsg)"  ;
        $ret  = test-ExoDnsRecordTDO @pltTDN ; 
        if($ret.Validated -eq $true){write-host "Valid CNAME DKIM Selector Record:`n$(($ret.DNSObject|out-string).trim())" } else { write-warning "Failed Validation" } ; 
        write-host "--> Re-resolve expanded CNAME target NameHost" 
        $pltTDN=[ordered]@{
            Name = $ret.dnsobject.namehost ; 
            Type = 'TXT' ; 
            fltr = '^v=DKIM1' ; # set filter for string match post filter
            tvalue = "" ; 
        } ;
        $smsg = "test-ExoDnsRecordTDO w`n$(($pltTDN|out-string).trim())" ; 
        write-host -foregroundcolor gray "$((get-date).ToString('HH:mm:ss')):$($smsg)"  ;
        $ret  = test-ExoDnsRecordTDO @pltTDN ; 
        if($ret.Validated -eq $true){write-host "Valid TXT target DKIM key Record:`n$(($ret.DNSObject|out-string).trim())" } else { write-warning "Failed Validation" } ; 
    } ; 
    Typical pass using non-o365 custom selictor specs
.EXAMPLE
$DomainName = 'myturf.com' ;
$sBnr3="`n#*~~~~~~v CHECK:$($DomainName): Domain Ownership Verification 'TXT' DNS record v~~~~~~" ;
write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($sBnr3)" ;
$pltTDN=[ordered]@{
    Name = $DomainName ;
    Type = 'TXT' ;
    fltr = '^MS=' ;
    tvalue = 'MS=ms60604949'
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
    fltr = '' ; # set filter for string match post filter
    tvalue = 'myturf-com.mail.protection.outlook.com'
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
    fltr = '' ; # set filter for string match post filter
    tvalue = 'autodiscover.outlook.com' ;
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
    fltr = '^v=spf1' ; # set filter for string match post filter
    tvalue = "v=spf1 ip4:148.163.146.158 ip4:148.163.142.153 ip4:170.92.0.0/16 ip4:205.142.232.90 include:spf.protection.outlook.com ~all" ;
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
        fltr = ''
        tvalue = "$($sel)-$($domainname.replace('.','-'))._domainkey.toroco.onmicrosoft.com" ;
    } ;
    $smsg = "test-ExoDnsRecordTDO w`n$(($pltTDN|out-string).trim())" ;
    write-host -foregroundcolor gray "$((get-date).ToString('HH:mm:ss')):$($smsg)"  ;
    $ret  = test-ExoDnsRecordTDO @pltTDN ;
    if($ret.Validated -eq $true){write-host "Valid CNAME DKIM Selector Record:`n$(($ret.DNSObject|out-string).trim())" } else { write-warning "Failed Validation" } ;
    write-host "--> Re-resolve expanded CNAME target NameHost"
    $pltTDN=[ordered]@{
        Name = $ret.dnsobject.namehost ;
        Type = 'TXT' ;
        fltr = '^v=DKIM1' ; # set filter for string match post filter
        tvalue = "" ;
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
        [Parameter(Mandatory=$True,HelpMessage="DNS Name[-Type TXT]")]
        [string]$Name,
        [Parameter(Mandatory=$True,HelpMessage="DNS Type[-Type TXT]")]
            [ValidateSet('MX','TXT','CNAME')]
            [string]$Type,
        [Parameter(Mandatory=$false,HelpMessage="String Type post-filter[-fltr TXT]")]
            [string]$fltr,
        [Parameter(Mandatory=$false,HelpMessage="Validating value string[-tvalue TXT]")]
            [string]$tvalue
    ) ; 
    $pltRvDN=[ordered]@{
        Name = $Name ; 
        Server = '1.1.1.1' ; 
        Type = $type ; 
        erroraction = 'STOP' ;
    } ;
    $oReturn = [ordered]@{
        DNSObject = $null ; 
        Type = $Type ; 
        Validated = $false ; 
    } ; 
    if($tvalue){$oReturn.add('TValue',$Tvalue)} ; 
    if($fltr){$oReturn.add('fltr',$fltr)} ; 
    $smsg = "resolve-DNSName w`n$(($pltRvDN|out-string).trim())" ; 
    write-host -foregroundcolor gray "$((get-date).ToString('HH:mm:ss')):$($smsg)" ;
    TRY{
        if($rec = resolve-DNSName @pltRvDN ){
            switch($pltRvDN.Type){
                'MX' {
                    if($rec.NameExchange -eq $tvalue ){
                        write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):==> NameExchange value matches spec: $($tvalue)" ; 
                        $oReturn.DNSObject = $rec ; 
                        $oReturn.Validated = $true ;
                        [pscustomobject]$oReturn | write-output  ; 
                    } else{
                        write-warning "$((get-date).ToString('HH:mm:ss')):String value DOES NOT MATCH MS specified validator!: $($tvalue)" ; 
                        $oReturn.DNSObject = $rec ; 
                        $oReturn.Validated = $false ;
                        [pscustomobject]$oReturn | write-output  ; 
                    }; 
                } ; 
                'TXT' {
                    $oReturn.add('strings',$null)
                    if($fltr -AND ($rec| ? strings -match $fltr)){
                        $rec = $rec| ? strings -match $fltr ; 
                    } ; 
                    write-host -foregroundcolor green "`n$(($rec | ft -a |out-string).trim())" ; 
                    if($tvalue -AND (($rec| select -expand strings) -eq $tvalue)){
                        write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):==> String value matches MS specified validator: $($tvalue)" ; 
                        $oReturn.DNSObject = $rec ; 
                        $oReturn.Strings = ($rec| select -expand strings) ; 
                        $oReturn.Validated = $true ;
                        [pscustomobject]$oReturn | write-output  ; 
                    }elseif($tvalue -AND (($rec| select -expand strings) -ne $tvalue)){
                        write-warning "$((get-date).ToString('HH:mm:ss')):String value DOES NOT MATCH MS specified validator!: $($tvalue)" ; 
                        $oReturn.DNSObject = $rec ; 
                        $oReturn.Strings = ($rec| select -expand strings) ; 
                        $oReturn.Validated = $false ;
                        [pscustomobject]$oReturn | write-output  ; 
                    }elseif($fltr -AND $rec){
                        write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):No comparison `$tvalue, but matched `$fltr:$($fltr): ==> String value matches specification" ; 
                        $oReturn.DNSObject = $rec ; 
                        $oReturn.Strings = ($rec| select -expand strings) ; 
                        $oReturn.Validated = $true ;
                        [pscustomobject]$oReturn | write-output  ; 
                    } else{
                        write-warning "$((get-date).ToString('HH:mm:ss')):String value DOES NOT MATCH MS specified validator!: $($tvalue)" ; 
                        $oReturn.DNSObject = $rec ; 
                        $oReturn.Strings = ($rec| select -expand strings) ; 
                        $oReturn.Validated = $false ;
                        [pscustomobject]$oReturn | write-output  ; 
                    }; 
                } 
                'CNAME' {
                    if($rec.NameHost -eq $tvalue ){
                        write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):NameHost value matches spec: $($tvalue)" ; 
                        $oReturn.DNSObject = $rec ; 
                        $oReturn.Validated = $true ;
                        [pscustomobject]$oReturn | write-output  ;
                    } else{
                        write-warning "$((get-date).ToString('HH:mm:ss')):value DOES NOT MATCH specification!: $($tvalue)" ; 
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
        throw $smsg ;
    } ; 
} ; 
#*------^ END Function test-ExoDnsRecordTDO ^------
