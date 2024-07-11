# test-ExoDnsRecordTDO_func.ps1

#*------v Function test-ExoDnsRecordTDO v------
if(-not (get-command test-ExoDnsRecordTDO -ea 0)){
    function test-ExoDnsRecordTDO{
        <#
        .SYNOPSIS
        test-ExoDnsRecordTDO - Boilerplate wrapper for Resolve-DNSName, that runs tests and validates proper returns, against specified testFalue. for Mail-related DNS Records (MX, TXT Domain Verific, SPF & DKIM; CNAME autodiscover record) 
        .NOTES
        Version     : 0.0.1
        Author      : Todd Kadrie
        Website     : http://www.toddomation.com
        Twitter     : @tostka / http://twitter.com/tostka
        CreatedDate : 2024-06-07
        FileName    : test-ExoDnsRecordTDO
        License     : MIT License
        Copyright   : (c) 2024 Todd Kadrie
        Github      : https://github.com/tostka/verb-EXO
        Tags        : Powershell,AzureAD,Authentication,Test
        AddedCredit : 
        AddedWebsite: 
        AddedTwitter: 
        REVISIONS
        * 3:16 PM 7/11/2024 cleaned up CBH params etc
        * 3:40 PM 6/18/2024 ren $fltr -> $filter ; $tvalue -> $testValue ; round out into full function; pull sources from Metas; shift into param intputs
        * 6:16 PM 6/12/2024 init
        .DESCRIPTION
        test-ExoDnsRecordTDO - Boilerplate wrapper for Resolve-DNSName, that runs tests and validates proper returns, against specified testFalue. for Mail-related DNS Records (MX, TXT Domain Verific, SPF & DKIM; CNAME autodiscover record) 
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
