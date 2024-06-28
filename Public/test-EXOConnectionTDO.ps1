# test-EXOConnectionTDO

function test-EXOConnectionTDO{
    <#
    .SYNOPSIS
    test-EXOConnectionTDO.ps1 - Evaluate status of existing ExchangeOnlineManagement connections into Exchange Online or Security & Compliance backends
    .NOTES
    Version     : 0.0.1
    Author      : Todd Kadrie
    Website     : http://www.toddomation.com
    Twitter     : @tostka / http://twitter.com/tostka
    CreatedDate : 2024-
    FileName    : test-EXOConnectionTDO.ps1
    License     : MIT License
    Copyright   : (c) 2024 Todd Kadrie
    Github      : https://github.com/tostka/verb-XXX
    Tags        : Powershell
    AddedCredit : REFERENCE
    AddedWebsite: URL
    AddedTwitter: URL
    REVISIONS
    * 4:16 PM 6/28/2024 init
    .DESCRIPTION
    test-EXOConnectionTDO.ps1 - Evaluate status of existing ExchangeOnlineManagement connections into Exchange Online or Security & Compliance backends
    Refactoring/simplifying test-EXOv2Connection() into stripped down equiv: get-connectioninformation natively returns a lot of points of comparison, obsoleting bulk of the older EXOv2 & basic-auth'd EOM versions. 

    Returns the following summary of any Exchange Online or Security & Compliance session connections found via get-connectioninformation:

        Property         | Type/Value                                                               | Description
        -----------------|--------------------------------------------------------------------------|---------------------------------------------
        EXOSessions      |  Microsoft.Exchange.Management.ExoPowershellSnapin.ConnectionInformation | array of any EXO sessions found
        isEXO            |  True                                                                    | indicates an EXO connection is present (not necessarily Active)
        isEXOValid       |  True                                                                    | indicates an EXO connection is State:'Connected' & TokenStatus:'Active'}
        isEXOCba         |  True                                                                    | indicates an EXO connection is present that used Certificate Based Authentication
        isEXOOrgMatched  |  True                                                                    | indicates an EXO connection's Organization specification matches the specified Organization/TenantDomain value
        EXOTokenLifeMins |  22                                                                      | reports the number of minutes remaining until TokenExpiryTimeUTC on the EXO session (blank if expired)
        EXOPrefix        |  xo                                                                      | array of ModulePrefix values found on EXO connections
        SCSessions       |  Microsoft.Exchange.Management.ExoPowershellSnapin.ConnectionInformation | array of any SC sessions found
        isSC             |  False                                                                   | indicates an SC connection is present (not necessarily Active)
        isSCValid        |  False                                                                   | indicates an SC connection is State:'Connected' & TokenStatus:'Active'}
        isSCCba          |  False                                                                   | indicates an SC connection is present that used Certificate Based Authentication
        isSCOrgMatched   |  False                                                                   | indicates an SC connection's Organization specification matches the specified Organization/TenantDomain value
        SCTokenLifeMins  |                                                                          | reports the number of minutes remaining until TokenExpiryTimeUTC on the SC Token (blank if expired)
        SCPrefix         |                                                                          | array of ModulePrefix values found on SC connections

    .PARAMETER Organization
    Office365 TenantDomain to be referenced for Connection Validation[-Organization 'TenantDomain.onmicrosoft.com']
    .PARAMETER UserPrincipalName
    Optional UPN to be tested against for current connection[-UserPrincipalName 'UPN@domain.com']
    .PARAMETER Silent
    Switch to specify suppression of all but warn/error echos.(unimplemented, here for cross-compat)
    .INPUTS
    None. Does not accepted piped input.(.NET types, can add description)
    .OUTPUTS
    None. Returns no objects or output (.NET types)
    System.Boolean
    [| get-member the output to see what .NET obj TypeName is returned, to use here]
    .EXAMPLE
    PS> if((test-EXOConnectionTDO).isEXO){}{connect-ExchangeOnline} ; get-xomailbox -ResultSize 1 ; 
    Simple pre-connection test/connection trigger prior to running an EXO command
    Run with whatif & verbose
    .EXAMPLE
    PS> $results = test-EXOConnectionTDO ; 
    PS> results ; 

        EXOSessions      : Microsoft.Exchange.Management.ExoPowershellSnapin.ConnectionInformation
        isEXO            : True
        isEXOValid       : True
        isEXOCba         : True
        isEXOOrgMatched  : True
        EXOPrefix        : xo
        EXOTokenLifeMins : 39
        SCSessions       : 
        isSC             : False
        isSCValid        : False
        isSCCba          : False
        isSCOrgMatched   : False
        SCTokenLifeMins  : 
        SCPrefix         :       

    Demo pass, captured into variable, dumping a single EXO connection of returned proeprties.
    .LINK
    https://github.com/tostka/verb-XXX
    .LINK
    https://bitbucket.org/tostka/powershell/
    .LINK
    [ name related topic(one keyword per topic), or http://|https:// to help, or add the name of 'paired' funcs in the same niche (enable/disable-xxx)]
    #>
    [CmdletBinding()]
    PARAM(
        [Parameter(HelpMessage="Office365 TenantDomain to be referenced for Connection Validation[-Organization 'TenantDomain.onmicrosoft.com']")]
            [string]$Organization = $TORMeta.o365_TenantDomain,
        [Parameter(HelpMessage="Optional UPN to be tested against for current connection[-UserPrincipalName 'UPN@domain.com']")]
            [string]$UserPrincipalName,
        [Parameter(HelpMessage="Optional UPN to be tested against for current connection[-UserPrincipalName 'UPN@domain.com']")]
            [string]$Prefix,
        [Parameter(HelpMessage="Silent output (suppress status echos)[-silent]")]
            [switch] $silent
        #,[switch]$isCBA
    ); 
    BEGIN{
        $rgxExoPsHostName="^(ps\.outlook\.com|outlook\.office365\.com)$" ; 
        $rgxConnectionUriCCMS = 'https://ps\.compliance\.protection\.outlook\.com' ; 
        TRY{
            $conns = get-connectioninformation -ea STOP ; 
        } CATCH {
            $ErrTrapd=$Error[0] ;
            $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
            else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        } ; 
    } ;  # BEG-E
    PROCESS{
        $oReturn = [ordered]@{
        #$oReturn = @{
            EXOSessions = $null ; 
            isEXO = $false ;
            isEXOValid = $false ;
            isEXOCba = $false ; 
            isEXOOrgMatched = $false ;
            EXOPrefix = $null ;  
            EXOTokenLifeMins = $null ; 
            SCSessions = $null ; 
            isSC = $false ; 
            isSCValid = $false ; 
            isSCCba = $false ; 
            isSCOrgMatched = $false ; 
            SCTokenLifeMins = $null ; 
            SCPrefix = $null ; 
        } ; 
        if($conns){
            $oReturn.EXOSessions = $conns | ?{($_.connectionuri -match $rgxConnectionUriEXO) -AND $_.IsEopSession -eq $false} ; 
            $oReturn.SCSessions = $conns | ?{($_.connectionuri -match $rgxConnectionUriCCMS) -AND $_.IsEopSession -eq $true }
            $oReturn.isEXO = [boolean]($oReturn.EXOSessions) ; 
            $oReturn.isSC = [boolean]($oReturn.SCSessions) ; 
            $oReturn.isEXOCba = [boolean]($oReturn.EXOSessions | ?{$_.CertificateAuthentication -eq $true }) ; 
            $oReturn.isSCCba = [boolean]($oReturn.SCSessions | ?{$_.CertificateAuthentication -eq $true }) ; 
            if($oReturn.EXOSessions){
                $oReturn.EXOPrefix = $results.EXOSessions | ?{$_.modulePrefix -eq 'xo'} | select -expand ModulePrefix ; 
            } ; 
            if($oReturn.EXOSessions | ?{$_.State -eq 'Connected' -AND $_.TokenStatus -eq 'Active'}){
                $oReturn.isEXOValid = $true ; 
                $oReturn.EXOTokenLifeMins = (new-timespan -start (get-date ) -end ($oReturn.EXOSessions.TokenExpiryTimeUTC  | sort | select -last 1).LocalDateTime).minutes ; 
            } else { 
                $oReturn.isEXOValid = $false ; 
            } ;  
            if($oReturn.SCSessions){
                $oReturn.SCPrefix = $results.SCSessions | ?{$_.modulePrefix -eq 'xo'} | select -expand ModulePrefix ; 
            } ; 
            if($oReturn.EXOSessions){
                if($Organization -AND ($oReturn.EXOSessions | ?{$_.Organization -eq $Organization}) ){
                    $oReturn.isEXOOrgMatched = $true 
                }elseif(-not $Organization){
                    $oReturn.isEXOOrgMatched = $true ; 
                } else {
                    $oReturn.isEXOOrgMatched = $false ;
                } ; 
            } ; 
            if($oReturn.SCSessions | ?{$_.State -eq 'Connected' -AND $_.TokenStatus -eq 'Active'}){
                $oReturn.isSCValid = $true ; 
                $oReturn.SCTokenLifeMins = (new-timespan -start (get-date ) -end ($oReturn.SCSessions.TokenExpiryTimeUTC | sort | select -last 1).LocalDateTime).minutes ; 
            } else { 
                $oReturn.isSCValid = $false ; 
            } ; 
            if($oReturn.SCSessions){
                if($Organization -AND ($oReturn.SCSessions | ?{$_.Organization -eq $Organization }) ){
                    $oReturn.isSCOrgMatched = $true 
                } elseif(-not $Organization){
                    $oReturn.isSCOrgMatched = $true ; 
                } else {
                    $oReturn.isSCOrgMatched = $false ;
                } ; 
            } ; 
        } else { 
            $smsg = "(get-connectioninformation  failed to return any configured EXO or SC connection)" ; 
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        } ; 
    } ;  # PROC-E
    END{
        #New-Object PSObject -Property $oReturn | write-output ;
        [pscustomobject]$oReturn | write-output ;
        #$cObj = [pscustomobject] @{DisplayName=$user.DisplayName;Prop2='string'} ;
    } ; 
} ; 

