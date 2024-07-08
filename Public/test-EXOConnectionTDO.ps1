# test-EXOConnectionTDO

function test-EXOConnectionTDO{
    <#
    .SYNOPSIS
    test-EXOConnectionTDO.ps1 - Evaluate status of existing ExchangeOnlineManagement connections into Exchange Online or Security & Compliance backends. Wraps underlying ExchangeOnlineManagement module's get-ConnectionInformation cmdlet, evaluating and simplying returned info, to make for very easy status evaluation.
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
    * 4:31 PM 7/8/2024 added CBH example demoing variant output ; added blank Org resolution (resolve TenentID into equiv TenantDomain); this should always return a full set of values even if we have to work around the bugs in the MS code.
    * 9:48 AM 7/3/2 add:$AppId to return, need a way to resolve CBA back to identifiable role; resolve-UserNameToUserRole() resolves cert thumb to role, 
    * 3:59 PM 7/2/2024 added TenantID, as the Organization has been coming back blank (post filter -OR against either); 
        added parametersets, as ConnectionID & Prefix are exclusive params in get-ConnectionInformation
        stripped back & simplified, from hybrid all in one summary, 
        to just looping connections profiling each with details, and dumping the 
        results as pscustomobjects into the pipeline, for post filtering on the 
        receiving end (rather than trying to summarize what could be a mixture of 
        differnt types - xo & sc - and Tenant connections). 
    * 4:16 PM 6/28/2024 init
    .DESCRIPTION
    test-EXOConnectionTDO.ps1 - Evaluate status of existing ExchangeOnlineManagement connections into Exchange Online or Security & Compliance backends. Wraps underlying ExchangeOnlineManagement module's get-ConnectionInformation cmdlet, evaluating and simplying returned info, to make for very easy status evaluation.

    Refactoring/simplifying test-EXOv2Connection() into stripped down equiv: get-connectioninformation natively returns a lot of points of comparison, obsoleting bulk of the older EXOv2 & basic-auth'd EOM versions. 

    Because there can be a mixture of EXO & S&C sessions to multiple tenants, this takes the client-side post-filter approach, to filter out an appropriate connection for the client needs, rather than 
    trying to build the logic into this simpler function. 

    Returns the following summary of any Exchange Online or Security & Compliance session connections found via get-connectioninformation:

        Property          | Type/Value                                                               | Description
        ------------------|--------------------------------------------------------------------------|---------------------------------------------
        Connection        |  Microsoft.Exchange.Management.ExoPowershellSnapin.ConnectionInformation | array of any sessions found
        Organization      |  TENANTNAME.onmicrosoft.com                                              | Connection Organization name (TenantDomainName)
        Prefix            |  xo                                                                      | ModulePrefix for the connection
        UserPrincipalName |  ACCOUNT@DOMAIN.TLD                                                      | The account that was used to connect
        ConnectionId      |  nabbnncf-nnab-nnnn-nbnn-nnnenncnnaee                                    | A unique GUID value for the connection
        TenantID          |  nnnnnnae-enna-nnbn-nadc-nndncnnbannb                                    | The tenant ID GUID value
        ModuleName        |  C:\Users\LOGON\AppData\Local\Temp\2\tmpEXO_ebmrzk2m.vvp                 | The filename and path of the temporary data for the session. 
        isXO              |  True                                                                    | indicates an EXO connection is present (not necessarily Active)
        isSC              |  True                                                                    | indicates a Security & Compliance connection is present (not necessarily Active)
        isCBA             |  True                                                                    | indicates an EXO connection is present that used Certificate Based Authentication
        isValid           |  True                                                                    | indicates an EXO connection is TokenStatus:'Active'}
        TokenLifeMins     |  22                                                                      | reports the number of minutes remaining until TokenExpiryTimeUTC on the session (blank if expired)
                
        ---

        [Get-ConnectionInformation (ExchangePowerShell) | Microsoft Learn](https://learn.microsoft.com/en-us/powershell/module/exchange/get-connectioninformation?view=exchange-ps)

        The Get-ConnectionInformation cmdlet returns the information about all active REST-based connections with Exchange Online in the current PowerShell instance. This cmdlet is equivalent to the Get-PSSession cmdlet that's used with remote PowerShell sessions.

        The output of the cmdlet contains the following properties:

            ConnectionId: A unique GUID value for the connection. For example, 8b632b3a-a2e2-8ff3-adcd-6d119d07694b.
            State: For example, Connected.
            Id: An integer that identifies the session in the PowerShell window. The first connection is 1, the second is 2, etc.
            Name: A unique name that's based on the PowerShell environment and Id value. For example, ExchangeOnline_1 for Exchange Online PowerShell or ExchangeOnlineProtection_1 for Security & Compliance PowerShell.
            UserPrincipalName: The account that was used to connect. For example, laura@contoso.onmicrosoft.com.
            ConnectionUri: The connection endpoint that was used. For example, https://outlook.office365.com for Exchange Online PowerShell or https://nam12b.ps.compliance.protection.outlook.com for Security & Compliance PowerShell.
            AzureAdAuthorizationEndpointUri : The Microsoft Entra authorization endpoint for the connection. For example, https://login.microsoftonline.com/organizations for Exchange Online PowerShell or https://login.microsoftonline.com/organizations for Security & Compliance PowerShell.
            TokenExpiryTimeUTC: When the connection token expires. For example, 9/30/2023 6:42:24 PM +00:00.
            CertificateAuthentication: Whether certificate based authentication (also known as CBA or app-only authentication) was used to connect. Values are True or False.
            ModuleName: The filename and path of the temporary data for the session. For example, C:\Users\laura\AppData\Local\Temp\tmpEXO_a54z135k.qgv
            ModulePrefix: The value specified using the Prefix parameter in the Connect-ExchangeOnline or Connect-IPPSSession command.
            Organization: The value specified using the Organization parameter in the Connect-ExchangeOnline or Connect-IPPSSession command for CBA or managed identity connections.
            DelegatedOrganization: The value specified using the DelegatedOrganization parameter in the Connect-ExchangeOnline or Connect-IPPSSession command.
            AppId: The value specified using the AppId parameter in the Connect-ExchangeOnline or Connect-IPPSSession command for CBA connections.
            PageSize: The default maximum number of entries per page in the connection. The default value is 1000, or you can use the PageSize parameter in the Connect-ExchangeOnline command to specify a lower number. Individual cmdlets might also have a PageSize parameter.
            TenantID: The tenant ID GUID value. For example, 3750b40b-a68b-4632-9fb3-5b1aff664079.
            TokenStatus: For example, Active.
            ConnectionUsedForInbuiltCmdlets
            IsEopSession: For Exchange Online PowerShell connections, the value is False. For Security & Compliance PowerShell connections, the value is True.

        Examples
        Example 1
        PowerShell

        Get-ConnectionInformation

        This example returns a list of all active REST-based connections with Exchange Online in the current PowerShell instance.
        Example 2
        PowerShell

        Get-ConnectionInformation -ConnectionId 1a9e45e8-e7ec-498f-9ac3-0504e987fa85

        This example returns the active REST-based connection with the specified ConnectionId value.
        Example 3
        PowerShell

        Get-ConnectionInformation -ModulePrefix Contoso,Fabrikam

        This example returns a list of active REST-based connections that are using the specified prefix values.
        Parameters
        -ConnectionId

        Note: This parameter is available in version 3.2.0 or later of the module.

        The ConnectionId parameter filters the connections by ConnectionId. ConnectionId is a GUID value in the output of the Get-ConnectionInformation cmdlet that uniquely identifies a connection, even if you have multiple connections open. You can specify multiple ConnectionId values separated by commas.

        Don't use this parameter with the ModulePrefix parameter.
        Type:	String[]
        Position:	Named
        Default value:	None
        Required:	True
        Accept pipeline input:	False
        Accept wildcard characters:	False
        Applies to:	Exchange Online
        -ModulePrefix

        Note: This parameter is available in version 3.2.0 or later of the module.

        The ModulePrefix parameter filters the connections by ModulePrefix. When you use the Prefix parameter with the Connect-ExchangeOnline cmdlet, the specified text is added to the names of all Exchange Online cmdlets (for example, Get-InboundConnector becomes Get-ContosoInboundConnector). The ModulePrefix value is visible in the output of the Get-ConnectionInformation cmdlet. You can specify multiple ModulePrefix values separated by commas.

        This parameter is meaningful only for connections that were created with the Prefix parameter.

        Don't use this parameter with the ConnectionId parameter.
        Type:	String[]
        Position:	Named
        Default value:	None
        Required:	True
        Accept pipeline input:	False
        Accept wildcard characters:	False
        Applies to:	Exchange Online    

        ---

    .PARAMETER Organization
    Office365 TenantDomain to be referenced for Connection Validation[-Organization 'TenantDomain.onmicrosoft.com']
    .PARAMETER UserPrincipalName
    Optional UPN to be tested against for current connection[-UserPrincipalName 'UPN@domain.com']
    .PARAMETER Silent
    Switch to specify suppression of all but warn/error echos.(unimplemented, here for cross-compat)
    .INPUTS
    None. Does not accepted piped input.(.NET types, can add description)
    .OUTPUTS
    System.Object returns per connection summary object 
    .EXAMPLE
    PS> if((test-EXOConnectionTDO | ?{$_.isXO -AND $_.isValid}){}{connect-ExchangeOnline} ; get-xomailbox -ResultSize 1 ; 
    Simple pre-connection test/connection trigger prior to running an EXO command
    .EXAMPLE
    PS> if((test-EXOConnectionTDO -prefix cc | ?{$_.isXO -AND $_.isValid}){}{connect-ExchangeOnline} ; get-xomailbox -ResultSize 1 ; 
    Simple pre-connection test/connection trigger prior to running an EXO command
    .EXAMPLE
    PS> $results = test-EXOConnectionTDO ; 
    PS> results ; 

        Connection        : Microsoft.Exchange.Management.ExoPowershellSnapin.ConnectionInformation
        Organization      : TENENTDOMAIN.onmicrosoft.com
        Prefix            : xo
        UserPrincipalName : LOGON@DOMAIN.TLD
        ConnectionId      : nabbnncf-nnab-nnnn-nbnn-nnnenncnnaee
        TenantID          : nnnnnnae-enna-nnbn-nadc-nndncnnbannb
        ModuleName        : C:\Users\LOGON\AppData\Local\Temp\2\tmpEXO_uxszgv2h.0r1
        isXO              : True
        isSC              : False
        isCBA             : False
        isValid           : True
        TokenLifeMins     : 0      

    Demo pass, captured into variable, dumping a single EXO connection of returned proeprties.
    .EXAMPLE
    PS> write-verbose "Input Data variable designator" ; 
    PS> $tenorg = 'ABC' ;
    PS> write-verbose "connection summary select properties" ; 
    PS> $prpCXO = 'Organization','Prefix','UserprincipalName','ConnectionID','isXO','isSC','isCBA','isValid','TokenLifeMins' ; 
    PS> disconnect-exchangeonline ; 
    PS> write-verbose "query status specifying explicit prefix 'xo'" ; 
    PS> write-verbose "and postfilter for isXO, and IsValid (TokenStatus: Active)" ; 
    PS> write-verbose "and org match to a local variables with target TenantDomain and TenantID guid" ; 
    PS> if( test-exoconnectiontdo -Prefix 'xo' |
    PS>     ?{$_.isXO -AND $_.IsValid -AND (
    PS>           $_.organization -match ((gv -name "$($TenOrg)Meta").Value.o365_TenantDomain) -OR $_.TenantID -match ((gv -name "$($TenOrg)Meta").Value.o365_TenantID )
    PS>     ) }){}else{
    PS>     connect-ExchangeOnline -Prefix xo -ShowBanner:$false ;
    PS>     write-verbose "above returns no status info; so pull the context again for future checks" ;
    PS>     if($PSXOContext = test-EXOConnectionTDO -prefix xo){
    PS>         $smsg = $PSXOContext | select $prpCXO[0..3] | convertTo-MarkdownTable -border ;
    PS>         $smsg += $PSXOContext | select $prpCXO[4..8] | convertTo-MarkdownTable -border ;
    PS>         write-host -foregroundcolor green $smsg ;
    PS>     } else {
    PS>         write-warning "Not Connected" ;
    PS>     } ;
    PS> } ; get-xomailbox -ResultSize 1 ;
    PS> write-verbose "Refresh Connection Status using cached connectionid guid (more specific than even Prefix" ;
    PS> $PSXOContext = test-EXOConnectionTDO -ConnectionID $PSXOContext.connectionID.guid ; 

        | Organization | Prefix | UserPrincipalName      | ConnectionId                         |
        | ------------ | ------ | ---------------------- | ------------------------------------ |
        |              | xo     | xxxxxx.xxxxxx@DOMO.TLD | nabbnncf-nnab-nnnn-nbnn-nnnenncnnaee |
        | isXO | isSC  | isCBA | isValid | TokenLifeMins |
        | ---- | ----- | ----- | ------- | ------------- |
        | True | False | False | True    | 6             |

    Fancier EXO connectivity pretest, post test, and cached context demo (output prettied up via verb-io:convertto-markdowntabl())
    .EXAMPLE
    PS> write-verbose "$XXXMeta variable designator" ; 
    PS> $tenorg = 'TOR' ;
    PS> write-verbose "connection summary select properties" ; 
    PS> $prpCXO = 'Organization','Prefix','UserprincipalName','ConnectionID','isXO','isSC','isCBA','isValid','TokenLifeMins' ; 
    PS> disconnect-exchangeonline ; 
    PS> write-verbose "query status specifying explicit prefix 'cc'" ; 
    PS> write-verbose "and postfilter for isSC, and IsValid (TokenStatus: Active)" ; 
    PS> write-verbose "and org match to a local variables" ; 
    PS> if( test-exoconnectiontdo -Prefix 'cc' |
    PS>     ?{$_.isSC -AND $_.IsValid -AND (
    PS>           $_.organization -match ((gv -name "$($TenOrg)Meta").Value.o365_TenantDomain) -OR $_.TenantID -match ((gv -name "$($TenOrg)Meta").Value.o365_TenantID )
    PS>     ) }){}else{
    PS>     connect-IPPSSession -Prefix cc  ;
    PS>     write-verbose "above returns no status info; so pull the context again for future checks" ;
    PS>     if($PSSCContext = test-EXOConnectionTDO -prefix cc){
    PS>         $smsg = $PSSCContext | select $prpCXO[0..3] | convertTo-MarkdownTable -border ;
    PS>         $smsg += $PSSCContext | select $prpCXO[4..8] | convertTo-MarkdownTable -border ;
    PS>         write-host -foregroundcolor green $smsg ;
    PS>     } else {
    PS>         write-warning "Not Connected" ;
    PS>     } ;
    PS> } ; 
    PS> write-verbose "Refresh Connection Status using cached connectionid guid (more specific than even Prefix" ;
    PS> $PSSCContext = test-EXOConnectionTDO -ConnectionID $PSSCContext.connectionID.guid -verbose ;   

        | Organization | Prefix | UserPrincipalName      | ConnectionId                         |
        | ------------ | ------ | ---------------------- | ------------------------------------ |
        |              | xo     | xxxxxx.xxxxxx@DOMO.TLD | nnnnnnae-enna-nnbn-nadc-nndncnnbannb |
        | isXO  | isSC | isCBA | isValid | TokenLifeMins |
        | ----- | ---- | ----- | ------- | ------------- |
        | False | True | False | True    | 6             |
  
    Fancier Sec & Compliance connectivity pretest, post test, and cached context demo
    .EXAMPLE
    PS> $results = test-exoconnectiontdo ; 
    PS> write-verbose "output returned CustomObject summary" ; 
    PS> $results ; 

        Connection        : Microsoft.Exchange.Management.ExoPowershellSnapin.ConnectionInformation
        Organization      : TENANTDOMAIN.onmicrosoft.com
        Prefix            : xo
        UserPrincipalName : OAuthUser@toroco.onmicrosoft.com
        ConnectionId      : nnnnncfe-fnan-nnbn-nbdn-annnnbnennfn
        AppId             : dannnnad-endn-nnan-nnnn-nenennnnnafe
        TenantID          : nnnnnnae-enna-nnbn-nadc-nndncnnbannb
        ModuleName        : C:\Users\LOGON\AppData\Local\Temp\2\tmpEXO_0imdddz4.3sl
        isXO              : True
        isSC              : False
        isCBA             : True
        isValid           : True
        TokenLifeMins     : 34

        Connection        : Microsoft.Exchange.Management.ExoPowershellSnapin.ConnectionInformation
        Organization      : TENANTDOMAIN.onmicrosoft.com
        Prefix            : cc
        UserPrincipalName : LOGON@DOMAIN.TLD
        ConnectionId      : nnnnnnde-nnnf-nndn-nnff-andbnnnnennn
        AppId             : 
        TenantID          : nnnnnnae-enna-nnbn-nadc-nndncnnbannb
        ModuleName        : C:\Users\LOGON\AppData\Local\Temp\2\tmpEXO_czwcyyvn.mem
        isXO              : False
        isSC              : True
        isCBA             : False
        isValid           : True
        TokenLifeMins     : 43
    PS> write-verbose "output TokenStatus of the connection property of the returned object (which reflects the entire output of the get-connectioninformation cmdlet)" ; 
    PS> ($results | ? isXO).connection.tokenstatus

        Active

    Demo default pass, which returns status info on all types of EXO-based connections. In above case upper return is a CBA-based EXO connection, and the lower is an account logon Sec & Compliance connection.
    .LINK
    https://github.com/tostka/verb-EXO
    #>
    [CmdletBinding(DefaultParameterSetName='Prefix')]
    PARAM(
        [Parameter(HelpMessage="Office365 TenantDomain to be filtered on returns[-Organization 'TENANTDOMAIN.onmicrosoft.com']")]
            [string]$Organization,
            # = $TORMeta.o365_TenantDomain,
        [Parameter(ParameterSetName='Prefix',HelpMessage="Prefix value to be filtered against existing connections[-Prefix xo]")]
            [Alias('ModulePrefix')]
            [string[]]$Prefix,
        [Parameter(ParameterSetName='ConnectionID',HelpMessage="The ConnectionId parameter filters the connections by ConnectionId. ConnectionId is a GUID value in the output of the Get-ConnectionInformation cmdlet that uniquely identifies a connection, even if you have multiple connections open. You can specify multiple ConnectionId values separated by commas.[-ConnectionId [guid]]")]
            [ValidateScript({
                [boolean]([guid]$_)
            })]
            [string[]]$ConnectionId,
        [Parameter(HelpMessage="Silent output (suppress status echos)[-silent]")]
            [switch] $silent
        #,[switch]$isCBA
    ); 
    BEGIN{
        $Verbose = ($VerbosePreference -eq 'Continue') ;
        $rgxExoPsHostName="^(ps\.outlook\.com|outlook\.office365\.com)$" ; 
        $rgxConnectionUriCCMS = 'https://ps\.compliance\.protection\.outlook\.com' ; 
        TRY{
            $pltGConn=[ordered]@{
                erroraction = 'STOP' ;
            } ;
            if($Prefix){
                $pltGConn.add('ModulePrefix',$Prefix) ; 
            } ;
            if($ConnectionID){
                $pltGConn.add('ConnectionID',$ConnectionID) ; 
            } ; 
            $smsg = "get-connectioninformation w`n$(($pltGConn|out-string).trim())" ; 
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            $connections = get-connectioninformation @pltGConn ; 
            
        } CATCH {
            $ErrTrapd=$Error[0] ;
            $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
            else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        } ; 
    } ;  # BEG-E
    PROCESS{
        $connections|%{
            $sessRet = [ordered]@{
                Connection = $_ ; 
                Organization = $_.Organization ; 
                Prefix = $_.ModulePrefix ;  
                UserPrincipalName = $_.UserPrincipalName ; 
                ConnectionId = $_.ConnectionId ; 
                AppId = $_.AppID ;
                TenantID = $_.TenantID ; 
                ModuleName = $_.ModuleName ; 
                isXO = [boolean](($_.connectionuri -match $rgxConnectionUriEXO) -AND $_.IsEopSession -eq $false)
                isSC = [boolean](($_.connectionuri -match $rgxConnectionUriCCMS) -AND $_.IsEopSession -eq $true)
                isCBA = [boolean]($_.CertificateAuthentication); 
                isValid = [boolean]($_.TokenStatus -eq 'Active') ; 
                TokenLifeMins = if($_.TokenExpiryTimeUTC){(new-timespan -start (get-date ) -end ($_.TokenExpiryTimeUTC).LocalDateTime).minutes}else{$null} ;  ; 
                #$null ; 
            } ; 
            #if($_.TokenExpiryTimeUTC){
            #    $sessRet.TokenLifeMins = (new-timespan -start (get-date ) -end ($_.TokenExpiryTimeUTC).LocalDateTime).minutes ; ; 
            #} ; 
            if($null -eq $sessRet.Organization -AND $sessRet.TenantID){
                $Tenantdomain = convert-TenantIdToDomainName -TenantId $sessRet.TenantID ;
                $smsg = "(coercing blank Session Org, to resolved TenantID equivelent TenantDomain)" ; 
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                $sessRet.Organization = $Tenantdomain ; 
            } ; 
            if($Organization){
                # instead of compare test, use it as a post-filter
                if($sessRet = $sessRet | ?{$_.Organization -match $Organization}){
                    [pscustomobject]$sessRet | Write-Output ;
                } else {
                    $smsg = "(no existing connection matched: `$_.Organization -match $($Organization))" ; 
                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                } ; 
            } else{ 
                [pscustomobject]$sessRet | Write-Output ; 
            } ; 
        } ;    
    } ;  # PROC-E
    END{} ; 
} ; 

