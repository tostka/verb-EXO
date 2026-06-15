# test-EXOConnectionTerseTDO

function test-EXOConnectionTerseTDO{
    <#
    .SYNOPSIS
    test-EXOConnectionTerseTDO.ps1 - Simplified wrapper of get-connection status with simpler confirmation output, on CBA v interactive long.
    .NOTES
    Version     : 0.0.1
    Author      : Todd Kadrie
    Website     : http://www.toddomation.com
    Twitter     : @tostka / http://twitter.com/tostka
    CreatedDate : 2024-
    FileName    : test-EXOConnectionTerseTDO.ps1
    License     : MIT License
    Copyright   : (c) 2024 Todd Kadrie
    Github      : https://github.com/tostka/verb-XXX
    Tags        : Powershell
    AddedCredit : REFERENCE
    AddedWebsite: URL
    AddedTwitter: URL
    REVISIONS
    * 8:32 AM 6/10/2026 init
    .DESCRIPTION
    test-EXOConnectionTerseTDO.ps1 - Simplified wrapper of get-connection status with simpler confirmation output, on CBA v interactive long.

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

    .PARAMETER Silent
    Switch to specify suppression of all but warn/error echos.(unimplemented, here for cross-compat)
    .INPUTS
    None. Does not accepted piped input.(.NET types, can add description)
    .OUTPUTS
    System.Boolean reflecting connected status; echos summary info
    .EXAMPLE
    PS> if((test-EXOConnectionTerseTDO | ?{$_.isXO -AND $_.isValid}){}{connect-ExchangeOnline} ; get-xomailbox -ResultSize 1 ; 
        
          State     TokenStat isCBA AppId                                Org
          -----     --------- ----- -----                                ---
          Connected Active     True da4551ad-e6d9-42a2-8738-3e7e90081afe toroco.onmicrosoft.com √ PASS

    Simple pre-connection test/connection trigger prior to running an EXO command, output demos return on a connected CBA    
    .EXAMPLE
    PS> if((test-EXOConnectionTerseTDO | ?{$_.isXO -AND $_.isValid}){}{connect-ExchangeOnline} ; get-xomailbox -ResultSize 1 ; 
        
          State  TokenStat isCBA AppId                                Org
          -----  --------- ----- -----                                ---
          Broken Expired    True da4551ad-e6d9-42a2-8738-3e7e90081afe toroco.onmicrosoft.com  /!\ FAIL

    Simple pre-connection test/connection trigger prior to running an EXO command, output demos return on a disconnected CBA
    .LINK
    https://github.com/tostka/verb-EXO
    #>
    [CmdletBinding(DefaultParameterSetName='Prefix')]
    [Alias('txot','test-xot')]
    PARAM(        
        [Parameter(HelpMessage="Silent output (suppress status echos)[-silent]")]
            [switch] $silent        
    ); 
    BEGIN{
        $Verbose = ($VerbosePreference -eq 'Continue') ;$ps
        $rgxExoPsHostName="^(ps\.outlook\.com|outlook\.office365\.com)$" ; 
        $rgxConnectionUriCCMS = 'https://ps\.compliance\.protection\.outlook\.com' ; 
        $prpTxoCBA = 'State',@{Name='TokenStat';Expression={$_.TokenStatus }},@{Name='Org';Expression={if($_.organization.length){$_.organization }else{' '}}},@{Name='isCBA';Expression={$_.CertificateAuthentication }},'AppId'
        $prpTxoSID = 'State',@{Name='TokenStat';Expression={$_.TokenStatus }},@{Name='Org';Expression={if($_.organization.length){$_.organization }else{' '}}},@{Name='UPN';Expression={$_.Userprincipalname }}
        #region WH_COLORCOMBOS ; #*------v WH_COLORCOMBOS v------
        # USE: write-host $smsg @whCFAIL # SEE 7psBnr for detailed WH_COLORCOMBOS_LOOP_TRY_CATCH_DEMO
        $whcBnr = if(gcm get-colorcombo -ea 0){get-colorcombo 62 | write-output }else{ @{BackgroundColor='Magenta';ForegroundColor='Black'}| write-output  } ; # BNR COLLORS
        $whcFAIL = if(gcm get-colorcombo -ea 0){get-colorcombo 20 | write-output }else{ @{BackgroundColor='DarkRed';ForegroundColor='Yellow'}| write-output  } ; # FAIL COLORS
        $whcSucc = if(gcm get-colorcombo -ea 0){get-colorcombo 49 | write-output }else{ @{BackgroundColor='Green';ForegroundColor='Blue'}| write-output  } ;  # SUCCESS COLORS
        $whcACT = if(gcm get-colorcombo -ea 0){get-colorcombo 52 | write-output }else{ @{BackgroundColor='Yellow';ForegroundColor='Black'}| write-output  } ; # PRE ATTEMPT NOTICE COLORS
        $whcRpt = if(gcm get-colorcombo -ea 0){get-colorcombo 48 | write-output }else{ @{BackgroundColor='White';ForegroundColor='Black'}| write-output  } ;  # POST REPORT  COLORS
        $whcStat = if(gcm get-colorcombo -ea 0){get-colorcombo 28 | write-output }else{ @{BackgroundColor='Gray';ForegroundColor='Blue'}| write-output  } ;  # LOOP PROCESS STATUS COLORS
        # RANDO W-H: $wh = get-colorcombo -rand ; write-host @wh "->Resolved XOMailbox: $($($ThisXoMbx.userprincipalname))" ; 
        #endregion WH_COLORCOMBOS ; #*------^ END WH_COLORCOMBOS ^------
        #region WHPASSFAIL ; #*======v WHPASSFAIL v======
        $whTPad = 72  ; $whTChar = '.' ; # scale $whTPad to longest Testing:xxx line you use in the test array
        if(-not $whPASS){$whPASS = @{ Object = "$([Char]8730) PASS`n" ; ForegroundColor = 'Green' ; NoNewLine = $true  } }
        if(-not $whFAIL){$whFAIL = @{'Object'= if ($env:WT_SESSION) { "$([char]0x274C) FAIL`n"} else {" /!\ FAIL`n"}; ForegroundColor = 'RED' ; NoNewLine = $true } } ;
        if(-not $psPASS){$psPASS = "$([Char]8730) PASS`n" } # $smsg = $pspass + " :Tested Drives" ; write-host $smsg ;
        if(-not $psFAIL){$psFAIL = if ($env:WT_SESSION) { "$([Char]8730) FAIL`n"} else {" /!\ FAIL`n"} } ; # $smsg = $psfail + " :Tested Drives" ; write-warning $smsg ;    
        # Update whPass/Fail to reflect colors in $whcSucc/Fail (if configured)
        if($whcSucc){$whcSucc.GetEnumerator() | %{if($whPass.containsKey($_.key)){$whPass[$_.Name]=$_.value}else{$whPass.add($_.name,$_.value)}}} ; 
        if($whcFAIL){$whcFAIL.GetEnumerator() | %{if($whFAIL.containsKey($_.key)){$whFAIL[$_.Name]=$_.value}else{$whFAIL.add($_.name,$_.value)}}} ; 
        
        TRY{
            $connections = get-connectioninformation  ; 
            
        } CATCH {
            $ErrTrapd=$Error[0] ;
            $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
            else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        } ; 
    } ;  # BEG-E
    PROCESS{
        if($connections){
            $connections|%{
            
                $thisxoc = $_ ;
                #if($thisxoc.state -eq 'Connected' -AND $thisxoc.TokenStatus -eq 'Active'){$pass = $true}else{$pass=$false} ;
                # per CPT, tokenstatus isn't dispositive, can still refresh, need to run get-exomailbox -resultsize 1 to know if refreshes cleanly
                <#
                ▒▒▒▒▒ [PS]:D:\s\build $ Get-ConnectionInformation
                ConnectionId                    : ac1992ad-8638-4837-bc86-35288c091967
                State                           : Broken
                Id                              : 11
                Name                            : ExchangeOnline_11
                UserPrincipalName               : OAuthUser@toroco.onmicrosoft.com
                ConnectionUri                   : https://outlook.office365.com
                AzureAdAuthorizationEndpointUri : https://login.microsoftonline.com/toroco.onmicrosoft.com
                TokenExpiryTimeUTC              :
                CertificateAuthentication       : True
                ModuleName                      : C:\Users\kadriTSS\AppData\Local\Temp\3\tmpEXO_js3n4315.qpn
                ModulePrefix                    : xo
                Organization                    : toroco.onmicrosoft.com
                DelegatedOrganization           :
                AppId                           : da4551ad-e6d9-42a2-8738-3e7e90081afe
                PageSize                        : 1000
                TenantID                        : 549366ae-e80a-44b9-8adc-52d0c29ba08b
                TokenStatus                     : Expired
                ConnectionUsedForInbuiltCmdlets : True
                IsEopSession                    : False
                
                Issue: get-exomailbox took ~1m to fail:
                ▒▒▒▒▒ [PS]:D:\s\build $ get-exomailbox -ResultSize 1
                get-exomailbox : The underlying connection was closed: An unexpected error occurred on a receive.
                At line:1 char:1
                + get-exomailbox -ResultSize 1
                + ~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                    + CategoryInfo          : ProtocolError: (:) [Get-EXOMailbox], DataServiceTransportException
                    + FullyQualifiedErrorId : The underlying connection was closed: An unexpected error occurred on a receive.,Microsoft.Exchange.Management.RestApiClient.GetExoMailbox                
                #>                
                if($thisxoc.state -eq 'Connected' -AND $thisxoc.TokenStatus -eq 'Active'){$pass = $true}else{$pass=$false} ;
                if($thisxoc.CertificateAuthentication){
                    $smsg = "`n$(($thisxoc|ft -a $prpTxoCBA|out-string).trim())" ; 
                    Write-Host "$($smsg) " -NoNewline ;                    
                }else{
                    $smsg = "`n$(($thisxoc|ft -a $prpTxoSID|out-string).trim())" ; 
                    Write-Host "$($smsg) " -NoNewline ;                                        
                }
                if ($pass) {
                    Write-Host @whPASS
                    $true | write-output  ; 
                } else {
                    write-host @whFAIL 
                    $false | write-output 
                }; 
            } ;    
        }else {
            $smsg = "(No EXO EOM connections found)" ; 
            if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
            #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
            # get-connectioninformation returns nothing when no connection, even with verbose
            # so we're not going to return a 'faked' summary to indicate a non-connection.
            $false | write-output 
        } ; 
    } ;  # PROC-E
    END{} ; 
} ; 

