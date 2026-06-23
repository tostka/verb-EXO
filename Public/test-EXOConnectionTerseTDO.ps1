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
    * 8:31 AM 6/19/2026 revised logic, to validate as: get-connection -AND get-xorecipient -resultsize 1; confirms autorefresh, fail runs dxo to clear, succ, post-runs disconnect-exchangeonline on each post broken id. 
    might need something akin to get-xorecipient for purview module connections confirm; now outputs summary object, or $false on no connect. Emits trailing cxo sample as well.
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
        [Parameter(ParameterSetName='Prefix',HelpMessage="Prefix value to be filtered against existing connections[-Prefix xo]")]
            [Alias('ModulePrefix')]
            [string[]]$Prefix,
        [Parameter(HelpMessage="Silent output (suppress status echos)[-silent]")]
            [switch] $silent        
    ); 

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
        
    TRY {
        $connections = Get-ConnectionInformation -ErrorAction SilentlyContinue ; 
        $isHealthy = $false
        if ($connections) {
            if($Prefix){
                set-alias -name g-xorecipient -Value "get-$($Prefix)Recipient"  ; 
                try {g-xorecipient -Res 1 -wa 0 -ErrorAction Stop | Out-Null ; $isGood = $true}catch {$isGood = $false}
                remove-alias -alias g-xorecipient ; 
            }else{
                try {get-recipient -Res 1 -wa 0 -ErrorAction Stop | Out-Null ; $isGood = $true}catch {$isGood = $false}
            } ;      
        }
        if ($isHealthy) {
            $connections | Where-Object { $_.State -eq 'Broken' } | 
                ForEach-Object {
                    Write-Verbose "Closing broken connection $($_.ConnectionId)"
                    Disconnect-ExchangeOnline -ConnectionId $_.ConnectionId -Confirm:$false -ErrorAction SilentlyContinue
                } ; 
            $connections | Where-Object { $_.State -ne 'Broken' } |
                ForEach-Object {
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
                            isSC = [boolean](($_.connectionuri -match $rgxConnectionUriPrvw) -AND $_.IsEopSession -eq $true)
                            isCBA = [boolean]($_.CertificateAuthentication); 
                            isValid = [boolean]($_.TokenStatus -eq 'Active') ; 
                            TokenLifeMins = if($_.TokenExpiryTimeUTC){(new-timespan -start (get-date ) -end ($_.TokenExpiryTimeUTC).LocalDateTime).minutes}else{$null} ;  ; 
                            #$null ; 
                        } ; 
                        [pscustomobject]$sessRet | Write-Output ; 
                    } ; 

        } else {
            write-host -foregroundcolor yellow "No valid session detected; resetting Exchange Online connection, for fresh attempt"
            # Clear everything (important for avoiding session buildup)
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
            # Reconnect cleanly
            #Connect-ExchangeOnline -ShowBanner:$false
            write-host -foregroundcolor yellow "To connect run:Connect-ExchangeOnline -ShowBanner:`$false " ;                 
            $false | write-output ; 
        }
    }CATCH {
        $ErrTrapd = $_
        $smsg = "`n$($ErrTrapd | Format-List * -Force | Out-String)".Trim()
        if ($logging) {Write-Log -LogContent $smsg -Path $logfile -UseHost -Level WARN}
        else {Write-Warning "$((Get-Date).ToString('HH:mm:ss')): $smsg"}
        $false | write-output ; 
    }
} ; 

