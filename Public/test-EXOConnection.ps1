#*----------v Function test-EXOConnection() v----------
function test-EXOConnection {
    <#
    .SYNOPSIS
    test-EXOConnection.ps1 - Validate EXO connection, and that the proper Tenant is connected (as per provided Credential)
    .NOTES
    Version     : 1.0.0
    Author      : Todd Kadrie
    Website     :	http://www.toddomation.com
    Twitter     :	@tostka / http://twitter.com/tostka
    CreatedDate : 2022-06-24
    FileName    : test-EXOConnection.ps1
    License     : MIT License
    Copyright   : (c) 2022 Todd Kadrie
    Github      : https://github.com/tostka/verb-EXO
    Tags        : Powershell
    REVISIONS
    * 10:18 AM 6/24/2022 init
    .DESCRIPTION
    test-EXOConnection.ps1 - Validate EXO connection, and that the proper Tenant is connected (as per provided Credential)
    .PARAMETER Credential
    Credential to be used for connection
    .OUTPUT
    System.Boolean
    .EXAMPLE
    PS> $isEXOValid = test-EXOConnection -User $AADUser user@domain.com -verbose
    PS> if($isEXOValid){write-host 'Has EXO Usermailbox Type License'} else { write-warning 'NO EXO USERMAILBOX TYPE LICENSE!'} ; 
    Evaluate IsLicensed status on passed UPN object
    .EXAMPLE
    PS> $ExMbxLicenses = get-ExoMailboxLicenses ;
    PS> [regex]$rgxExLics = ('(' + (($ExMbxLicenses.GetEnumerator().name |%{[regex]::escape($_)}) -join '|') + ')') ; 
    Demo pulling the underlying licenses list and building a regex for static use
    .LINK
    https://github.com/tostka/verb-EXO
    #>
    #Requires -Version 3
    ##Requires -Modules AzureAD, verb-Text
    ##Requires -RunasAdministrator
    # VALIDATORS: [ValidateNotNull()][ValidateNotNullOrEmpty()][ValidateLength(24,25)][ValidateLength(5)][ValidatePattern("some\sregex\sexpr")][ValidateSet("USEA","GBMK","AUSYD")][ValidateScript({Test-Path $_ -PathType 'Container'})][ValidateScript({Test-Path $_})][ValidateRange(21,65)][ValidateCount(1,3)]
    [CmdletBinding()]
     Param(
        [Parameter(Mandatory=$True,HelpMessage="Credentials [-Credentials [credential object]]")]
        [System.Management.Automation.PSCredential]$Credential = $global:credo365TORSID
    )
    BEGIN {
        ${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name ;
        $Verbose = ($VerbosePreference -eq 'Continue') ;
        if(-not $rgxCertFNameSuffix){$rgxCertFNameSuffix = '-([A-Z]{3})$' } ; 
        #*------v PSS & GMO VARIS v------
        # get-pssession session varis
        # select key differentiating properties:
        $pssprops = 'Id','ComputerName','ComputerType','State','ConfigurationName','Availability', 
            'Description','Guid','Name','Path','PrivateData','RootModuleModule', 
            @{name='runspace.ConnectionInfo.ConnectionUri';Expression={$_.runspace.ConnectionInfo.ConnectionUri} },  
            @{name='runspace.ConnectionInfo.ComputerName';Expression={$_.runspace.ConnectionInfo.ComputerName} },  
            @{name='runspace.ConnectionInfo.Port';Expression={$_.runspace.ConnectionInfo.Port} },  
            @{name='runspace.ConnectionInfo.AppName';Expression={$_.runspace.ConnectionInfo.AppName} },  
            @{name='runspace.ConnectionInfo.Credentialusername';Expression={$_.runspace.ConnectionInfo.Credential.username} },  
            @{name='runspace.ConnectionInfo.AuthenticationMechanism';Expression={$_.runspace.ConnectionInfo.AuthenticationMechanism } },  
            @{name='runspace.ExpiresOn';Expression={$_.runspace.ExpiresOn} } ; 
        $EXOv1ConfigurationName = $EXOv2ConfigurationName = $EXoPConfigurationName = "Microsoft.Exchange" ;

        if(-not $EXOv1ConfigurationName){$EXOv1ConfigurationName = "Microsoft.Exchange" };
        if(-not $EXOv2ConfigurationName){$EXOv2ConfigurationName = "Microsoft.Exchange" };
        if(-not $EXoPConfigurationName){$EXoPConfigurationName = "Microsoft.Exchange" };

        if(-not $EXOv1ComputerName){$EXOv1ComputerName = 'ps.outlook.com' };
        if(-not $EXOv1runspaceConnectionInfoAppName){$EXOv1runspaceConnectionInfoAppName = '/PowerShell-LiveID'  };
        if(-not $EXOv1runspaceConnectionInfoPort){$EXOv1runspaceConnectionInfoPort -eq '443' };

        if(-not $EXOv2ComputerName){$EXOv2ComputerName = 'outlook.office365.com' ;}
        if(-not $EXOv2Name){$EXOv2Name = "ExchangeOnlineInternalSession*" ; }
        if(-not $rgxEXoPrunspaceConnectionInfoAppName){$rgxEXoPrunspaceConnectionInfoAppName = '^/(exadmin|powershell)$'}; 
        if(-not $EXoPrunspaceConnectionInfoPort){$EXoPrunspaceConnectionInfoPort = '80' } ; 
        # gmo varis
        if(-not $rgxEXOv1gmoDescription){$rgxEXOv1gmoDescription = "^Implicit\sremoting\sfor\shttps://ps\.outlook\.com/PowerShell" }; 
        if(-not $EXOv1gmoprivatedataImplicitRemoting){$EXOv1gmoprivatedataImplicitRemoting = $true };
        if(-not $rgxEXOv2gmoDescription){$rgxEXOv2gmoDescription = "^Implicit\sremoting\sfor\shttps://outlook\.office365\.com/PowerShell" }; 
        if(-not $EXOv2gmoprivatedataImplicitRemoting){$EXOv2gmoprivatedataImplicitRemoting = $true } ;
        if(-not $rgxExoPsessionstatemoduleDescription){$rgxExoPsessionstatemoduleDescription = '/(exadmin|powershell)$' };
        if(-not $PSSStateOK){$PSSStateOK = 'Opened' };
        if(-not $PSSAvailabilityOK){$PSSAvailabilityOK = 'Available' };
        $modname = 'ExchangeOnlineManagement' ;
        #*------^ END PSS & GMO VARIS ^------

        # check if using Pipeline input or explicit params:
        if ($PSCmdlet.MyInvocation.ExpectingInput) {
            write-verbose "Data received from pipeline input: '$($InputObject)'" ;
        } else {
            # doesn't actually return an obj in the echo
            write-verbose "Data received from parameter input:" # '$($InputObject)'" ;
        } ;
    } 
    PROCESS {
        $isEXOValid = $false ;
        if($pssEXOv2 = Get-PSSession | 
                where-object {$_.ConfigurationName -like "Microsoft.Exchange" -AND (
                    $_.Name -like "ExchangeOnlineInternalSession*") -AND (
                    $_.ComputerName -eq $EXOv2ComputerName) }  -AND (
                    $_.State -eq $PSSStateOK)  -AND (
                    $_.Availability -eq $PSSAvailabilityOK
                    )){
            $smsg += "`n`nEXOv2 PSSessions:`n$(($pssEXOv2 | fl $pssprops|out-string).trim())" ; 
             # verify the exov2 cmdlets actually imported as a tmp_ module w specifid prefix & 1st cmdlet
            if ( (get-module -name tmp_* | ForEach-Object { Get-Command -module $_.name -name 'Add-xoAvailabilityAddressSpace' -ea 0 }) -AND (test-EXOToken) ) {
                $bExistingEXOGood = $true ;
            } else { $bExistingEXOGood = $false ; }
            # implement caching of accepteddoms into the XXXMeta, in the session (cut back on queries to EXO on acceptedom)
            # swap in non-looping
            if( get-command Get-xoAcceptedDomain) {
                 #$TenOrg = get-TenantTag -Credential $Credential ;
                if(-not (Get-Variable  -name "$($TenOrg)Meta").value.o365_AcceptedDomains){
                    set-Variable  -name "$($TenOrg)Meta" -value ( (Get-Variable  -name "$($TenOrg)Meta").value += @{'o365_AcceptedDomains' = (Get-xoAcceptedDomain).domainname} )
                } ;
            } ;
            
            $smsg = "(validating Tenant:Credential alignment)" ; 
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            
            if( ($credential.username -match $rgxCertThumbprint) -AND ((Get-Variable  -name "$($TenOrg)Meta").value.o365_Prefix -eq $certTag )){
                # 9:59 AM 6/24/2022 need a case for CBA cert (thumbprint username)
                # compare cert fname suffix to $xxxMeta.o365_Prefix
                # validate that the connected EXO is to the CBA Cert tenant
                #$smsg = "(EXO Authenticated & Functional CBA cert:$($Credential.username.split('@')[1].tostring())),($($Credential.username))" ;
                $smsg = "(EXO Authenticated & Functional CBA cert:$($certTag),($($certUname))" ;
                if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                $bExistingEXOGood = $isEXOValid = $true ;
            }elseif((Get-Variable  -name "$($TenOrg)Meta").value.o365_AcceptedDomains.contains($Credential.username.split('@')[1].tostring())){
                # validate that the connected EXO is to the $Credential tenant
                $smsg = "(EXO Authenticated & Functional:$($Credential.username.split('@')[1].tostring())),($($Credential.username))" ;
                if($silent){}elseif($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                $bExistingEXOGood = $isEXOValid = $true ;
            # issue: found fresh bug in cxo: svcacct UPN suffix @tenantname.onmicrosoft.com, but testing against AccepteDomain, it's not in there (tho @DOMAIN.mail.onmicrosoft.comis)
            }elseif((Get-Variable  -name "$($TenOrg)Meta").value.o365_TenantDomain -eq ($Credential.username.split('@')[1].tostring())){
                $smsg = "(EXO Authenticated & Functional(TenDom):$($Credential.username.split('@')[1].tostring()))" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                $bExistingEXOGood = $isEXOValid = $true ;
            } else {
                $smsg = "(Credential mismatch:disconnecting from existing EXO:$($eEXO.Identity) tenant)" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } 
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                Disconnect-exo ;
                $bExistingEXOGood = $isEXOValid = $false ;
            } ;
            
        } else { 
            $smsg = "Unable to detect EXOv2 PSSession!" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } 
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
            throw $smsg ;
            Break ; 
        } ; 

    }  # PROC-E
    END{
        $smsg = "(Returning `$isEXOValid:$($isEXOValid) to pipeline)" ; 
        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        $isEXOValid | write-output ; 
    } ;
} ; 
#*------^ END Function test-EXOConnection() ^------