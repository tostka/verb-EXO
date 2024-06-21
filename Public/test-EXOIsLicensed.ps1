# test-EXOIsLicensed_func.ps1

#*------v test-EXOIsLicensed.ps1 v------
function test-EXOIsLicensed {
    <#
    .SYNOPSIS
    test-EXOIsLicensed.ps1 - Evaluate IsLicensed status, to indicate license support for Exchange online UserMailbox type, on passed in AzureADUser object
    .NOTES
    Version     : 1.0.0
    Author      : Todd Kadrie
    Website     :	http://www.toddomation.com
    Twitter     :	@tostka / http://twitter.com/tostka
    CreatedDate : 2022-03-22
    FileName    : test-EXOIsLicensed.ps1
    License     : MIT License
    Copyright   : (c) 2022 Todd Kadrie
    Github      : https://github.com/tostka/verb-XXX
    Tags        : Powershell
    REVISIONS
    * 3:52 PM 5/23/2023 implemented @rxo @rxoc split, (silence all connectivity, non-silent feedback of functions); flipped all r|cxo to @pltrxoC, and left all function calls as @pltrxo; 
    * 9:13 AM 5/22/2023 added Silent back, for broad call compatibility (pltrxo consistency); 
    * 2:39 PM 5/17/2023 add pltrxo support
    * 3:15 PM 5/15/2023:test-EXOIsLicensed() works w latest aad/exo-eom updates
    * 1:06 PM 4/4/2022 updated CBH example to reflect $AADU obj, not UPN input
    3:08 PM 3/23/2022 init
    .DESCRIPTION
    test-EXOIsLicensed.ps1 - Evaluate IsLicensed status, to indicate license support for Exchange online UserMailbox type, on passed in AzureADUser object
    Coordinates with verb-exo:get-ExoMailboxLicenses() to retrieve a static list of UserMailbox -supporting license names & sku's in our Tenant. 

    The get-EXOMailboxLicenses list is *not* interactive with AzureAD or EXO, 
    -- CORRECTION: the dependant get-AADlicensePlanList() includes an AAD call to pull the sku's: Connect-AAD -Credential:$Credential -verbose:$($verbose) -silent ;
    but that func needs working access, not the code w/in this.
    ...and it *will* have to be tuned for local Tenants, and maintained for currency over time. 

    It's a simple test, but it beats..
        ...the prior get-Msoluser |?{$_.islicensed} (which indicates:*some* license is assigned - could be a worthless 'FREEFLOW'!) 
        
        ... or testing |?{$_.LicenseReconciliationNeeded } 
        ( which used to indicate a mailbox *exists* but lacks a suitable mailbox-supporting license, 
        and continues to be mounted, *solely* due to being within 30days of onboarding to cloud.).  

    Not to mention get-AzureADuser's complete lack of any native evaluation on either front. [facepalm]
    Nor any similar native support in the gap from the ExchangeOnlineManagement module. 

    <rant>
        I *love* coding coverage for slipshod MS module providers that write to replace *force*-deprecated critical infra tools, 
        but can't be bothered to deliver equiv function, equiv parameters, or even similar outputs, 
        for long-standing higher-functioning tools, when they write the half-implemented *new* ones.

        And no, "Just make calls to GraphAPI!", is not a viable answer, for *working* admins, mandated to deliver working solutions on tight schedules. 
        If we wanted to be REST web devs, we wouldn't be running o365 services!
    </rant>

    .PARAMETER  User
    AzureADUser [Microsoft.Open.AzureAD.Model.User] object
    .PARAMETER TenOrg
    Tenant Tag (3-letter abbrebiation)[-TenOrg 'XYZ']
    .PARAMETER Credential
    Use specific Credentials (defaults to Tenant-defined SvcAccount)[-Credentials [credential object]]
    .PARAMETER UserRole
    Credential User Role spec (SID|CSID|UID|B2BI|CSVC|ESVC|LSVC|ESvcCBA|CSvcCBA|SIDCBA)[-UserRole @('SIDCBA','SID','CSVC')]
    .PARAMETER useEXOv2
    Use EXOv2 (ExchangeOnlineManagement) over basic auth legacy connection [-useEXOv2]
    .PARAMETER Silent
    Switch to specify suppression of all but warn/error echos.(unimplemented, here for cross-compat)
    .OUTPUT
    System.Boolean
    .EXAMPLE
    PS> $isEXOLicensed = test-EXOIsLicensed -User $AADUser -verbose
    PS> if($isEXOLicensed){write-host 'Has EXO Usermailbox Type License'} else { write-warning 'NO EXO USERMAILBOX TYPE LICENSE!'} ; 
    Evaluate IsLicensed status on passed UPN object
    .LINK
    https://github.com/tostka/verb-EXO
    #>
    #Requires -Version 3
    ##Requires -Modules AzureAD, verb-Text
    ##Requires -RunasAdministrator
    # VALIDATORS: [ValidateNotNull()][ValidateNotNullOrEmpty()][ValidateLength(24,25)][ValidateLength(5)][ValidatePattern("some\sregex\sexpr")][ValidateSet("USEA","GBMK","AUSYD")][ValidateScript({Test-Path $_ -PathType 'Container'})][ValidateScript({Test-Path $_})][ValidateRange(21,65)][ValidateCount(1,3)]
    [CmdletBinding()]
    
     Param(
        [Parameter(Position=0,Mandatory=$True,ValueFromPipeline=$true,HelpMessage="Either Msoluser object or UserPrincipalName for user[-User upn@domain.com|`$msoluserobj ]")]
            [Microsoft.Open.AzureAD.Model.User]$User,
        # Service Connection Supporting Varis (AAD, EXO, EXOP)
        [Parameter(Mandatory=$FALSE,HelpMessage="TenantTag value, indicating Tenants to connect to[-TenOrg 'TOL']")]
            [ValidateNotNullOrEmpty()]
            #[ValidatePattern("^\w{3}$")]
            [string]$TenOrg = $global:o365_TenOrgDefault,
        [Parameter(Mandatory = $false, HelpMessage = "Use specific Credentials (defaults to Tenant-defined SvcAccount)[-Credentials [credential object]]")]
            [System.Management.Automation.PSCredential]$Credential,
        [Parameter(Mandatory = $false, HelpMessage = "Credential User Role spec (SID|CSID|UID|B2BI|CSVC|ESVC|LSVC|ESvcCBA|CSvcCBA|SIDCBA)[-UserRole @('SIDCBA','SID','CSVC')]")]
            # sourced from get-admincred():#182: $targetRoles = 'SID', 'CSID', 'ESVC','CSVC','UID','ESvcCBA','CSvcCBA','SIDCBA' ;
            #[ValidateSet("SID","CSID","UID","B2BI","CSVC","ESVC","LSVC","ESvcCBA","CSvcCBA","SIDCBA")]
            # pulling the pattern from global vari w friendly err
            [ValidateScript({
                if(-not $rgxPermittedUserRoles){$rgxPermittedUserRoles = '(SID|CSID|UID|B2BI|CSVC|ESVC|LSVC|ESvcCBA|CSvcCBA|SIDCBA)'} ;
                if(-not ($_ -match $rgxPermittedUserRoles)){throw "'$($_)' doesn't match `$rgxPermittedUserRoles:`n$($rgxPermittedUserRoles.tostring())" ; } ;
                return $true ;
            })]
            [string[]]$UserRole = @('ESvcCBA','CSvcCBA','SIDCBA'),
            #@('SID','CSVC'),
            # svcAcct use: @('ESvcCBA','CSvcCBA','SIDCBA')
        [Parameter(HelpMessage="Use EXOv2 (ExchangeOnlineManagement) over basic auth legacy connection [-useEXOv2]")]
            [switch] $useEXOv2=$true,
        [Parameter(HelpMessage="Silent output (suppress status echos)[-silent]")]
            [switch] $silent
    )
    BEGIN {
        ${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name ;
        $Verbose = ($VerbosePreference -eq 'Continue') ;
        
        <# recycling the inbound above into next call in the chain
        $pltRXO = [ordered]@{
            Credential = $Credential ; 
            verbose = $($VerbosePreference -eq "Continue")  ; 
            silent = $silent ; 
        } ;
        # default connectivity cmds - force silent false
        $pltRXOC = [ordered]@{} ; $pltRXO.GetEnumerator() | ?{ $_.Key -notmatch 'silent' }  | ForEach-Object { $pltRXOC.Add($_.Key, $_.Value) } ; $pltRXOC.Add('silent',$true) ;
        if((gcm Reconnect-EXO).Parameters.keys -notcontains 'silent'){ $pltRxo.remove('Silent') } ; 
        #>
        # 9:26 AM 6/17/2024 this needs cred resolution splice over latest get-exomailboxlicenses
        $o365Cred = $null ;
        if($Credential){
            $smsg = "`Credential:Explicit credentials specified, deferring to use..." ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                # get-TenantCredentials() return format: (emulating)
                $o365Cred = [ordered]@{
                Cred=$Credential ;
                credType=$null ;
            } ;
            $uRoleReturn = resolve-UserNameToUserRole -UserName $Credential.username -verbose:$($VerbosePreference -eq "Continue") ; # Username
            #$uRoleReturn = resolve-UserNameToUserRole -Credential $Credential -verbose = $($VerbosePreference -eq "Continue") ;   # full Credential support
            if($uRoleReturn.UserRole){
                $o365Cred.credType = $uRoleReturn.UserRole ;
            } else {
                $smsg = "Unable to resolve `$credential.username ($($credential.username))"
                $smsg += "`nto a usable 'UserRole' spec!" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                throw $smsg ;
                Break ;
            } ;
        } else {
            $pltGTCred=@{TenOrg=$TenOrg ; UserRole=$null; verbose=$($verbose)} ;
            if($UserRole){
                $smsg = "(`$UserRole specified:$($UserRole -join ','))" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                $pltGTCred.UserRole = $UserRole;
            } else {
                $smsg = "(No `$UserRole found, defaulting to:'CSVC','SID' " ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                $pltGTCred.UserRole = 'CSVC','SID' ;
            } ;
            $smsg = "get-TenantCredentials w`n$(($pltGTCred|out-string).trim())" ;
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level verbose }
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
            $o365Cred = get-TenantCredentials @pltGTCred
        } ;
        if($o365Cred.credType -AND $o365Cred.Cred -AND $o365Cred.Cred.gettype().fullname -eq 'System.Management.Automation.PSCredential'){
            $smsg = "(validated `$o365Cred contains .credType:$($o365Cred.credType) & `$o365Cred.Cred.username:$($o365Cred.Cred.username)" ;
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
            # 9:58 AM 6/13/2024 populate $credential with return, if not populated (may be required for follow-on calls that pass common $Credentials through)
            if((gv Credential) -AND $Credential -eq $null){
                $credential = $o365Cred.Cred ;
            }elseif($credential.gettype().fullname -eq 'System.Management.Automation.PSCredential'){
                $smsg = "(`$Credential is properly populated; explicit -Credential was in initial call)" ; 
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            } else {
                $smsg = "`$Credential is `$NULL, AND $o365Cred.Cred is unusable to populate!" ;
                $smsg = "downstream commands will *not* properly pass through usable credentials!" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                throw $smsg ;
                break ;
            } ;
        } else {
            $smsg = "UNABLE TO RESOLVE FUNCTIONAL CredType/UserRole from specified explicit -Credential:$($Credential.username)!" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            break ;
        } ; 

        # downstream commands
        $pltRXO = [ordered]@{
            Credential = $Credential ;
            verbose = $($VerbosePreference -eq "Continue")  ;
        } ;
        if((get-command Reconnect-EXO).Parameters.keys -contains 'silent'){
            $pltRxo.add('Silent',$silent) ;
        } ;
        # default connectivity cmds - force silent false
        $pltRXOC = [ordered]@{} ; $pltRXO.GetEnumerator() | ?{ $_.Key -notmatch 'silent' }  | ForEach-Object { $pltRXOC.Add($_.Key, $_.Value) } ; $pltRXOC.Add('silent',$true) ; 
        if((get-command Reconnect-EXO).Parameters.keys -notcontains 'silent'){
            $pltRxo.remove('Silent') ;
        } ; 

        #$ExMbxLicenses = get-ExoMailboxLicenses -verbose:$($VerbosePreference -eq "Continue")  ;
        # add outdetail & unfiltered support
        $ExMbxLicenses = get-ExoMailboxLicenses -Unfiltered -OutDetail -credential $pltRXO.Credential -verbose:$($VerbosePreference -eq "Continue")  ;
        # pull the full Tenant list, for performing sku-> name conversions
        #$lplist =  get-AADlicensePlanList -verbose -IndexOnName ;

        $pltGLPList=[ordered]@{ 
            TenOrg= $TenOrg;
            IndexOnName=$false ;
            Credential = $pltRXO.Credential ; 
            verbose = $pltRXO.verbose  ; 
            silent = $false ; 
        } ; 
        $smsg = "get-AADlicensePlanList w`n$(($pltGLPList|out-string).trim())" ; 
        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 

        # this *does* require working AAD access logon. 
        $skus  = get-AADlicensePlanList @pltGLPList ;

        # check if using Pipeline input or explicit params:
        if ($PSCmdlet.MyInvocation.ExpectingInput) {
            write-verbose "Data received from pipeline input: '$($InputObject)'" ;
        } else {
            # doesn't actually return an obj in the echo
            write-verbose "Data received from parameter input:" # '$($InputObject)'" ;
        } ;
    } 
    PROCESS {
        if($ExMbxLicenses){
            $IsExoLicensed = $false ;
            foreach($pLic in $User.AssignedLicenses.skuid){
                $smsg = "(resolving $($plic) for EXO support)" ; 
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                if($tsku = $skus[$pLic]){
                    if($xlic = $ExMbxLicenses[$tsku.SkuPartNumber]){
                        $IsExoLicensed = $true ;
                        $smsg = "$($User.userprincipalname) HAS EXO UserMailbox-supporting License:$($xlic.SKU)|$($xlic.Label)|$($tsku.skuid)" ;
                        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                        break ; 
                    } ;
                } else { 
                    $smsg = "($($plic):mbx support no match)" ; 
                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                } ; 
            } ;
        } else { 
            $smsg = "Unable to resolve get-ExoMailboxLicenses!" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } 
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
            throw $smsg ;
            Break ; 
        } ; 

    }  # PROC-E
    END{
        $IsExoLicensed | write-output ; 
    } ;
}

#*------^ test-EXOIsLicensed.ps1 ^------