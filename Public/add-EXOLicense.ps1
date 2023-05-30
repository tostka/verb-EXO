﻿# add-EXOLicense.ps1

#*------v add-EXOLicense.ps1 v------
function add-EXOLicense {
    <#
    .SYNOPSIS
    remove-EXOLicense.ps1 - Add a temporary o365 license from specified AADUser account. Returns updated MSOLUser object to pipeline.
    .NOTES
    Version     : 1.0.0.
    Author      : Todd Kadrie
    Website     : http://www.toddomation.com
    Twitter     : @tostka / http://twitter.com/tostka
    CreatedDate : 2021-07-13
    FileName    : add-EXOLicense.ps1
    License     : MIT License
    Copyright   : (c) 2021 Todd Kadrie
    Github      : https://github.com/tostka/verb-XXX
    Tags        : Powershell, ExchangeOnline, AzureAD, License
    AddedCredit : REFERENCE
    AddedWebsite: URL
    AddedTwitter: URL
    REVISIONS
    * 3:14 PM 5/30/2023 updated CBH; udpt CBH; consold 222+223 into 1 line; add pswl compliance; expanded demos ; rem'd unused
    * 3:52 PM 5/23/2023 implemented @rxo @rxoc split, (silence all connectivity, non-silent feedback of functions); flipped all r|cxo to @pltrxoC, and left all function calls as @pltrxo; generic'd the meta vari name ; general cleanup rem'd; added expanded licname to echo ; 
    * 4:11 PM 5/22/2023 flipped all lic status testing to use of test-exoislicensed ; logic fixes
    * 9:49 AM 5/19/2023 trimmed rem's; ++ adv func/pipeline supp ; shifted usr reso into thge process loop ; rem'd unused $TenantShortName; wrapped plts ; 
        rem'd END transcript stop - these are util functions: connectivity; transcription & control of logging should occur in the calling script/func, not the stripped down leaf function.
    * 3:37 PM 5/17/2023 added pltRXO support; replaced hard-coded LicenseSkuIds w dyn LicenseSkuKeys pulled from global Meta. Updated UserRole validator to modern; 
        added AADUser detect, deprected MSolUser; stripped out all logging & fancy connectivity, the pltRXO support provides enough to chain through existing creds ; 
        removed dangling xow support
    * 4:01 PM 4/19/2023 roughed in, untested EOM310 updates: pasted in generic services block, sub'd -exo -> -xo. No further testing.
    * 2:54 PM 12/21/2022 tested through non-debug of shared, no-add lic mbx ;  more recent retooling for EXOv2/MFA support/Loss of MSONLINE/MSOL module support/cmdlets around AADU status and licensure.
    * 2:29 PM 8/12/2022 sync'd back to last _func.ps1 chgs as well ; fixed inacc warning, when lic's all burned (was echo'ing failure to update usageloc, not lic fail).
    * 5:17 PM 3/23/2022 more retooling to remove msonline module dependance, and shift to AzureAD (crappy implementation GraphAPI) module
    * 1:50 PM 3/23/2022 hunting the VerbosePreference toggle midway through, found 2 more verbose tests lacking leading verbose = $($VerbosePreference -eq "Continue"); prefixed examples with PS>
    * 5:00 PM 3/22/2022 extensive rewrite: Sec mandate to disable all basic auth == complete loss of the long-standing MS MSOnline powershell module:
        net effect: have to reimplement & rewraite all verb-MsolNoun cmdlet calls into
        the new AzureAD module's equivelents (which fail to match msol cmdlets names,
        parameters, or even the data returned, and property names)
        - had to write 3 new functions, ground up, to reimplement loss of the 1-liner Set-MsolUserLicense cmdlet functions:
        - wrote verb-aad: add-AADUserLicense()
        - wrote verb-aad: remove-AADUserLicense()
        - wrote verb-aad: set-AADUserUsageLocation()
        - wrote verb-aad: get-AADlicensePlanList, to workaround loss of useful sku reporting from the prior equiv msol sku cmdlet (new output is unformatted json [facepalm])
        - rewrote most of the license testing & handling code in this verb-exo:Add-EXOLicense() (roughly 11:20am 3/21/2022 to 5:03 PM 3/22/2022, and I still have a verbose state bug to workout on this script).
    * 11:51 AM 3/2 1/2022 update: because *any* licenes, including worthless FLOWFREE, toggles IsLicensed:$true, logic below fails to detect the lack of an EXO lic.
    Have to splice over from get-mailboxuserStatus, that evaluates existing aaduser/msolu licenses against the ones that actually support a UserMailbox type.
    * 12:57 PM 1/31/2022 addded -ea 0 to gv PassStatus_$($tenorg) (spurious error suppress)
    * 2:14 PM 1/18/2022 updated Example 1 to include echo of the returned msolu.licenses value.
    * 12:08 PM 1/11/2022 ren add-EXOLicenseTemp -> add-EXOLicense ; add
    $XXXMETA.o365LicSkuExStd == EXCHANGESTANDARD (Office 365 Exchange Online Only
    ,commonly used for App Access) & stick in front of $LicenseSkuIds,
    $XXXMETA.o365LicSkuExStd; added examples with explicit cmdlines for the adds;
    spliced over UsageLocation test/assert code from add-o365license.
    * 1:34 PM 1/5/2022 init
    .DESCRIPTION
    remove-EXOLicense.ps1 - Add a temporary o365 license from specified AADUser account. Returns updated MSOLUser object to pipeline.
    .PARAMETER Ticket
    Ticket Number [-Ticket '999999']
    .PARAMETER TenOrg
    TenantTag value, indicating Tenants to connect to[-TenOrg 'TOL']
    .PARAMETER  users
    Array of UserPrincipalNames (or MSOLUser objects) to have a temporary Exchange License applied
    .PARAMETER LicenseSkuKeys
    Array, in preference order, of XXXMeta global value LicenseSkuKey names (resolves SKUId from TenOrg global Meta vari ; first working lic assignment, will be applied)[-LicenseSkuIds 'o365LicSkuExStd','o365LicSkuF1']
    .PARAMETER LicenseSkuIds
    Optional Array, in preference order, of LicenseSkuID (e.g. TenantName:SPE_F1) to be added, runs list until first sucess (default process is to dynamically resolve id's from Meta LicenseSkuKeys specifications)[-LicenseSkuIds @(`$XXXMETA.o365LicSkuExStd,`$XXXMETA.o365LicSkuF1)]
    .PARAMETER Force
    switch to override normal 'skipped' license application to existing Mailbox (needed for licensed-Shared, or upgraded existing lic).
    .PARAMETER UserRole
    Credential User Role spec (SID|CSID|UID|B2BI|CSVC)[-UserRole SID]
    .PARAMETER Credential
    Use specific Credentials (defaults to Tenant-defined SvcAccount)[-Credentials [credential object]]
    .PARAMETER silent
    Switch to specify suppression of all but warn/error echos.(unimplemented, here for cross-compat)
    .PARAMETER useEXOv2
    Use EXOv2 (ExchangeOnlineManagement) over basic auth legacy connection [-useEXOv2]
    .PARAMETER outObject
    switch to return a system.object summary to the pipeline[-outObject]
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    Microsoft.Online.Administration.User
    Returns updated MSOLUser object to pipeline
    .EXAMPLE
    PS> add-EXOLicense -users 'Test@domain.com','Test2@domain.com' -Ticket 999999 -Credential $pltrxo.Credential ; 
    Process an array of users, with default 'hunting' -LicenseSkuIds array.
    .EXAMPLE
    PS> $whatif = $false ;
    PS> $target = '999998,TestSharedMbxConversion@toro.com' ;
    PS> pushd;
    PS> $prpADU1 = 'UserPrincipalName','DisplayName',@{Name='IsLicensed'; Expression={[boolean]($_.AssignedLicenses.count -gt 0) }}  ;
    PS> $prpADU2 = @{Name='Licenses';Expression={($_ | Get-AzureADUserLicenseDetail).SkuPartNumber -join ','}} ;
    PS> if($target.contains(',')){
    PS>     $ticket,$trcp = $target.split(',') ;
    PS>     $pltAxLic = [ordered]@{
    PS>         users = $trcp ;
    PS>         ticket = $ticket ;
    PS>         whatif = $($whatif) ;
    PS>         Verbose = $false ;
    PS>         Credential  =  $credO365TORSIDCBA ;
    PS>         silent = $false ;
    PS>     } ;
    PS>     $smsg = "add-EXOLicense w`n$(($pltAxLic|out-string).trim())" ;
    PS>     if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
    PS>     $updatedAadu = add-EXOLicense @pltAxLic ;
    PS>     write-verbose "re -refresh AADU" ;
    PS>     $updatedAadu  = get-AzureAdUser -obj $updatedAadu.UserPrincipalName ;
    PS>     $smsg = "UpdatedAadu: w`n$(($updatedAadu| ft -auto $prpADU1 |out-string).trim())" ;
    PS>     $smsg += "`n:$(($updatedAadu| fl $prpADU2 |out-string).trim())" ;
    PS>     write-host -foregroundcolor green $smsg ;
    PS> } else { write-warning "`$target does *not* contain comma delimited ticket,UPN string!"} ;    
    Fancier variant of above, with more post-confirm reporting
    .EXAMPLE
    PS> add-EXOLicense -users 'Test@domain.com' -LicenseSkuIds $XXXMETA.o365LicSkuExStd -ticket TICKETNUMBER;
    add an explicitly specified lic to a user (in this case, using the LicenseSku for EXCHANGESTANDARD, as stored in a global variable)
    .EXAMPLE
    PS> add-EXOLicense -users 'Test@domain.com' -LicenseSkuIds $XXXMETA.o365LicSkuF1 -ticket TICKETNUMBER;
    add an explicitly specified lic to a user (in this case, using the LicenseSku for SPE_F1 - web-only o365 - lic as stored in a global variable)
    .EXAMPLE
    PS> add-EXOLicense -users 'Test@domain.com' -LicenseSkuIds $XXXMETA.o365LicSkuE3 -ticket TICKETNUMBER ;
    add an explicitly specified lic to a user (in this case, using the LicenseSku for ENTERPRISEPACK - E3 o365 - lic as stored in a global variable)
    .EXAMPLE
    PS> add-EXOLicense -users 'Test@domain.com' -LicenseSkuIds 'TENANTNAME:EXCHANGESTANDARD' -ticket TICKETNUMBER ;
    add an explicitly specified lic to a user by specifying the Tenant-specific LicenseSkuID directly
    .EXAMPLE
    PS> add-EXOLicense -users 'Test@domain.com' -LicenseSkuIds @('TENANTNAME:EXCHANGESTANDARD',$XXXMETA.o365LicSkuF1) -ticket TICKETNUMBER ;
    Explicitly specify a preference order array of Tenant-specific LicenseSkuIDs (one string, another pulleed from Meta global vari; attempted in order until first success)
    .EXAMPLE
    PS> add-o365License -users $Aaduser.UserprincipalName -ticket TICKETNUMBER ;
    add-o365License compatibility option
    .LINK
    https://github.com/tostka/verb-exo
    #>
    # migr to verb-exo, pull the dupe spec...
    #Requires -Modules AzureAD, MSOnline, ExchangeOnlineManagement, verb-AAD, verb-Auth, verb-IO, verb-logging, verb-Mods, verb-Text
    #Requires -RunasAdministrator
    # VALIDATORS: [ValidateNotNull()][ValidateNotNullOrEmpty()][ValidateLength(24,25)][ValidateLength(5)][ValidatePattern("some\s\regex\sexpr")][ValidateSet("USEA","GBMK","AUSYD")][ValidateScript({Test-Path $_ -PathType 'Container'})][ValidateScript({Test-Path $_})][ValidateRange(21,65)][ValidateCount(1,3)]
    ## [OutputType('bool')] # optional specified output type
    [CmdletBinding()]
    #[Alias('add-o365License')]
    PARAM(
        [Parameter(Position=0,Mandatory=$False,ValueFromPipeline=$true,HelpMessage="Array of UserPrincipalNames (or AzureADUser objects) to have a temporary Exchange License applied")]
            #[ValidateNotNullOrEmpty()]
            #[Alias('ALIAS1', 'ALIAS2')]
            #[ValidatePattern("^([0-9a-zA-Z]+[-._+&'])*[0-9a-zA-Z]+@([-0-9a-zA-Z]+[.])+[a-zA-Z]{2,63}$")]
            [array]$users,
        [Parameter(Mandatory=$True,HelpMessage="Ticket Number [-Ticket '999999']")]
            [string]$Ticket,
        [Parameter(,HelpMessage="Array, in preference order, of XXXMeta global value LicenseSkuKey names (resolves SKUId from TenOrg global Meta vari ; first working lic assignment, will be applied)[-LicenseSkuIds 'o365LicSkuExStd','o365LicSkuF1']")]
            [ValidateNotNullOrEmpty()]
            [array]$LicenseSkuKeys=@('o365LicSkuExStd','o365LicSkuF1','o365LicSkuE3'),
        [Parameter(,HelpMessage="Optional Array, in preference order, of LicenseSkuID (e.g. TenantName:SPE_F1) to be added, runs list until first sucess (default process is to dynamically resolve id's from Meta LicenseSkuKeys specifications)[-LicenseSkuIds @(`$XXXMETA.o365LicSkuExStd,`$XXXMETA.o365LicSkuF1)]")]
            #[ValidateNotNullOrEmpty()]
            [array]$LicenseSkuIds = @(), 
        [Parameter(HelpMessage="switch to override normal 'skipped' license application to existing Mailbox (needed for licensed-Shared, or upgraded existing lic).[-Force]")]
            [switch] $Force,
        [Parameter(Mandatory=$FALSE,HelpMessage="TenantTag value, indicating Tenants to connect to[-TenOrg 'TOL']")]
            [ValidateNotNullOrEmpty()]
            #[ValidatePattern("^\w{3}$")]
            [string]$TenOrg = $global:o365_TenOrgDefault,
        [Parameter(HelpMessage="Use EXOv2 (ExchangeOnlineManagement) over basic auth legacy connection [-useEXOv2]")]
            [switch] $useEXOv2=$true,
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
            [string[]]$UserRole = @('SID','CSVC'),
        [Parameter(HelpMessage="Silent output (suppress status echos)[-silent]")]
            [switch] $silent,
        [Parameter(HelpMessage="switch to show extended debugging output [-showdebug]")]
            # included solely for backward compatibility with add-o365License()
            [switch] $showDebug,
        [Parameter(HelpMessage="Whatif Flag  [-whatIf]")]
            [switch] $whatIf
    ) ;
    BEGIN{
        #region CONSTANTS-AND-ENVIRO #*======v CONSTANTS-AND-ENVIRO v======
        # function self-name (equiv to script's: $MyInvocation.MyCommand.Path) ;
        ${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name ;
        $Verbose = ($VerbosePreference -eq 'Continue') ;
        # Get parameters this function was invoked with
        $PSParameters = New-Object -TypeName PSObject -Property $PSBoundParameters ;
        $smsg = "(ParameterSetName $($PSCmdlet.ParameterSetName) is in effect)" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

        if(-not $rgxEmailAddr){ $rgxEmailAddr = "^([0-9a-zA-Z]+[-._+&'])*[0-9a-zA-Z]+@([-0-9a-zA-Z]+[.])+[a-zA-Z]{2,63}$"}

        # recycling the inbound above into next call in the chain
        # downstream commands
        $pltRXO = [ordered]@{
            Credential = $Credential ;
            verbose = $($VerbosePreference -eq "Continue")  ;
        } ;
        if((gcm Reconnect-EXO).Parameters.keys -contains 'silent'){
            $pltRxo.add('Silent',$silent) ;
        } ;
        # default connectivity cmds - force silent false
        $pltRXOC = [ordered]@{} ; $pltRXO.GetEnumerator() | ?{ $_.Key -notmatch 'silent' }  | ForEach-Object { $pltRXOC.Add($_.Key, $_.Value) } ; $pltRXOC.Add('silent',$true) ;
        if((gcm Reconnect-EXO).Parameters.keys -notcontains 'silent'){ $pltRxo.remove('Silent') } ; 

        #[array]$LicenseSkuIds = @() ; # moved to param , to permit direct lic spec when using indep of formal scripts
        if(-not $LicenseSkuIds){
            $smsg = "Retrieve & build LicenseSkuIDS from global Meta vari" ;  
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            $LicenseSkuKeys | foreach-object { $LicenseSkuIds += @((get-variable -name "$($tenorg)META").value[$_]) } ; 
        } else { 
            $smsg = "Explicit -LicenseSkuIds specified, using those licenses (in preference order)" ; 
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            $LicenseSkuKeys = $LicenseSkuIds
        } ;

        $smsg = $sBnr="`n#*======v $(${CmdletName}) : v======" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

        $admin = "$env:username" ;

        # check if using Pipeline input or explicit params:
        if ($PSCmdlet.MyInvocation.ExpectingInput) {
            $smsg = "Data received from pipeline input: '$($InputObject)'" ;
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        } else {
            # doesn't actually return an obj in the echo
            #$smsg = "Data received from parameter input: '$($InputObject)'" ;
            #if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
            #else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        } ;

    } ;  # BEGIN-E
    PROCESS{
        
        $ttl = ($users|measure).count ; $Procd=0 ;
        [array]$Rpt =@() ;
        
        foreach ($usr in $users){

            switch($usr.GetType().FullName){
                'Microsoft.Online.Administration.User' {
                    #$smsg = "(-user:MsolU detected:$($usr.userprincipalname), extracting the UPN...)" ;
                    $smsg = "MSOLUSER OBJECT IS NO LONGER SUPPORTED BY THIS FUNCTION!" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    $usr = $usr.userprincipalname ;
                } ;
                'Microsoft.Open.AzureAD.Model.User' {
                    $smsg = "(-user:AzureADU detected)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    $usr = $usr.userprincipalname ;
                } ;
                'System.String'{
                    $smsg = "(-user:string detected)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                    if($usr -match $rgxEmailAddress){

                        $smsg = "(-user:EmailAddress/UPN detected:$($usr))" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        $usr = $usr ;
                    } else {
                        $smsg = "-Users: Unable to recognize either an AzureAD user object, or a UPN string, from the specified input:`n$($usr)" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                        break ; 
                    } ; 
                }
                default{
                    $smsg = "Unrecognized format for -User:$($usr)!. Please specify either a user UPN, or pass a full MsolUser object." ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    Break ;
                }
            } ;
            
            $tUPN=$usr ;
            #$LicenseSkuIds=$XXXMETA.o365LicSkuF1; # dyn from global XXXmeta
            $error.clear() ;
            TRY {

                $Exit = 0 ;
                Do {
                    Try {
                        connect-aad @pltRXOC ; 
                        $AADUser=$null ;
                        #$TenantShortName = ((Get-AzureADTenantDetail).verifieddomains |?{$_._default}).name.split('.')[0] ;
                        $pltGAADU=[ordered]@{ 
                            ObjectId = $tUPN ;
                            ErrorAction = 'STOP' ;
                            verbose = $($VerbosePreference -eq "Continue") ;
                        } ;
                        $AADUser = Get-AzureADUser @pltGAADU ;
                        $Exit = $Retries ;
                    } CATCH {
                        Start-Sleep -Seconds $RetrySleep ;
                        $Exit ++ ;
                        $smsg = "Failed to exec cmd because: $($Error[0])" ;
                        $smsg += "`nTry #: $Exit" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                        else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        If ($Exit -eq $Retries) {
                            $smsg =  "Unable to exec cmd!" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                            else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        } ;
                        Continue ;
                    }  ;
                } Until ($Exit -eq $Retries) ;

                # confirm/set UsageLoc (reqd for updates)
                if (-not $AADUser.UsageLocation) {
                    $smsg = "AADUser: MISSING USAGELOCATION, FORCING" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    $spltSAADUUL = [ordered]@{
                        Users = $AADUser.UserPrincipalName ;
                        UsageLocation = "US" ;
                        whatif = $($whatif) ;
                        Credential = $pltRXO.Credential ;
                        verbose = $pltRXO.verbose  ;
                        silent = $false ;
                    } ;
                    $smsg = "set-AADUserUsageLocationw`n$(($spltSAADUUL|out-string).trim())" ;
                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                    $bRet = set-AADUserUsageLocation @spltSAADUUL ;
                    if($bRet.Success){
                        $smsg = "set-AADUserUsageLocation updated UsageLocation:$($bRet.AzureADuser.UsageLocation)" ;
                        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                        # update the local AADUser to reflect the updated AADU returned
                        $AADUser = $bRet.AzureADuser ;
                        #$Report.FixedUsageLocation = $true ;
                    } else {
                        $smsg = "set-AADUserUsageLocation: FAILED TO UPDATE USAGELOCATION!" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        #$Report.FixedUsageLocation = $false ;
                        if(-not $whatif){
                            BREAK;
                        }
                    } ;
                } ;
                #

                # if lic'd and has a mailbox, shouldn't need a new license
                # w shift from msol -> aad, $aadU doesn 't even *have* an islicensed property! Have to interpolate:
                # nope!: IsLicensed:true, even if nothing but FLOWFREE is set. Worthless, for determining why there's no mailbox.
                # have to splice over the full exolic-testing code from verb-ex2010:get-mailboxUserStatus():

                # 8:44 AM 12/21/2022 no, use the verb-EXO:test-EXOIsLicensed(): test-EXOIsLicensed -User $AADUser -verbose
                $IsExoLicensed = test-EXOIsLicensed -User $AADUser -Credential:$pltRXO.Credential -verbose:$pltRXO.verbose -silent:$pltRXO.silent ;
                $pltGLPList=[ordered]@{ 
                    TenOrg= $TenOrg;
                    verbose=$($VerbosePreference -eq "Continue") ;
                    credential= $pltRXO.credential ;
                    silent = $false ; 
                } ;
                $smsg = "$($tenorg):get-AADlicensePlanList wn$(($pltGLPList|out-string).trim())" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                $skus = $null ;
                $skus = get-AADlicensePlanList @pltGLPList ;

                $ombx = get-xomailbox -id $AADUser.UserPrincipalName -ea continue  ;
                $ombx = $ombx | ?{$_ -is [System.Management.Automation.PSObject]} # filtering any aberrant obj returned (legacy of prior problematic xow reliance to work around hybrid stepable pipeline bug)
                $AADLicDetails = get-AADUserLicenseDetails -UPNs $AADUser.userprincipalname -Credential:$pltRXO.Credential -verbose:$pltRXO.verbose -silent:$pltRXO.silent ; 
                $smsg = "`nExisting Mbx:`n$(($ombx | ft -a 'RecipientType','RecipientTypeDetails'|out-string).trim())" ;
                if($AADLicDetails){
                    $smsg += "`n`$AADLicDetails`n$(($AADLicDetails|out-string).trim())" ;
                } else { 
                    $smsg += "`n`$AADLicDetails:(empty return)" ;
                } ; 
                if($ombx.RecipientTypeDetails -eq 'SharedMailbox'){
                    $smsg += "`nSharedMailbox does not *require* a license" ;
                } ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                #if( -not($Force) -AND $ombx){
                <#if( -not($Force) -AND ($ombx.RecipientTypeDetails -eq 'SharedMailbox') ){
                    $smsg += "`n -- SKIPPING EXO-RELATED LICENSE ADD! --" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                } else
                #>
                if( (-not $IsExoLicensed) -OR ($Force) ){
                    # not supported on aadu: defer to: verb-AAD:test-AADUserIsLicensed(): $isLicensed = test-AADUserIsLicensed -user $AzureADUser -verbose
                    if($IsAADIsLicensed = test-AADUserIsLicensed -user $AADUser -Verbose:($VerbosePreference -eq 'Continue')){
                        # has a bozo lic that doesn't support a mailbox
                        $smsg = "AADUser:$($tUPN):  isLicensed (has some form of license added), but has *NO* EXO UserMailbox-supporting license!" ;
                        $smsg += "`n(or is being -Force upgraded to an elevated license)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN }
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ;
                    $smsg="confirmed $($AADUser.UserPrincipalName):is unlicensed/underlicensed" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                    # 9:55 AM 11/15/2019 per BB: apply a license, and notify lic staff to record
                    
                    # Set-AzureADUserLicense  vers
                    $pltALic=[ordered]@{
                        UserPrincipalName=$AADUser.userprincipalname ;
                         AssignedLicenses=$null ;
                    } ;

                    foreach($LicenseSkuId in $LicenseSkuIds){
                        if( $LicenseSkuId.contains(':') ){
                            $LicenseSkuId = $LicenseSkuId.split(':')[1] ;
                            # need the skuid, not the name, could pull another licplan list indexedonName, but can also post-filter the hashtable, and get it.
                            $LicenseSkuId = ($skus.values | ?{$_.SkuPartNumber -eq $LicenseSkuId}).skuid ;
                        } ;
                        #$smsg = "(attempting license:$($LicenseSkuId)...)" ;
                        $smsg = "(attempting license:$(($skus.values | ?{$_.Skuid -eq $LicenseSkuId}).SkuPartNumber):$($LicenseSkuId)...)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                        $pltAAADUL=[ordered]@{
                            Users=$AADUser.UserPrincipalName ;
                            skuid=$LicenseSkuId ;
                            Credential = $pltRXO.Credential ; 
                            verbose = $pltRXO.verbose  ; 
                            silent = $false ; 
                            erroraction = 'STOP' ;
                            whatif = $($whatif) ;
                        } ;
                        $smsg = "add-AADUserLicense w`n$(($pltAAADUL|out-string).trim())" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                        $bRet = add-AADUserLicense @pltAAADUL ;
                        if($bRet.Success){
                            $smsg = "add-AADUserLicense added  Licenses:$($bRet.AddedLicense)" ;
                            # $AADUser.AssignedLicenses.skuid
                            $smsg += "`n$(($AADUser.AssignedLicenses.skuid|out-string).trim())" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                            $smsg = "Detailed Return:`n$(($bRet|out-string).trim())" ;
                            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                            BREAK ; # abort further loops if one successfully applied
                        } elseif($whatif){
                            $smsg = "(whatif pass, exec skipped), " ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        } elseif( -not [boolean]($bRet.AddedLicenses)){
                            # failed add
                            $smsg = "Failed Lic Add:$($LicenseSkuId) (exhausted units?, moving on to next if avail...)" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN }
                            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            Continue ;
                        } else {
                            $smsg = "add-AADUserLicense : UNAVAIL LIC UNIT, OR FAILED TO UPDATE USAGELOCATION!" ;
                            $smsg += "`n$(($bRet|out-string).trim())" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            #$Report.FixedUsageLocation = $false ;
                            if(-not $whatif){
                                BREAK;
                            }
                        } ;

                    } ;  # loop-E $LicenseSkuIds

                };  # if-E $ombx
            } CATCH {
                $ErrTrapd=$Error[0] ;
                $smsg = "$('*'*5)`nFailed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: `n$(($ErrTrapd|out-string).trim())`n$('-'*5)" ;
                $smsg += "`n$($ErrTrapd.Exception.Message)" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                BREAK ;
            } ; 
            if(!$whatif){
                $smsg = "dawdling until License reinflates mbx..." ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                $1F=$false ;
                Do {
                    if($1F){Sleep -s 30} ;
                    write-host "." -NoNewLine ;
                    $1F=$true ;
                } Until ($ombx = get-xomailbox -id $AADUser.userprincipalname -EA 0) ; # capture return (prevent from dropping into pipe)
                # get-xomailbox returns: System.Management.Automation.PSObject; not a real Mailbox object class
                $ombx = $ombx | ?{$_ -is [System.Management.Automation.PSObject]} ; # looks like an attempt to filter just the mailbox out of the pipeline return
                $smsg = "xo Mailbox confirmed!" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } ;

            # return $AADUser to pipeline if populated
            $smsg = "refresh updated AADUser:" ; 
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 

            $pltGAADU.ObjectId = $AADUser.UserPrincipalName ;
            TRY {
                $AADUser = Get-AzureADUser @pltGAADU ;
            } CATCH {
                $ErrTrapd=$Error[0] ;
                $smsg = "$('*'*5)`nFailed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: `n$(($ErrTrapd|out-string).trim())`n$('-'*5)" ;
                $smsg += "`n$($ErrTrapd.Exception.Message)" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                BREAK ;
            } ; 

            $smsg = "Return updated AADUser to pipeline" ; 
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
            $AADUser | write-output ;

            $smsg =  $sBnr.replace('=v','=^').replace('v=','^=') ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        } ; # ($usr in $users)

    } # PROC-E
    END{

    } ;
 } ; 

#*------^ add-EXOLicense.ps1 ^------