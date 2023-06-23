#*------v remove-EXOLicense.ps1 v------
function remove-EXOLicense {
    <#
    .SYNOPSIS
    remove-EXOLicense.ps1 - Remove a temporary o365 license from specified MsolUser account. Returns updated MSOLUser object to pipeline.
    .NOTES
    Version     : 1.0.0.
    Author      : Todd Kadrie
    Website     : http://www.toddomation.com
    Twitter     : @tostka / http://twitter.com/tostka
    CreatedDate : 2021-07-13
    FileName    : remove-EXOLicense.ps1
    License     : MIT License
    Copyright   : (c) 2021 Todd Kadrie
    Github      : https://github.com/tostka/verb-XXX
    Tags        : Powershell
    AddedCredit : REFERENCE
    AddedWebsite: URL
    AddedTwitter: URL
    REVISIONS
    * 4:25 PM 5/22/2023 scrapped, rewrote adapted from functional add-exoLicense(); 
     * 3:37 PM 5/17/2023 added pltRXO support
    * 12:10 PM 6/7/2022 updated cbh params
    * 5:17 PM 3/23/2022 retooling to remove msonline module dependance, and shift to AzureAD (crappy implementation GraphAPI) module
    * 12:57 PM 1/31/2022 addded -ea 0 to gv PassStatus_$($tenorg) (spurious error suppress)
    * 3:14 PM 1/18/2022 REM'D EXOP conn's (match add-exolic), this is pure msolu & exo. 
    * 1:02 PM 1/17/2022 port of add-EXOLicense to removal process
    .DESCRIPTION
    remove-EXOLicense.ps1 - Remove a temporary o365 license from specified MsolUser account. Returns updated MSOLUser object to pipeline.
    .PARAMETER users
    Array of UserPrincipalNames (or MSOLUser objects) to have a temporary Exchange License applied
    .PARAMETER Ticket
    Ticket Number [-Ticket '999999']
    .PARAMETER LicenseSkuKeys
    Array, in preference order, of XXXMeta global value LicenseSkuKey names (resolves SKUId from TenOrg global Meta vari ; runs full list removing all matches)[-LicenseSkuIds 'o365LicSkuExStd','o365LicSkuF1']
    .PARAMETER LicenseSkuIds
    Optional Array, in preference order, of LicenseSkuID (e.g. TenantName:SPE_F1) to be added, runs full list removing all matches (default process is to dynamically resolve id's from Meta LicenseSkuKeys specifications)[-LicenseSkuIds @(`$XXXMETA.o365LicSkuExStd,`$XXXMETA.o365LicSkuF1)]
    .PARAMETER Force
    switch to override normal 'skipped' license application to existing Mailbox (needed for licensed-Shared, or upgraded existing lic).
    .PARAMETER TenOrg
    TenantTag value, indicating Tenants to connect to[-TenOrg 'TOL']
    .PARAMETER useEXOv2
    Use EXOv2 (ExchangeOnlineManagement) over basic auth legacy connection [-useEXOv2]
    .PARAMETER Credential
    Use specific Credentials (defaults to Tenant-defined SvcAccount)[-Credentials [credential object]]
    .PARAMETER UserRole
    Credential User Role spec (SID|CSID|UID|B2BI|CSVC)[-UserRole SID]
    .PARAMETER Credential
    Use specific Credentials (defaults to Tenant-defined SvcAccount)[-Credentials [credential object]]
    .PARAMETER silent
    Switch to specify suppression of all but warn/error echos.(unimplemented, here for cross-compat)
    .PARAMETER showDebug
    switch to show extended debugging output [-showdebug]
    .PARAMETER whatIf
    Whatif Flag  [-whatIf]
    .INPUTS
    None. Does not accepted piped input.(.NET types, can add description)
    .OUTPUTS
    System.Object - returns summary report to pipeline
    .EXAMPLE
    PS> remove-EXOLicense -users 'Test@domain.com','Test2@domain.com' -ticket TICKETNUMBER -verbose  ;
    Process an array of users, with default 'hunting' -LicenseSkuIds array. 
    .EXAMPLE
    PS> $updatedMSOLU = remove-EXOLicense -users 'Test@domain.com','Test2@domain.com' -verbose -ticket TICKETNUMBER;
        if($updatedMSOLU.islicensed){'Y'} else { 'N'} ; 
    Process licnse removal for specified user, and post-test isLicensed status, using default license array configured on the -LicenseSkuIDs default value.
    .EXAMPLE
    PS> $whatif=$true ;
        $target = '99999,lynctest15@toro.com' ;
        if($target.contains(',')){
            $ticket,$trcp = $target.split(',') ;
            $updatedmsolu = remove-EXOLicense -users $trcp -Ticket $ticket -whatif:$($whatif) ;
            $props1 = 'UserPrincipalName','DisplayName','IsLicensed' ;
            $props2 = @{Name='Licenses';
            Expression={$_.licenses.accountskuid -join ', '}}  ;
            $smsg = "$((get-date).ToString('HH:mm:ss')):UpdatedMsolU: w`n$(($updatedmsolu| ft -auto $props1|out-string).trim())" ;
            $smsg += "`n$(($updatedmsolu| fl $props2 |out-string).trim())" ;
            write-host -foregroundcolor green $smsg ;
        } else { write-warning "`$target does *not* contain comma delimited ticket,UPN string!"} ;
    Fancier variant of above, with more post-confirm reporting
    .EXAMPLE
    PS> remove-EXOLicense -users 'Test@domain.com' -LicenseSkuIds $TORMETA.o365LicSkuExStd -ticket TICKETNUMBER;
    removal an explictly specified lic to a user (in this case, using the LicenseSku for EXCHANGESTANDARD, as stored in a global variable)
    .EXAMPLE
    PS> remove-EXOLicense -users 'Test@domain.com' -LicenseSkuIds $TORMETA.o365LicSkuF1 -ticket TICKETNUMBER;
    removal an explicitly specified lic to a user (in this case, using the LicenseSku for SPE_F1 - web-only o365 - lic as stored in a global variable)
    .EXAMPLE
    PS> remove-EXOLicense -users 'Test@domain.com' -LicenseSkuIds $TORMETA.o365LicSkuE3 -ticket TICKETNUMBER;
    removal an explicitly specified lic to a user (in this case, using the LicenseSku for ENTERPRISEPACK - E3 o365 - lic as stored in a global variable)
    .EXAMPLE
    PS> remove-EXOLicense -users 'Test@domain.com' -LicenseSkuIds 'TENANTNAME:EXCHANGESTANDARD' -ticket TICKETNUMBER;
    removal an explicitly specified lic to a user by specifying the Tenant-specific LicenseSkuID directly
    .EXAMPLE
    PS> remove-o365License -$MsoLUser.UserprincipalName -ticket TICKETNUMBER ;
    remove-o365License compatibility option
    .LINK
    https://github.com/tostka/verb-exo
    .LINK
    #>
    # migr to verb-exo, pull the dupe spec...
    #Requires -Modules AzureAD, MSOnline, ExchangeOnlineManagement, verb-AAD, verb-Auth, verb-IO, verb-logging, verb-Mods, verb-Text
    #Requires -RunasAdministrator
    # VALIDATORS: [ValidateNotNull()][ValidateNotNullOrEmpty()][ValidateLength(24,25)][ValidateLength(5)][ValidatePattern("some\s\regex\sexpr")][ValidateSet("USEA","GBMK","AUSYD")][ValidateScript({Test-Path $_ -PathType 'Container'})][ValidateScript({Test-Path $_})][ValidateRange(21,65)][ValidateCount(1,3)]
    ## [OutputType('bool')] # optional specified output type
    [CmdletBinding()]
    #[Alias('remove-o365License')]
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
         [Parameter(,HelpMessage="Optional Array, in preference order, of LicenseSkuID (e.g. TenantName:SPE_F1) to be added, runs full list removing all matches (default process is to dynamically resolve id's from Meta LicenseSkuKeys specifications)[-LicenseSkuIds @(`$XXXMETA.o365LicSkuExStd,`$XXXMETA.o365LicSkuF1)]")]
            #[ValidateNotNullOrEmpty()]
            [array]$LicenseSkuIds = @(), 
        [Parameter(HelpMessage="switch to override normal 'skipped' license application to existing Mailbox.[-Force]")]
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
            # included solely for backward compatibility with Removal-o365License()
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

        $rgxOPLic = '^CN\=ENT\-APP\-Office365\-(EXOK|F1|MF1)-DL$' ;
        $rgxXLic = '^CN\=ENT\-APP\-Office365\-(EXOK|F1|MF1)-DL$' ;

        # recycling the inbound above into next call in the chain
        $pltRXO = [ordered]@{
            Credential = $Credential ; 
            verbose = $($VerbosePreference -eq "Continue")  ; 
            silent = $silent ; 
        } ;

        $smsg = "Retrieve & build LicenseSkuIDS from global Meta vari" ;  
        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        [array]$LicenseSkuIds = @() ; 
        $LicenseSkuKeys | foreach-object { $LicenseSkuIds += @((get-variable -name "$($tenorg)META").value[$_]) } ; 

        #$rgxEmailAddr = '^([0-9a-zA-Z]+[-._+&'])*[0-9a-zA-Z]+@([-0-9a-zA-Z]+[.])+[a-zA-Z]{2,63}$ '
 
        $sBnr="`n#*======v $(${CmdletName}) : v======" ;
        $smsg = $sBnr ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

        $admin = "$env:username" ;

        # check if using Pipeline input or explicit params:
        if ($PSCmdlet.MyInvocation.ExpectingInput) {
            write-verbose "Data received from pipeline input: '$($InputObject)'" ;
        } else {
            # doesn't actually return an obj in the echo
            #write-verbose "Data received from parameter input: '$($InputObject)'" ;
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
            #-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
            # Looks like 1/5/2022, there are no spare E3's, maybe shift to the F3 ($TORMETA.o365LicSkuF1 = works to get mbx back).
            # (below defaults to the 'office 365 F3', the E3 alt is:  $tormeta.o365LicSkuE3 )
            # 12:06 PM 1/11/2022 add ExOnly: EXCHANGESTANDARD # Office 365 Exchange Online Only (commonly used for App
            $tUPN=$usr ;
            #$LicenseSkuIds=$TORMETA.o365LicSkuF1; # dyn from global XXXmeta
            $error.clear() ;
            TRY {

                $Exit = 0 ;
                Do {
                    Try {
                        connect-aad @pltRXO ; 
                        $AADUser=$null ;
                        #$TenantShortName = ((Get-AzureADTenantDetail).verifieddomains |?{$_._default}).name.split('.')[0] ;
                        $pltGAADU=[ordered]@{ 
                            ObjectId = $tUPN ;
                            ErrorAction = 'STOP' ;
                            verbose = $($VerbosePreference -eq "Continue") ;
                        } ;
                        $AADUser = Get-AzureADUser @pltGAADU ;
                        $Exit = $Retries ;
                    } Catch {
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

                # 8:44 AM 12/21/2022 no, use the verb-EXO:test-EXOIsLicensed(): test-EXOIsLicensed -User $AADUser -verbose
                $IsExoLicensed = test-EXOIsLicensed -User $AADUser -Credential:$pltRXO.Credential -verbose:$pltRXO.verbose -silent:$pltRXO.silent ;
                $pltGLPList=[ordered]@{ 
                    TenOrg= $TenOrg;
                    verbose=$($VerbosePreference -eq "Continue") ;
                    credential= $pltRXO.credential ;
                    silent = $silent ; 
                } ;
                $smsg = "$($tenorg):get-AADlicensePlanList wn$(($pltGLPList|out-string).trim())" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                $skus = $null ;
                $skus = get-AADlicensePlanList @pltGLPList ;

                $ombx = get-xomailbox -id $AADUser.UserPrincipalName -ea continue  ;
                $ombx = $ombx | ?{$_ -is [System.Management.Automation.PSObject]} # filtering any aberrant obj returned (legacy of prior problematic xow reliance to work around hybrid stepable pipeline bug)
                $AADLicDetails = get-AADUserLicenseDetails -UPNs $AADUser.userprincipalname -Credential:$pltRXO.Credential -verbose:$pltRXO.verbose -silent:$pltRXO.silent ; 
                #$AADLicDetails = get-AADUserLicenseDetails -UPNs $AADUser.userprincipalname -Verbose:$($VerbosePreference -eq "Continue")
                $smsg = "`nExisting Mbx:`n$(($ombx | ft -a 'RecipientType','RecipientTypeDetails'|out-string).trim())" ;
                $smsg += "`n`$AADLicDetails`n$(($AADLicDetails|out-string).trim())" ;
                if($ombx.RecipientTypeDetails -eq 'SharedMailbox'){
                    $smsg += "`nSharedMailbox does not *require* a license" ;
                } ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                #if( -not($Force) -AND $ombx){
                <#if( -not($Force) -AND ($ombx.RecipientTypeDetails -eq 'SharedMailbox') ){
                    $smsg += "`n -- SKIPPING EXO-RELATED LICENSE Removal! --" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                } else
                #>
                if( ($IsExoLicensed) -OR ($Force) ){
                    # not supported on aadu: defer to: verb-AAD:test-AADUserIsLicensed(): $isLicensed = test-AADUserIsLicensed -user $AzureADUser -verbose
                    if($IsAADIsLicensed = test-AADUserIsLicensed -user $AADUser -Verbose:($VerbosePreference -eq 'Continue')){
                        # has a bozo lic that doesn't support a mailbox
                        $smsg = "AADUser:$($tUPN):  isLicensed (has some form of license added), and has an EXO UserMailbox-supporting license!" ;
                        $smsg += "`n(or is being -Force upgraded to an elevated license)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN }
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ;
                    $smsg="confirmed $($AADUser.UserPrincipalName):is licensed/overlicensed" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;


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

                        $pltRAADUL=[ordered]@{
                            Users=$AADUser.UserPrincipalName ;
                            skuid=$LicenseSkuId ;
                            Credential = $pltRXO.Credential ; 
                            verbose = $pltRXO.verbose  ; 
                            silent = $false ; 
                            erroraction = 'STOP' ;
                            whatif = $($whatif) ;
                        } ;
                        $smsg = "remove-AADUserLicense w`n$(($pltRAADUL|out-string).trim())" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                        $bRet = remove-AADUserLicense @pltRAADUL ;
                        if($bRet.Success){
                            $smsg = "remove-AADUserLicense removed Licenses:$($bRet.RemovedLicenses)" ;
                            # $AADUser.AssignedLicenses.skuid
                            $smsg += "`n$(($AADUser.AssignedLicenses.skuid|out-string).trim())" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                            $smsg = "Detailed Return:`n$(($bRet|out-string).trim())" ;
                            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                            #BREAK ; # process through all lic potential removals
                        } elseif($whatif){
                            $smsg = "(whatif pass, exec skipped), " ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        } elseif( -not [boolean]($bRet.RemovedLicensess)){
                            # failed removal
                            $smsg = "Failed Lic Removal:$($LicenseSkuId) (moving on to next if avail...)" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN }
                            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            Continue ;
                        } else {
                            $smsg = "FAILED TO UPDATE !" ;
                            $smsg += "`n$(($bRet|out-string).trim())" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            #$Report.FixedUsageLocation = $false ;
                            if(-not $whatif){
                                BREAK;
                            }
                        } ;

                    } ;  # loop-E $LicenseSkuIds

                    connect-aad @pltRXO ; 
                    $AADUser=$null ;
                    #$TenantShortName = ((Get-AzureADTenantDetail).verifieddomains |?{$_._default}).name.split('.')[0] ;
                    $pltGAADU=[ordered]@{ 
                        ObjectId = $tUPN ;
                        ErrorAction = 'STOP' ;
                        verbose = $($VerbosePreference -eq "Continue") ;
                    } ;
                    # refresh ADU post chgs & test xmbx lic stat
                    $AADUser = Get-AzureADUser @pltGAADU ;
                    if(-not $LicenseSkuIds){
                        # running explicit LicenseSkuIds may not have removed all EXO lic's, so no point in doing a followup confirm license-free
                        $IsExoLicensed = test-EXOIsLicensed -User $AADUser -Credential:$pltRXO.Credential -verbose:$pltRXO.verbose -silent:$pltRXO.silent ;
                        if($IsExoLicensed){
                            $smsg = "AADUser still coming back with mounted EXO-supporting license. `nRe-Running remove-EXOLicesnse pass..." ; 
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                            $pltRxLic=[ordered]@{
                                users=$AADUser.userprincipalname ;
                                ticket=$ticket ;
                                whatif=$($whatif) ;
                                Verbose=$($VerbosePreference -eq "Continue")
                                Credential = $pltRXO.Credential ;
                                silent = $false ; 
                            } ;
                            if($UserRole){
                                $smsg = "(recycle `$UserRole from script)" ; 
                                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                                $pltRxLic.UserRole = $UserRole ; 
                            } ; 
                            #$smsg = "remove-EXOLicense w`n$(($pltRxLic|out-string).trim())" ;
                            #if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                            #else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            $bRet = remove-EXOLicense @pltRxLic ;
                            if($bRet) {$AADUser = $bRet } ; 
                        } else { 
                            $smsg = "Validated:$($AADUser.userprincipalname): is now EXO-unlicensed" ; 
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                        } ; 
                    } elseif($whatif){
                        $smsg = "-whatif: skipping post-validation" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                    } else { 
                        # running explicit LicenseSkuIds may not have removed all EXO lic's, so no point in doing a followup confirm license-free
                        $smsg = "-LicenseSkuId specified: Skipping broad test-ExoIsLicensed confirmations" ; 
                        $smsg += "`nsolely the licenses specified would have been removed," ; 
                        $smsg += "`nand may not be the complete EXO-license set" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                    } ; 

                };  # if-E $ombx

            } CATCH {     
                $ErrTrapd=$Error[0] ;
                $smsg = "$('*'*5)`nFailed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: `n$(($ErrTrapd|out-string).trim())`n$('-'*5)" ;
                write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" ;
                Break ;
            } ;
            if(!$whatif){
                $smsg = "dawdling to ensure License change doesn't soft-delete mailbox..." ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                $1F=$false ;
                Do {
                    if($1F){Sleep -s 30} ;
                    write-host "." -NoNewLine ;
                    $1F=$true ;
                #} Until (get-xomailbox -id $oMSUsr.userprincipalname -EA 0) ;
                } Until ($ombx = get-xomailbox -id $AADUser.userprincipalname -EA 0) ; # capture return (prevent from dropping into pipe)
                # get-xomailbox returns: System.Management.Automation.PSObject; not a real Mailbox object class
                $ombx = $ombx | ?{$_ -is [System.Management.Automation.PSObject]} ; # looks like an attempt to filter just the mailbox out of the pipeline return
                $smsg = "xo Mailbox confirmed!" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } ;

            # return $AADUser to pipeline if populated

            $AADUser | write-output ;

            $smsg =  $sBnr.replace('=v','=^').replace('v=','^=') ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        } ; # ($usr in $users)

    } # PROC-E
    END{
        <#
        $stopResults = stop-transcript ;
        $smsg = $stopResults ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        #>
     } ;
 }

#*------^ remove-EXOLicense.ps1 ^------