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
    Array, in preference order, of XXXMeta global value LicenseSkuKey names (resolves SKUId from TenOrg global Meta vari ; first working lic assignment, will be applied)[-LicenseSkuIds 'o365LicSkuExStd','o365LicSkuF1']
    .PARAMETER Force
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
    ###Requires -Version 5
    ###Requires -Modules ActiveDirectory, MSOnline, AzureAD, ExchangeOnlineManagement, verb-AAD, verb-ADMS, verb-Ex2010
    ##Requires -Modules ActiveDirectory, AzureAD, MSOnline, ExchangeOnlineManagement, verb-AAD, verb-ADMS, verb-Auth, verb-Ex2010, verb-EXO, verb-IO, verb-logging, verb-Mods, verb-Network, verb-Text, verb-logging
    # stripped down, doesn't really need AAD, may not need balance.
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
            [ValidatePattern("^([0-9a-zA-Z]+[-._+&'])*[0-9a-zA-Z]+@([-0-9a-zA-Z]+[.])+[a-zA-Z]{2,63}$")]
            [array]$users,
        [Parameter(Mandatory=$True,HelpMessage="Ticket Number [-Ticket '999999']")]
            [string]$Ticket,
        [Parameter(,HelpMessage="Array, in preference order, of XXXMeta global value LicenseSkuKey names (resolves SKUId from TenOrg global Meta vari ; first working lic assignment, will be applied)[-LicenseSkuIds 'o365LicSkuExStd','o365LicSkuF1']")]
            [ValidateNotNullOrEmpty()]
            [array]$LicenseSkuKeys=@('o365LicSkuExStd','o365LicSkuF1','o365LicSkuE3'),
        [Parameter(Mandatory=$FALSE,HelpMessage="TenantTag value, indicating Tenants to connect to[-TenOrg 'TOL']")]
            [ValidateNotNullOrEmpty()]
            #[ValidatePattern("^\w{3}$")]
            [string]$TenOrg = $global:o365_TenOrgDefault,
        [Parameter(HelpMessage="Use EXOv2 (ExchangeOnlineManagement) over basic auth legacy connection [-useEXOv2]")]
            [switch] $useEXOv2,
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
    <# add-o365License parms: (compatib): takes an MSolUser, workaround, spec $MsoLUser.UserprincipalName 
        [Parameter(Position = 0, Mandatory = $True, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, HelpMessage = "MSolUser [-MSolUser `$UserObjectVariable ]")]
        $MSolUser,
        [Parameter(Position = 0, Mandatory = $false, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, HelpMessage = "MS LicenseSku value for license to be applied (defaults to EXCHANGESTANDARD) [-LicenseSku tenantname:LICENSESKU]")]
        $LicenseSku = "toroco:EXCHANGESTANDARD",
        [switch] $showDebug,
        [Parameter(HelpMessage = "Whatif Flag  [-whatIf]")]
        [switch] $whatIf
    #>
    <# remove-EXOLicense -users fname.lname@domain.com -ticket 99999 -whatif -verbose 

    #>
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
        #$rgxEmailAddr = '^([0-9a-zA-Z]+[-._+&'])*[0-9a-zA-Z]+@([-0-9a-zA-Z]+[.])+[a-zA-Z]{2,63}$ '
        
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

        # finally if we're using pipeline, and aggregating, we need to aggreg outside of the process{} block
        if($PSCmdlet.MyInvocation.ExpectingInput){
            # pipeline instantiate an aggregator here
        } ;

        
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

        <# if we want pipeline to work have to move the clipboard grab out or down into process{}, where pipeline binding will be actually populated
        if(!$users){
            $users= (get-clipboard).trim().replace("'",'').replace('"','') ;
            if($users){
                write-verbose "No -users specified, detected value on clipboard:`n$($users)" ;
            } else {
                write-warning "No -users specified, nothing matching dname, emailaddress or samaccountname, found on clipboard. EXITING!" ;
                Break ;
            } ;
        } else {
            write-verbose "($(($users|measure).count)) user(s) specified:`n'$($users -join "','")'" ;
        } ;
        #>

    }
    PROCESS{


        $ttl = ($users|measure).count ; $Procd=0 ;
        [array]$Rpt =@() ;
        # with pipeline input, the pipeline evals as either $_ (if unmapped to a param in binding), or iterating on the mapped value.
        #     the foreach loop below doesn't actually loop. Process{} is the loop with a pipeline-fed param, and the bound - $users - variable once per pipeline bound element - per array item on an array -
        #     is run with the $users value populated with each element in turn. IOW, the foreach is a single-run pass, and the Process{} block is the loop.
        # you need both a bound $users at the top - to handle explicit assigns remove-EXOLicense -users $variable.
        # with a process {} block to handle any pipeline passed input. The pipeline still maps to the bound param: $users, but the e3ntire process{} is run per element, rather than iteratign the internal $users foreach.

        foreach ($usr in $users){

            #-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
            # Looks like 1/5/2022, there are no spare E3's, maybe shift to the F3 ($TORMETA.o365LicSkuF1 = works to get mbx back).
            # (below defaults to the 'office 365 F3', the E3 alt is:  $tormeta.o365LicSkuE3 )
            # 12:06 PM 1/11/2022 add ExOnly: EXCHANGESTANDARD # Office 365 Exchange Online Only (commonly used for App

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
            #$LicenseSkuIds=$TORMETA.o365LicSkuF1; # dyn from global XXXmeta
            $error.clear() ;
            TRY {

                $Exit = 0 ;
                
                connect-aad @pltRXO ; 

                # need to pop the pltGLPListist ;
                $pltGLPList=[ordered]@{ 
                    indexonname = $true ; 
                    TenOrg= $TenOrg;
                    verbose=$($VerbosePreference -eq "Continue") ;
                    credential= $pltRXO.credential ;
                    silent = $silent ; 
                } ;
                $smsg = "$($tenorg):get-AADlicensePlanList wn$(($pltGLPList|out-string).trim())" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                $skus = $null ;
                $skus = get-AADlicensePlanList @pltGLPList ;

                #Get-MsolAccountSku -ea STOP ;

                #$pltALic=[ordered]@{UserPrincipalName=$usr ; RemoveLicenses=$null ;} ;
                foreach($LicenseSkuId in $LicenseSkuIds){
                                        
                    $smsg = "(attempting license:$($LicenseSkuId)...)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                    <# skus come back as:
                        @($TORMETA.o365LicSkuExStd,$TORMETA.o365LicSkuF1,$TORMETA.o365LicSkuE3)

                        toroco:EXCHANGESTANDARD
                        toroco:SPE_F1
                        toroco:ENTERPRISEPACK
                    #>
                    # split and process the entry
                    $SkuPartNumber = $LicenseSkuId.split(':')[1] ; 
                    $skuid = $skus[$SkuPartNumber].skuid ; 
                    # convert it to a skuid
                    <# $skuid = $lplistn['EXCHANGESTANDARD'].skuid ; 
                    # then pass it to 
                    PS> $bRet = remove-AADUserLicense -users 'upn@domain.com','upn2@domain.com' -skuid $skuid -verbose -whatif ; 
                    #>
                     #$bRes = remove-AADUserLicense -Users $AADUser.UserPrincipalName -skuid $LicenseSkuId -verbose -whatif
                    $pltRAADUL=[ordered]@{
                        Users=$usr ;
                        skuid=$skuid ;
                        Credential = $pltRXO.Credential ; 
                        verbose = $pltRXO.verbose  ; 
                        silent = $pltRXO.silent ; 
                        erroraction = 'STOP' ;
                        whatif = $($whatif) ;
                    } ;
                    $smsg = "remove-AADUserLicense w`n$(($pltRAADUL|out-string).trim())" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                    $bRet = remove-AADUserLicense @pltRAADUL ; 
                    if($bRet.Success){
                        # AzureADUser, RemovedLicenses
                        $smsg = "remove-AADUserLicense removed License:$($bRet.RemovedLicenses)" ; 
                        # $AADUser.AssignedLicenses.skuid
                        $smsg += "`n$(($AADUser.AssignedLicenses.skuid|out-string).trim())" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            
                        $smsg = "Detailed Return:`n$(($bRet|out-string).trim())" ; 
                        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                        # update the local AADUser to reflect the updated AADU returned
                        $AADUser = $bRet.AzureADUser ; 
                        #$Report.FixedUsageLocation = $true ; 
                        BREAK ; # abort further loops if one successfully applied
                    } elseif($whatif){
                        $smsg = "(whatif pass, exec skipped), " ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } else { 
                        $smsg = "remove-AADUserLicense : FAILED TO REMOVE LICENSE!" ;
                        $smsg += "`n$(($bRet|out-string).trim())" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug 
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        #$Report.FixedUsageLocation = $false ; 
                        if(-not $whatif){
                            BREAK; 
                        } 
                    } ; 
                   
                } ;  # loop-E $LicenseSkuIds
                  



            } CATCH {     $ErrTrapd=$Error[0] ;
                $smsg = "$('*'*5)`nFailed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: `n$(($ErrTrapd|out-string).trim())`n$('-'*5)" ;
                write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" ;
                Break ;
            } ;
            if(!$whatif){
                <# don't care if mbx evap's, just want lic back
                $smsg = "dawdling until License reinflates mbx..." ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                $1F=$false ;
                Do {
                    if($1F){Sleep -s 30} ;
                    write-host "." -NoNewLine ;
                    $1F=$true ;
                } Until ($ombx = ps1GetxMbx -id $oMSUsr.userprincipalname -EA 0) ; # capture return (prevent from dropping into pipe)
                $smsg = "xo Mailbox confirmed!" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                #>
            } ;

            #-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=


            $smsg =  $sBnr.replace('=v','=^').replace('v=','^=') ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        } ; # ($usr in $users)

    } # PROC-E
    END{

        # return $oMSUsr to pipeline if populated

        $AzureADUser | write-output ;
            
        <#
        if($outObject -AND -not ($PSCmdlet.MyInvocation.ExpectingInput)){
            $Rpt | write-output ;
            $smsg = "(-outObject: Output summary object to pipeline)"
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

        }elseif($outObject -AND ($PSCmdlet.MyInvocation.ExpectingInput)){
            write-verbose "(pipeline input, individual objects dropped into pipeline)" ;

        } else {
            $oput = ($Rpt | select-object -unique) -join ',' ;
            $oput | out-clipboard ;
            $smsg = "(output copied to clipboard)"
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

            $oput |  write-output ;
        } ;
        #>
        $stopResults = stop-transcript ;
        $smsg = $stopResults ; 
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
     } ;
 }

#*------^ remove-EXOLicense.ps1 ^------