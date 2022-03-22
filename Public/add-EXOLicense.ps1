{
    <#
    .SYNOPSIS
    add-EXOLicense.ps1 - Add a temporary o365 license to specified MsolUser account. Returns updated MSOLUser object to pipeline.
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
    Tags        : Powershell
    AddedCredit : REFERENCE
    AddedWebsite: URL
    AddedTwitter: URL
    REVISIONS
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
    $TORMETA.o365LicSkuExStd == EXCHANGESTANDARD (Office 365 Exchange Online Only 
    ,commonly used for App Access) & stick in front of $LicenseSkuIds, 
    $TORMETA.o365LicSkuExStd; added examples with explicit cmdlines for the adds; 
    spliced over UsageLocation test/assert code from add-o365license. 
    * 1:34 PM 1/5/2022 init
    .DESCRIPTION
    add-EXOLicense.ps1 - Add a temporary o365 license to specified MsolUser account. Returns updated MSOLUser object to pipeline.
    .PARAMETER Ticket
    Ticket Number [-Ticket '999999']
    .PARAMETER TenOrg
    TenantTag value, indicating Tenants to connect to[-TenOrg 'TOL']
    .PARAMETER  users
    Array of UserPrincipalNames (or MSOLUser objects) to have a temporary Exchange License applied
    .PARAMETER Credential
    Use specific Credentials (defaults to Tenant-defined SvcAccount)[-Credentials [credential object]]
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
    PS> add-EXOLicense -users 'Test@domain.com','Test2@domain.com' -verbose  ;
    Process an array of users, with default 'hunting' -LicenseSkuIds array. 
    .EXAMPLE
    PS> $updatedMSOLU = add-EXOLicense -users 'Test@domain.com','Test2@domain.com' -verbose ;
        if($updatedMSOLU.islicensed){'Y' ; $updatedMSOLU.licenses } else { 'N' } ; 
    Process license add for specified user, and post-test isLicensed status, using default license array configured on the -LicenseSkuIDs default value. Then echo the current licenses list (as returned in the updated MSOLUser object). 
    .EXAMPLE
PS> $whatif=$true ;
    $target = 'TICKETNUMBER,USERUPN' ;
    if($target.contains(',')){
        $ticket,$trcp = $target.split(',') ;
        $updatedmsolu = add-EXOLicense -users $trcp -Ticket $ticket -whatif:$($whatif) ;
        $props1 = 'UserPrincipalName','DisplayName','IsLicensed' ;
        $props2 = @{Name='Licenses';
        Expression={$_.licenses.accountskuid -join ', '}}  ;
        $smsg = "UpdatedMsolU: w`n$(($updatedmsolu| ft -auto $props1 |out-string).trim())" ;
        $smsg += "`n:$(($updatedmsolu| fl $props2 |out-string).trim())" ;
        write-host -foregroundcolor green $smsg ;
        if(!$whatif){
            write-host "dawdling until License reinflates mbx..." ;
            $1F=$false ;
            Do {
                if($1F){Sleep -s 30} ;
                write-host "." -NoNewLine ;
                $1F=$true ;
            } Until (get-exomailbox -id $trcp  -EA 0) ;
            write-host "Mailbox reattached: Ready for conversion!" ;
        } ;
    } else { write-warning "`$target does *not* contain comma delimited ticket,UPN string!"} ;
    Fancier variant of above, with more post-confirm reporting
    .EXAMPLE
    PS> add-EXOLicense -users 'Test@domain.com' -LicenseSkuIds $TORMETA.o365LicSkuExStd -ticket TICKETNUMBER;
    add an explicitly specified lic to a user (in this case, using the LicenseSku for EXCHANGESTANDARD, as stored in a global variable)
    .EXAMPLE
    PS> add-EXOLicense -users 'Test@domain.com' -LicenseSkuIds $TORMETA.o365LicSkuF1 -ticket TICKETNUMBER;
    add an explicitly specified lic to a user (in this case, using the LicenseSku for SPE_F1 - web-only o365 - lic as stored in a global variable)
    .EXAMPLE
    PS> add-EXOLicense -users 'Test@domain.com' -LicenseSkuIds $TORMETA.o365LicSkuE3 -ticket TICKETNUMBER ;
    add an explicitly specified lic to a user (in this case, using the LicenseSku for ENTERPRISEPACK - E3 o365 - lic as stored in a global variable)
    .EXAMPLE
    PS> add-EXOLicense -users 'Test@domain.com' -LicenseSkuIds 'TENANTNAME:EXCHANGESTANDARD' -ticket TICKETNUMBER ;
    add an explicitly specified lic to a user by specifying the Tenant-specific LicenseSkuID directly
    .EXAMPLE
    PS> add-o365License -$MsoLUser.UserprincipalName -ticket TICKETNUMBER ;
    add-o365License compatibility option
    .LINK
    https://github.com/tostka/verb-exo
    #>
    ###Requires -Version 5
    ###Requires -Modules ActiveDirectory, MSOnline, AzureAD, ExchangeOnlineManagement, verb-AAD, verb-ADMS, verb-Ex2010
    ##Requires -Modules ActiveDirectory, AzureAD, MSOnline, ExchangeOnlineManagement, verb-AAD, verb-ADMS, verb-Auth, verb-Ex2010, verb-EXO, verb-IO, verb-logging, verb-Mods, verb-Network, verb-Text, verb-logging
    # stripped down, doesn't really need AAD, may not need balance.
    ##Requires -Modules AzureAD, MSOnline, ExchangeOnlineManagement, verb-AAD, verb-Auth, verb-EXO, verb-IO, verb-logging, verb-Mods, verb-Text
    # migr to verb-exo, pull the dupe spec...
    #Requires -Modules AzureAD, MSOnline, ExchangeOnlineManagement, verb-AAD, verb-Auth, verb-IO, verb-logging, verb-Mods, verb-Text
    #Requires -RunasAdministrator
    # VALIDATORS: [ValidateNotNull()][ValidateNotNullOrEmpty()][ValidateLength(24,25)][ValidateLength(5)][ValidatePattern("some\s\regex\sexpr")][ValidateSet("USEA","GBMK","AUSYD")][ValidateScript({Test-Path $_ -PathType 'Container'})][ValidateScript({Test-Path $_})][ValidateRange(21,65)][ValidateCount(1,3)]
    ## [OutputType('bool')] # optional specified output type
    [CmdletBinding()]
    #[Alias('add-o365License')]
    PARAM(
        [Parameter(Position=0,Mandatory=$False,ValueFromPipeline=$true,HelpMessage="Array of UserPrincipalNames (or MSOLUser objects) to have a temporary Exchange License applied")]
        #[ValidateNotNullOrEmpty()]
        #[Alias('ALIAS1', 'ALIAS2')]
        [ValidatePattern("^([0-9a-zA-Z]+[-._+&'])*[0-9a-zA-Z]+@([-0-9a-zA-Z]+[.])+[a-zA-Z]{2,63}$")]
        [array]$users,
        [Parameter(Mandatory=$True,HelpMessage="Ticket Number [-Ticket '999999']")]
        [string]$Ticket,
        [Parameter(,HelpMessage="Array, in preference order, of Tenant-specific LicenseSku names (first working lic assignment will be applied)[-LicenseSkuIds 'tenantname:SPE_F1','tenantname:ENTERPRISEPACK']")]
        [Alias('LicenseSku')]
        [ValidateNotNullOrEmpty()]
        [array]$LicenseSkuIds=@($TORMETA.o365LicSkuExStd,$TORMETA.o365LicSkuF1,$TORMETA.o365LicSkuE3),
        [Parameter(Mandatory=$False,HelpMessage="TenantTag value, indicating Tenants to connect to[-TenOrg 'TOL']")]
        [ValidateNotNullOrEmpty()]
        [ValidatePattern("^\w{3}$")]
        [string]$TenOrg = 'TOR',
        [Parameter(HelpMessage="Use EXOv2 (ExchangeOnlineManagement) over basic auth legacy connection [-useEXOv2]")]
        [switch] $useEXOv2,
        [Parameter(Mandatory = $false, HelpMessage = "Use specific Credentials (defaults to Tenant-defined SvcAccount)[-Credentials [credential object]]")]
        [System.Management.Automation.PSCredential]$Credential,
        [Parameter(Mandatory = $false, HelpMessage = "Credential User Role spec (SID|CSID|UID|B2BI|CSVC)[-UserRole SID]")]
        [ValidateSet('SID', 'CSID', 'UID', 'B2BI', 'CSVC')]
        [string]$UserRole = 'SID',
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
    <# add-EXOLicense -users fname.lname@domain.com -ticket 99999 -whatif -verbose 

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
        if ($PSScriptRoot -eq "") {
            if ($psISE) { $ScriptName = $psISE.CurrentFile.FullPath }
            elseif ($context = $psEditor.GetEditorContext()) {$ScriptName = $context.CurrentFile.Path }
            elseif ($host.version.major -lt 3) {
                $ScriptName = $MyInvocation.MyCommand.Path ;
                $PSScriptRoot = Split-Path $ScriptName -Parent ;
                $PSCommandPath = $ScriptName ;
            } else {
                if ($MyInvocation.MyCommand.Path) {
                    $ScriptName = $MyInvocation.MyCommand.Path ;
                    $PSScriptRoot = Split-Path $MyInvocation.MyCommand.Path -Parent ;
                } else {throw "UNABLE TO POPULATE SCRIPT PATH, EVEN `$MyInvocation IS BLANK!" } ;
            };
            $ScriptDir = Split-Path -Parent $ScriptName ;
            $ScriptBaseName = split-path -leaf $ScriptName ;
            $ScriptNameNoExt = [system.io.path]::GetFilenameWithoutExtension($ScriptName) ;
        } else {
            $ScriptDir = $PSScriptRoot ;
            if ($PSCommandPath) {$ScriptName = $PSCommandPath }
            else {
                $ScriptName = $myInvocation.ScriptName
                $PSCommandPath = $ScriptName ;
            } ;
            $ScriptBaseName = (Split-Path -Leaf ((& { $myInvocation }).ScriptName))  ;
            $ScriptNameNoExt = [system.io.path]::GetFilenameWithoutExtension($MyInvocation.InvocationName) ;
        } ;
        if ($showDebug) { write-debug -verbose:$true "`$ScriptDir:$($ScriptDir)`n`$ScriptBaseName:$($ScriptBaseName)`n`$ScriptNameNoExt:$($ScriptNameNoExt)`n`$PSScriptRoot:$($PSScriptRoot)`n`$PSCommandPath:$($PSCommandPath)" ; } ;

        # finally if we're using pipeline, and aggregating, we need to aggreg outside of the process{} block
        if($PSCmdlet.MyInvocation.ExpectingInput){
            # pipeline instantiate an aggregator here
        } ;

        # to sketch in support for passing either a UPN or an MSOLUser (convert the Msolu to upn)
        [array]$userstemp = $()  ; 
        foreach($user in $users){
            switch($user.GetType().FullName){
                'Microsoft.Online.Administration.User' {
                    $smsg = "(-user:MsolU detected:$($user.userprincipalname), extracting the UPN...)" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                    $userstemp +=$user.userprincipalname ; 
                } ; 
                'System.String'{

                    $smsg = "(-user:string detected)" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 

                    if($user -match $rgxEmailAddress){

                        $smsg = "(-user:EmailAddress/UPN detected:$($user))" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                        $userstemp +=$user ; 
                    } ; 
                } 
                default{
                    $smsg = "Unrecognized format for -User:$($User)!. Please specify either a user UPN, or pass a full MsolUser object." ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                    Break ; 
                } 
            } ;
        } ;  # loop-E $users

        $users = $userstemp ; 

        #*======V CONFIGURE DEFAULT LOGGING FROM PARENT SCRIPT NAME v======
        $pltSL=@{ NoTimeStamp=$false ; Tag="$($ticket)-$($TenOrg)-LASTPASS-$($users -join ',')" ; showdebug=$($showdebug) ; whatif=$($whatif) ; Verbose=$($VerbosePreference -eq 'Continue') ; } ;
        $smsg = "start-Log w`n$()$(($pltSL|out-string).trim())" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

        if($PSCommandPath){
            $logspec = start-Log -Path $PSCommandPath @pltSL ;
            $smsg += " -Path $($PSCommandPath)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        } else {
            $logspec = start-Log -Path ($MyInvocation.MyCommand.Definition) @pltSL ;
            $smsg += " -Path $($MyInvocation.MyCommand.Definition)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        } ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        if($logspec){
            $logging=$logspec.logging ;
            $logfile=$logspec.logfile ;
            $transcript=$logspec.transcript ;
            if(Test-TranscriptionSupported){
                $stopResults = try {Stop-transcript -ErrorAction stop} catch {} ;
                $startResults = start-transcript -Path $transcript ;
                # start-tra is winding up in pipeline, cap and log it.
                $smsg = $startResults ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } ;
        } else {throw "Unable to configure logging!" } ;
        #*======^ CONFIGURE DEFAULT LOGGING FROM PARENT SCRIPT NAME ^======

        $sBnr="`n#*======v $(${CmdletName}) : v======" ;
        $smsg = $sBnr ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

        $admin = "$env:username" ;

        #*======v EXO/EXOv2 CMDLET ALIASING v======
        #-=-=-=-=-=-=-=-=
        # simple loop to stock the set, no set->get conversion, roughed in $Exov2 exo->xo replace. Do specs in exo, and flip to suit under $exov2
        #configure EXO EMS aliases to cover useEXOv2 requirements
        # have to preconnect, as it gcm's the targets
        if ($script:useEXOv2) { reconnect-eXO2 }
        else { reconnect-EXO } ;
        # aliased ExOP|EXO|EXOv2 cmdlets (permits simpler single code block for any of the three variants of targets & syntaxes)
        # each is '[aliasname];[exOcmd] (xOv2cmd & exop are converted from [exocmd])
        <#[array]$cmdletMaps = 'ps1GetxRcp;get-exorecipient;','ps1GetxMbx;get-exomailbox;','ps1SetxMbx;Set-exoMailbox;','ps1GetxUser;get-exoUser;',
            'ps1GetxMUsr;Get-exoMailUser','ps1SetxMUsr;Set-exoMailUser','ps1SetxCalProc;set-exoCalendarprocessing;',
            'ps1GetxCalProc;get-exoCalendarprocessing;','ps1GetxMbxFldrPerm;get-exoMailboxfolderpermission;',
            'ps1GetxMbxFldrPerm;get-exoMailboxfolderpermission','ps1AddxMbxPrm;Add-exoMailboxPermission','ps1GetxMbxPrm;Get-exoMailboxPermission',
            'ps1RmvxMbxPrm;Remove-exoMailboxPermission','ps1AddRcpPrm;Add-exoRecipientPermission','ps1GetRcpPrm;Get-exoRecipientPermission',
            'ps1RmvRcpPrm;Remove-exoRecipientPermission','ps1GetxAccDom;Get-exoAcceptedDomain;','ps1GetxRetPol;Get-exoRetentionPolicy',
            'ps1GetxDistGrp;get-exoDistributionGroup;','ps1GetxDistGrpMbr;get-exoDistributionGroupmember;','ps1GetxMsgTrc;get-exoMessageTrace;',
            'ps1GetxMsgTrcDtl;get-exoMessageTraceDetail;','ps1GetxMbxFldrStats;get-exoMailboxfolderStatistics','ps1GetxMContact;Get-exomailcontact;',
            'ps1SetxMContact;Set-exomailcontact;','ps1NewxMContact;New-exomailcontact','ps1TestxMapi;Test-exoMAPIConnectivity',
            'ps1GetxOrgCfg;Get-exoOrganizationConfig','ps1GetxMbxRegionCfg;Get-exoMailboxRegionalConfiguration',
            'ps1TestxOAuthConn;Test-exoOAuthConnectivity','ps1NewxDistGrp;new-exoDistributionGroup','ps1SetxDistGrp;set-exoDistributionGroup',
            'ps1AddxDistGrpMbr;Add-exoDistributionGroupMember','ps1RmvxDistGrpMbr;remove-exoDistributionGroupMember',
            'ps1GetxDDG;Get-exoDynamicDistributionGroup','ps1NewxDDG;New-exoDynamicDistributionGroup','ps1SetxDDG;Set-exoDynamicDistributionGroup' ;
            'ps1GetxCasMbx;Get-exoCASMailbox','ps1GetxMbxStat;Get-exoMailboxStatistics','ps1GetxMobilDevStats;Get-exoMobileDeviceStatistics'
            #>
        [array]$cmdletMaps = 'ps1GetxMbx;get-exomailbox;' ; # reduced to single cmd, 
        [array]$XoOnlyMaps = 'ps1GetxMsgTrcDtl','ps1TestxOAuthConn' ; # cmdlet alias names from above that are skipped for aliasing in EXOP
        # cmdlets from above that have diff names EXO v EXoP: these each have  schema: [alias];[xoCmdlet];[opCmdlet]; op Aliases use the opCmdlet as target
        [array]$XoRenameMaps = 'ps1GetxMsgTrc;get-exoMessageTrace;get-MessageTrackingLog','ps1AddRcpPrm;Add-exoRecipientPermission;Add-AdPermission',
                'ps1GetRcpPrm;Get-exoRecipientPermission;Get-AdPermission','ps1RmvRcpPrm;Remove-exoRecipientPermission;Remove-ADPermission' ;
        [array]$Xo2VariantMaps =   'ps1GetxCasMbx;Get-exoCASMailbox', 'ps1GetxMbx;get-exomailbox;', 'ps1GetxMbxFldrPerm;get-exoMailboxfolderpermission;',
            'ps1GetxMbxFldrStats;get-exoMailboxfolderStatistics', 'ps1GetxMbxPrm;Get-exoMailboxPermission', 'ps1GetxMbxStat;Get-exoMailboxStatistics',
            'ps1GetxMobilDevStats;Get-exoMobileDeviceStatistics', 'ps1GetxRcp;get-exorecipient;', 'ps1AddRcpPrm;Add-exoRecipientPermission' ; 
        # cmdlets above have XO2 enhanced variant-named versions to target (they never are prefixed verb-xo[noun], always/only verb-exo[noun])
        # code to summarize & indexed-hash the renamed cmdlets for variant processing
        $XoRenameMapNames = @() ; 
        $oxoRenameMaps = @{} ;
        $XoRenameMaps | foreach {     $XoRenameMapNames += $_.split(';')[0] ;     $name = $_.split(';')[0] ;     $oxoRenameMaps[$name] = $_.split(';')  ;  } ;
        $Xo2VariantMapNames = @() ;
        $oXo2VariantMaps = @{} ;
        $Xo2VariantMaps | foreach {  $Xo2VariantMapNames += $_.split(';')[0] ;  $name = $_.split(';')[0] ;  $oXo2VariantMaps[$name] = $_.split(';') ; } ; 
        #$cmdletMapsFltrd = $cmdletmaps|?{$_.split(';')[1] -like '*DistributionGroup*'} ;  # filtering subset
        #$cmdletMapsFltrd += $cmdletmaps|?{$_.split(';')[1] -like '*recipient'}
        $cmdletMapsFltrd = $cmdletmaps ; # or use full set
        foreach($cmdletMap in $cmdletMapsFltrd){
            if($script:useEXOv2){
                if($Xo2VariantMapNames -contains $cmdletMap.split(';')[0]){
                    write-verbose "$($cmdletMap.split(';')[1]) has an XO2-VARIANT cmdlet, renaming for XOV2 enhanced variant" ;
                    # sub -exoNOUN -> -NOUN using ExOP variant cmdlet
                    if(!($cmdlet= Get-Command $oXo2VariantMaps[($cmdletMap.split(';')[0])][2] )){ throw "unable to gcm Alias definition!:$($oxoRenameMaps[($cmdletMap.split(';')[0])][2])" ; break }
                    $nAName = ($cmdletMap.split(';')[0]);
                    if(-not(get-alias -name $naname -ea 0 |?{$_.Definition -eq $cmdlet.name})){
                        $nalias = set-alias -name ($cmdletMap.split(';')[0]) -value ($cmdlet.name) -passthru ;
                        write-verbose "$($nalias.Name) -> $($nalias.ResolvedCommandName)" ;
                    } ;
                } else { 
                    # common cmdlets between all 3 systems
                    if(!($cmdlet= Get-Command $cmdletMap.split(';')[1].replace('-exo','-xo') )){ throw "unable to gcm Alias definition!:$($cmdletMap.split(';')[1])" ; break }
                    $nAName = ($cmdletMap.split(';')[0]) ;
                    if(-not(get-alias -name $naname -ea 0 |?{$_.Definition -eq $cmdlet.name})){
                        $nalias = set-alias -name $nAName -value ($cmdlet.name) -passthru ;
                        write-verbose "$($nalias.Name) -> $($nalias.ResolvedCommandName)" ;
                    } ;
                } ; 
            } else {
                if(!($cmdlet= Get-Command $cmdletMap.split(';')[1])){ throw "unable to gcm Alias definition!:$($cmdletMap.split(';')[1])" ; break }
                $nAName = ($cmdletMap.split(';')[0]);
                if(-not(get-alias -name $naname -ea 0 |?{$_.Definition -eq $cmdlet.name})){
                    $nalias = set-alias -name ($cmdletMap.split(';')[0]) -value ($cmdlet.name) -passthru ;
                    write-verbose "$($nalias.Name) -> $($nalias.ResolvedCommandName)" ;
                } ;
            } ;
        } ;# ...
        # cleanup example:
        #get-alias -scope Script |?{$_.name -match '^ps1.*'} | %{Remove-Alias -alias $_.name} ; 
        #*======^ EXO/EXOv2 CMDLET ALIASING ^======
        #

        #region SERVICE_CONNECTIONS #*======v SERVICE_CONNECTIONS v======
        #region useEXO ; #*------v useEXO v------
        $useEXO = $true ; # non-dyn setting, drives variant EXO reconnect & query code
        #if($CloudFirst){ $useEXO = $true } ; # expl: steering on a parameter
        if($useEXO){
            #region GENERIC_EXO_CREDS_&_SVC_CONN #*------v GENERIC EXO CREDS & SVC CONN BP v------
            # o365/EXO creds
            <### Usage: Type defaults to SID, if not spec'd - Note: there must be a *logged in & configured *profile*
            $o365Cred=get-TenantCredentials -TenOrg $TenOrg -verbose -userrole SID ;
            Returns a credential set for the $TenOrg Hybrid OnPrem Exchange Org
            .EXAMPLE
            $o365Cred=get-TenantCredentials -TenOrg $TenOrg -verbose -userrole CSVC ;
            Returns the CSVC Userrole credential for the $TenOrg Hybrid OnPrem Exchange Org
            .EXAMPLE
            $o365Cred=get-TenantCredentials -TenOrg $TenOrg -verbose -userrole B2BI ;
            Returns the B2BI Userrole credential for the $TenOrg Hybrid OnPrem Exchange Org
            ###>
            $o365Cred=$null ;
            if($o365Cred=(get-TenantCredentials -TenOrg $TenOrg -UserRole 'CSVC','SID' -verbose:$($verbose))){
                # make it script scope, so we don't have to predetect & purge before using new-variable
                if(get-Variable -Name cred$($tenorg) -scope Script -ea 0 ){ remove-Variable -Name cred$($tenorg) -scope Script } ;
                New-Variable -Name cred$($tenorg) -scope Script -Value $o365Cred.cred ;
                $smsg = "Resolved $($Tenorg) `$o365cred:$($o365Cred.cred.username) (assigned to `$cred$($tenorg))" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } else {
                $statusdelta = ";ERROR";
                $script:PassStatus += $statusdelta ;
                set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta) ;
                $smsg = "Unable to resolve $($tenorg) `$o365Cred value!"
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                throw "Unable to resolve $($tenorg) `$o365Cred value!`nEXIT!"
                Break ;
            } ;
            <### CALLS ARE IN FORM: (cred$($tenorg))
            $pltRXO = @{Credential = $Credential ; verbose = $($verbose) ; }
            $pltRXO = @{
                Credential = (Get-Variable -name cred$($tenorg) ).value ;
                #verbose = $($verbose) ;
                Verbose = $FALSE ; Silent = $true ;} ;
            if ($script:useEXOv2) { reconnect-eXO2 @pltRXO }
            else { reconnect-EXO @pltRXO } ;
            # or with Tenant-specific cred($Tenorg) lookup
            #$pltRXO creds & .username can also be used for AzureAD connections
            Connect-AAD @pltRXO ;
            ###>
            # configure splat for connections: (see above useage)
            $pltRXO = @{
                Credential = (Get-Variable -name cred$($tenorg) ).value ;
                #verbose = $($verbose) ;
                Verbose = $FALSE ; Silent = $true ;} ;
            #endregion GENERIC_EXO_CREDS_&_SVC_CONN #*------^ END GENERIC EXO CREDS & SVC CONN BP ^------
        } # if-E $useEXO
        #endregion useEXO ; #*------^ END useEXO ^------

        #region GENERIC_EXOP_CREDS_&_SRVR_CONN #*------v GENERIC EXOP CREDS & SRVR CONN BP v------
        # steer all onprem code on $XXXMeta.ExOPAccessFromToro & Ex10Server values
        $UseExOP=$true ;
        <# no onprem dep
        if((Get-Variable  -name "$($TenOrg)Meta").value.ExOPAccessFromToro -AND (Get-Variable  -name "$($TenOrg)Meta").value.Ex10Server){
            $UseExOP = $true ;
            $smsg = "$($TenOrg):Meta.ExOPAccessFromToro($((Get-Variable  -name "$($TenOrg)Meta").value.ExOPAccessFromToro)) -AND/OR Meta.Ex10Server($((Get-Variable  -name "$($TenOrg)Meta").value.Ex10Server)),`ENABLING use of OnPrem Ex system this pass." ;
            if($verbose){ if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
        } else {
            $UseExOP = $false ;
            $smsg = "$($TenOrg):Meta.ExOPAccessFromToro($((Get-Variable  -name "$($TenOrg)Meta").value.ExOPAccessFromToro)) -AND/OR Meta.Ex10Server($((Get-Variable  -name "$($TenOrg)Meta").value.Ex10Server)),`nDISABLING use of OnPrem Ex system this pass." ;
            if($verbose){ if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
        } ;
        #>
        if($UseExOP){
            #*------v GENERIC EXOP CREDS & SRVR CONN BP v------
            # do the OP creds too
            $OPCred=$null ;
            # default to the onprem svc acct
            $pltGHOpCred=@{TenOrg=$TenOrg ;userrole='ESVC','SID'; verbose=$($verbose)} ;
            if($OPCred=(get-HybridOPCredentials @pltGHOpCred).cred){
                # make it script scope, so we don't have to predetect & purge before using new-variable
                if(get-Variable -Name "cred$($tenorg)OP" -scope Script -ea 0 ){ remove-Variable -Name "cred$($tenorg)OP" -scope Script } ;
                New-Variable -Name "cred$($tenorg)OP" -scope Script -Value $OPCred ;
                $smsg = "Resolved $($Tenorg) `$OPCred:$($OPCred.username) (assigned to `$cred$($tenorg)OP)" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } else {
                $statusdelta = ";ERROR";
                $script:PassStatus += $statusdelta ;
                set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta) ;
                $smsg = "Unable to resolve get-HybridOPCredentials -TenOrg $($TenOrg) -userrole 'ESVC' value!"
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                throw "Unable to resolve $($tenorg) `$OPCred value!`nEXIT!"
                Break ;
            } ;
            $smsg= "Using EXOP cred:`$cred$($tenorg)OP:$((Get-Variable -name "cred$($tenorg)OP" ).value.username)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            <# CALLS ARE IN FORM: (cred$($tenorg))
                $pltRX10 = @{
                Credential = (Get-Variable -name "cred$($tenorg)OP" ).value ;
                #verbose = $($verbose) ;
                Verbose = $FALSE ; Silent = $true ; } ;
            Reconnect-Ex2010 @pltRX10 ; # local org conns
            #$pltRx10 creds & .username can also be used for local ADMS connections
            #>
            $pltRX10 = @{
                Credential = (Get-Variable -name "cred$($tenorg)OP" ).value ;
                #verbose = $($verbose) ;
                Verbose = $FALSE ; Silent = $true ; } ;

            # defer cx10/rx10, until just before get-recipients qry
            #endregion GENERIC_EXOP_CREDS_&_SRVR_CONN #*------^ END GENERIC EXOP CREDS & SRVR CONN BP ^------
            # connect to ExOP X10
            if($pltRX10){
                #ReConnect-Ex2010XO @pltRX10 ;
                ReConnect-Ex2010 @pltRX10 ;
            } else { Reconnect-Ex2010 ; } ;
        } ;  # if-E $useEXOP


        #region UseOPAD #*------v UseOPAD v------
        if($UseExOP){
            #region GENERIC_ADMS_CONN_&_XO #*------v GENERIC ADMS CONN & XO  v------
            $smsg = "(loading ADMS...)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            # always capture load-adms return, it outputs a $true to pipeline on success
            $ADMTLoaded = load-ADMS -Verbose:$FALSE ;
            <#
            # resolve $domaincontroller dynamic, cross-org
            # setup ADMS PSDrives per tenant
            if(!$global:ADPsDriveNames){
                $smsg = "(connecting X-Org AD PSDrives)" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                $global:ADPsDriveNames = mount-ADForestDrives -verbose:$($verbose) ;
            } ;
            if(($global:ADPsDriveNames|measure).count){
                $useEXOforGroups = $false ;
                $smsg = "Confirming ADMS PSDrives:`n$(($global:ADPsDriveNames.Name|%{get-psdrive -Name $_ -PSProvider ActiveDirectory} | ft -auto Name,Root,Provider|out-string).trim())" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                # returned object
                #         $ADPsDriveNames
                #         UserName                Status Name
                #         --------                ------ ----
                #         DOM\Samacctname   True  [forestname wo punc]
                #         DOM\Samacctname   True  [forestname wo punc]
                #         DOM\Samacctname   True  [forestname wo punc]

            } else {
                #-=-record a STATUS=-=-=-=-=-=-=
                $statusdelta = ";ERROR";
                $script:PassStatus += $statusdelta ;
                set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta) ;
                #-=-=-=-=-=-=-=-=
                $smsg = "Unable to detect POPULATED `$global:ADPsDriveNames!`n(should have multiple values, resolved to $()"
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                throw "Unable to resolve $($tenorg) `$o365Cred value!`nEXIT!"
                Break ;
            } ;
            #>
            #endregion GENERIC_ADMS_CONN_&_XO #*------^ END GENERIC ADMS CONN & XO ^------
        } ;
        #if (!$domaincontroller) { $domaincontroller = get-gcfast } ;
        #if(!$domaincontroller){ if(test-path function:get-gcfast){$domaincontroller=get-gcfast} else { throw "no get-gcfast()!" } ;} else {"(existing `$domaincontroller:$($domaincontroller))"} ;
        # use new get-GCFastXO cross-org dc finde
        # default to Op_ExADRoot forest from $TenOrg Meta
        if($UseExOP -AND -not $domaincontroller){
            #$domaincontroller = get-GCFastXO -TenOrg $TenOrg -subdomain ((gv -name "$($TenOrg)Meta").value['OP_ExADRoot']) -verbose:$($verbose) |?{$_.length};
            # need to debug the above, credential issue?
            # just get it done
            $domaincontroller = get-GCFast
        } ;
        #endregion UseOPAD #*------^ END UseOPAD ^------

        #region MSOL_CONNECTION ; #*------v  MSOL CONNECTION v------
        $reqMods += "connect-msol".split(";") ;
        if ( !(check-ReqMods $reqMods) ) { write-error "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Missing function. EXITING." ; Break ; }  ;
        $smsg = "(loading AAD...)" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        #connect-msol ;
        connect-msol @pltRXO ;
        #endregion MSOL_CONNECTION ; #*------^  MSOL CONNECTION ^------
        #

        #
        #region AZUREAD_CONNECTION ; #*------v AZUREAD CONNECTION v------
        $reqMods += "Connect-AAD".split(";") ;
        if ( !(check-ReqMods $reqMods) ) { write-error "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Missing function. EXITING." ; Break ; }  ;
        $smsg = "(loading AAD...)" ;
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        #connect-msol ;
        Connect-AAD @pltRXO ;
        #region AZUREAD_CONNECTION ; #*------^ AZUREAD CONNECTION ^------
        #

        <# defined above
        # EXO connection
        $pltRXO = @{
            Credential = (Get-Variable -name cred$($tenorg) ).value ;
            verbose = $($verbose) ; } ;
        #>
        <#
        if($VerbosePreference = "Continue"){
            $VerbosePrefPrior = $VerbosePreference ;
            $VerbosePreference = "SilentlyContinue" ;
            $verbose = ($VerbosePreference -eq "Continue") ;
        } ;
        disconnect-exo ;
        if ($script:useEXOv2) { reconnect-eXO2 @pltRXO }
        else { reconnect-EXO @pltRXO } ;
        # reenable VerbosePreference:Continue, if set, during mod loads
        if($VerbosePrefPrior -eq "Continue"){
            $VerbosePreference = $VerbosePrefPrior ;
            $verbose = ($VerbosePreference -eq "Continue") ;
        } ;
        #>
        #-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
        #endregion SERVICE_CONNECTIONS #*======^ END SERVICE_CONNECTIONS ^======

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
        # you need both a bound $users at the top - to handle explicit assigns add-EXOLicense -users $variable.
        # with a process {} block to handle any pipeline passed input. The pipeline still maps to the bound param: $users, but the e3ntire process{} is run per element, rather than iteratign the internal $users foreach.
        foreach ($usr in $users){

            #-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
            # Looks like 1/5/2022, there are no spare E3's, maybe shift to the F3 ($TORMETA.o365LicSkuF1 = works to get mbx back).
            # (below defaults to the 'office 365 F3', the E3 alt is:  $tormeta.o365LicSkuE3 )
            # 12:06 PM 1/11/2022 add ExOnly: EXCHANGESTANDARD # Office 365 Exchange Online Only (commonly used for App
            $tUPN="$usr" ;
            #$LicenseSkuIds=$TORMETA.o365LicSkuF1;
            $error.clear() ;
            TRY {

                $Exit = 0 ;
                Do {
                    Try {
                        connect-msol @pltRXO;
                        $oMSUsr=$null ;
                        $TenantShortName = ((Get-AzureADTenantDetail).verifieddomains |?{$_._default}).name.split('.')[0] ; 
                        $oMSUsr = get-msoluser -UserPrincipalName $tUPN -EA STOP ; 
                        $pltGAADU=[ordered]@{ ObjectId = $tUPN ; ErrorAction = 'STOP' ; verbose = $($VerbosePreference -eq "Continue") ; } ; 
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

                # confirm/set UsageLoc (reqd for updates)
                if (-not $AADUser.UsageLocation) {
                    $smsg = "AADUser: MISSING USAGELOCATION, FORCING" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug 
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                    $spltSAADUUL = [ordered]@{ 
                        ObjectID = $AADUser.UserPrincipalName ;
                        UsageLocation = "US" ;
                        whatif = $($whatif) ;
                        verbose = ($VerbosePreference -eq "Continue") ;
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
                #if($oMSUsr.isLicensed -eq $true -AND (ps1GetxMbx -id $oMSUsr.UserPrincipalName -ea stop)){
                # nope!: IsLicensed:true, even if nothing but FLOWFREE is set. Worthless, for determining why there's no mailbox. 
                # have to splice over the full exolic-testing code from verb-ex2010:get-mailboxUserStatus():
                
                $pltGLPList=[ordered]@{ TenOrg= $TenOrg; verbose=$($VerbosePreference -eq "Continue") ; credential= $pltRXO.credential ; } ; 

                $skus  = get-AADlicensePlanList @pltGLPList ;

                $ExMbxLicenses = get-ExoMailboxLicenses -verbose:$($VerbosePreference -eq "Continue") ;
                if($skus -AND $ExMbxLicenses){
                    $IsExoLicensed = $false ;
                    foreach($pLic in $AADUser.AssignedLicenses.skuid){
                        if($ExMbxLicenses[$plic]){
                            $IsExoLicensed = $true ;
                            $smsg = "$($mbx.userprincipalname) HAS EXO UserMailbox-supporting License:$($ExMbxLicenses[$sku].SKU)|$($ExMbxLicenses[$sku].Label)" ;
                            write-host "$((get-date).ToString('HH:mm:ss')):$($smsg)"  ;
                            break ; 
                        } ;
                    } ;
                } else { 
                    throw "Unable to resolve either get-AADlicensePlanList OR -ExoMailboxLicenses!" ; 
                    Break ; 
                } ; 

                if((ps1GetxMbx -id $AADUser.UserPrincipalName -ea continue)){
                    $MSOLLicDetails = get-MsolUserLicenseDetails -UPNs $oMSUsr.userprincipalname -showdebug:$($showdebug) -Verbose:$($VerbosePreference -eq "Continue") ;
                    $smsg= "`MSOLLicDetails`n$(($MSOLLicDetails|out-string).trim())" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                } elseif(-not $IsExoLicensed){
                    if($oMSUsr.isLicensed -eq $true){
                        # has a bozo lic that doesn't support a mailbox
                        $smsg = "AADUser:$($tUPN): has .isLicensed:true, but has *NO* EXO UserMailbox-supporting license!" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } 
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                    } ; 
                    $smsg="confirmed $($oMSUsr.UserPrincipalName):is unlicensed/underlicensed" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                    # 9:55 AM 11/15/2019 per Bruce, apply a license, and notify Janel to record
                    #$bRet = add-o365License -MsolUser $oMSUsr -whatif:$($whatif) -showDebug:$($showdebug) -Verbose:$($VerbosePreference -eq "Continue") ;

                    #$smsg = "(Get-MsolAccountSku...)" ; 
                    # get-AADlicensePlanList does a formatted return of Get-AzureADSubscribedSku, returns an indexed hash on skuid, for quick lookups
                    <#$smsg = "(Get-AzureADSubscribedSku...)" ; 
                    $smsg = "(get-AADlicensePlanList...)" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                    #>
                    #$skus = Get-MsolAccountSku -ea STOP ;
                    #$skus = Get-AzureADSubscribedSku -ea STOP ; 
                    #$skus = get-AADlicensePlanList ;
                    <# raw Get-AzureADSubscribedSku return
                        LicenseName     = (Convert-LicenseSku -LicenseSku $License.SkuPartNumber)
                        LicenseSku      = $License.SkuPartNumber
                        LicenseConsumed = $License.ConsumedUnits
                        LicenseTotal    = $License.PrepaidUnits.Enabled
                        LicenseFree     = ($License.PrepaidUnits.Enabled - $License.ConsumedUnits)
                        LicenseEnabled  = if ($License.CapabilityStatus -eq 'Enabled') { $true }else { $false }
                        PlanName        = $Plan.ServicePlanName
                        PlanDescription = (Convert-ServicePlan -ServicePlanID $Plan.ServicePlanID)
                        PlanId          = $Plan.ServicePlanID
                        PlanStatus      = $Plan.ProvisioningStatus
                        AppliesTo       = $Plan.AppliesTo

                    # get-AADlicensePlanList return, with indexed-hashlookup on the SkuPartNumber/plan name (default out: summary table, default index on skuid)
                    $skus['18181a46-0d4e-45cd-891e-60aabd171b4e'] | ft -auto ;
                    SkuId                                SkuPartNumber    Enabled Consumed Available Warning Suspended
                    -----                                -------------    ------- -------- --------- ------- ---------
                    4b9405b0-7788-4568-add1-99614e613b69 EXCHANGESTANDARD      58       53         5       0         0
                    #>

                    <# AADv2 lic assign:
                    # terrible prime doc assign lic example
                    $userUPN="<user sign-in name (UPN)>" ;
                    $planName="<license plan name from the list of license plans>" ;
                    $License = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense ; 
                    $License.SkuId = (Get-AzureADSubscribedSku | Where-Object -Property SkuPartNumber -Value $planName -EQ).SkuID
                    $LicensesToAssign = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
                    $LicensesToAssign.AddLicenses = $License ;
                    Set-AzureADUserLicense -ObjectId $userUPN -AssignedLicenses $LicensesToAssign ;

                    # better route to a named pack lic
                    # get the pack skuid's you want to target:
                    Get-AzureADSubscribedSku | Select Sku*
 
                    SkuId                                SkuPartNumber
                    -----                                -------------
                    6fd2c87f-b296-42f0-b197-1e91e994b900 ENTERPRISEPACK
                    efccb6f7-5641-4e0e-bd10-b4976e1bf68e EMS
                    # new lic obj for ea pack
                    $E3License = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense ;
                    $EMSLicense = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense
                    # set each lic skuid with the matching skuid from the above dump (pulls in the mapping)
                    $E3License.SkuId = "6fd2c87f-b296-42f0-b197-1e91e994b900"
                    $EMSLicense.SkuId = "efccb6f7-5641-4e0e-bd10-b4976e1bf68e"
                    # new lic2assign obj:
                    $LicensesToAssign = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses ;
                    # stack the array of licenses
                    $LicensesToAssign.AddLicenses = $E3License,$EMSLicense ;
                    Set-AzureADUserLicense -ObjectId $User.ObjectId -AssignedLicenses $LicensesToAssign ;
                    Get-AzureADUser -ObjectId $User.ObjectId | Select -ExpandProperty AssignedPlans

                    # loop out a list to assign
                    $tpacks = 'ENTERPRISEPACK','EMS'  ; 
                    $askus = Get-AzureADSubscribedSku ; 
                    $licAssignArray=@() ; 
                    foreach($tpk in $tpacks){
                        $newLicense = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense ;
                        $newLicense.SkuId = $askus|?{$_.skupartnumber -eq 'ENTERPRISEPACK'} | select -expand SkuId ;
                        $licAssignArray+= $newLicense ; 
                    } ; 
                    $LicensesToAssign = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses ;
                    $LicensesToAssign.AddLicenses = $licAssignArray ;
                    Set-AzureADUserLicense -ObjectId $User.ObjectId -AssignedLicenses $LicensesToAssign ;
                    Get-AzureADUser -ObjectId $User.ObjectId | Select -ExpandProperty AssignedPlans
                    
                    # lic removal
                    PS C:\> $User = Get-AzureAdUser -SearchString sue.cooper@exchangeserverpro.net
                    PS C:\> $License = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense
                    PS C:\> $License.SkuId = "6fd2c87f-b296-42f0-b197-1e91e994b900"
                    PS C:\> $LicensesToAssign = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
                    PS C:\> $LicensesToAssign.AddLicenses = @()
                    PS C:\> $LicensesToAssign.RemoveLicenses = $License.SkuId
                    PS C:\> Set-AzureADUserLicense -ObjectId $User.ObjectId -AssignedLicenses $LicensesToAssign

                    #>

                    <# #$LicenseSkuIds=$TORMETA.o365LicSkuF1;
                    toroco:SPE_F1
                    #> 

                    #$pltALic=[ordered]@{ObjectId=$oMSUsr.userprincipalname ; AddLicenses=$null ;} ;
                    # Set-AzureADUserLicense  vers
                    $pltALic=[ordered]@{UserPrincipalName=$oMSUsr.userprincipalname ; AssignedLicenses=$null ;} ;

                    foreach($LicenseSkuId in $LicenseSkuIds){
                        if( $LicenseSkuId.contains(':') ){
                            $LicenseSkuId = $LicenseSkuId.split(':')[1] ;
                            # need the skuid, not the name, could pull another licplan list indiexedonName, but can also post-filter the hashtable, and get it.
                            $LicenseSkuId = ($skus.values | ?{$_.SkuPartNumber -eq $LicenseSkuId}).skuid ; 
                        } ; 
                        $smsg = "(attempting license:$($LicenseSkuId)...)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;

                        #$bRes = add-AADUserLicense -Users $AADUser.UserPrincipalName -skuid $LicenseSkuId -verbose -whatif
                        $pltAAADUL=[ordered]@{
                            Users=$AADUser.UserPrincipalName ;
                            skuid=$LicenseSkuId ;
                            verbose = $($VerbosePreference -eq "Continue") ;
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
                            # update the local AADUser to reflect the updated AADU returned
                            #$AADUser = $bRet.AzureADuser ; 
                            #$Report.FixedUsageLocation = $true ; 
                            BREAK ; # abort further loops if one successfully applied
                        } elseif($whatif){
                            $smsg = "(whatif pass, exec skipped), " ; 
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        } else { 
                            $smsg = "add-AADUserLicense : FAILED TO UPDATE USAGELOCATION!" ;
                            $smsg += "`n$(($bRet|out-string).trim())" ; 
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug 
                            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            #$Report.FixedUsageLocation = $false ; 
                            if(-not $whatif){
                                BREAK; 
                            } 
                        } ; 
                        
                    } ;  # loop-E $LicenseSkuIds

                }



            } CATCH {     $ErrTrapd=$Error[0] ;
                $smsg = "$('*'*5)`nFailed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: `n$(($ErrTrapd|out-string).trim())`n$('-'*5)" ;
                write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" ;
                Break ;
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
                #} Until (ps1GetxMbx -id $oMSUsr.userprincipalname -EA 0) ;
                } Until ($ombx = ps1GetxMbx -id $oMSUsr.userprincipalname -EA 0) ; # capture return (prevent from dropping into pipe)
                $smsg = "xo Mailbox confirmed!" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } ;

            #-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=


            $smsg =  $sBnr.replace('=v','=^').replace('v=','^=') ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        } ; # ($usr in $users)

    } # PROC-E
    END{

        # return $oMSUsr to pipeline if populated

        $oMSUsr | write-output ;
            
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

#*------^ add-EXOLicense.ps1 ^------