#*----------v Function get-ADUsersWithSoftDeletedxoMailboxes() v----------
function get-ADUsersWithSoftDeletedxoMailboxes {
    <#
    .SYNOPSIS
    get-ADUsersWithSoftDeletedxoMailboxes.ps1 - Get *existing* ADUsers with SoftDeleted xoMailboxes
    .NOTES
    Version     : 1.0.0
    Author      : Todd Kadrie
    Website     : http://www.toddomation.com
    Twitter     : @tostka / http://twitter.com/tostka
    CreatedDate : 2022-01-14
    FileName    : get-ADUsersWithSoftDeletedxoMailboxes
    License     : MIT License
    Copyright   : (c) 2021 Todd Kadrie
    Github      : https://github.com/tostka/verb-xo
    Tags        : Powershell
    AddedCredit : REFERENCE
    AddedWebsite: URL
    AddedTwitter: URL
    REVISIONS
    * 2:51 PM 1/14/2022 init
    .DESCRIPTION
    get-ADUsersWithSoftDeletedxoMailboxes.ps1 - Get *existing* ADUsers with SoftDeleted xoMailboxes
    .PARAMETER useEXOv2
    Use EXOv2 (ExchangeOnlineManagement) over basic auth legacy connection [-useEXOv2]
    .INPUTS
    None. Does not accepted piped input.(.NET types, can add description)
    .OUTPUTS
    None. Returns no objects or output (.NET types)
    System.Boolean
    [| get-member the output to see what .NET obj TypeName is returned, to use here]
    .EXAMPLE
    PS> .\get-ADUsersWithSoftDeletedxoMailboxes.ps1 -verbose
    Run with verbose
    .LINK
    https://github.com/tostka/verb-EXO
    #>
    ##Requires -Version 2.0
    #Requires -Version 3
    #requires -PSEdition Desktop
    ##requires -PSEdition Core
    ##Requires -Modules ActiveDirectory, AzureAD, MSOnline, ExchangeOnlineManagement, verb-AAD, verb-ADMS, verb-Auth, verb-Ex2010, verb-EXO, verb-IO, verb-logging, verb-Network, verb-Text
    #Requires -RunasAdministrator
    ##Requires -PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
    ##Requires -Modules ActiveDirectory, AzureAD, MSOnline, ExchangeOnlineManagement, MicrosoftTeams, SkypeOnlineConnector, Lync,  verb-AAD, verb-ADMS, verb-Auth, verb-Azure, VERB-CCMS, verb-Desktop, verb-dev, verb-Ex2010, verb-EXO, verb-IO, verb-logging, verb-Mods, verb-Network, verb-L13, verb-SOL, verb-Teams, verb-Text, verb-logging
    #Requires -Modules ActiveDirectory, ExchangeOnlineManagement, verb-ADMS, verb-Auth, verb-Ex2010, verb-EXO, verb-IO, verb-logging, verb-Network, verb-Text
    # VALIDATORS: [ValidateNotNull()][ValidateNotNullOrEmpty()][ValidateLength(24,25)][ValidateLength(5)][ValidatePattern("some\sregex\sexpr")][ValidateSet("USEA","GBMK","AUSYD")][ValidateScript({Test-Path $_ -PathType 'Container'})][ValidateScript({Test-Path $_})][ValidateRange(21,65)][ValidateCount(1,3)]
    ## [OutputType('bool')] # optional specified output type
    [CmdletBinding()]
    ###[Alias('Alias','Alias2')]
    PARAM(
        [Parameter(HelpMessage="Use EXOv2 (ExchangeOnlineManagement) over basic auth legacy connection [-useEXOv2]")]
        [switch] $useEXOv2
    ) ;
    
    BEGIN { 
        #region CONSTANTS-AND-ENVIRO #*======v CONSTANTS-AND-ENVIRO v======
        # function self-name (equiv to script's: $MyInvocation.MyCommand.Path) ;
        ${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name ;
        # Get parameters this function was invoked with
        $PSParameters = New-Object -TypeName PSObject -Property $PSBoundParameters ;
        $Verbose = ($VerbosePreference -eq 'Continue') ; 
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
        $ComputerName = $env:COMPUTERNAME ;
        $NoProf = [bool]([Environment]::GetCommandLineArgs() -like '-noprofile'); # if($NoProf){# do this};
        # silently stop any running transcripts
        $stopResults = try {Stop-transcript -ErrorAction stop} catch {} ; 
        #endregion CONSTANTS-AND-ENVIRO #*======^ END CONSTANTS-AND-ENVIRO ^======
        
        #region START-LOG #*======v START-LOG OPTIONS v======
        #region START-LOG-HOLISTIC #*------v START-LOG-HOLISTIC v------
        # Single log for script/function example that accomodates detect/redirect from AllUsers scope'd installed code, and hunts a series of drive letters to find an alternate logging dir (defers to profile variables)
        #${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name ;
        if(!(get-variable LogPathDrives -ea 0)){$LogPathDrives = 'd','c' };
        foreach($budrv in $LogPathDrives){if(test-path -path "$($budrv):\scripts" -ea 0 ){break} } ;
        if(!(get-variable rgxPSAllUsersScope -ea 0)){
            $rgxPSAllUsersScope="^$([regex]::escape([environment]::getfolderpath('ProgramFiles')))\\((Windows)*)PowerShell\\(Scripts|Modules)\\.*\.(ps(((d|m))*)1|dll)$" ;
        } ;
        if(!(get-variable rgxPSCurrUserScope -ea 0)){
            $rgxPSCurrUserScope="^$([regex]::escape([Environment]::GetFolderPath('MyDocuments')))\\((Windows)*)PowerShell\\(Scripts|Modules)\\.*\.(ps((d|m)*)1|dll)$" ;
        } ;
        $pltSL=[ordered]@{Path=$null ;NoTimeStamp=$false ;Tag=$null ;showdebug=$($showdebug) ; Verbose=$($VerbosePreference -eq 'Continue') ; whatif=$($whatif) ;} ;
        $pltSL.Tag = $ModuleName ; 
        if($script:PSCommandPath){
            if(($script:PSCommandPath -match $rgxPSAllUsersScope) -OR ($script:PSCommandPath -match $rgxPSCurrUserScope)){
                $bDivertLog = $true ; 
                switch -regex ($script:PSCommandPath){
                    $rgxPSAllUsersScope{$smsg = "AllUsers"} 
                    $rgxPSCurrUserScope{$smsg = "CurrentUser"}
                } ;
                $smsg += " context script/module, divert logging into [$budrv]:\scripts" 
                write-verbose $smsg  ;
                if($bDivertLog){
                    if((split-path $script:PSCommandPath -leaf) -ne $cmdletname){
                        # function in a module/script installed to allusers|cu - defer name to Cmdlet/Function name
                        $pltSL.Path = (join-path -Path "$($budrv):\scripts" -ChildPath "$($cmdletname).ps1") ;
                    } else {
                        # installed allusers|CU script, use the hosting script name
                        $pltSL.Path = (join-path -Path "$($budrv):\scripts" -ChildPath (split-path $script:PSCommandPath -leaf)) ;
                    }
                } ;
            } else {
                $pltSL.Path = $script:PSCommandPath ;
            } ;
        } else {
            if(($MyInvocation.MyCommand.Definition -match $rgxPSAllUsersScope) -OR ($MyInvocation.MyCommand.Definition -match $rgxPSCurrUserScope) ){
                 $pltSL.Path = (join-path -Path "$($budrv):\scripts" -ChildPath (split-path $script:PSCommandPath -leaf)) ;
            } elseif(test-path $MyInvocation.MyCommand.Definition) {
                $pltSL.Path = $MyInvocation.MyCommand.Definition ;
            } elseif($cmdletname){
                $pltSL.Path = (join-path -Path "$($budrv):\scripts" -ChildPath "$($cmdletname).ps1") ;
            } else {
                $smsg = "UNABLE TO RESOLVE A FUNCTIONAL `$CMDLETNAME, FROM WHICH TO BUILD A START-LOG.PATH!" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Warn } #Error|Warn|Debug 
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                BREAK ;
            } ; 
        } ;
        write-verbose "start-Log w`n$(($pltSL|out-string).trim())" ; 
        $logspec = start-Log @pltSL ;
        $error.clear() ;
        TRY {
            if($logspec){
                $logging=$logspec.logging ;
                $logfile=$logspec.logfile ;
                $transcript=$logspec.transcript ;
                $stopResults = try {Stop-transcript -ErrorAction stop} catch {} ;
                start-Transcript -path $transcript ;
            } else {throw "Unable to configure logging!" } ;
        } CATCH {
            $ErrTrapd=$Error[0] ;
            $smsg = "Failed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: $($ErrTrapd)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        } ;
        #endregion START-LOG-HOLISTIC #*------^ END START-LOG-HOLISTIC ^------
        #region START-LOG-SIMPLE #*------v START-LOG-SIMPLE v------
        #Configure default logging from parent script name
        # Configure default logging from parent script name
        $pltSL=@{ NoTimeStamp=$true ; Tag="($TenOrg)-LASTPASS" ; showdebug=$($showdebug) ; whatif=$($whatif) ; Verbose=$($VerbosePreference -eq 'Continue') ; } ;
        if($PSCommandPath){   $logspec = start-Log -Path $PSCommandPath @pltSL ;
        } else { $logspec = start-Log -Path ($MyInvocation.MyCommand.Definition) @pltSL ; } ;
        if($logspec){
            $logging=$logspec.logging ;
            $logfile=$logspec.logfile ;
            $transcript=$logspec.transcript ;
            if(Test-TranscriptionSupported){
                $stopResults = try {Stop-transcript -ErrorAction stop} catch {} ;
                start-transcript -Path $transcript ;
            } ;
        } else {throw "Unable to configure logging!" } ;
        #endregion START-LOG-SIMPLE #*------^ END START-LOG SIMPLE ^------
        #endregion START-LOG #*======^ START-LOG OPTIONS ^======
        
        #region EXCH-CMD-ALIASING-OPTS ; #*======v EXCH-CMD-ALIASING-OPTS v======
        #region EXO-v-EXOv2-ALIASING #*------v Function EXO v EXOv2 ALIASING on $useEXOv2 v------
        # simple loop to stock the set, no set->get conversion, roughed in $Exov2 exo->xo replace. Do specs in exo, and flip to suit under $exov2
        #configure EXO EMS aliases to cover useEXOv2 requirements
        # have to preconnect, as it gcm's the targets
        if ($script:useEXOv2) { reconnect-eXO2 }
        else { reconnect-EXO } ;
        # aliased ExOP|EXO|EXOv2 cmdlets (permits simpler single code block for any of the three variants of targets & syntaxes)
        # each is '[aliasname];[exOcmd] (xOv2cmd & exop are converted from [exocmd])
        [array]$cmdletMaps = 'ps1GetxRcp;get-exorecipient;','ps1GetxMbx;get-exomailbox;','ps1SetxMbx;Set-exoMailbox;','ps1GetxUser;get-exoUser;',
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
        # cleanup
        get-alias -scope Script |?{$_.name -match '^ps1.*'} | %{Remove-Alias -alias $_.name} ; 
        #endregion EXO-v-EXOv2-ALIASING #*------^ END Function EXO V EXOv2 ALIASING ^------
        
        #endregion EXCH-CMD-ALIASING-OPTS #*======^ END EXCH-CMD-ALIASING-OPTS ^======
        
        #region useEXOP ; #*------v useEXOP v------
        $useEXOP = $false ; 
        if($useEXOP){
            #*------v GENERIC EXOP CREDS & SRVR CONN BP v------
            # do the OP creds too
            $OPCred=$null ;
            # default to the onprem svc acct
            $pltGHOpCred=@{TenOrg=$TenOrg ;userrole='ESVC' ;verbose=$($verbose)} ;
            if($OPCred=(get-HybridOPCredentials @pltGHOpCred).cred){
                # make it script scope, so we don't have to predetect & purge before using new-variable
                New-Variable -Name "cred$($tenorg)OP" -scope Script -Value $OPCred ;
                $smsg = "Resolved $($Tenorg) `$OPCred:$($OPCred.username) (assigned to `$cred$($tenorg)OP)" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } else {
                #-=-record a STATUSERROR=-=-=-=-=-=-=
                $statusdelta = ";ERROR"; # CHANGE|INCOMPLETE|ERROR|WARN|FAIL ;
                if(gv passstatus -scope Script){$script:PassStatus += $statusdelta } ;
                if(gv -Name PassStatus_$($tenorg) -scope Script){set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta)} ; 
                #-=-=-=-=-=-=-=-=
                $smsg = "Unable to resolve get-HybridOPCredentials -TenOrg $($TenOrg) -userrole 'ESVC' value!"
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                throw "Unable to resolve $($tenorg) `$OPCred value!`nEXIT!"
                BREAK ;
            } ;
            $smsg= "Using EXOP cred:`$cred$($tenorg)OP:$((Get-Variable -name "cred$($tenorg)OP" ).value.username)" ;  
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
            <# CALLS ARE IN FORM: (cred$($tenorg))
            $pltRX10 = @{
                Credential = (Get-Variable -name "cred$($tenorg)OP" ).value ;
                verbose = $($verbose) ; }
            ReConnect-Ex2010XO @pltRX10 ; # cross-prem conns
            Reconnect-Ex2010 @pltRX10 ; # local org conns
            #$pltRx10 creds & .username can also be used for local ADMS connections
            #>
            $pltRX10 = @{
                Credential = (Get-Variable -name "cred$($tenorg)OP" ).value ;
                verbose = $($verbose) ; } ;     
            # TEST
            #if( !(check-ReqMods $reqMods) ) {write-error "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Missing function. EXITING." ; BREAK ;}  ;
            # defer cx10/rx10, until just before get-recipients qry
            #*------^ END GENERIC EXOP CREDS & SRVR CONN BP ^------
            # connect to ExOP X10
            <#
            if($pltRX10){
                ReConnect-Ex2010 @pltRX10 ;
            } else { Reconnect-Ex2010 ; } ; 
            #>
        } ;  # if-E $useEXOP
        #endregion useEXOP ; #*------^ END useEXOP ^------
        #region useOPAD ; #*------v useOPAD v------
        if($useEXOP){
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
                #-=-record a STATUSERROR=-=-=-=-=-=-=
                $statusdelta = ";ERROR"; # CHANGE|INCOMPLETE|ERROR|WARN|FAIL ;
                if(gv passstatus -scope Script){$script:PassStatus += $statusdelta } ;
                if(gv -Name PassStatus_$($tenorg) -scope Script){set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatus_$($tenorg)).value + $statusdelta)} ; 
                #-=-=-=-=-=-=-=-=
                $smsg = "Unable to detect POPULATED `$global:ADPsDriveNames!`n(should have multiple values, resolved to $()"
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                throw "Unable to resolve $($tenorg) `$o365Cred value!`nEXIT!"
                BREAK ;
            } ; 
        }  else { 
            <# have to defer to get-azuread, or use EXO's native cmds to poll grp members
            # TODO 1/15/2021
            $useEXOforGroups = $true ; 
            $smsg = "$($TenOrg):HAS NO ON-PREM ACTIVEDIRECTORY, DEFERRING ALL GROUP ACCESS & MGMT TO NATIVE EXO CMDS!" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            #>
        } ; 
        #endregion useOPAD ; #*------^ END useOPAD ^------

        <#
        if($pltRX10){
            ReConnect-Ex2010 @pltRX10 ;
        } else { Reconnect-Ex2010 ; } ;     
        #>
        if ($PSCmdlet.MyInvocation.ExpectingInput) {
            $smsg = "Data received from pipeline input: '$($InputObject)'" ; 
        } else {
            #write-verbose "Data received from parameter input: '$($InputObject)'" ; 
        } ; 
        
        #### NEW CODE/CONSTANTS HERE ####
        
    } ; # BEGIN-E
    PROCESS {
        $Error.Clear() ; 
        # call func with $PSBoundParameters and an extra (includes Verbose)
        #call-somefunc @PSBoundParameters -anotherParam
        
        # - Pipeline support will iterate the entire PROCESS{} BLOCK, with the bound - $array - 
        #   param, iterated as $array=[pipe element n] through the entire inbound stack. 
        # $_ within PROCESS{}  is also the pipeline element (though it's safer to declare and foreach a bound $array param).
        
        # - foreach() below alternatively handles _named parameter_ calls: -array $objectArray
        # which, when a pipeline input is in use, means the foreach only iterates *once* per 
        #   Process{} iteration (as process only brings in a single element of the pipe per pass) 
        
        #foreach($item in $array) {
            # dosomething w $item
            
            # put your real processing in here, and assume everything that needs to happen per loop pass is within this section.
            # that way every pipeline or named variable param item passed will be processed through. 
            
            $smsg = "getting *existing* ADUsers with SoftDeletedxoMailboxes" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
            
            $smsg = "(get all SoftDeleted xoMbxs)" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
            $error.clear() ;
            TRY {
                $pltGxMbx=[ordered]@{ Resultsize='Unlimited' ;SoftDeletedMailbox=$true ;ErrorAction = 'STOP';} ; 
                $smsg = "$((get-alias ps1GetxMbx).definition) w`n$(($pltGxMbx|out-string).trim())" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                $allsdmbx = ps1GetxMbx @pltGxMbx ;
                $smsg = "(get all LegalHeld mailboxes (InactiveMailboxOnly))" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                $pltGxMbx=[ordered]@{ Resultsize='Unlimited' ;InactiveMailboxOnly=$true ;ErrorAction = 'STOP';} ; 
                $smsg = "$((get-alias ps1GetxMbx).definition) w`n$(($pltGxMbx|out-string).trim())" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                $allimbx = ps1GetxMbx @pltGxMbx ;
                $smsg = "(compare the populations)" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                $pltComp=[ordered]@{ReferenceObject=$allsdmbx ;DifferenceObject=$allimbx ;PassThru=$true;Property='userprincipalname' ;} ; 
                $smsg = "compare-object w`n$(($pltComp|out-string).trim())" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                $cmpare = compare-object @pltComp ;
                $smsg = "(isolate all SoftDeleted mbxs that are *not* Inactive/legal-held)" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                $nonHeldSDs = $allsdmbx | ? isinactivemailbox -eq $false | sort whensoftdeleted ;  
                $smsg = "Filter for non-LegalHeld SoftDeleted mailboxes, with non-Deleted ADUsers...`n" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                $availSoftDeleteADUsers = $nonHeldSDs | %{$upn = $_.userprincipalname ; get-adobject -filter 'userprincipalname -eq  $upn' -IncludeDeletedObjects -properties IsDeleted,LastKnownParent,userprincipalname -ea continue} |?{$_.isdeleted -eq $false } ; 
                if($availSoftDeleteADUsers){ 
                    #$availSoftDeleteADUsers | ft -auto name,IsDeleted,lastknownparent,userp*
                    $smsg = "`n$(($availSoftDeleteADUsers | ft -auto name,IsDeleted,lastknownparent,userp*|out-string).trim())" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    $smsg = "(returning matches to pipeline)" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    $availSoftDeleteADUsers | write-output ; 
                } else {
                    $smsg = "`$availSoftDeleteADUsers: none found" 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                } ; 
            } CATCH {
                $ErrTrapd=$Error[0] ;
                $smsg = "$('*'*5)`nFailed processing $($ErrTrapd.Exception.ItemName). `nError Message: $($ErrTrapd.Exception.Message)`nError Details: `n$(($ErrTrapd|out-string).trim())`n$('-'*5)" ;
                write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)"  ;
                Break ; 
            } ; 
        #} ;  # loop-E process loop

    } ;  # PROC-E
    END {
        # clean-up dyn-created vars & those created by a dot sourced script.
        #((Compare-Object -ReferenceObject (Get-Variable).Name -DifferenceObject $DefVaris).InputObject).foreach{Remove-Variable -Name $_} ; 
    } ;  # END-E
} ; 
#*------^ END Function get-ADUsersWithSoftDeletedxoMailboxes() ^------



