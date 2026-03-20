# Update-EXOLinkedHybridObjectsTDO.ps1
# Update-LinkHybridObjects.ps1

function Update-EXOLinkedHybridObjectsTDO {
    <#
    .SYNOPSIS
    Update-EXOLinkedHybridObjectsTDO - Checks an EXO mailbox for critical hybrid matches: mgUser,ADUser,HardMatch,RemoteMailbox,ExGuidMatch, creates ADUser and RemoteMailbox if possible, Updates ImmutableID and ExchangeGuid to bring into sync.
    .NOTES
    Version     : 0.0.1
    Author      : Todd Kadrie
    Website     : http://www.toddomation.com
    Twitter     : @tostka / http://twitter.com/tostka
    CreatedDate : 2026-
    FileName    : Update-EXOLinkedHybridObjectsTDO.ps1
    License     : MIT License
    Copyright   : (c) 2026 Todd Kadrie
    Github      : https://github.com/tostka/verb-XXX
    Tags        : Powershell
    AddedCredit : REFERENCE
    AddedWebsite: URL
    AddedTwitter: URL
    REVISIONS
    * 1:52 PM 3/20/2026 latest rev, mostly func, missing some outlier parts, did work past the update-mguser scope block. Needs more debugging
    * 5:36 PM 3/16/2026 init
    .DESCRIPTION

    # NOTE: UPDATE-MGUSER -onpremisesimmutableid requires beyond global defaults: throws 'Access Denied' unless scopes includes Directory.AccessAsUser.All
    which reflects performing anything the user themselves can do.

    "Directory.AccessAsUser.All",'Directory.ReadWrite.All','User.ReadWrite.All','AuditLog.Read.All','openid','profile','User.Read','User.Read.All','email' ;

    Tried backing out MG scopes:
    PS> $MGScopesRqd = verb-mg\get-MGCodeCmdletPermissionsTDO -path D:\scripts\Update-EXOLinkedHybridObjectsTDO.ps1 ;

    [How do i run a bulk update for the 'Employee Type' and 'Employee Hire Date' attribute for users in you tenant using a CSV file in MS graph PowerShell - Microsoft Q&A](https://learn.microsoft.com/en-us/answers/questions/1471000/how-do-i-run-a-bulk-update-for-the-employee-type-a)
    says: 
    ```
    Update-MgUser is a graph command and based upon the link you have attached is meant for cloud users only, could you please check if the user is cloud user only. if not, you can use permissions Directory.ReadWrite.All, User.ReadWrite.All (application)

    In case nothing works, maybe you can also give a try to use the following permission.
    ```
    => Nothing except  Directory.AccessAsUser.All actually works, even for a Global Admin. 

    .PARAMETER ThisXoMbx
    Mailbox Object to be checked
    .PARAMETER ca5
    CustomAttribute5 Update Value
    .PARAMETER ticket
    TicketNumber
    .PARAMETER waitPostChangeSecs
    Seconds to wait after change made, to refresh object
    .PARAMETER emlCoTag
    Tag string to be appended to conflicting email address objects, to create unique addresses
    .PARAMETER RequiredScopes
    Scopes required for planned cmdlets to be executed[-RequiredScopes @('User.Read.All', 'Group.Read.All', 'Domain.Read.All')]
    .PARAMETER Force
    Force (Confirm-override switch, overrides ShouldProcess testing, executes somewhat like legacy -whatif:`$false)[-force]
    .INPUTS
    None. Does not accepted piped input.(.NET types, can add description)
    .OUTPUTS
    None. Returns no objects or output    
    .EXAMPLE
    PS> .\Update-EXOLinkedHybridObjectsTDO.ps1 -whatif -verbose
    EXSAMPLEOUTPUT
    Run with whatif & verbose
    .EXAMPLE
    PS> .\Update-EXOLinkedHybridObjectsTDO.ps1
    EXSAMPLEOUTPUT
    EXDESCRIPTION
    .LINK
    https://github.com/tostka/verb-XXX
    .LINK
    https://github.com/tostka/powershellbb/
    .LINK
    [ name related topic(one keyword per topic), or http://|https:// to help, or add the name of 'paired' funcs in the same niche (enable/disable-xxx)]
    #>
    [CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact = 'High')]
    #[CmdletBinding()]
    PARAM(
        [Parameter(Mandatory=$TRUE,HelpMessage="Mailbox Object to be checked")]
            [psobject]$ThisXoMbx,
        [Parameter(HelpMessage="CustomAttribute5 Update Value")]
            [string]$ca5 = 'Spartanmowers',
        [Parameter(HelpMessage="TicketNumber")]
            [string]$ticket = 'RFC15319',
        #[Parameter(HelpMessage="Domain Name to be used for constructing ADUser proxyAddress Additions (for missing onmicrosoft.com")]
        #    [string]$newdom = 'spartanmowers.com',
        [Parameter(HelpMessage="Seconds to wait after change made, to refresh object")]
            [int]$waitPostChangeSecs = 30,
        [Parameter(HelpMessage="Tag string to be appended to conflicting email address objects, to create unique addresses")]
            [string]$emlCoTag = '-INT',
        [Parameter(Mandatory=$False,HelpMessage="Scopes required for planned cmdlets to be executed[-RequiredScopes @('User.Read.All', 'Group.Read.All', 'Domain.Read.All')]")]
            [Alias('scopes')] # alias the connect-mggraph underlying param, for passthru
            [array]$RequiredScopes = @("Directory.AccessAsUser.All",'Directory.ReadWrite.All','User.ReadWrite.All','AuditLog.Read.All','openid','profile','User.Read','User.Read.All','email'),
        [Parameter(HelpMessage="Force (Confirm-override switch, overrides ShouldProcess testing, executes somewhat like legacy -whatif:`$false)[-force]")]
            [switch]$Force
        #[switch]$WHATIF = $true # whatif is implied by SSP, throws: A parameter with the name 'WhatIf' was defined multiple times for the command. if both $whatif param and SupportsShouldProcess=$true
    )
    BEGIN{
        # resolved functional, per [How do i run a bulk update for the 'Employee Type' and 'Employee Hire Date' attribute for users in you tenant using a CSV file in MS graph PowerShell - Microsoft Q&A](https://learn.microsoft.com/en-us/answers/questions/1471000/how-do-i-run-a-bulk-update-for-the-employee-type-a)
        #$RequiredScopes = "Directory.AccessAsUser.All",'Directory.ReadWrite.All','User.ReadWrite.All','AuditLog.Read.All','openid','profile','User.Read','User.Read.All','email' ; 
        # non-func list using get-mgcommandlet permissions: bogus. 
        #'User.ReadBasic.All','User.ReadWrite.All','DeviceManagementApps.ReadWrite.All','User.Read.All','Directory.ReadWrite.All','Directory.Read.All','DeviceManagementServiceConfig.ReadWrite.All','User.ReadWrite.CrossCloud','DeviceManagementManagedDevices.ReadWrite.All','DeviceManagementManagedDevices.Read.All','DeviceManagementConfiguration.ReadWrite.All','DeviceManagementConfiguration.Read.All','DeviceManagementServiceConfig.Read.All','DeviceManagementApps.Read.All','User.ReadWrite','User-Mail.ReadWrite.All','User-PasswordProfile.ReadWrite.All','User-Phone.ReadWrite.All','User.EnableDisableAccount.All','User.ManageIdentities.All' ; 
        #region ENVIRO_DISCOVER ; #*------v ENVIRO_DISCOVER v------
        #push-TLSLatest
        $Verbose = [boolean]($VerbosePreference -eq 'Continue') ; 
        # USE: -whatif:$($whatifswitch) # proxy for -whatif when SupportsShouldProcess
        [boolean]$whatIfSwitch = ($WhatIf.IsPresent -or $whatif -eq $true -OR $WhatIfPreference -eq $true);  $smsg = "-Verbose:$($Verbose)`t-Whatif:$($whatifswitch) " ;  write-host -foregroundcolor yellow $smsg 
        $rPSCmdlet = $PSCmdlet ; # an object that represents the cmdlet or advanced function that's being run. Available on functions w CmdletBinding (& $args will not be available). (Blank on non-CmdletBinding/Non-Adv funcs).
        $rPSScriptRoot = $PSScriptRoot ; # the full path of the executing script's parent directory., PS2: valid only in script modules (.psm1). PS3+:it's valid in all scripts. (Funcs: ParentDir of the file that hosts the func)
        $rPSCommandPath = $PSCommandPath ; # the full path and filename of the script that's being run, or file hosting the funct. Valid in all scripts.
        $rMyInvocation = $MyInvocation ; # populated only for scripts, function, and script blocks.
        # - $MyInvocation.MyCommand.Name returns name of a function, to identify the current command,  name of the current script (pop'd w func name, on Advfuncs)
        # - Ps3+:$MyInvocation.PSScriptRoot : full path to the script that invoked the current command. The value of this property is populated only when the caller is a script (blank on funcs & Advfuncs)
        # - Ps3+:$MyInvocation.PSCommandPath : full path and filename of the script that invoked the current command. The value of this property is populated only when the caller is a script (blank on funcs & Advfuncs)
        #     ** note: above pair contain information about the _invoker or calling script_, not the current script
        $rPSBoundParameters = $PSBoundParameters ; 
        #region LOAD_MODS ; #*------v LOAD_MODS v------
        if(-not $tDepModules){$tDepModules = @('ExchangeOnlineManagement','ActiveDirectory','Microsoft.Graph.Users','VERB-MG','verb-Ex2010','verb-EXO','verb-IO','verb-logging','verb-Text') } ; 
        $tDepModules |%{$tmod = $_ ; if(-not (get-module $tmod -ea 0 )){ write-verbose "Loading missing module:$($tmod)" ; Import-Module $tmod -ea STOP ; } ; }; get-variable $tmod -ea 0| remove-variable $tmod ;
        #endregion LOAD_MODS ; #*------^ END LOAD_MODS ^------

        TRY{
            #if([system.io.fileinfo]$rPSCmdlet.MyInvocation.MyCommand.Source){
            if($ps1Path = [system.io.fileinfo]$rMyInvocation.mycommand.definition){
                $transcript = (join-path -path $ps1Path.DirectoryName -ChildPath 'logs') ; 
                if(-not (test-path $transcript  -PathType Container -ea 0)){ mkdir $transcript -verbose }
                $transcript = join-path -path $transcript -childpath $ps1Path.BaseName ;                 
            }else{$throw} ;
        }CATCH{
            if($rPSCmdlet.MyInvocation.InvocationName){
                if(gcm gcim -ea 0){$drvs = gcim Win32_LogicalDisk }elseif(gcm gwmi -ea 0){$drvs = gwmi Win32_LogicalDisk} 
                if($drvs = $drvs |?{$_.deviceid -match '[A-Z]:' -AND $_.drivetype -eq 3}){
                    foreach($logdrive in @('D:','C:')){if($drvs |?{$_.deviceid -eq $logdrive}){break} } ; 
                }else{write-warning "unable to gcim/gwmi Win32_LogicalDisk class!" ; break } ; 
                $transcript = (join-path -path (join-path -path $logdrive -ChildPath 'scripts') -childpath 'logs') ; 
                if(-not (test-path $transcript  -PathType Container -ea 0)){ mkdir $transcript -verbose }
                $transcript = join-path -path $transcript -childpath $rPSCmdlet.MyInvocation.InvocationName ;                 
            } ELSE{
                $smsg = "FUNCTION: Unable to resolve a the function name (blank `$rPSCmdlet.MyInvocation.InvocationName)" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                BREAK ; 
            }; 
        }
        if($ticket){$transcript += "-$($ticket)" }
        if($whatif -OR $WhatIf.IsPresent){$transcript += "-WHATIF"}ELSE{$transcript += "-EXEC"} ; 
        if($thisXoMbx.userprincipalname){$transcript += "-$($thisXoMbx.userprincipalname)"} ; 
        $transcript += "-$(get-date -format 'yyyyMMdd-HHmmtt')-trans-log.txt" ;
        #$transcript = "d:\scripts\24381-CA5-$($ca5)-updates-$(get-date -format 'yyyyMMdd-HHmmtt')-trans-log.txt" ;
        $stopResults = try {Stop-transcript -ErrorAction stop} CATCH {} ;
        if($stopResults){
            $smsg = "Stop-transcript:$($stopResults)" ;
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
        } ;
        $startResults = start-Transcript -path $transcript -whatif:$false -confirm:$false;
        if($startResults){
            $smsg = "start-transcript:$($startResults)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        } ;        

        $VerbosePreference = 'Continue'
        #region SVCCONN_LITE ; #*------v SVCCONN_LITE v------
        $isXoConn = [boolean]( (gcm Get-ConnectionInformation -ea 0) -AND (Get-ConnectionInformation -ea 0 |?{$_.State -eq 'Connected' -AND $_.TokenStatus -eq 'Active'})) ; if(-not $isXoConn){Connect-EXO}else{write-verbose "EXO connected"};
        #region MG_CONNECT ; #*------v MG_CONNECT v------
        #$isMgConn = [boolean]( (gcm get-mgcontext -ea 0) -AND (get-mgcontext -ea 0 )); if(-not $isMgConn ){connect-mg }else{write-verbose "MG connected"};
        #$RequiredScopes = "Directory.AccessAsUser.All",'Directory.ReadWrite.All','User.ReadWrite.All','AuditLog.Read.All','openid','profile','User.Read','User.Read.All','email' ;
        if(-not (get-command  test-mgconnection)){
            if(-not (get-module -list Microsoft.Graph -ea 0)){
                $smsg = "MISSING Microsoft.Graph!" ;
                $smsg += "`nUse: install-module Microsoft.Graph -scope CurrentUser" ;
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } ;
        } ;
        $MGConn = test-mgconnection -Verbose:($VerbosePreference -eq 'Continue') ;
        if($RequiredScopes){$addScopes = @() ;$RequiredScopes |foreach-object{ $thisPerm = $_ ;if($mgconn.scopes -contains $thisPerm){write-verbose "has scope: $($thisPerm)"} else{$addScopes += @($thisPerm) ; write-verbose "ADD scope: $($thisPerm)"} } ;} ;
        $pltCcMG = [ordered]@{NoWelcome=$true; ErrorAction = 'STOP'}
        if($addScopes){ $pltCcMG.add('RequiredScopes',$addscopes); $pltCcMG.add('ContextScope','Process'); $pltCCMG.add('silent',$false); write-verbose "Adding non-default Scopes, setting non-persistant single-process ContextScope"  } ; 
        if($MGConn.isConnected -AND $addScopes -AND $mgconn.CertificateThumbprint){
            $smsg = "CBA cert lacking scopes :$($addscopes -join ',')!"  ;  $smsg += "`nDisconnecting to use interactive connection: connect-mg -RequiredScopoes `"'$($addscopes -join "','")'`"" ; $smsg += "`n(alt: : connect-mggraph -Scopes `"'$($addscopes -join "','")'`" )" ; write-warning $smsg ; 
            disconnect-mggraph ; 
        }elseif($MGConn.isConnected -AND $addScopes -and -not ($mgconn.CertificateThumbprint)){
        }elseif(-NOT ($MGConn.isConnected) -AND $addScopes -and -not ($mgconn.CertificateThumbprint)){$pltCCMG.add('Credential',$credO365TORSID)            
        }else {write-verbose "(currently connected with any specifically specified required Scopes)"
            $pltCcMG = $null ; 
        }
        if($pltCcMG){
            $smsg = "connect-mg w`n$(($pltCCMG.getenumerator() | ?{$_.name -notmatch 'requiredscopes'} | ft -a | out-string|out-string).trim())" ;
            $smsg += "`n`n-requiredscopes:`n$(($pltCCMG.requiredscopes|out-string).trim())`n" ;
            if($silent){} else {
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            } ; 
            connect-mg @pltCCMG ;
        } ; 
        if(-not (get-command Get-MgUser)){
            $smsg = "Missing Get-MgUser!" ;
            $smsg += "`nPre-connect to Microsoft.Graph via:" ;
            $smsg += "`nConnect-MgGraph -Scopes `'$($requiredscopes -join "','")`'" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            BREAK ;
        } ;
        #endregion MG_CONNECT ; #*------^ END MG_CONNECT ^------
        $isXoPConn = [boolean]( (gcm get-pssession -ea 0) -AND (get-pssession -ea 0 |?{$_.State -eq 'Opened' -AND $_.Availability -eq 'Available'})); if(-not $isXoPConn){Reconnect-Ex2010}else{write-verbose "XOP connected"};
        $isADConn = [boolean](gcm get-aduser -ea 0) ; if(!$isADConn){$env:ADPS_LoadDefaultDrive = 0 ; $sName="ActiveDirectory"; if (!(Get-Module | where {$_.Name -eq $sName})) {Import-Module $sName -ea Stop}}else{write-verbose "ADMS connected"};
        #endregion SVCCONN_LITE ; #*------^ END SVCCONN_LITE ^------

        #region UPDATE_MGUIMMUT ; #*------v Update-MGUIMmut v------
        Function Update-MGUIMmut{
            <# call: Update-MGUIMmut -thisMGU $thisMGU -thisADU $thisADU -Verbose:$($VerbosePreference -eq 'Continue') -whatif:$($whatifswitch) 
            #>
            [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'HIGH')] 
            PARAM(
                [Parameter(Mandatory=$true)]
                    $thisMGU, 
                [Parameter(Mandatory=$true)]
                    $thisADU
            ) ; 
            TRY{
                $OpImmutableId = [System.Convert]::ToBase64String($ThisADU.ObjectGUID.ToByteArray())
                $pltUdMgu = [ordered]@{
                    UserId = $ThisMgu.Id ;
                    OnPremisesImmutableId = $OpImmutableId ;
                    ErrorAction = 'Stop' ;
                    WhatIf = $whatIfSwitch ;
                }
                write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):Update-MgUser w`n$(($pltUdMgu|out-string).trim())" ;
                if ($Force -or $PSCmdlet.ShouldProcess($ThisMgu.displayname, "Update-MgUser")) {
                    #Update-MgUser -UserId $ThisMgu.Id -OnPremisesImmutableId $OpImmutableId -ErrorAction Stop -WhatIf:$whatIfSwitch
                    Update-MgUser @pltUdMgu ;
                    $doUpdtMGUOnPremImmut = $true ;
                    write-verbose "refresh thisMGU" ;
                    $ThisMgu = Get-MgUserFull -UserId $ThisXoMbx.ExternalDirectoryObjectId -ErrorAction Stop
                    $hasImmutSync = $true ;
                    $ThisMgu | write-output  ; 
                } else {
                    Write-Host "(-Whatif or `"No`" to the prompt)"
                    $ThisMgu | write-output  ; 
                } ; 
            } CATCH [System.Exception] {
                $ErrTrapd=$Error[0] ;
                if($ErrTrapd.ErrorDetails -match 'Insufficient\sprivileges\sto\scomplete\sthe\soperation'){
                    $smsg = "Update-MgUse: PERMISSION FAIL!" ;
                    $smsg += "`nIF USING ISE, RUN THE CMD OUTSIDE OF ISE IN A FULL PS SESSION!" ;
                    $smsg += "`n`nPS> Update-MgUser -UserId $($ThisMgu.Id) -OnPremisesImmutableId $($OpImmutableId) -ErrorAction Stop -WhatIf"
                    write-warning $smsg ;
                    write-warning $smsg ;
                }else{
                    THROW $ErrTrapd
                } ;
            }CATCH {                  
                #write-warning "POPULATED onPremisesImmutableId and no MGUser.UPN matching ADuser"
                $ErrTrapd=$Error[0] ;
                write-host -foregroundcolor gray "TargetCatch:} CATCH [$($ErrTrapd.Exception.GetType().FullName)] {"  ;
                $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" ;
            }
        } ;
        #endregion UPDATE_MGUIMMUT ; #*------^ END Update-MGUIMmut ^------

    } ;  # BEG-E
    PROCESS{
        $hasXoMbx = $false ;
        $hasXoMbxDirSync = $false ; 
        $hasMgUser = $false ;
        $hasADUser = $false ; 
        $hasImmutSync = $false ; 
        $hasRmbx = $false ; 
        $hasRmbxExGuidMatch = $false ; 
    
        $doAddADUser = $false ; 
        $doUpdtMGUOnPremImmut = $false ; 
        $doAddRmbx = $false ; 
        $doRmbxExGuidMatch = $false ; 

        # 0. Resolve xombx to live object (csv collapsed)
        try {
            $DC = GET-GCFAST ;
            $ThisXoMbx = get-xomailbox -id $ThisXoMbx.ExchangeGuid -ea STOP ; 
            $hasXoMbx = $true ;
        }CATCH {
            write-host "Failed to resolve MG user for ExternalDirectoryObjectId: $($_.Exception.Message)"
            throw
        }
        # 1. Resolve MG User from ExternalDirectoryObjectId
        try {
            if($ThisMgu = Get-MgUserFull -UserId $ThisXoMbx.ExternalDirectoryObjectId -ErrorAction Stop){
            $hasMgUser = $true ;}
        }CATCH {
            write-host "Failed to resolve MG user for ExternalDirectoryObjectId: $($_.Exception.Message)"
            throw
        }

        # 2. If onPremisesImmutableId exists, convert base64 -> GUID, then find AD user
        $ThisOnPremisesImmutableId = (Get-MgUser -UserId $ThisXoMbx.ExternalDirectoryObjectId -Property OnPremisesImmutableId -ErrorAction Stop | Select-Object -ExpandProperty OnPremisesImmutableId) 

        $ThisADU = $null
        if ($ThisOnPremisesImmutableId) {
            try {
                $GuidBytes = [System.Convert]::FromBase64String($ThisOnPremisesImmutableId)
                $GuidObj = New-Object -TypeName guid -ArgumentList (,$GuidBytes)
                $ThisADU = Get-ADUser -Identity $GuidObj.Guid -Properties * -server $dc -ErrorAction Stop
                $hasADUser = $true ; 
                write-host "Found matching AD user by immutableId: $($ThisADU.DistinguishedName)"
            }CATCH {
                write-host "onPremisesImmutableId present but resolving AD user failed: $($_.Exception.Message)"
                try {            
                    #$ThisADU = Get-ADUser -Identity $GuidObj.Guid -Properties * -ErrorAction Stop
                    $thisUPN = $thisXoMbx.userprincipalname ; 
                    write-warning "onPremisesImmutableId present, UNMATCHED, checking for an ADUser with MGUser.UPN:$($thisUPN)" ; 
                    $ThisADU = get-aduser -Filter {userprincipalname -eq $thisUPN} -Properties * -server $dc -ErrorAction STOP ;
                    if($thisADU){
                        $hasADUser = $true ; 
                        $smsg = "Found matching AD user by UPN: $($ThisADU.DistinguishedName)" ; 
                        $smsg += "`nUpdating MgUser OnPremisesImmutableId to match ADUser OpImmutableId" ;
                        write-warning $smsg ;
                        <#
                        # Hard-match MG user to newly created AD user
                        $OpImmutableId = [System.Convert]::ToBase64String($ThisADU.ObjectGUID.ToByteArray())
                        $pltUdMgu = [ordered]@{
                            UserId = $ThisMgu.Id ;
                            OnPremisesImmutableId = $OpImmutableId ;
                            ErrorAction = 'Stop' ;
                            WhatIf = $whatIfSwitch ;
                        }
                        write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):Update-MgUser w`n$(($pltUdMgu|out-string).trim())" ;
                        if ($Force -or $PSCmdlet.ShouldProcess($ThisMgu.displayname, "Update-MgUser")) {
                            #Update-MgUser -UserId $ThisMgu.Id -OnPremisesImmutableId $OpImmutableId -ErrorAction Stop -WhatIf:$whatIfSwitch
                            Update-MgUser @pltUdMgu ;
                            $doUpdtMGUOnPremImmut = $true ;
                            write-verbose "refresh thisMGU" ;
                            $ThisMgu = Get-MgUserFull -UserId $ThisXoMbx.ExternalDirectoryObjectId -ErrorAction Stop
                            $hasImmutSync = $true ;
                        } else {
                            Write-Host "(-Whatif or `"No`" to the prompt)"
                        } ;                      
                        #>
                        $thisMGU = Update-MGUIMmut -thisMGU $thisMGU -thisADU $thisADU -Verbose:$($VerbosePreference -eq 'Continue') -whatif:$($whatifswitch) ;
                    } ; 
                <#
                } CATCH [System.Exception] {
                    $ErrTrapd=$Error[0] ;
                    if($ErrTrapd.ErrorDetails -match 'Insufficient\sprivileges\sto\scomplete\sthe\soperation'){
                        $smsg = "Update-MgUse: PERMISSION FAIL!" ;
                        $smsg += "`nIF USING ISE, RUN THE CMD OUTSIDE OF ISE IN A FULL PS SESSION!" ;
                        $smsg += "`n`nPS> Update-MgUser -UserId $($ThisMgu.Id) -OnPremisesImmutableId $($OpImmutableId) -ErrorAction Stop -WhatIf"
                        write-warning $smsg ;
                        write-warning $smsg ;
                    }else{
                        THROW $ErrTrapd
                    } ;
                #>
                }CATCH {                  
                    #write-warning "POPULATED onPremisesImmutableId and no MGUser.UPN matching ADuser"
                    $ErrTrapd=$Error[0] ;
                    write-host -foregroundcolor gray "TargetCatch:} CATCH [$($ErrTrapd.Exception.GetType().FullName)] {"  ;
                    $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                    write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" ;
                }
            } # CATCH
        }else{
            TRY {            
                #$ThisADU = Get-ADUser -Identity $GuidObj.Guid -Properties * -ErrorAction Stop
                $thisUPN = $thisXoMbx.userprincipalname ; 
                write-warning "NO onPremisesImmutableId present, checking for an ADUser with MGUser.UPN:$($thisUPN)" ; 
                $ThisADU = get-aduser -Filter {userprincipalname -eq $thisUPN} -Properties * -server $dc -ErrorAction STOP -Server $dc ;            
                if($thisADU){
                    $hasADUser = $true ; 
                    $smsg = "Found matching AD user by UPN: $($ThisADU.DistinguishedName)" ; 
                    $smsg += "`nUpdating MgUser OnPremisesImmutableId to match ADUser OpImmutableId" ;
                    write-warning $smsg ; 
                    # Hard-match MG user to newly created AD user
                    <#
                    $OpImmutableId = [System.Convert]::ToBase64String($ThisADU.ObjectGUID.ToByteArray())
                    $pltUdMgu = [ordered]@{
                        UserId = $ThisMgu.Id ;
                        OnPremisesImmutableId = $OpImmutableId ;
                        ErrorAction = 'Stop' ;
                        WhatIf = $whatIfSwitch ;
                    }
                    write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):Update-MgUser w`n$(($pltUdMgu|out-string).trim())" ;
                    if ($Force -or $PSCmdlet.ShouldProcess($ThisMgu.displayname, "Update-MgUser")) {
                        #Update-MgUser -UserId $ThisMgu.Id -OnPremisesImmutableId $OpImmutableId -ErrorAction Stop -WhatIf:$whatIfSwitch
                        Update-MgUser @pltUdMgu ;
                        $doUpdtMGUOnPremImmut = $true ;
                        write-verbose "refresh thisMGU" ;
                        $ThisMgu = Get-MgUserFull -UserId $ThisXoMbx.ExternalDirectoryObjectId -ErrorAction Stop
                        $hasImmutSync = $true ;
                    } else {
                        Write-Host "(-Whatif or `"No`" to the prompt)"
                    } ;                     
                    #>
                    $thisMGU = Update-MGUIMmut -thisMGU $thisMGU -thisADU $thisADU -Verbose:$($VerbosePreference -eq 'Continue') -whatif:$($whatifswitch) 
                } ; 
            <#
            } CATCH [System.Exception] {
                    $ErrTrapd=$Error[0] ;
                    if($ErrTrapd.ErrorDetails -match 'Insufficient\sprivileges\sto\scomplete\sthe\soperation'){
                        $smsg = "Update-MgUse: PERMISSION FAIL!" ;
                        $smsg += "`nIF USING ISE, RUN THE CMD OUTSIDE OF ISE IN A FULL PS SESSION!" ;
                        $smsg += "`n`nPS> Update-MgUser -UserId $($ThisMgu.Id) -OnPremisesImmutableId $($OpImmutableId) -ErrorAction Stop -WhatIf"
                        write-warning $smsg ;
                        write-warning $smsg ;
                    }else{
                        THROW $ErrTrapd
                    } ;
            #>
            }CATCH {
                write-warning "NO onPremisesImmutableId and no MGUser.UPN matching ADuser"
            }
        }

        # 3. If mailbox is DirSynced, check for RemoteMailbox by ExchangeGuid
        $ThisRmbx = $null
        if ($ThisXoMbx.IsDirSynced) {
            $hasXoMbxDirSync = $true ; 
            try {
                $ThisRmbx = Get-RemoteMailbox -Identity $ThisXoMbx.ExchangeGuid.Guid -DomainController $dc -ErrorAction SilentlyContinue
                if ($ThisRmbx) { 
                    write-host "Found RemoteMailbox: $($ThisRmbx.Identity)" ; 
                    $hasRmbx = $true ; 
                    $hasRmbxExGuidMatch = $true ; 
                }else{
                    $ThisRmbx = Get-RemoteMailbox -Identity $ThisXoMbx.UserPrincipalName -DomainController $dc -ErrorAction SilentlyContinue
                    if ($ThisRmbx) { 
                        write-host "Found RemoteMailbox (via ExGuid): $($ThisRmbx.Identity)" ; 
                        $hasRmbx = $true ;                     
                    } ; 
                } ; 
            }CATCH {
                write-host "Get-RemoteMailbox lookup error: $($_.Exception.Message)"
            }
        }ELSE{
            # intersync, it's possible for there to be an RMBX, and the XoMbx to be .isDirsynced:$false
            try {
                $ThisRmbx = Get-RemoteMailbox -Identity $ThisXoMbx.UserPrincipalName -DomainController $dc -ErrorAction SilentlyContinue
                if ($ThisRmbx) { 
                    write-host "Found RemoteMailbox (via UPN): $($ThisRmbx.Identity)" ; 
                    $hasRmbx = $true ;                     
                } ;             
            }CATCH {
                write-host "Get-RemoteMailbox lookup error: $($_.Exception.Message)"
            }
        }

        # PreCheck ExchangeGuid on RemoteMailbox matches Exchange Online mailbox's ExchangeGuid
        if ($ThisRmbx -and $ThisXoMbx.ExchangeGuid) {
            if ($ThisRmbx.ExchangeGuid -ne $ThisXoMbx.ExchangeGuid.Guid) {
                write-host "Setting RemoteMailbox ExchangeGuid to match XO mailbox ExchangeGuid"
                #$whatIfSwitch = $WhatIf.IsPresent        
                #if ($PSCmdlet.ShouldProcess($ThisRmbx.Identity, 'Set-RemoteMailbox ExchangeGuid')) {
                if ($Force -OR $PSCmdlet.ShouldProcess($ThisRmbx.Identity, 'Set-RemoteMailbox ExchangeGuid')) {
                    Write-Host "Write Action Here: $InputObject" ; 
                    Set-RemoteMailbox -Identity $ThisRmbx.Identity -ExchangeGuid $ThisXoMbx.ExchangeGuid.Guid -DomainController $dc -WhatIf:$whatIfSwitch -ErrorAction Stop
                    $doRmbxExGuidMatch = $false ; 
                    # refresh the rmbx for trailing report
                    $ThisRmbx = Get-RemoteMailbox -Identity $ThisRmbx.Identity -DomainController $dc -ErrorAction STOP
                    $hasRmbxExGuidMatch = $true ; 
                } else {
                    Write-Host "(-Whatif or `"No`" to the prompt)" ; 
                }  ;                      
            }else{
                $hasRmbxExGuidMatch = $true ; 
            }
        }

        # 4. If AD user exists but MG user is not hard-matched, set OnPremisesImmutableId to AD.ObjectGUID
        if ($ThisADU -and $ThisMgu) {
            # Compare base64 of AD GUID to MG OnPremisesImmutableId
            $OpImmutableId = [System.Convert]::ToBase64String($ThisADU.ObjectGUID.ToByteArray())
            if (-not $ThisMgu.OnPremisesImmutableId -or ($ThisMgu.OnPremisesImmutableId -ne $OpImmutableId)) {
                write-host "Updating MG user OnPremisesImmutableId to AD ObjectGUID base64 to hard-match"
                <#
                if ($force -OR  $PSCmdlet.ShouldProcess($ThisMgu.Id, 'Update OnPremisesImmutableId')) {
                $pltUdMgu = [ordered]@{
                    UserId = $ThisMgu.Id ;
                    OnPremisesImmutableId = $OpImmutableId ;
                    ErrorAction = 'Stop' ;
                    WhatIf = $whatIfSwitch ;
                }
                write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):Update-MgUser w`n$(($pltUdMgu|out-string).trim())" ;
                TRY {
                    if ($Force -or $PSCmdlet.ShouldProcess($ThisMgu.displayname, "Update-MgUser")) {
                        #Update-MgUser -UserId $ThisMgu.Id -OnPremisesImmutableId $OpImmutableId -ErrorAction Stop -WhatIf:$whatIfSwitch
                        Update-MgUser @pltUdMgu ;
                        $doUpdtMGUOnPremImmut = $true ;
                        write-verbose "refresh thisMGU" ;
                        $ThisMgu = Get-MgUserFull -UserId $ThisXoMbx.ExternalDirectoryObjectId -ErrorAction Stop
                        $hasImmutSync = $true ;
                    } else {
                        Write-Host "(-Whatif or `"No`" to the prompt)"
                    } ;
                } CATCH [System.Exception] {
                    $ErrTrapd=$Error[0] ;
                    if($ErrTrapd.ErrorDetails -match 'Insufficient\sprivileges\sto\scomplete\sthe\soperation'){
                        $smsg = "Update-MgUse: PERMISSION FAIL!" ;
                        $smsg += "`nIF USING ISE, RUN THE CMD OUTSIDE OF ISE IN A FULL PS SESSION!" ;
                        $smsg += "`n`nPS> Update-MgUser -UserId $($ThisMgu.Id) -OnPremisesImmutableId $($OpImmutableId) -ErrorAction Stop -WhatIf"
                        write-warning $smsg ;
                        write-warning $smsg ;
                    }else{
                        THROW $ErrTrapd
                    } ;
                }CATCH {
                    #write-warning "POPULATED onPremisesImmutableId and no MGUser.UPN matching ADuser"
                    $ErrTrapd=$Error[0] ;
                    write-host -foregroundcolor gray "TargetCatch:} CATCH [$($ErrTrapd.Exception.GetType().FullName)] {"  ;
                    $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                    write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" ;
                }
                #>
                $thisMGU = Update-MGUIMmut -thisMGU $thisMGU -thisADU $thisADU -Verbose:$($VerbosePreference -eq 'Continue') -whatif:$($whatifswitch) 
            }else{
                $hasImmutSync = $true ; 
            }
        } ; 

        # 5. If no AD user exists, create one under the specified OU
        if (-not $ThisADU) {
            # Build attributes from MG user
            $ouPath = 'OU=Azure Email Enabled,OU=System Accounts,OU=Other Accounts,OU=LYN,DC=global,DC=ad,DC=toro,DC=com'
            $displayName = $ThisMgu.DisplayName
            $givenName = $ThisMgu.GivenName
            $surname = $ThisMgu.Surname
            $userPrincipalName = $ThisMgu.UserPrincipalName

            # SamAccountName: first 20 word characters from displayname
            $SamAccountNameBase = (($displayName.ToCharArray() | Where-Object { $_ -match '\w' } | Select-Object -First 20) -join '')
            $SamAccountName = $SamAccountNameBase
            $suffix = 1
            while (Get-ADUser -Filter { SamAccountName -eq $SamAccountName } -server $dc -ErrorAction SilentlyContinue) {
                $SamAccountName = "$SamAccountNameBase$suffix"
                $suffix++
            }

            $pltNADU = [ordered]@{
                Name = $displayName ;
                GivenName = $givenName ;
                Surname = $surname ;
                SamAccountName = $SamAccountName ;
                UserPrincipalName = $userPrincipalName ;
                Path = $ouPath ;
                AccountPassword = $secureRandomPassword ;
                Enabled = $false ;       
                PasswordNeverExpires = $true ;
                ChangePasswordAtLogon = $False ;
                Server = $dc ;
                #WhatIf = $whatIfSwitch 
                ErrorAction = 'stop' ; 
            } ;        
            #write-host "Creating AD user $displayName in $ouPath with SamAccountName $SamAccountName"
            if ($PSCmdlet.ShouldProcess($displayName, 'New-ADUser,Update-MgUser')) {
                $secureRandomPassword = (ConvertTo-SecureString -String (('Pfx' + [guid]::NewGuid().ToString()) ) -AsPlainText -Force)
                write-host -foregroundcolor YELLOW "`n$((get-date).ToString('HH:mm:ss')):new-ADUser w:`n$(($pltNADU|out-string).trim())`n" ;

                #New-ADUser -Name $displayName -GivenName $givenName -Surname $surname -SamAccountName $SamAccountName -UserPrincipalName $userPrincipalName -Path $ouPath -AccountPassword $secureRandomPassword -Enabled $false -ErrorAction Stop
                New-ADUser @pltNADU 
                $doAddADUser = $true ; 
                # Re-fetch created AD user
                $ThisADU = Get-ADUser -Filter { SamAccountName -eq $SamAccountName } -server $dc -Properties * -ErrorAction Stop
                <#
                # Hard-match MG user to newly created AD user
                $OpImmutableId = [System.Convert]::ToBase64String($ThisADU.ObjectGUID.ToByteArray())
                Update-MgUser -UserId $ThisMgu.Id -OnPremisesImmutableId $OpImmutableId -ErrorAction Stop  -WhatIf:$whatIfSwitch ; 
                $doUpdtMGUOnPremImmut = $true ; 
                write-verbose "refresh thisMGU" ; 
                $ThisMgu = Get-MgUserFull -UserId $ThisXoMbx.ExternalDirectoryObjectId -ErrorAction Stop
                $hasImmutSync = $true ; 
                #>
                $thisMGU = Update-MGUIMmut -thisMGU $thisMGU -thisADU $thisADU -Verbose:$($VerbosePreference -eq 'Continue') -whatif:$($whatifswitch) 
            } else {
                Write-Host "(-Whatif or `"No`" to the prompt)"
            } ; 
        } # if-E thisADU

        # 6. If no RemoteMailbox exists and AD user exists and is matched, create RemoteMailbox and set ExchangeGuid
        if (-not $ThisRmbx -and $ThisADU -and $ThisMgu) {
            # Find an .onmicrosoft.com routing address from XO mailbox
            $RemRouteAddr = ($ThisXoMbx.EmailAddresses | Where-Object { $_ -match '\.onmicrosoft\.com$' } | ForEach-Object { $_ -replace '^smtp:' , '' } | Select-Object -Last 1)
            #if (-not $RemRouteAddr) {
            if(-not $RemRouteAddr -and $thisADU){
                $smsg = "No suitable .onmicrosoft.com remote routing address found on $($ThisXoMbx.Identity)"
                write-warning $smsg ;
                $smsg = "$($ThisXoMbx.userprincipalname).isDirsync, but has *no* onmicrosoft.com suitable address!";
                $smsg += "`nAttempting to calculate an address and push it into the ADUser.proxyaddresses list (to populate on the xmbx, for a future pass, after ADC sync)" ;
                WRITE-WARNING $SMSG ;
                if($ThisXoMbx.EmailAddressPolicyEnabled -eq $false){
                    $dirname = "$($ThisXoMbx.primarysmtpaddress.split('@')[0])$($emlCoTag)" ;
                    $newdom = 'toroco.onmicrosoft.com' ;
                    $newpEml = @($dirname,$newdom) -join '@' ;
                    $newpEml = "smtp:$($newpEml)" ;
                    $xproxy = $thisADU.proxyaddresses  | ?{$_ -match '^smtp\:'} ;
                    if($xproxy -contains $newpEml){
                        $smsg = "ADU.$($ThisXoMbx.userprincipalname) already has the necessary RemoteRouting Address added, skipping dupe addition ";
                        write-warning $smsg ;
                    }else{
                        #$ThisXoMbx | set-xomailbox -EmailAddresses @{add="smtp:$($newpEml)"} -whatif:$($whatif) -ea STOP ;
                        #Set-ADUser -Identity $thisADU.DistinguishedName -Add @{proxyAddresses=$newpEml}  -whatif:$($whatif) -ea STOP -server $dc -VERBOSE  ;
                        if ($Force -or $PSCmdlet.ShouldProcess($thisadu.userprincipalname, "set-ADUser")) {
                            Set-ADUser -Identity $thisADU.objectguid.guid -Add @{proxyAddresses=$newpEml}  -whatif:$($whatifswitch) -ea STOP -server $dc -VERBOSE  ;
                            # refresh the updated obj
                            $thisADU = get-aduser -Identity $thisADU.objectguid.guid -ea STOP -server $dc
                        } else {
                            Write-Host "(-Whatif or `"No`" to the prompt)"
                        } ; 
                    } ;
                    $ADUFix = $true ; 
                    $smsg = "WAIT FULL ADC CYCLE AND RECHECK IF THE XMBX HAS A SUITABLE ONMICROSOFT.COM ADDRESS (and MGU has matched OPimmuntable), THEN RERUN AN UPDATE TO CREATE RMBX" ;
                    WRITE-WARNING $SMSG ;
                } ;
            } ; 

            if($RemRouteAddr -and $thisADU){
                write-host "Enabling RemoteMailbox for $($ThisADU.UserPrincipalName) with RemoteRoutingAddress $RemRouteAddr"
                if ($PSCmdlet.ShouldProcess($ThisADU.UserPrincipalName, 'Enable-RemoteMailbox')) {
                    $ThisRmbx = Enable-RemoteMailbox -Identity $ThisADU.UserPrincipalName -RemoteRoutingAddress $RemRouteAddr -DomainController $dc -ErrorAction Stop ; 
                    $doAddRmbx = $true ; 
                    <#
                    VERBOSE: Enabling RemoteMailbox for Marketing@intimidatorutv.com with RemoteRoutingAddress Marketing1@toroco.onmicrosoft.com
                    This task does not support recipients of this type. The specified recipient global.ad.toro.com/LYN/Other Accounts/System Accounts/Azure Email Enabled/Marketing - Intimidator is of type RemoteUserMailbox. Please make sure
                    that this recipient matches the required recipient type for this task.
                    + CategoryInfo          : InvalidArgument: (global.ad.toro....g - Intimidator:ADObjectId) [Enable-RemoteMailbox], RecipientTaskException
                    + FullyQualifiedErrorId : 2882FF3D,Microsoft.Exchange.Management
                    #>
                    # above throws error, but still created rmbx: redisco it as UPN
                    try {
                        #$ThisRmbx = Get-RemoteMailbox -Identity $ThisXoMbx.ExchangeGuid.Guid -ErrorAction SilentlyContinue
                        # newly created won't have the guid match
                        $ThisRmbx = Get-RemoteMailbox -Identity $ThisXoMbx.userprincipalname -DomainController $dc -ErrorAction SilentlyContinue
                        if ($ThisRmbx) { write-host "Found RemoteMailbox: $($ThisRmbx.Identity)" }
                    }CATCH {
                        write-host "Get-RemoteMailbox lookup error: $($_.Exception.Message)"
                    }
                }
            }else{
                $smsg = "thisADU but No matched RemoteRoutingAddress, deferrring Rmbx creation until RRA is available (future pass)" ; 
                write-warning $smsg;
            }
            

            # Ensure ExchangeGuid on RemoteMailbox matches Exchange Online mailbox's ExchangeGuid
            if ($ThisRmbx -and $ThisXoMbx.ExchangeGuid) {
                if ($ThisRmbx.ExchangeGuid -ne $ThisXoMbx.ExchangeGuid.Guid) {
                    write-host "Setting RemoteMailbox ExchangeGuid to match XO mailbox ExchangeGuid"
                    $whatIfSwitch = $WhatIf.IsPresent        
                    if ($PSCmdlet.ShouldProcess($ThisRmbx.Identity, 'Set-RemoteMailbox ExchangeGuid')) {
                        Set-RemoteMailbox -Identity $ThisRmbx.Identity -ExchangeGuid $ThisXoMbx.ExchangeGuid.Guid -DomainController $dc -WhatIf:$whatIfSwitch -ErrorAction Stop
                        $doRmbxExGuidMatch = $false ; 
                        # refresh the rmbx for trailing report
                        $ThisRmbx = Get-RemoteMailbox -Identity $ThisRmbx.Identity -DomainController $dc -ErrorAction STOP
                        $hasRmbxExGuidMatch = $true ; 
                    } else {
                        Write-Host "(-Whatif or `"No`" to the prompt)"
                    } ; 
                }else{
                    $hasRmbxExGuidMatch = $true ; 
                }
            }
        }ELSE{
            $smsg = "MISSING COMPONENT: " ; # $ThisRmbx -and $ThisXoMbx.ExchangeGuid
            $smsg += "thisRmbx w`n$(($ThisRmbx | ft -a |out-string).trim())" ; 
            $smsg += "ThisXoMbx.ExchangeGuid:`n$($ThisXoMbx.ExchangeGuid)" ; 
            write-warning $smsg ; 
        }
        $hsReport = @"

###ThisXoMailbox: 
$(($ThisXoMbx| ft -a Name,Alias,ServerName,isDirSynced,PrimarySMTPAddress|out-string).trim())
$(($ThisXoMbx| ft -a ExchangeGuid |out-string).trim())

###ThisMgu: 
$(($ThisMgu| ft -a DisplayName,Id,Mail,UserPrincipalName  |out-string).trim())
$(($ThisMgu| ft -a OnPremisesImmutableId,OnPremisesSyncEnabled,OnPremisesSyncEnabled|out-string).trim())
$(($ThisMgu| ft -a OnPremisesDistinguishedName,OnPremisesProvisioningErrors|out-string).trim())

###ThisADU: 
$(($ThisADU| ft -a name,DistinguishedName,Enabled |out-string).trim())
$(($ThisADU| ft -a 'GivenName','Surname','Name' |out-string).trim())
$(($ThisADU| ft -a 'SamAccountName','UserPrincipalName','ObjectClass','ObjectGUID' |out-string).trim())


###ThisOnPremisesImmutableId: $(($ThisOnPremisesImmutableId|out-string).trim())
###(equiv converted ADUser:OpImmutableId:$(($OpImmutableId|out-string).trim())

###ThisRmbx: 
$(($ThisRmbx| ft -a Name,RecipientTypeDetails,RemoteRecipientType |out-string).trim())
$(($ThisRmbx| ft -a ExchangeGuid |out-string).trim())

"@;
        WRITE-HOST -FOREGROUNDCOLOR GREEN $hsReport ; 


        $actions = @('doAddADUser','doUpdtMGUOnPremImmut','doAddRmbx','doRmbxExGuidMatch')    
        $actions | foreach-object{
            $thisActName = $_ ; 
            if((gv -Name $thisActName -ea 0).value -eq $true){
                write-warning "ACTION:`$$($thisActName):$((gv -Name $thisActName).value)!"
            } else{
                write-host -foregroundcolor green  "ACTION:$($thisActName):$((gv -Name $thisActName).value)"
            } ; 
        } ; 
        write-host "`n" ; 
        $tests = @('hasXoMbx','hasXoMbxDirSync','hasMgUser','hasADUser','hasImmutSync','hasRmbx','hasRmbxExGuidMatch'); 
        $tests | foreach-object{
        $thisTestName = $_ ; 
            if((gv -Name $thisTestName -ea 0).value -eq $false){
                write-warning "TEST:`$$($thistestName):$((gv -Name $thistestname).value)!"
            } else{
                write-host -foregroundcolor green  "TEST:$($thistestName):$((gv -Name $thistestname).value)"
            } ; 
        } ; 
        write-host "`nUpdate-EXOLinkedHybridObjectsTDO completed for $($ThisXoMbx.Identity)"
    } ;  # PROC-E
    END{
        if($stopResults){
            $smsg = "Stop-transcript:$($stopResults)" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
        } ;
    }
}
