# connect-O365Services.ps1

    #region CONNECT_O365SERVICES ; #*======v connect-O365Services v======
    #if(-not (get-childitem function:connect-O365Services -ea 0)){
        function connect-O365Services {
            <#
            .SYNOPSIS
            connect-O365Services - logic wrapper for my histortical scriptblock that resolves creds, svc avail and relevent status, to connect to range of Services (in o365)
            .NOTES
    
            REVISIONS
            * 4:59 PM 1/20/2026 fixed returning populated CredentialO365 in return object
            * 4:12 PM 1/6/2026 revised for verb-MG compait
            *8:17 PM 6/1/2025 debugs functional for useexo & usesc now; 
    
            .PARAMETER EnvSummary
            Pre-resolved local environrment summary (product of output of verb-io\resolve-EnvironmentTDO())[-EnvSummary `$rvEnv]
            .PARAMETER NetSummary
            Pre-resolved local network summary (product of output of verb-network\resolve-NetworkLocalTDO())[-NetSummary `$netsettings]
            .PARAMETER useEXO
            Connect to O365 ExchangeOnlineManagement)[-useEXO]
            .PARAMETER useSC
            Connect to O365 Security & Compliance/Purview)[-useSC]
            .PARAMETER UseMSOL
            Connect to O365 MSOnline powershell module)[-UseMSOL]
            .PARAMETER UseAAD
            Connect to O365 AzureAD  powershell module)[-UseAAD]
            .PARAMETER UseMG
            Connect to O365 Microsoft.Graph powershell module(s))[-UseMG]
            .PARAMETER MGPermissionsScope
            Optional Array of MG delegated Permission Names(avoids manual discovery against configured cmdlets)[-MGPermissionsScope @('Domain.Read.All','Domain.ReadWrite.All','Directory.Read.All') ]
            .PARAMETER MGCmdlets
            Microsoft.Graph powershell module cmdlets used for this connection scope (avoids lengthy manual AST Parse of source script; used with verb-MG\get-MGCodeCmdletPermissionsTDO() to resolve connect-mgGraph delegated permissions connection -scope)[-MGCmdlets]
            .PARAMETER TenOrg
            Tenant Tag (3-letter abbrebiation)[-TenOrg 'XYZ']
            .PARAMETER Credential
            Use specific Credentials (defaults to Tenant-defined SvcAccount)[-Credentials [credential object]]
            .PARAMETER AdminAccount
            Use specific AdminAccount for service connections (defaults to Tenant-defined SvcAccount)[-AdminAccount LOGON@DOMAIN.COM]
            .PARAMETER UserRole
            Credential User Role spec (SID|CSID|UID|B2BI|CSVC|ESVC|LSVC|ESvcCBA|CSvcCBA|SIDCBA)[-UserRole @('SIDCBA','SID','CSVC')]
            .PARAMETER useEXOv2
            Use EXOv2 (ExchangeOnlineManagement) over basic auth legacy connection [-useEXOv2]
            .PARAMETER Silent
            Silent output (suppress status echos)[-silent]
            .PARAMETER MGPermissionsScope
            Optional Array of MG Permission Names(avoids manual discovery against configured cmdlets)[-MGPermissionsScope @('Domain.Read.All','Domain.ReadWrite.All','Directory.Read.All') ]
    
            #>
            ##Requires -Modules AzureAD, verb-AAD
            [CmdletBinding()]
            ## PSV3+ whatif support:[CmdletBinding(SupportsShouldProcess)]
            ###[Alias('Alias','Alias2')]
            PARAM(
                # environment parameters:
                [Parameter(Mandatory=$true,HelpMessage="Pre-resolved local environrment summary (product of output of verb-io\resolve-EnvironmentTDO())[-EnvSummary `$rvEnv]")]
                    $EnvSummary, # $rvEnv
                [Parameter(Mandatory=$true,HelpMessage="Pre-resolved local network summary (product of output of verb-network\resolve-NetworkLocalTDO())[-NetSummary `$netsettings]")]
                    $NetSummary, # $netsettings
                # service choices
                [Parameter(HelpMessage="Connect to O365 ExchangeOnlineManagement)[-useEXO]")]
                    [switch]$useEXO,
                [Parameter(HelpMessage="Connect to O365 Security & Compliance/Purview)[-useSC]")]
                    [switch]$useSC,
                [Parameter(HelpMessage="Connect to O365 MSOnline powershell module)[-UseMSOL]")]
                    [switch]$UseMSOL,
                [Parameter(HelpMessage="Connect to O365 AzureAD  powershell module)[-UseAAD]")]
                    [switch]$UseAAD,
                [Parameter(HelpMessage="Connect to O365 Microsoft.Graph powershell module(s))[-UseMG]")]
                    [switch]$UseMG,
                # Service Connection parameters
                [Parameter(Mandatory=$FALSE,HelpMessage="TenantTag value, indicating Tenants to connect to[-TenOrg 'TOL']")]
                    [ValidateNotNullOrEmpty()]
                    #[ValidatePattern("^\w{3}$")]
                    [string]$TenOrg = $global:o365_TenOrgDefault,
                [Parameter(Mandatory = $false, HelpMessage = "Use specific Credentials (defaults to Tenant-defined SvcAccount)[-Credentials [credential object]]")]
                    [System.Management.Automation.PSCredential]$Credential,
                [Parameter(Mandatory = $false, HelpMessage = "Use specific AdminAccount for service connections (defaults to Tenant-defined SvcAccount)[-AdminAccount LOGON@DOMAIN.COM]")]
                        [string]$AdminAccount,
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
                    # svcAcct use: @('ESvcCBA','CSvcCBA','SIDCBA')
                [Parameter(HelpMessage="Use EXOv2 (ExchangeOnlineManagement) over basic auth legacy connection [-useEXOv2]")]
                    [switch] $useEXOv2=$true,
                [Parameter(HelpMessage="Silent output (suppress status echos)[-silent]")]
                    [switch] $silent,            
                [Parameter(HelpMessage="Optional Array of MG delegated Permission Names(avoids manual discovery against configured cmdlets)[-MGPermissionsScope @('Domain.Read.All','Domain.ReadWrite.All','Directory.Read.All') ]")]
                    [string[]]$MGPermissionsScope,
                [Parameter(HelpMessage="Microsoft.Graph powershell module cmdlets used for this connection scope (avoids lengthy manual AST Parse of source script; used with verb-MG\get-MGCodeCmdletPermissionsTDO() to resolve connect-mgGraph delegated permissions connection -scope)[-MGCmdlets]")]
                    [string[]]$MGCmdlets
            );
            BEGIN {
                # for scripts wo support, can use regions to fake BEGIN;PROCESS;END: (tho' can use the real deal in scripts as well as adv funcs, as long as all code is inside the blocks)
                # ps1 faked:#region BEGIN ; #*------v BEGIN v------
                # 8:59 PM 4/23/2025 with issues in CMW - funcs unrecog'd unless loaded before any code use - had to move the entire FUNCTIONS block to the top of BEGIN{}

                #region FUNCTIONS_INTERNAL ; #*======v FUNCTIONS_INTERNAL v======
                # Pull the CUser mod dir out of psmodpaths:
                #$CUModPath = $env:psmodulepath.split(';')|?{$_ -like '*\Users\*'} ;

                #region get_CodeProfileAST ; #*------v get-CodeProfileAST v------
                if(-not (get-childitem function:get-CodeProfileAST -ea 0)){
                      function get-CodeProfileAST {
                      <#
                      .SYNOPSIS
                      get-CodeProfileAST - Parse and return script/module/function compoonents, Module using Language.FunctionDefinitionAst parser
                      .NOTES
                  
                      REVISIONS
                      * 10:57 AM 5/19/2025 add: CBH for more extensive code profiling demo (for targeting action-verb cmds in code, from specific modules); fixed some missing CBH info.
                  
                      .PARAMETER  File
                      Path to script/module file
                      .PARAMETER scriptblock
                      Scriptblock of code[-scriptblock `$sbcode]
                      .PARAMETER Functions
                      Flag to return Functions-only [-Functions]
                      .PARAMETER Parameter
                      Flag to return Parameter-only [-Functions]
                      .PARAMETER Variables
                      Flag to return Variables-only [-Variables]
                      .PARAMETER Aliases
                      Flag to return Aliases-only [-Aliases]
                      .PARAMETER GenericCommands
                      Flag to return GenericCommands-only [-GenericCommands]
                      .PARAMETER All
                      Flag to return All [-All]
                      .PARAMETER ShowDebug
                      Parameter to display Debugging messages [-ShowDebug switch]
                      .PARAMETER Whatif
                      Parameter to run a Test no-change pass [-Whatif switch]
                  
                      #>
                      [CmdletBinding()]
                      [Alias('get-ScriptProfileAST')]
                      PARAM(
                          [Parameter(Position = 0, Mandatory = $false, ValueFromPipeline = $true, HelpMessage = "Path to script[-File path-to\script.ps1]")]
                              [ValidateScript( {Test-Path $_})][ValidatePattern( "\.(ps1|psm1|txt)$")]
                              [Alias('PSPath','File')]
                              [system.io.fileinfo]$Path,
                          [Parameter(Position = 1, Mandatory = $false, ValueFromPipeline = $true, HelpMessage = "Scriptblock of code[-scriptblock `$sbcode]")]
                              [Alias('code')]
                              $scriptblock,
                          [Parameter(HelpMessage = "Flag to return Functions-only [-Functions]")]
                              [switch] $Functions,
                          [Parameter(HelpMessage = "Flag to return Parameters-only [-Functions]")]
                              [switch] $Parameters,
                          [Parameter(HelpMessage = "Flag to return Variables-only [-Variables]")]
                              [switch] $Variables,
                          [Parameter(HelpMessage = "Flag to return Aliases-only [-Aliases]")]
                              [switch] $Aliases,
                          [Parameter(HelpMessage = "Flag to return GenericCommands-only [-GenericCommands]")]
                              [switch] $GenericCommands,
                          [Parameter(HelpMessage = "Flag to return All [-All]")]
                              [switch] $All,
                          [Parameter(HelpMessage = "Debugging Flag [-showDebug]")]
                              [switch] $showDebug,
                          [Parameter(HelpMessage = "Whatif Flag  [-whatIf]")]
                              [switch] $whatIf
                      ) ;
                      BEGIN {
                          TRY{
                              $Verbose = ($VerbosePreference -eq "Continue") ;
                              if(-NOT ($path -OR $scriptblock)){
                                  throw "neither -Path or -Scriptblock specified: Please specify one or the other when running" ; 
                                  break ; 
                              } elseif($path -AND $scriptblock){
                                  throw "BOTH -Path AND -Scriptblock specified: Please specify EITHER one or the other when running" ; 
                                  break ; 
                              } ;  
                              if ($Path -AND $Path.GetType().FullName -ne 'System.IO.FileInfo') {
                                  write-verbose "(convert path to gci)" ; 
                                  $Path = get-childitem -path $Path ; 
                              } ;
                              if ($scriptblock -AND $scriptblock.GetType().FullName -ne 'System.Management.Automation.ScriptBlock') {
                                  write-verbose "(recast -scriptblock to [scriptblock])" ; 
                                  $scriptblock= [scriptblock]::Create($scriptblock) ; 
                              } ;
                          } CATCH {
                              $ErrTrapd=$Error[0] ;
                              $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                              if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                              else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                          } ; 
                      } ;
                      PROCESS {
                          $sw = [Diagnostics.Stopwatch]::StartNew();
                          TRY{
                              write-verbose "$((get-date).ToString('HH:mm:ss')):(running AST parse...)" ; 
                              New-Variable astTokens -Force -whatif:$false -confirm:$false ; New-Variable astErr -Force -whatif:$false -confirm:$false ; 
                              if($Path){            
                                  $AST = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$astTokens, [ref]$astErr) ; 
                              }elseif($scriptblock){
                                  $AST = [System.Management.Automation.Language.Parser]::ParseInput($scriptblock, [ref]$astTokens, [ref]$astErr) ; 
                              } ;     
                              if($host.version.major -ge 3){$objReturn=[ordered]@{Dummy = $null ;} }
                              else {$objReturn = @{Dummy = $null ;} } ;
                              if ($Functions -OR $All) {
                                  write-verbose "$((get-date).ToString('HH:mm:ss')):(parsing Functions from AST...)" ; 
                                  $ASTFunctions = $AST.FindAll( { $args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true) ;
                                  $objReturn.add('Functions', $ASTFunctions) ;
                              } ;
                              if ($Parameters -OR $All) {
                                  write-verbose "$((get-date).ToString('HH:mm:ss')):(parsing Parameters from AST...)" ; 
                                  $ASTParameters = $ast.ParamBlock.Parameters.Name.variablepath.userpath ;
                                  $objReturn.add('Parameters', $ASTParameters) ;
                              } ;
                              if ($Variables -OR $All) {
                                  write-verbose "$((get-date).ToString('HH:mm:ss')):(parsing Variables from AST...)" ; 
                                  $AstVariableAssignments = $ast.FindAll( { $args[0] -is [System.Management.Automation.Language.VariableExpressionAst] }, $true) ;
                                  $objReturn.add('Variables', $AstVariableAssignments) ;
                              } ;
                              if ($($Aliases -OR $GenericCommands) -OR $All) {
                                  write-verbose "$((get-date).ToString('HH:mm:ss')):(parsing ASTGenericCommands from AST...)" ; 
                                  $ASTGenericCommands = $ast.FindAll( { $args[0] -is [System.Management.Automation.Language.CommandAst] }, $true) ;
                                  if ($Aliases -OR $All) {
                                      write-verbose "$((get-date).ToString('HH:mm:ss')):(post-filtering (set|new)-Alias from AST...)" ; 
                                      $ASTAliasAssigns = ($ASTGenericCommands | ? { $_.extent.text -match '(set|new)-alias' }) ;
                                      $objReturn.add('Aliases', $ASTAliasAssigns) ;
                                  } ;
                                  if ($GenericCommands -OR $All) {
                                      $objReturn.add('GenericCommands', $ASTGenericCommands) ;
                                  } ;
                              } ;
                              #$objReturn | Write-Output ;
                              New-Object PSObject -Property $objReturn | Write-Output ;
                          } CATCH {
                              $ErrTrapd=$Error[0] ;
                              $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                              if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                              else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                          } ; 
                      } ;
                      END {
                          $sw.Stop() ;
                          $smsg = ("Elapsed Time: {0:dd}d {0:hh}h {0:mm}m {0:ss}s {0:fff}ms" -f $sw.Elapsed) ; 
                          if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                          else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                      } ;
                  } ; 
                } ; 
                #endregion get_CodeProfileAST ; #*------^ END get-CodeProfileAST ^------

                #region get_MGCodeCmdletPermissionsTDO ; #*------v get-MGCodeCmdletPermissionsTDO v------
                if(-not (get-childitem function:get-MGCodeCmdletPermissionsTDO -ea 0)){
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                            function get-MGCodeCmdletPermissionsTDO {
                    <#
                    .SYNOPSIS
                    get-MGCodeCmdletPermissionsTDO - wrapper for verb-dev\get-codeprofileAST() that parses [verb]-MG[noun] cmdlets from a specified -file or -scriptblock, and reseolves the necessary connect-mgGraph -scope permissions, using the Find-MgGraphCommand  command.
                    .NOTES
                
                    REVISIONS
                    * 1:49 PM 5/14/2025 add: -cmdlets, bypasses AST parsing cuts right to find-mgGraphCommand expansion; additional verbose status echos (as it's returning very limited set of perms)
                
                    .PARAMETER  File
                    Path to script/module file to be parsed for matching cmdlets[-Path path-to\script.ps1]
                    .PARAMETER scriptblock
                    Scriptblock of code to be parsed for matching cmdlets[-scriptblock `$sbcode]
                    .PARAMETER CommandFilterRegex
                    Regular expression filter to match commands within GenericCommand lines parsed from subject code (defaults \w+-mg\w+)[-CommandFilterRegex '\w+-mgDomain\w+']
                    .PARAMETER ModuleFilterRegex 
                    Regular expression filter to match commands solely in matching Module (defaults 'Microsoft\.Graph')[-CommandFilterRegex 'Microsoft\.Graph\.Identity\.DirectoryManagement\s\s\s']
                    .PARAMETER Cmdlets
                    MGGraph cmdlet names to be Find-MgGraphCommand'd into delegated access -scope permissions (bypasses ASTParser discovery)
                
                    #>  
                    [CmdletBinding()]
                    ## PSV3+ whatif support:[CmdletBinding(SupportsShouldProcess)]
                    ###[Alias('Alias','Alias2')]
                    PARAM(
                        [Parameter(Position = 0,ValueFromPipeline = $true, HelpMessage = "Path to script/module file to be parsed for matching cmdlets[-Path path-to\script.ps1]")]
                            [ValidateScript( {Test-Path $_})][ValidatePattern( "\.(ps1|psm1|txt)$")]
                            [Alias('PSPath','File')]
                            [system.io.fileinfo]$Path,
                        [Parameter(Position = 1,HelpMessage = "Scriptblock of code to be parsed for matching cmdlets[-scriptblock `$sbcode]")]
                            [Alias('code')]
                            $scriptblock,
                        [Parameter(HelpMessage = "Regular expression filter to match commands within GenericCommand lines parsed from subject code (defaults \w+-mg\w+)[-CommandFilterRegex '\w+-mgDomain\w+']")]
                            [regex]$CommandFilterRegex = '\w+\-mg\w+',
                        [Parameter(HelpMessage = "Regular expression filter to match commands solely in matching Module (defaults 'Microsoft\.Graph')[-CommandFilterRegex 'Microsoft\.Graph\.Identity\.DirectoryManagement\s\s\s']")]
                            [regex]$ModuleFilterRegex = '^Microsoft\.Graph',
                        [Parameter(HelpMessage = "MGGraph cmdlet names to be Find-MgGraphCommand'd into delegated access -scope permissions (bypasses ASTParser discovery)[-Cmdlets @('get-MgDomain','get-MGContext')]")]
                            [string[]]$Cmdlets
                    );  
                    BEGIN {
                        $Verbose = ($VerbosePreference -eq "Continue") ;
                        # MG Cmdlets that don't have perms (don't bother FindMGCommanding them, wastes ~3mins for no return)
                        $MGNonPermCmdlets = 'Find-MgGraphCommand','Connect-MgGraph','Get-MgContext','Confirm-MgDomain','Get-MgDomainServiceConfigurationRecord' ; 
                        [regex]$rgxMGNonPermCmdlets = ('(' + (($MGNonPermCmdlets |%{[regex]::escape($_)}) -join '|') + ')') ;
                        if($Cmdlets){
                            $smsg = "-Cmdlets (skipping -path/-scriptblock AST parsing)" ; 
                            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                        }else{
                            TRY{
                                if(-NOT ($path -OR $scriptblock)){
                                    throw "neither -Path or -Scriptblock specified: Please specify one or the other when running" ; 
                                    break ; 
                                } elseif($path -AND $scriptblock){
                                    throw "BOTH -Path AND -Scriptblock specified: Please specify EITHER one or the other when running" ; 
                                    break ; 
                                } ;  
                                if ($Path -AND $Path.GetType().FullName -ne 'System.IO.FileInfo'){
                                    $smsg = "(convert path to gci)" ; 
                                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                                    $Path = get-childitem -path $Path ; 
                                } ;
                                if ($scriptblock -AND $scriptblock.GetType().FullName -ne 'System.Management.Automation.ScriptBlock'){
                                    $smsg = "(recast -scriptblock to [scriptblock])" ; 
                                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                                    $scriptblock= [scriptblock]::Create($scriptblock) ; 
                                } ;
                            } CATCH {
                                $ErrTrapd=$Error[0] ;
                                $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                                else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            } ; 
                        } ; 
                    } ;
                    PROCESS {
                        $sw = [Diagnostics.Stopwatch]::StartNew();
                        if($Cmdlets){
                            $smsg = "-cmdlets specified:`n$(($Cmdlets|out-string).trim())" ;                     
                        }else{
                            if($host.version.major -ge 3){$pltgCPA=[ordered]@{Dummy = $null ;} }
                            else {$pltgCPA = @{Dummy = $null ;} } ;
                            if($pltgCPA.keys -contains 'dummy'){$pltgCPA.remove('Dummy') };
                            $pltgCPA.add('erroraction','STOP' ) ;
                            $pltgCPA.add('GenericCommands',$true )  ;
                            if($Path){ $pltgCPA.add('Path',$Path.fullname)}
                            if($ScriptBlock){ $pltgCPA.add('ScriptBlock',$ScriptBlock)}
                            $smsg = "get-CodeProfileAST  w`n$(($pltgCPA|out-string).trim())" ; 
                            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                            $GCmds = (get-CodeProfileAST @pltgCPA).GenericCommands ; 
                            # shouldn't need .tostring() on a regex type, but w returns full list, wo returns just 1 item.
                            $GCmds.extent.text | ?{$_ -match $CommandFilterRegex.tostring()} | foreach-object {$cmdlets += $matches[0]} ; 
                            $smsg = "Normalize & unique names"; 
                            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                            if($ModuleFilterRegex){
                                $cmdlets = $cmdlets | select -unique | foreach-object { 
                                    get-command -name $_| ?{$_.source -match $ModuleFilterRegex} 
                                } | select -expand name | select -unique ;         
                            }else {
                                $cmdlets = $cmdlets | foreach-object { 
                                    get-command -name $_| select -expand name 
                                } | select -unique ;
                            }
                            $smsg = "Parsed following matching cmdlets:`n$(($cmdlets|out-string).trim())" ;   
                        } ;               
                        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                        if($Cmdlets | ?{$_ -match $rgxMGNonPermCmdlets}){
                            $smsg = "(Excluding non-Permission MGCmdlets from Permission discovery:" ; 
                            $smsg += "`n$(($Cmdlets | ?{$_ -match $rgxMGNonPermCmdlets}|out-string).trim())`n)" ; 
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                            $Cmdlets = $Cmdlets | ?{$_ -notmatch $rgxMGNonPermCmdlets} ; 
                        } ; 
                        write-host -foregroundcolor yellow "Resolving $($cmdlets.count) cmdlets against Find-MgGraphCommand..." ; 
                        $PermsRqd = @() ;         
                        write-host -foregroundcolor yellow "[" -nonewline ; 
                        $cmdlets |foreach-object{
                            $thisCmdlet = $_ ; 
                            $smsg = "$($thisCmdlet)" ; 
                            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                            write-host -NoNewline '.' ; 
                            #$PermsRqd += Find-MgGraphCommand -command $thisCmdlet -ea 0| Select -First 1 -ExpandProperty Permissions | Select -Unique name ; 
                            $thisPerm = $null ; 
                            #$thisPerm = Find-MgGraphCommand -command $thisCmdlet -ea 0| Select -First 1 -ExpandProperty Permissions | Select -Unique name ; 
                            $thisPerm = Find-MgGraphCommand -command $thisCmdlet -ea 0 |?{$_.permissions} | select -expand permissions | Select -Unique name ;   ; 
                            if($thisPerm){
                                $PermsRqd += $thisPerm ; 
                                $smsg = "(Find-MgGraphCommand -command $($thisCmdlet) returned Permissions:`n$(($thisPerm -join ','|out-string).trim()))" ; 
                            }else {
                                $smsg = "($($Cmdlet):no Permissions returned" ; 
                            } ; 
                            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                        } ; 
                        write-host -foregroundcolor yellow "]" ; 
                        $PermsRqd = $PermsRqd.name | select -unique ;
                    } ; # PROC-E  
                    END {
                        $sw.Stop() ;
                        $smsg = ("Elapsed Time: {0:dd}d {0:hh}h {0:mm}m {0:ss}s {0:fff}ms" -f $sw.Elapsed) ; 
                        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                        if($PermsRqd){
                            $PermsRqd | write-output ; 
                            $smsg = "(Resolved Perm Names:" ; 
                            #$smsg += "`n$((|out-string).trim())" ; 
                            $smsg += "`n'$(($PermsRqd) -join "','")'" ; 
                            $smsg += "`nCan be cached into a `$MGPermissionsScope etc, to skip this lengthy -scope discovery process)" ; 
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                        } else { 
                            $false | write-output 
                        } ; 
                    } ; # END-E
                } ; 
                } ; 
                #endregion get_MGCodeCmdletPermissionsTDO ; #*------^ END get-MGCodeCmdletPermissionsTDO ^------
                #endregion FUNCTIONS_INTERNAL ; #*======^ END FUNCTIONS_INTERNAL ^======

                #region CONSTANTS_AND_ENVIRO ; #*======v CONSTANTS_AND_ENVIRO v======
        
                #region NETWORK_INFO ; #*======v NETWORK_INFO v======
                if($env:Userdomain){
                    switch($env:Userdomain){
                        'CMW'{
                            #$logon_SID = $CMW_logon_SID
                        }
                        'TORO'{
                            #$o365_SIDUpn = $o365_Toroco_SIDUpn ;
                            #$logon_SID = $TOR_logon_SID ;
                        }
                        $env:COMPUTERNAME{
                            $smsg = "%USERDOMAIN% -EQ %COMPUTERNAME%: $($env:computername) => non-domain-connected, likely edge role Ex server!" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            if($NetSummary.Workgroup){
                                $smsg = "WorkgroupName:$($NetSummary.Workgroup)" ;
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                                else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                            } ;
                        } ;
                        default{
                            $smsg = "$($env:userdomain):UNRECOGIZED/UNCONFIGURED USER DOMAIN STRING!" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            THROW $SMSG
                            BREAK ;
                        }
                    } ;
                } ;  # $env:Userdomain-E
                #endregion NETWORK_INFO ; #*======^ END NETWORK_INFO ^======

                #region COMMON_CONSTANTS ; #*------v COMMON_CONSTANTS v------

                if(-not $DoRetries){$DoRetries = 4 } ;    # # times to repeat retry attempts
                if(-not $RetrySleep){$RetrySleep = 10 } ; # wait time between retries
                if(-not $RetrySleep){$DawdleWait = 30 } ; # wait time (secs) between dawdle checks
                if(-not $DirSyncInterval){$DirSyncInterval = 30 } ; # AADConnect dirsync interval
                if(-not $ThrottleMs){$ThrottleMs = 50 ;}
                if(-not $rgxDriveBanChars){$rgxDriveBanChars = '[;~/\\\.:]' ; } ; # ;~/\.:,
                if(-not $rgxCertThumbprint){$rgxCertThumbprint = '[0-9a-fA-F]{40}' } ; # if it's a 40char hex string -> cert thumbprint
                if(-not $rgxSmtpAddr){$rgxSmtpAddr = "^([0-9a-zA-Z]+[-._+&'])*[0-9a-zA-Z]+@([-0-9a-zA-Z]+[.])+[a-zA-Z]{2,63}$" ; } ; # email addr/UPN
                if(-not $rgxDomainLogon){$rgxDomainLogon = '^[a-zA-Z][a-zA-Z0-9\-\.]{0,61}[a-zA-Z]\\\w[\w\.\- ]+$' } ; # DOMAIN\samaccountname
                if(-not $exoMbxGraceDays){$exoMbxGraceDays = 30} ;
                if(-not $XOConnectionUri ){$XOConnectionUri = 'https://outlook.office365.com'} ;
                if(-not $SCConnectionUri){$SCConnectionUri = 'https://ps.compliance.protection.outlook.com'} ;
                if(-not $XODefaultPrefix){$XODefaultPrefix = 'xo' };
                if(-not $SCDefaultPrefix){$SCDefaultPrefix = 'sc' }; 
                #$rgxADDistNameGAT = ",$(($TORMeta.UnreplicatedOU -split ',' | select -skip 1 ) -join ',')"
                #$rgxADDistNameAT = ",$(($TORMeta.UnreplicatedOU -split ',' | select -skip 2 ) -join ',')"

                write-verbose "Coerce configured but blank Resultsize to Unlimited" ;
                if(get-variable -name resultsize -ea 0){
                    if( ($null -eq $ResultSize) -OR ('' -eq $ResultSize) ){$ResultSize = 'unlimited' }
                    elseif($Resultsize -is [int]){} else {throw "Resultsize must be an integer or the string 'unlimited' (or blank)"} ;
                } ;
                #$ComputerName = $env:COMPUTERNAME ;
                #$NoProf = [bool]([Environment]::GetCommandLineArgs() -like '-noprofile'); # if($NoProf){# do this};
                # XXXMeta derived constants:
                # - AADU Licensing group checks
                # calc the rgxLicGrpName fr the existing $xxxmeta.rgxLicGrpDN: (get-variable tormeta).value.rgxLicGrpDN.split(',')[0].replace('^','').replace('CN=','')
                #$rgxLicGrpName = (get-variable -name "$($tenorg)meta").value.rgxLicGrpDN.split(',')[0].replace('^','').replace('CN=','')
                # use the dn vers LicGrouppDN = $null ; # | ?{$_ -match $tormeta.rgxLicGrpDN}
                #$rgxLicGrpDN = (get-variable -name "$($tenorg)meta").value.rgxLicGrpDN
                # email trigger vari, it will be semi-delimd list of mail-triggering events
                $script:PassStatus = $null ;
                # TenOrg or other looped-specific PassStatus (auto supported by 7pswlt)
                #New-Variable -Name PassStatus_$($tenorg) -scope Script -Value $null ;
                [array]$SmtpAttachment = $null ;
                #write-verbose "start-Timer:Master" ;
                $swM = [Diagnostics.Stopwatch]::StartNew() ;
                # $ByPassLocalExchangeServerTest = $true # rough in, code exists below for exempting service/regkey testing on this variable status. Not yet implemented beyond the exemption code, ported in from orig source.
                #endregion COMMON_CONSTANTS ; #*------^ END COMMON_CONSTANTS ^------

                #region LOCAL_CONSTANTS ; #*------v LOCAL_CONSTANTS v------
                # BELOW TRIGGERS/DRIVES TEST_MODS: array of: "[modname];[modDLUrl,or pscmdline install]"
                <#$tDepModules = @("Microsoft.Graph.Authentication;https://www.powershellgallery.com/packages/Microsoft.Graph/",
                "ExchangeOnlineManagement;https://www.powershellgallery.com/packages/ExchangeOnlineManagement/",
                "ActiveDirectory;get-windowscapability -name RSAT* -Online | ?{$_.name -match 'Rsat\.ActiveDirectory'} | %{Add-WindowsCapability -online -name $_.name}"
                #,"AzureAD;https://www.powershellgallery.com/packages/AzureAD"
                ) ;
                #>
                $tDepModules = @() ; 
                if($useEXO){$tDepModules += @("ExchangeOnlineManagement;https://www.powershellgallery.com/packages/ExchangeOnlineManagement/;Get-xoOrganizationConfig",'verb-exo;localRepo;connect-exo')} ;
                if($UseMSOL){$tDepModules += @("MSOnline;https://www.powershellgallery.com/packages/MSOnline/;Get-MsolDomain")} ;
                if($UseAAD){$tDepModules += @("AzureAD;https://www.powershellgallery.com/packages/AzureAD/;Get-AzureADTenantDetail")} ;
                if($UseExOP){$tDepModules += @('verb-Ex2010;localRepo;Connect-Ex2010')} ;
                if($UseMG){$tDepModules += @("Microsoft.Graph.Authentication;https://www.powershellgallery.com/packages/Microsoft.Graph/;Get-MgOrganization")} ;
                if($UseOPAD){$tDepModules += @("ActiveDirectory;get-windowscapability -name RSAT* -Online | ?{$_.name -match 'Rsat\.ActiveDirectory'} | %{Add-WindowsCapability -online -name $_.name};Get-ADDomain")} ;

                $prpMGConnDeleg = 'Account','ClientId','TenantId','AuthType','ContextScope' ; 
                $prpMGConnCBA = 'CertificateSubjectName','CertificateThumbprint','Certificate' ; 
                $prpMGConnRet = $($prpMGConnDeleg;$prpMGConnCBA) ; 

                #region ENCODED_CONTANTS ; #*------v ENCODED_CONTANTS v------
                # ENCODED CONsTANTS & SUPPORT FUNCTIONS:
                #region 2B4 ; #*------v 2B4 v------
                if(-not (get-command 2b4 -ea 0)){function 2b4{[CmdletBinding()][Alias('convertTo-Base64String')] PARAM([Parameter(ValueFromPipeline=$true)][string[]]$str) ; PROCESS{$str|%{[Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($_))}  };} ; } ;
                #endregion 2B4 ; #*------^ END 2B4 ^------
                #region 2B4C ; #*------v 2B4C v------
                # comma-quoted return
                if(-not (get-command 2b4c -ea 0)){function 2b4c{ [CmdletBinding()][Alias('convertto-Base64StringCommaQuoted')] PARAM([Parameter(ValueFromPipeline=$true)][string[]]$str) ;BEGIN{$outs = @()} PROCESS{[array]$outs += $str | %{[Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($_))} ; } END {'"' + $(($outs) -join '","') + '"' | out-string | set-clipboard } ; } ; } ;
                #endregion 2B4C ; #*------^ END 2B4C ^------
                #region FB4 ; #*------v FB4 v------
                # DEMO: $SitesNameList = 'THluZGFsZQ==','U3BlbGxicm9vaw==','QWRlbGFpZGU=' | fb4 ;
                if(-not (get-command fb4 -ea 0)){function fb4{[CmdletBinding()][Alias('convertFrom-Base64String')] PARAM([Parameter(ValueFromPipeline=$true)][string[]]$str) ; PROCESS{$str | %{ [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($_)) }; } ; } ; };
                #endregion FB4 ; #*------^ END FB4 ^------
                # FOLLOWING CONSTANTS ARE USED FOR DEPENDANCY-LESS CONNECTIONS
                if(-not $CMW_logon_SID){$CMW_logon_SID = 'Q01XXGQtdG9kZC5rYWRyaWU=' | fb4 } ;
                if(-not $o365_Toroco_SIDUpn){$o365_Toroco_SIDUpn = 'cy10b2RkLmthZHJpZUB0b3JvLmNvbQ==' | fb4 } ;
                if(-not $TOR_logon_SID){$TOR_logon_SID = 'VE9ST1xrYWRyaXRzcw==' | fb4 } ;

                #endregion ENCODED_CONTANTS ; #*------^ END ENCODED_CONTANTS ^------

                #endregion CONSTANTS_AND_ENVIRO ; #*======^ CONSTANTS_AND_ENVIRO ^======

                #region SUBMAIN ; #*======v SUB MAIN v======

                #region TEST_MODS ; #*------v TEST_MODS v------
                if($tDepModules){
                    if( (test-ModulesAvailable -ModuleSpecifications $tDepModules) -contains $false ){
                        $smsg += "MISSING DEPENDANT MODULE!(see errors above)" ;
                        $smsg += "`n(may require provisioning internal function versions for this niche)" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ; 
                } ;
                #endregion TEST_MODS ; #*------^ END TEST_MODS ^------

                if($host.version.major -ge 3){$ret_ccO365S=[ordered]@{Dummy = $null ;} }
                else {$ret_ccO365S = @{Dummy = $null ;} } ;
                if($ret_ccO365S.keys -contains 'dummy'){$ret_ccO365S.remove('Dummy') };
                $fieldsBoolean = 'hasEXO','hasSC','hasMSOL','hasAAD','hasMG' | select -unique  | sort ; $fieldsBoolean | % { $ret_ccO365S.add($_,$false) } ;
                $fieldsnull = 'CredentialO365','UserPrincipalNameO365','MGContext','MGtoken' | select -unique  | sort ; $fieldsnull | % { $ret_ccO365S.add($_,$null) } ;

                 # PRETUNE STEERING separately *before* pasting in balance of region
                # THIS BLOCK DEPS ON VERB-* FANCY CRED/AUTH HANDLING MODULES THAT *MUST* BE INSTALLED LOCALLY TO FUNCTION
                # NOTE: *DOES* INCLUDE *PARTIAL* DEP-LESS $useExopNoDep=$true OPT THAT LEVERAGES Connect-ExchangeServerTDO, VS connect-ex2010 & CREDS ARE ASSUMED INHERENT TO THE ACCOUNT)
                # Connect-ExchangeServerTDO HAS SUBSTANTIAL BENEFIT, OF WORKING SEAMLESSLY ON EDGE SERVER AND RANGE OF DOMAIN-=CONNECTED EXOP ROLES
                <#
                $useO365 = $true ;
                $useEXO = $true ;
                $useSC = $TRUE ; 
                $UseMSOL = $false ; # should be hard disabled now in o365
                $UseAAD = $false  ;
                $UseMG = $true ;
                #>
                #Optional Array of MG Permission Names(avoids manual discovery against configured cmdlets) @('Domain.Read.All','Domain.ReadWrite.All','Directory.Read.All') ]")]
                if($UseMG -AND -not (get-variable -name MGPermissionsScope -ea 0).value){
                    [string[]]$MGPermissionsScope = @() ;
                    # if $MGPermissionsScope is omitted, get-MGCodeCmdletPermissionsTDO will be run to discover -  via Find-MGGraphCommand - and resolve into working ACL Scopes for connect-mgGraph
                    # if $MgCmdlets is populated with a an Array of -MG*/Microsoft.Graph* cmdlets, AST Parser details will not be run by get-MGCodeCmdletPermissionsTDO, solely the leaf Find-MGGraphCommand
                    if(-not (get-variable -name MGCmdlets  -ea 0).value){[string[]]$MGCmdlets = @()} ;
                } ;
                if($env:userdomain -eq $env:computername){
                    $isNonDomainServer = $true ;
                    $UseOPAD = $false ;
                }
                if($IsEdgeTransport){
                    #$UseExOP = $true ;
                    if($IsEdgeTransport -AND $psise){
                        $smsg = "powershell_ISE UNDER Exchange Edge Transport role!"
                        $smsg += "`nThis script is likely to fail the get-messagetrackingLog calls with Access Denied errors"
                        $smsg += "`nif run with this combo."
                        $smsg += "`nEXIT POWERSHELL ISE, AND RUN THIS DIRECTLY UNDER EMS FOR EDGE USE";
                        $smsg += "`n(bug appears to be a conflict in Remote EMS v EMS access permissions, not resolved yet)" ;
                        write-warning $msgs ;
                    } ;
                } ;
                $useO365 = [boolean]($useO365 -OR $useEXO -or $useSC -OR $UseMSOL -OR $UseAAD -OR $UseMG) ; 
                #$UseOP = [boolean]($UseOP -OR $UseExOP -OR $UseOPAD) ;
                #*------^ END STEERING VARIS ^------
                #*------v EXO V2/3 steering constants v------
                $EOMModName =  'ExchangeOnlineManagement' ;
                $EOMMinNoWinRMVersion = $MinNoWinRMVersion = '3.0.0' ; # support both names
                #*------^ END EXO V2/3 steering constants ^------
                # assert Org from Credential specs (if not param'd)
                # 1:36 PM 7/7/2023 and revised again -  revised the -AND, for both, logic wasn't working
                if($TenOrg){
                    $smsg = "Confirmed populated `$TenOrg" ;
                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                } elseif(-not($tenOrg) -and $Credential){
                    $smsg = "(unconfigured `$TenOrg: asserting from credential)" ;
                    if($silent){}elseif ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Prompt }
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    if((get-command get-TenantTag).Parameters.keys -contains 'silent'){
                        $TenOrg = get-TenantTag -Credential $Credential -silent ;;
                    }else {
                        $TenOrg = get-TenantTag -Credential $Credential ;
                    }
                } elseif(-not($tenOrg) -and $AdminAccount){
                        $smsg = "(unconfigured `$TenOrg: asserting from AdminAccount)" ;
                        if($silent){}elseif ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Prompt }
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        # convert UPN to cred for get-tenanttag handling
                        $tmpCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList ($AdminAccount,(convertto-securestring -string "passworddummy" -asplaintext -force)) ;
                        if((get-command get-TenantTag).Parameters.keys -contains 'silent'){
                            $TenOrg = get-TenantTag -Credential $tmpCredential -silent ;;
                        }else {
                            $TenOrg = get-TenantTag -Credential $tmpCredential ;
                        }
                } else {
                    # if not using Credentials or a TargetTenants/TenOrg loop, default the $TenOrg on the $env:USERDOMAIN
                    $smsg = "(unconfigured `$TenOrg & *NO* `$Credential: fallback asserting from `$env:USERDOMAIN)" ;
                    if($silent){}elseif ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Prompt }
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    switch -regex ($env:USERDOMAIN){
                        ([regex]('(' + (( @($TORMeta.legacyDomain,$CMWMeta.legacyDomain)  |foreach-object{[regex]::escape($_)}) -join '|') + ')')).tostring() {$TenOrg = $env:USERDOMAIN.substring(0,3).toupper() } ;
                        $TOLMeta.legacyDomain {$TenOrg = 'TOL' }
                        $env:COMPUTERNAME {
                            # non-domain-joined, no domain, but the $NetSummary.fqdn has a dns suffix that can be steered.
                            if($NetSummary.fqdn){
                                switch -regex (($NetSummary.fqdn.split('.') | select -last 2 ) -join '.'){
                                    'toro\.com$' {$tenorg = 'TOR' ; } ;
                                    '(charlesmachineworks\.com|cmw\.internal)$' { $TenOrg = 'CMW'} ;
                                    '(torolab\.com|snowthrower\.com)$'  { $TenOrg = 'TOL'} ;
                                    default {throw "UNRECOGNIZED DNS SUFFIX!:$(($NetSummary.fqdn.split('.') | select -last 2 ) -join '.')" ; break ; } ;
                                } ;
                            }else{
                                throw "NIC.ip $($NetSummary.ipaddress) does not PTR resolve to a DNS A with a full fqdn!" ;
                            } ;
                        } ;
                        default {throw "UNRECOGNIZED `$env:USERDOMAIN!:$($env:USERDOMAIN)" ; exit ; } ;
                    } ;
                } ;
                #region useO365 ; #*------v useO365 v------
                #$useO365 = $false ; # non-dyn setting, drives variant EXO reconnect & query code
                #if($CloudFirst){ $useO365 = $true } ; # expl: steering on a parameter
                if($useO365){
                    # creds are handled in cxo, don't need them for calls
                } else {
                    $smsg = "(`$useO365:$($useO365))" ;
                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                } ; # if-E if($useO365 ){
                #endregion useO365 ; #*------^ END useO365 ^------

            } ; # BEG-E
            PROCESS {

                #region SERVICE_CONNECTIONS #*======v SERVICE_CONNECTIONS v======

                #region useEXO ; #*------v useEXO v------
                # 1:29 PM 9/15/2022 as of MFA & v205, have to load EXO *before* any EXOP, or gen get-steppablepipeline suffix conflict error
                if($useEXO){
                    $pltCXO = [ordered]@{
                        Prefix = $XODefaultPrefix ;
                        TenOrg = $TenOrg ; 
                        Silent = $($silent) ; 
                        #Verbose = ($PSBoundParameters['Verbose'] -eq $true); 
                    } ;
                    if($AdminAccount){
                        $pltCXO.add('UserPrincipalName',$AdminAccount) ; 
                    } ; 
                    if($Credential){
                        $pltCXO.add('Credential',$Credential) ; 
                    } ; 
                    if(-not ($AdminAccount -OR $Credential) -AND $UserRole){
                        $pltCXO.add('UserRole',$UserRole) ; 
                    } ; 
                    $smsg = "Connect-EXO w`n$(($pltCXO|out-string).trim())" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Prompt }
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    connect-exo @pltCXO ; 
                } else {
                    $smsg = "(`$useEXO:$($useEXO))" ;
                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                } ; # if-E
                #endregion useEXO ; #*------^ END useEXO ^------
                #region useSC ; #*------v useSC v------
                # 1:29 PM 9/15/2022 as of MFA & v205, have to load EXO *before* any EXOP, or gen get-steppablepipeline suffix conflict error
                if($useSC){
                    $pltCSC = [ordered]@{
                        Prefix = $SCDefaultPrefix ;
                        TenOrg = $TenOrg ; 
                        connectPurview = $true ; 
                        Silent = $($silent) ; 
                        #Verbose = ($PSBoundParameters['Verbose'] -eq $true); 
                    } ;
                    if($AdminAccount){
                        $pltCSC.add('UserPrincipalName',$AdminAccount) ; 
                    } ; 
                    if($Credential){
                        $pltCSC.add('Credential',$Credential) ; 
                    } ; 
                    if(-not ($AdminAccount -OR $Credential) -AND $UserRole){
                        $pltCSC.add('UserRole',$UserRole) ; 
                    } ; 
                    $smsg = "Connect-SC (Connect-IPPSSession Purview) w`n$(($pltCSC|out-string).trim())" ; 
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Prompt }
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    connect-exo @pltCSC ; 
                } else {
                    $smsg = "(`$useSC:$($useSC))" ;
                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                } ; # if-E
                #endregion useSC ; #*------^ END useSC ^------
                #region VALIDATE_XOSC ; #*------v VALIDATE_XOSC v------
                if($useEXO -OR $useSC){
                    $XOconnections = test-exoconnectiontdo ; 
                    foreach($xcon in $XOconnections){
                        if($xcon.connection -ANd $xcon.isXO -ANd $xcon.isValid -AND $xcon.TokenLifeMins -gt 0){
                            $ret_rxo = $xcon; $ret_ccO365S.hasEXO = $true ; 
                            if($rai = (resolve-AppIDToCBAFriendlyName -AppId $xcon.AppID)){
                                $o365Cred = (gv -name "cred$($rai.service)$($rai.tenorg)$($rai.friendlyname.split('_')[1].split('-')[0].replace('Cert',''))" -ea STOP).value ; 
                            }
                        } # else {$ret_rxo = $null ; $ret_ccO365S.hasEXO = $false } ;
                        if($xcon.connection -ANd $xcon.isSC -ANd $xcon.isValid -AND $xcon.TokenLifeMins -gt 0){
                            $ret_rSC = $xcon; $ret_ccO365S.hasSC = $true
                            if(-not $o365Cred){
                                if($rai = (resolve-AppIDToCBAFriendlyName -AppId $xcon.AppID)){
                                    $o365Cred = (gv -name "cred$($rai.service)$($rai.tenorg)$($rai.friendlyname.split('_')[1].split('-')[0].replace('Cert',''))" -ea STOP).value ; 
                                }
                            }
                        } # else {$ret_rSC = $null; $ret_ccO365S.hasSC = $false } ;
                    } ; 
                } ; 
                if(($ret_ccO365S.hasEXO -OR $ret_ccO365S.hasSC) -AND $o365Cred){
                    $smsg = "Returning EXO Credential" ; 
                    if($VerbosePreference -eq "Continue"){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                    $ret_ccO365S.CredentialO365 = $o365Cred  ; 
                }
                #endregion VALIDATE_XOSC ; #*------^ END VALIDATE_XOSC ^------
                #region MSOL_CONNECTION ; #*------v  MSOL CONNECTION v------
                #$UseMSOL = $false
                if($UseMSOL){
                    $smsg = "(loading MSOL...)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } Error|Warn|Debug
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    connect-msol @pltRXOC ;
                    TRY{$MsolCo = Get-MsolCompanyInformation -ea stop ; $ret_ccO365S.hasMSOL = $true} CATCH {$ret_ccO365S.hasMSOL = $false } 
                } else {
                    $smsg = "(`$UseMSOL:$($UseMSOL))" ;
                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                } ;
                #endregion MSOL_CONNECTION ; #*------^  MSOL CONNECTION ^------

                #region AZUREAD_CONNECTION ; #*------v AZUREAD CONNECTION v------
                #$UseAAD = $false
                if($UseAAD){
                    #$reqMods += "Connect-AAD".split(";") ;
                    #if ( !(check-ReqMods $reqMods) ) { write-error "$((get-date).ToString("yyyyMMdd HH:mm:ss")):Missing function. EXITING." ; Break ; }  ;
                    $smsg = "(loading AAD...)" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    Connect-AAD @pltRXOC ;
                    TRY{$AADTenant = Get-AzureADTenantDetail -ea stop ; $ret_ccO365S.hasAAD = $true} CATCH {$ret_ccO365S.hasAAD = $false }  ; 
                } else {
                    $smsg = "(`$UseAAD:$($UseAAD))" ;
                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                } ;
                #endregion AZUREAD_CONNECTION ; #*------^ AZUREAD CONNECTION ^------

                #region CONNECT_MG ; #*------v CONNECT_MG v------
                #$UseMG = $false;
                if($UseMG){
                    if(-not (get-command  test-mgconnection)){
                        TRY{
                            ipmo -fo -verb verb-mg -verbose 
                        } CATCH {$ErrTrapd=$Error[0] ;
                            write-host -foregroundcolor gray "TargetCatch:} CATCH [$($ErrTrapd.Exception.GetType().FullName)] {"  ;
                            $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                            write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" ;
                            BREAK ; 
                        } ;
                    } ;
                    $MGCntxt = test-mgconnection -Verbose:($VerbosePreference -eq 'Continue') ;
                    if($MGCntxt.CertificateThumbprint -match $rgxCertThumbprint -AND $MgCntxt.AppName -AND $mgCntxt.isCBA){
                        $smsg = "CBA-auth'd existing conn: skipping RequirePermissions scope discovery" ; 
                        if($VerbosePreference -eq "Continue"){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                    }elseif(-not $MGPermissionsScope){
                        if(gi function:get-MGCodeCmdletPermissionsTDO -ea 0){
                            $pltGMGP=[ordered]@{
                                #whatif = $($whatif) ;
                            } ;
                            if($MgCmdlets){
                                $pltGMGP.add('Cmdlets',$MgCmdlets)  ;
                            }else{
                                if($EnvSummary.isScript){
                                    if($EnvSummary.PSCommandPathproxy){ $prxPath = $EnvSummary.PSCommandPathproxy }
                                    elseif($script:PSCommandPath){$prxPath = $script:PSCommandPath}
                                    elseif($rPSCommandPath){$prxPath = $rPSCommandPath} ;
                                    $pltGMGP.add('Path',$prxPath)  ;
                                }elseif($EnvSummary.isFunc){
                                    $pltGMGP.add('scriptblock',(get-command -name $EnvSummary.FuncName).definition) ;
                                }else{
                                    $smsg = "MISSING or INDETERMINANT `$EnvSummary.isScript/`$EnvSummary.isFunc (should be output of verb-io\resolve-EnvironmentTDO())" ;
                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                    throw $smsg ;
                                    BREAK ;
                                } ;
                            } ;
                            $smsg = "get-MGCodeCmdletPermissionsTDO w`n$(($pltGMGP|out-string).trim())" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                            $RequiredScopes = $PermsRqd = get-MGCodeCmdletPermissionsTDO @pltGMGP ;
                            $smsg = "`nResolved MGPermissionsScope:`n$(($PermsRqd |out-string).trim())" ;
                            $smsg +="`n(can be hardcoded into script's `$MGPermissionsScope to save query time on future passes)" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        } else {
                            if(-not $MGCmdlets){
                                $smsg = "-useMG:$($useMG): Performing *manual* MGCodeCmdletPermissions expansion: (missing function)" ;
                                $smsg += "`ncannot procede with CURRENTLY EMPTY `$MGCmdlets!"
                                $smsg = "`n(should contain all [verb]-mg[noun] cmdlets to be used this session)" ;
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                                throw $smsg ;
                                BREAK ;
                            } ;
                            $ttl = $MGCmdlets  |  measure | select -expand count ;
                            write-host -foregroundcolor yellow "$($ttl) Cmdlets:Collecting Permissions[" -NoNewline ;
                            $PermsRqd = @() ;
                            $MGCmdlets  |foreach-object{
                                write-host -NoNewline '.' ;
                                if($ACL = Find-MgGraphCommand -command $_ -ea 0){
                                    $PermsRqd += $ACL | Select -First 1 -ExpandProperty Permissions | Select -Unique name ;
                                } ;
                            } ;
                            write-host -foregroundcolor yellow "]" ;
                            $PermsRqd = $PermsRqd.name | select -unique ;
                            $RequiredScopes = $PermsRqd ; 
                            $smsg = "`nResolved MGPermissionsScope:`n$(($PermsRqd |out-string).trim())" ;
                            $smsg +="`n(can be hardcoded into script's `$MGPermissionsScope to save query time on future passes)" ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                            else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        } ;
                    }else{
                        $smsg = "Using explicit -MGPermissionsScope specified: $(($MGPermissionsScope | select -first 3) -join ',')..."
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                        $PermsRqd = $MGPermissionsScope ;
                    }

                    #region cMG_SCAFFOLD ; #*------v cMG_SCAFFOLD v------
                    if(-not (get-command  test-mgconnection)){
                        TRY{
                            ipmo -fo -verb verb-mg -verbose 
                        } CATCH {$ErrTrapd=$Error[0] ;
                            write-host -foregroundcolor gray "TargetCatch:} CATCH [$($ErrTrapd.Exception.GetType().FullName)] {"  ;
                            $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                            write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" ;
                            BREAK ; 
                        } ;
                    } ;
                    $MGCntxt = test-mgconnection -Verbose:($VerbosePreference -eq 'Continue') ;
                    $o365Cred = $null ;
                    if($Credential -AND $MGCntxt.isConnected){
                        $smsg = "Explicit -Credential:$($Credential.username) -AND `$MGCntxt.isConnected: running pre:Disconnect-MgGraph" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        # Dmg returns a get-mgcontext into pipe, if you don't cap it corrupts the pipe on your current flow
                        $dOut = Disconnect-MgGraph -Verbose:($VerbosePreference -eq 'Continue')
                        $MGCntxt = test-mgconnection -Verbose:($VerbosePreference -eq 'Continue') ;
                    };
                    if($Credential){
                        $smsg = "`Credential:Explicit credentials specified, deferring to use..." ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        write-verbose "get-TenantCredentials() return format: (emulating)" ; 
                        $o365Cred = [ordered]@{
                            Cred=$Credential ;
                            credType=$null ;
                        } ;
                        $uRoleReturn = resolve-UserNameToUserRole -UserName $Credential.username -verbose:$($VerbosePreference -eq "Continue") ; # Username
                        write-verbose "w full cred opt: $uRoleReturn = resolve-UserNameToUserRole -Credential $Credential -verbose = $($VerbosePreference -eq 'Continue')"  ; 
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
                        if($MGCntxt.isConnected){
                            if($MgCntxt.isUser){
                                $TenantTag = $TenOrg = get-TenantTag -Credential $MgCntxt.Account ;
                                $uRoleReturn = resolve-UserNameToUserRole -UserName $MgCntxt.CertificateThumbprint -verbose:$($VerbosePreference -eq "Continue") ;
                                $credential = get-TenantCredentials -TenOrg $TenOrg -UserRole $uRoleReturn.UserRole -verbose:$($VerbosePreference -eq "Continue") ;
                            } elseif($MgCntxt.isCBA -AND $MgCntxt.AppName -match 'CBACert-(\w{3})'){
                                    #$MgCntxt.AppName.split('-')[-1]
                                    $TenantTag = $TenOrg = $matches[1]
                                    # also need credential
                                    $uRoleReturn = resolve-UserNameToUserRole -UserName $MgCntxt.CertificateThumbprint -verbose:$($VerbosePreference -eq "Continue") ;
                                    write-verbose "ret'd obj:$uRoleReturn = [ordered]@{     UserRole = $null ;     Service = $null ;     TenOrg = $null ; } " ;  
                                    $credRet = get-TenantCredentials -TenOrg $TenOrg -UserRole $uRoleReturn.UserRole -verbose:$($VerbosePreference -eq "Continue")
                                    $credential = $credRet.Cred ;
                            }else{
                                $smsg = "UNABLE TO RESOLVE mgContext to a working TenOrg!" ;
                                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            }
                        } ; 
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
                        write-verbose "populate $credential with return, if not populated (may be required for follow-on calls that pass common $Credentials through)" ; 
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
                    $pltCMG = [ordered]@{
                        Credential = $Credential ;
                        verbose = $($VerbosePreference -eq "Continue")  ;
                        #ErrorVariable = 'err_ccMG' # connect-mg doesn't support -errvari!
                    } ;
                    <#if((get-command Connect-MG).Parameters.keys -contains 'silent'){
                        $pltCMG.add('Silent',$silent) ;
                    } ;
                    #>
                    #endregion cMG_SCAFFOLD ; #*------^ END cMG_SCAFFOLD ^------
                    #------
                    connect-MG @pltCMG 
                    # need to update the hasMG flag: TRY{$AADTenant = Get-AzureADTenantDetail -ea stop ; $ret_ccO365S.hasAAD = $true} CATCH {$ret_ccO365S.hasAAD = $false }  ; 
                    TRY{
                        $MGCntxt = test-mgconnection -Verbose:($VerbosePreference -eq 'Continue') -ea Stop ;
                        if($MGCntxt.isConnected){
                            $ret_ccO365S.hasMG = $true ; 
                            if($ret_ccO365S.hasMG -AND $pltCMG.Credential -AND -not $ret_ccO365S.CredentialO365){
                                $smsg = "returning MG Credential" ; 
                                if($VerbosePreference -eq "Continue"){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                                $ret_ccO365S.CredentialO365 = $pltCMG.Credential ; 
                            } ; 
                        }ELSE{$ret_ccO365S.hasMG = $FALSE}                        
                    } CATCH {
                        $ret_ccO365S.hasMG = $false 
                    }  
                    <#
                    if($err_ccMG -ne $null){
                        $smsg = $err_ccMG  ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        throw $smsg ;
                        BREAK ;
                    } ;
                    #>
                    <# CBA connect-mgGraph use: [Connect to Microsoft Graph PowerShell using Certificate - AdminDroid Blog](https://blog.admindroid.com/connect-to-microsoft-graph-powershell-using-certificate/)
                    # has full demo of suite of -MG cmdlets to provision & perm Apps w CBA, local copy:C:\usr\work\o365\scripts\ConnectMSGraphCertificate_admindroid-com.ps1
                    # -scope is perm hard-coded non-dyn w/in app, must have full suite of all acls it will ever need
                    $pltCCMG=[ordered]@{TenantId = $TenantID ;ClientId = $ClientID ;CertificateThumbprint = $CertificateThumbprint ;ErrorAction = 'SilentlyContinue' ;ErrorVariable = 'ApplicationConnectionError' ; } ;
                    write-host "Connect-mgGraph w`n$(($pltGMGP|out-string).trim())" ;
                    Connect-MgGraph @pltCCMG ;
                    if($ApplicationConnectionError -ne $null){
                        Write-Host $ApplicationConnectionError -ForegroundColor Red ;
                        Exit ;
                    } ; Get-MgContext ;
                    #>
                   <# 1:58 PM 1/6/2026 throws up crashes out conn attempt 
                    # resolve & store the MG goken:
                    $pltIGR = @{
                        Method = "GET" ; 
                        URI = "/v1.0/me" ; 
                        OutputType = "HttpResponseMessage" ; 
                    } ; 
                    $Response = Invoke-GraphRequest @pltIGR ; 
                    $Headers = $Response.RequestMessage.Headers ; 
                    if($TokenString = $Headers.Authorization.Parameter){
                        #$ret_ccO365S.MGtoken = $Headers.Authorization.Parameter ; 
                        $ret_ccO365S.MGtoken = ConvertTo-SecureString -String $TokenString -AsPlainText -Force ; 
                    }else {
                        $smsg = "Unable To Invoke-GraphRequest back a Token object!" ; 
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                        throw $smsg ; 
                        break ; 
                    } ; 
                    #>
                } ;  # if-E $useMG
                #endregion CONNECT_MG ; #*------^ END CONNECT_MG ^------
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
                if ($script:useEXOv2) { Connect-EXO2 @pltRXOC }
                else { Connect-EXO @pltRXOC } ;
                # reenable VerbosePreference:Continue, if set, during mod loads
                if($VerbosePrefPrior -eq "Continue"){
                    $VerbosePreference = $VerbosePrefPrior ;
                    $verbose = ($VerbosePreference -eq "Continue") ;
                } ;
                #>
                #-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
                #endregion SERVICE_CONNECTIONS #*======^ END SERVICE_CONNECTIONS ^======
            } ; # PROC-E
            END {
                $swM.Stop() ;
                $smsg = ("Elapsed Time: {0:dd}d {0:hh}h {0:mm}m {0:ss}s {0:fff}ms" -f $swM.Elapsed) ;
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                <# return status obj
                $ret_ccO365S = [ordered]@{
                    CredentialO365 = $null ; 
                    hasEXO = $false ;
                    hasMSOL = $false ;
                    hasAAD = $false ;
                    hasMG = $false ;
                    MGContext = $null ; # $ctxMG
                    MGtoken = $null ; 
                } ; 
                #>
                $smsg = "Returning connection summary to pipeline:`n$(($ret_ccO365S|out-string).trim())" ; 
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                [pscustomobject]$ret_ccO365S | write-output ;
            } ; # END-E
        } ;
    #} ;
    #endregion CONNECT_O365SERVICES ; #*======^ END connect-o365services ^======