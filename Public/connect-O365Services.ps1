# connect-O365Services.ps1

#region CONNECT_O365SERVICES ; #*======v connect-O365Services v======
if(-not (get-childitem function:connect-O365Services -ea 0)){
    function connect-O365Services {
        <#
        .SYNOPSIS
        connect-O365Services - logic wrapper for my histortical scriptblock that resolves creds, svc avail and relevent status, to connect to range of Services (in o365)
        .NOTES
        Version     : 0.0.
        Author      : Todd Kadrie
        Website     : http://www.toddomation.com
        Twitter     : @tostka / http://twitter.com/tostka
        CreatedDate : 2024-06-07
        FileName    : connect-O365Services
        License     : MIT License
        Copyright   : (c) 2024 Todd Kadrie
        Github      : https://github.com/tostka/verb-AAD
        Tags        : Powershell,AzureAD,Authentication,Test
        AddedCredit :
        AddedWebsite:
        AddedTwitter:
        REVISIONS
        * 1:01 PM 5/19/2025 rem'd $prefVaris dump (blank values, throws errors)
        * 3:34 PM 5/16/2025 spliced over local dep internal_funcs (out of the main paramt block) ; fixed typo in return vari name ret_ccO365S
        * 8:16 AM 5/15/2025 init
        .DESCRIPTION
        connect-O365Services - logic wrapper for my histortical scriptblock that resolves creds, svc avail and relevent status, to connect to range of Services (in o365)
        .PARAMETER EnvSummary
        Pre-resolved local environrment summary (product of output of verb-io\resolve-EnvironmentTDO())[-EnvSummary `$rvEnv]
        .PARAMETER NetSummary
        Pre-resolved local network summary (product of output of verb-network\resolve-NetworkLocalTDO())[-NetSummary `$netsettings]
        .PARAMETER XoPSummary
        Pre-resolved local ExchangeServer summary (product of output of verb-ex2010\test-LocalExchangeInfoTDOO())[-XoPSummary `$lclExOP]
        .PARAMETER useEXO
        Connect to O365 ExchangeOnlineManagement)[-useEXO]
        .PARAMETER UseExOP
        Connect to OnPrem ExchangeManagementShell(Remote (Local,Edge))[-UseExOP]
        .PARAMETER useExopNoDep
        Connect to OnPrem ExchangeManagementShell using No Dependancy options)[-useEXO]
        .PARAMETER ExopVers
        Connect to OnPrem ExchangeServer version (Ex2019|Ex2016|Ex2013|Ex2010|Ex2007|Ex2003|Ex2000). An array represents a min/max range of all between; null indicates all versions returned by verb-Ex2010\get-ADExchangeServerTDO())[-useEXO]
        XOP Switch to set ForestWide Exchange EMS scope(e.g. Set-AdServerSettings -ViewEntireForest `$True)[-useForestWide]
        .PARAMETER UseOPAD
        Connect to OnPrem ActiveDirectory powershell module)[-UseOPAD]
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
        .PARAMETER UserRole
        Credential User Role spec (SID|CSID|UID|B2BI|CSVC|ESVC|LSVC|ESvcCBA|CSvcCBA|SIDCBA)[-UserRole @('SIDCBA','SID','CSVC')]
        .PARAMETER useEXOv2
        Use EXOv2 (ExchangeOnlineManagement) over basic auth legacy connection [-useEXOv2]
        .PARAMETER Silent
        Silent output (suppress status echos)[-silent]
        .PARAMETER MGPermissionsScope
        Optional Array of MG Permission Names(avoids manual discovery against configured cmdlets)[-MGPermissionsScope @('Domain.Read.All','Domain.ReadWrite.All','Directory.Read.All') ]
        .PARAMETER useExOPVers
        String array to indicate target OnPrem Exchange Server version to target with connections, if an array, will be assumed to reflect a span of versions to include, connections will aways be to a random server of the latest version specified (Ex2000|Ex2003|Ex2007|Ex2010|Ex2000|Ex2003|Ex2007|Ex2010|Ex2016|Ex2019), used with verb-Ex2010\get-ADExchangeServerTDO() dyn location via ActiveDirectory.[-useExOPVers @('Ex2010','Ex2016')]")]
        .INPUTS
        Does not accept piped input
        .OUTPUTS
        None (records transcript file)
        .EXAMPLE
        PS> $PermsRqd = connect-O365Services -path D:\scripts\new-MGDomainRegTDO.ps1 ;
        Typical pass script pass, using the -path param
        .EXAMPLE
        PS> $PermsRqd = connect-O365Services -scriptblock (gcm -name connect-O365Services).definition ;
        Typical function pass, using get-command to return the definition/scriptblock for the subject function.
        .EXAMPLE
        PS> write-verbose "Typically from the BEGIN{} block of an Advanced Function, or immediately after PARAM() block" ;
        PS> $Verbose = [boolean]($VerbosePreference -eq 'Continue') ;
        PS> $rPSCmdlet = $PSCmdlet ; 
        PS> $rPSScriptRoot = $PSScriptRoot ; 
        PS> $rPSCommandPath = $PSCommandPath ; 
        PS> $rMyInvocation = $MyInvocation ; 
        PS> $rPSBoundParameters = $PSBoundParameters ; 
        PS> $pltRvEnv=[ordered]@{
        PS>     PSCmdletproxy = $rPSCmdlet ; 
        PS>     PSScriptRootproxy = $rPSScriptRoot ; 
        PS>     PSCommandPathproxy = $rPSCommandPath ; 
        PS>     MyInvocationproxy = $rMyInvocation ;
        PS>     PSBoundParametersproxy = $rPSBoundParameters
        PS>     verbose = [boolean]($PSBoundParameters['Verbose'] -eq $true) ; 
        PS> } ;
        PS> $smsg = "resolve-EnvironmentTDO w`n$(($pltRVEnv|out-string).trim())" ; 
        PS> if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        PS> $rvEnv = resolve-EnvironmentTDO @pltRVEnv ; 
        PS> $smsg = "`$rvEnv returned:`n$(($rvEnv |out-string).trim())" ; 
        PS> if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
        PS> else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
        PS> $netsettings = resolve-NetworkLocalTDO ; 
        PS> $lclExOP = test-LocalExchangeInfoTDO ; 
        PS> $pltCco365Svcs=[ordered]@{
        PS>     EnvSummary = $rvEnv ;
        PS>     NetSummary = $netsettings ;
        PS>     XoPSummary = $lclExOP ;
        PS>     useEXO = $true ;
        PS>     UseMSOL = $false ;
        PS>     UseAAD = $false ;
        PS>     UseMG = $true ;
        PS>     TenOrg = $global:o365_TenOrgDefault ;
        PS>     Credential = $null ;
        PS>     UserRole = @('SID','CSVC') ;
        PS>     # svcAcct use: @('ESvcCBA','CSvcCBA','SIDCBA')
        PS>     useEXOv2 = $true ;
        PS>     silent = $false ;
        PS>     MGPermissionsScope = $null ;
        PS>     MGCmdlets = $null ;
        PS> } ;
        PS> $smsg = "connect-O365Services w`n$(($pltCco365Svcs|out-string).trim())" ;
        PS> if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        PS> $ret_ccO365S = connect-O365Services @pltCco365Svcs ; 
        Demo leveraging verb-io\resolve-EnvironmentTDO(), verb-network\resolve-NetworkLocalTDO() & verb-ex2010\test-LocalExchangeInfoTDO() to provide relevent inputs
        .LINK
        https://bitbucket.org/tostka/verb-dev/
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
            [Parameter(Mandatory=$true,HelpMessage="Pre-resolved local ExchangeServer summary (product of output of verb-ex2010\test-LocalExchangeInfoTDOO())[-XoPSummary `$lclExOP]")]
                $XoPSummary, # $lclExOP = test-LocalExchangeInfoTDO ;
            # service choices
            #$useO365 = $true ; - intterpolate it from the other svcs
            [Parameter(HelpMessage="Connect to O365 ExchangeOnlineManagement)[-useEXO]")]
                [switch]$useEXO,
            <# OP switches
            #[Parameter(HelpMessage="Connect to OnPrem ExchangeManagementShell(Remote (Local,Edge))[-UseOP]")]
            #    [switch]$UseOP, # interpolate from below
            [Parameter(HelpMessage="Connect to OnPrem ExchangeManagementShell(Remote (Local,Edge))[-UseExOP]")]
                [switch]$UseExOP,
            [Parameter(HelpMessage="Connect to OnPrem ExchangeManagementShell using No Dependancy options)[-useEXO]")]
                [switch]$useExopNoDep,
            [Parameter(HelpMessage="Connect to OnPrem ExchangeServer version (Ex2019|Ex2016|Ex2013|Ex2010|Ex2007|Ex2003|Ex2000). An array represents a min/max range of all between; null indicates all versions returned by verb-Ex2010\get-ADExchangeServerTDO())[-useEXO]")]
                [AllowNull()]
                [ValidateSet('Ex2019','Ex2016','Ex2013','Ex2010','Ex2007','Ex2003','Ex2000')]
                [switch]$ExopVers, # = 'Ex2010' # 'Ex2019','Ex2016','Ex2013','Ex2010','Ex2007','Ex2003','Ex2000', Null for All versions
                #if($Version){
                #   $ExopVers = $Version ; #defer to local script $version if set
                #} ;
            [Parameter(HelpMessage="XOP Switch to set ForestWide Exchange EMS scope(e.g. Set-AdServerSettings -ViewEntireForest `$True)[-useForestWide]")]
                [switch]$useForestWide,
            [Parameter(HelpMessage="Connect to OnPrem ActiveDirectory powershell module)[-UseOPAD]")]
                [switch]$UseOPAD,
            #>
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
                    Version     : 1.1.0
                    Author      : Todd Kadrie
                    Website     : https://www.toddomation.com
                    Twitter     : @tostka / http://twitter.com/tostka
                    CreatedDate : 3:56 PM 12/8/2019
                    FileName    : get-CodeProfileAST.ps1
                    License     : MIT License
                    Copyright   : (c) 2019 Todd Kadrie
                    Github      : https://github.com/tostka
                    AddedCredit :
                    AddedWebsite:
                    AddedTwitter:
                    REVISIONS
                    * 4:11 PM 5/15/2025 add psv2-ordered compat
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
                    .INPUTS
                    None
                    .OUTPUTS
                    Outputs a system.object containing:
                    * Parameters : Details on all Parameters in the file
                    * Functions : Details on all Functions in the file
                    * VariableAssignments : Details on all Variables assigned in the file
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
                    Version     : 0.0.
                    Author      : Todd Kadrie
                    Website     : http://www.toddomation.com
                    Twitter     : @tostka / http://twitter.com/tostka
                    CreatedDate : 2024-06-07
                    FileName    : get-MGCodeCmdletPermissionsTDO
                    License     : MIT License
                    Copyright   : (c) 2024 Todd Kadrie
                    Github      : https://github.com/tostka/verb-AAD
                    Tags        : Powershell,AzureAD,Authentication,Test
                    AddedCredit : 
                    AddedWebsite: 
                    AddedTwitter: 
                    REVISIONS
                    * 1:49 PM 5/14/2025 add: -cmdlets, bypasses AST parsing cuts right to find-mgGraphCommand expansion; additional verbose status echos (as it's returning very limited set of perms)
                    * 4:37 PM 5/12/2025 retweaked expansion; found that the cmdlet name filtering wasn't working as a raw [regex], had to .tostring() the regex to get it to return more than a single item
                    * 12:23 PM 5/6/2025 wrapper for verb-dev\get-codeprofileAST() that parses [verb]-MG[noun] cmdlets from a specified -file or -scriptblock, and reseolves the necessary connect-mgGraph delegated access -scope permissions, using the Find-MgGraphCommand command.
                    .DESCRIPTION
                    wrapper for verb-dev\get-codeprofileAST() that parses [verb]-MG[noun] cmdlets from a specified -file or -scriptblock, and reseolves the necessary connect-mgGraph -scope permissions, using the Find-MgGraphCommand command.
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
                    .INPUTS
                    Does not accept piped input
                    .OUTPUTS
                    None (records transcript file)
                    .EXAMPLE
                    PS> $PermsRqd = get-MGCodeCmdletPermissionsTDO -path D:\scripts\new-MGDomainRegTDO.ps1 ; 
                    Typical pass script pass, using the -path param
                    .EXAMPLE
                    PS> $PermsRqd = get-MGCodeCmdletPermissionsTDO -scriptblock (gcm -name get-MGCodeCmdletPermissionsTDO).definition ; 
                    Typical function pass, using get-command to return the definition/scriptblock for the subject function.
                    .EXAMPLE
                    PS> write-verbose "Typically from the BEGIN{} block of an Advanced Function, or immediately after PARAM() block" ; 
                    PS> $Verbose = [boolean]($VerbosePreference -eq 'Continue') ;
                    PS> $rPSCmdlet = $PSCmdlet ;
                    PS> $rPSScriptRoot = $PSScriptRoot ;
                    PS> $rPSCommandPath = $PSCommandPath ;
                    PS> $rMyInvocation = $MyInvocation ;
                    PS> $rPSBoundParameters = $PSBoundParameters ;
                    PS> $pltRvEnv=[ordered]@{
                    PS>     PSCmdletproxy = $rPSCmdlet ;
                    PS>     PSScriptRootproxy = $rPSScriptRoot ;
                    PS>     PSCommandPathproxy = $rPSCommandPath ;
                    PS>     MyInvocationproxy = $rMyInvocation ;
                    PS>     PSBoundParametersproxy = $rPSBoundParameters
                    PS>     verbose = [boolean]($PSBoundParameters['Verbose'] -eq $true) ;
                    PS> } ;
                    PS> $rvEnv = resolve-EnvironmentTDO @pltRVEnv ;  
                    PS> if($rvEnv.isScript){
                    PS>     if($rvEnv.PSCommandPathproxy){ $prxPath = $rvEnv.PSCommandPathproxy }
                    PS>     elseif($script:PSCommandPath){$prxPath = $script:PSCommandPath}
                    PS>     elseif($rPSCommandPath){$prxPath = $rPSCommandPath} ; 
                    PS>     $PermsRqd = get-MGCodeCmdletPermissionsTDO -Path $prxPath  ; 
                    PS> } ; 
                    PS> if($rvEnv.isFunc){
                    PS>     $PermsRqd = get-MGCodeCmdletPermissionsTDO -Path (gcm -name $rvEnv.FuncName).definition ; 
                    PS> } ; 
                    Demo leveraging resolve-environmentTDO outputs
                    .LINK
                    https://bitbucket.org/tostka/verb-dev/
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
            #region ENVIRO_DISCOVER ; #*------v ENVIRO_DISCOVER v------
            <#
            $Verbose = [boolean]($VerbosePreference -eq 'Continue') ;
            $rPSCmdlet = $PSCmdlet ; # an object that represents the cmdlet or advanced function that's being run. Available on functions w CmdletBinding (& $args will not be available). (Blank on non-CmdletBinding/Non-Adv funcs).
            $rPSScriptRoot = $PSScriptRoot ; # the full path of the executing script's parent directory., PS2: valid only in script modules (.psm1). PS3+:it's valid in all scripts. (Funcs: ParentDir of the file that hosts the func)
            $rPSCommandPath = $PSCommandPath ; # the full path and filename of the script that's being run, or file hosting the funct. Valid in all scripts.
            $rMyInvocation = $MyInvocation ; # populated only for scripts, function, and script blocks.
            # - $MyInvocation.MyCommand.Name returns name of a function, to identify the current command,  name of the current script (pop'd w func name, on Advfuncs)
            # - Ps3+:$MyInvocation.PSScriptRoot : full path to the script that invoked the current command. The value of this property is populated only when the caller is a script (blank on funcs & Advfuncs)
            # - Ps3+:$MyInvocation.PSCommandPath : full path and filename of the script that invoked the current command. The value of this property is populated only when the caller is a script (blank on funcs & Advfuncs)
            #     ** note: above pair contain information about the _invoker or calling script_, not the current script
            $rPSBoundParameters = $PSBoundParameters ;
            #>
            #region PREF_VARI_DUMP ; #*------v PREF_VARI_DUMP v------
            <#$script:prefVaris = @{
                whatifIsPresent = $whatif.IsPresent
                whatifPSBoundParametersContains = $rPSBoundParameters.ContainsKey('WhatIf') ;
                whatifPSBoundParameters = $rPSBoundParameters['WhatIf'] ;
                WhatIfPreferenceIsPresent = $WhatIfPreference.IsPresent ; # -eq $true
                WhatIfPreferenceValue = $WhatIfPreference;
                WhatIfPreferenceParentScopeValue = (Get-Variable WhatIfPreference -Scope 1).Value ;
                ConfirmPSBoundParametersContains = $rPSBoundParameters.ContainsKey('Confirm') ;
                ConfirmPSBoundParameters = $rPSBoundParameters['Confirm'];
                ConfirmPreferenceIsPresent = $ConfirmPreference.IsPresent ; # -eq $true
                ConfirmPreferenceValue = $ConfirmPreference ;
                ConfirmPreferenceParentScopeValue = (Get-Variable ConfirmPreference -Scope 1).Value ;
                VerbosePSBoundParametersContains = $rPSBoundParameters.ContainsKey('Confirm') ;
                VerbosePSBoundParameters = $rPSBoundParameters['Verbose'] ;
                VerbosePreferenceIsPresent = $VerbosePreference.IsPresent ; # -eq $true
                VerbosePreferenceValue = $VerbosePreference ;
                VerbosePreferenceParentScopeValue = (Get-Variable VerbosePreference -Scope 1).Value;
                VerboseMyInvContains = '-Verbose' -in $rPSBoundParameters.UnboundArguments ;
                VerbosePSBoundParametersUnboundArgumentContains = '-Verbose' -in $rPSBoundParameters.UnboundArguments
            } ;
            write-verbose "`n$(($script:prefVaris.GetEnumerator() | Sort-Object Key | Format-Table Key,Value -AutoSize|out-string).trim())`n" ;
            #>
            #endregion PREF_VARI_DUMP ; #*------^ END PREF_VARI_DUMP ^------
            #region RV_ENVIRO ; #*------v RV_ENVIRO v------
            <#
            $pltRvEnv=[ordered]@{
                PSCmdletproxy = $rPSCmdlet ;
                PSScriptRootproxy = $rPSScriptRoot ;
                PSCommandPathproxy = $rPSCommandPath ;
                MyInvocationproxy = $rMyInvocation ;
                PSBoundParametersproxy = $rPSBoundParameters
                verbose = [boolean]($PSBoundParameters['Verbose'] -eq $true) ;
            } ;
            write-verbose "(Purge no value keys from splat)" ;
            $mts = $pltRVEnv.GetEnumerator() |?{$_.value -eq $null} ; $mts |%{$pltRVEnv.remove($_.Name)} ; rv mts -ea 0 -whatif:$false -confirm:$false;
            $smsg = "resolve-EnvironmentTDO w`n$(($pltRVEnv|out-string).trim())" ;
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
            $rvEnv = resolve-EnvironmentTDO @pltRVEnv ;
            $smsg = "`$rvEnv returned:`n$(($rvEnv |out-string).trim())" ;
            if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
            #>
            #endregion RV_ENVIRO ; #*------^ END RV_ENVIRO ^------
            #region NETWORK_INFO ; #*======v NETWORK_INFO v======
            #$NetSummary = resolve-NetworkLocalTDO ;
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
            #region TEST_EXOPLOCAL ; #*------v TEST_EXOPLOCAL v------
            #
            #$XoPSummary = test-LocalExchangeInfoTDO ;
            write-verbose "Expand returned NoteProperty properties into matching local variables" ;
            if($host.version.major -gt 2){
                $XoPSummary.PsObject.Properties | ?{$_.membertype -eq 'NoteProperty'} | foreach-object{set-variable -name $_.name -value $_.value -verbose -whatif:$false -Confirm:$false ;} ;
            }else{
                write-verbose "Psv2 lacks the above expansion capability; just create simpler variable set" ;
                $ExVers = $XoPSummary.ExVers ; $isLocalExchangeServer = $XoPSummary.isLocalExchangeServer ; $IsEdgeTransport = $XoPSummary.IsEdgeTransport ;
            } ;
            #endregion TEST_EXOPLOCAL ; #*------^ END TEST_EXOPLOCAL ^------
            #

            #endregion ENVIRO_DISCOVER ; #*------^ END ENVIRO_DISCOVER ^------
            #region TLS_LATEST_FORCE ; #*------v TLS_LATEST_FORCE v------
            $CurrentVersionTlsLabel = [Net.ServicePointManager]::SecurityProtocol ; # Tls, Tls11, Tls12 ('Tls' == TLS1.0)  ;
            write-verbose "PRE: `$CurrentVersionTlsLabel : $($CurrentVersionTlsLabel )" ;
            # psv6+ already covers, test via the SslProtocol parameter presense
            if ('SslProtocol' -notin (Get-Command Invoke-RestMethod).Parameters.Keys) {
                $currentMaxTlsValue = [Math]::Max([Net.ServicePointManager]::SecurityProtocol.value__,[Net.SecurityProtocolType]::Tls.value__) ;
                write-verbose "`$currentMaxTlsValue : $($currentMaxTlsValue )" ;
                $newerTlsTypeEnums = [enum]::GetValues('Net.SecurityProtocolType') | Where-Object { $_ -gt $currentMaxTlsValue }
                if($newerTlsTypeEnums){
                    write-verbose "Appending upgraded/missing TLS `$enums:`n$(($newerTlsTypeEnums -join ','|out-string).trim())" ;
                } else {
                    write-verbose "Current TLS `$enums are up to date with max rev available on this machine" ;
                };
                $newerTlsTypeEnums | ForEach-Object {
                    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor $_
                } ;
            } ;
            #endregion TLS_LATEST_FORCE ; #*------^ END TLS_LATEST_FORCE ^------

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
            if($useEXO){$tDepModules += @("ExchangeOnlineManagement;https://www.powershellgallery.com/packages/ExchangeOnlineManagement/")} ; 
            if($UseMSOL){$tDepModules += @("MSOnline;https://www.powershellgallery.com/packages/MSOnline/")} ; 
            if($UseAAD){$tDepModules += @("AzureAD;https://www.powershellgallery.com/packages/AzureAD/")} ; 
            if($useEXO){$tDepModules += @("ExchangeOnlineManagement;https://www.powershellgallery.com/packages/ExchangeOnlineManagement/")} ; 
            if($UseMG){$tDepModules += @("Microsoft.Graph.Authentication;https://www.powershellgallery.com/packages/Microsoft.Graph/")} ; 
            if($UseOPAD){$tDepModules += @("ActiveDirectory;get-windowscapability -name RSAT* -Online | ?{$_.name -match 'Rsat\.ActiveDirectory'} | %{Add-WindowsCapability -online -name $_.name}")} ; 

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
                foreach($tmod in $tDepModules){
                    $tmodName,$tmodURL = $tmod.split(';') ;
                    if (-not(Get-Module $tmodName -ListAvailable)) {
                        $smsg = "This script requires a recent version of the $($tmodName) PowerShell module. Download it here:`n$($tmodURL )";
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Warn }
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                        return
                    } else {
                        write-verbose "$tModName confirmed available" ;
                    } ;
                } ;
            } ;
            #endregion TEST_MODS ; #*------^ END TEST_MODS ^------

            
            # return status obj
            <#
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
            if($host.version.major -ge 3){$ret_ccO365S=[ordered]@{Dummy = $null ;} }
            else {$ret_ccO365S = @{Dummy = $null ;} } ;
            if($ret_ccO365S.keys -contains 'dummy'){$ret_ccO365S.remove('Dummy') };
            $fieldsBoolean = 'hasEXO','hasMSOL','hasAAD','hasMG' | select -unique  | sort ; $fieldsBoolean | % { $ret_ccO365S.add($_,$false) } ;
            $fieldsnull = 'CredentialO365','MGContext','MGtoken' | select -unique  | sort ; $fieldsnull | % { $ret_ccO365S.add($_,$null) } ;


            # PRETUNE STEERING separately *before* pasting in balance of region
            # THIS BLOCK DEPS ON VERB-* FANCY CRED/AUTH HANDLING MODULES THAT *MUST* BE INSTALLED LOCALLY TO FUNCTION
            # NOTE: *DOES* INCLUDE *PARTIAL* DEP-LESS $useExopNoDep=$true OPT THAT LEVERAGES Connect-ExchangeServerTDO, VS connect-ex2010 & CREDS ARE ASSUMED INHERENT TO THE ACCOUNT)
            # Connect-ExchangeServerTDO HAS SUBSTANTIAL BENEFIT, OF WORKING SEAMLESSLY ON EDGE SERVER AND RANGE OF DOMAIN-=CONNECTED EXOP ROLES
            <#
            $useO365 = $true ;
            $useEXO = $true ;
            $UseOP=$true ;
            $UseExOP=$true ;
            $useExopNoDep = $true ; # switch to use Connect-ExchangeServerTDO, vs connect-ex2010 (creds are assumed inherent to the account)
            $ExopVers = 'Ex2010' # 'Ex2019','Ex2016','Ex2013','Ex2010','Ex2007','Ex2003','Ex2000', Null for All versions
            if($Version){
                $ExopVers = $Version ; #defer to local script $version if set
            } ;
            $useForestWide = $true ; # flag to trigger cross-domain/forest-wide code in AD & EXoP
            $UseOPAD = $true ;
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
                $UseExOP = $true ;
                if($IsEdgeTransport -AND $psise){
                    $smsg = "powershell_ISE UNDER Exchange Edge Transport role!"
                    $smsg += "`nThis script is likely to fail the get-messagetrackingLog calls with Access Denied errors"
                    $smsg += "`nif run with this combo."
                    $smsg += "`nEXIT POWERSHELL ISE, AND RUN THIS DIRECTLY UNDER EMS FOR EDGE USE";
                    $smsg += "`n(bug appears to be a conflict in Remote EMS v EMS access permissions, not resolved yet)" ;
                    write-warning $msgs ;
                } ;
            } ;
            $useO365 = [boolean]($useO365 -OR $useEXO -OR $UseMSOL -OR $UseAAD -OR $UseMG)
            $UseOP = [boolean]($UseOP -OR $UseExOP -OR $UseOPAD) ;
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
                #region GENERIC_EXO_CREDS_&_SVC_CONN #*------v GENERIC EXO CREDS & SVC CONN BP v------
                # o365/EXO creds
                <### Usage: Type defaults to SID, if not spec'd - Note: there must be a *logged in & configured *profile*
                $o365Cred = get-TenantCredentials -TenOrg $TenOrg -verbose -userrole SID ;
                Returns a credential set for the $TenOrg Hybrid OnPrem Exchange Org
                .EXAMPLE
                $o365Cred = get-TenantCredentials -TenOrg $TenOrg -verbose -userrole CSVC ;
                Returns the CSVC Userrole credential for the $TenOrg Hybrid OnPrem Exchange Org
                .EXAMPLE
                $o365Cred = get-TenantCredentials -TenOrg $TenOrg -verbose -userrole B2BI ;
                Returns the B2BI Userrole credential for the $TenOrg Hybrid OnPrem Exchange Org
                ###>
                $o365Cred = $null ;
                if($Credential){
                    $smsg = "`Credential:Explicit credentials specified, deferring to use..." ;
                    if($silent){}elseif ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
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
                } elseif(get-item function:get-TenantCredentials -ea stop){
                    $pltGTCred=@{TenOrg=$TenOrg ; UserRole=$null; verbose=$($verbose)} ;
                    if($UserRole){
                        $smsg = "(`$UserRole specified:$($UserRole -join ','))" ;
                        if($silent){}elseif ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        $pltGTCred.UserRole = $UserRole;
                    } else {
                        switch -Regex ($TenOrg){
                            'TOL' {
                                [string[]]$UserRole = @('SIDCBA','SID','ESvcCBA')
                            }
                            'CMW|TOR' {
                                [string[]]$UserRole = @('ESvcCBA','SID','CSVC')
                                # svcAcct use: @('ESvcCBA','CSvcCBA','SIDCBA')
                            }
                            default {
                                [string[]]$UserRole = @('SID','CSVC')
                            }
                        } ;
                        $smsg = "(No -UserRole specified, defaulting to:SIDCBA,SID,ESvcCBA )" ;
                        if($silent){}elseif ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                        $pltGTCred.UserRole = $UserRole ;
                    } ;
                    $smsg = "get-TenantCredentials w`n$(($pltGTCred|out-string).trim())" ;
                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level verbose }
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                    $o365Cred = get-TenantCredentials @pltGTCred
                } else{
                    if(get-variable "$($Tenorg)Meta" -ea 0){
                        if((get-variable "$($Tenorg)Meta" -ea 0).value.o365_SIDUpn){
                            #$o365Cred.Cred = Get-Credential -Credential ((get-variable "$($Tenorg)Meta" -ea 0).value.o365_SIDUp) ;
                            # wo basicAuth, no pw exchanges, it's MFA only now, only need the logon
                            $o365Cred.Cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList (((get-variable "$($Tenorg)Meta" -ea 0).value.o365_SIDUp),(convertto-securestring -string 'dummy' -asplaintext -force)) ;
                            $o365Cred.credtype = 'SID' ;
                        }else{
                            $smsg = "No resolvable *Meta.o365_SIDUpn: prompting for SID (use dummy pw, will be unused for o365 MAuth logon) " ;
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                            $o365Cred.Cred = Get-Credential -Credential  ;
                            $o365Cred.credtype = 'SID' ;
                        }
                    } else {
                        $smsg = "`$TenOrg:$($Tenorg) specified: No matching $($Tenorg)Meta variable found locally!"  ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        #throw $smsg ;
                        $smsg = "Prompting for SID (use dummy pw, will be unused for o365 MAuth logon) " ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        $o365Cred.Cred = Get-Credential -Credential  ;
                        $o365Cred.credtype = 'SID' ;

                    } ;
                } ;
                if($o365Cred.credType -AND $o365Cred.Cred -AND $o365Cred.Cred.gettype().fullname -eq 'System.Management.Automation.PSCredential'){
                    $smsg = "(validated `$o365Cred contains .credType:$($o365Cred.credType) & `$o365Cred.Cred.username:$($o365Cred.Cred.username)" ;
                    if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                    else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                    # 9:58 AM 6/13/2024 populate $credential with return, if not populated (may be required for follow-on calls that pass common $Credentials through)
                    if( ((get-variable Credential -ea 0) -AND $null -eq $Credential) -OR -not (get-variable Credential -ea 0)  ){
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
                if($o365Cred){
                    # make it script scope, so we don't have to predetect & purge before using new-variable
                    if(get-Variable -Name cred$($tenorg) -scope Script -ea 0 ){ remove-Variable -Name cred$($tenorg) -scope Script } ;
                    New-Variable -Name cred$($tenorg) -scope Script -Value $o365Cred.cred -whatif:$false -confirm:$false;
                    $smsg = "Resolved $($Tenorg) `$o365cred:$($o365Cred.cred.username) (assigned to `$cred$($tenorg))" ;
                    if($silent){}elseif ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                } else {
                    $statusdelta = ";ERROR"; # CHANGE|INCOMPLETE|ERROR|WARN|FAIL ;
                    $script:PassStatus += $statusdelta ;
                    set-Variable -Name PassStatus_$($tenorg) -scope Script -Value ((get-Variable -Name PassStatettus_$($tenorg)).value + $statusdelta) ;
                    $smsg = "Unable to resolve $($tenorg) `$o365Cred value!"
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    throw "Unable to resolve $($tenorg) `$o365Cred value!`nEXIT!"
                    Break ;
                } ;
                # if we get here, wo a $Credential, w resolved $o365Cred, assign it
                # we've got both pltRXO & pltRXOC: pltRXOC is used in connect-exo; connect-msol & connect-aad; pltrxo isn't.
                if(-not $Credential -AND $o365Cred){$Credential = $o365Cred.cred } ;
                # configure splat for connections: (see above useage)
                # downstream commands
                $pltRXO = [ordered]@{
                    Credential = $Credential ;
                    verbose = $($VerbosePreference -eq "Continue")  ;
                } ;
                if($silent -AND ((get-command Reconnect-EXO).Parameters.keys -contains 'silent')){
                    $pltRxo.add('Silent',[boolean]$silent) ;
                } ;
                # default connectivity cmds - force silent
                $pltRXOC = [ordered]@{} ; $pltRXO.GetEnumerator() | ?{ $_.Key -notmatch 'silent' }  | ForEach-Object { $pltRXOC.Add($_.Key, $_.Value) } ; $pltRXOC.Add('silent',[boolean]$silent) ;
                if((get-command Reconnect-EXO).Parameters.keys -notcontains 'silent'){
                    $pltRxo.remove('Silent') ;
                } ;
                write-verbose "store the `$credential" ; 
                $ret_ccO365S.CredentialO365 = $pltRXOC.Credential ; 

                #region EOMREV ; #*------v EOMREV Check v------
                #$EOMmodname = 'ExchangeOnlineManagement' ;
                $pltIMod = @{Name = $EOMmodname ; ErrorAction = 'Stop' ; verbose=$false} ;
                # do a gmo first, faster than gmo -list
                if([version]$EOMMv = (Get-Module @pltIMod).version){}
                elseif([version]$EOMMv = (Get-Module -ListAvailable @pltIMod).version){}
                else {
                    $smsg = "$($EOMmodname) PowerShell v$($MinNoWinRMVersion) module is required, do you want to install it?" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Prompt }
                    else{ $smsg = "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    $bRet = Read-Host "Enter YYY to continue. Anything else will exit"  ;
                    if ($bRet.ToUpper() -eq "YYY") {
                        $smsg = "Installing $($EOMmodname) module..." ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        Install-Module $EOMmodname -Repository PSGallery -AllowClobber -Force ;
                    } else {
                        $smsg = "Please install $($EOMmodname) PowerShell v$($MinNoWinRMVersion)  module." ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN }
                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        #exit 1
                        break ;
                    }  ;
                } ;
                $smsg = "(Checking for WinRM support in this EOM rev...)" ;
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE }
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
                if([version]$EOMMv -ge [version]$MinNoWinRMVersion){
                    $MinNoWinRMVersion = $EOMMv.tostring() ;
                    $IsNoWinRM = $true ;
                }elseif([version]$EOMMv -lt [version]$MinimumVersion){
                    $smsg = "Installed $($EOMmodname) is v$($MinNoWinRMVersion): This module is obsolete!" ;
                    $smsg += "`nAnd unsupported by this function!" ;
                    $smsg += "`nPlease install $($EOMmodname) PowerShell v$($MinNoWinRMVersion)  module!" ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    Break ;
                } else {
                    $IsNoWinRM = $false ;
                } ;
                [boolean]$UseConnEXO = [boolean]([version]$EOMMv -ge [version]$MinNoWinRMVersion) ;
                #endregion EOMREV ; #*------^ END EOMREV Check  ^------
                #-=-=-=-=-=-=-=-=
                <### CALLS ARE IN FORM: (cred$($tenorg))
                # downstream commands
                $pltRXO = @{
                    Credential = (Get-Variable -name cred$($tenorg) ).value ;
                    #verbose = $($verbose) ;
                    Verbose = $FALSE ;
                } ;
                if((get-command Reconnect-EXO).Parameters.keys -contains 'silent'){
                    $pltRxo.add('Silent',$false) ;
                } ;
                # default connectivity cmds - force silent false
                $pltRXOC = [ordered]@{} ; $pltRXO.GetEnumerator() | ?{ $_.Key -notmatch 'silent' }  | ForEach-Object { $pltRXOC.Add($_.Key, $_.Value) } ; $pltRXOC.Add('silent',$true) ;
                if((get-command Reconnect-EXO).Parameters.keys -notcontains 'silent'){
                    $pltRxo.remove('Silent') ;
                } ;
                #$pltRXO creds & .username can also be used for AzureAD connections:
                #Connect-AAD @pltRXOC ;
                ###>
                #endregion GENERIC_EXO_CREDS_&_SVC_CONN #*------^ END GENERIC EXO CREDS & SVC CONN BP ^------

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
                if ($script:useEXOv2 -OR $useEXOv2) { reconnect-eXO2 @pltRXOC }
                else { reconnect-EXO @pltRXOC } ;
                $ret_rxo = test-exoconnectiontdo ;
                if($ret_rxo.connection -ANd $ret_rxo.isXO -ANd $ret_rxo.isValid -AND $ret_rxo.TokenLifeMins -gt 0){$ret_ccO365S.hasEXO = $true}else {$ret_ccO365S.hasEXO = $false } ;
            } else {
                $smsg = "(`$useEXO:$($useEXO))" ;
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ;
            } ; # if-E
            #endregion  ; #*------^ END useEXO ^------
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
                <# Find-MgGraphCommand -command Get-MgUser | Select -First 1 -ExpandProperty Permissions

                $Cmdlets = 'Get-MgUser','Get-MgSubscribedSku';
                $prpMgu = 'BusinessPhones','DisplayName','GivenName','JobTitle','Mail','MobilePhone','OfficeLocation','Surname','UserPrincipalName' ;
                $PermsRqd = @() ; $Cmdlets |%{$PermsRqd += Find-MgGraphCommand -command $_ -ea STOP| Select -First 1 -ExpandProperty Permissions | Select -Unique name ; } ; $PermsRqd = $PermsRqd.name | select -unique ;
                $smsg = "Connect-mgGraph -scope`n`n$(($PermsRqd|out-string).trim())" ;
                $smsg += "`n`n(Perms reflects Cmdlets:$($Cmdlets -join ','))" ;
                write-host $smsg ;
                Connect-mgGraph -scope $PermsRqd -ea STOP ;

                $prpMgu = 'BusinessPhones','DisplayName','GivenName','JobTitle','Mail','MobilePhone','OfficeLocation','Surname','UserPrincipalName' ;
                #>
                if(-not $MGPermissionsScope){
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
                        $PermsRqd = get-MGCodeCmdletPermissionsTDO @pltGMGP ;
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
                $pltCCMG=[ordered]@{
                    ErrorAction = 'SilentlyContinue' ;ErrorVariable = 'err_ccMG'
                } ;
                if($PermsRqd){ $pltCCMG.add('scope',$PermsRqd)} ;
                $smsg = "Connect-mgGraph w`n$(($pltCCMG|out-string).trim())" ;
                if($pltCCMG.scope){
                    $smsg += "`nwith -scope`n`n$(($PermsRqd|out-string).trim())" ;
                } ;
                if($MGCmdlets){
                    $smsg += "`n`n(Perms reflects Cmdlets:$($MGCmdlets  -join ','))" ;
                } ;
                if($silent){} else {
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Prompt }
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                } ;
                $ccResults = Connect-mgGraph @pltCCMG ;
                if($err_ccMG -ne $null){
                    $smsg = $err_ccMG  ;
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent}
                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    throw $smsg ;
                    BREAK ;
                } ;
                $smsg = "Connect-MgGraph result`n$(($ccresults|out-string).trim())" ;   ;
                #TRY{$MGContext = Get-MgContext -ea stop ; $ret_ccO365S.hasMG = $true} CATCH {$ret_ccO365S.hasMG = $false }  ; 
                TRY{
                    if($ctxMG = [Microsoft.Graph.PowerShell.Authentication.GraphSession]::Instance.AuthContext){
                    #if($ctxMG = Get-MgContext){
                        $ret_ccO365S.hasMG = $true ; 
                        if($ctxMG.AuthType -eq 'Delegated'){
                            $smsg = "`n$(($ctxMG | ft -a $prpMGConnDeleg|out-string).trim())" ; 
                        } else { 
                            if($ctxMG.CertificateThumbprint){
                                $smsg = "`n$(($ctxMG | ft -a ($prpMGConnDeleg | select -Skip 1)|out-string).trim())" ;
                                $smsg += "`n$(($ctxMG | ft -a $prpMGConnCBA|out-string).trim())" ;
                            } else { 
                                $smsg = "`n$(($ctxMG | ft -a $prpMGConnDeleg|out-string).trim())" ; 
                            }
                        } ;
                        $smsg += "`n$(($ctxMG |select @{name="Scopes";expression={$_.Scopes -join ","}}|out-string).trim())" ;
                        if($silent){}elseif ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Success } 
                        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                        #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                        #return $true ;
                        #return ($ctxMG | select $prpMGConnRet) ; 
                        $ret_ccO365S.MGContext = ($ctxMG | select $prpMGConnRet) ; 
                    } ; 
                }CATCH{
                    write-host -foregroundcolor yellow  "No MG Connection!" ; 
                    #return $false  ;
                } ; 
                <#$smsg += "`nMGContext:`n$((Get-MgContext|out-string).trim())" ;
                #$smsg += "`nMGContext:`n$(($MGContext|out-string).trim())" ;
                if($silent){} else {
                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
                    else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    #Levels:Error|Warn|Info|H1|H2|H3|H4|H5|Debug|Verbose|Prompt|Success
                } ;
                #>
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
            if ($script:useEXOv2) { reconnect-eXO2 @pltRXOC }
            else { reconnect-EXO @pltRXOC } ;
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
} ;
#endregion CONNECT_O365SERVICES ; #*======^ END CONNECT_O365SERVICES ^======

