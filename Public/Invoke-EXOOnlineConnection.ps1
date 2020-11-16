#*------v Function Invoke-ExoOnlineConnection v------
function Invoke-ExoOnlineConnection{
    <#
    .SYNOPSIS
    Invoke-ExoOnlineConnection.ps1 - EXO non-ending MFA session, that renews it self ; once you connect to EXO with this it will stay open
    .NOTES
    Version     : 1.0.0
    Author      : Todd Kadrie
    Website     : http://www.toddomation.com
    Twitter     : @tostka / http://twitter.com/tostka
    CreatedDate : 2020-11-10
    FileName    : Invoke-ExoOnlineConnection.ps1
    License     : [none specified]
    Copyright   : [none specified]
    Github      : https://github.com/tostka/verb-XXX
    Tags        : Powershell
    AddedCredit : Mahmoud Badran
    AddedWebsite: https://techcommunity.microsoft.com/t5/exchange/60-minutes-timeout-on-mfa-session/m-p/559224
    REVISIONS
    .DESCRIPTION
    Invoke-ExoOnlineConnection.ps1 - EXO non-ending MFA session, that renews it self ; once you connect to EXO with this it will stay open
    normally came as a .ps1 with a local function. Haven't tested, looks like it should work, trick is to preregister the timer/check interval outside of the function, prior to call.
    .PARAMETER  Checktimer
    Switch to trigger a timercheck. [-Checktimer]
    PARAMETERRepairPSSession
    Switch to trigger a session repair. [-RepairPSSession]
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    None. Returns no objects or output
    .EXAMPLE
    ## Create an Timer instance to trackand recheck status
    $timer = New-Object Timers.Timer
    ## Now setup the Timer instance to fire events
    $timer.Interval = 600000
    $timer.AutoReset = $true  # enable the event again after its been fired
    $timer.Enabled = $true
    ## register your event
    ## $args[0] Timer object
    ## $args[1] Elapsed event properties
    Register-ObjectEvent -InputObject $timer -EventName Elapsed -SourceIdentifier Repair  -Action {Invoke-ExoOnlineConnection -Checktimer}
    .EXAMPLE
    .LINK
    https://github.com/tostka/verb-EXO
    .LINK
    https://techcommunity.microsoft.com/t5/exchange/60-minutes-timeout-on-mfa-session/m-p/559224
    #>
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$false,HelpMessage = "Switch to trigger a timercheck. [-Checktimer]")]
        [switch]$Checktimer,
        [Parameter(mandatory=$false, valuefrompipeline=$false,HelpMessage = "Switch to trigger a session repair. [-RepairPSSession]")]
        [switch]$RepairPSSession,
        [Parameter(HelpMessage = "Credential to use for this connection [-credential [credential obj variable]")]
        [System.Management.Automation.PSCredential]$Credential = $global:credo365TORSID
    )
    BEGIN{
        if(!$Global:ErrorActionPreference){$Global:ErrorActionPreference = "Stop"} ; 
        if(!$Global:VerbosePreference){$Global:VerbosePreference = "Continue"} ; 
        #if(!$office365UserPrincipalName){$office365UserPrincipalName = "ADMIN@o365.com" } ; 
        if(!$PSExoPowershellModuleRoot){$PSExoPowershellModuleRoot = (Get-ChildItem -Path $($env:LOCALAPPDATA+"\Apps\2.0\") -Filter CreateExoPSSession.ps1 -Recurse -ErrorAction SilentlyContinue -Force | Select -Last 1).DirectoryName } ; 
        if(!$ExoPowershellModule){$ExoPowershellModule = "Microsoft.Exchange.Management.ExoPowershellModule.dll"} ; 
        if(!$ExoPowershellModulePath){$ExoPowershellModulePath = [System.IO.Path]::Combine($PSExoPowershellModuleRoot, $ExoPowershellModule)} ; 
        if(!(get-module $ExoPowershellModule.replace('.dll','') )){Import-Module $ExoPowershellModulePath} ; 
    }
    PROCESS{
        #determine if  PsSession is loaded in memory
        $ExosessionInfo = Get-PsSession
        #calculate session time style: $global:_EXO_ExchangeEnvironmentName = $ExchangeEnvironmentName;
        # MS uses these global name
        if ($global:_EXO_ExosessionStartTime){
             $global:_EXO_ExosessionTotalTime = ((Get-Date) - $global:_EXO_ExosessionStartTime)
        }
        #need to loop through each session a user might have opened previously
        foreach ($ExosessionItem in $ExosessionInfo){
            #check session timer to know if we need to break the connection in advance of a timeout. Break and make new after 40 minutes.
            if ($ExosessionItem.ComputerName.Contains("outlook.office365.com") -and $ExosessionItem.State -eq "Opened" -and $global:_EXO_ExosessionTotalTime.TotalSeconds -ge "2400"){
                Write-Verbose -Message "The PowerShell session has been running for $($global:_EXO_ExosessionTotalTime.TotalMinutes) minutes. We need to shut it down and create a new session due to the access token expiration at 60 minutes."
                $ExosessionItem | Remove-PSSession
                Start-Sleep -Seconds 3
                $strSessionFound = $false
                $global:_EXO_ExosessionTotalTime = $null #reset the timer
            } else { Write-Verbose -Message "The PowerShell session has been running for $($global:_EXO_ExosessionTotalTime.TotalMinutes) minutes.)"}
            #Force repair PSSession
            if ($ExosessionItem.ComputerName.Contains("outlook.office365.com") -and $RepairPSSession){
                Write-Verbose -Message "Attempting to repair broken PowerShell session to Exchange Online using cached credential."
                $ExosessionItem | Remove-PSSession
                Start-Sleep -Seconds 3
                $strSessionFound = $false
                $global:_EXO_ExosessionTotalTime = $null
            }elseif ($ExosessionItem.ComputerName.Contains("outlook.office365.com") -and $ExosessionItem.State -eq "Opened"){
                $strSessionFound = $true
            }
        }
        if (!$strSessionFound){
            Write-Verbose -Message "Creating new Exchange Online PowerShell session..."
            try{
                $pltNEXOS = @{
                    ExchangeEnvironmentName         = $ExchangeEnvironmentName ;
                    ConnectionUri                   = "https://outlook.office365.com/powershell-liveid/" ;
                    #AzureADAuthorizationEndpointUri = $AzureADAuthorizationEndpointUri ;
                    UserPrincipalName               = $Credential.username ;
                    PSSessionOption                 = $PSSessionOption ;
                    #Credential                      = $Credential ;
                    BypassMailboxAnchoring          = $($BypassMailboxAnchoring) ;
                    #ShowProgress                    = $($showProgress) # isn't a param of new-exopssessoin, is used with set-exo
                    #DelegatedOrg                    = $DelegatedOrganization ;
                    ErrorAction                      = 'SilentlyContinue' ; 
                    ErrorVariable                    = $newOnlineSessionError ; 
                }
                #$ExoSession  = New-ExoPSSession -UserPrincipalName $Credential.username -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -ErrorAction SilentlyContinue -ErrorVariable $newOnlineSessionError
                write-verbose "New-ExoPSSession w`n$(($pltNEXOS|out-string).trim())" ; 
                $ExoSession  = New-ExoPSSession @pltNEXOS ; 
            }catch{
                Write-Verbose -Message "Throw error..."
                throw;
            } finally {
                if ($newOnlineSessionError) {
                 Write-Verbose -Message "Final error..."
                    throw $newOnlineSessionError
                }
            }
            Write-Verbose -Message "Importing remote PowerShell session..."
            $global:_EXO_ExosessionStartTime = (Get-Date)
            #Import-PSSession $ExoSession -AllowClobber | Out-Null
            Import-PSSession $ExoSession -AllowClobber -DisableNameChecking
        } ;
    } ;
    END{} ;
} ; 
#*------^ END Function Invoke-ExoOnlineConnection ^------

