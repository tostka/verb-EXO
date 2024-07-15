#*------v RemoveExistingPSSessionTargeted.ps1 v------
function RemoveExistingPSSessionTargeted() {
    <#
    .SYNOPSIS
    RemoveExistingPSSessionTargeted.ps1 - Tweaked version of the Exchangeonline module:RemoveExistingPSSession() to avoid purging CCMW sessions on connect. Intent is to permit *concurrent* EXO & CCMS sessions.
    .NOTES
    Version     : 1.0.0
    Author      : Todd Kadrie
    Website     :	http://www.toddomation.com
    Twitter     :	@tostka / http://twitter.com/tostka
    CreatedDate : 20201109-0833AM
    FileName    : RemoveExistingPSSessionTargeted.ps1
    License     : [none specified]
    Copyright   : [none specified]
    Github      : https://github.com/tostka/verb-exo
    Tags        : Powershell
    AddedCredit : Microsoft (edited version of published commands in the module)
    AddedWebsite:	https://docs.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2
    REVISIONS
    * 12:01 PM 7/15/2024 long obso pssession target func, delete
    * 8:34 AM 11/9/2020 init
    * 8:34 AM 11/9/2020 init
    .DESCRIPTION
    RemoveExistingPSSessionTargeted.ps1 - Tweaked version of the Exchangeonline module:RemoveExistingPSSession() to avoid purging CCMW sessions on connect. Intent is to permit *concurrent* EXO & CCMS sessions.
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    None. Returns no objects or output.
    .EXAMPLE
    RemoveExistingPSSessionTargeted
    Stock call
    .LINK
    https://github.com/tostka/verb-EXO
    .LINK
    https://docs.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2
    #>
    [CmdletBinding()]
    param()
    
    #$existingPSSession = Get-PSSession | Where-Object {$_.ConfigurationName -like "Microsoft.Exchange" -and $_.Name -like "ExchangeOnlineInternalSession*"}
    <# filter *ONLY* EXO sessions, exclude CCMS, they differ on ComputerName endpoint:
    #-=EXO-=-=-=-=-=-=-=
    ConfigurationName : Microsoft.Exchange
    ComputerName      : outlook.office365.com
    Name              : ExchangeOnlineInternalSession_2
    #-=CCMS-=-=-=-=-=-=-=
    ConfigurationName : Microsoft.Exchange
    ComputerName      : nam02b.ps.compliance.protection.outlook.com
    Name              : ExchangeOnlineInternalSession_1
    #-=-=-=-=-=-=-=-=
    #>
    $rgxExoPsHostName = "^(ps\.outlook\.com|outlook\.office365\.com)$"
    $existingPSSession = Get-PSSession | Where-Object {$_.ConfigurationName -like "Microsoft.Exchange" -and $_.Name -like "ExchangeOnlineInternalSession*" -AND $_.ComputerName -match$rgxExoPsHostName} ; 

        if ($existingPSSession.count -gt 0) 
        {
            for ($index = 0; $index -lt $existingPSSession.count; $index++)
            {
                $session = $existingPSSession[$index]
                Remove-PSSession -session $session

                Write-Host "Removed the PSSession $($session.Name) connected to $($session.ComputerName)"
            }
        }

        # Clear any left over PS tmp modules
        if ($global:_EXO_PreviousModuleName -ne $null)
        {
            Remove-Module -Name $global:_EXO_PreviousModuleName -ErrorAction SilentlyContinue
            $global:_EXO_PreviousModuleName = $null
        }
    }

#*------^ RemoveExistingPSSessionTargeted.ps1 ^------