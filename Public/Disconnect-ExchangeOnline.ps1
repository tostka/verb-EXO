#*------v Disconnect-ExchangeOnline.ps1 v------
function Disconnect-ExchangeOnline{
    <#
    .SYNOPSIS
    Disconnect-ExchangeOnline.ps1 - localized verb-EXO vers of non-'$global:' funct from ExchangeOnlineManagement. The globals export fine, these don't and appear to need to be loaded manually
    .NOTES
    Version     : 1.0.0
    Author      : Todd Kadrie
    Website     :	http://www.toddomation.com
    Twitter     :	@tostka / http://twitter.com/tostka
    CreatedDate : 20201109-0833AM
    FileName    : Disconnect-ExchangeOnline.ps1
    License     : [none specified]
    Copyright   : [none specified]
    Github      : https://github.com/tostka/verb-exo
    Tags        : Powershell
    AddedCredit : Microsoft (edited version of published commands in the module)
    AddedWebsite:	https://docs.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2
    REVISIONS
    * 8:34 AM 11/9/2020 init
    .DESCRIPTION
    Disconnect-ExchangeOnline.ps1 - localized verb-EXO vers of non-'$global:' funct from ExchangeOnlineManagement. The globals export fine, these don't and appear to need to be loaded manually
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    None. Returns no objects or output.
    .EXAMPLE
    Disconnect-ExchangeOnline
    Stock call
    .LINK
    https://github.com/tostka/verb-EXO
    .LINK
    https://docs.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact='High')]
    param()

    process {
        if ($PSCmdlet.ShouldProcess(
            "Running this cmdlet clears all active sessions created using Connect-ExchangeOnline or Connect-IPPSSession.",
            "Press(Y/y/A/a) if you want to continue.",
            "Running this cmdlet clears all active sessions created using Connect-ExchangeOnline or Connect-IPPSSession. "))
        {

            # Keep track of error count at beginning.
            $errorCountAtStart = $global:Error.Count;

            try
            {
                # Cleanup current exchange online PSSessions
                #RemoveExistingPSSession
                RemoveExistingPSSessionTargeted

                # Import the module once more to ensure that Test-ActiveToken is present
                Import-Module $global:_EXO_ModulePath -Cmdlet Clear-ActiveToken;

                # Remove any active access token from the cache
                Clear-ActiveToken

                Write-Host "Disconnected successfully !"

                if ($global:_EXO_EnableErrorReporting -eq $true)
                {
                    if ($global:_EXO_TelemetryFilePath -eq $null)
                    {
                        $global:_EXO_TelemetryFilePath = New-EXOClientTelemetryFilePath
                    }

                    Push-EXOTelemetryRecord -TelemetryFilePath $global:_EXO_TelemetryFilePath -CommandName Disconnect-ExchangeOnline -CommandParams $PSCmdlet.MyInvocation.BoundParameters -OrganizationName  $global:_EXO_ExPSTelemetryOrganization -ScriptName $global:_EXO_ExPSTelemetryScriptName  -ScriptExecutionGuid $global:_EXO_ExPSTelemetryScriptExecutionGuid
                }
            }
            catch
            {
                # If telemetry is enabled, log errors generated from this cmdlet also. 
                if ($global:_EXO_EnableErrorReporting -eq $true)
                {
                    $errorCountAtProcessEnd = $global:Error.Count 

                    if ($global:_EXO_TelemetryFilePath -eq $null)
                    {
                        $global:_EXO_TelemetryFilePath = New-EXOClientTelemetryFilePath
                    }

                    # Log errors which are encountered during Disconnect-ExchangeOnline execution. 
                    Write-Warning("Writing Disconnect-ExchangeOnline errors to " + $global:_EXO_TelemetryFilePath)

                    Push-EXOTelemetryRecord -TelemetryFilePath $global:_EXO_TelemetryFilePath -CommandName Disconnect-ExchangeOnline -CommandParams $PSCmdlet.MyInvocation.BoundParameters -OrganizationName  $global:_EXO_ExPSTelemetryOrganization -ScriptName $global:_EXO_ExPSTelemetryScriptName  -ScriptExecutionGuid $global:_EXO_ExPSTelemetryScriptExecutionGuid -ErrorObject $global:Error -ErrorRecordsToConsider ($errorCountAtProcessEnd - $errorCountAtStart) 
                }

                throw $_
            }
        }
    }
}

#*------^ Disconnect-ExchangeOnline.ps1 ^------