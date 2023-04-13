# test-EXOIsLicensed

#*----------v Function test-EXOIsLicensed() v----------
function test-EXOIsLicensed {
    <#
    .SYNOPSIS
    test-EXOIsLicensed.ps1 - Evaluate IsLicensed status, to indicate license support for Exchange online UserMailbox type, on passed in AzureADUser object
    .NOTES
    Version     : 1.0.0
    Author      : Todd Kadrie
    Website     :	http://www.toddomation.com
    Twitter     :	@tostka / http://twitter.com/tostka
    CreatedDate : 2022-03-22
    FileName    : 
    License     : MIT License
    Copyright   : (c) 2022 Todd Kadrie
    Github      : https://github.com/tostka/verb-XXX
    Tags        : Powershell
    REVISIONS
    * 1:06 PM 4/4/2022 updated CBH example to reflect $AADU obj, not UPN input
    3:08 PM 3/23/2022 init
    .DESCRIPTION
    test-EXOIsLicensed.ps1 - Evaluate IsLicensed status, to indicate license support for Exchange online UserMailbox type, on passed in AzureADUser object
    Coordinates with verb-exo:get-ExoMailboxLicenses() to retrieve a static list of UserMailbox -supporting license names & sku's in our Tenant. 

    The get-EXOMailboxLicenses list is *not* interactive with AzureAD or EXO, 
    and it *will* have to be tuned for local Tenants, and maintained for currency over time. 

    It's a simple test, but it beats..
        ...the prior get-Msoluser |?{$_.islicensed} (which indicates:*some* license is assigned - could be a worthless 'FREEFLOW'!) 
        
        ... or testing |?{$_.LicenseReconciliationNeeded } 
        ( which used to indicate a mailbox *exists* but lacks a suitable mailbox-supporting license, 
        and continues to be mounted, *solely* due to being within 30days of onboarding to cloud.).  

    Not to mention get-AzureADuser's complete lack of any native evaluation on either front. [facepalm]
    Nor any similar native support in the gap from the ExchangeOnlineManagement module. 

    <rant>
        I *love* coding coverage for slipshod MS module providers that write to replace *force*-deprecated critical infra tools, 
        but can't be bothered to deliver equiv function, equiv parameters, or even similar outputs, 
        for long-standing higher-functioning tools, when they write the half-implemented *new* ones.

        And no, "Just make calls to GraphAPI!", is not a viable answer, for *working* admins, mandated to deliver working solutions on tight schedules. 
        If we wanted to be REST web devs, we wouldn't be running o365 services!
    </rant>

    .PARAMETER  User
    AzureADUser [Microsoft.Open.AzureAD.Model.User] object
    .PARAMETER TenOrg
    TenantTag value, indicating Tenants to connect to[-TenOrg 'TOL']
    .PARAMETER Credential
    Credential to be used for connection
    .OUTPUT
    System.Boolean
    .EXAMPLE
    PS> $isEXOLicensed = test-EXOIsLicensed -User $AADUser -verbose
    PS> if($isEXOLicensed){write-host 'Has EXO Usermailbox Type License'} else { write-warning 'NO EXO USERMAILBOX TYPE LICENSE!'} ; 
    Evaluate IsLicensed status on passed UPN object
    .LINK
    https://github.com/tostka/verb-EXO
    #>
    #Requires -Version 3
    ##Requires -Modules AzureAD, verb-Text
    ##Requires -RunasAdministrator
    # VALIDATORS: [ValidateNotNull()][ValidateNotNullOrEmpty()][ValidateLength(24,25)][ValidateLength(5)][ValidatePattern("some\sregex\sexpr")][ValidateSet("USEA","GBMK","AUSYD")][ValidateScript({Test-Path $_ -PathType 'Container'})][ValidateScript({Test-Path $_})][ValidateRange(21,65)][ValidateCount(1,3)]
    [CmdletBinding()]
    
     Param(
        [Parameter(Position=0,Mandatory=$True,ValueFromPipeline=$true,HelpMessage="Either Msoluser object or UserPrincipalName for user[-User upn@domain.com|`$msoluserobj ]")]
        [Microsoft.Open.AzureAD.Model.User]$User,
        [Parameter(Mandatory=$False,HelpMessage="TenantTag value, indicating Tenants to connect to[-TenOrg 'TOL']")]
        [ValidateNotNullOrEmpty()]
        [ValidatePattern("^\w{3}$")]
        [string]$TenOrg = $global:o365_TenOrgDefault,
        [Parameter(Mandatory=$False,HelpMessage="Credentials [-Credentials [credential object]]")]
        [System.Management.Automation.PSCredential]$Credential = $global:credo365TORSID
        #[switch]$silent # removed, there's no echos enabled
    )
    BEGIN {
        ${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name ;
        $Verbose = ($VerbosePreference -eq 'Continue') ;
        #Connect-AAD -Credential:$Credential -verbose:$($verbose) ;
        
        $ExMbxLicenses = get-ExoMailboxLicenses -verbose:$($VerbosePreference -eq "Continue") ;
        # pull the full Tenant list, for performing sku-> name conversions
        #$lplist =  get-AADlicensePlanList -verbose -IndexOnName ;

        $pltGLPList=[ordered]@{ TenOrg= $TenOrg; credential= $Credential ; IndexOnName=$false ; verbose=$($VerbosePreference -eq "Continue") ;} ; 
        $smsg = "get-AADlicensePlanList w`n$(($pltGLPList|out-string).trim())" ; 
        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 

        $skus  = get-AADlicensePlanList @pltGLPList ;

        # check if using Pipeline input or explicit params:
        if ($PSCmdlet.MyInvocation.ExpectingInput) {
            write-verbose "Data received from pipeline input: '$($InputObject)'" ;
        } else {
            # doesn't actually return an obj in the echo
            write-verbose "Data received from parameter input:" # '$($InputObject)'" ;
        } ;
    } 
    PROCESS {
        if($ExMbxLicenses){
            $IsExoLicensed = $false ;
            foreach($pLic in $User.AssignedLicenses.skuid){
                $smsg = "(resolving $($plic) for EXO support)" ; 
                if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                if($tsku = $skus[$pLic]){
                    if($xlic = $ExMbxLicenses[$tsku.SkuPartNumber]){
                        $IsExoLicensed = $true ;
                        $smsg = "$($User.userprincipalname) HAS EXO UserMailbox-supporting License:$($xlic.SKU)|$($xlic.Label)|$($tsku.skuid)" ;
                        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } 
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                        break ; 
                    } ;
                } else { 

                } ; 
            } ;
        } else { 
            $smsg = "Unable to resolve get-ExoMailboxLicenses!" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } 
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
            throw $smsg ;
            Break ; 
        } ; 

    }  # PROC-E
    END{
        $IsExoLicensed | write-output ; 
    } ;
} ; 
#*------^ END Function test-EXOIsLicensed() ^------