﻿#*------v get-ExoMailboxLicenses.ps1 v------
function get-ExoMailboxLicenses {
<#
    .SYNOPSIS
    get-ExoMailboxLicenses - Provides a prefab array indexed hash of Exchange-Online mailbox-supporting licenses (at least one of which is required to accomodate an EXO Usermailbox - Note: This is a static non-query-based list of license. The function must be manually updated to accomodate MS licensure changes over time).
    .PARAMETER Mailboxes
    .NOTES
    Version     : 1.0.0
    Author      : Todd Kadrie
    Website     :	http://www.toddomation.com
    Twitter     :	@tostka / http://twitter.com/tostka
    CreatedDate : 2022-02-25
    FileName    : get-ExoMailboxLicenses.ps1
    License     : MIT License
    Copyright   : (c) 2022 Todd Kadrie
    Github      : https://github.com/tostka/verb-ex2010
    Tags        : Powershell
    REVISIONS
    * 1:15 PM 3/21/2022 refactored, updated CBH example
    * 4:14 PM 3/7/2022 updated CBH exmpl
    * 2:21 PM 3/1/2022 updated CBH
    * 4:27 PM 2/25/2022 init vers
    .DESCRIPTION
    get-ExoMailboxLicenses - Provides a prefab array indexed hash of Exchange-Online mailbox-supporting licenses (at least one of which is required to accomodate an EXO Usermailbox - Note: This is a static non-query-based list of license. The function must be manually updated to accomodate MS licensure changes over time).
    .PARAMETER TenOrg
    TenantTag value, indicating Tenants to connect to[-TenOrg 'TOL']
    .EXAMPLE
    PS>  $hSummary=[ordered]@{ AADUAssignedLicenses = $null ; AADUserPrincipalName = $null ; IsExoLicensed = $false ; } ;
    PS>  $pltGLPList=[ordered]@{ TenOrg= $TenOrg; verbose=$($VerbosePreference -eq "Continue") ; credential= $pltRXO.credential ; } ;
    PS>  $smsg = "$($tenorg):get-AADlicensePlanList w`n$(($pltGLPList|out-string).trim())" ;
    PS>  if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
    PS>  else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
    PS>  $licensePlanListHash = get-AADlicensePlanList @pltGLPList ;
    PS>  $smsg = "$($tenorg):get-ExoMailboxLicenses" ;
    PS>  if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
    PS>  else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
    PS>  $ExMbxLicenses = get-ExoMailboxLicenses -verbose:$($VerbosePreference -eq "Continue") ;
    PS>  $smsg = "$(($ExMbxLicenses.Values|measure).count) EXO UserMailbox-supporting License summaries returned)" ;
    PS>  if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
    PS>  else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
    PS>  $hSummary.AADUserPrincipalName = $UserPrincipalName ;
    PS>  $pltGAADU=[ordered]@{ ObjectId = $UserPrincipalName ; ErrorAction = 'STOP' ; verbose = ($VerbosePreference -eq "Continue") ; } ; 
    PS>  $smsg = "Get-AzureADUser on UPN:`n$(($pltGAADU|out-string).trim())" ;
    PS>  if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
    PS>  else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
    PS>  if($AADUser = Get-AzureADUser @pltGAADU){
    PS>      $userList = $AADUser | Select -ExpandProperty AssignedLicenses | Select SkuID  ;
    PS>      $userLicenses=@() ;
    PS>      $userList | ForEach {
    PS>          $sku=$_.SkuId ;
    PS>          $userLicenses+=$licensePlanListHash[$sku].SkuPartNumber ;
    PS>      } ;
    PS>      $hSummary.AADUAssignedLicenses = $userLicenses ;
    PS>      $IsExoLicensed = $false ;
    PS>      foreach($pLic in $hSummary.AADUAssignedLicenses){
    PS>          $smsg = "--(LicSku:$($plic): checking EXO UserMailboxSupport)" ;
    PS>          if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
    PS>          else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
    PS>          if($ExMbxLicenses[$plic]){
    PS>              $hSummary.IsExoLicensed = $true ;
    PS>              $smsg = "$($mbx.userprincipalname) HAS EXO UserMailbox-supporting License:$($ExMbxLicenses[$sku].SKU)|$($ExMbxLicenses[$sku].Label)" ;
    PS>              if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info }
    PS>              else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
    PS>              break ; # no sense running whole set, break on 1st mbx-support match
    PS>          } ;
    PS>      } ;
    PS>      if(-not $hSummary.IsExoLicensed){
    PS>          $smsg = "$($mbx.userprincipalname) WAS FOUND TO HAVE *NO* EXO UserMailbox-supporting License!" ;
    PS>          if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN }
    PS>          else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
    PS>      } ;
    PS>  } else {
    PS>      $smsg = "=>Get-AzureADUser NOMATCH" ;
    PS>      $smsg = $recursetag,$smsg -join '' ;
    PS>      if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN }
    PS>      else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
    PS>  } ;
    PS>  $hSummary | write-output ;
    Expanded demo leveraging verb-AAD:get-AADlicensePlanList & verb-EXO:get-ExoMailboxLicenses, 
    to evaluate an AADUser.AssignedLicenses list for a license that supports a UserMailbox. 
    Demoes use of the get-ExoMailboxLicenses() hash returned, to perform license lookups for EXO support. 
    .EXAMPLE
    PS>  $hSummary=[ordered]@{ AADUAssignedLicenses = $null ; AADUserPrincipalName = $null ; IsExoLicensed = $false ; } ;
    PS>  $pltGLPList=[ordered]@{ TenOrg= $TenOrg; verbose=$($VerbosePreference -eq "Continue") ; credential= $pltRXO.credential ; } ;
    PS>  $licensePlanListHash = get-AADlicensePlanList @pltGLPList ;
    PS>  $ExMbxLicenses = get-ExoMailboxLicenses -verbose:$($VerbosePreference -eq "Continue") ;
    PS>  $hSummary.AADUserPrincipalName = $UserPrincipalName ;
    PS>  $pltGAADU=[ordered]@{ ObjectId = $UserPrincipalName ; ErrorAction = 'STOP' ; verbose = ($VerbosePreference -eq "Continue") ; } ;
    PS>  if($AADUser = Get-AzureADUser @pltGAADU){
    PS>      $userList = $AADUser | Select -ExpandProperty AssignedLicenses | Select SkuID  ;
    PS>      $userLicenses=@() ;
    PS>      $userList | ForEach {
    PS>          $sku=$_.SkuId ;
    PS>          $userLicenses+=$licensePlanListHash[$sku].SkuPartNumber ;
    PS>      } ;
    PS>      $hSummary.AADUAssignedLicenses = $userLicenses ;
    PS>      $IsExoLicensed = $false ;
    PS>      foreach($pLic in $hSummary.AADUAssignedLicenses){
    PS>          if($ExMbxLicenses[$plic]){
    PS>              $hSummary.IsExoLicensed = $true ;
    PS>              $smsg = "$($mbx.userprincipalname) HAS EXO UserMailbox-supporting License:$($ExMbxLicenses[$sku].SKU)|$($ExMbxLicenses[$sku].Label)" ;
    PS>              write-host "$((get-date).ToString('HH:mm:ss')):$($smsg)"  ;
    PS>              break ;
    PS>          } ;
    PS>      } ;
    PS>      if(-not $hSummary.IsExoLicensed){
    PS>          $smsg = "$($mbx.userprincipalname) WAS FOUND TO HAVE *NO* EXO UserMailbox-supporting License!" ;
    PS>          write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" ;
    PS>      } ;
    PS>  } else {
    PS>      $smsg = "=>Get-AzureADUser NOMATCH" ;
    PS>      write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" ;
    PS>  } ;
    PS>  $hSummary | write-output ;
    Simplfied 'terse' example that does the above, wo the echoing or testing. 
    .LINK
    https://github.com/tostka/verb-ex2010
    #>
    #Requires -Modules verb-IO, verb-logging, verb-Text
    [OutputType('System.Collections.Hashtable')]
    [CmdletBinding()]
    PARAM(
        [Parameter(Mandatory=$FALSE,HelpMessage="TenantTag value, indicating Tenants to connect to[-TenOrg 'TOL']")]
        [ValidateNotNullOrEmpty()]
        [string]$TenOrg = 'TOR',
        [Parameter(HelpMessage="Credential to use for this connection [-credential [credential obj variable]")]
        [System.Management.Automation.PSCredential]$Credential = $global:credTORSID
    ) ;
    
    ${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name ;
    $verbose = ($VerbosePreference -eq "Continue") ;
    
    # check if using Pipeline input or explicit params:
    if ($PSCmdlet.MyInvocation.ExpectingInput) {
        write-verbose "Data received from pipeline input: '$($InputObject)'" ;
    } else {
        # doesn't actually return an obj in the echo
        #write-verbose "Data received from parameter input: '$($InputObject)'" ;
    } ;
    

    # input table of Exchange Online assignable licenses that include a UserMailbox:
    $ExMbxLicensesTbl = @"
|SKU|Label|Notes|
|ENTERPRISEPACK|Office 365 Enterprise E3|OfficE; EXO (OL,OWA,OM,100G mbx)|
|EXCHANGESTANDARD|Exchange Online Plan 1|No Office; no Services; 50G mbx, No ArchiveMbx|
|SPE_F1|Microsoft 365 F3| OfficeWeb, OfficeMobile; EXO (OWA,OM 2G Mbx)|(formerly Microsoft 365 F1, renamed Mar2020)
|STANDARDPACK|OFFICE 365 E1| OfficeWeb, OfficeMobile; EXO (OWA,OM 50G Mbx)
|EXCHANGEENTERPRISE_FACULTY|Exch Online Plan 2 for Faculty|No Office; no Services; 100G mbx, +ArchiveMbx, +vmail, +DLP|
|EXCHANGE_L_STANDARD|Exchange Online (Plan 1)|No Office; no Services; 50G mbx, No ArchiveMbx|
|EXCHANGE_S_ENTERPRISE|Exchange Online Plan 2 S|No Office; no Services; 100G mbx, +ArchiveMbx, +vmail, +DLP|
|EXCHANGEENTERPRISE|Exchange Online Plan 2|No Office; no Services; 50G mbx, +ArchiveMbx, +vmail, +DLP|
|STANDARDWOFFPACK_STUDENT|O365 Education E1 for Students|OfficeWeb, OfficeMobile; EXO (OWA,OM 50G Mbx)|
|STANDARDWOFFPACK_IW_FACULTY|O365 Education for Faculty||
|STANDARDWOFFPACK_IW_STUDENT|O365 Education for Students||
|STANDARDPACK_STUDENT|Office 365 (Plan A1) for Students||
|ENTERPRISEPACKLRG|Office 365 (Plan E3)||
|STANDARDWOFFPACK_FACULTY|Office 365 Education E1 for Faculty|OfficeWeb, OfficeMobile; EXO (OWA,OM 50G Mbx)|
|ENTERPRISEWITHSCAL_FACULTY|Office 365 Education E4 for Faculty||
|ENTERPRISEWITHSCAL_STUDENT|Office 365 Education E4 for Students||
|STANDARDPACK|Office 365 Enterprise E1|OfficeWeb, OfficeMobile; EXO (OWA,OM 50G Mbx)|
|STANDARDWOFFPACK|Office 365 Enterprise E2|OfficeWeb, OfficeMobile; EXO (OWA,OM 50G Mbx), No ArchiveMbx|
|ENTERPRISEPACKWITHOUTPROPLUS|Office 365 Enterprise E3 without ProPlus Add-on||
|ENTERPRISEWITHSCAL|Office 365 Enterprise E4||
|ENTERPRISEPREMIUM|Office 365 Enterprise E5|OfficE; EXO (OL,OWA,OM,100G mbx),AAD P1 & P2, Az Info Protection Plan 2; UC; ATP|
|DESKLESSPACK_YAMMER|Office 365 Enterprise K1 with Yammer||
|DESKLESSPACK|Office 365 Enterprise K1 without Yammer||
|DESKLESSWOFFPACK|Office 365 Enterprise K2||
|MIDSIZEPACK|Office 365 Midsize Business||
|STANDARDWOFFPACKPACK_FACULTY|Office 365 Plan A2 for Faculty||
|STANDARDWOFFPACKPACK_STUDENT|Office 365 Plan A2 for Students||
|ENTERPRISEPACK_FACULTY|Office 365 Plan A3 for Faculty||
|ENTERPRISEPACK_STUDENT|Office 365 Plan A3 for Students||
|OFFICESUBSCRIPTION_FACULTY|Office 365 ProPlus for Faculty||
|LITEPACK_P2|Office 365 Small Business Premium||
|SPE_E3|MICROSOFT 365 E3|OfficeWeb, OfficeMobile; EXO (OL,OWA,OM 2G Mbx)||
|SPE_E5|MICROSOFT 365 E5||
"@ ;
    $ExMbxLicenses = $ExMbxLicensesTbl | convertfrom-markdowntable ;

    # building a CustObj (actually an indexed hash) with the default quota specs from all db's. The 'index' for each db, is the db's Name (which is also stored as Database on the $mbx)
    $smsg = "(converting $(($ExMbxLicenses|measure).count) UserMailbox-supporting o365 Licenses to indexed hash)" ;     
    if($verbose){
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
    } ; 
    if($host.version.major -gt 2){$hExMbxLicenses = [ordered]@{} } 
    else { $hExMbxLicenses = @{} } ;
    
    $ttl = ($ExMbxLicenses|measure).count ; $Procd = 0 ; 
    foreach ($Sku in $ExMbxLicenses){
        $Procd ++ ; 
        $sBnrS="`n#*------v PROCESSING : ($($Procd)/$($ttl)) $($Sku.SKU) v------" ; 
        $smsg = $sBnrS ; 
        if($verbose){
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        } ; 
        
        $name =$($Sku | select -expand SKU) ; 
        $hExMbxLicenses[$name] = $Sku ; 

        $smsg = "$($sBnrS.replace('-v','-^').replace('v-','^-'))" ;
        if($verbose){
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        } ; 
    } ;  # loop-E

    if($hExMbxLicenses){
        $smsg = "(Returning summary objects to pipeline)" ; 
        if($verbose){
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug
            else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        } ; 
        $hExMbxLicenses | Write-Output ; 
    } else {
        $smsg = "NO RETURNABLE `$hExMbxLicenses OBJECT!" ; 
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug 
        else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
        THROW $smsg ;
    } ; 
}

#*------^ get-ExoMailboxLicenses.ps1 ^------