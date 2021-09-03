#*------v Function new-xoDGFromProperty v------
function new-xoDGFromProperty{
    <#
    .SYNOPSIS
    new-xoDGFromProperty.ps1 - run a get-exorecipient to re-resolve an array of Recipients into the matching primarysmtpaddress
    .NOTES
    Version     : 1.0.0
    Author      : Todd Kadrie
    Website     :	http://www.toddomation.com
    Twitter     :	@tostka / http://twitter.com/tostka
    CreatedDate : 2021-09-02
    FileName    : new-xoDGFromProperty
    License     : MIT License
    Copyright   : (c) 2020 Todd Kadrie
    Github      : https://github.com/tostka/verb-XXX
    Tags        : Powershell
    AddedCredit : REFERENCE
    AddedWebsite:	URL
    AddedTwitter:	URL
    REVISIONS
    * 9:45 AM 9/2/2021 rev: added CBH, fixed existing block: Add-DistributionGroupMember -> propr xo alias:ps1AddxDistGrpMbr
    .DESCRIPTION
    new-xoDGFromProperty.ps1 - run a get-exorecipient to re-resolve an array of Recipients into the matching primarysmtpaddress
    .PARAMETER Members
    Array of Members to be resolved against current Exchange environment [-Members `$members ]
    .PARAMETER NewDGName
    Name to be used for New DG to be populated[-NewDGName (`"`$(`$preDDG.name)-ApprovedSenders`
    .PARAMETER ManagedBy (override; defaults to ManagedBy of specified DDG)# [-ManagedBy `$preDDG.ManagedBy]
    .PARAMETER useEXOv2
    Use EXOv2 (ExchangeOnlineManagement) over basic auth legacy connection [-useEXOv2]
    .PARAMETER Whatif
    Parameter to run a Test no-change pass [-Whatif switch]
    .EXAMPLE
    PS> $pltNxoDGfP=[ordered]@{
        Members=$preDDG.AcceptMessagesOnlyFrom  ;
        NewDGName=("$($preDDG.name)-ApprovedSenders") ;
        ManagedBy=$preDDG.ManagedBy ;
        whatIf=$true ;
    } ; 
    if($nDG = new-xoDGFromProperty @pltNxoDGfP){
        set-exoDynamicDistributionGroup -id $preDDG.primarysmtpaddress -AcceptMessagesOnlyFromDLMembers $nDG.primarysmtpaddress -AcceptMessagesOnlyFrom $null -whatif ; 
    } ; 
    Generate a new DG to host a transplanted recipients value (to shift static AcceptMessagesOnlyFrom to a setparte SD-managable DG).
    Then demo's updating a the source DDG, adding the new created DG onto the DDG.AcceptMessagesOnlyFromDLMembers, 
    and blanking the original DDG.AcceptMessagesOnlyFrom.
    .LINK
    https://github.com/tostka/verb-Exo
    #>
    [CmdletBinding()]
    PARAM(
        [Parameter(Mandatory=$False,HelpMessage="Array of Members to be resolved against current Exchange environment [-Members `$members ]")]
        [array]$Members,
        [Parameter(Mandatory=$True,HelpMessage="Name to be used for New DG to be populated[-NewDGName (`"`$(`$preDDG.name)-ApprovedSenders`" ;)]")]
        [string]$NewDGName,
        [Parameter(Mandatory = $false, HelpMessage = "ManagedBy (override; defaults to ManagedBy of specified DDG)# [-ManagedBy `$preDDG.ManagedBy]")]
        $ManagedBy,
        [Parameter(HelpMessage="Use EXOv2 (ExchangeOnlineManagement) over basic auth legacy connection [-useEXOv2]")]
        [switch] $useEXOv2,
        [Parameter(HelpMessage="Whatif Flag (defaults true, override -whatif:`$false) [-whatIf]")]
        [switch]$whatIf
    ) 
    if ($script:useEXOv2) { reconnect-eXO2 }
    else { reconnect-EXO } ;
    [array]$cmdletMaps = 'ps1GetxRcp;get-exorecipient;','ps1GetxDistGrp;get-exoDistributionGroup',
        'ps1NewxDistGrp;new-exoDistributionGroup' ,'ps1SetxDistGrp;set-exoDistributionGroup',
        'ps1GetxDistGrpMbr;get-exoDistributionGroupMember','ps1RmvxDistGrpMbr;remove-exoDistributionGroupMember',
        'ps1AddxDistGrpMbr;Add-exoDistributionGroupMember','ps1GetxDDG;Get-exoDynamicDistributionGroup',
        'ps1NewxDDG;New-exoDynamicDistributionGroup','ps1SetxDDG;Set-exoDynamicDistributionGroup',
        'ps1GetxOrgCfg;Get-exoOrganizationConfig' ;
    foreach($cmdletMap in $cmdletMaps){
        if($script:useEXOv2){
            if(!($cmdlet= Get-Command $cmdletMap.split(';')[1].replace('-exo','-xo') )){ throw "unable to gcm Alias definition!:$($cmdletMap.split(';')[1])" ; break }
            $nAName = ($cmdletMap.split(';')[0]) ;
            if(-not(get-alias -name $naname -ea 0 |?{$_.Definition -eq $cmdlet.name})){
                $nalias = set-alias -name $nAName -value ($cmdlet.name) -passthru ;
                write-verbose "$($nalias.Name) -> $($nalias.ResolvedCommandName)" ;
            } ;
        } else {
            if(!($cmdlet= Get-Command $cmdletMap.split(';')[1])){ throw "unable to gcm Alias definition!:$($cmdletMap.split(';')[1])" ; break }
            $nAName = ($cmdletMap.split(';')[0]);
            if(-not(get-alias -name $naname -ea 0 |?{$_.Definition -eq $cmdlet.name})){
                $nalias = set-alias -name ($cmdletMap.split(';')[0]) -value ($cmdlet.name) -passthru ;
                write-verbose "$($nalias.Name) -> $($nalias.ResolvedCommandName)" ;
            } ;
        } ;
    } ; 
    if($ManagedBy){$oManagedBy = ps1GetxRcp $ManagedBy -ea 'STOP' | select -expand primarysmtpaddress  | select -unique ;} ; 
    if($members){
        $members = $members | ps1GetxRcp -ErrorAction Continue | select -expand primarysmtpaddress  | select -unique ;
    } ; 
    $pltNDG=[ordered]@{
        DisplayName=$NewDGName;
        Name=$NewDGName;
        Members=$members ;
        #DomainController=$domaincontroller;
        Alias=([System.Text.RegularExpressions.Regex]::Replace($NewDGName,"[^1-9a-zA-Z_]",""));
        ManagedBy=$oManagedBy;
        #OrganizationalUnit = (get-organizationalunit (($preDDG.DistinguishedName.tostring().split(",") | select -Skip 1) -join ",").tostring()).CanonicalName ;
        ErrorAction = 'Stop' ; 
        whatif=$($whatif);
    } ; 
    if($existDG=ps1GetxDistGrp -id $pltndg.alias -ResultSize 1 -ea 0){
        $pltSetDG=[ordered]@{
            identity = $existDG.primarysmtpaddress ; 
            #Members=$members ; # not supported have to add-DistributionGroupMember them in on existings
            #DomainController=$domaincontroller;
            ManagedBy=$oManagedBy;
            whatif=$($whatif);
            ErrorAction = 'Stop' ; 
        } ; 
        $smsg = "UpdateExisting DG:$((get-alias ps1SetxDistGrp).definition)  w`n$(($pltSetDG|out-string).trim())" ; 
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        ps1SetxDistGrp @pltSetDG ;
        # pre-purge
        $prembrs = ps1GetxDistGrpMbr -id $pltSetDG.identity ;
        $pltModDGMbr=[ordered]@{identity= $pltSetDG.identity ;whatif = $($whatif) ;erroraction = 'STOP'  ;confirm=$false ;}
        $smsg = "Clear existing members:$((get-alias ps1RmvxDistGrpMbr).definition) w`n$(($pltModDGMbr|out-string).trim())`n$(($prembrs |out-string).trim())" ; 
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        #$prembrs | %{ps1RmvxDistGrpMbr @$pltModDGMbr -Member $_.alias  } ; 
        $prembrs.distinguishedname | ps1RmvxDistGrpMbr @pltModDGMbr ; 
        # ps1GetxDistGrpMbr -id $pltSetDG.identity | ps1RmvxDistGrpMbr -id $pltSetDG.identity –whatif:$($whatif) -ea STOP ; 
        # then add validated from scratch
        $smsg = "re-add VALIDATED members:add-DistributionGroupMember w`n$(($pltModDGMbr|out-string).trim())`n$(($members|out-string).trim())" ; 
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        $members | ps1AddxDistGrpMbr @pltModDGMbr ; 
        $pdg =  ps1GetxDistGrp -id $pltSetDG.identity ;
    } else { 
        $smsg = "$((get-alias ps1NewxDistGrp).definition)  w`n$(($pltNDG|out-string).trim())" ; 
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        $pdg = ps1NewxDistGrp @pltNDG ;
    } ; 
    if(!$whatif){
        # was getting notfounds, trying to update the $pdg, so re-qry it from scratch, if it comes back it's *there* for updates
        $1F=$false ;Do {if($1F){Sleep -s 5} ;  write-host "." -NoNewLine ; $1F=$true ; } Until ($existDG = ps1GetxDistGrp $pltNDG.alias -EA 0) ;
        # set hidden (can't be done with new-dg command): -HiddenFromAddressListsEnabled
        $pltSetDG=[ordered]@{
            identity = $existDG.primarysmtpaddress ; 
            HiddenFromAddressListsEnabled = $true ; 
            whatif=$($whatif);
            ErrorAction = 'Stop' ; 
        } ; 
        $smsg = "HiddenFromAddressListsEnabled:UpdateExisting DG:$((get-alias ps1SetxDistGrp).definition)  w`n$(($pltSetDG|out-string).trim())" ; 
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        ps1SetxDistGrp @pltSetDG ;

        $pdg =  ps1GetxDistGrp -id $pltSetDG.identity ;
        write-verbose "Returning new DG object to pipeline" ; 
        $pdg | write-output ; 
        
    } else {
        $smsg = "(-whatif: skipping balance of process)" ; 
        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level Info } #Error|Warn|Debug 
        else{ write-host -foregroundcolor green "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
        $false | write-output ; 
    }  ;

} ; 
#*------^ END Function new-xoDGFromProperty  ^------