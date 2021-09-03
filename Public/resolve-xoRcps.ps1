#*------v Function Resolve-xoRcps v------
function Resolve-xoRcps {
    <#
    .SYNOPSIS
    Resolve-xoRcps.ps1 - run a get-exorecipient to re-resolve an array of Recipients into the matching primarysmtpaddress
    .NOTES
    Version     : 1.0.0
    Author      : Todd Kadrie
    Website     :	http://www.toddomation.com
    Twitter     :	@tostka / http://twitter.com/tostka
    CreatedDate : 2021-09-02
    FileName    : Resolve-xoRcps
    License     : MIT License
    Copyright   : (c) 2020 Todd Kadrie
    Github      : https://github.com/tostka/verb-XXX
    Tags        : Powershell
    AddedCredit : REFERENCE
    AddedWebsite:	URL
    AddedTwitter:	URL
    REVISIONS
    .DESCRIPTION
    Resolve-xoRcps.ps1 - run a get-exorecipient to re-resolve an array of Recipients into the matching primarysmtpaddress
    .PARAMETER Recipients
    Array of Recipients to be resolved against current Exchange environment [-Recipients `$ModeratedBy ]
    .PARAMETER MatchRecipientTypeDetails
    Regex for RecipientTypeDetails value to require for matched Recipients [-MatchRecipientTypeDetails '(UserMailbox|MailUser)']
    .PARAMETER BlockRecipientTypeDetails
    Regex for RecipientTypeDetails value to filter out of matched Recipients [-Block '(MailContact|GuestUser)']
    .PARAMETER useEXOv2
    Use EXOv2 (ExchangeOnlineManagement) over basic auth legacy connection [-useEXOv2]
    .PARAMETER Whatif
    Parameter to run a Test no-change pass [-Whatif switch]
    .EXAMPLE
    PS> .\Resolve-xoRcps.ps1
    .EXAMPLE
    PS> if($pltSDdg.RejectMessagesFrom){
            $pltSDdg.RejectMessagesFrom = (Resolve-xoRcps -Recipients $srcDg.RejectMessagesFrom -MatchRecipientTypeDetails '(UserMailbox|MailUser|MailContact)' -Verbose:($VerbosePreference -eq 'Continue') )  ; 
        } ;
        Resolve recip designators on the RejectMessagesFrom value, to EXO recipient objects, and return the primarysmtpaddress
    .LINK
    https://github.com/tostka/verb-EXO
    #>
    [CmdletBinding()]
    PARAM(
        [Parameter(Mandatory=$True,HelpMessage="Array of Recipients to be resolved against current Exchange environment [-Recipients `$ModeratedBy ]")]
        [array]$Recipients,
        [Parameter(HelpMessage="Regex for RecipientTypeDetails value to require for matched Recipients [-MatchRecipientTypeDetails '(UserMailbox|MailUser)']")]
        [string]$MatchRecipientTypeDetails,
        [Parameter(HelpMessage="Regex for RecipientTypeDetails value to filter out of matched Recipients [-Block '(MailContact|GuestUser)']")]
        [string]$BlockRecipientTypeDetails,
        [Parameter(HelpMessage="Use EXOv2 (ExchangeOnlineManagement) over basic auth legacy connection [-useEXOv2]")]
        [switch] $useEXOv2
    ) 
    if ($script:useEXOv2) { reconnect-eXO2 }
    [array]$cmdletMaps = 'ps1GetxRcp;get-exorecipient;' ;
    foreach($cmdletMap in $cmdletMaps){
        if($script:useEXOv2){
            if(!($cmdlet= Get-Command $cmdletMap.split(';')[1].replace('-exo','-xo') )){ throw "unable to gcm Alias definition!:$($cmdletMap.split(';')[1])" ; break }
            $nAName = ($cmdletMap.split(';')[0]) ; 
            if(!($nalias = get-alias -name $nAName -ea 0 )){
                $nalias = set-alias -name $nAName -value ($cmdlet.name) -passthru ;
                write-verbose "$($nalias.Name) -> $($nalias.ResolvedCommandName)" ;
            } ;
        } else {
            if(!($cmdlet= Get-Command $cmdletMap.split(';')[1])){ throw "unable to gcm Alias definition!:$($cmdletMap.split(';')[1])" ; break }
            $nAName = ($cmdletMap.split(';')[0]);
            if(!($nalias = get-alias -name $nAName -ea 0 )){
                $nalias = set-alias -name ($cmdletMap.split(';')[0]) -value ($cmdlet.name) -passthru ;
                write-verbose "$($nalias.Name) -> $($nalias.ResolvedCommandName)" ;
            } ; 
        } ;
    } ;
    if ($script:useEXOv2) { reconnect-eXO2 }
    else { reconnect-EXO } ;
    if($Recipients){
        $resolvedRecipients = $Recipients | foreach-object {
            ps1GetxRcp $_ 
        } ; 
        if($MatchRecipientTypeDetails){
            $resolvedRecipients = $resolvedRecipients |?{$_.RecipientTypeDetails -match $MatchRecipientTypeDetails} ; 
        } ; 
        if($BlockRecipientTypeDetails){
            $resolvedRecipients = $resolvedRecipients |?{$_.RecipientTypeDetails -notmatch $BlockRecipientTypeDetails} ; 
        } ; 
        $resolvedRecipients.primarysmtpaddress |write-output ;
    } else { 
        write-host "No Recipients specified" ;
        $null | write-output ;
    } ; 
} ; 
#*------^ END Function Resolve-xoRcps ^------