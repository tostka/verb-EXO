# resolve-AppIDToCBAFriendlyName.ps1

#*------v resolve-AppIDToCBAFriendlyName.ps1 v------
function resolve-AppIDToCBAFriendlyName {
    <#
    .SYNOPSIS
    resolve-AppIDToCBAFriendlyName - Delivers a 'username' equivelent for CBA authenticated connections: Resolves AppID (commonly from EOM:get-connectioninformation cmdlet) to equiv CBA cert FriendlyName (resolves AppID to password on a local profile\Keys\*.psxml cred file (which contins the cred AppID); then resolves the matched cred.username (which is a local installed cert Thumbprint) to the locally-installed cert, and returns the cert's FriendlyName.
    .NOTES
    Version     : 1.0.0.0
    Author      : Todd Kadrie
    Website     :	http://www.toddomation.com
    Twitter     :	@tostka / http://twitter.com/tostka
    CreatedDate : 2024-07-03
    FileName    : resolve-AppIDToCBAFriendlyName.ps1
    License     : MIT License
    Copyright   : (c) 2024 Todd Kadrie
    Github      : https://bitbucket.com/tostka/verb-Auth
    REVISIONS
    * 10:19 AM 7/3/2024 init, adapted from resolve-usernametouserrole(); 
    .DESCRIPTION
    resolve-AppIDToCBAFriendlyName - Delivers a 'username' equivelent for CBA authenticated connections: Resolves AppID (commonly from EOM:get-connectioninformation cmdlet) to equiv CBA cert FriendlyName (resolves AppID to password on a local profile\Keys\*.psxml cred file (which contins the cred AppID); then resolves the matched cred.username (which is a local installed cert Thumbprint) to the locally-installed cert, and returns the cert's FriendlyName.

    Returns local cert's FriendlyName, as a connected 'username' equivelent for Certificate-Based-Authentication (CBA) Exchange Online connections 

    UserRole: (SID|CSID|UID|B2BI|CSVC|ESVC|LSVC|ESvcCBA|CSvcCBA|SIDCBA)
    Service = (Op|O365)
    TenOrg = Tenant 3-letter 'Tag' ; 
    And for Certificate-Based-Authentication (CBA), it also returns: 
    FriendlyName: [certificate's configured FriendlyName value, which generally summarizes the Service, TenOrg & User role as well].

    .PARAMETER Credential
    Credential to be resolved to UserRole[-Credential [credential object]]
    .PARAMETER AppId
    Guid corresponding to an Entra/AzureAD Registered AppID to be resolved to UserRole summary[-AppID 'dannnnad-endn-nnan-nnnn-nenennnnnafe']
    .EXAMPLE
    PS> if($xosession = get-connectioninformation){
    PS>     if($xosession.CertificateAuthentication){
    PS>         $uRoleReturn = resolve-AppIDToCBAFriendlyName -AppID $xosession.AppId -verbose:$($VerbosePreference -eq "Continue")  ; 
    PS>         $certUname = $uRoleReturn.FriendlyName ; 
    PS>         $certTag = $uRoleReturn.TenOrg ; 
    PS>     } ; 
    PS>     if($xoSession.TokenStatus -eq 'Active'){
    PS>         $smsg = "Connected to " ; 
    PS>         switch ($xosession.IsEopSession){
    PS>           $true { $smsg += "Sec & Compl PS "}  
    PS>           $false {$smsg += "XO EOM PS " } 
    PS>         } ; 
    PS>         if($xosession.CertificateAuthentication){
    PS>             $smsg += "using CBA:" ; 
    PS>             $smsg += " $($certUname)" ; 
    PS>         } ; 
    PS>         write-host $smsg ; 
    PS>         $prpConn = 'Organization','UserPrincipalName','ModulePrefix','CertificateAuthentication','AppId','TenantID','ConnectionId','IsEopSession','TokenStatus','State' ; 
    PS>         $hsDetails = @"
    PS> Conneciton Details:
    PS> $(($xosession | select $prpConn[0..2] | ConvertTo-Markdowntable -Border -NoDashRow|out-string).trim())
    PS> $(($xosession | select $prpConn[3..5] | ConvertTo-Markdowntable -Border -NoDashRow|out-string).trim())
    PS> $(($xosession | select $prpConn[6..9] | ConvertTo-Markdowntable -Border -NoDashRow|out-string).trim())
    PS> "@ ; 
    PS>         write-verbose $hsDetails ;  
    PS>     } else { 
    PS>         $smsg = "Not currently connected (TokenStatus:$($xoSession.TokenStatus))" ; 
    PS>         $smsg += "`nPreviously: " 
    PS>         switch ($xosession.IsEopSession){
    PS>           $true { $smsg += "Sec & Compl PS"}  
    PS>           $false {$smsg += "XO EOM PS" } 
    PS>         } ; 
    PS>         if($xosession.CertificateAuthentication){
    PS>             $smsg += " using CBA:" ; 
    PS>             $smsg += " $($certUname)" ; 
    PS>         } ; 
    PS>         write-host -foregroundcolor yellow $smsg ; 
    PS>     } ; 
    PS> } else { 
    PS>     write-host -foregroundcolor yellow "No connection info returned" ; 
    PS> } ;    
    Demo parsing get-connectioninformation results to report connection details
    .EXAMPLE
    $uRoleReturn = resolve-AppIDToCBAFriendlyName -UserName $Credential.username -verbose:$($VerbosePreference -eq "Continue") ; 
    Resolve Username string into UserRole value
    .EXAMPLE
    $uRoleReturn = resolve-AppIDToCBAFriendlyName -Credential $Credential -verbose = $($VerbosePreference -eq "Continue") ; 
    Resolve Credential object into UserRole value
    .LINK
    https://bitbucket.com/tostka/verb-Auth
    #>
    [CmdletBinding()] 
    Param(
        #[Parameter(Mandatory = $false, HelpMessage = "Credential to be resolved to UserRole[-Credential [credential object]]")]
        #    [System.Management.Automation.PSCredential]$Credential,
        [Parameter(Mandatory=$false,HelpMessage="Guid corresponding to an Entra/AzureAD Registered AppID to be resolved to UserRole summary[-AppID 'dannnnad-endn-nnan-nnnn-nenennnnnafe']")]
            [ValidateScript({
                [boolean]([guid]$_)
            })]
            [string[]]$AppId
    ) ;
    BEGIN {
        $verbose = ($VerbosePreference -eq "Continue") 
        if(-not $rgxCertThumbprint){$rgxCertThumbprint = '[0-9a-fA-F]{40}' } ; # if it's a 40char hex string -> cert thumbprint  
        if(-not $rgxSmtpAddr){$rgxSmtpAddr = "^([0-9a-zA-Z]+[-._+&'])*[0-9a-zA-Z]+@([-0-9a-zA-Z]+[.])+[a-zA-Z]{2,63}$" ; } ; # email addr/UPN
        if(-not $rgxDomainLogon){$rgxDomainLogon = '^[a-zA-Z][a-zA-Z0-9\-\.]{0,61}[a-zA-Z]\\\w[\w\.\- ]+$' } ; # DOMAIN\samaccountname 
        if(-not $rgxCBACertFname){$rgxCBACertFname = 'O365-o365_.*cbacert-\w{3}-\w-.*\.psxml'} ; 
        $prpcert = 'Subject','Issuer','FriendlyName','NotBefore','NotAfter','HasPrivateKey','Thumbprint' ; 
    } ;
    PROCESS {
        
        if(test-path -path (join-path (split-path $profile) 'keys')){
            if($credfiles = gci "$(join-path (split-path $profile) 'keys')\*.psxml" |? name -match $rgxCBACertFname){
                foreach($cfile in $credfiles){
                    TRY{
                        $sBnr4="`n#*``````v processing $($cfile.fullname) v``````" ; 
                        $smsg = $sBnr4 ; 
                        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                        
                        if($tmpcred= Import-Clixml $cfile.fullname -ea STOP){
                            $pw = $tmpcred.GetNetworkCredential().Password ; 
                            $smsg = "comparing AppID:`n$($AppId)" ; 
                            $smsg += "`nto $($cfile.name) cert pw:`n$($pw)" ; 
                            write-verbose $smsg ; 
                            if($AppID -eq $pw){
                                write-verbose "Matched AppID->tmpcred.Password:`ntmpcred Uname is the cert thumb: $($tmpcred.username)"
                                $uRoleReturn = [ordered]@{
                                    UserRole = $null ; 
                                    Service = $null ; 
                                    TenOrg = $null ; 
                                } ;
                                if($tcert = gci "cert:\currentuser\my\$($tmpcred.username)"){
                                    if($tcert  | ?{$_.notbefore -le (get-date ) -le $_.notafter}){
                                        $smsg = "(cert is still within NotBefore<>NotAfter range)" ; 
                                        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                                    } else { 
                                        $smsg = "matched cert for thumbprint:cert:\currentuser\my\$($tmpcred.username)" ; 
                                        $smsg += "`nIS EXPIRED!" ; 
                                        $smsg += "`n$(($tcert | FL $prpcert |out-string).trim())" ; 
                                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                                    } ; 
                                    if($tcert.FriendlyName){
                                        $smsg = "adding cert FriendlyName to return..." ; 
                                        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                                        $uRoleReturn.add('FriendlyName',$tcert.friendlyname) ; 
                                    }elseif($tcert.Subject){
                                        $tempFname = $tcert.subject.split('.')[0].replace('CN=','').replace('o365','o365_') ; 
                                        $uRoleReturn.add('FriendlyName',$tempFname) ;
                                    } else { 
                                        $smsg = "Unable to find/parse either a FriendlyName or Subject on the cert, into a suitable RoleName analog" ; 
                                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                                        else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                                    } ; 
                                    if($uRoleReturn.FriendlyName){
                                        # [regex]::match($variX.name,'^cred(?<Svc>\w{2})(?<TenOrg>\w{3})(?<URole>.*)$').captures[0].groups['URole'].value
                                        #[regex]::match($tcert.subject.split('.')[0].replace('CN=','').replace('o365','o365_'),
                                        $regexCBAFriendName = "(?<Service>.*)_(?<URole>.*)CBACert-(?<TenOrg>\w{3})" ; 
                                        #"o365_(?<URole>.*)CBACert-(?<TenOrg>\w{3})" ; 
                                        if($hits = [regex]::match($uRoleReturn.FriendlyName,$regexCBAFriendName)){
                                            $uRoleReturn.UserRole = $hits.Groups['URole'].value ; 
                                            # $ServiceNames = 'o365','OP' 
                                            $uRoleReturn.Service = $hits.Groups['Service'].value ; #'o365'
                                            $uRoleReturn.TenOrg =  $hits.Groups['TenOrg'].value
                                        }else { 
                                            $smsg = "Unable to rgx `$uRoleReturn.FriendlyName" ; 
                                            $smsg += "`n$($uRoleReturn.FriendlyName)" ; 
                                            $smsg += "`nwith $($regexCBAFriendName)" ; 
                                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                                            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                                        } ; 
                                        $smsg = "Returning `$uRoleReturn summary to pipeline`n$(($uRoleReturn|out-string).trim())" ; 
                                        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 
                                        [pscustomobject]$uRoleReturn | Write-Output ; 
                                        break ; 
                                    } ; 
                                }else {
                                    $smsg = "Unable to: gci cert:\currentuser\my\$($tmpcred.username)" ; 
                                    if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                                    else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                                } ;  
                            } ;
                        }else {
                            $smsg = "Unable to import a content from: $($cfile.fullname) " ; 
                            $SMSG += "Does not appear configured for local .psxml credential storage!" ; 
                            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
                        } ; 
                        $smsg = $sBnr4.replace('`v','`^').replace('v`','^`') ;
                        if($verbose){if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level VERBOSE } 
                        else{ write-verbose "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; } ; 

                    } CATCH {
                        $ErrTrapd=$Error[0] ;
                        $smsg = "`n$(($ErrTrapd | fl * -Force|out-string).trim())" ;
                        if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN } #Error|Warn|Debug
                        else{ write-warning "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ;
                    } ; 

                } ;  # loop-E

            } else { 
                $smsg = "No local CBA-Auth `$rgxCBACertFname files found ($($rgxCBACertFname))!" ; 
                $SMSG += "Does not appear configured for local .psxml credential storage!" ; 
                if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
                else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
            } ; 
        } else { 
            $smsg = "LOCAL PROFILE LACKS A KEYS SUBDIR!" ; 
            $SMSG += "Does not appear configured for local .psxml credential storage!" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
        } ; 

    } ;  # PROC-E
    END{
        if(-not $uRoleReturn.UserRole){
            $smsg = "FAILED TO RESOLVE AppId :$($AppId) succesffully against an installed local .psxml file & installed certificate combo!" ; 
            if ($logging) { Write-Log -LogContent $smsg -Path $logfile -useHost -Level WARN -Indent} 
            else{ write-WARNING "$((get-date).ToString('HH:mm:ss')):$($smsg)" } ; 
            $uRoleReturn = [ordered]@{
                UserRole = 'FAILED_RESOLUTION' ; 
                Service = 'FAILED_RESOLUTION' ; 
                TenOrg = 'FAILED_RESOLUTION' ;  
            } ;
            [pscustomobject]$uRoleReturn | Write-Output ; 
        } ;  
    } ;
}
#*------^ resolve-AppIDToCBAFriendlyName.ps1 ^------
