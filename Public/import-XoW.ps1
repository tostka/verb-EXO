#*------v import-XoW v------
function import-XoW_func {
    <#
    .SYNOPSIS
    import-XoW - import freestanding local invoke-XOWrapper_func.ps1 (back fill lack of xow support in verb-exo mod)
    .NOTES
    Version     : 1.0.0.
    Author      : Todd Kadrie
    Website     : http://www.toddomation.com
    Twitter     : @tostka / http://twitter.com/tostka
    CreatedDate : 2021-07-13
    FileName    : import-XoW_func.ps1
    License     : MIT License
    Copyright   : (c) 2021 Todd Kadrie
    Github      : https://github.com/tostka/verb-XXX
    Tags        : Powershell
    AddedCredit : REFERENCE
    AddedWebsite: URL
    AddedTwitter: URL
    REVISIONS
    * 10:07 AM 3/28/2023 typo, didn't have $ModuleName in trailing w-h
    * 10:20 AM 3/27/2023 completing work, added CBH demo
    * 10:32 AM 3/24/2023 flip wee lxoW into full function call
    .DESCRIPTION
    import-XoW - import freestanding local invoke-XOWrapper_func.ps1 (back fill lack of xow support in verb-exo mod)
    .INPUTS
    None. Does not accepted piped input.
    .OUTPUTS
    None.
    .EXAMPLE
    PS> if(-not(get-command invoke-XoWrapper -ea 0)){
    PS>     write-verbose "need the _func.ps1 to target, gcm doesn't do substrings, wo a wildcard" ; 
    PS>     if(-not($lmod = get-command import-XoW_func.ps1)){
    PS>         write-verbose "found local $($lmod.source), deferring to..." ; 
    PS>         ipmo -fo -verb $lmod ; 
    PS>     } else {
    PS>         #*------v import-XoW v------
    PS>             ## pasted copy of this function
    PS>         #*------^ import-XoW ^------
    PS>     } ; ;
    PS>     lxoW -verbose ;
    PS> } ; 
    Call as Local override demo for host scripts
    .LINK
    https://github.com/tostka/verb-exo
    #>
    [CmdletBinding()]
    [Alias('lxoW')]
    PARAM(
        [Parameter(Mandatory=$false,HelpMessage="Tenant Tag (3-letter abbrebiation)[-TenOrg 'XYZ']")]
        #[ValidateNotNullOrEmpty()]
        [string]$ModuleName = 'invoke-XOWrapper_func.ps1'
    ) ;
    write-verbose "ipmo invoke-XOWrapper/xOW function" ;
    if($iflpath = get-command $ModuleName | select -expand source){ 
        if(test-path $iflpath){
            $tMod = $iflpath ; 
        }elseif(test-path (join-path -path 'C:\usr\work\o365\scripts\' -childpath $ModuleName)){
            $tMod = (join-path -path 'C:\usr\work\o365\scripts\' -childpath $ModuleName) ;  
        } else {throw 'Unable to locate xoW_func.ps1!' ;
            break ;
        } ;
        if($tmod){
            write-verbose 'Check for preloaded target function' ; 
            if(-not(get-command (split-path $tmod -leaf).replace('_func.ps1',''))){ 
                write-verbose "`$tMod:$($tMod)" ;
                Import-Module -force -verbose $tMod ;
            } else { write-host "($tmod already loaded)" } ;
        } else { write-warning "unable to resolve `$tmod!" } ;
    } else { 
        throw "Unable to locate/gcm $($ModuleName)" ; 
    } ;  
 }
 #*------^ import-XoW ^------