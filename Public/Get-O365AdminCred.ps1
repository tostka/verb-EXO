#*------v Function Get-O365AdminCred v------
Function Get-O365AdminCred {
    If (-not($o365cred)) {
        #$global:credo365XXXSID = Get-Credential -Credential $o365adminusername
        # leverage the function
        Get-AdminCred ;
    }
}#*------^ END Function Get-O365AdminCred ^------