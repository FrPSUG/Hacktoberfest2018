Function Get-TeamViewerID {

    param(
        [string] $Hostname,
        [switch] $Copy
    )


    #Variables
    $Target = $Hostname
    If (!$Target) {$Target = $env:COMPUTERNAME}
    
    
    #Start Remote Registry Service
    If ($Target -ne $env:COMPUTERNAME) {
        $Service = Get-Service -Name "Remote Registry" -ComputerName $Target
        $Service.Start()
    }


    #Suppresses errors (comment to disable error suppression)
    $ErrorActionPreference = "SilentlyContinue"


    #Attempts to pull clientID value from remote registry and display it if successful
    $RegCon = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $Target)
    $RegKey = $RegCon.OpenSubKey("SOFTWARE\\WOW6432Node\\TeamViewer")
    $ClientID = $RegKey.GetValue("clientID")


    #If previous attempt was unsuccessful, attempts the same from a different location
    If (!$clientid) {
        $RegKey = $RegCon.OpenSubKey("SOFTWARE\\WOW6432Node\\TeamViewer\Version9")
        $ClientID = $RegKey.GetValue("clientID")
    }


    #If previous attempt was unsuccessful, attempts the same from a different location
    If (!$clientid) {
        $RegKey = $RegCon.OpenSubKey("SOFTWARE\\TeamViewer")
        $ClientID = $RegKey.GetValue("clientID")
    }


    #Stop Remote Registry service
    If ($Target -ne $env:COMPUTERNAME) {
        $Service.Stop()
    }


    #Display results
    Write-Host
    If (!$clientid) {Write-Host "ERROR: Unable to retrieve clientID value via remote registry!" -ForegroundColor Red}
    Else {Write-Host "TeamViewer client ID for $Target is $Clientid." -ForegroundColor Yellow}
    Write-Host


    #Copy to clipboard
    If ($copy -and $ClientID) {$ClientID | clip}
    return $ClientID
}