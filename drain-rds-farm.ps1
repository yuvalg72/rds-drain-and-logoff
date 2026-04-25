#Requires -RunAsAdministrator
<#
    Version:  1.1
    Author:   Yuval Grimblat
    Title:    Network Security Solutions Architect and Project Manager
    Company:  Mornex LTD
    Year:     2026
    LinkedIn: https://www.linkedin.com/in/yuvalgrimblat
#>
$ErrorActionPreference = "Stop"

# Get-CimInstance replaces the deprecated Get-WmiObject
$domain           = (Get-CimInstance Win32_ComputerSystem).Domain
$connectionBroker = "$env:COMPUTERNAME.$domain"

$collections = Get-RDSessionCollection -ConnectionBroker $connectionBroker |
               Select-Object -ExpandProperty CollectionName

$sessionHosts = foreach ($c in $collections) {
    Get-RDSessionHost -ConnectionBroker $connectionBroker -CollectionName $c |
        Select-Object -ExpandProperty SessionHost
}
$sessionHosts = $sessionHosts | Sort-Object -Unique

foreach ($ts in $sessionHosts) {
    Write-Host "Configuring $ts..."
    try {
        Set-RDSessionHost -SessionHost $ts `
                          -NewConnectionAllowed NotUntilReboot `
                          -ConnectionBroker $connectionBroker
        Write-Host "Completed configuration for $ts"
    } catch {
        Write-Warning "Failed to drain $ts`: $_"
    }
}

# Pass -ConnectionBroker so sessions are retrieved from the correct broker
$sessions = Get-RDUserSession -ConnectionBroker $connectionBroker

foreach ($session in $sessions) {
    try {
        Invoke-RDUserLogoff -HostServer      $session.HostServer `
                            -UnifiedSessionID $session.UnifiedSessionId `
                            -ConnectionBroker $connectionBroker `
                            -Force
        Write-Host "Logged off session $($session.UnifiedSessionId) on $($session.HostServer)"
    } catch {
        Write-Warning "Failed to log off session $($session.UnifiedSessionId) on $($session.HostServer)`: $_"
    }
}
