# Define variables
$domain = (Get-WmiObject Win32_ComputerSystem).Domain
$connectionBroker = "$env:COMPUTERNAME.$domain"

$collections = Get-RDSessionCollection -ConnectionBroker $connectionBroker |
              Select-Object -ExpandProperty CollectionName
$sessionHosts = foreach ($c in $collections) {
    Get-RDSessionHost -ConnectionBroker $connectionBroker -CollectionName $c |
        Select-Object -ExpandProperty SessionHost
}
$sessionHosts = $sessionHosts | Sort-Object -Unique

# Loop through each session host
foreach ($ts in $sessionHosts) {
    
    Write-Host "Configuring $ts..."
    
    Set-RDSessionHost -SessionHost $ts -NewConnectionAllowed NotUntilReboot -ConnectionBroker $connectionBroker
    
    Write-Host "Completed configuration for $ts"
}

$sessions = Get-RDUserSession

foreach($session in $sessions)
{
    Invoke-RDUserLogoff -HostServer $session.HostServer -UnifiedSessionID $session.UnifiedSessionId -Force
}
