# Intune Remediation Detect script
# which will gather network information including
# Network Interface Name, IP Address and MAC Address
#
# Results are printed to StdOut with Write-Host to show results in Intune console
#
# Petri.Paavola@yodamiitti.fi
# Microsoft MVP - Window and Intune
# 29.1.2025

# Get all network interfaces
$NetAdapters = Get-NetAdapter
$NetIPConfiguration = Get-NetIPConfiguration
$NetConnectionProfiles = Get-NetConnectionProfile

# Loop through all NetAdapters
# Filter in NetworkDevices with Status "Up"
# Get Name, IPAddress, MACAddress and Wireless network SSID
# Save results to array of objects

$NetInfo = @()
foreach ($NetAdapter in $NetAdapters) {
    if ($NetAdapter.Status -eq "Up") {
        $NetIPConfig   = $NetIPConfiguration | Where-Object InterfaceIndex -eq $NetAdapter.ifIndex

        $InterfaceName = $NetAdapter.Name
        $IPv4Address   = $NetIPConfig.IPv4Address.IPv4Address
        $IPv4DefaultGateway = $NetIPConfig.IPv4DefaultGateway
        $MACAddress    = $NetAdapter.MacAddress
        $NetConnectionProfileName          = $null

        # Get Wired Network and Wireless Network Name (SSID) if available
        $NetConnectionProfile = $NetConnectionProfiles | Where-Object InterfaceIndex -eq $NetAdapter.ifIndex
        if ($NetConnectionProfile.Name) {
            $NetConnectionProfileName = $NetConnectionProfile.Name
        }

        # Save results to array of objects
        # only if there is IPv4Address and DefaultGateway
        if ($IPv4Address -and $IPv4DefaultGateway) {
            $NetInfo += [PSCustomObject]@{
                InterfaceName = $InterfaceName
                IPv4Address   = $IPv4Address
                MACAddress    = $MACAddress
                NetConnectionProfileName  = $NetConnectionProfileName
            }
        }
    }
}


# Workaround to get all data shown in Intune console
# Convert objects to compressed JSON string and output to StdOut
$NetInfo | ConvertTo-Json -Compress | Write-Host
Exit 0