# Get WLAN profiles installed to device
#
# Petri.Paavola@yodamiitti.fi
# Microsoft MVP - Windows and Intune
# 6.2.2025


# Run netsh command to get WLAN profiles
$netshOutput = netsh wlan show profiles

# Extract profile names from netsh output
$profileNames = $netshOutput | Select-String -Pattern ':\s+(.+)$' | ForEach-Object { $_.Matches.Groups[1].Value.Trim() }

# Create custom objects
$profileObjects = $profileNames | ForEach-Object {
    [PSCustomObject]@{
        ProfileName = $_
    }
}

# Convert to compressed JSON
$json = $profileObjects | ConvertTo-Json -Compress

# Output the JSON
$json

exit 0