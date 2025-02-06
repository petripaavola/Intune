# Intune Win32App custom Detection script will check if device TimeZone is 'FLE Standard Time'
#
#
# Petri.Paavola@yodamiitti.fi
# Microsoft MVP - Windows and Intune
# 2025-02-06

# Get current TimeZone
$currentTimeZone = (Get-TimeZone).Id

if ($currentTimeZone -eq "FLE Standard Time") {
    Write-Host "TimeZone is correctly set to 'FLE Standard Time'."
    exit 0
} else {
    # Write-Host "TimeZone is NOT set to 'FLE Standard Time'. We need to 'Remediate'"
    exit 1
}
