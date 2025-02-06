# Intune Remediation Detection script to check Windows Update registry values
# This script checks the following registry values:
#   - HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\DoNotConnectToWindowsUpdateInternetLocations
#   - HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\UseWUServer
# The expected values are:
#   - DoNotConnectToWindowsUpdateInternetLocations: 0
#   - UseWUServer: 1
# If any of the registry values have an unexpected value, the script will return a non-compliant status.
# It is ok to not have those registry values at all
#
# Petri.Paavola@yodamiitti.fi
# Microsoft MVP - Windows and Intune
# 2025-02-06


# Define compliance flag
$compliant = $true

# Define return string
$returnString = ''

# Define Windows Update registry paths and expected values
$wuRegPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate"
$key1Name  = "DoNotConnectToWindowsUpdateInternetLocations"
$expectedValue1 = 0

$auRegPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU"
$key2Name  = "UseWUServer"
$expectedValue2 = 1

# Check the first registry key
if (Test-Path $wuRegPath) {

    # First we need to determine if value DoNotConnectToWindowsUpdateInternetLocations exists
    $key1 = Get-ItemProperty -Path $wuRegPath -Name $key1Name -ErrorAction SilentlyContinue
    if($key1) {
        # Value DoNotConnectToWindowsUpdateInternetLocations exists
        # Now we need to check if the value is correct
        if ($key1.$key1Name -ne $expectedValue1) {
            $returnString = "Non-compliant: '$wuRegPath\$key1Name' has value $($key1.$key1Name) (expected $expectedValue1). "
            $compliant = $false
        }
    }
}

# Check the second registry key
if (Test-Path $auRegPath) {
        
    # First we need to determine if value UseWUServer exists
    $key2 = Get-ItemProperty -Path $auRegPath -Name $key2Name -ErrorAction SilentlyContinue
    if($key2) {
        # Value UseWUServer exists
        # Now we need to check if the value is correct
        if ($key2.$key2Name -ne $expectedValue2) {
            $returnString =  "$returnString Non-compliant: '$auRegPath\$key2Name' has value $($key2.$key2Name) (expected $expectedValue2)."
            $compliant = $false
        }
    }
}



# Output compliance status and exit appropriately
if ($compliant) {
    Write-Output "WindowsUpdate registry values are compliant."
    exit 0
} else {
    Write-Host "$returnString"
   exit 1
}