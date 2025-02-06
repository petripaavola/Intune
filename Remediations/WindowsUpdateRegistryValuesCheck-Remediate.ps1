# Intune Remediation Remediate script to change Windows Update registry values
# This script changes the following registry values:
#   - HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\DoNotConnectToWindowsUpdateInternetLocations
#   - HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU\UseWUServer
# The expected values to be set are:
#   - DoNotConnectToWindowsUpdateInternetLocations: 0
#   - UseWUServer: 1
#
# Petri.Paavola@yodamiitti.fi
# Microsoft MVP - Windows and Intune
# 2025-02-06


# Define compliance flag
$remediationSuccessful = $true

# Define return string
$returnString = $null

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
        if ($key1.$key1Name -eq $expectedValue1) {
            # Registry key value is correct, no need to change it
            $returnString = "Compliant: '$wuRegPath\$key1Name' has value $expectedValue1. "
        } else {
            # Value is wrong so we need to change it
            Set-ItemProperty -Path $wuRegPath -Name $key1Name -Value $expectedValue1 -Force
            $Success = $?
            if($Success) {
                $returnString = "Remediated: '$wuRegPath\$key1Name' value changed to $expectedValue1. "
            } else {
                $returnString = "Failed to remediate: '$wuRegPath\$key1Name' value change to $expectedValue1. "
                $remediationSuccessful = $false
            }
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
        if ($key2.$key2Name -eq $expectedValue2) {
            # Registry key value is correct, no need to change it
            $returnString =  "$returnString Remediated: '$auRegPath\$key2Name' value changed to $expectedValue2."
        } else {
            # Value is wrong so we need to change it
            Set-ItemProperty -Path $auRegPath -Name $key2Name -Value $expectedValue2 -Force
            $Success = $?
            if($Success) {
                $returnString =  "$returnString Remediated: '$auRegPath\$key2Name' value changed to $expectedValue2."
            } else {
                $returnString =  "$returnString Failed to remediate: '$auRegPath\$key2Name' value change to $expectedValue2."
                $remediationSuccessful = $false
            }
        }
    }
}



# Output compliance status and exit appropriately
if ($remediationSuccessful) {
    if ($returnString) {
        # Remediation was successful and was done because we have values in $returnString
        Write-Host "$returnString"
    } else {
        # Variable $returnString is empty so registry values did not exist at all so we are Compliant
        Write-Host "WindowsUpdate registry keys do not exist. No remediation needed, we are Compliant"
    }
    exit 0
} else {
    Write-Host "$returnString"
   exit 1
}