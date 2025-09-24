# Get Windows MDM Assigned Access / Kiosk configuration
# We are especially interested on XML configuration
#
# Run this script in SYSTEM Powershell using: psexec.exe -sid Powershell.exe
#
# Petri.Paavola@yodamiitti.fi
# Microsoft MVP - Windows and Intune

# Getting MDM policy works only as SYSTEM
$aa = Get-WmiObject -Namespace "root\cimv2\mdm\dmmap" -Class MDM_AssignedAccess -ErrorAction Stop

if ([string]::IsNullOrWhiteSpace($aa.Configuration)) {
    Write-Host "No AssignedAccess (kiosk) policy applied."
}
else {
    Write-Host "=== MDM_AssignedAccess Properties ==="
    $aa | Select-Object InstanceID, ParentID, KioskModeApp | Format-List
    Write-Host ""

    Write-Host "=== AssignedAccess Configuration (raw XML) ==="
    $aa.Configuration
    Write-Host ""
}
