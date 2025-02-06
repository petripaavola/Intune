# Disable Built-in Administrator account
# Intune Remediation Detection script
# which can be run on-demand from Intune console
#
#
# Petri.Paavola@yodamiitti.fi
# Microsoft MVP - Windows and Intune
# 6.2.2025


# Get the Administrator account name
$adminAccount = Get-LocalUser | Where-Object { $_.SID -like 'S-1-5-*-500' }
$adminName = $adminAccount.Name
#Write-Output "Administrator account name: $adminName"

# Enable the Administrator account
Disable-LocalUser -Name $adminName
$Success = $?

if($Success) {
	Write-Output "Administrator account disabled: $adminName"
	Exit 0
} else {
	Write-Output "Error disabling local administrator account: $adminName"
	Exit 1
}








