# Check Windows licensing activation status
# use with Intune Proactive Remediations Detection script
#
#
# Petri.Paavola@yodamiitti.fi
# Microsoft MVP - Windows and Intune
# 6.2.2024


<#
	0 {$Status = "Unlicensed"}
	1 {$Status = "Licensed"}
	2 {$Status = "Out-Of-Box Grace Period"}
	3 {$Status = "Out-Of-Tolerance Grace Period"}
	4 {$Status = "Non-Genuine Grace Period"}
	5 {$Status = "Notification"}
	6 {$Status = "Extended Grace"}
#>					


#defined initial data
$LicenseStatusText = @("Unlicensed","Activated","OOB Grace","OOT Grace","Non-Genuine Grace","Notification","Extended Grace")
		
$LicenseStatus = Get-CimInstance -ClassName SoftwareLicensingProduct | Where-Object { $_.PartialProductKey -and $_.Name -like "*Windows*"} | Select-Object -ExpandProperty licenseStatus

if($LicenseStatus -eq 1) {
	Write-Host "Activated"
	Exit 0
} else {
	Write-Host "$($LicenseStatusText[$LicenseStatus])"
	Exit 1
}