# Intune Win32 App Audacity Custom Detection check script for version 3.5.1
#
# This returns compliant with equal to 3.5.1 version
#
#
# Petri.Paavola@yodamiitti.fi
# Windows MVP - Windows and Intune
# 2024-11-01
#
# Original script source:
# https://github.com/petripaavola/Intune/blob/master/Apps/Audacity_3.5.1_Intune_Win32_Custom_DetectionCheck_Script.ps1



# Check if file exists
if(-not (Test-Path "C:\Program Files\Audacity\audacity.exe")) {
	# File does not exist
	exit 1
}


$FileVersionString = [System.Diagnostics.FileVersionInfo]::GetVersionInfo("C:\Program Files\Audacity\audacity.exe").FileVersion

# Convert , -> .
$FileVersionString = $FileVersionString.Replace(',','.')

#The below line trims the spaces before and after the version name
$FileVersionString = $FileVersionString.Trim();

# Cast variable to [version]
$FileVersion = [version]$FileVersionString


# We could also get version with this command in most other cases
#$file = Get-ChildItem "C:\Program Files\Audacity\audacity.exe"
#[version]$FileVersion = $file.versioninfo.fileversion


if ($FileVersion -eq "3.5.1.0" ) {
    # App detected
	
	# Write the version to STDOUT by default
	# Write anything to StdOut and exit 0 to make application show as detected
    Write-Host "$FileVersion"
    exit 0
}
else {
    # App NOT detected
	
	#Exit with non-zero failure code
    exit 1
}
