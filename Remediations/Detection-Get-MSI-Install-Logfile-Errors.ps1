# Check if MSI Application installer has failed
# and return possible error line to STDOUT
# This can be used for example with Intune Remediation script
# to get possible installation error code and explanation remotely
#
# You can create MSI installation log file with parameter /l*v
# For example: msiexec /i installer.msi /qn /l*v C:\Windows\Temp\WPNinja-Install-Application.log
#
# Run Remediation in System context and in 64-bit PowerShell
#
# Author:
# Petri.Paavola@yodamiitti.fi
# Microsoft MVP - Windows and Intune
# 2025-09-21


        #
        #
        #
      #####
       ###
        #

# Text to search from MSI installation log file
$ErrorTextToSearch = 'INSTALL ERROR:'

# MSI Application installation log file path
$MSIApplicationInstallerLogFilePath = 'C:\Windows\Temp\WPNinja-Install-Application.log'

       #
      ###
     #####
       #
       #
       #


# Read log file if it exists
if(Test-Path $MSIApplicationInstallerLogFilePath) {
	# Log file exists
	$LogFileContent = Get-Content -Path $MSIApplicationInstallerLogFilePath
} else {
	# Log file does not exists
	Write-Host "MSI Application log file does not exist: $MSIApplicationInstallerLogFilePath"
	Exit 0
}

# Try to find possible error code and text from log file
if($LogFileContent) {
	# $LogFileContent has content

	# Find MSI Application failure texts defined in variable $ErrorTextToSearch
    #
    # Cast variable as String array so we don't get any other (hidden) object properties derived from Get-Content
    # which could/would clutter possible json conversion output later
	[String[]]$MSIApplicationFailureCode = $LogFileContent | Where-Object { $_ -like "*$ErrorTextToSearch*" }
} else {
	# Could not read log file
	Write-Host "Could not read log file: $MSIApplicationInstallerLogFilePath"
	Exit 0
}

# Print to STDOUT if we have error code and text
if($MSIApplicationFailureCode) {
	# We found installation error code and text

	if($MSIApplicationFailureCode.Length -gt 1) {
		# We have more than 1 line so convert results to compressed json
		# So we can see all error code lines in Intune Remediation console
		# which supports only showing one (last) line from STDOUT
		$MSIApplicationFailureCode | ConvertTo-Json -Compress
	} else {
        # We have only one line
        # Print installation error text to STDOUT as is
        # So we can see it from Intune Remedation console remotely
        $MSIApplicationFailureCode
    }    

	# Exit with failure so we can find failed device easier in Intune Remediations view
	Exit 1
} else {
	# We did NOT find installation error code and text
	# So installation was successful
	Write-Host "Did not find installation error codes and text from log file: $MSIApplicationInstallerLogFilePath"
	Exit 0
}
