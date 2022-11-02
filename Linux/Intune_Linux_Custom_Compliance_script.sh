#!/bin/bash

# Intune Linux Custom Compliance script
# This script runs actual compliance check as Powershell script
# which is included in this script
#
# Compliance checks:
#	Reboot Required check (file should not exist /var/run/reboot-required)
#
#
# Script creates 2 log files for debugging
# 	/tmp/IntuneCustomComplianceScript_Bash.log
# 	/tmp/IntuneCustomComplianceScript_Powershell.log
#
# Note: Make sure you have installed Powershell before running this script
#
# Note: make sure you are NOT putting anything
# to STDOUT because that breaks the script
# Powershell script returns compressed JSON to STDOUT
# which is passed to Intune for compliance validation check
#
#
# Petri.Paavola@yodamiitti.fi
# Microsoft MVP - Windows and Devices for IT
# 2.11.2022
#
# https://github.com/petripaavola/Intune/tree/master/Linux


Version=1.60
BashLogFilePath=/tmp/IntuneCustomComplianceScript_Bash.log

# Create log for debugging
echo "Intune Linux Custom Compliance bash script version $Version" > $BashLogFilePath
echo "Runtime: $(date +%Y-%m-%d_%H:%M:%S)" >> $BashLogFilePath
echo "Running script: $0" >> $BashLogFilePath
echo "Running as user: $( whoami; )" >> $BashLogFilePath
echo "The present working directory is $( pwd; )" >> $BashLogFilePath

# Create Powershell commands string
PowerShellCommandsString='
	# Start Powershell logging
	# Note: Start-Transcript breaks the script
	# Note: using Write-Host breaks the script
	# Note: using Write-Output breaks the script
	
	$logData = @()
	$logData += "###################################################"
	$logData += "Starting Intune Linux Custom Compliance script Powershell part"

	$RunTime = Get-Date -Format yyyyMMdd-HHmmss
	$logData += "RunTime = $RunTime"

	$Whoami = whoami
	$logData += "Whoami = $Whoami"
	
	$CurrentPath = Get-Location | Select-Object -ExpandProperty Path
	$logData += "CurrentPath = $CurrentPath"

	# Check reboot pending
	$logData += "###################################################"
	$logData += "Run Reboot Pending compliance check"
	$logData += "file should not exist /var/run/reboot-required"
	if(Test-Path /var/run/reboot-required) {
		$RebootRequired = $True
	} else {
		$RebootRequired = $False
	}
	$logData += "RebootRequired=$RebootRequired"
	$logData += "###################################################"

	# Create hash table which we convert to compressed JSON
	$hash = @{ RebootRequired = $RebootRequired }
	$returnJson = $hash | ConvertTo-Json -Compress

	$logData += "Custom Compliance returnJson = $returnJson"
	$logData += "Powershell script end"
	$logData | Out-File -FilePath /tmp/IntuneCustomComplianceScript_Powershell.log -Force

	# Return compressed JSON
	return $returnJson
'

echo "####################################################" >> $BashLogFilePath
echo "Running Powershell script: $PowerShellCommandsString" >> $BashLogFilePath
echo "####################################################" >> $BashLogFilePath

# Run Powershell Commands
pwsh -Command "& { ${PowerShellCommandsString} }"
ExitCode=$?

echo "Powershell script run exitcode: $ExitCode" >> $BashLogFilePath
echo "Bash script end" >> $BashLogFilePath
