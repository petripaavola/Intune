#!/bin/bash

# Intune Linux Custom Compliance script
# This script runs actual compliance check as Powershell script
# which is included in this script
#
# Compliance checks:
#	Check Powershell is installed
#	Check Powershell version
#	Reboot Required check (file should not exist /var/run/reboot-required)
#   	Check MS Edge is installed (file should exist /opt/microsoft/msedge/msedge
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


Version=2.25
BashLogFilePath=/tmp/IntuneCustomComplianceScript_Bash.log

# Create log for debugging
echo "Intune Linux Custom Compliance bash script version $Version" > $BashLogFilePath
echo "Runtime: $(/bin/date +%Y-%m-%d_%H:%M:%S)" >> $BashLogFilePath
echo "Running script: $0" >> $BashLogFilePath
echo "Running as user: $( /bin/whoami; )" >> $BashLogFilePath
echo "Working directory: $( /bin/pwd; )" >> $BashLogFilePath

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

	# Check MSEdge is installed
	$logData += "###################################################"
	$logData += "Run Microsoft Edge is installed check"
	$logData += "file should exist /opt/microsoft/msedge/msedge"
	if(Test-Path /opt/microsoft/msedge/msedge) {
		$MSEdgeInstalled = $True
	} else {
		$MSEdgeInstalled = $False
	}
	$logData += "MSEdgeInstalled=$MSEdgeInstalled"

	# Check Powershell version
	$logData += "###################################################"
	$logData += "Run Powershell version check"
	[String]$PSVersion = $PSVersionTable | Select-Object -ExpandProperty PSVersion
	$logData += "PSVersion=$PSVersion"

	$logData += "###################################################"
	# Create hash table which we convert to compressed JSON
	$hash = @{
		RebootRequired = $RebootRequired
		MSEdgeInstalled = $MSEdgeInstalled
		PSVersion = $PSVersion
		PowershellInstalled = $True
	}
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

# Check that pwsh exist
PWSHPATH=/opt/microsoft/powershell/7/pwsh
if [ -f "$PWSHPATH" ]; then
	# Powershell found
	# Run Powershell Commands
	
	echo "Run Powershell script part" >> $BashLogFilePath
	$PWSHPATH -Command "& { ${PowerShellCommandsString} }"
	ExitCode=$?
	echo "Powershell script run exitcode: $ExitCode" >> $BashLogFilePath
	echo "Bash script end" >> $BashLogFilePath	
else
	# pwsh not found
	# return hardcoded json which will make device Not Compliant

	echo "pwsh not found from path: $PWSHPATH" >> $BashLogFilePath
	echo "Compliance Check will fail" >> $BashLogFilePath
	JSON='{"RebootRequired":false,"MSEdgeInstalled":true,"PSVersion":"7.2.7","PowershellInstalled":false}'
	echo "Return json: $JSON" >> $BashLogFilePath
	echo "Bash script end" >> $BashLogFilePath
	echo "$JSON"
fi
