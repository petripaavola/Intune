#!/bin/bash

# Intune Linux Custom Compliance script
# This script runs actual compliance check as Powershell script
# which is included in this script
#
# Compliance checks:
#	Check Powershell is installed
#	Reboot Required check (file should not exist /var/run/reboot-required)
#   Check MS Edge is installed (file should exist /opt/microsoft/msedge/msedge
#	Check MS Edge version
#	Check Powershell version
#	Check Kernel version
#	Check Kernel patch level
#	Check Kernel flavour
#	Check Kernel tainted state
#	Check SecureBoot status
#   Check sysctrl values (placeholder to check any value)
#		user.max_user_namespaces
#	Check gsettings values  (placeholder to check any value)
#		org.gnome.desktop.screensaver lock-enabled
#		org.gnome.desktop.screensaver idle-activation-enabled
#		org.gnome.desktop.session idle-delay
#	Check Defender for Endpoint on Linux status
#		MicrosoftDefenderForEndpointOnLinux_Installed
#		MicrosoftDefenderForEndpointOnLinux_RegisteredToOrganization
#		MicrosoftDefenderForEndpointOnLinux_Healthy
#		MicrosoftDefenderForEndpointOnLinux_DefinitionsStatus_up_to_date
#		MicrosoftDefenderForEndpointOnLinux_real_time_protection_enabled
#
#
# Script creates 2 log files for debugging (which you may want to disable in production?)
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
# 15.3.2023
#
# https://github.com/petripaavola/Intune/tree/master/Linux


Version=3.0
BashLogFilePath=/tmp/IntuneCustomComplianceScript_Bash.log

# Create log for debugging
echo "Intune Linux Custom Compliance bash script version $Version" > $BashLogFilePath
echo "Runtime: $(/bin/date +%Y-%m-%d_%H:%M:%S)" >> $BashLogFilePath
echo "Running script: $0" >> $BashLogFilePath
echo "Running as user: $( /bin/whoami; )" >> $BashLogFilePath
echo "Working directory: $( /bin/pwd; )" >> $BashLogFilePath

# Create Powershell commands string
PowerShellCommandsString='
	# Note: Do not use single quotes in this Powershell script
	# because it will break the bash script
	#
	# Note: Start-Transcript breaks the script
	# Note: using Write-Host causes unnecessary data returned to bash script and may break the script
	# Note: using Write-Output breaks the script

	$ScriptStartTime = Get-Date

	# Create hash for JSON
	# We will add values
	$hash = @{}

	# Add PowershellInstalled=$True
	$hash.Add("PowershellInstalled", $True)

	# Create logData array
	$logData = @()
	$logData += "###################################################"
	$logData += "Starting Intune Linux Custom Compliance script Powershell`n"

	$RunTime = Get-Date -Format yyyyMMdd-HHmmss
	$logData += "RunTime = $RunTime"

	$Whoami = whoami
	$logData += "Whoami = $Whoami"
	
	$CurrentPath = Get-Location | Select-Object -ExpandProperty Path
	$logData += "CurrentPath = $CurrentPath"

	
	####################################################################################
	# Check reboot pending

	$logData += "###################################################"
	$logData += "Run Reboot Pending compliance check`n"
	$logData += "file should not exist /var/run/reboot-required"
	if(Test-Path /var/run/reboot-required) {
		$RebootRequired = $True
	} else {
		$RebootRequired = $False
	}
	$logData += "RebootRequired=$RebootRequired"
	$hash.Add("RebootRequired", $RebootRequired)


	####################################################################################
	# Check MSEdge is installed
	
	$logData += "###################################################"
	$logData += "Run Microsoft Edge is installed check`n"
	$logData += "file should exist /opt/microsoft/msedge/msedge"
	if(Test-Path /opt/microsoft/msedge/msedge) {
		$MSEdgeInstalled = $True
	} else {
		$MSEdgeInstalled = $False
	}
	$logData += "MSEdgeInstalled=$MSEdgeInstalled"
	$hash.Add("MSEdgeInstalled", $MSEdgeInstalled)


	####################################################################################
	# Check MSEdge version
	
	$logData += "###################################################"
	$logData += "Run Microsoft Edge version check`n"
	$logData += "with command /opt/microsoft/msedge/msedge --version"

	$MSEdgeVersionString = /opt/microsoft/msedge/msedge --version
	# Example string which can be tested with regex101.com
	# Microsoft Edge 109.0.1518.70 unknown
	$regex = "^.* ([0-9]{1,}\.[0-9]{1,}\.[0-9]{1,}\.[0-9]{1,}).*"
	if($MSEdgeVersionString -match $regex) {
		$MSEdgeVersion = $Matches[1]
	} else {
		# Regex did not find MSEdge version
		$MSEdgeVersion = "0.0.0.0"
	}

	$logData += "MSEdgeVersion=$MSEdgeVersion"
	$hash.Add("MSEdgeVersion", $MSEdgeVersion)


	####################################################################################
	# Check Powershell version
	
	$logData += "###################################################"
	$logData += "Run Powershell version check`n"
	[String]$PSVersion = $PSVersionTable | Select-Object -ExpandProperty PSVersion
	$logData += "PSVersion=$PSVersion"
	$hash.Add("PSVersion", $PSVersion)


	####################################################################################
	# Check Kernel version, patch level and flavour
	
	$logData += "###################################################"
	$logData += "Run Kernel version check`n"
	# Example kernel version: 5.15.0-52-generic
	# Kernel version is: 5.15.0
	# Patch level is: 52
	#  Linux distro/kernel specific additional info: generic

	# Get kernel information
	$KernelVersionUname = & /bin/uname -r
	
	# Extract version info to own variables
	$KernelVersion =    ($KernelVersionUname.Split("-"))[0]
	$logData += "KernelVersion=$KernelVersion"
	$hash.Add("KernelVersion", $KernelVersion)

	$KernelPatchLevel = ($KernelVersionUname.Split("-"))[1]
	# Add .0 to kernel patch level version 52 -> 52.0
	# Otherwise our compliance JSON will no work correctly for Version comparison
	$KernelPatchLevel = "$($KernelPatchLevel).0"
	$logData += "KernelPatchLevel=$KernelPatchLevel"
	$hash.Add("KernelPatchLevel", $KernelPatchLevel)

	$KernelFlavour = ($KernelVersionUname.Split("-"))[2]
	$logData += "KernelFlavour=$KernelFlavour"
	$hash.Add("KernelFlavour", $KernelFlavour)


	####################################################################################
	# Check Kernel tainted state
	#
	#   There is a lot going on with tainted kernel situation
	#   Kernel could be in tainted mode intentionally
	#	when using for example proprietary NVIDIA or AMD graphics drivers
	#   But many times tainted kernel state is causing bigger problems
	#   https://www.kernel.org/doc/html/latest/admin-guide/tainted-kernels.html#decoding-tainted-state-at-runtime
	#   Fix for tainted kernel is to reboot the device
	# 	Value 0 is normal when kernel is NOT tainted

	$logData += "###################################################"
	$logData += "Run Kernel tainted state check`n"

	[int]$TaintedKernelState = Get-Content /proc/sys/kernel/tainted
	$logData += "/proc/sys/kernel/tainted"
	$logData += "$TaintedKernelState"

	$logData += "TaintedKernelState = $TaintedKernelState"
	$hash.Add("TaintedKernelState", $TaintedKernelState)


	####################################################################################
	# Check SecureBoot status
	
	$logData += "###################################################"
	$logData += "Check SecureBoot status`n"
	
	$SecureBootStatus = & /bin/mokutil --sb-state
	$logData += "/bin/mokutil --sb-state result"
	$logData += "$SecureBootStatus"

	# SecureBoot status can be:
	# 	SecureBoot enabled
	# 	SecureBoot disabled
	if($SecureBootStatus -eq "SecureBoot enabled") {
		$SecureBootEnabled = $True
	} else {
		$SecureBootEnabled = $False
	}
	$logData += "SecureBootEnabled = $SecureBootEnabled"
	$hash.Add("SecureBootEnabled", $SecureBootEnabled)


	####################################################################################
	# Get sysctl values
	#	This is another pandoras box of values
	#	Either disable this step, include only necessary values or pass on all the values
	#	If you pass all the values then you can later deside what values to check in JSON
	#   Note: passing all values might slow Compliance check

	$logData += "###################################################"
	$logData += "Get sysctl values and convert to hashtable`n"
	
	# Get sysctrl values
	$sysctlArray = & /sbin/sysctl -a

	$sysctlValuesAddedToHash = 0


       #	
       #
     # # #
      ###
       #

	# Process only specified sysctl settings
	$sysctlArray = $sysctlArray | Where-Object {
		$_ -like "user.max_user_namespaces*" -or
		$_ -like "net.ipv4.conf.all.accept_redirects*"
	}

       #
      ###
     # # #
       #
	   #


	# Process all sysctl values except ignore specified values
	# Ignore dev.cdrom.info -values because there are many values with overlapping name
	#$sysctlArray = $sysctlArray | Where-Object { $_ -notlike "dev.cdrom.info*"}

	
	# Create hash table from values
	foreach($sysctl in $sysctlArray) {
		$sysctlValueName = $null
		$sysctlValue = $null
		$Value = $null

		# Validate and extract data with regex
		# You can test and validate regex in https://regex101.com
		# String:  user.max_user_namespaces = 15233
		# Regex:  ^(.*) = (.*)$
		
		$regex = "^(.*) = (.*)$"
		$Success = $sysctl -match $regex
		
		# Continue only if regex was match
		if($Success) {
			# Extract values
			$sysctlValueName = "sysctl $($Matches[1])"
			$sysctlValue = ($Matches[2]).Trim()
			
			if($sysctlValue) {
				# We have name and value

				# Trim empty spaces
				$sysctlValue = $sysctlValue.Trim()

				$logData += "`$sysctlValue=$sysctlValue"
				$logData += "GetType `$sysctlValue=$($sysctlValue.GetType())"

				# Check if value is integer
				if(($sysctlValue -is [Int32]) -or ($sysctlValue -is [Int64]) -or ($sysctlValue -is [uint32])) {
					$Value = [int64]$sysctlValue
					$logData += "`$Value=$Value is integer"
				} else {
					# Cast value as String
					$Value = [String]$sysctlValue
					$logData += "`$Value=$Value is string"
				}
				
			} else {
				# We have name but value is null
				$Value = $null
			}

			# Add hash key/value if it does not already exist
			# This prevents us getting unnecessary errors which could fail this bash script
			# and help debugging possible problems
			if(-not ($hash.ContainsKey($sysctlValueName))) {
				$hash.Add($sysctlValueName, $Value)
				
				$sysctlValuesAddedToHash++
			} else {
				$logData += "Skipped duplicate hash valueName: $sysctlValueName"
			}
			
		} else {
			# sysctl line did not match our regex
			# Print value to log for debugging
			$logData += "Did not match regex: $sysctl"
		}
	}

	$logData += "Added $sysctlValuesAddedToHash sysctl values to hash"


    ####################################################################################
	# Get gsettings

	# Configure below gsettings to be included in compliance check

       #	
       #
     # # #
      ###
       #

	# Add only selected settings to $hash
	$SelectedGsettingsArray = @(
		"gsettings get org.gnome.desktop.screensaver lock-enabled",`
		"gsettings get org.gnome.desktop.screensaver idle-activation-enabled",`
		"gsettings get org.gnome.desktop.session idle-delay"
	)

       #
      ###
     # # #
       #
	   #
	
	$logData += "###################################################"
	$logData += "Get gsettings`n"

	### function start ###
	function Add_gsettingsArrayToHash {
		Param(
			[Parameter(Mandatory=$true,
			ValueFromPipeline=$true,
			Position=0)]
			$gsettingsArray
		)

		$gsettingsHash = @{}

		foreach($gsetting in $gsettingsArray) {
			$gsettingName = $null
			$gsettingValue = $null
			$Value = $null

			# gsettings are either boolean, uint32, integer or string types (all else are strings)
			#	org.gnome.desktop.screensaver idle-activation-enabled true
			# 	org.gnome.desktop.screensaver lock-delay uint32 0
			#	org.gnome.desktop.screensaver picture-opacity 100
			#	org.gnome.desktop.screensaver picture-options singlequotezoomsinglequote
			#	org.gnome.desktop.search-providers disabled $as []


			# 	org.gnome.desktop.screensaver lock-delay uint32 0
			#   Get uint32 regex match lines
			if($gsetting -match "^(.*) (.*) uint32 (.*)$") {
				# setting is uint32 type
				$gsettingName = "gsettings get $($matches[1]) $($matches[2])"
				$gsettingValue = [uint32]$matches[3]

			} elseif ($gsetting -match "^(.*?) (.*?) (.*)$") {
				# Using ? in first 2 groups to get lazy regex
				# so it would match lines ending with (includes additional space): @a{ss} {}
				# org.gnome.desktop.search-providers disabled $as []
				
				# There was no uint32 type value

				# setting is boolean, integer or string type (we ignore other variable types)
				$gsettingName = "gsettings get $($matches[1]) $($matches[2])"
				$Value = $matches[3]
				
				if(($Value -is [int32]) -or ($Value -is [uint32]) -or ($Value -is [int64])) {

					# $Value is integer
					$gsettingValue = [int64]$Value
				
				} elseif(($Value -eq "true") -or ($Value -eq "false")) {

					# value is boolean
					if($Value -eq "true") {
						$gsettingValue = $True
					} else {
						$gsettingValue = $False
					}
					
				} else {
					# Save Value as string
					$gsettingValue = [string]$Value
				}
			
			} else {
				# Regex did not match anything
				$logData += "regex did not catch this: $gsetting"

			}

			# Add gsetting to hash if we have Hash Key-value -> our regex worked
			if($gsettingName) {	
				if(-not ($gsettingsHash.ContainsKey($gsettingName))) {
					$gsettingsHash.Add($gsettingName, $gsettingValue)
				} else {
					$logData += "Skipped duplicate hash Key: $gsettingName"
				}
			} else {
				$logData += "Regex did not catch gsettings line. Fix regex in this script! $gsettingName"
			}
		}
		
		return $gsettingsHash
	}
	### function end ###


	# Get all gsettings to array
	$logData += "Get all gsettings"
	$gsettingsArray = & /bin/gsettings list-recursively

	# Parse gsettingsArray results and return hashtable
	$logData += "Parse gsettingsArray results and return hashtable"
	$gsettingsHashtable = Add_gsettingsArrayToHash $gsettingsArray


	# $SelectedGsettingsArray this array was manually specified in the beginning of this block
	# Add specified gsettings to $hash
	foreach($gsetting in $SelectedGsettingsArray) {
		# Make sure hash Key exists in $gsettingsHashtable before adding it to $hash
		if($gsettingsHashtable.ContainsKey($gsetting)) {
			if(-not ($hash.ContainsKey($gsetting))) {
				$logData += "Add to `$hash: Key: $gsetting Value: $($gsettingsHashtable[$gsetting])"
				$hash.Add($gsetting, $gsettingsHashtable[$gsetting])
			} else {
				$logData += "Skipped duplicate hash Name: $gsetting"
				$logData += "This should not be possible case normally"
			}
		} else {
			$logData += "`$gsettingsHashtable Key not found: $gsetting"
			$logData += "usually this means there is typo in `$SelectedGsettingsArray gsettings entries"
		}
	}


	<#
	# Do not do below because JSON will get BIG.
	# Left code here just in case needed in the future
	#
	# Add all $gsettingsHashtable hash values to $hash which we will convert to JSON in the end
	foreach($key in $gsettingsHashtable.keys) {
		if(-not ($hash.ContainsKey($key))) {
			$hash.Add($key, $gsettingsHashtable[$key])
		} else {
			$logData += "Skipped duplicate hash Key: $($key)"
			$logData += "This should not be possible case normally"
		}
	}
	#>


	####################################################################################
	# Check Microsoft Defender Endpoint for Linux status
	
	$logData += "###################################################"
	$logData += "Check Microsoft Defender Endpoint for Linux status`n"

	# Check mdatp is installed in default path
	$mdatpPath = "/opt/microsoft/mdatp/sbin/wdavdaemonclient"
	if(Test-Path $mdatpPath) {
		$MicrosoftDefenderForEndpointOnLinux_Installed = $True

		# Check Defender is registered to organization
		$mdatp_org_id = & $mdatpPath --field org_id

		# Example value includes also double quotes
		# "0899fda5-fcad-42ef-9325-f8a0c7e30560"
		$regex = "^?[a-f0-9]{8}-([a-f0-9]{4}-){3}[a-f0-9]{12}?$"
		if($mdatp_org_id -match $regex) {
			# Found valid guid
			$MicrosoftDefenderForEndpointOnLinux_RegisteredToOrganization = $True
		} else {
			# Did not find valid guid
			$MicrosoftDefenderForEndpointOnLinux_RegisteredToOrganization = $True
		}

		# Check Defender is functioning properly (is healthy)
		$mdatp_healthy = & $mdatpPath health --field healthy
		
		if($mdatp_healthy -eq "true") {
			# Defender is healthy
			$MicrosoftDefenderForEndpointOnLinux_Healthy = $True
		} else {
			# Defender is NOT healthy
			$MicrosoftDefenderForEndpointOnLinux_Healthy = $False
		}
		
		# Check Defender definitions are up_to_date
		$mdatp_definitions_status = & $mdatpPath health --field definitions_status
		
		if($mdatp_definitions_status -eq "`"up_to_date`"") {
			# Definitions are up to date
			$MicrosoftDefenderForEndpointOnLinux_DefinitionsStatus_up_to_date = $True
		} else {
			# Definitions are NOT up to date
			$MicrosoftDefenderForEndpointOnLinux_DefinitionsStatus_up_to_date = $False
		}

		# Check Defender real_time_protection_enabled
		$mdatp_real_time_protection_enabled = & $mdatpPath health --field real_time_protection_enabled
		
		if($mdatp_real_time_protection_enabled -eq "true") {
			# Defender Realtime Protection is enabled
			$MicrosoftDefenderForEndpointOnLinux_real_time_protection_enabled = $True
		} else {
			# Defender Realtime Protection is NOT enabled
			$MicrosoftDefenderForEndpointOnLinux_real_time_protection_enabled = $False
		}
	
	} else {
		# Microsoft Defender Endpoint for Linux status is NOT installed

		$MicrosoftDefenderForEndpointOnLinux_Installed = $False
		$MicrosoftDefenderForEndpointOnLinux_RegisteredToOrganization = $False
		$MicrosoftDefenderForEndpointOnLinux_Healthy = $False
		$MicrosoftDefenderForEndpointOnLinux_DefinitionsStatus_up_to_date = $False
		$MicrosoftDefenderForEndpointOnLinux_real_time_protection_enabled = $False
		
	}

	# Add Defender information to $hash
	
	$logData += "MicrosoftDefenderForEndpointOnLinux_Installed=$MicrosoftDefenderForEndpointOnLinux_Installed"
	$hash.Add("MicrosoftDefenderForEndpointOnLinux_Installed", $MicrosoftDefenderForEndpointOnLinux_Installed)
	
	$logData += "MicrosoftDefenderForEndpointOnLinux_RegisteredToOrganization=$MicrosoftDefenderForEndpointOnLinux_RegisteredToOrganization"
	$hash.Add("MicrosoftDefenderForEndpointOnLinux_RegisteredToOrganization", $MicrosoftDefenderForEndpointOnLinux_RegisteredToOrganization)
	
	$logData += "MicrosoftDefenderForEndpointOnLinux_Healthy=$MicrosoftDefenderForEndpointOnLinux_Healthy"
	$hash.Add("MicrosoftDefenderForEndpointOnLinux_Healthy", $MicrosoftDefenderForEndpointOnLinux_Healthy)

	$logData += "MicrosoftDefenderForEndpointOnLinux_DefinitionsStatus_up_to_date=$MicrosoftDefenderForEndpointOnLinux_DefinitionsStatus_up_to_date"
	$hash.Add("MicrosoftDefenderForEndpointOnLinux_DefinitionsStatus_up_to_date", $MicrosoftDefenderForEndpointOnLinux_DefinitionsStatus_up_to_date)

	$logData += "MicrosoftDefenderForEndpointOnLinux_real_time_protection_enabled=$MicrosoftDefenderForEndpointOnLinux_real_time_protection_enabled"
	$hash.Add("MicrosoftDefenderForEndpointOnLinux_Real_time_protection_enabled", $MicrosoftDefenderForEndpointOnLinux_real_time_protection_enabled)


	$logData += "###################################################"

	# Convert hashtable to JSON
	$returnJson = $hash | ConvertTo-Json -Compress

	$logData += "Custom Compliance returnJson:`n"
	$logData += "$returnJson"
	$logData += ""

	$logData += "###################################################"
	$logData += "Powershell script end`n"

	# Calculate script runtime
	$ScriptEndTime = Get-Date
	$ScriptRunTime = New-Timespan $ScriptStartTime $ScriptEndTime
	$logData += "Script runtime $($ScriptRunTime.Minutes) min $($ScriptRunTime.Seconds) sec $($ScriptRunTime.Milliseconds) millisec"

	$logData += "###################################################"
	$logData += "Possible error data below`n"

	# Export log array to log file
	$logData | Out-File -FilePath /tmp/IntuneCustomComplianceScript_Powershell.log -Force

	# Export error data to help debugging errors in Powershell
	# Of course we will not make any errors but just in case :)
	if($Error) {
		$Error | Out-File -FilePath /tmp/IntuneCustomComplianceScript_Powershell.log -Append
	}
	
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
	JSON=$($PWSHPATH -Command "& { ${PowerShellCommandsString} }")
	ExitCode=$?
	echo "Powershell script run exitcode: $ExitCode" >> $BashLogFilePath
	echo "JSON From Powershell" >> $BashLogFilePath
	echo "" >> $BashLogFilePath
	echo "$JSON" >> $BashLogFilePath
	echo "" >> $BashLogFilePath
	echo "Bash script end" >> $BashLogFilePath
	echo "$JSON"
else
	# pwsh (Powershell) not found
	# return hardcoded json which will make device Not Compliant

	echo "pwsh not found from path: $PWSHPATH" >> $BashLogFilePath
	echo "Compliance Check will fail" >> $BashLogFilePath
	JSON='{"PowershellInstalled":false}'
	echo "Return json:" >> $BashLogFilePath
	echo "$JSON" >> $BashLogFilePath
	echo "Bash script end" >> $BashLogFilePath
	echo "$JSON"
fi
