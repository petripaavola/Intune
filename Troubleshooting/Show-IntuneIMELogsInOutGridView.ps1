# Show Intune Management Extension (IME) logs in Out-GridView
#
# Petri.Paavola@yodamiitti.fi
# Microsof MVP - Windows and Devices for IT
# 28.2.2023
#
# https://github.com/petripaavola/Intune/tree/master/Troubleshooting
#
# ver 1.0


[CmdletBinding()]
Param(
	$LogFilePath=$null
)

Write-Host "Starting IME Log Out-GridView Tool`n"

# If LogFilePath is not specified then show log files
# from folder C:\ProgramData\Microsoft\intunemanagementextension\Logs
if(-not $LogFilePath) {
	$LogFiles = Get-ChildItem -Path 'C:\ProgramData\Microsoft\intunemanagementextension\Logs' -Filter *.log
	
	# Show log files in Out-GridView
	$SelectedLogFile = $LogFiles | Out-GridView -Title 'Select log file to show in Out-GridView' -OutputMode Single
	
	if($SelectedLogFile) {
		$LogFilePath = $SelectedLogFile.FullName
	} else {
		Write-Host "No log file selected. Script will exit!`n" -ForegroundColor Yellow
		Exit 0
	}
}


$LineNumber=1

# Create Generic list where log entry custom objects are added
$LogEntryList = [System.Collections.Generic.List[PSObject]]@()

$Log = Get-Content -Path $LogFilePath

# This matches for cmtrace type logs
# Test with https://regex101.com
# String: <![LOG[ExecutorLog AgentExecutor gets invoked]LOG]!><time="21:38:06.3814532" date="2-14-2023" component="AgentExecutor" context="" type="1" thread="1" file="">

$regex = '^\<\!\[LOG\[(.*)]LOG\].*\<time="([0-9]{1,2}):([0-9]{1,2}):([0-9]{1,2}).([0-9]{1,})".*date="([0-9]{1,2})-([0-9]{1,2})-([0-9]{4})" component="(.*?)" context="(.*?)" type="(.*?)" thread="(.*?)" file="(.*?)">$'

Foreach ($CurrentLogEntry in $Log) {

	# Get data from CurrentLogEntry
	if($CurrentLogEntry -Match $regex) {
		# Regex found match
		$LogMessage = $Matches[1].Trim()
		
		$Hour = $Matches[2]
		$Minute = $Matches[3]
		$Second = $Matches[4]
		
		$MilliSecond = $Matches[5]
		# Cut milliseconds to 0-999
		# Time unit is so small that we don't even bother to round the value
		$MilliSecond = $MilliSecond.Substring(0,3)
		
		$Month = $Matches[6]
		$Day = $Matches[7]
		$Year = $Matches[8]
		
		$Component = $Matches[9]
		$Context = $Matches[10]
		$Type = $Matches[11]
		$Thread = $Matches[12]
		$File = $Matches[13]

		$Param = @{
			Hour=$Hour
			Minute=$Minute
			Second=$Second
			MilliSecond=$MilliSecond
			Year=$Year
			Month=$Month
			Day=$Day
		}

		#$LogEntryDateTime = Get-Date @Param
		#Write-Host "DEBUG `$LogEntryDateTime: $LogEntryDateTime"

		# This works for humans but does not sort
		#$DateTimeToLogFile = "$($Hour):$($Minute):$($Second).$MilliSecond $Day/$Month/$Year"

		# This sorts right way
		$DateTimeToLogFile = "$Year-$Month-$Day $($Hour):$($Minute):$($Second).$MilliSecond"

		# Create Powershell custom object and add it to list
		$LogEntryList.add([PSCustomObject]@{
			'Line' = $LineNumber;
			'DateTime' = $DateTimeToLogFile;
			'Message' = $LogMessage;
			'Component' = $Component;
			'Context' = $Context;
			'Type' = $Type;
			'Thread' = $Thread;
			'File' = $File
			})

	}
	$LineNumber++
}
	
$LogEntryList | Out-GridView -Title "Intune IME Log Viewer $LogFilePath"

Write-Host "Script end"
