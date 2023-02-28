<#
.Synopsis
   Show Intune Management Extension (IME) logs in Out-GridView ver 1.2
   This looks a lot like cmtrace.exe tool
   
   Accepts log file as parameter.
   
   If log file is not specified then log files can be
   selected from graphical Out-GridView
   
   
   Author:
   Petri.Paavola@yodamiitti.fi
   Modern Management Principal
   Microsoft MVP - Windows and Devices for IT
   
   2023-02-28
   
   https://github.com/petripaavola/Intune/tree/master/Troubleshooting
.DESCRIPTION

.EXAMPLE
   .\Show-IntuneIMELogsInOutGridView.ps1
.EXAMPLE
   .\Show-IntuneIMELogsInOutGridView.ps1 "C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\IntuneManagementExtension.log"
.EXAMPLE
    .\Show-IntuneIMELogsInOutGridView.ps1 -LogFilePath "C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\IntuneManagementExtension.log"
.EXAMPLE
    Get-ChildItem C:\programdata\Microsoft\IntuneManagementExtension\Logs\IntuneManagementExtension.log | .\Show-IntuneIMELogsInOutGridView.ps1

.INPUTS
   Intune Management Extension (IME) log file full path
     or
   Intune Management Extension (IME) log file type object which has property FullName
.OUTPUTS
   None
.NOTES
.LINK
   https://github.com/petripaavola/Intune/tree/master/Troubleshooting
#>

[CmdletBinding()]
Param(
    [Parameter(Mandatory=$false,
				HelpMessage = 'Enter Intune IME log file path',
                ValueFromPipeline=$true,
                ValueFromPipelineByPropertyName=$true)]
	[Alias("FullName")]
    [String]$LogFilePath = $null
)


Write-Host "Starting IME Log Out-GridView Tool`n"

# If LogFilePath is not specified then show log files in Out-GridView
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

# Initialize variables
$LineNumber=1
$MultilineLogEntryStartsLineNumber=0

# Create Generic list where log entry custom objects are added
$LogEntryList = [System.Collections.Generic.List[PSObject]]@()

$Log = Get-Content -Path $LogFilePath

# This matches for cmtrace type logs
# Test with https://regex101.com
# String: <![LOG[ExecutorLog AgentExecutor gets invoked]LOG]!><time="21:38:06.3814532" date="2-14-2023" component="AgentExecutor" context="" type="1" thread="1" file="">

# This matches single line full log entry
$SingleLineRegex = '^\<\!\[LOG\[(.*)]LOG\].*\<time="([0-9]{1,2}):([0-9]{1,2}):([0-9]{1,2}).([0-9]{1,})".*date="([0-9]{1,2})-([0-9]{1,2})-([0-9]{4})" component="(.*?)" context="(.*?)" type="(.*?)" thread="(.*?)" file="(.*?)">$'

# Start of multiline log entry
$FirstLineOfMultiLineLogRegex = '^\<\!\[LOG\[(.*)$'

# End of multiline log entry
$LastLineOfMultiLineLogRegex = '^(.*)\]LOG\]\!>\<time="([0-9]{1,2}):([0-9]{1,2}):([0-9]{1,2}).([0-9]{1,})".*date="([0-9]{1,2})-([0-9]{1,2})-([0-9]{4})" component="(.*?)" context="(.*?)" type="(.*?)" thread="(.*?)" file="(.*?)">$'


# Process each log file line one by one
Foreach ($CurrentLogEntry in $Log) {

	# Get data from CurrentLogEntry
	if($CurrentLogEntry -Match $SingleLineRegex) {
		# This matches single line full log entry
		
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
		#Write-Host "DEBUG `$LogEntryDateTime: $LogEntryDateTime" -ForegroundColor Yellow

		# This works for humans but does not sort
		#$DateTimeToLogFile = "$($Hour):$($Minute):$($Second).$MilliSecond $Day/$Month/$Year"

		# This does sorting right way
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

	} elseif ($CurrentLogEntry -Match $FirstLineOfMultiLineLogRegex) {
		# Single line regex did not get results so we are dealing multiline case separately here
		# Test if this is start of multiline log entry

		#Write-Host "DEBUB Start of multiline regex: $CurrentLogEntry" -ForegroundColor Yellow

		$MultilineLogEntryStartsLineNumber = $LineNumber

		# Regex found match
		$LogMessage = $Matches[1].Trim()
		
		$DateTimeToLogFile = ''
		$Component = ''
		$Context = ''
		$Type = ''
		$Thread = ''
		$File = ''

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

	} elseif ($CurrentLogEntry -Match $LastLineOfMultiLineLogRegex) {
		# Single line regex did not get results so we are dealing multiline case separately here
		# Test if this is end of multiline log entry

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
		#Write-Host "DEBUG `$LogEntryDateTime: $LogEntryDateTime" -ForegroundColor Yellow

		# This works for humans but does not sort
		#$DateTimeToLogFile = "$($Hour):$($Minute):$($Second).$MilliSecond $Day/$Month/$Year"

		# This does sorting right way
		$DateTimeToLogFile = "$Year-$Month-$Day $($Hour):$($Minute):$($Second).$MilliSecond"

		# Create Powershell custom object and add it to list
		$LogEntryList.add([PSCustomObject]@{
			'Line' = $LineNumber;
			'DateTime' = '';
			'Message' = $LogMessage;
			'Component' = $Component;
			'Context' = $Context;
			'Type' = $Type;
			'Thread' = $Thread;
			'File' = $File
			})
		
		# Add DateTime, Component, Context, Type, Thread, File information to object which is starting multiline log entry
		$LogEntryList[$MultilineLogEntryStartsLineNumber-1].DateTime = $DateTimeToLogFile
		$LogEntryList[$MultilineLogEntryStartsLineNumber-1].Component = $Component
		$LogEntryList[$MultilineLogEntryStartsLineNumber-1].Context = $Context
		$LogEntryList[$MultilineLogEntryStartsLineNumber-1].Type = $Type
		$LogEntryList[$MultilineLogEntryStartsLineNumber-1].Thread = $Thread
		$LogEntryList[$MultilineLogEntryStartsLineNumber-1].File = $File
		
	} else {
		# We didn't catch log entry with our regex
		# This should be multiline log entry but not first or last line in that log entry
		# This can also be some line that should be matched with (other) regex
		
		#Write-Host "DEBUG: $CurrentLogEntry"  -ForegroundColor Yellow
		
		# Create Powershell custom object and add it to list
		$LogEntryList.add([PSCustomObject]@{
			'Line' = $LineNumber;
			'DateTime' = '';
			'Message' = $CurrentLogEntry;
			'Component' = '';
			'Context' = '';
			'Type' = '';
			'Thread' = '';
			'File' = ''
			})
	}

	$LineNumber++
}
	
$LogEntryList | Out-GridView -Title "Intune IME Log Viewer $LogFilePath"

Write-Host "Script end"
