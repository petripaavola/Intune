# Install Audacity + MP3 Lame encoder
# IntuneWin32 EXE Powershell script install example
#
#
# Petri.Paavola@yodamiitti.fi
# Windows MVP - Windows and Intune
# 2024-11-01
#
# Original script source:
# https://github.com/petripaavola/Intune/blob/master/Apps/install_Audacity_3.5.1-EXE-example.ps1
#
# Intune install command:
# Powershell.exe -ExecutionPolicy Bypass -File install_Audacity_3.5.1-EXE-example.ps1


    #
    #
  # # #
   # #
    #

# Configure these values
$SoftwareName = 'Audacity'
$version = "3.5.1"

$InstallerFilePath = "$PSScriptRoot\audacity-win-3.5.1-64bit.exe"
$Arguments = @( 
    '/verysilent'
    '/norestart'
)

    #
   # #
  # # #
    #
    #



# Start script in 64bit environment if script was started in 32bit environment
# Intune Win32 application install process starts in 32bit process by default (2024-11-01)
# Few commands will require native bit command to work. Examples pnputil.exe and dism.exe
# Below 64bit workaround is provided by Oliver Kieselbach
# Original example:
# https://github.com/okieselbach/Intune/blob/master/ManagementExtension-Samples/IntunePSTemplate.ps1
if (-not [System.Environment]::Is64BitProcess)
{
     # start new PowerShell as x64 bit process, wait for it and gather exit code and standard error output
    $sysNativePowerShell = "$($PSHOME.ToLower().Replace("syswow64", "sysnative"))\powershell.exe"

    $pinfo = New-Object System.Diagnostics.ProcessStartInfo
    $pinfo.FileName = $sysNativePowerShell
    $pinfo.Arguments = "-ex bypass -file `"$PSCommandPath`""
    $pinfo.RedirectStandardError = $true
    $pinfo.RedirectStandardOutput = $true
    $pinfo.CreateNoWindow = $true
    $pinfo.UseShellExecute = $false
    $p = New-Object System.Diagnostics.Process
    $p.StartInfo = $pinfo
    $p.Start() | Out-Null
	$p.WaitForExit()  # Wait for the 64-bit process to complete

    $exitCode = $p.ExitCode

    $stderr = $p.StandardError.ReadToEnd()

    if ($stderr) { Write-Error -Message $stderr }
	
	exit $exitCode

} else {

	#############################################################
	# Main script starts here

	# Start logging
	Start-Transcript "C:\Windows\Logs\Install_$($SoftwareName)_$($version)_IntuneWin.log" -Append

	# Running inside Try so we can catch fatal errors
	try {

		Write-Host "Install $SoftwareName $version"


		# Check that $InstallerFilePath exists
		if (-not (Test-Path $InstallerFilePath)) {
			Write-Host "Installer file not found: $InstallerFilePath"
			Write-Host "Exiting installation with error code 9999"
			Stop-Transcript
			Exit 9999
		}


		# Log file path for stdOut and stdError
		$RedirectStandardOutputFilePath = "C:\Windows\Logs\Install_$($SoftwareName)_$($version)_IntuneWin_RedirectStandardOutput.log"
		$RedirectStandardErrorFilePath = "C:\Windows\Logs\Install_$($SoftwareName)_$($version)_IntuneWin_RedirectStandardError.log"


		# We use Start-Process which gets most probability the ExitCode
		# Some really really rare cases $LastExitCode does not work, so this works better
		$Process = Start-Process $InstallerFilePath -ArgumentList $Arguments -Wait -Passthru -NoNewWindow -RedirectStandardOutput $RedirectStandardOutputFilePath -RedirectStandardError $RedirectStandardErrorFilePath
		$Process.WaitForExit()
		$ExitCode = $Process.ExitCode

		if($ExitCode -eq 0) {
			# Installation was successful
			
			Write-Host "$SoftwareName $version installation succeeded with ExitCode $ExitCode"
		} else {
			# Installation failed
			
			Write-Host "Error installing $SoftwareName $version. ExitCode: $ExitCode"
			Write-Host "Exiting installation with ExitCode $ExitCode"
			Stop-Transcript
			Exit $ExitCode
		}

		# Remove Desktop shortcut
		# This is best try so app installation will not fail if this for any reason did not succeed
		$DesktopShortcutPath = "C:\Users\Public\Desktop\Audacity.lnk"
		if(Test-Path $DesktopShortcutPath) {
			Write-Host "Delete desktop shortcut: $DesktopShortcutPath"
			Remove-Item $DesktopShortcutPath -Force -ErrorAction SilentlyContinue
			Write-Host "Success: $?"
		}


		Write-Host "Copy lame_enc.dll to Audacity installation folder"
		Copy-Item "$PSScriptRoot\lame_enc.dll" -Destination "C:\Program Files\Audacity" -Force
		$Success = $?

		if($Success) {
			Write-Host "Success copying file."

			Write-Host "$SoftwareName $version installation succeeded"
			Write-Host "Exiting installation with ExitCode 0"
			Stop-Transcript
			Exit 0
		} else {
			Write-Host "Error copying lame_enc.dll file"
			Write-Host "Exiting installation with ExitCode 2"
			Stop-Transcript
			Exit 2
		}

	} catch {
		# Handle any unexpected errors

		Write-Host "An unexpected fatal error occurred: $_"
		Write-Host "Exiting installation with ExitCode 99999"
		Stop-Transcript
		Exit 99999
		
	}
}