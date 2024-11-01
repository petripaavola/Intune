# Install Audacity + MP3 Lame encoder
# IntuneWin32 EXE Powershell script install example
#
#
# Petri.Paavola@yodamiitti.fi
# Windows MVP - Windows and Intune
# 1.11.2024
#
#
# Intune install command: Powershell.exe -ExecutionPolicy Bypass -File install_Audacity_3.5.1-EXE-example.ps1



$Software = 'Audacity'
$version = "3.5.1"

$Installer = 'audacity-win-3.5.1-64bit.exe'
$Arguments = @( 
    '/verysilent'
    '/norestart'
)


# Start logging
Start-Transcript "C:\Windows\Logs\Install_$($Software)_$($version)_IntuneWin.log" -Append

Write-Host "Install $Software $version"

# Log file path for stdOut and stdError
$RedirectStandardOutputFilePath = "C:\Windows\Logs\Install_$($Software)_$($version)_IntuneWin_RedirectStandardOutput.log"
$RedirectStandardErrorFilePath = "C:\Windows\Logs\Install_$($Software)_$($version)_IntuneWin_RedirectStandardError.log"


# We use Start-Process which gets most probability the ExitCode
# Some really really rare cases $LastExitCode does not work, so this works better
$Process = Start-Process $Installer -ArgumentList $Arguments -Wait -Passthru -NoNewWindow -RedirectStandardOutput $RedirectStandardOutputFilePath -RedirectStandardError $RedirectStandardErrorFilePath
$Process.WaitForExit()
$ExitCode = $Process.ExitCode

if($ExitCode -eq 0) {
	Write-Host "$Software $version installation succeeded"
} else {
	Write-Host "Error installing $Software $version. Return code: $Success"
	Write-Host "Exiting installation with error code $Success"
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
	Write-Host "Installation ready"
	Stop-Transcript
	Exit 0
} else {
	Write-Host "Error copying lame_enc.dll file"
	Write-Host "Exiting installation with error code 2"
	Stop-Transcript
	Exit 2
}

