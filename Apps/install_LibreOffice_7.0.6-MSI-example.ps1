# Install LibreOffice 7.0.6 MSI
# IntuneWin32 MSI Powershell script install example
#
# UI Languages: FI, EN, SE
# Proofing languages: English, German, French, Swedish and Spanish
#
#
# Petri.Paavola@yodamiitti.fi
# Microsoft MVP - Windows and Intune
# 2024-11-01
#
# Original script source:
# https://github.com/petripaavola/Intune/blob/master/Apps/install_LibreOffice_7.0.6-MSI-example.ps1
#
# Intune install command:
# Powershell.exe -ExecutionPolicy Bypass -File install_LibreOffice_7.0.6-MSI-example.ps1
#
# Intune uninstall command:
# msiexec /x {9F9A9C01-5A65-4C2E-A243-FC88C81BC35F} /qn /l*v C:\Windows\Logs\Uninstall_LibreOffice_7.0.6_MSI.log
#
# MSI ProductCode={9F9A9C01-5A65-4C2E-A243-FC88C81BC35F}
#
# This is MSI application install example script for MSI-files
# In this case there is so long parameter that it will not fit in Intune install command line


    #
    #
  # # #
   # #
    #

# Configure these values
$SoftwareName = "LibreOffice"
$version = "7.0.6"
$MSIFilePath = "$PSScriptRoot\LibreOffice_7.0.6_Win_x64.msi"

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
	Start-Transcript "C:\Windows\Logs\Install_$($SoftwareName)_$($version)_IntuneWin32App.log" -Append


	# Running inside Try so we can catch fatal errors
	try {
		Write-Host "Install $SoftwareName $version"


		# Check that $MSIFilePath exists
		if (-not (Test-Path $MSIFilePath)) {
			Write-Host "Installer file not found: $MSIFilePath"
			Write-Host "Exiting installation with error code 9999"
			Stop-Transcript
			Exit 9999
		}


		# Log file path for stdOut and stdError
		$RedirectStandardOutputFilePath = "C:\Windows\Logs\Install_$($SoftwareName)_$($version)_IntuneWin_RedirectStandardOutput.log"
		$RedirectStandardErrorFilePath = "C:\Windows\Logs\Install_$($SoftwareName)_$($version)_IntuneWin_RedirectStandardError.log"

		$Installer = "msiexec.exe"
		$Arguments = @( 
			'/i'
			"`"$MSIFilePath`""
			'/qn'
			"/l*v C:\Windows\Logs\Install_$($SoftwareName)_$($version)_MSI.log"
			'/norestart'
			'RebootYesNo=No'
			'ALLUSERS=1'
			'CREATEDESKTOPLINK=0'
			'REGISTER_ALL_MSO_TYPES=0'
			'REGISTER_NO_MSO_TYPES=1'
			'ISCHECKFORPRODUCTUPDATES=0'
			'QUICKSTART=0'
			'ADDLOCAL=ALL'
			'UI_LANGS=en_US,sv,fi'
			'REMOVE=gm_r_ex_Dictionary_Af,gm_r_ex_Dictionary_Sq,gm_r_ex_Dictionary_Tr,gm_r_ex_Dictionary_An,gm_r_ex_Dictionary_Ar,gm_r_ex_Dictionary_Be,gm_r_ex_Dictionary_Bg,gm_r_ex_Dictionary_Bn,gm_r_ex_Dictionary_Br,gm_r_ex_Dictionary_Pt_Pt,gm_r_ex_Dictionary_Pt_Br,gm_r_ex_Dictionary_Bs,gm_r_ex_Dictionary_Ca,gm_r_ex_Dictionary_Bo,gm_r_ex_Dictionary_Cs,gm_r_ex_Dictionary_Da,gm_r_ex_Dictionary_Nl,gm_r_ex_Dictionary_Et,gm_r_ex_Dictionary_Gd,gm_r_ex_Dictionary_Gl,gm_r_ex_Dictionary_Gu,gm_r_ex_Dictionary_He,gm_r_ex_Dictionary_Hi,gm_r_ex_Dictionary_Hu,gm_r_ex_Dictionary_Hr,gm_r_ex_Dictionary_Id,gm_r_ex_Dictionary_It,gm_r_ex_Dictionary_Is,gm_r_ex_Dictionary_Lt,gm_r_ex_Dictionary_Lo,gm_r_ex_Dictionary_Lv,gm_r_ex_Dictionary_Ne,gm_r_ex_Dictionary_No,gm_r_ex_Dictionary_Oc,gm_r_ex_Dictionary_Pl,gm_r_ex_Dictionary_Ro,gm_r_ex_Dictionary_Ru,gm_r_ex_Dictionary_Sr,gm_r_ex_Dictionary_Si,gm_r_ex_Dictionary_Sk,gm_r_ex_Dictionary_Sl,gm_r_ex_Dictionary_El,gm_r_ex_Dictionary_Te,gm_r_ex_Dictionary_Th,gm_r_ex_Dictionary_Uk,gm_r_ex_Dictionary_Vi,gm_r_ex_Dictionary_Zu'
		)


		# We use Start-Process which gets most probability the ExitCode
		# Some really really rare cases $LastExitCode does not work, so this works better
		$Process = Start-Process $Installer -ArgumentList $Arguments -Wait -Passthru -NoNewWindow -RedirectStandardOutput $RedirectStandardOutputFilePath -RedirectStandardError $RedirectStandardErrorFilePath
		$Process.WaitForExit()
		$ExitCode = $Process.ExitCode

		if ($ExitCode -eq 0) {
			# MSI installation was successful

			Write-Host "$SoftwareName $version installation succeeded with ExitCode $ExitCode"
			Stop-Transcript
			exit $ExitCode
		
		} elseif ($ExitCode -eq 3010) {
			# MSI installation was successful but reboot is required after software install

			Write-Host "$SoftwareName $version installed successfully but requires a reboot (ExitCode: $ExitCode)"
			Stop-Transcript
			exit $ExitCode

		} else {
			# MSI installation failed
			
			Write-Host "Error installing $SoftwareName. ExitCode: $ExitCode"
			Write-Host "Exiting installation with ExitCode $ExitCode"
			Stop-Transcript
			Exit $ExitCode
		}
	} catch {
		# Handle any unexpected errors

		Write-Host "An unexpected fatal error occurred: $_"
		Write-Host "Exiting installation with ExitCode 99999"
		Stop-Transcript
		Exit 99999
		
	}
}