# Intune Device Query examples

## Event logs Device Query
```
// Get Intune MDM Management policy errors
WindowsEvent('Microsoft-Windows-DeviceManagement-Enterprise-Diagnostics-Provider/Admin',12h)
| where EventId == 404 and Message !contains "Fake"
| project LoggedDateTime, EventId, Message
```
```
// Get MSI Application installations from last 7 days
WindowsEvent('Application', 7d)
| where Message contains 'Installation completed successfully' 
| project LogName, EventId , Message, LoggedDateTime
```
```
// Check for Windows 11 24H2 2025-01 Cumulative Update installation steps and times
// You can see when CU was installed and when computer was restarted -> Staged state changed to Installed State
WindowsEvent('Setup', 30d)
| where Message contains "KB5050009" and (EventId == 1 or EventId == 4 or EventId == 2)
| project LoggedDateTime, Message, EventId
| order by LoggedDateTime desc
```
```
// Get Bitlocker events from last 1 days
WindowsEvent('Microsoft-Windows-Bitlocker/Bitlocker Management',1d)
| order by LoggedDateTime
```
```
// Get Bitlocker Errors and Warnings from last 7 days
WindowsEvent('Microsoft-Windows-Bitlocker/Bitlocker Management',7d)
| order by LoggedDateTime
| where Level == 'ERROR' or Level == 'WARNING'
```
```
// Get AppLocker warning for EXE and DLL
WindowsEvent('Microsoft-Windows-AppLocker/EXE and DLL',7d)
| where Level == 'Warning'
```
```
// Get AppLocker warning for MSIs and scripts
WindowsEvent('Microsoft-Windows-AppLocker/MSI and Script',1d)
| where Level == 'Warning'
```

## Registry Device Query

```
// Get IP-address from registry
WindowsRegistry('HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Interfaces\*')
| where ValueName contains 'IPaddress' and ValueData != '0.0.0.0'
| project ValueName, ValueData
```
```
// Get Defender for Endpoint DeviceTag value from registry
WindowsRegistry('HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows Advanced Threat Protection\DeviceTagging') 
| where ValueName == 'Group'
| project ValueData
```
```
// Get Chrome policies registry keys
WindowsRegistry('HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Google\Chrome\*')
| project RegistryKey, ValueName, ValueType, ValueData
```
```
WindowsRegistry('HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows Defender\Policy Manager') 
| where ValueData == '0'
```

## Files Device Query
```
// Get cmd.exe file information
FileInfo('c:\\windows\\system32\\cmd.exe')
| project FileName, Directory, ProductVersion, DateTime=substring(tostring(CreatedDateTime), 0, 10)
```
```
// Get all files and subfiles and subfolders in C:\Windows\Logs
FileInfo('C:\Windows\Logs')
```
```
// Get all user profiles in computer - workaround
FileInfo('c:\users\*\ntuser.ini')
| project User=substring(Directory,9,)
| where User != 'Administrator' and User !contains 'default' and User !contains '$'
| order by User asc
```

## Certificates Device Query

```
// Get Intune Suite CloudPKI certificate
Certificate
| where CommonName startswith 'CloudPKI'
| project CommonName, ValidToDateTime, SelfSigned, SigningAlgorithm, KeyStrength
```
```
// Get DigiCert issuer certificates
Certificate
| where Issuer contains 'DigiCert'
| project CommonName, ValidToDateTime, SelfSigned, SigningAlgorithm, KeyStrength
```

## Process Device Query
```
// Get explorer.exe process information
Process
| where ProcessName == 'explorer.exe'
| project ProcessId, ProcessName, Path, TotalSizeBytes
```

## EncryptableVolume Device Query
```
EncryptableVolume
| join Tpm
| project WindowsDriveLetter, ProtectionStatus, EncryptionMethod, EncryptionPercentage, Activated, Enabled, SpecVersion, Manufacturer
| where ProtectionStatus == 'protected'
```

## LocalUserAccount Device Query
```
// List all local user accounts
LocalUserAccount
| project UserId, Username, HomeDirectory, WindowsSid
```

## WindowsDriver Device Query

```
// Get Microsoft drivers
WindowsDriver
| where DriverDescription contains 'Microsoft'
```

## WindowsService Device Query

```
// Get Windows services which have StartMode Automatic but current state is stopped
WindowsService
| where StartMode == "AUTO"
| where State == "STOPPED"
```
