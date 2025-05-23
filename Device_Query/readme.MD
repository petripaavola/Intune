## Table of Contents

### Intune Device Query examples
- [Event logs Device Query](#event-logs-device-query)
- [Registry Device Query](#registry-device-query)
- [Files Device Query](#files-device-query)
- [Certificates Device Query](#certificates-device-query)
- [Process Device Query](#process-device-query)
- [EncryptableVolume Device Query](#encryptablevolume-device-query)
- [LocalUserAccount Device Query](#localuseraccount-device-query)
- [WindowsDriver Device Query](#windowsdriver-device-query)
- [WindowsService Device Query](#windowsservice-device-query)

### Intune Multi Device Query examples
- [Battery](#Battery)
- [BiosInfo](#BiosInfo)
- [Cpu](#Cpu)
- [Tpm](#Tpm)
- [WindowsQfe](#windowsqfe)


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
```
// Get Boot times in last 40 days
WindowsEvent('System',40d)
| where EventId == 6005
| order by LoggedDateTime
| project LoggedDateTime, Message
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

# Intune Multi Device Query examples

## Battery
```
// Get devices Battery CycleCount and sort CycleCount descending and show only few properties (with project)
Battery
| project Device, CycleCount, DesignedCapacity, FullChargedCapacity
| order by CycleCount desc
```
```
// Get devices where Battery FullChargeCapacity is lower than DesignedCapacity.
// This tells us if Battery is not 100% anymore
Battery
| where FullChargedCapacity < DesignedCapacity
| order by CycleCount desc
```

## BiosInfo
```
// Get BiosInfo for HP devices where BIOS information is not the latest version (P78 Ver. 01.49)
BiosInfo
| where Manufacturer == "HP" and SmBiosVersion != "P78 Ver. 01.49"
```

## Cpu
```
// Get ARM64 devices using exact operator
Cpu
| where Architecture == "ARM64"
```
```
// Get ARM devices using startswith operator
Cpu
| where Architecture startswith "ARM"
```
```
// Get ARM devices using contains operator
Cpu
| where Architecture contains "ARM"
```

## Tpm
```
// Get devices where TPM is NOT in desired state
Tpm
| where Activated == false or Enabled == false or Owned == false
```
```
// Get devices where TPM IS Activated
Tpm
| where Activated == true
```
```
// Get devices where TPM is NOT Activated
Tpm
| where Activated == false
```


## WindowsQfe
```
// Get 2025-04 Cumulative Update for W11 23H2 and W11 24H2
WindowsQfe
| where HotFixId == "KB5055528" or HotFixId == "KB5055523"
```
