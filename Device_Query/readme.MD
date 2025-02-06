# Intune Device Query examples
## Registry
**Get IP-address from registry**
```
// Get IP-address from registry
WindowsRegistry('HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Interfaces\*')
| where ValueName contains 'IPaddress' and ValueData != '0.0.0.0'
| project ValueName, ValueData
```
