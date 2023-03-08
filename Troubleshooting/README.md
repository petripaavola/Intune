# Intune troubleshooting tools and tips

### Intune Management Extension log analyzer and LogViewerUI ###
Check my [Get-IntuneManagementExtensionDiagnostics](https://github.com/petripaavola/Get-IntuneManagementExtensionDiagnostics) tool.  
<img src="https://github.com/petripaavola/Get-IntuneManagementExtensionDiagnostics/blob/main/pics/Get-IntuneManagementExtensionDiagnostics-Observed_Timeline.png" width=25% height=25%>
<img src="https://github.com/petripaavola/Get-IntuneManagementExtensionDiagnostics/blob/main/pics/Get-IntuneManagementExtensionDiagnostics-LogViewerUI01.png" width=25% height=25%>

### Firewall ####
Check Firewall rules configured from Intune
```
Get-NetFirewallRule -PolicyStore ActiveStore | Where-Object { $_.PolicyStoreSource -eq 'Mdm' } | Select-Object -Property DisplayName,Action,Direction
```
Firewall configurations and rules registry path is:  
Computer\HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\SharedAccess\Parameters\FirewallPolicy\Mdm


#### [Show-IntuneIMELogsInOutGridView.ps1](./Show-IntuneIMELogsInOutGridView.ps1)
**You really should look my [Get-IntuneManagementExtensionDiagnostics](https://github.com/petripaavola/Get-IntuneManagementExtensionDiagnostics) tool for Intune LogViewerUI.**  

* Show Intune Management Extension logs in cmtrace style window using Out-GridView cmdlet
* Shows list of IME log files to show in Out-GridView if log file is not specified as command line parameter

Either give log file as parameter or select file from Out-GridView  
![01-Show-IntuneIMELogsInOutGridView.png](https://github.com/petripaavola/Intune/blob/master/Troubleshooting/pics/01-Show-IntuneIMELogsInOutGridView.png)

Out-Gridview looks like cmtrace.exe. Filter enables quick search by typing any text  
![02-Show-IntuneIMELogsInOutGridView.png](https://github.com/petripaavola/Intune/blob/master/Troubleshooting/pics/02-Show-IntuneIMELogsInOutGridView.png)
