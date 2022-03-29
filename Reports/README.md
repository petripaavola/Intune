# Intune HTML reports
Reports which we don't have in Intune (at least at this time)

Quick links to reports:
* [Create_IntuneConfigurationAssignments_HTML_Report.ps1](https://github.com/petripaavola/Intune/blob/master/Reports/Create_IntuneConfigurationAssignments_HTML_Report.ps1)
* [Create_IntuneMobileAppAssignments_HTML_Report_v2.0.ps1](https://github.com/petripaavola/Intune/blob/master/Reports/Create_IntuneMobileAppAssignments_HTML_Report_v2.0.ps1)

Make sure you have Intune Powershell Cmdlets installed and updated

**Install-Module -Name Microsoft.Graph.Intune -Force**

Check more tips on module management
(https://github.com/petripaavola/Intune/blob/master/Powershell_Commands/Intune_Powershell_Commands_Examples.ps1)

## Create_IntuneConfigurationAssignments_HTML_Report.ps1
[Create_IntuneConfigurationAssignments_HTML_Report.ps1](https://github.com/petripaavola/Intune/blob/master/Reports/Create_IntuneConfigurationAssignments_HTML_Report.ps1)

2 views:
* Find where single Configuration Profile is assigned to
* Find all Configuration Profiles that are assigned to single AzureAD Group (this view does not exist in Intune)

Platform based tables and quick links

### Overview
![01-DeviceConfigurationsReport_Overview.png](https://github.com/petripaavola/Intune/blob/master/Reports/pics/01-DeviceConfigurationsReport_Overview.png)

### Where Configuration Profile is assigned to
![02-DeviceConfigurationsReport_Platform_by_profilename.png](https://github.com/petripaavola/Intune/blob/master/Reports/pics/02-DeviceConfigurationsReport_Platform_by_profilename.png)

### What Configuration Profiles are assigned to AzureAD Group
![03-DeviceConfigurationsReport_Targeted_to_AzureADGroup.png](https://github.com/petripaavola/Intune/blob/master/Reports/pics/03-DeviceConfigurationsReport_Targeted_to_AzureADGroup.png)

---

## Create_IntuneMobileAppAssignments_HTML_Report_v2.0.ps1
[Create_IntuneMobileAppAssignments_HTML_Report_v2.0.ps1](https://github.com/petripaavola/Intune/blob/master/Reports/Create_IntuneMobileAppAssignments_HTML_Report_v2.0.ps1)

2 views:
* Find where single Application is assigned to
* Find all Apps that are assigned to single AzureAD Group (this view does not exist in Intune)

Platform based tables and quick links

### Overview
![Report_WindowsAppAssignmentsOverview2.png](https://github.com/petripaavola/Intune/blob/master/Reports/Report_WindowsAppAssignmentsOverview2.png)

### Where application is assigned to
![Report_WindowsAppAssignmentsSortedWithAppDisplayName2.png](https://github.com/petripaavola/Intune/blob/master/Reports/Report_WindowsAppAssignmentsSortedWithAppDisplayName2.png)

### What Apps are assigned to AzureAD Group
![Report_WindowsAppAssignmentsGroupedByAssignmentGroupName2.png](https://github.com/petripaavola/Intune/blob/master/Reports/Report_WindowsAppAssignmentsGroupedByAssignmentGroupName2.png)
