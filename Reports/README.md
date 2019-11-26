# Intune HTML reports
Reports which we don't have in Intune (at least at this time)

Make sure you have Intune Powershell Cmdlets installed and updated

**Install-Module -Name Microsoft.Graph.Intune -Force**

Check more tips on module management
(https://github.com/petripaavola/Intune/blob/master/Powershell_Commands/Intune_Powershell_Commands_Examples.ps1)

## Create_IntuneMobileAppAssignments_HTML_Report.ps1
[Create_IntuneMobileAppAssignments_HTML_Report.ps1](https://github.com/petripaavola/Intune/blob/master/Reports/Create_IntuneMobileAppAssignments_HTML_Report.ps1)

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
