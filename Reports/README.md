# Intune HTML reports (updating page at the moment, stay tuned for updates soon! :)
Reports which we don't have in Intune (at least at this time)

Quick links to reports:
* **[Create-IntuneAppAssignmentsReport.ps1](https://github.com/petripaavola/Intune/blob/master/Reports/Create-IntuneAppAssignmentsReport.ps1)**
  * **Huge update to ver3.0. Check it out!**
  * <img src="./pics/IntuneApplicationAssignmentReport.png" width=33% height=33%>
* [Create_IntuneConfigurationAssignments_HTML_Report.ps1](https://github.com/petripaavola/Intune/blob/master/Reports/Create_IntuneConfigurationAssignments_HTML_Report.ps1)
  * Updated version 3.0 work in progress


## Create-IntuneAppAssignmentsReport.ps1 ver3.0 ###
**Huge update to this 3.0 version! (updated 24.5.2023)**

Link to script [Create-IntuneAppAssignmentsReport.ps1](https://github.com/petripaavola/Intune/blob/master/Reports/Create-IntuneAppAssignmentsReport.ps1)

### Features: ###

* 2 main views usually used:
  * Find AzureAD Groups where single Application is assigned to
    * Sort by **App name** column (default)
  * Find all Apps that are assigned to single AzureAD Group
    * **this view does not exist in Intune**
    * Sort by **Target Group** column
* See **impact** of Assignment
  * Get number of devices and users in Assignment group
* Realtime **filtering and free text search**
  * Filter by OS, App Type, Assignment Target Group, Intune Filter Name
  * Free text search
* **Sort** by any column
* **Hover** on ApplicationName, TargetGroup and/or FilterName to get more information
* **Web link** to Intune Application, Target AzureAD Group and Intune Filter
* **Export** CSV file, json file and paste to Excel

### Overview
![IntuneApplicationAssignmentReport.png](https://github.com/petripaavola/Intune/blob/master/Reports/pics/IntuneApplicationAssignmentReport.png)

### What Apps are assigned to AzureAD Group
![IntuneApplicationAssignmentReportSortByAssignmentGroup.pngg](https://github.com/petripaavola/Intune/blob/master/Reports/pics/IntuneApplicationAssignmentReportSortByAssignmentGroup.png)

### Usage: ###

Make sure you have Intune Powershell module installed and updated  
You can install Intune Powershell management module to your user account with command
```
Install-Module -Name Microsoft.Graph.Intune -Scope CurrentUser
```
**Run script**  
```
./Create-IntuneAppAssignmentsReport.ps1
```
### Parameters ###

**-UseOfflineCache**  
Create report using files from local cache folder.
```
./Create-IntuneAppAssignmentsReport.ps1 -UseOfflineCache
```
**-ExportCSV**  
Export report as ; limited CSV file.
```
./Create-IntuneAppAssignmentsReport.ps1 -ExportCSV
```
**-ExportJSON**  
Export report as JSON file
```
./Create-IntuneAppAssignmentsReport.ps1 -ExportJSON
```
**-ExportToExcelCopyPaste**  
Export report to Clipboard. You can paste it to Excel and excel will paste data to columns automatically.
```
./Create-IntuneAppAssignmentsReport.ps1 -ExportToExcelCopyPaste
```
**-DoNotOpenReportAutomatically**  
Do not automatically open HTML report to Web browser. Can be used when automating report creation.
```
./Create-IntuneAppAssignmentsReport.ps1 -DoNotOpenReportAutomatically
```
**-UpdateIconsCache**  
Update App icon cache. New Apps will always get icons downloaded automatically but existing icons are not automatically updated
```
./Create-IntuneAppAssignmentsReport.ps1 -UpdateIconsCache
```
**-IncludeAppsWithoutAssignments**  
Include Intune Application without Assignments. This will get a lot of Apps you didn't even know exists inside Intune/Graph API.
```
./Create-IntuneAppAssignmentsReport.ps1 -IncludeAppsWithoutAssignments
```
**-DoNotDownloadAppIcons**  
Do not download Application icons.
```
./Create-IntuneAppAssignmentsReport.ps1 -DoNotDownloadAppIcons
```
**-IncludeIdsInReport**  
Include Appication Ids in report. This makes wider so it is disabled by default.
```
./Create-IntuneAppAssignmentsReport.ps1 -IncludeIdsInReport
```
**-IncludeBase64ImagesInReport**  
Includes Application icons inside HTML file so report will have icons if HTML if copied somewhere else. Note! This is slow and creates huge HTML file.
```
./Create-IntuneAppAssignmentsReport.ps1 -IncludeBase64ImagesInReport
```

---

## Create_IntuneConfigurationAssignments_HTML_Report.ps1
[Create_IntuneConfigurationAssignments_HTML_Report.ps1](https://github.com/petripaavola/Intune/blob/master/Reports/Create_IntuneConfigurationAssignments_HTML_Report.ps1)

**Work is in progress to update this to 3.0 version which has same features than the Application report has**

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
