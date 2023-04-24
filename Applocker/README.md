# Create Intune Applocker configuration profile from Applocker XML file #

Create Intune Applocker Configuration Profile with this script.

## Quick links to files:
* [Create_Applocker_Intune_policy.ps1](https://github.com/petripaavola/Intune/blob/master/Applocker/Create_Applocker_Intune_policy.ps1)

## Description

This script reads Applocker exported XML rules file and creates Intune Applocker custom oma-uri configuration profile.

Script makes XML syntax check first. Only **Enabled** or **AuditOnly** mode rules are created to Intune. NotConfigured rules are ignored.

**Double check Intune rules configurations before applying to devices. Always test first on test devices before applying policy to production devices.**

Script saves 2 JSON files locally to script folder:
  * **Intune Applocker configuration profile**
    * For example: *Applocker Policy - Created 2023-04-24 103049.json*
  * **Graph API response are saved in current directory**
    * For example. *Applocker Policy - Created 2023-04-24 103712 - GraphAPI response.json*

With option **-WhatIf** script does XML syntax check and creates Intune configuration JSON file locally.  
Filename example is: *Applocker Policy - Created 2023-04-24 103044 - WhatIf.json*

### Intune Applocker configuration policy ###
Script creates Intune Applocker policies with naming syntax:  
**Applocker Policy - Created 2023-04-24 103712 by admin.username@organization.com**

**Test and double check Applocker configuration profile XML syntax before production deployment.**

### Parameters ###

**-FilePath**
Path to Applocker XML file  
```
.\Create_Applocker_Intune_policy.ps1 -Filepath .\AppLocker.xml
```

**-WhatIf**
Does XML file syntax check and creates rules json file. Filename is like: Applocker Policy - Created 2023-04-24 102623 - WhatIf.json
```
.\Create_Applocker_Intune_policy.ps1 -Filepath .\AppLocker.xml -WhatIf
```

### Prerequisities ###

Script uses Intune Powershell management module **Microsoft.Graph.Intune**.  

You can install Intune Powershell management module to your user account with command
```
Install-Module -Name Microsoft.Graph.Intune -Scope CurrentUser
```
