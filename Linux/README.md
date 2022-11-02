# Intune Linux Custom Compliance scripts
Example scripts how to run Linux Custom Compliance checks with Intune.
This script is Bash script which includes Powershell script inside.

## Quick links to files:
* [Intune_Linux_Custom_Compliance_script.sh](https://github.com/petripaavola/Intune/blob/master/Linux/Intune_Linux_Custom_Compliance_script.sh)
* [Intune_Linux_Custom_Compliance_script_Rules_file.json](https://github.com/petripaavola/Intune/blob/master/Linux/Intune_Linux_Custom_Compliance_script_Rules_file.json)

## Compliance checks in this script
* Reboot Required check (file should not exist /var/run/reboot-required)

## How script works
Compliance script itself is Bash script because that is requirement for Intune. However script includes Powershell script and compliance checks are done in Powershell part.

So Bash script is just launcher for Powershell script

Script creates 2 log files for debugging and testing
* /tmp/IntuneCustomComplianceScript_Bash.log
* /tmp/IntuneCustomComplianceScript_Powershell.log

Script is run in user context who enrolled Ubuntu device to Intune. You can verify this from log files.

Be sure NOT to write anything to STDOUT because that will break the compliance check. Only Powershell part of script is allowed to return compressed JSON to STDOUT which is then passed to Intune for custom Compliance Check.

## Requirements
* Powershell must be installed for this script to work
  * Note! Script does not check existence of Powershell at this time
* You can create Dynamic Azure AD Group for targeting the Linux Compliance Policy using this rule
  * **(device.deviceOSType -eq "Linux")**


## Screenshots
![Intune_Linux_CustomComplianceCheck_RebootRequired_Compliant.png](https://github.com/petripaavola/Intune/blob/master/Linux/Intune_Linux_CustomComplianceCheck_RebootRequired_Compliant.png)
![Intune_Linux_CustomComplianceCheck_RebootRequired_NotCompliant.png](https://github.com/petripaavola/Intune/blob/master/Linux/Intune_Linux_CustomComplianceCheck_RebootRequired_NotCompliant.png)

**Intune Linux Compliance Policy**
![Intune_Linux_Compliance_Policy.png](https://github.com/petripaavola/Intune/blob/master/Linux/Intune_Linux_Compliance_Policy.png)
