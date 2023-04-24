<#
.Synopsis
   This script reads Applocker exported XML rules file and creates Intune Applocker configuration profile.

.DESCRIPTION
   This script reads Applocker exported XML rules file and creates Intune Applocker custom oma-uri configuration profile.
   
   Script makes XML syntax check first. Only Enabled or AuditOnly mode rules are created to Intune. NotConfigured rules are ignored.

   Double check Intune rules configurations before applying to devices. Always test first on test devices before applying policy to production devices.

   Created json files (Intune Applocker configuration profile) and Graph API response are saved in current directory.


   Script uses Microsoft Intune Powershell management module.
   
   You can install Intune Powershell module to your user account without Administrator permissions with command:
   Install-Module -Name Microsoft.Graph.Intune -Scope CurrentUser


   Author:
   Petri.Paavola@yodamiitti.fi
   Senior Modern Management Principal
   Microsoft MVP - Windows and Devices for IT
   
   2023-04-24

   Version 1.1

   https://github.com/petripaavola/Intune/tree/master/Applocker

.PARAMETER FilePath
   Applocker exported rules XML file path

.PARAMETER WhatIf
   Does XML file syntax check and creates rules json file. Filename is like: Applocker Policy - Created 2023-04-24 102623 - WhatIf.json

.EXAMPLE
   .\Create_Applocker_Intune_policy.ps1 .\AppLocker.xml
.EXAMPLE
   .\Create_Applocker_Intune_policy.ps1 -Filepath .\AppLocker.xml
.EXAMPLE
   .\Create_Applocker_Intune_policy.ps1 -Filepath .\AppLocker.xml -WhatIf

.INPUTS
.OUTPUTS
   Script creates 2 json files. Intune Applocker XML configuration json file and Microsoft Graph API response json file.
   
   With option -WhatIf script creates Intune Applocker XML configuration json file which ends with -WhatIf.json
.NOTES
.LINK
   https://github.com/petripaavola/Intune/tree/master/Applocker
#>

[CmdletBinding()]
Param(
	[Parameter(Mandatory=$True,
				HelpMessage = 'Enter Applocker XML file fullpath',
                ValueFromPipeline=$true,
                ValueFromPipelineByPropertyName=$true)]
	[Alias("XMLFilePath")]
    [String]$FilePath,
	[Parameter(Mandatory=$False,
                ValueFromPipeline=$true,
                ValueFromPipelineByPropertyName=$true)]
    [Switch]$WhatIf
)

# Script version
$Version='1.1'


# region Functions
Function Validate-XMLSyntax {
	Param(
		$XMLFilePath
	)

	Try {
		[xml]$xmlTest = Get-Content $XMLFilePath -Encoding UTF8 -Raw
		return $True
	} catch {
		Write-Host "Error:`n$($_.Exception.Message)" -ForegroundColor Red
		return $False
	}	
}

# Create a helper function to format the XmlNode string indents
#
# This function was provided and automatically documented by ChatGPT-4
# Good example of AI usage in coding
function FormatXmlNodeString {
	Param(
		$node
	)
	
    $xmlString = $node.OuterXml

    # Add newline before every opening and closing tag (excluding self-closing tags)
    $formattedXmlString = $xmlString -replace '(?<=</[^<>]+>)(?=<[^<>]+>)', "`r`n" -replace '(?<=<[^<>]+>)(?=<[^<>]+>)', "`r`n"
    
    # Initialize indent level
    $indent = 0

    # Process XML string line by line to add proper indentation
    $formattedXmlString = ($formattedXmlString -split "`r`n" | ForEach-Object {
        # If the line starts with a closing tag, decrease the indent level by 2
        if ($_ -match '^</') {
            $indent -= 2
        }
        
        # Create a new line with the current indent level
        $line = (' ' * $indent) + $_

        # If the line starts with an opening tag (not a closing or self-closing tag), increase the indent level by 2
        if ($_ -match '^<[^(?:\!|/)].*[^/]>$') {
            $indent += 2
        }

        # Return the formatted line
        $line
    }) -join "`r`n"

    return $formattedXmlString
}

# endregion Functions

# region main

Write-Host "Import Applocker XML rules and create Applocker policy to Intune. Version $version`n" -ForegroundColor Cyan

if($WhatIf) {
	Write-Host "-WhatIf specified so not making any changes to Intune`n" -ForegroundColor Yellow
}

Write-Host "Processing file: $FilePath`n"

# Check that XML file exists
if(-not (Test-Path $FilePath)) {
	Write-Host "Could not find XML file: $FilePath" -ForegroundColor Red
	Write-Host "Script will exit`n" -ForegroundColor Yellow
	Exit 1
}

# Check XML file syntax
Write-Host "Validate Applocker XML file syntax"
if(Validate-XMLSyntax $FilePath) {
	Write-Host "OK`n" -ForegroundColor Green
} Else {
	Write-Host "FAILED: XML syntax is NOT valid`n" -ForegroundColor Red
	Write-Host "Script will exit`n" -ForegroundColor Yellow
	Exit 1	
}

# Read XML file string to variable casted as XML type
[xml]$ApplockerXML = Get-Content $FilePath -Encoding UTF8 -Raw

Write-Host "Import Applocker configurations from XML file (EXE/MSI/Appx/Script/DLL)"

# Extract RuleCollections
$exeRuleCollection = $ApplockerXML.AppLockerPolicy.RuleCollection | Where-Object { $_.Type -eq 'Exe' -and (($_.EnforcementMode -eq 'Enabled') -or ($_.EnforcementMode -eq 'AuditOnly'))}
$msiRuleCollection = $ApplockerXML.AppLockerPolicy.RuleCollection | Where-Object { $_.Type -eq 'Msi' -and (($_.EnforcementMode -eq 'Enabled') -or ($_.EnforcementMode -eq 'AuditOnly')) }
$scriptRuleCollection = $ApplockerXML.AppLockerPolicy.RuleCollection | Where-Object { $_.Type -eq 'Script'  -and (($_.EnforcementMode -eq 'Enabled') -or ($_.EnforcementMode -eq 'AuditOnly'))}
$dllRuleCollection = $ApplockerXML.AppLockerPolicy.RuleCollection | Where-Object { $_.Type -eq 'Dll'  -and (($_.EnforcementMode -eq 'Enabled') -or ($_.EnforcementMode -eq 'AuditOnly'))}
$appxRuleCollection = $ApplockerXML.AppLockerPolicy.RuleCollection | Where-Object { $_.Type -eq 'Appx'  -and (($_.EnforcementMode -eq 'Enabled') -or ($_.EnforcementMode -eq 'AuditOnly'))}


# Format RuleCollections XML strings
# Add Newlines and indents for human readability
if($exeRuleCollection) {
	$exeRuleCollectionXmlString = FormatXmlNodeString $exeRuleCollection
} else {
	$exeRuleCollectionXmlString = $null
}

if($msiRuleCollection) {
	$msiRuleCollectionXmlString = FormatXmlNodeString $msiRuleCollection
} else {
	$msiRuleCollectionXmlString = $null
}

if($scriptRuleCollection) {
	$scriptRuleCollectionXmlString = FormatXmlNodeString $scriptRuleCollection
} else {
	$scriptRuleCollectionXmlString = $null
}

if($dllRuleCollection) {
	$dllRuleCollectionXmlString = FormatXmlNodeString $dllRuleCollection
} else {
	$dllRuleCollectionXmlString = $null
}

if($appxRuleCollection) {
	$appxRuleCollectionXmlString = FormatXmlNodeString $appxRuleCollection
} else {
	$appxRuleCollectionXmlString = $null
}

Write-Host "OK`n" -ForegroundColor Green

<#
# DEBUG
# Output XML strings
Write-Output "Exe RuleCollection XML:`n$exeRuleCollectionXmlString`n"
Write-Output "Msi RuleCollection XML:`n$msiRuleCollectionXmlString`n"
Write-Output "Script RuleCollection XML:`n$scriptRuleCollectionXmlString`n"
Write-Output "Dll RuleCollection XML:`n$dllRuleCollectionXmlString`n"
Write-Output "Appx RuleCollection XML:`n$appxRuleCollectionXmlString`n"
Write-Host ""
#>

if(-not $WhatIf) {

	# Connect to Intune
	$IntunePowershellModule = Import-Module Microsoft.Graph.Intune -PassThru -ErrorAction SilentlyContinue
	if (-not $IntunePowershellModule) {
		Write-Host "Intune Powershell module not found!`n" -ForegroundColor Red
		Write-Host "You can install Intune Powershell module to your user account with command:"
		Write-Host "Install-Module -Name Microsoft.Graph.Intune -Scope CurrentUser" -ForegroundColor Cyan
		Write-Host "`nor you can install machine-wide Intune module with command:`nInstall-Module -Name Microsoft.Graph.Intune`n"
		Exit 1
	}


	Write-Host "Connecting to Intune using Powershell Intune-Module"
	Update-MSGraphEnvironment -SchemaVersion 'beta'
	$Success = $?

	if (-not $Success) {
		Write-Host "Failed to update MSGraph Environment schema to Beta!`n" -ForegroundColor Red
		Write-Host "Make sure you have installed Intune Powershell module"
		
		Write-Host "You can install Intune Powershell module to your user account with command:"
		Write-Host "Install-Module -Name Microsoft.Graph.Intune -Scope CurrentUser" -ForegroundColor Cyan
		Write-Host "`nor you can install machine-wide Intune module with command:`nInstall-Module -Name Microsoft.Graph.Intune`n"
		Exit 1
	}

	$MSGraphEnvironment = Connect-MSGraph
	$Success = $?

	if ($Success -and $MSGraphEnvironment) {
		$TenantId = $MSGraphEnvironment.tenantId
		$AdminUserUPN = $MSGraphEnvironment.upn

		Write-Host "Connected to Microsoft Intune / Microsoft Graph as user:`n$AdminUserUPN`n"
		
	} else {
		Write-Host "Could not connect to MSGraph!" -ForegroundColor Red
		Exit 1	
	}
} else {
	Write-Host "Skipping connecting to Intune Graph API management because -WhatIf is specified`n" -ForegroundColor Yellow
}

##########################################
# Create Applocker policy to Intune

$DateTime = Get-Date -Format "yyyy-MM-dd HHmmss"

$DisplayName = "Applocker Policy - Created $DateTime by $AdminUserUPN"
$Description = "Creation date: $DateTime`nCreated by: $AdminUserUPN`nApplocker rules uploaded with Powershell script v$Version`n`nApplocker rules configured:`n"
$Grouping = Get-Date -Format "yyyyMMddHHmmss"

# Applocker custom oma-uri policy JSON template
# We will add all rules to this template's omaSettings property if XML found
$BodyTemplate = @"
{
    "displayName":  "$DisplayName",
	"description":  "$Description",
    "roleScopeTagIds":  [
                            "0"
                        ],
    "@odata.type":  "#microsoft.graph.windows10CustomConfiguration",
    "omaSettings":  [
                    ]
}
"@

$Body = $BodyTemplate | ConvertFrom-Json

# Add rules XML information to Intune Policy creation request body
if($exeRuleCollectionXmlString) {
	$ExeEnforcementMode = $exeRuleCollection.EnforcementMode
	Write-Host "Found Applocker EXE rules with EnforcementMode $ExeEnforcementMode"
	Write-Host "Add EXE rules to Intune policy creation request body"

	$ExeJSON = @"
	{
		"displayName":  "Applocker EXE Rule with EnforcementMode $ExeEnforcementMode",
		"description":  "",
		"omaUri":  "./Vendor/MSFT/AppLocker/ApplicationLaunchRestrictions/EXE$Grouping/EXE/Policy",
		"@odata.type":  "#microsoft.graph.omaSettingString",
		"value":  ""
	}
"@

	# Create Powershell object from JSON
	$ExeRule = $ExeJSON | ConvertFrom-Json

	# Add rule XML value to object
	$ExeRule.value = $exeRuleCollectionXmlString

	# Add EXE Rule to Body object
	$Body.omaSettings += $ExeRule
	$Success = $?
	
	if($Success) {
		Write-Host "OK`n" -ForegroundColor Green
		
		# Add rule type and enforcement information to Description"
		$Body.Description += "`tEXE    with EnforcementMode $ExeEnforcementMode`n"
	} else {
		Write-Host "Failed" -ForegroundColor Red
		Write-Host "Script will exit`n" -ForegroundColor Yellow
		Exit 1
	}
	
} else {
	Write-Host "Did NOT find Applocker EXE rules with EnforcementMode Enabled or AuditOnly`n" -ForegroundColor Yellow
}


if($msiRuleCollectionXmlString) {
	$MsiEnforcementMode = $msiRuleCollection.EnforcementMode
	Write-Host "Found Applocker MSI rules with EnforcementMode $MsiEnforcementMode"
	Write-Host "Add MSI rules to Intune policy creation request body"

	$MsiJSON = @"
	{
		"displayName":  "Applocker MSI Rule with EnforcementMode $MsiEnforcementMode",
		"description":  "",
		"omaUri":  "./Vendor/MSFT/AppLocker/ApplicationLaunchRestrictions/MSI$Grouping/MSI/Policy",
		"@odata.type":  "#microsoft.graph.omaSettingString",
		"value":  ""
	}
"@

	# Create Powershell object from JSON
	$MsiRule = $MsiJSON | ConvertFrom-Json

	# Add rule XML value to object
	$MsiRule.value = $msiRuleCollectionXmlString

	# Add MSI Rule to Body object
	$Body.omaSettings += $MsiRule
	$Success = $?
	
	if($Success) {
		Write-Host "OK`n" -ForegroundColor Green
		
		# Add rule type and enforcement information to Description"
		$Body.Description += "`tMSI    with EnforcementMode $MsiEnforcementMode`n"
	} else {
		Write-Host "Failed" -ForegroundColor Red
		Write-Host "Script will exit`n" -ForegroundColor Yellow
		Exit 1
	}
	
} else {
	Write-Host "Did NOT find Applocker MSI rules with EnforcementMode Enabled or AuditOnly`n" -ForegroundColor Yellow
}


if($scriptRuleCollectionXmlString) {
	$scriptEnforcementMode = $scriptRuleCollection.EnforcementMode
	Write-Host "Found Applocker Script rules with EnforcementMode $scriptEnforcementMode"
	Write-Host "Add Script rules to Intune policy creation request body"

	$ScriptJSON = @"
	{
		"displayName":  "Applocker Script Rule with EnforcementMode $scriptEnforcementMode",
		"description":  "",
		"omaUri":  "./Vendor/MSFT/AppLocker/ApplicationLaunchRestrictions/Script$Grouping/Script/Policy",
		"@odata.type":  "#microsoft.graph.omaSettingString",
		"value":  ""
	}
"@

	# Create Powershell object from JSON
	$ScriptRule = $ScriptJSON | ConvertFrom-Json

	# Add rule XML value to object
	$ScriptRule.value = $scriptRuleCollectionXmlString

	# Add MSI Rule to Body object
	$Body.omaSettings += $ScriptRule
	$Success = $?

	if($Success) {
		Write-Host "OK`n" -ForegroundColor Green
		
		# Add rule type and enforcement information to Description"
		$Body.Description += "`tScript with EnforcementMode $ScriptEnforcementMode`n"
	} else {
		Write-Host "Failed" -ForegroundColor Red
		Write-Host "Script will exit`n" -ForegroundColor Yellow
		Exit 1
	}
	
} else {
	Write-Host "Did NOT find Applocker Script rules with EnforcementMode Enabled or AuditOnly`n" -ForegroundColor Yellow
}


if($dllRuleCollectionXmlString) {
	$dllEnforcementMode = $dllRuleCollection.EnforcementMode
	Write-Host "Found Applocker DLL rules with EnforcementMode $dllEnforcementMode"
	Write-Host "Add DLL rules to Intune policy creation request body"

	$DLLJSON = @"
	{
		"displayName":  "Applocker DLL Rule with EnforcementMode $dllEnforcementMode",
		"description":  "",
		"omaUri":  "./Vendor/MSFT/AppLocker/ApplicationLaunchRestrictions/DLL$Grouping/DLL/Policy",
		"@odata.type":  "#microsoft.graph.omaSettingString",
		"value":  ""
	}
"@

	# Create Powershell object from JSON
	$DLLRule = $DLLJSON | ConvertFrom-Json

	# Add rule XML value to object
	$DLLRule.value = $dllRuleCollectionXmlString

	# Add MSI Rule to Body object
	$Body.omaSettings += $DLLRule
	$Success = $?
	
	if($Success) {
		Write-Host "OK`n" -ForegroundColor Green
		
		# Add rule type and enforcement information to Description"
		$Body.Description += "`tDLL    with EnforcementMode $dllEnforcementMode`n"
	} else {
		Write-Host "Failed" -ForegroundColor Red
		Write-Host "Script will exit`n" -ForegroundColor Yellow
		Exit 1
	}
	
} else {
	Write-Host "Did NOT find Applocker DLL rules with EnforcementMode Enabled or AuditOnly`n" -ForegroundColor Yellow
}


if($appxRuleCollectionXmlString) {
	$appxEnforcementMode = $appxRuleCollection.EnforcementMode
	Write-Host "Found Applocker Appx rules with EnforcementMode $appxEnforcementMode"
	Write-Host "Add Appx rules to Intune policy creation request body"

	$AppxJSON = @"
	{
		"displayName":  "Applocker Appx Rule with EnforcementMode $appxEnforcementMode",
		"description":  "",
		"omaUri":  "./Vendor/MSFT/AppLocker/ApplicationLaunchRestrictions/Appx$Grouping/StoreApps/Policy",
		"@odata.type":  "#microsoft.graph.omaSettingString",
		"value":  ""
	}
"@

	# Create Powershell object from JSON
	$AppxRule = $AppxJSON | ConvertFrom-Json

	# Add rule XML value to object
	$AppxRule.value = $appxRuleCollectionXmlString

	# Add MSI Rule to Body object
	$Body.omaSettings += $AppxRule
	$Success = $?
	
	if($Success) {
		Write-Host "OK`n" -ForegroundColor Green
		
		# Add rule type and enforcement information to Description"
		$Body.Description += "`tAppx   with EnforcementMode $appxEnforcementMode`n"
		
	} else {
		Write-Host "Failed" -ForegroundColor Red
		Write-Host "Script will exit`n" -ForegroundColor Yellow
		Exit 1
	}
	
} else {
	Write-Host "Did NOT find Applocker Appx rules with EnforcementMode Enabled or AuditOnly`n" -ForegroundColor Yellow
}


$BodyJson = $Body | ConvertTo-Json -Depth 5

# Save JSON to file
if($WhatIf) {
	$ExportFilePath = "$PSScriptRoot\Applocker Policy - Created $DateTime - WhatIf.json"
} else {
	$ExportFilePath = "$PSScriptRoot\Applocker Policy - Created $DateTime.json"
}
Write-Host "Export Applocker Intune policy JSON to file: $ExportFilePath"
$BodyJson | Out-File -FilePath $ExportFilePath -Encoding UTF8
$Success = $?
if($Success) {
	Write-Host "OK`n" -ForegroundColor Green
} else {
	Write-Host "Failed`n" -ForegroundColor Red
}

if($WhatIf) {
	Write-Host "What if: Would upload Applocker configuration to Intune with POST body json:`n`n$BodyJson`n"
	
} else {
	Write-Host "Create Applocker Configuration profile to Intune"

	$Url = "https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations"
	$MSGraphRequest = Invoke-MSGraphRequest -Url $url -Content $BodyJson.ToString() -HttpMethod 'POST'
	$Success = $?
	if($Success -and $MSGraphRequest) {
		Write-Host "OK`n" -ForegroundColor Green
		Write-Host "Graph API response:`n"
		$MSGraphRequest
		Write-Host ""
		
		$ResponseFilePath = "$PSScriptRoot\Applocker Policy - Created $DateTime - GraphAPI response.json"
		Write-Host "Save Graph API response to file just in case that is needed later:"
		Write-Host "$ResponseFilePath"
		
		$MSGraphRequest | ConvertTo-Json -Depth 5 | Out-File -FilePath $ResponseFilePath -Encoding UTF8
		$Success = $?
		if($Success) {
			Write-Host "OK`n" -ForegroundColor Green
			Write-Host "Check Applocker Configuration profile in Intune:`n$DisplayName`n"
		} else {
			Write-Host "Failed`n" -ForegroundColor Red
		}		
		
	} else {
		Write-Host "There was possible problems uploading Applocker policy to Intune!" -ForegroundColor Yellow
		$MSGraphRequest	
	}
}

Write-Host "All done."


# endregion main