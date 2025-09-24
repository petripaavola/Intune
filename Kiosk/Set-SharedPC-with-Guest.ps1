$namespaceName = "root\cimv2\mdm\dmmap"
$parentID="./Vendor/MSFT/Policy/Config"
$className = "MDM_SharedPC"
$cimObject = Get-CimInstance -Namespace $namespaceName -ClassName $className
if (-not ($cimObject)) {
   $cimObject = New-CimInstance -Namespace $namespaceName -ClassName $className -Property @{ParentID=$ParentID;InstanceID=$instance}
}
$cimObject.EnableSharedPCMode = $True
$cimObject.SetEduPolicies = $True
$cimObject.SetPowerPolicies = $True
$cimObject.MaintenanceStartTime = 0
$cimObject.SignInOnResume = $True
$cimObject.SleepTimeout = 0
$cimObject.EnableAccountManager = $True
$cimObject.AccountModel = 2
$cimObject.DeletionPolicy = 1
$cimObject.DiskLevelDeletion = 25
$cimObject.DiskLevelCaching = 50
$cimObject.RestrictLocalStorage = $False
$cimObject.KioskModeAUMID = ""
$cimObject.KioskModeUserTileDisplayText = ""
$cimObject.InactiveThreshold = 0
Set-CimInstance -CimInstance $cimObject