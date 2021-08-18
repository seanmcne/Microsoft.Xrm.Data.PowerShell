# Microsoft.Xrm.Data.PowerShell
[![PowerShell Gallery Version](https://img.shields.io/powershellgallery/v/Microsoft.Xrm.Data.PowerShell)](https://www.powershellgallery.com/packages/Microsoft.Xrm.Data.Powershell)
[![PowerShell Gallery Downloads](https://img.shields.io/powershellgallery/dt/Microsoft.Xrm.Data.PowerShell)](https://www.powershellgallery.com/packages/Microsoft.Xrm.Data.Powershell)

[![GitHub forks](https://img.shields.io/github/forks/seanmcne/Microsoft.Xrm.Data.PowerShell?style=social)](https://github.com/seanmcne/Microsoft.Xrm.Data.PowerShell/network/members) [![GitHub Repo stars](https://img.shields.io/github/stars/seanmcne/Microsoft.Xrm.Data.PowerShell?style=social)](https://github.com/seanmcne/Microsoft.Xrm.Data.PowerShell/stargazers)

### Overview 
**Microsoft.Xrm.Data.Powershell** contains two modules, the first is Microsoft.Xrm.Tooling.CrmConnector.Powershell which is owned and maintained by Microsoft, the second is Microsoft.Xrm.Data.Powershell which is a wrapper over this connector providing helpful functions. 

**Installation Options:**
- **PREFERRED** [Install via PowerShell](/README.md#preferred-install-the-module-via-powershell-gallery) 
- **Use only if absolutely required** [Install via a downloaded zip](/README.md#alternative-how-to-file-copy-or-manually-deploy-this-module) (Primarily used for  troubleshooting or cannot install from the gallery)

[How the module works](/README.md#how-microsoftxrmdatapowershell-works)  
[How to get a list of the commands](/README.md#how-to-get-command-details)  
[About the Authors](/README.md#about-authors)  

New releases of this can be  found on the [Release Page](https://github.com/seanmcne/Microsoft.Xrm.Data.PowerShell/releases) or can be downloaded using OneGet (Install-Module) from the Powershell Gallery. 

**Microsoft.Xrm.Data.Powershell** 

This module builds from Microsoft.Xrm.Tooling.CrmConnector.Powershell, on top of this we are providing common functions to create, delete, query, and update data.  We have also included many helpful functions for common tasks such as publishing, and manipulating System & CRM User Settings, etc. The module will function for both Dynamics CRM Online and On-Premise environments. Note: while you can import or create data this utility was not specifically designed to do high throughput data imports. For data import please refer to our blog at https://aka.ms/CRMInTheField - you may also review sample code written to add high speed/high throughput data manipulation to .NET projects posted by Austin Jones & Sean McNellis at: https://github.com/seanmcne/XrmCoreLibrary

**Microsoft.Xrm.Tooling.CrmConnector.Powershell**

This module comes from Dynamics CRM SDK and it exposes two functions, Get-CrmOrganizations and Get-CrmConnection. See the link for more detail. [Use PowerShell cmdlets for XRM tooling to connect to CRM](https://docs.microsoft.com/en-us/powershell/module/microsoft.xrm.tooling.crmconnector.powershell/?view=pa-ps-latest) - (For reference, click for the [Previous Documentation Url](https://technet.microsoft.com/en-us/library/dn689040.aspx))

### Preferred: Install the module via PowerShell Gallery
Note this method requires: Powershell Management Framework 5 or higher - details: https://www.powershellgallery.com/ 

```Powershell
Install-Module Microsoft.Xrm.Data.PowerShell -Scope CurrentUser
```
To Update to a newer release of the module
```Powershell
Update-Module Microsoft.Xrm.Data.PowerShell -Force
```

Troubleshooting: 
1. Try adding the -verbose flag to your install and update module commands - this should give you more information
2. As this module is not signed, you may need to change Execution Policy to load the module. You can do so by executing the following command then try again
```PowerShell
 Set-ExecutionPolicy –ExecutionPolicy RemoteSigned –Scope CurrentUser
```
3. If powershell is out of date and doesn't support TLS 1.2 by default - enable it to use TLS 1.2 by executing the following command, then try again 
```PowerShell
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
```

### Alternative: How to file copy or manually deploy this module - Only if absolutely required
1. Go to Releases(https://github.com/seanmcne/Microsoft.Xrm.Data.PowerShell/releases) and Download Microsoft.Xrm.Data.Powershell.zip.
2. Right click the downloaded zip file and click "Properties". 
3. Check "Unblock" checkbox and click "OK", or simply click "Unblock" button depending on OS versions. 

&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;![image](https://user-images.githubusercontent.com/2292260/114417281-14f5d200-9b77-11eb-8d5f-28f3e8795c0f.png)

4. Extract the zip file and copy "Microsoft.Xrm.Data.PowerShell" folder to *either* one of the following folders - one is for a user only scope, the second is for a system-wide scope (for all users): 
  * %USERPROFILE%\Documents\WindowsPowerShell\Modules
  * %WINDIR%\System32\WindowsPowerShell\v1.0\Modules

5. As this module is not signed, you may need to change Execution Policy to load the module. You can do so by executing following command.
```PowerShell
 Set-ExecutionPolicy –ExecutionPolicy RemoteSigned –Scope CurrentUser
```
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;*Refer to 
[Set-ExecutionPolicy](https://technet.microsoft.com/en-us/library/ee176961.aspx) 
for more information.*

6. Open PowerShell and run following command to load the module.
``` powershell
#Import Micrsoft.Xrm.Data.Powershell module 
Import-Module Microsoft.Xrm.Data.Powershell
```
* *The module is not compatible yet with PowerShell Core and must use v4 or v5 - this is due to the dependency on XrmTooling. If there is an update for XrmTooling we will explore porting this to work with that new module depending on the amount of effort required.*

### How Microsoft.Xrm.Data.Powershell works
Microsoft.Xrm.Data.Powershell module exposes many functions, but you can use Connect-CrmOnlineDiscovery, Connect-CrmOnPremDiscovery to connect to any CRM organization by using Discovery Service. Use Connect-CrmOnline function for Azure Automation. By executing these function, it creates $conn global variable. Any other functions which needs to connect to the CRM Organization takes connection parameter. You can explicitly specify the connection by using -conn parameter, but if you omit the connection, functions retrieve connection from global variable.

Alternatively, you can create multiple connection objects and pass them into each function under the –conn parameter.

### Example
This example shows how to create connection and do CRUD operation as well as manipulate System Settings.
1. Run following command to connect to Dynamics CRM Organization via  the Xrm Tooling GUI.
```PowerShell
# Online - use oAuth and XrmTooling Ui by providing your UPN and the enviroment url
connect-crmonline -Username "user@domain.com" -ServerUrl <orgurl>.crm.dynamics.com

# OnPrem sample using discovery
Connect-CrmOnPremDiscovery -InteractiveMode

# Azure Automation example
# this uses an application user - see here for more details: https://docs.microsoft.com/en-us/power-platform/admin/create-users-assign-online-security-roles#create-an-application-user
$oAuthClientId = "00000000-0000-0000-0000-000000000000"
$encryptedClientSecret = Get-AutomationVariable -Name ClientSecret
Connect-CrmOnline -ClientSecret $encryptedClientSecret -OAuthClientId $oAuthClientId -ServerUrl "https://<org>.crm.dynamics.com"
```
For Azure Automation, write all scripts inside inlinescript block as Runbook or use PowerShell type.  When using Azure automation or any non-end-user connection (*headless*) an ApplicationUser should be used - for details on how to create an application user see here: https://docs.microsoft.com/en-us/power-platform/admin/create-users-assign-online-security-roles#create-an-application-user

2. Run following command to test CRUD.
```PowerShell
# Create an account and store record Guid to a variable 
$accountId = New-CrmRecord -conn $conn -EntityLogicalName account -Fields @{"name"="Sample Account";"telephone1"="555-5555"} 
 
# Display the Guid 
$accountid 
 
# Retrieve a record and store record to a variable 
$account = Get-CrmRecord -conn $conn -EntityLogicalName account -Id $accountId -Fields name,telephone1 
 
# Display the record 
$account 
 
# Set new name value for the record 
$account.name = "Sample Account Updated" 
 
# Update the record 
Set-CrmRecord -conn $conn -CrmRecord $account 
 
# Retrieve the record again and display the result 
Get-CrmRecord -conn $conn -EntityLogicalName account -Id $accountId -Fields name 
 
# Delete the record 
Remove-CrmRecord -conn $conn -CrmRecord $account
```
3. Run following command to manipulate SystemSettings.
```PowerShell
# Display the current setting 
Get-CrmSystemSettings -conn $conn -ShowDisplayName 
 
# Change the PricingDecimalPrecision system setting from 0 to 1 
Set-CrmSystemSettings -conn $conn -PricingDecimalPrecision 1 
 
# Display the current setting 
Get-CrmSystemSettings -conn $conn -ShowDisplayName
```
### How to get command details
Each command has detail explanation.
1. Run following command to get all commands.
```PowerShell
Get-Command *crm*
```
2. Run following command to get help.
```PowerShell
Get-Help New-CrmRecord -Detailed
```
### About Authors
This module is implemented by Sean McNellis and Kenichiro Nakamura, these helper functions make use of the Microsoft provided powershell module hosted (standalone) here: https://www.powershellgallery.com/packages/Microsoft.Xrm.Tooling.CrmConnector.PowerShell/. 

<a href="https://twitter.com/seanmcne" target="_blank">Sean McNellis</a>, Principal Customer Engineer based out of North America and works supporting Dynamics & PowerPlatform Customers.
Kenichiro Nakamura, Sr. Software Engineer based out of Japan.

Blog (English): <a href="http://aka.ms/CrmInTheField" target="_blank">http://aka.ms/CrmInTheField</a>
Blog (Japanese): <a href="http://blogs.msdn.com/CrmJapan" target="_blank">http://blogs.msdn.com/CrmJapan</a>
Twitter: [@pfedynamics](https://twitter.com/pfedynamics)

Current Powershell Gallery location is here: https://www.powershellgallery.com/packages/Microsoft.Xrm.Data.Powershell
