# Microsoft.Xrm.Data.PowerShell
### Overview 
**Microsoft.Xrm.Data.Powershell.zip** contains one primary module, Microsoft.Xrm.Data.Powershell, but also relies on other included dll's such as Microsoft.Xrm.Tooling.CrmConnector.Powershell which we are loading as a secondary module (instead of a snap-in). 

**Installation Options:**
- [Install via PowerShell](/README.md#preferred-install-the-module-via-powershell-gallery)
- [Install via a downloaded zip](/README.md#alternative-how-to-file-copy-or-manually-deploy-this-module)

[How the module works](/README.md#how-microsoftxrmdatapowershell-works)  
[How to get a list of the commands](/README.md#how-to-get-command-details)  
[About the Authors](/README.md#about-authors)  

New releases of this can be  found on the [Release Page](https://github.com/seanmcne/Microsoft.Xrm.Data.PowerShell/releases) or can be downloaded using OneGet (Install-Module) from the Powershell Gallery. 

**Microsoft.Xrm.Data.Powershell** 

This module builds from Microsoft.Xrm.Tooling.CrmConnector.Powershell, on top of this we are providing common functions to create, delete, query, and update data.  We have also included many helpful functions for common tasks such as publishing, and manipulating System & CRM User Settings, etc. The module will function for both Dynamics CRM Online and On-Premise environments. Note: while you can import or create data this utility was not specifically designed to do high throughput data imports. For data import please refer to our blog at https://aka.ms/CRMInTheField - you may also review sample code written to add high speed/high throughput data manipulation to .NET projects posted by Austin Jones & Sean McNellis at: https://pfexrmcore.codeplex.com/ 

**Microsoft.Xrm.Tooling.CrmConnector.Powershell**

This module comes from Dynamics CRM SDK and it exposes two functions, Get-CrmOrganizations and Get-CrmConnection. See the link for more detail. [Use PowerShell cmdlets for XRM tooling to connect to CRM](https://technet.microsoft.com/en-us/library/dn689040.aspx)

### Preferred: Install the module via PowerShell Gallery
Note this method requires: Powershell Management Framework 5 or higher - details: https://www.powershellgallery.com/ 

```Powershell
Install-Module Microsoft.Xrm.Data.PowerShell -Scope CurrentUser
```
To Update to a newer release:
```Powershell
Update-Module Microsoft.Xrm.Data.PowerShell -Force
```

Troubleshooting: 
1. Try adding the -verbose flag to your install and update module commands - this should give you more information
2. As this module is not signed, you may need to change Execution Policy to load the module. You can do so by executing following command.
```PowerShell
 Set-ExecutionPolicy –ExecutionPolicy RemoteSigned –Scope CurrentUser
```

### Alternative: How to file copy or manually deploy this module
1. Go to Releases(https://github.com/seanmcne/Microsoft.Xrm.Data.PowerShell/releases) and Download Microsoft.Xrm.Data.Powershell.zip.
2. Right click the downloaded zip file and click "Properties". 
3. Check "Unblock" checkbox and click "OK", or simply click "Unblock" button depending on OS versions. 

&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src="https://i1.gallery.technet.s-msft.com/powershell-functions-for-16c5be31/image/file/142582/1/unblock.png" width="250">

4. Extract the zip file and copy "Microsoft.Xrm.Data.PowerShell" folder to one of the following folders:
  * %USERPROFILE%\Documents\WindowsPowerShell\Modules
  * %WINDIR%\System32\WindowsPowerShell\v1.0\Modules
Following image shows this module copied to User Profile. If you want anyone to use the module on the computer, copy them to System Wide PowerShell module folder instead. If you do not have the folder, you can manually create them.

&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src="https://i1.gallery.technet.s-msft.com/scriptcenter/powershell-functions-for-16c5be31/image/file/142578/1/individual.png" width="275">

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
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* *The module requires PowerShell v4.0.*

### How Microsoft.Xrm.Data.Powershell works
Microsoft.Xrm.Data.Powershell module exposes many functions, but you can use Connect-CrmOnlineDiscovery, Connect-CrmOnPremDiscovery to connect to any CRM organization by using Discovery Service. Use Connect-CrmOnline function for Azure Automation. By executing these function, it creates $conn global variable. Any other functions which needs to connect to the CRM Organization takes connection parameter. You can explicitly specify the connection by using -conn parameter, but if you omit the connection, functions retrieve connection from global variable.

Alternatively, you can create multiple connection objects and pass them into each function under the –conn parameter.

### Example
This example shows how to create connection and do CRUD operation as well as manipulate System Settings.
1. Run following command to connect to Dynamics CRM Organization via  the Xrm Tooling GUI.
```PowerShell
# Online
Connect-CrmOnlineDiscovery -InteractiveMode
# OnPrem
Connect-CrmOnPremDiscovery -InteractiveMode
# Azure Automation
Connect-CrmOnline -Credential $cred -ServerUrl "https://<org>.crm.dynamics.com"
```

For Azure Automation, write all scripts inside inlinescript block as Runbook or use PowerShell type.

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
This module is implemented by Sean McNellis and Kenichiro Nakamura.
 
<a href="https://twitter.com/seanmcne" target="_blank">Sean McNellis</a>, Sr. Premier Field Engineer, is based out of North America and works supporting Dynamics CRM customers.
Kenichiro Nakamura, Sr. Premier Mission Critical Specialist, is based out of Japan and works supporting PMC customers.

Blog (English): <a href="http://blogs.msdn.com/CrmInTheField" target="_blank">http://blogs.msdn.com/CrmInTheField</a>
Blog (Japanese): <a href="http://blogs.msdn.com/CrmJapan" target="_blank">http://blogs.msdn.com/CrmJapan</a>
Twitter: [@pfedynamics](https://twitter.com/pfedynamics)

Refer to previous release information [here](https://gallery.technet.microsoft.com/PowerShell-functions-for-16c5be31).
