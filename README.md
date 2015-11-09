
# Microsoft.Xrm.Data.PowerShell

### Overview 
**Microsoft.Xrm.Data.Powershell.zip** contains one primary module, Microsoft.Xrm.Data.Powershell, but also relies on an included dll module Microsoft.Xrm.Tooling.CrmConnector.Powershell.  

**Microsoft.Xrm.Data.Powershell** 

This module uses the CRM connection from Microsoft.Xrm.Tooling.CrmConnector.Powershell and provides common functions to create, delete, query, and update data as well as functions for common tasks such as publishing, and manipulating System & CRM User Settings, etc. The module should function for both Dynamics CRM Online and On-Premise environment.  

**Microsoft.Xrm.Tooling.CrmConnector.Powershell**

This module comes from Dynamics CRM SDK and it exposes two functions, Get-CrmOrganizations and Get-CrmConnection. See the link for more detail.

[Use PowerShell cmdlets for XRM tooling to connect to CRM](https://technet.microsoft.com/en-us/library/dn689040.aspx)

###How to setup modules
<p>1. Download Microsoft.Xrm.Data.Powershell.zip.</p> 
<p>2. Right click the downloaded zip file and click "Properties". </p> 
<p>3. Check "Unblock" checkbox and click "OK", or simply click "Unblock" button depending on OS versions. </p> 
![Image of Unblock](https://i1.gallery.technet.s-msft.com/powershell-functions-for-16c5be31/image/file/142582/1/unblock.png)
<p>4. Extract the zip file and copy "Microsoft.Xrm.Data.PowerShell" folder to one of the following folders:<br/>
  * %USERPROFILE%\Documents\WindowsPowerShell\Modules<br/>
  * %WINDIR%\System32\WindowsPowerShell\v1.0\Modules<br/>
Following image shows this module copied to User Profile. If you want anyone to use the module on the computer, copy them to System Wide PowerShell module folder instead. If you do not have the folder, you can manually create them.</p> 
![Image of individual](https://i1.gallery.technet.s-msft.com/scriptcenter/powershell-functions-for-16c5be31/image/file/142578/1/individual.png)
<p>5. As this module is not signed, you may need to change Execution Policy to load the module. You can do so by executing following command. </p> 
```PowerShell
 Set-ExecutionPolicy –ExecutionPolicy RemoteSigned –Scope CurrentUser
```
Please refer to 
[Set-ExecutionPolicy](https://technet.microsoft.com/en-us/library/ee176961.aspx) 
for more information.
<p>6. Open PowerShell and run following command to load the module. </p> 
```PowerShell
# Import Micrsoft.Xrm.Data.Powershell module 
Import-Module Microsoft.Xrm.Data.Powershell
```
*The module requires PowerShell v4.0.

###How module works
Microsoft.Xrm.Data.Powershell module exposes many functions, but you can use Connect-CrmOnlineDiscovery to connect to any CRM Online organization. By executing the function, it creates $conn global variable. Any other functions which needs to connect to the CRM Organization takes connection parameter. You can explicitly specify the connection by using -conn parameter, but if you omit the connection, functions retrieve connection from global variable.

If you are using On-Premise, use following command instead to create connection. This function creates connection by prompting you GUI signin page.  Alternatively, you can create multiple connection objects and pass them into each function under the –conn parameter.

```PowerShell
$global:conn = Get-CrmConnection -InteractiveMode
```
See more detail about [Get-CrmConnection](https://technet.microsoft.com/en-us/library/dn756303.aspx) function.<br/>

###Example
This example shows how to create connection and do CRUD operation as well as manipulate System Settings.
<p>1. Run following command to connect to Dynamics CRM Organization via  the Xrm Tooling GUI.</p>
```PowerShell
Connect-CrmOnlineDiscovery -InteractiveMode
```
<p>2. Run following command to test CRUD.</p>
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
<p>3. Run following command to manipulate SystemSettings.</p>
```PowerShell
# Display the current setting 
Get-CrmSystemSettings -conn $conn -ShowDisplayName 
 
# Change the PricingDecimalPrecision system setting from 0 to 1 
Set-CrmSystemSettings -conn $conn -PricingDecimalPrecision 1 
 
# Display the current setting 
Get-CrmSystemSettings -conn $conn -ShowDisplayName
```
###How to get command details
Each command has detail explanation.
<p>1. Run following command to get all commands.</p>
```PowerShell
Get-Command *crm*
```
<p>2. Run following command to get help.</p>
```PowerShell
Get-Help New-CrmRecord -Detailed
```
####About Authors
This module is implemented by Sean McNellis and Kenichiro Nakamura.
 
Sean McNellis, Sr. Premier Field Engineer, is based out of North America and works supporting Dynamics CRM customers.<br/>
Kenichiro Nakamura, Sr. Premier Mission Critical Specialist, is based out of Japan and works supporting PMC customers.
 
Blog (English): [http://blogs.msdn.com/CrmInTheField](http://blogs.msdn.com/CrmInTheField) <br/>
Blog (Japanese): [http://blogs.msdn.com/CrmJapan](http://blogs.msdn.com/CrmJapan) <br/>
Twitter: @pfedynamics

####Release History
**Version 1.6 - 2015/11/2**

Added following function.
- Approve-CrmEmailAddress to approve EmailAddress for User or Queue.

Updating the following function.
- Remove-CrmRecord to accept CrmRecord object from pipeline so that you can do like (Get-CrmRecords account).CrmRecords | Remove-CrmRecord to remove all account records.

Fixed following issues
- Import-CrmSolution issue when importing multiple solutions.
- Set-CrmRecord didn't work as expected when passing $true/$false for bool field.
- Set-CrmRecord didn't update SystemForm correctly.
- Test-XrmTimerStop had wrong scope for variable.
- Added CrmError check for CRUD functions.

**Version 1.5 - 2015/10/24**

Updated the following functions.
- Get-CrmSystemSettings to retrieve all values from Organization entity.
- Set-CrmSystemSettings to accept more parameters. Please note that not all parameters work in all Dynamics CRM versions.
- Set-CrmRecordOwner to let you assign CRM record to a Team in addition to a User. We changed the parameter name from UserId to PrincipalId but it has UserId as alias.

Fixed following issues
- Export-CrmSolution didn't work as expected when using wildcard searches.
- Set-CrmSystemSettings fails when passing $conn parameter explicitly.

**Version 1.4 - 2015/9/30**

Added following function.
- Get-CrmRecordsCount

Fixed following issue.
- Get-CrmRecordsByFetchXml with AllRecords parameter didn't work as expected, which causes the same issue for Get-CrmRecords as it uses the function underneath.
- Connect-CrmOnlineDiscovery: Display retrieved CRM organizations by friendly name.

**Version 1.3 - 2015/9/23**

Added following functions.
- Export-CrmSolutionTranslation
- Import-CrmSolutionTranslation

Fixed following issues.
- Revised Import-CrmSolution and Export-CrmSolution help comments.
- Return nothing when Import-CrmSolution succeeded.
- Return nothing when Publish-CrmAllCustomization succeeded.

**Version 1.2 - 2015/9/21**

Added following functions.
- Get-CrmAllLanguagePacks
- Enable-CrmLanguagePack
- Disable-CrmLanguagePack

**Version 1.1 - 2015/9/19**

Fixed following issues.
- Revise Get-CrmRecords Columns parameter name to Fields to addhere naming convension.
- The Import-CrmSolution  function didn't throw an error even if it fails to import a solution. We added import status check to confirm the status and throw an error if it fails. We also include logic to check if solution is managed or not to handle publishing properly.
- Revised Set-CrmUserSettings help comment.

**Version 1.0 - 2015/9/16**

Initial release.
