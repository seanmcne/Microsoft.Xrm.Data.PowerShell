####Sample Overview
This script works against Microsoft Dynamics CRM Online and On-Premise. <br/>
Do not specify $serverUrl if you connect to Dynamics CRM Online.
This sample includes:
-	Retrieve CRM User Settings which has UI Language setting of specified LCID
-	Deprovision specified LCID
-	Provision specified LCID
-	Update UI Language setting of CRM User Settings with original LCID

####How to use the sample
1.	Open UpdateCrmUsersSettings.ps1 file with PowerShell ISE or any preferred text editor.
2.	Modify following parameters.<br/>
  a.	$crmAdminUser: A username to connect to Dynamics CRM <br/>
  b.	$crmAdminPassword: A passrod of above user.<br/>
  c.	$organizationName: Dynamics CRM organization unique name. If you connect Dynamics CRM Online, do not specify any value for this.<br/>
  d.	$serverUrl: URL of Dynamics CRM Server. If you connect Dynamics CRM Online, do not specify any value for this.<br/>
  e.	$LCID: Specify Language ID for reprovision.<br/>
3.	Save the ps1 file.
4.	Run the ps1 file via PowerShell.
5.	Connect to your Dynamics CRM organization to see user option. It may take several seconds before actual settings are reflected.
