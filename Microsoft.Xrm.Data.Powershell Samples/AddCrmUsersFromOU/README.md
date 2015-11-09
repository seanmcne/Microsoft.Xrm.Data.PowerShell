####Sample Overview
This script works against Microsoft Dynamics CRM On-Premise environment only. If you use Online, please find Online version of the script. (AddCrmOLUsersFromCSV.ps1)
This sample includes:
-	Read AD Users from specified OU by using LDAP query.
-	Create CRM user under specified Business Unit.<br/>
  The script reads Department AD attribute value as Business Unit name. If no value specified, User will be created under default Business Unit.
-	Assign a specified security role to a CRM user.

####How to use the sample
1.	Open AddCrmUsersFromOU.ps1 file with PowerShell ISE or any preferred text editor.
2.	Modify following parameters.<br/>
  a.	$ouPath: LDAP to specify OU to get AD Users from.<br/>
  b.	$domain: AD Domain Name.<br/>
  c.	$organizationName: Dynamics CRM organization unique name.<br/>
  d.	$serverUrl: URL of Dynamics CRM Server.<br/>
  e.	$crmAdminUser: A username to connect to Dynamics CRM <br/>
  f.	$crmAdminPassword: A passrod of above user.<br/>
  g.	$securityRoleName: A SecurityRole name to assign to users.<br/>
3.	Save the ps1 file.
4.	Create or move AD users to the OU. Update Department attribute as needed.
5.	Run the ps1 file via PowerShell.
6.	Connect to your Dynamics CRM organization to see the result.
