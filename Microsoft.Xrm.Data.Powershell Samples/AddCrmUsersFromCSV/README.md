####Prerequisites
You need to run this scripts as a domain user who has privilege to create AD user.

####Sample Overview
This script works against Microsoft Dynamics CRM On-Premise environment only. If you use Online, please find Online version of the script. (AddCrmOLUsersFromCSV.ps1)<br/>
This sample includes:
-	Read CSV file as source data<br/>
	The CSV file contains both Active Directory and Dynamics CRM related information. You can find CSV file in the same folder of this ReadMe.
-	Create AD user and enable the user.
-	Create CRM user under specified Business Unit.
-	Assign a specified security role to a CRM user.

####How to use the sample
1.	Open AddCrmUserFromCSV.ps1 file with PowerShell ISE or any preferred text editor.
2.	Modify following parameters.<br/>
  a.	$organizationName: Dynamics CRM organization unique name.<br/>
  b.	$serverUrl: URL of Dynamics CRM Server.<br/>
  c.	$csvPath: A path to CSV file.<br/>
  d.	$crmAdminUser: A username to connect to Dynamics CRM <br/>
  e.	$crmAdminPassword: A passrod of above user.<br/>
3.	Save the ps1 file.
4.	Update CSV file to include appropriate data.
5.	Run the ps1 file via PowerShell.
6.	Open Active Directory Users and Computers snap-in to see the result.
7.	Also connect to your Dynamics CRM organization to see the result.
