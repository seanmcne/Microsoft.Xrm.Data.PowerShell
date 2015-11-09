####Prerequisites
To run this sample, you need to install following modules which is Office 365 PowerShell modules.<br/>
Microsoft Online Services Sign In Assistant (32/64 bit)<br/>
http://www.microsoft.com/en-us/download/details.aspx?id=41950<br/>
Azure AD Module for PowerShell (32 bit)<br/>
http://go.microsoft.com/fwlink/p/?linkid=236298<br/>
Azure AD Module for PowerShell (64 bit)<br/>
http://go.microsoft.com/fwlink/p/?linkid=236297<br/>
Please see the following link for more information about PowerShell for Office 365 <br/>
http://powershell.office.com/<br/>

####Sample Overview
This script works against Microsoft Dynamics CRM Online environment only. If you use On-Premise edition, please find On-Premise version of the script. (AddCrmUsersFromCSV.ps1)
This sample includes:
-	Create Office 365 Group.
-	Read CSV file as source data
o	The CSV file contains both Office 365 and Dynamics CRM Online related information. You can find CSV file in the same folder of this ReadMe.
-	Create Office 365 User, add the user to the group, and assign Dynamics CRM Professional License.
-	Confirm CRM user synchronization from Office 365.
-	Move a CRM user to specified Business Unit.
-	Assign a specified security role to a CRM user.

####How to use the sample
1.	Open AddCrmOLUserFromCSV.ps1 file with PowerShell ISE or any preferred text editor.
2.	Modify following parameters.<br/>
  a.	$msolUser: A username to connect to Office 365.<br/>
  b.	$msolPassword: A password of above user.<br/>
  c.	$msolDomainName: A domain name of Office 365. This is used to determine CRM License name.<br/>
  d.	$msolGroupName: Office 365 Group name, to which you want to add users.<br/>
  e.	$csvPath: A path to CSV file.<br/>
  f.	$crmAdminUser: A username to connect to Dynamics CRM Online. Can be same as $msolUser.<br/>
  g.	$crmAdminPassword: A passrod of above user.<br/>
3.	Save the ps1 file.
4.	Update CSV file to include appropriate data.
5.	Run the ps1 file via PowerShell.
6.	Connect to https://portal.office.com to see the result.
7.	Also connect to your Dynamics CRM Online organization to see the result.
