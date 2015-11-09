####Sample Overview
This script demonstrates how to work with multiple CRM organizations for both Online and On-Premise.<br/>
This sample includes:
-	Read connection information from CSV.
-	Create connections by using loaded information.
-	Add an Account record to all organizations.

####How to use the sample
1.	Open connectionssource.csv file and modify connection information. <br/>
For On-Premise, make sure to enter ServerUrl, whereas if online, make sure to enter Deployment Region. <br/>
If you are unsure about the region, then you can run Get-CrmOrganizations –Credential $cred –OnlineType Office365 to list all CRM Organizations the credential belongs to, and you can take DiscoveryServerShortName as DeploymentRegion.
2.	Modify following parameters.<br/>
  a.	$connectionsSourceCsvPath: Path to the csv.<br/>
3.	Save the ps1 file.
4.	Run the ps1 file via PowerShell.
5.	Connect to your Dynamics CRM organizations to see the result.
