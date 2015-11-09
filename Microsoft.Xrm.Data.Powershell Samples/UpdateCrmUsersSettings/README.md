####Sample Overview
This script works against Microsoft Dynamics CRM Online and On-Premise. Do not specify $serverUrl if you connect to Dynamics CRM Online. 
This sample includes:
-	Retrieve CRM Users
-	Update CRM User settings.

####How to use the sample
1.	Open UpdateCrmUsersSettings.ps1 file with PowerShell ISE or any preferred text editor.
2.	Modify following parameters.<br/>
  a.	$crmAdminUser: A username to connect to Dynamics CRM <br/>
  b.	$crmAdminPassword: A passrod of above user.<br/>
  c.	$organizationName: Dynamics CRM organization unique name. If you connect Dynamics CRM Online, do not specify any value for this.<br/>
  d.	$serverUrl: URL of Dynamics CRM Server. If you connect Dynamics CRM Online, do not specify any value for this.<br/>
  e.	$AdvancedFindStartupMode: Specify AdvancedFind mode. 1:simple, 2:detail.<br/>
  f.	$timeZoneCode: Specify TimeZoneCode for users. Use Get-CrmTimeZones to see all options.<br/>
  g.	$pagingLimit: Specify how many records per view. Value can be 25,50,75,100,250<br/>
  h.	$reportScriptErrorOption: Set options for Privacy tab.<br/>
  i.	$uiLanguageId: Set user’s UI Language LCID.<br/>
  j.	$transactionCurrencyName: Set user’s transaction currency by ISO name.<br/>
3.	Save the ps1 file.
4.	Run the ps1 file via PowerShell.
5.	Connect to your Dynamics CRM organization to see user option. It may take several seconds before actual settings are reflected.
