﻿# Generated by: Sean McNellis (seanmcn)
#
# Copyright © Microsoft Corporation.  All Rights Reserved.
# This code released under the terms of the 
# Microsoft Public License (MS-PL, http://opensource.org/licenses/ms-pl.html.)
# Sample Code is provided for the purpose of illustration only and is not intended to be used in a production environment. 
# THIS SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, 
# INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE. 
# We grant You a nonexclusive, royalty-free right to use and modify the Sample Code and to reproduce and distribute the object code form of the Sample Code, provided that. 
# You agree: 
# (i) to not use Our name, logo, or trademarks to market Your software product in which the Sample Code is embedded; 
# (ii) to include a valid copyright notice on Your software product in which the Sample Code is embedded; 
# and (iii) to indemnify, hold harmless, and defend Us and Our suppliers from and against any claims or lawsuits, including attorneys’ fees, that arise or result from the use or distribution of the Sample Code 

# Script parameters #
#$crmAdminUser = "<CRM admin account. i.e.) administator@contoso.onmicrosoft.com>"
#$crmAdminPassword = ConvertTo-SecureString -String "<password>" -AsPlainText -Force

# Blank following parameters if you use Online environment.
#$organizationName = "<organization name> i.e.) contoso>"
#$serverUrl ="<CRM Server URL i.e.) http://crmserver:5555>"

$crmAdminUser = "kenakamu@crm2015training15.onmicrosoft.com"
$crmAdminPassword = ConvertTo-SecureString -String "Pa`$`$w0rd" -AsPlainText -Force

# Blank following parameters if you use Online environment.
$organizationName = ""
$serverUrl ="<CRM Server URL i.e.) http://crmserver:5555>"

$LCID = 1033
# Script parameters #

Write-Output "Connecting to CRM Online as $crmAdminUser"
$crmCred = New-Object System.Management.Automation.PSCredential ($crmAdminUser,$crmAdminPassword) 
Try
{
    # You can also use Get-CrmConnection to directly create connection.
    # See https://msdn.microsoft.com/en-us/library/dn756303.aspx for more detail.
    If($organizationName -eq "")
    {
        Connect-CrmOnlineDiscovery -Credential $crmCred -ErrorAction Stop
    }
    Else
    {
        $global:conn = Get-CrmConnection -OrganizationName $organizationName -ServerUrl $serverUrl -Credential $crmCred -ErrorAction Stop
    }
}
Catch
{
     throw
}

Write-Output "Retrieve user settings with LanguageId: $LCID"
$settings = Get-CrmRecords -EntityLogicalName usersettings -FilterAttribute uilanguageid -FilterOperator eq -FilterValue $LCID -Fields uilanguageid

Write-Output "Disable Langauge Pack"
Disable-CrmLanguagePack $LCID

Write-Output "Enable Language Pack"
Enable-CrmLanguagePack $LCID

Write-Output "Set user settings with LanguageId: $LCID"
$settings.CrmRecords | % { $_.uilanguageid = $LCID; Set-CrmUserSettings $_;}