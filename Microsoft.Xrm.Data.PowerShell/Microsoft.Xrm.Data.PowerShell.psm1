
### https://msdn.microsoft.com/en-us/library/microsoft.xrm.tooling.connector.crmserviceclient_methods(v=crm.6).aspx ###
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

<#
.Synopsis
   A means of running multiple instances of a cmdlet/function/scriptblock
.DESCRIPTION
   This function allows you to provide a cmdlet, function or script block with a set of data to allow multithreading.
.EXAMPLE
   $sb = [scriptblock] {param($system) gwmi win32_operatingsystem -ComputerName $system | select csname,caption}
   $servers = Get-Content servers.txt
   $rtn = Invoke-Async -Set $server -SetParam system  -ScriptBlock $sb
.EXAMPLE
   $servers = Get-Content servers.txt
   $rtn = Invoke-Async -Set $servers -SetParam computername -Params @{count=1} -Cmdlet Test-Connection -ThreadCount 50 
.INPUTS
   
.OUTPUTS
   Determined by the provided cmdlet, function or scriptblock.
.NOTES
    This can often times eat up a lot of memory due in part to how some cmdlets work. Test-Connection is a good example of this. 
    Although it is not a good idea to manually run the garbage collector it might be needed in some cases and can be run like so:
    [gc]::Collect()
#>

#Get-CrmOrganizations && Get-CrmConnection
function Connect-CrmOnlineDiscovery{

<#
 .SYNOPSIS
 Retrieves all CRM Online Organization you belong to, let you select which organization to login, then returns connection information.

 .DESCRIPTION
 The Connect-CrmOnlineDiscovery cmdlet lets you retrieves all CRM Online Organization you belong to, let you select which organization to login, then returns connection information.
 
 You can use Get-Credential to create Credential information, or you can simply invoke Connect-CrmOnlineDiscovery which prompts you to enter username/password.

 .PARAMETER Credential
 A PS-Credential. You can invoke Get-Credential.

 .EXAMPLE
 Connect-CrmOnlineDiscovery
 Supply values for the following parameters:
 
 IsReady                        : True
 IsBatchOperationsAvailable     : True
 OrganizationServiceProxy       : Microsoft.Xrm.Tooling.Connector.CrmWebSvc+ManagedTokenOrganizationServiceProxy
 LastCrmError                   : 
 LastCrmException               : 
 CrmConnectOrgUriActual         : https://contoso.api.crm.dynamics.com/XRMServices/2011/Organization.svc
 ConnectedOrgFriendlyName       : contoso
 ConnectedOrgUniqueName         : contoso
 ConnectedOrgPublishedEndpoints : {[WebApplication, https://contoso.crm.dynamics.com/], [OrganizationService, 
                                  https://contoso.api.crm.dynamics.com/XRMServices/2011/Organization.svc], 
                                  [OrganizationDataService, 
                                  https://contoso.api.crm.dynamics.com/XRMServices/2011/OrganizationData.svc]}
 ConnectionLockObject           : System.Object
 ConnectedOrgVersion            : 7.1.0.1086

 This example prompts you to enter username/password, displays all CRM organization, and returns connection.

 .EXAMPLE
 PS C:\>$cred = Get-Credential
 PS C:\>Connect-CrmOnlineDiscovery $cred
 
 IsReady                        : True
 IsBatchOperationsAvailable     : True
 OrganizationServiceProxy       : Microsoft.Xrm.Tooling.Connector.CrmWebSvc+ManagedTokenOrganizationServiceProxy
 LastCrmError                   : 
 LastCrmException               : 
 CrmConnectOrgUriActual         : https://contoso.api.crm.dynamics.com/XRMServices/2011/Organization.svc
 ConnectedOrgFriendlyName       : contoso
 ConnectedOrgUniqueName         : contoso
 ConnectedOrgPublishedEndpoints : {[WebApplication, https://contoso.crm.dynamics.com/], [OrganizationService, 
                                  https://contoso.api.crm.dynamics.com/XRMServices/2011/Organization.svc], 
                                  [OrganizationDataService, 
                                  https://contoso.api.crm.dynamics.com/XRMServices/2011/OrganizationData.svc]}
 ConnectionLockObject           : System.Object
 ConnectedOrgVersion            : 7.1.0.1086

 This example displays all CRM organization, and returns connection.
 
 .EXAMPLE
 PS C:\>Connect-CrmOnlineDiscovery -InteractiveMode
 
 IsReady                        : True
 IsBatchOperationsAvailable     : True
 OrganizationServiceProxy       : Microsoft.Xrm.Tooling.Connector.CrmWebSvc+ManagedTokenOrganizationServiceProxy
 LastCrmError                   : 
 LastCrmException               : 
 CrmConnectOrgUriActual         : https://contoso.api.crm.dynamics.com/XRMServices/2011/Organization.svc
 ConnectedOrgFriendlyName       : contoso
 ConnectedOrgUniqueName         : contoso
 ConnectedOrgPublishedEndpoints : {[WebApplication, https://contoso.crm.dynamics.com/], [OrganizationService, 
                                  https://contoso.api.crm.dynamics.com/XRMServices/2011/Organization.svc], 
                                  [OrganizationDataService, 
                                  https://contoso.api.crm.dynamics.com/XRMServices/2011/OrganizationData.svc]}
 ConnectionLockObject           : System.Object
 ConnectedOrgVersion            : 7.1.0.1086

 This example shows how to use -InteractiveMode switch. By specifying the switch, you can login via GUI tool.
#>

    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false)]
        [PSCredential]$Credential, 
        [Parameter(Mandatory=$false)]
        [switch]$UseCTP,
        [Parameter(Mandatory=$false)]
        [switch]$InteractiveMode
    )
    
    #Need to change when XrmTooling is updated to remove -OrganizationName parameter
    if($InteractiveMode)
    {
        $global:conn = Get-CrmConnection -InteractiveMode
        
        Write-Verbose "You are now connected and may run any of the CRM Commands."
        return $global:conn 
    }
    else
    {
        $onlineType = "Office365"
        if($UseCTP)
        {
            $onlineType = "LiveID"
        }
        if($Credential -eq $null -And !$Interactive)
        {
            $Credential = Get-Credential
        }
        $crmOrganizations = Get-CrmOrganizations -Credential $Credential -OnLineType $onlineType -Verbose 
        $i = 0
          
        if($crmOrganizations.Count -gt 0)
        {    

	    if($crmOrganizations.Count -eq 1)
            {
                $orgNumber = 0
            }
	    else
            {
				$crmOrganizations = $crmOrganizations | sort-object FriendlyName;
                foreach($crmOrganization in $crmOrganizations)
                {   $friendlyName = $crmOrganization.FriendlyName

                    $message = "[$i] $friendlyName (" + $crmOrganization.WebApplicationUrl + ")"
                    Write-Host $message 
                    $i++
                }
                $orgNumber = Read-Host "Select CRM Organization"
    
                Write-Verbose ($crmOrganizations[$orgNumber]).UniqueName
	    }
            $global:conn = Get-CrmConnection -Credential $Credential -DeploymentRegion $crmOrganizations[$orgNumber].DiscoveryServerShortname -OnLineType $onlineType -OrganizationName ($crmOrganizations[$orgNumber]).UniqueName

            Write-Verbose "You are now connected and may run any of the CRM Commands."
            return $global:conn    
        }
    }
}

### Core CRUD Cmdlets ###
#CreateNewRecord
function New-CrmRecord{

<#
 .SYNOPSIS
 Creates a new CRM record by specifying field name/value set, and returns record guid.

 .DESCRIPTION
 The New-CrmRecord cmdlet lets you create a record to your CRM organization. 
 Use @{"field logical name"="value"} syntax to create Fields, and make sure you specify correct type of value for the field. 

 You can use Get-CrmEntityAttributeMetadata cmdlet and check AttributeType to see the field type. In addition, for CRM specific types, you can use New-CrmMoney, New-CrmOptionSetValue or New-CrmEntityReference cmdlets.

 .PARAMETER conn
 A connection to your CRM organizatoin. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .PARAMETER EntityLogicalName
 A logicalname for an Entity to create. i.e.)accout, contact, lead, etc..

 .PARAMETER Fields
 A List of field name/value pair. Use @{"field logical name"="value"} syntax to create Fields, and make sure you specify correct type of value for the field. 
 You can use Get-CrmEntityAttributeMetadata cmdlet and check AttributeType to see the field type. In addition, for CRM specific types, you can use New-CrmMoney, New-CrmOptionSetValue or New-CrmEntityReference cmdlets.

 .EXAMPLE
 New-CrmRecord -conn $conn -EntityLogicalName account -Fields @{"name"="account name";"telephone1"="555-5555"}
 Guid
 ----
 57bd1c45-2b17-e511-80dc-c4346bc4fc6c  

 This example creates an account record and set value for account and telephone1 fields.

 .EXAMPLE
 New-CrmRecord account @{"name"="account name";"industrycode"=New-CrmOptionSetValue -Value 1}
 Guid
 ----
 57bd1c45-2b17-e511-80dc-c4346bc4fc6c  
 
 This example creates an account record by specifying OptionSetValue to IndustryCode field by omitting parameter names.
 When ommiting parameter names, you do not provide $conn, cmdlets automatically finds it.

 .EXAMPLE
 New-CrmRecord account @{"name"="account name";"overriddencreatedon"=[datetime]"2000-01-01"}
 Guid
 ----
 57bd1c45-2b17-e511-80dc-c4346bc4fc6c  
 
 This example creates an account record by specifying past date for CreatedOn. You can create DateTime object by casting like this Example or use Get-Date cmdlets.

 .EXAMPLE 
 PS C:\>$parentId = New-CrmRecord account @{"name"="parent account name"}
 
 PS C:\>$parentReference = New-CrmEntityReference -EntityLogicalName account -Id $parentId
 
 PS C:\>$childId = New-CrmRecord account @{"name"="child account name";"parentaccountid"=$parentReference}
 Guid
 ----
 59bd1c45-2b17-e511-80dc-c4346bc4fc6c

 This example creates an account record, then assign the created account as Parent Account field to second account record.
#>

    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1)]
        [string]$EntityLogicalName,
        [parameter(Mandatory=$true, Position=2)]
        [hashtable]$Fields
    )

    if($conn -eq $null)
    {
        $connobj = Get-Variable conn -Scope global -ErrorAction SilentlyContinue
        if($connobj.Value -eq $null)
        {
            Write-Warning 'You need to create Connect to CRM Organization. Use Get-CrmConnection to create it.'
            break;
        }
        else
        {
            $conn = $connobj.Value
        }
    }

    $newfields = New-Object 'System.Collections.Generic.Dictionary[[String], [Microsoft.Xrm.Tooling.Connector.CrmDataTypeWrapper]]'
    
    foreach($field in $Fields.GetEnumerator())
    {  
        $newfield = New-Object -TypeName 'Microsoft.Xrm.Tooling.Connector.CrmDataTypeWrapper'
        
        switch($field.Value.GetType().Name)
        {
            "Boolean" {
                $newfield.Type = [Microsoft.Xrm.Tooling.Connector.CrmFieldType]::CrmBoolean
                break              
            }
            "DateTime" {
                $newfield.Type = [Microsoft.Xrm.Tooling.Connector.CrmFieldType]::CrmDateTime
                break
            }
            "Decimal" {
                $newfield.Type = [Microsoft.Xrm.Tooling.Connector.CrmFieldType]::CrmDecimal
                break
            }
            "Single" {
                $newfield.Type = [Microsoft.Xrm.Tooling.Connector.CrmFieldType]::CrmFloat
                break
            }
            "Money" {
                $newfield.Type = [Microsoft.Xrm.Tooling.Connector.CrmFieldType]::Raw
                break
            }
            "Int32" {
                $newfield.Type = [Microsoft.Xrm.Tooling.Connector.CrmFieldType]::CrmNumber
                break
            }
            "EntityReference" {
                $newfield.Type = [Microsoft.Xrm.Tooling.Connector.CrmFieldType]::Raw
                break
            }
            "OptionSetValue" {
                $newfield.Type = [Microsoft.Xrm.Tooling.Connector.CrmFieldType]::Raw
                break
            }
            "String" {
                $newfield.Type = [Microsoft.Xrm.Tooling.Connector.CrmFieldType]::String
                break
            }
        }
        
        $newfield.Value = $field.Value
        $newfields.Add($field.Key, $newfield)
    }

    try
    {        
        $result = $conn.CreateNewRecord($EntityLogicalName, $newfields, $null, $false, [Guid]::Empty)
        if(!$result)
        {
            return $conn.LastCrmError
        }
    }
    catch
    {
        return $conn.LastCrmException        
    }

    return $result
}

#GetEntityDataById 
function Get-CrmRecord{

<#
 .SYNOPSIS
 Retrieves a CRM record by specifying EntityLogicalName, record's Id (guid) and field names.

 .DESCRIPTION
 The Get-CrmRecord cmdlet lets you retrieve a record from your CRM organization. The retireve results has two properties for each field like name and name_Property. 
 The field with "_Property" contains field's logicalname and its Type information, and another field contains "readable" value, which is either FormattedValue or extracted value from CRM special type like EntityReference.
 
 The output also contains logical name of the Entity and original field which contains raw data. You can ignore these two fields.

 You can specify fields as fieldname1,fieldname2,fieldname3 syntax or * to retrieve all fields (not recommended for performance reason.)
 You can use Get-CrmEntityAttributes cmdlet to see all fields logicalname. The retrieved data can be passed to several cmdlets like Set-CrmRecord, Removed-CrmRecord to further process it.

 .PARAMETER conn
 A connection to your CRM organizatoin. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .PARAMETER EntityLogicalName
 A logicalname for an Entity to retrieve. i.e.)accout, contact, lead, etc..

 .PARAMETER Id
 An Id (guid) of the record

 .PARAMETER Fields
 A List of field logicalnames. Use "fieldname1, fieldname2, fieldname3" syntax to speficy Fields, or ues "*" to retrieve all fields (not recommended for performance reason.)

 .EXAMPLE
 Get-CrmRecord -conn $conn -EntityLogicalName account -Id be02caab-6c16-e511-80d6-c4346bc43dc0 -Fields name,accountnumber
 name_Property      : [name, A. Datum Corporation (sample)]
 name               : A. Datum Corporation (sample)
 accountid_Property : [accountid, be02caab-6c16-e511-80d6-c4346bc43dc0]
 accountid          : be02caab-6c16-e511-80d6-c4346bc43dc0
 original           : {[name_Property, [name, A. Datum Corporation (sample)]], [name, A. Datum Corporation (sample)], [accountid_Property, [accountid, 
                      be02caab-6c16-e511-80d6-c4346bc43dc0]], [accountid, be02caab-6c16-e511-80d6-c4346bc43dc0]}
 logicalname        : account  

 This example retrieves an account record with name and accountnumber fields

 .EXAMPLE
 Get-CrmRecord account be02caab-6c16-e511-80d6-c4346bc43dc0 primarycontactid,statuscode
 accountid_Property        : [accountid, be02caab-6c16-e511-80d6-c4346bc43dc0]
 accountid                 : be02caab-6c16-e511-80d6-c4346bc43dc0
 primarycontactid_Property : [primarycontactid, Microsoft.Xrm.Sdk.EntityReference]
 primarycontactid          : Rene Valdes (sample)
 statuscode_Property       : [statuscode, Microsoft.Xrm.Sdk.OptionSetValue]
 statuscode                : Active
 original                  : {[accountid_Property, [accountid, be02caab-6c16-e511-80d6-c4346bc43dc0]], [accountid, be02caab-6c16-e511-80d6-c4346bc43dc0], 
                             [primarycontactid_Property, [primarycontactid, Microsoft.Xrm.Sdk.EntityReference]], [primarycontactid, Microsoft.Xrm.Sdk.EntityReference]...}
 logicalname               : account
 
 This example retrieves an account record with PrimaryContact and StateCode fields by omitting parameter names.
 When ommiting parameter names, you do not provide $conn, cmdlets automatically finds it.

 .EXAMPLE
 Get-CrmRecord account 5b2fe1f3-b503-e511-80d4-c4346bc43dc0 *
 address2_addresstypecode_Property    : [address2_addresstypecode, Microsoft.Xrm.Sdk.OptionSetValue]
 address2_addresstypecode             : Default Value
 merged_Property                      : [merged, False]
 merged                               : No
 statecode_Property                   : [statecode, Microsoft.Xrm.Sdk.OptionSetValue]
 statecode                            : Active
 emailaddress1_Property               : [emailaddress1, someone9@example.com]
 emailaddress1                        : someone9@example.com
 ......

 This example retrieve an account record with all fields (which have values).
#>

    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1)]
        [string]$EntityLogicalName,
        [parameter(Mandatory=$true, Position=2)]
        [guid]$Id,
        [parameter(Mandatory=$true, Position=3)]
        [string[]]$Fields
    )

    if($conn -eq $null)
    {
        $connobj = Get-Variable conn -Scope global -ErrorAction SilentlyContinue
        if($connobj.Value -eq $null)
        {
            Write-Warning 'You need to create Connect to CRM Organization. Use Get-CrmConnection to create it.'
            break;
        }
        else
        {
            $conn = $connobj.Value
        }
    }

    if($Fields -eq "*")
    {
        [Collections.Generic.List[String]]$x = $null
    }
    else
    {
        [Collections.Generic.List[String]]$x = $Fields;
    }

    try
    {
        $record = $conn.GetEntityDataById($EntityLogicalName, $Id, $x, [Guid]::Empty)
    }
    catch
    {
        return $conn.LastCrmException        
    }    
    
    if($record -eq $null)
    {
        $error = "Record Id: " + $Id + "Does Not Exist" 
        return $error
    }
        
    $psobj = New-Object -TypeName System.Management.Automation.PSObject
        
    foreach($att in $record.GetEnumerator())
    {
        if($att.Value -is [Microsoft.Xrm.Sdk.EntityReference])
        {
            Add-Member -InputObject $psobj -MemberType NoteProperty -Name $att.Key -Value $att.Value.Name
        }
		elseif($att.Value -is [Microsoft.Xrm.Sdk.AliasedValue])
        {
			Add-Member -InputObject $psobj -MemberType NoteProperty -Name $att.Key -Value $att.Value.Value
        }
        else
        {
            Add-Member -InputObject $psobj -MemberType NoteProperty -Name $att.Key -Value $att.Value
        }
    }  

    Add-Member -InputObject $psobj -MemberType NoteProperty -Name "original" -Value $record
    Add-Member -InputObject $psobj -MemberType NoteProperty -Name "logicalname" -Value $EntityLogicalName

    return $psobj
}

#UpdateEntity 
function Set-CrmRecord{

<#
 .SYNOPSIS
 Updates a new CRM record by specifying EntityLogicalName, record's Id (guid) and field name/value sets.

 .DESCRIPTION
 The Set-CrmRecord cmdlet lets you update a record of your CRM organization. 

 There are two ways to update a record.
 
 1. Pass EntityLogicalName and record's Id and field name/value set. Use @{"field logical name"="value"} syntax to update Fields, and make sure you specify correct type of value for the field. 
 You can use Get-CrmEntityAttributeMetadata cmdlet and check AttributeType to see the field type. In addition, for CRM specific types, you can use New-CrmMoney, New-CrmOptionSetValue or New-CrmEntityReference cmdlets.

 2. Get a record object by using Get-CrmRecord/Get-CrmRecords cmdlets, then assign new value for fields. if a field you want to assign value is not available, you need to use first method to update the field.

 .PARAMETER conn
 A connection to your CRM organizatoin. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .PARAMETER CrmRecord
 A record object which is obtained via Get-CrmRecord/Get-CrmRecords. When you pass CrmRecord, then you don't use EntityLogicalName/Id.

 .PARAMETER EntityLogicalName
 A logicalname for an Entity to update. i.e.)accout, contact, lead, etc..

 .PARAMETER Id
 An Id (guid) of the record

 .PARAMETER Fields
 A List of field name/value pair. Use @{"field logical name"="value"} syntax to create Fields, and make sure you specify correct type of value for the field. 
 You can use Get-CrmEntityAttributeMetadata cmdlet and check AttributeType to see the field type. In addition, for CRM specific types, you can use New-CrmMoney, New-CrmOptionSetValue or New-CrmEntityReference cmdlets.

 .EXAMPLE
 Set-CrmRecord -conn $conn -EntityLogicalName account -Id 52a17637-5617-e511-80dc-c4346bc4fc6c -Fields @{"name"="updated name";"telephone1"="555-5555"}

 This example updates an account record by using Id and new value for name/telephone1 fields.
 Though any of field value is same as current record, it still tries to update the field.

 .EXAMPLE
 Set-CrmRecord account 52a17637-5617-e511-80dc-c4346bc4fc6c @{"industrycode"=New-CrmOptionSetValue -Value 1}
  
 This example updates an account record by using Id and IndustryCode filed by specifying OptionSetValue.
 When ommiting parameter names, you do not provide $conn, cmdlets automatically finds it.

 .EXAMPLE 
 PS C:\>$account = Get-CrmRecord account b202caab-6c16-e511-80d6-c4346bc43dc0 name

 PS C:\>$account
 name_Property      : [name, Adventure Works (sample)]
 name               : Adventure Works (sample)
 accountid_Property : [accountid, b202caab-6c16-e511-80d6-c4346bc43dc0]
 accountid          : b202caab-6c16-e511-80d6-c4346bc43dc0
 original           : {[name_Property, [name, Adventure Works (sample)]], [name, Adventure Works (sample)], [accountid_Property, [accountid, 
                      b202caab-6c16-e511-80d6-c4346bc43dc0]], [accountid, b202caab-6c16-e511-80d6-c4346bc43dc0]}
 logicalname        : account

 PS C:\>$account.name = $account.name + " updated!"

 PS C:\>Set-CrmRecord $account

 PS C:\>Get-CrmRecord account b202caab-6c16-e511-80d6-c4346bc43dc0 name
 name_Property      : [name, Adventure Works (sample) updated!]
 name               : Adventure Works (sample)updated!
 accountid_Property : [accountid, b202caab-6c16-e511-80d6-c4346bc43dc0]
 accountid          : b202caab-6c16-e511-80d6-c4346bc43dc0
 original           : {[name_Property, [name, Adventure Works (sample) updated!]], [name, Adventure Works (sample) updated!], [accountid_Property, [accountid, 
                      b202caab-6c16-e511-80d6-c4346bc43dc0]], [accountid, b202caab-6c16-e511-80d6-c4346bc43dc0]}
 logicalname        : account

 This example retrieves and store an account record to $account object, then assign new value for name field. Then update it by using Set-CrmRecord cmdlet.
 Finally retrieves it again to display the result. 
 When ommiting parameter names, you do not provide $conn, cmdlets automatically finds it.

 .EXAMPLE 
 PS C:\>$fetch = @"
 <fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="false" no-lock="true">
    <entity name="account">
      <attribute name="name" />
      </entity>
  </fetch>
 "@

 PS C:\>(Get-CrmRecordsByFetch $fetch).CrmRecords | % { $_.name = $_.name + " updated!"; Set-CrmRecord -CrmRecord $_}
 
 This example retrieves and stores account records by using FetchXML and pipe results (CrmRecords). In the next pipe, it does foreach operation (%) and assign new value for name. Then it updates each record using Set-CrmRecord.

#>

    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1, ParameterSetName="CrmRecord")]
        [PSObject]$CrmRecord,
        [parameter(Mandatory=$true, Position=1, ParameterSetName="Fields")]
        [string]$EntityLogicalName,
        [parameter(Mandatory=$true, Position=2, ParameterSetName="Fields")]
        [guid]$Id,
        [parameter(Mandatory=$true, Position=3, ParameterSetName="Fields")]
        [hashtable]$Fields
    )

    if($conn -eq $null)
    {
        $connobj = Get-Variable conn -Scope global -ErrorAction SilentlyContinue
        if($connobj.Value -eq $null)
        {
            Write-Warning 'You need to create Connect to CRM Organization. Use Get-CrmConnection to create it.'
            break;
        }
        else
        {
            $conn = $connobj.Value
        }
    }
    
    if($CrmRecord -ne $null)
    { 
        $entityLogicalName = $CrmRecord.logicalname
    }
    else
    {
        $entityLogicalName = $EntityLogicalName
    }
        
    # Some Entity has different pattern for id name.
    if($entityLogicalName -eq "usersettings")
    {
        $primaryKeyField = "systemuserid"
    }
    elseif($entityLogicalName -eq "systemform")
    {
        $primaryKeyField = "formid"
    }
    elseif($entityLogicalName -in ("opportunityclose","socialactivity","campaignresponse","letter","orderclose","appointment","recurringappointmentmaster","fax","email","activitypointer","incidentresolution","bulkoperation","quoteclose","task","campaignactivity","serviceappointment","phonecall"))
    {
        $primaryKeyField = "activityid"
    }
    else
    {
        $primaryKeyField = $entityLogicalName + "id"
    }

    $newfields = New-Object 'System.Collections.Generic.Dictionary[[String], [Microsoft.Xrm.Tooling.Connector.CrmDataTypeWrapper]]'
    
    if($CrmRecord -ne $null)
    {                
        $originalRecord = $CrmRecord.original        
        $Id = $originalRecord[$primaryKeyField]
        
        foreach($crmFieldKey in ($CrmRecord | Get-Member -MemberType NoteProperty).Name)
        {
            $crmFieldValue = $CrmRecord.($crmFieldKey)
            if(($crmFieldKey -eq "original") -or ($crmFieldKey -eq "logicalname") -or ($crmFieldKey -like "*_Property"))
            {
                continue
            }
            elseif($originalRecord[$crmFieldKey+"_Property"].Value -is [bool])
            {
                if($crmFieldValue -is [Int32])
                {
                    if(($originalRecord[$crmFieldKey+"_Property"].Value -and $crmFieldValue -eq 1) -or `
                    (!$originalRecord[$crmFieldKey+"_Property"].Value -and $crmFieldValue -eq 0))
                    {
                        continue 
                    }  
                }
                elseif($crmFieldValue -is [bool])
                {
                    if($crmFieldValue -eq $originalRecord[$crmFieldKey+"_Property"].Value)
                    {
                        continue
                    }
                }
                elseif($crmFieldValue -eq $originalRecord[$crmFieldKey])
                {
                    continue
                }                             
            }
            elseif($originalRecord[$crmFieldKey+"_Property"].Value -is [Microsoft.Xrm.Sdk.OptionSetValue])
            { 
                if($crmFieldValue -is [Microsoft.Xrm.Sdk.OptionSetValue])
                {
                    if($crmFieldValue.Value -eq $originalRecord[$crmFieldKey+"_Property"].Value.Value)
                    {
                        continue
                    }
                } 
                elseif($crmFieldValue -is [Int32])
                {
                    if($crmFieldValue -eq $originalRecord[$crmFieldKey+"_Property"].Value.Value)
                    {
                        continue
                    }
                }
                elseif($crmFieldValue -eq $originalRecord[$crmFieldKey])
                {
                    continue
                }
            }            
            elseif($originalRecord[$crmFieldKey+"_Property"].Value -is [Microsoft.Xrm.Sdk.Money])
            { 
                if($crmFieldValue -is [Microsoft.Xrm.Sdk.Money])
                {
                    if($crmFieldValue.Value -eq $originalRecord[$crmFieldKey+"_Property"].Value.Value)
                    {
                        continue
                    }
                }
                elseif($crmFieldValue -is [decimal] -or $crmFieldValue -is [Int32])
                {
                    if($crmFieldValue -eq $originalRecord[$crmFieldKey+"_Property"].Value.Value)
                    {
                        continue
                    }
                }
                elseif($crmFieldValue -eq $originalRecord[$crmFieldKey])
                {
                    continue
                }
            }
            elseif($originalRecord[$crmFieldKey+"_Property"].Value -is [Microsoft.Xrm.Sdk.EntityReference])
            { 
                if(($crmFieldValue -is [Microsoft.Xrm.Sdk.EntityReference]) -and ($crmFieldValue.Name -eq $originalRecord[$crmFieldKey].Name))
                {
                    continue
                }
                elseif($crmFieldValue -eq $originalRecord[$crmFieldKey].Name)
                {
                    continue
                }
            }
            elseif($crmFieldValue -eq $originalRecord[$crmFieldKey])
            { 
                continue 
            }

            $newfield = New-Object -TypeName 'Microsoft.Xrm.Tooling.Connector.CrmDataTypeWrapper'
            $value = New-Object psobject
            switch($CrmRecord.($crmFieldKey + "_Property").Value.GetType().Name)
            {
                "Boolean" {
                    $newfield.Type = [Microsoft.Xrm.Tooling.Connector.CrmFieldType]::CrmBoolean
                    if($crmFieldValue -is [Boolean])
                    {
                        $value = $crmFieldValue
                    }
                    else
                    {
                        $value = [Int32]::Parse($crmFieldValue)
                    }
                    break
                }
                "DateTime" {
                    $newfield.Type = [Microsoft.Xrm.Tooling.Connector.CrmFieldType]::CrmDateTime
                    if($crmFieldValue -is [DateTime])
                    {
                        $value = $crmFieldValue
                    }
                    else
                    {
                        $value = [DateTime]::Parse($crmFieldValue)
                    }
                    break
                }
                "Decimal" {
                    $newfield.Type = [Microsoft.Xrm.Tooling.Connector.CrmFieldType]::CrmDecimal
                    if($crmFieldValue -is [Decimal])
                    {
                        $value = $crmFieldValue
                    }
                    else
                    {
                        $value = [Decimal]::Parse($crmFieldValue)
                    }
                    break
                }
                "Single" {
                    $newfield.Type = [Microsoft.Xrm.Tooling.Connector.CrmFieldType]::CrmFloat
                    if($crmFieldValue -is [Single])
                    {
                        $value = $crmFieldValue
                    }
                    else
                    {
                        $value = [Single]::Parse($crmFieldValue)
                    }
                    break
                }
                "Money" {
                    $newfield.Type = [Microsoft.Xrm.Tooling.Connector.CrmFieldType]::Raw
                    if($crmFieldValue -is [Microsoft.Xrm.Sdk.Money])
                    {                
                        $value = $crmFieldValue
                    }
                    else
                    {                
                        $value = New-Object -TypeName 'Microsoft.Xrm.Sdk.Money'
                        $value.Value = $crmFieldValue
                    }
                    break
                }
                "Int32" {
                    $newfield.Type = [Microsoft.Xrm.Tooling.Connector.CrmFieldType]::CrmNumber
                    if($crmFieldValue -is [Int32])
                    {
                        $value = $crmFieldValue
                    }
                    else
                    {
                        $value = [Int32]::Parse($crmFieldValue)
                    }
                    break
                }
                "EntityReference" {
                    $newfield.Type = [Microsoft.Xrm.Tooling.Connector.CrmFieldType]::Raw
                    $value = $crmFieldValue
                    break
                }
                "OptionSetValue" {
                    $newfield.Type = [Microsoft.Xrm.Tooling.Connector.CrmFieldType]::Raw
                    if($crmFieldValue -is [Microsoft.Xrm.Sdk.OptionSetValue])
                    {
                        $value = $crmFieldValue                        
                    }
                    else
                    {
                        $value = New-Object -TypeName 'Microsoft.Xrm.Sdk.OptionSetValue'
                        $value.Value = [Int32]::Parse($crmFieldValue)
                    }
                    break
                }
                "String" {
                    $newfield.Type = [Microsoft.Xrm.Tooling.Connector.CrmFieldType]::String
                    $value = $crmFieldValue
                    break
                }
            }
        
            $newfield.Value = $value
            $newfields.Add($crmFieldKey, $newfield)
        }
    }
    else
    {
        foreach($field in $Fields.GetEnumerator())
        {  
           $newfield = New-Object -TypeName 'Microsoft.Xrm.Tooling.Connector.CrmDataTypeWrapper'
        
            switch($field.Value.GetType().Name)
            {
                "Boolean" {
                    $newfield.Type = [Microsoft.Xrm.Tooling.Connector.CrmFieldType]::CrmBoolean
                    break             
                }
                "DateTime" {
                    $newfield.Type = [Microsoft.Xrm.Tooling.Connector.CrmFieldType]::CrmDateTime
                    break
                }
                "Decimal" {
                    $newfield.Type = [Microsoft.Xrm.Tooling.Connector.CrmFieldType]::CrmDecimal
                    break
                }
                "Single" {
                    $newfield.Type = [Microsoft.Xrm.Tooling.Connector.CrmFieldType]::CrmFloat
                    break
                }
                "Money" {
                    $newfield.Type = [Microsoft.Xrm.Tooling.Connector.CrmFieldType]::Raw
                    break
                }
                "Int32" {
                    $newfield.Type = [Microsoft.Xrm.Tooling.Connector.CrmFieldType]::CrmNumber
                    break
                }
                "EntityReference" {
                    $newfield.Type = [Microsoft.Xrm.Tooling.Connector.CrmFieldType]::Raw
                    break
                }
                "OptionSetValue" {
                    $newfield.Type = [Microsoft.Xrm.Tooling.Connector.CrmFieldType]::Raw
                    break
                }
                "String" {
                    $newfield.Type = [Microsoft.Xrm.Tooling.Connector.CrmFieldType]::String
                    break
               }
            }
        
            $newfield.Value = $field.Value
            $newfields.Add($field.Key, $newfield)
        }
    }
    
    try
    {
        $result = $conn.UpdateEntity($entityLogicalName, $primaryKeyField, $Id, $newfields, $null, $false, [Guid]::Empty)
        if(!$result)
        {
            return $conn.LastCrmError
        }
    }
    catch
    {
        #TODO: Throw Exceptions back to user
        return $conn.LastCrmException
    }
}

#DeleteEntity 
function Remove-CrmRecord{

<#
 .SYNOPSIS
 Delete a CRM record by specifying EntityLogicalName and record's Id (guid)

 .DESCRIPTION
 The Remove-CrmRecord cmdlet lets you delete a record of your CRM organization. 
 
 There are two ways to delete a record. 
 1. Pass EntityLogicalName and record's Id.
 2. Get a record object by using Get-CrmRecord/Get-CrmRecords cmdlets, then pass it.

 .PARAMETER conn
 A connection to your CRM organizatoin. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .PARAMETER CrmRecord
 A record object which is obtained via Get-CrmRecord/Get-CrmRecords. When you pass CrmRecord, then you don't use EntityLogicalName/Id.

 .PARAMETER EntityLogicalName
 A logicalname for an Entity to delete. i.e.)accout, contact, lead, etc..

 .PARAMETER Id
 An Id (guid) of the record

 .EXAMPLE
 Remove-CrmRecord -conn $conn -EntityLogicalName account -Id 52a17637-5617-e511-80dc-c4346bc4fc6c

 This example deletes an account record by using Id

 .EXAMPLE 
 PS C:\>$account = Get-CrmRecord account b202caab-6c16-e511-80d6-c4346bc43dc0 name

 PS C:\>$account
 accountid_Property : [accountid, b202caab-6c16-e511-80d6-c4346bc43dc0]
 accountid          : b202caab-6c16-e511-80d6-c4346bc43dc0
 original           : {[name_Property, [name, Adventure Works (sample)]], [name, Adventure Works (sample)], [accountid_Property, [accountid, 
                      b202caab-6c16-e511-80d6-c4346bc43dc0]], [accountid, b202caab-6c16-e511-80d6-c4346bc43dc0]}
 logicalname        : account

 PS C:\>Remove-CrmRecord $account

 PS C:\>Get-CrmRecord account b202caab-6c16-e511-80d6-c4346bc43dc0 name
 WARNING: Record Id: b202caab-6c16-e511-80d6-c4346bc43dc0Does Not Exist

 This example retrieves and store an account record to $account object, then pass it to Remove-CrmRecord cmdlet.
 Finally retrieves it again to confirm it is deleted.
 When ommiting parameter names, you do not provide $conn, cmdlets automatically finds it.

 .EXAMPLE 
 PS C:\>$fetch = @"
 <fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="false" no-lock="true">
    <entity name="account">
      <attribute name="name" />
      </entity>
  </fetch>
 "@

 PS C:\>(Get-CrmRecordsByFetch $fetch).CrmRecords | % { Remove-CrmRecord $conn -CrmRecord $_}
 
 PS C:\>Get-CrmRecordsByFetch $fetch
 WARNING: No Result

 This example retrieves and stores account records by using FetchXML and pipe results (CrmRecords). In the next pipe, it updates each record using Remove-CrmRecord.
 Finally retrieves them again to confirm all records are deleted.
#>

    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1, ParameterSetName="CrmRecord", ValueFromPipeline=$True)]
        [PSObject]$CrmRecord,
        [parameter(Mandatory=$true, Position=1, ParameterSetName="Fields")]
        [string]$EntityLogicalName,
        [parameter(Mandatory=$true, Position=2, ParameterSetName="Fields")]
        [guid]$Id
    )

    begin
    {
        if($conn -eq $null)
        {
            $connobj = Get-Variable conn -Scope global -ErrorAction SilentlyContinue
            if($connobj.Value -eq $null)
            {
                Write-Warning 'You need to create Connect to CRM Organization. Use Get-CrmConnection to create it.'
                break;
            }
            else
            {
                $conn = $connobj.Value
            }
        }
    }
    process
    {
        if($CrmRecord -ne $null)
        {
            $EntityLogicalName = $CrmRecord.logicalname
            $Id = $CrmRecord.($EntityLogicalName + "id")
        }

        try
        {
            $result = $conn.DeleteEntity($EntityLogicalName, $Id, [Guid]::Empty)
            if(!$result)
            {
                return $conn.LastCrmError
            }
        }
        catch
        {
            return $conn.LastCrmException
        }
    }
}

### Other Cmdlets from Xrm Tooling ###
#AddEntityToQueue 
function Move-CrmRecordToQueue{

<#
 .SYNOPSIS
 Move a CRM record to a Queue

 .DESCRIPTION
 The Move-CrmRecordToQueue cmdlet lets you move a record to a queue. 

 There are two ways to specify a record. 
 1. Pass EntityLogicalName and record's Id.
 2. Get a record object by using Get-CrmRecord/Get-CrmRecords cmdlets, then pass it.

 .PARAMETER conn
 A connection to your CRM organizatoin. Use $conn = Get-CrmConnection <Parameters> to generate it.

  .PARAMETER CrmRecord
 A record object which is obtained via Get-CrmRecord/Get-CrmRecords. When you pass CrmRecord, then you don't use EntityLogicalName/Id.

 .PARAMETER EntityLogicalName
 A logicalname for an Entity to move. i.e.)accout, contact, lead, etc..

 .PARAMETER Id
 An Id (guid) of the record

 .PARAMETER QueueName
 Queue's Name to move the record.

 .PARAMETER WorkingUserId
 An Id (guid) of SystemUser who works on the queue item.

 .PARAMETER SetWorkingByUser
 Specify if the record needs to be marked as WorkingByUser.

 .EXAMPLE
 Move-CrmRecordToQueue -conn $conn -EntityLogicalName incident -Id 1e005a70-6317-e511-80da-c4346bc43d94 -QueueName "Support Queue" -WorkingUserId f9d40920-7a43-4f51-9749-0549c4caf67d
 
 This example moves an incident (Case) record to "Support Queue" queue and assigned it to a User as WorkingUser.

 .EXAMPLE
 Move-CrmRecordToQueue incident 1e005a70-6317-e511-80da-c4346bc43d94 "Support Queue" f9d40920-7a43-4f51-9749-0549c4caf67d $True
 
 This example moves an incident (Case) record to "Support Queue" queue and assigned it to a User as WorkingUser and mark it as User Working 
 When ommiting parameter names, you do not provide $conn, cmdlets automatically finds it.

 .EXAMPLE 
 PS C:\>$incident = Get-CrmRecord incident 20005a70-6317-e511-80da-c4346bc43d94 title
 
 PS C:\>Move-CrmRecordToQueue $incident "Support Queue" f9d40920-7a43-4f51-9749-0549c4caf67d $True

 This example retrieves and store an incident record, then pass it to Move-CrmRecordToQueue.
#>

    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1, ParameterSetName="CrmRecord")]
        [PSObject]$CrmRecord,
        [parameter(Mandatory=$true, Position=1, ParameterSetName="NameWithId")]
        [string]$EntityLogicalName,
        [parameter(Mandatory=$true, Position=2, ParameterSetName="NameWithId")]
        [guid]$Id,
        [parameter(Mandatory=$true, Position=3)]
        [string]$QueueName,
        [parameter(Mandatory=$true, Position=4)]
        [guid]$WorkingUserId,
        [parameter(Mandatory=$false, Position=5)]
        [bool]$SetWorkingByUser
    )

    if($conn -eq $null)
    {
        $connobj = Get-Variable conn -Scope global -ErrorAction SilentlyContinue
        if($connobj.Value -eq $null)
        {
            Write-Warning 'You need to create Connect to CRM Organization. Use Get-CrmConnection to create it.'
            break;
        }
        else
        {
            $conn = $connobj.Value
        }
    }  

    if($CrmRecord -ne $null)
    {
        $EntityLogicalName = $CrmRecord.logicalname
        $Id = $CrmRecord.($EntityLogicalName + "id")
    }

    try
    {
        $result = $conn.AddEntityToQueue($Id, $EntityLogicalName, $QueueName, $WorkingUserId, $SetWorkingByUser, [Guid]::Empty)
    }
    catch
    {
        return $conn.LastCrmException
    }
}

#AssignEntityToUser
function Set-CrmRecordOwner{

<#
 .SYNOPSIS
 Assign an user as a CRM record's owner

 .DESCRIPTION
 The Set-CrmRecordOwner cmdlet lets you assign a user to a record's owner. 
 
 There are two ways to specify a record. 
 1. Pass EntityLogicalName and record's Id.
 2. Get a record object by using Get-CrmRecord/Get-CrmRecords cmdlets, then pass it.

 .PARAMETER conn
 A connection to your CRM organizatoin. Use $conn = Get-CrmConnection <Parameters> to generate it.

  .PARAMETER CrmRecord
 A record object which is obtained via Get-CrmRecord/Get-CrmRecords. When you pass CrmRecord, then you don't use EntityLogicalName/Id.

 .PARAMETER EntityLogicalName
 A logicalname for an Entity to set an owner. i.e.)accout, contact, lead, etc..

 .PARAMETER Id
 An Id (guid) of the record

 .PARAMETER PrincipalId
 An Id (guid) of SystemUser or team to be assigned as owner.

 .PARAMETER AssignToTeam
 Switch indicating the PrincipalId Supplied is for a CRM Ownership Team and NOT a CRM User

 .EXAMPLE
 Set-CrmRecordOwner -conn $conn -EntityLogicalName contact -Id e1d47674-4017-e511-80db-c4346bc42d18 -PrincipalId f9d40920-7a43-4f51-9749-0549c4caf67d
 
 This example assigns a contact record to an User.

 .EXAMPLE
 Set-CrmRecordOwner -conn $conn -EntityLogicalName contact -Id e1d47674-4017-e511-80db-c4346bc42d18 -PrincipalId f0d40920-7a43-4f51-9749-0549c4caf673 -AssignToTeam
 
 This example assigns a contact record to a Team.

 .EXAMPLE
 Set-CrmRecordOwner contact e1d47674-4017-e511-80db-c4346bc42d18 f9d40920-7a43-4f51-9749-0549c4caf67d
 
 This example assigns a contact record to an User  by ommiting parameter names.
 When ommiting parameter names, you do not provide $conn, cmdlets automatically finds it.

 .EXAMPLE 
 PS C:\>$contact = Get-CrmRecord contact e1d47674-4017-e511-80db-c4346bc42d18 fullname
 
 PS C:\>Set-CrmRecordOwner $contact f9d40920-7a43-4f51-9749-0549c4caf67d

 This example retrieves and store a contact record, then pass it to Set-CrmRecordOwner.
#>

    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,        
        [parameter(Mandatory=$true, Position=1, ParameterSetName="CrmRecord")]
        [PSObject]$CrmRecord,
        [parameter(Mandatory=$true, Position=1, ParameterSetName="NameWithId")]
        [string]$EntityLogicalName,
        [parameter(Mandatory=$true, Position=2, ParameterSetName="NameWithId")]
        [guid]$Id,
        [parameter(Mandatory=$true, Position=3)][alias("UserId")]
        [guid]$PrincipalId,
		[parameter(Mandatory=$true, Position=4)]
		[switch]$AssignToTeam
    )

    if($conn -eq $null)
    {
        $connobj = Get-Variable conn -Scope global -ErrorAction SilentlyContinue
        if($connobj.Value -eq $null)
        {
            Write-Warning 'You need to create Connect to CRM Organization. Use Get-CrmConnection to create it.'
            break;
        }
        else
        {
            $conn = $connobj.Value
        }
    }  

    if($CrmRecord -ne $null)
    {
        $EntityLogicalName = $CrmRecord.logicalname
        $Id = $CrmRecord.($EntityLogicalName + "id")
    }

    try
    {
		if($conn.LastCrmException -ne $null){
			$errTrack = $conn.LastCrmException.getHashCode(); 
			#Write-Verbose "Before Hashcode: $errTrack"
		}
		else{
			$errTrack = $null;
		}
		if($AssignToTeam){
			write-verbose "Assigning record with Id: $Id to Team with Id: $PrincipalId"
			
			$req = New-Object Microsoft.Crm.Sdk.Messages.AssignRequest
			$req.target = New-CrmEntityReference -EntityLogicalName $EntityLogicalName -Id $Id; 
			$req.Assignee = New-CrmEntityReference -EntityLogicalName "team" -Id $PrincipalId; 
			$result = [Microsoft.Crm.Sdk.Messages.AssignResponse]$conn.ExecuteCrmOrganizationRequest($req, $null);
		}
		else{
	        $result = $conn.AssignEntityToUser($PrincipalId, $EntityLogicalName, $Id, [Guid]::Empty)
		}
		#Checks to see if the hashcode of the last exception is present, and compares it to the current exception hashcode (if present)
		if($conn.LastCrmException -ne $null -And $errTrack -ne $null){
			#Write-Verbose "After Hashcode: $errTrack"
			if($errTrack -ne $conn.LastCrmException.getHashCode() ){
				write-error $conn.LastCrmException
				throw $conn.LastCrmException; 
			}
		}
		write-verbose "Completed..."
    }
    catch
    {
        return $conn.LastCrmException
    }
}

#CloseActivity 
function Set-CrmActivityRecordToCloseState{

<#
 .SYNOPSIS
 Close an activity record.

 .DESCRIPTION
 The Set-CrmActivityRecordToCloseState cmdlet lets you close an activity record. 
 
 There are two ways to specify a record. 
 1. Pass ActivityEntityType and record's Id.
 2. Get a record object by using Get-CrmRecord/Get-CrmRecords cmdlets, then pass it.

 .PARAMETER conn
 A connection to your CRM organizatoin. Use $conn = Get-CrmConnection <Parameters> to generate it.

  .PARAMETER CrmRecord
 A record object which is obtained via Get-CrmRecord/Get-CrmRecords. When you pass CrmRecord, then you don't use ActivityEntityType/ActivityId

 .PARAMETER ActivityEntityType
 A logicalname for an Activity Entity to close. i.e.)phonecall, email, task, etc..

 .PARAMETER ActivityId
 An Id (guid) of the activity record

 .PARAMETER StateCode
 A State code name. You can use (Get-CrmEntityOptionSet <EntityLogicalName> statecode).Items to get StateCode strings.

 .PARAMETER StatusCode
 A Status code name. You can use (Get-CrmEntityOptionSet <EntityLogicalName> statuscode).Items to get StateCode strings.

 .EXAMPLE
 Set-CrmActivityRecordToCloseState -conn $conn -ActivityEntityType task -ActivityId a0025a70-6317-e511-80da-c4346bc43d94 -StateCode Completed -StatusCode Completed
 
 This example closes a task record as Completed/Completed.

 .EXAMPLE
 Set-CrmActivityRecordToCloseState task a0025a70-6317-e511-80da-c4346bc43d94 Open "In Progress"
 
 This example closes a task record as InProgress/Open.
 When ommiting parameter names, you do not provide $conn, cmdlets automatically finds it.

 .EXAMPLE 
 PS C:\>$task = Get-CrmRecord task a0025a70-6317-e511-80da-c4346bc43d94 subject
 
 PS C:\>Set-CrmActivityRecordToCloseState $task Open "Not Started"

 This example retrieves and store a task record, then pass it to Set-CrmActivityRecordToCloseState
#>

    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1, ParameterSetName="CrmRecord")]
        [PSObject]$CrmRecord,
        [parameter(Mandatory=$true, Position=1, ParameterSetName="NameWithId")]
        [string]$ActivityEntityType,
        [parameter(Mandatory=$true, Position=2, ParameterSetName="NameWithId")]
        [guid]$ActivityId,
        [parameter(Mandatory=$true, Position=3)]
        [string]$StateCode,
        [parameter(Mandatory=$true, Position=4)]
        [string]$StatusCode
    )

    if($conn -eq $null)
    {
        $connobj = Get-Variable conn -Scope global -ErrorAction SilentlyContinue
        if($connobj.Value -eq $null)
        {
            Write-Warning 'You need to create Connect to CRM Organization. Use Get-CrmConnection to create it.'
            break;
        }
        else
        {
            $conn = $connobj.Value
        }
    }   

    if($CrmRecord -ne $null)
    {
        $ActivityEntityType = $CrmRecord.logicalname
        $ActivityId = $CrmRecord.("activityid")
    }

    try
    {
        $result = $conn.CloseActivity($ActivityEntityType, $ActivityId, $StateCode, $StatusCode, [Guid]::Empty)
    }
    catch
    {
        return $conn.LastCrmException
    }
}

#CreateAnnotation 
function Add-CrmNoteToCrmRecord{

<#
 .SYNOPSIS
 Create a new note (annotation) to a record.

 .DESCRIPTION
 The Add-CrmNoteToCrmRecord cmdlet lets you add a note (annoation) to a record. 

 There are two ways to specify a target record.
 
 1. Pass EntityLogicalName and record's Id.
 2. Get a record object by using Get-CrmRecord/Get-CrmRecords cmdlets, and pass it.

 You can specify note subject and body by using -Subject and -NoteText parameters

 .PARAMETER conn
 A connection to your CRM organizatoin. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .PARAMETER CrmRecord
 A record object which is obtained via Get-CrmRecord/Get-CrmRecords. When you pass CrmRecord, then you don't use EntityLogicalName/Id.

 .PARAMETER EntityLogicalName
 A logicalname of the target record. i.e.)accout, contact, lead, etc..

 .PARAMETER Id
 An Id (guid) of the record

 .PARAMETER Subject
 Subject of a note (annotation).

 .PARAMETER NoteText
 Body Text of a note (annotation).

 .EXAMPLE
 Add-CrmNoteToCrmRecord -conn $conn -EntityLogicalName account -Id 00005a70-6317-e511-80da-c4346bc43d94 -Subject "sample subject" -NoteText "sample body"

 This example add a note (annotation) to an account record.

 .EXAMPLE
 Add-CrmNoteToCrmRecord account 00005a70-6317-e511-80da-c4346bc43d94 "sample subject" "sample body"
 
 This example add a note (annotation) to an account record by ommiting parameters names.
 When ommiting parameter names, you do not provide $conn, cmdlets automatically finds it.

#>

    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1, ParameterSetName="CrmRecord")]
        [PSObject]$CrmRecord,
        [parameter(Mandatory=$true, Position=1, ParameterSetName="NameWithId")]
        [string]$EntityLogicalName,
        [parameter(Mandatory=$true, Position=2, ParameterSetName="NameWithId")]
        [guid]$Id,
        [parameter(Mandatory=$true, Position=3)]
        [string]$Subject,
        [parameter(Mandatory=$true, Position=4)]
        [string]$NoteText 
    )

    if($conn -eq $null)
    {
        $connobj = Get-Variable conn -Scope global -ErrorAction SilentlyContinue
        if($connobj.Value -eq $null)
        {
            Write-Warning 'You need to create Connect to CRM Organization. Use Get-CrmConnection to create it.'
            break;
        }
        else
        {
            $conn = $connobj.Value
        }
    }   

    if($CrmRecord -ne $null)
    {
        $EntityLogicalName = $CrmRecord.logicalname
        $Id = $CrmRecord.($EntityLogicalName + "id")
    }

    $newfields = New-Object 'System.Collections.Generic.Dictionary[[String], [Microsoft.Xrm.Tooling.Connector.CrmDataTypeWrapper]]'
    $subjectfield = New-Object -TypeName 'Microsoft.Xrm.Tooling.Connector.CrmDataTypeWrapper'
    $subjectfield.Type = [Microsoft.Xrm.Tooling.Connector.CrmFieldType]::String
    $subjectfield.Value = $Subject
    $noteTextfield = New-Object -TypeName 'Microsoft.Xrm.Tooling.Connector.CrmDataTypeWrapper'
    $noteTextfield.Type = [Microsoft.Xrm.Tooling.Connector.CrmFieldType]::String
    $noteTextfield.Value = $NoteText
    $newfields.Add("subject", $subjectfield)
    $newfields.Add("notetext", $noteTextfield)

    try
    {
        $result = $conn.CreateAnnotation($EntityLogicalName, $Id, $newfields, [Guid]::Empty)
    }
    catch
    {
        return $conn.LastCrmException
    }
}

#CreateEntityAssociation
function Add-CrmRecordAssociation{

<#
 .SYNOPSIS
 Associates two records for N:N relationship.

 .DESCRIPTION
 The Add-CrmRecordAssociation cmdlet lets you associate two records for N:N relationship by specifying relatioship logical name. 

 There are two ways to specify records.
 
 1. Pass EntityLogicalName and record's Id for both records.
 2. Get a record object by using Get-CrmRecord/Get-CrmRecords cmdlets, and pass it for both records.

 You can specify relationship logical name for the association.

 .PARAMETER conn
 A connection to your CRM organizatoin. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .PARAMETER CrmRecord1
 A first record object which is obtained via Get-CrmRecord/Get-CrmRecords. When you pass CrmRecord, then you don't use EntityLogicalName/Id.

 .PARAMETER CrmRecord2
 A second record object which is obtained via Get-CrmRecord/Get-CrmRecords. When you pass CrmRecord, then you don't use EntityLogicalName/Id.

 .PARAMETER EntityLogicalName1
 A logicalname for first Entity. i.e.)accout, contact, lead, etc..

 .PARAMETER Id1
 An Id (guid) of first record

 .PARAMETER EntityLogicalName2
 A logicalname for second Entity. i.e.)accout, contact, lead, etc..

 .PARAMETER Id2
 An Id (guid) of second record

 .PARAMETER RelationshipName
 A N:N relationship logical name.

 .EXAMPLE
 Add-CrmRecordAssociation -conn $conn -EntityLogicalName1 account -Id1 00005a70-6317-e511-80da-c4346bc43d94 -EntityLogicalName2 contact -Id2 66005a70-6317-e511-80da-c4346bc43d94 -RelationshipName new_accounts_contacts

 This example associates an account and a contact records through new_accounts_contacts custom N:N relationship.

 .EXAMPLE
 Add-CrmRecordAssociation account 00005a70-6317-e511-80da-c4346bc43d94 contact 66005a70-6317-e511-80da-c4346bc43d94 new_accounts_contacts
 
 This example associates an account and a contact records through new_accounts_contacts custom N:N relationship by ommiting parameters names.
 When ommiting parameter names, you do not provide $conn, cmdlets automatically finds it.

 .EXAMPLE
 PS C:\>$account = Get-CrmRecord account 00005a70-6317-e511-80da-c4346bc43d94 name

 PS C:\>$contact = Get-CrmRecord contact 66005a70-6317-e511-80da-c4346bc43d94 fullname

 PS C:\>Add-CrmRecordAssociation -conn $conn -CrmRecord1 $account -CrmRecord2 $contact -RelationshipName new_accounts_contacts

 This example retrieves and stores an account and a contact records to variables, then pass them to Add-CrmRecordAssociation cmdlets.

 .EXAMPLE
 PS C:\>$account = Get-CrmRecord account 00005a70-6317-e511-80da-c4346bc43d94 name

 PS C:\>$contact = Get-CrmRecord contact 66005a70-6317-e511-80da-c4346bc43d94 fullname

 PS C:\>Add-CrmRecordAssociation $account $contact new_accounts_contacts

 This example retrieves and stores an account and a contact records to variables, then pass them to Add-CrmRecordAssociation cmdlets.

#>

    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1, ParameterSetName="CrmRecord")]
        [PSObject]$CrmRecord1,
        [parameter(Mandatory=$true, Position=2, ParameterSetName="CrmRecord")]
        [PSObject]$CrmRecord2,
        [parameter(Mandatory=$true, Position=1, ParameterSetName="NameWithId")]
        [string]$EntityLogicalName1,
        [parameter(Mandatory=$true, Position=2, ParameterSetName="NameWithId")]
        [guid]$Id1,
        [parameter(Mandatory=$true, Position=3, ParameterSetName="NameWithId")]
        [string]$EntityLogicalName2,
        [parameter(Mandatory=$true, Position=4, ParameterSetName="NameWithId")]
        [guid]$Id2,
        [parameter(Mandatory=$true, Position=5)]
        [string]$RelationshipName
    )

    if($conn -eq $null)
    {
        $connobj = Get-Variable conn -Scope global -ErrorAction SilentlyContinue
        if($connobj.Value -eq $null)
        {
            Write-Warning 'You need to create Connect to CRM Organization. Use Get-CrmConnection to create it.'
            break;
        }
        else
        {
            $conn = $connobj.Value
        }
    }

    if($CrmRecord1 -ne $null)
    {
        $EntityLogicalName1 = $CrmRecord1.logicalname
        $Id1 = $CrmRecord1.($EntityLogicalName1 + "id")
    }

    if($CrmRecord2 -ne $null)
    {
        $EntityLogicalName2 = $CrmRecord2.logicalname
        $Id2 = $CrmRecord2.($EntityLogicalName2 + "id")
    }

    try
    {
        $result = $conn.CreateEntityAssociation($EntityLogicalName1, $Id1, $EntityLogicalName2, $Id2, $RelationshipName, [Guid]::Empty)
    }
    catch
    {
        return $conn.LastCrmException
    }
}

#CreateMultiEntityAssociation
function Add-CrmMultiRecordAssociation{

<#
 .SYNOPSIS
 Associates multiple records to single record for N:N relationship.

 .DESCRIPTION
 The Add-CrmMultiRecordAssociation cmdlet lets you associate multiple records to single record for N:N relationship by specifying relatioship logical name. 
 Use @('<object>','<object>') syntax to specify multiple ids or records.
 if the relationship is self-referencing, specify $True for -IsReflexiveRelationship Parameter.

 There are two ways to specify records.
 
 1. Pass EntityLogicalName and record's Id for both records.
 2. Get record object(s) by using Get-CrmRecord/Get-CrmRecords cmdlets, and pass them.

 You can specify relationship logical name for the association.

 .PARAMETER conn
 A connection to your CRM organizatoin. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .PARAMETER CrmRecord1
 A first record object which is obtained via Get-CrmRecord/Get-CrmRecords. When you pass CrmRecord, then you don't use EntityLogicalName/Id.

 .PARAMETER CrmRecord2s
 An array of records object which are obtained via Get-CrmRecord/Get-CrmRecords. When you pass CrmRecord, then you don't use EntityLogicalName/Id.

 .PARAMETER EntityLogicalName1
 A logicalname for first Entity. i.e.)accout, contact, lead, etc..

 .PARAMETER Id1
 An Id (guid) of first record

 .PARAMETER EntityLogicalName2
 A logicalname for second Entity. i.e.)accout, contact, lead, etc..

 .PARAMETER Id2s
 An array of Ids (guid) of second records. Specify by using @('66005a70-6317-e511-80da-c4346bc43d94','62005a70-6317-e511-80da-c4346bc43d94') synctax.

 .PARAMETER RelationshipName
 A N:N relationship logical name.

 .PARAMETER IsReflexiveRelationship
 Specify $True if the N:N relationship is self-referencing.

 .EXAMPLE
 Add-CrmMultiRecordAssociation -conn $conn -EntityLogicalName1 account -Id1 00005a70-6317-e511-80da-c4346bc43d94 -EntityLogicalName2 contact -Id2s @('66005a70-6317-e511-80da-c4346bc43d94','62005a70-6317-e511-80da-c4346bc43d94') -RelationshipName new_accounts_contacts
 This example associates an account and two contact records through new_accounts_contacts custom N:N relationship.

 .EXAMPLE
 Add-CrmMultiRecordAssociation account 00005a70-6317-e511-80da-c4346bc43d94 contact @('66005a70-6317-e511-80da-c4346bc43d94','62005a70-6317-e511-80da-c4346bc43d94') new_accounts_contacts
 
 This example associates an account and two contact records through new_accounts_contacts custom N:N relationship by ommiting parameters names.
 When ommiting parameter names, you do not provide $conn, cmdlets automatically finds it.

 .EXAMPLE
 PS C:\>$account = Get-CrmRecord account 00005a70-6317-e511-80da-c4346bc43d94 name

 PS C:\>$fetch = @"
<fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="false" no-lock="true">
  <entity name="contact">
    <attribute name="fullname" />
    <filter type="and">
      <condition attribute="lastname" operator="like" value="%sample%" />
    </filter>
  </entity>
</fetch>
"@

 PS C:\>$contacts = Get-CrmRecordsByFetch $fetch

 PS C:\>Add-CrmMultiRecordAssociation $account $contacts.CrmRecords new_accounts_contacts

 This example retrieves contacts by using FetchXML and stores to a variable, then retrieves and store an account record to another variable.
 Then passes those variables Add-CrmMultiRecordAssociation.
#>

    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1, ParameterSetName="CrmRecord")]
        [PSObject]$CrmRecord1,
        [parameter(Mandatory=$true, Position=2, ParameterSetName="CrmRecord")]
        [PSObject[]]$CrmRecord2s,
        [parameter(Mandatory=$true, Position=1, ParameterSetName="NameWithId")]
        [string]$EntityLogicalName1,
        [parameter(Mandatory=$true, Position=2, ParameterSetName="NameWithId")]
        [guid]$Id1,
        [parameter(Mandatory=$true, Position=3, ParameterSetName="NameWithId")]
        [string]$EntityLogicalName2,
        [parameter(Mandatory=$true, Position=4, ParameterSetName="NameWithId")]
        [guid[]]$Id2s,
        [parameter(Mandatory=$true, Position=5)]
        [string]$RelationshipName,
        [parameter(Mandatory=$false, Position=6)]
        [bool]$IsReflexiveRelationship
    )

    if($conn -eq $null)
    {
        $connobj = Get-Variable conn -Scope global -ErrorAction SilentlyContinue
        if($connobj.Value -eq $null)
        {
            Write-Warning 'You need to create Connect to CRM Organization. Use Get-CrmConnection to create it.'
            break;
        }
        else
        {
            $conn = $connobj.Value
        }
    }  

    if($CrmRecord1 -ne $null)
    {
        $EntityLogicalName1 = $CrmRecord1.logicalname
        $Id1 = $CrmRecord1.($EntityLogicalName1 + "id")
    }

    if($CrmRecord2s -ne $null)
    {
        if($CrmRecord2s.Count -ne 0)
        {
            $EntityLogicalName2 = $CrmRecord2s[0].logicalname
            $Ids = New-Object 'System.Collections.Generic.List[System.Guid]'
            foreach($CrmRecord2 in $CrmRecord2s)
            {
                $Ids.Add($CrmRecord2.($EntityLogicalName2 + "id"))
            }
            $Id2s = $Ids.ToArray()
        }
         else
        {
            Write-Warning 'CrmRecords2 does not include any records.'
            break;
        }
    }   

    try
    {
        $result = $conn.CreateMultiEntityAssociation($EntityLogicalName1, $Id1, $EntityLogicalName2, $Id2s, $RelationshipName, [Guid]::Empty, $IsReflexiveRelationship)
    }
    catch
    {
        return $conn.LastCrmException
    }
}

#CreateNewActivityEntry 
function Add-CrmActivityToCrmRecord{

<#
 .SYNOPSIS
 Create a new activity to a record.

 .DESCRIPTION
 The Add-CrmActivityToCrmRecord cmdlet lets you add an activity to a record. You use ActivityEntityType to specify Activity Type and Subject/Description to set values. 
 You can use Fields optional Parameter to specify additional Field values. Use @{"field logical name"="value"} syntax to create Fields , and make sure you specify correct type of value for the field. 
 You can use Get-CrmEntityAttributeMetadata cmdlet and check AttributeType to see the field type. In addition, for CRM specific types, you can use New-CrmMoney, New-CrmOptionSetValue or New-CrmEntityReference cmdlets.

 There are two ways to specify a target record.
 
 1. Pass EntityLogicalName and record's Id.
 2. Get a record object by using Get-CrmRecord/Get-CrmRecords cmdlets, and pass it.

 You can specify note subject and body by using -Subject and -NoteText parameters

 .PARAMETER conn
 A connection to your CRM organizatoin. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .PARAMETER CrmRecord
 A record object which is obtained via Get-CrmRecord/Get-CrmRecords. When you pass CrmRecord, then you don't use EntityLogicalName/Id.

 .PARAMETER EntityLogicalName
 A logicalname of the target record. i.e.)accout, contact, lead, etc..

 .PARAMETER Id
 An Id (guid) of the record

 .PARAMETER ActivityEntityType
 A logicalname for an Activity Entity to add. i.e.)phonecall, email, task, etc..

 .PARAMETER Subject
 Subject of the activity record.

 .PARAMETER Description
 Description of the activity record.

 .PARAMETER Fields
 A List of field name/value pair. Use @{"field logical name"="value"} syntax to create Fields, and make sure you specify correct type of value for the field. 
 You can use Get-CrmEntityAttributeMetadata cmdlet and check AttributeType to see the field type. In addition, for CRM specific types, you can use New-CrmMoney, New-CrmOptionSetValue or New-CrmEntityReference cmdlets.

 .EXAMPLE
 Add-CrmActivityToCrmRecord -conn $conn -EntityLogicalName account -Id feff5970-6317-e511-80da-c4346bc43d94 -ActivityEntityType task -Subject "sample task" -Description "sample task description" -OnwerUserId f9d40920-7a43-4f51-9749-0549c4caf67d

 This example adds a task an account record.

 .EXAMPLE
 Add-CrmActivityToCrmRecord account feff5970-6317-e511-80da-c4346bc43d94 task "sample task" "sample task description" f9d40920-7a43-4f51-9749-0549c4caf67d
 
 This example adds a task to an account record by ommiting parameters names.
 When ommiting parameter names, you do not provide $conn, cmdlets automatically finds it.

 .EXAMPLE
 Add-CrmActivityToCrmRecord account feff5970-6317-e511-80da-c4346bc43d94 task "sample task" "sample task description" f9d40920-7a43-4f51-9749-0549c4caf67d @{"scheduledend"=(Get-Date).AddDays(3);"prioritycode"=New-CrmOptionSetValue 2}

 This example adds a task to an account record with Due Date and Priority fields.

 .EXAMPLE
 PS C:\>$account = Get-CrmRecord account feff5970-6317-e511-80da-c4346bc43d94 name

 PS C:\>user = Get-MyCrmUserId

 PS C:\>Add-CrmActivityToCrmRecord $account task "sample task" "sample task description" $user
 
 This example retrieves and stores an account, and login user Id (guid) to variables. Then passes them to Add-CrmActivityToCrmRecord.
#>

    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,        
        [parameter(Mandatory=$true, Position=1, ValueFromPipeline=$true, ParameterSetName="CrmRecord")]
        [PSObject]$CrmRecord,
        [parameter(Mandatory=$true, Position=1, ParameterSetName="NameWithId")]
        [string]$EntityLogicalName,
        [parameter(Mandatory=$true, Position=2, ParameterSetName="NameWithId")]
        [guid]$Id,
        [parameter(Mandatory=$true, Position=3)]
        [string]$ActivityEntityType,
        [parameter(Mandatory=$true, Position=4)]
        [string]$Subject,
        [parameter(Mandatory=$true, Position=5)]
        [string]$Description,
        [parameter(Mandatory=$true, Position=6)]
        [string]$OnwerUserId,
        [parameter(Mandatory=$false, Position=7)]
        [hashtable]$Fields
    )

    
    if($conn -eq $null)
    {
        $connobj = Get-Variable conn -Scope global -ErrorAction SilentlyContinue
        if($connobj.Value -eq $null)
        {
            Write-Warning 'You need to create Connect to CRM Organization. Use Get-CrmConnection to create it.'
            break;
        }
        else
        {
            $conn = $connobj.Value
        }
    }  

    if($CrmRecord -ne $null)
    {
        $EntityLogicalName = $CrmRecord.logicalname
        $Id = $CrmRecord.($EntityLogicalName + "id")
    }

    $newfields = New-Object 'System.Collections.Generic.Dictionary[[String], [Microsoft.Xrm.Tooling.Connector.CrmDataTypeWrapper]]'
    
    if($Fields -ne $null)
    {
        foreach($field in $Fields.GetEnumerator())
        {  
            $newfield = New-Object -TypeName 'Microsoft.Xrm.Tooling.Connector.CrmDataTypeWrapper'
        
            switch($field.Value.GetType().Name)
            {
            "Boolean" {
                $newfield.Type = [Microsoft.Xrm.Tooling.Connector.CrmFieldType]::CrmBoolean
                break              
            }
            "DateTime" {
                $newfield.Type = [Microsoft.Xrm.Tooling.Connector.CrmFieldType]::CrmDateTime
                break
            }
            "Decimal" {
                $newfield.Type = [Microsoft.Xrm.Tooling.Connector.CrmFieldType]::CrmDecimal
                break
            }
            "Single" {
                $newfield.Type = [Microsoft.Xrm.Tooling.Connector.CrmFieldType]::CrmFloat
                break
            }
            "Money" {
                $newfield.Type = [Microsoft.Xrm.Tooling.Connector.CrmFieldType]::Raw
                break
            }
            "Int32" {
                $newfield.Type = [Microsoft.Xrm.Tooling.Connector.CrmFieldType]::CrmNumber
                break
            }
            "EntityReference" {
                $newfield.Type = [Microsoft.Xrm.Tooling.Connector.CrmFieldType]::Raw
                break
            }
            "OptionSetValue" {
                $newfield.Type = [Microsoft.Xrm.Tooling.Connector.CrmFieldType]::Raw
                break
            }
            "String" {
                $newfield.Type = [Microsoft.Xrm.Tooling.Connector.CrmFieldType]::String
                break
            }
        }
        
            $newfield.Value = $field.Value
            $newfields.Add($field.Key, $newfield)
        }
    }

    try
    {
        $result = $conn.CreateNewActivityEntry($ActivityEntityType, $EntityLogicalName, $Id,
                $Subject, $Description, $OnwerUserId, $newfields, [Guid]::Empty)
    }
    catch
    {
        return $conn.LastCrmException
    }

    return $result
}

#DeleteEntityAssociation
function Remove-CrmRecordAssociation{

<#
 .SYNOPSIS
 Associates two records for N:N relationship.

 .DESCRIPTION
 The Remove-CrmRecordAssociation cmdlet lets you disassociate two records for N:N relationship by specifying relatioship logical name. 

 There are two ways to specify records.
 
 1. Pass EntityLogicalName and record's Id for both records.
 2. Get a record object by using Get-CrmRecord/Get-CrmRecords cmdlets, and pass it for both records.

 You can specify relationship logical name for the disassociation.

 .PARAMETER conn
 A connection to your CRM organizatoin. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .PARAMETER CrmRecord1
 A first record object which is obtained via Get-CrmRecord/Get-CrmRecords. When you pass CrmRecord, then you don't use EntityLogicalName/Id.

 .PARAMETER CrmRecord2
 A second record object which is obtained via Get-CrmRecord/Get-CrmRecords. When you pass CrmRecord, then you don't use EntityLogicalName/Id.

 .PARAMETER EntityLogicalName1
 A logicalname for first Entity. i.e.)accout, contact, lead, etc..

 .PARAMETER Id1
 An Id (guid) of first record

 .PARAMETER EntityLogicalName2
 A logicalname for second Entity. i.e.)accout, contact, lead, etc..

 .PARAMETER Id2
 An Id (guid) of second record

 .PARAMETER RelationshipName
 A N:N relationship logical name.

 .EXAMPLE
 Remove-CrmRecordAssociation -conn $conn -EntityLogicalName1 account -Id1 00005a70-6317-e511-80da-c4346bc43d94 -EntityLogicalName2 contact -Id2 66005a70-6317-e511-80da-c4346bc43d94 -RelationshipName new_accounts_contacts

 This example associates an account and a contact records through new_accounts_contacts custom N:N relationship.

 .EXAMPLE
 Remove-CrmRecordAssociation account 00005a70-6317-e511-80da-c4346bc43d94 contact 66005a70-6317-e511-80da-c4346bc43d94 new_accounts_contacts
 
 This example associates an account and a contact records through new_accounts_contacts custom N:N relationship by ommiting parameters names.
 When ommiting parameter names, you do not provide $conn, cmdlets automatically finds it.

 .EXAMPLE
 PS C:\>$account = Get-CrmRecord account 00005a70-6317-e511-80da-c4346bc43d94 name

 PS C:\>$contact = Get-CrmRecord contact 66005a70-6317-e511-80da-c4346bc43d94 fullname

 PS C:\>Remove-CrmRecordAssociation -conn $conn -CrmRecord1 $account -CrmRecord2 $contact -RelationshipName new_accounts_contacts

 This example retrieves and stores an account and a contact records to variables, then pass them to Remove-CrmRecordAssociation cmdlets.

 .EXAMPLE
 PS C:\>$account = Get-CrmRecord account 00005a70-6317-e511-80da-c4346bc43d94 name

 PS C:\>$contact = Get-CrmRecord contact 66005a70-6317-e511-80da-c4346bc43d94 fullname

 PS C:\>Remove-CrmRecordAssociation $account $contact new_accounts_contacts

 This example retrieves and stores an account and a contact records to variables, then pass them to Remove-CrmRecordAssociation cmdlets.
#>

    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1, ParameterSetName="CrmRecord")]
        [PSObject]$CrmRecord1,
        [parameter(Mandatory=$true, Position=2, ParameterSetName="CrmRecord")]
        [PSObject]$CrmRecord2,
        [parameter(Mandatory=$true, Position=1, ParameterSetName="NameWithId")]
        [string]$EntityLogicalName1,
        [parameter(Mandatory=$true, Position=2, ParameterSetName="NameWithId")]
        [guid]$Id1,
        [parameter(Mandatory=$true, Position=3, ParameterSetName="NameWithId")]
        [string]$EntityLogicalName2,
        [parameter(Mandatory=$true, Position=4, ParameterSetName="NameWithId")]
        [guid]$Id2,
        [parameter(Mandatory=$true, Position=5)]
        [string]$RelationshipName
    )

    if($conn -eq $null)
    {
        $connobj = Get-Variable conn -Scope global -ErrorAction SilentlyContinue
        if($connobj.Value -eq $null)
        {
            Write-Warning 'You need to create Connect to CRM Organization. Use Get-CrmConnection to create it.'
            break;
        }
        else
        {
            $conn = $connobj.Value
        }
    }    

    if($conn -eq $null)
    {
        $connobj = Get-Variable conn -Scope global -ErrorAction SilentlyContinue
        if($connobj.Value -eq $null)
        {
            Write-Warning 'You need to create Connect to CRM Organization. Use Get-CrmConnection to create it.'
            break;
        }
        else
        {
            $conn = $connobj.Value
        }
    }

    if($CrmRecord1 -ne $null)
    {
        $EntityLogicalName1 = $CrmRecord1.logicalname
        $Id1 = $CrmRecord1.($EntityLogicalName1 + "id")
    }

    if($CrmRecord2 -ne $null)
    {
        $EntityLogicalName2 = $CrmRecord2.logicalname
        $Id2 = $CrmRecord2.($EntityLogicalName2 + "id")
    }

    try
    {
        $result = $conn.DeleteEntityAssociation($EntityLogicalName1, $Id1, $EntityLogicalName2, $Id2, $RelationshipName, [Guid]::Empty)
    }
    catch
    {
        return $conn.LastCrmException
    }
}

#ExecuteWorkflowOnEntity  
function Invoke-CrmRecordWorkflow{

<#
 .SYNOPSIS
 Runs an on-demand workflow for a record.

 .DESCRIPTION
 The Invoke-CrmRecordWorkflow cmdlet lets you run an on-demand workflow for a record. 

 There are two ways to specify records.
 
 1. Pass record's Id.
 2. Get a record object by using Get-CrmRecord/Get-CrmRecords cmdlets, and pass it.

 You can specify on-demand workflow by using its name.

 .PARAMETER conn
 A connection to your CRM organizatoin. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .PARAMETER CrmRecord
 A record object which is obtained via Get-CrmRecord/Get-CrmRecords. When you pass CrmRecord, then you don't use Id.

 .PARAMETER Id
 An Id (guid) of first record.

 .PARAMETER WorkflowName
 An on-demand workflow name.

 .EXAMPLE
 Invoke-CrmRecordWorkflow -conn $conn -Id faff5970-6317-e511-80da-c4346bc43d94 -WorkflowName "Sample Workflow for Account"

 This example runs an on-demand workflow named "Sample Workflow for Accoutn" for an account.

 .EXAMPLE
 Invoke-CrmRecordWorkflow faff5970-6317-e511-80da-c4346bc43d94 "Sample Workflow for Account"
 
 This example runs an on-demand workflow named "Sample Workflow for Accoutn" for an account by ommiting parameter names.
 When ommiting parameter names, you do not provide $conn, cmdlets automatically finds it.

 .EXAMPLE
 PS C:\>$account = Get-CrmRecord account 00005a70-6317-e511-80da-c4346bc43d94 name

 PS C:\>Invoke-CrmRecordWorkflow $account "Sample Workflow for Account"

 This example runs an on-demand workflow named "Sample Workflow for Accoutn" for an account by ommiting parameter names.

#>

    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1, ParameterSetName="CrmRecord")]
        [PSObject]$CrmRecord,
        [parameter(Mandatory=$true, Position=1, ParameterSetName="Id")]
        [Alias("Id")]
        [string]$StringId,
        [parameter(Mandatory=$true, Position=2)]
        [string]$WorkflowName
    )

    if($conn -eq $null)
    {
        $connobj = Get-Variable conn -Scope global -ErrorAction SilentlyContinue
        if($connobj.Value -eq $null)
        {
            Write-Warning 'You need to create Connect to CRM Organization. Use Get-CrmConnection to create it.'
            break;
        }
        else
        {
            $conn = $connobj.Value
        }
    }

    if($CrmRecord -ne $null)
    {        
        $Id = $CrmRecord.($CrmRecord.logicalname + "id")
    }
    else
    {
        $Id = [guid]::Parse($StringId)
    }

    try
    {
        $result = $conn.ExecuteWorkflowOnEntity($WorkflowName, $Id, [Guid]::Empty)
    }
    catch
    {
        return $conn.LastCrmException
    }

    return $result
}

#GetMyCrmUserId
function Get-MyCrmUserId{

<#
 .SYNOPSIS
 Retrieves login user's CRM UserId (guid).

 .DESCRIPTION
 The Get-MyCrmUserId cmdlet retrieves login user's CRM UserId. 

 .PARAMETER conn
 A connection to your CRM organizatoin. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .EXAMPLE
 Get-MyCrmUserId -conn $conn

 This example returns login user's CRM UserId.

 .EXAMPLE
 Get-MyCrmUserId
 
 This example returns login user's CRM UserId by ommiting -conn parameter. 
 To omit conn parameter, you need creating $conn in advance, then cmdlets automatically finds it.

#>

    [CmdletBinding()]
    PARAM( 
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn
    )

    if($conn -eq $null)
    {
        $connobj = Get-Variable conn -Scope global -ErrorAction SilentlyContinue
        if($connobj.Value -eq $null)
        {
            Write-Warning 'You need to create Connect to CRM Organization. Use Get-CrmConnection to create it.'
            break;
        }
        else
        {
            $conn = $connobj.Value
        }
    }

    try
    {
        $response = $conn.GetMyCrmUserId()
    }
    catch
    {
        return $conn.LastCrmException
    }    

    return $response
}

#GetAllAttributesForEntity
function Get-CrmEntityAttributes{

<#
 .SYNOPSIS
 Retrieves all attributes metadata for an Entity.

 .DESCRIPTION
 The Get-CrmEntityAttributes cmdlet lets you retrieve all attributes metadata for an Entity. 

 .PARAMETER conn
 A connection to your CRM organizatoin. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .PARAMETER EntityLogicalName
 A logicalname for Entity. i.e.)accout, contact, lead, etc..

 .EXAMPLE
 Get-CrmEntityAttributes -conn $conn -EntityLogicalName account
 AttributeOf                 : preferredcontactmethodcode
 AttributeType               : Virtual
 AttributeTypeName           : Microsoft.Xrm.Sdk.Metadata.AttributeTypeDisplayName
 ColumnNumber                : 149
 Description                 : Microsoft.Xrm.Sdk.Label
 DisplayName                 : Microsoft.Xrm.Sdk.Label
 DeprecatedVersion           : 
 IntroducedVersion           : 5.0.0.0
 EntityLogicalName           : account
 IsAuditEnabled              : Microsoft.Xrm.Sdk.BooleanManagedProperty
 IsCustomAttribute           : False
 ...

 This example retrieves all attributes metadata for Account Entity.

 .EXAMPLE
 Get-CrmEntityAttributes account
 AttributeOf                 : preferredcontactmethodcode
 AttributeType               : Virtual
 AttributeTypeName           : Microsoft.Xrm.Sdk.Metadata.AttributeTypeDisplayName
 ColumnNumber                : 149
 Description                 : Microsoft.Xrm.Sdk.Label
 DisplayName                 : Microsoft.Xrm.Sdk.Label
 DeprecatedVersion           : 
 IntroducedVersion           : 5.0.0.0
 EntityLogicalName           : account
 IsAuditEnabled              : Microsoft.Xrm.Sdk.BooleanManagedProperty
 IsCustomAttribute           : False
 ...

 This example retrieves all attributes metadata for Account Entity by ommiting parameter names.
 When ommiting parameter names, you do not provide $conn, cmdlets automatically finds it.

 .EXAMPLE Get-CrmEntityAttributes account | Where {$_.IsCustomAttribute -eq $true} | Select logicalname

 This example retrieves all attributes metadata for an account, and filter for only custom fields. Then displays logicalname field only.

#>

 [CmdletBinding()]
    PARAM( 
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1)]
        [string]$EntityLogicalName
    )

    if($conn -eq $null)
    {
        $connobj = Get-Variable conn -Scope global -ErrorAction SilentlyContinue
        if($connobj.Value -eq $null)
        {
            Write-Warning 'You need to create Connect to CRM Organization. Use Get-CrmConnection to create it.'
            break;
        }
        else
        {
            $conn = $connobj.Value
        }
    } 
       
    try
    {
        $results = $conn.GetAllAttributesForEntity($EntityLogicalName)
    }
    catch
    {
        return $conn.LastCrmException
    }    

    return $results
}

#GetAllEntityMetadata 
function Get-CrmEntityAllMetadata{

<#
 .SYNOPSIS
 Retrieves all Metadata for CRM organization.

 .DESCRIPTION
 The Get-CrmEntityAllMetadata cmdlet lets you retrieve all Metadata for CRM organization. You can specify which type of Metadata you want to retrive by using EntityFilters parameter.

 .PARAMETER conn
 A connection to your CRM organizatoin. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .PARAMETER OnlyPublished
 Specify $True if you retrieve only published Metadata.

 .PARAMETER EntityFilters
 Specify which type of Metadata you want to retrieve. Valid options are "all", "attributes", "entity", "privileges", "relationships".

 .EXAMPLE
 Get-CrmEntityAllMetadata -conn $conn -OnlyPublished $True EntityFilters all
 ActivityTypeMask                  : 0
 Attributes                        : {, , Microsoft.Xrm.Sdk.Metadata.AttributeTypeDisplayName, ...}
 AutoRouteToOwnerQueue             : False
 CanTriggerWorkflow                : False
 Description                       : Microsoft.Xrm.Sdk.Label
 DisplayCollectionName             : Microsoft.Xrm.Sdk.Label
 DisplayName                       : Microsoft.Xrm.Sdk.Label
 EntityHelpUrlEnabled              : False
 EntityHelpUrl                     : 
 ...

 This example retrieves all published Metadata.

 .EXAMPLE
 Get-CrmEntityAllMetadata $True Entity
 ActivityTypeMask                  : 0
 Attributes                        : 
 AutoRouteToOwnerQueue             : False
 CanTriggerWorkflow                : False
 Description                       : Microsoft.Xrm.Sdk.Label
 DisplayCollectionName             : Microsoft.Xrm.Sdk.Label
 DisplayName                       : Microsoft.Xrm.Sdk.Label
 EntityHelpUrlEnabled              : False
 ...

 This example retrieves all published Entity metadata by ommiting parameter names.
 When ommiting parameter names, you do not provide $conn, cmdlets automatically finds it.

#>

 [CmdletBinding()]
    PARAM( 
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$false, Position=1)]
        [bool]$OnlyPublished=$true, 
        [parameter(Mandatory=$false, Position=2)]
        [string]$EntityFilters
    )

    if($conn -eq $null)
    {
        $connobj = Get-Variable conn -Scope global -ErrorAction SilentlyContinue
        if($connobj.Value -eq $null)
        {
            Write-Warning 'You need to create Connect to CRM Organization. Use Get-CrmConnection to create it.'
            break;
        }
        else
        {
            $conn = $connobj.Value
        }
    } 
    
    switch($EntityFilters.ToLower())
    {
        "all" {
            $filter = [Microsoft.Xrm.Sdk.Metadata.EntityFilters]::All
            break             
        }
        "attributes" {
            $filter = [Microsoft.Xrm.Sdk.Metadata.EntityFilters]::Attributes
            break 
        } 
        "entity" {
            $filter = [Microsoft.Xrm.Sdk.Metadata.EntityFilters]::Entity
            break
        }  
        "privileges" {
            $filter = [Microsoft.Xrm.Sdk.Metadata.EntityFilters]::Privileges
            break
        }  
        "relationships" {
            $filter = [Microsoft.Xrm.Sdk.Metadata.EntityFilters]::Relationships
            break
        }
        default {
            $filter = [Microsoft.Xrm.Sdk.Metadata.EntityFilters]::Default
            break
        }               
    }

    try
    {
        $results = $conn.GetAllEntityMetadata($OnlyPublished, $filter)
    }
    catch
    {
        return $conn.LastCrmException
    }    

    return $results
}

#GetEntityAttributeMetadataForAttribute  
function Get-CrmEntityAttributeMetadata{

<#
 .SYNOPSIS
 Retrieves an attribute metadata for an Entity.

 .DESCRIPTION
 The Get-CrmEntityAttributeMetadata cmdlet lets you retrieve an attribute metadata for an Entity. 

 .PARAMETER conn
 A connection to your CRM organizatoin. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .PARAMETER EntityLogicalName
 A logicalname for Entity. i.e.)accout, contact, lead, etc..

 .PARAMETER FieldLogicalName
 A logicalname for a field.

 .EXAMPLE
 Get-CrmEntityAttributeMetadata -conn $conn -EntityLogicalName account -FieldLogicalName parentaccountid
 Targets                     : {account}
 AttributeOf                 : 
 AttributeType               : Lookup
 AttributeTypeName           : Microsoft.Xrm.Sdk.Metadata.AttributeTypeDisplayName
 ColumnNumber                : 57
 Description                 : Microsoft.Xrm.Sdk.Label
 DisplayName                 : Microsoft.Xrm.Sdk.Label
 DeprecatedVersion           : 
 IntroducedVersion           : 5.0.0.0
 EntityLogicalName           : account
 ...

 This example retrieves Parent Account attribute metadata for Account Entity.

 .EXAMPLE
 Get-CrmEntityAttributeMetadata account parentaccountid
 AttributeOf                 : preferredcontactmethodcode
 AttributeType               : Virtual
 AttributeTypeName           : Microsoft.Xrm.Sdk.Metadata.AttributeTypeDisplayName
 ColumnNumber                : 149
 Description                 : Microsoft.Xrm.Sdk.Label
 DisplayName                 : Microsoft.Xrm.Sdk.Label
 DeprecatedVersion           : 
 IntroducedVersion           : 5.0.0.0
 EntityLogicalName           : account
 IsAuditEnabled              : Microsoft.Xrm.Sdk.BooleanManagedProperty
 IsCustomAttribute           : False
 ...

 This example retrieves Parent Account attribute metadata for Account Entity by ommiting parameter names.
 When ommiting parameter names, you do not provide $conn, cmdlets automatically finds it.
 
#>

 [CmdletBinding()]
    PARAM( 
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1)]
        [string]$EntityLogicalName,
        [parameter(Mandatory=$true, Position=2)]
        [string]$FieldLogicalName
    )

    if($conn -eq $null)
    {
        $connobj = Get-Variable conn -Scope global -ErrorAction SilentlyContinue
        if($connobj.Value -eq $null)
        {
            Write-Warning 'You need to create Connect to CRM Organization. Use Get-CrmConnection to create it.'
            break;
        }
        else
        {
            $conn = $connobj.Value
        }
    } 
    
    try
    {
        $result = $conn.GetEntityAttributeMetadataForAttribute($EntityLogicalName, $FieldLogicalName)
    }
    catch
    {
        return $conn.LastCrmException
    }    

    return $result
}

#GetEntityDataByFetchSearch
function Get-CrmRecordsByFetch{

<#
 .SYNOPSIS
 Retrieves CRM records by using FetchXML query.

 .DESCRIPTION
 The Get-CrmRecordsByFetch cmdlet lets you retrieve up to 5,000 records from your CRM organization by using FetchXML query. 
 The output contains CrmRecords (List or retrieved records), PagingCookie (for next iteration), and NextPage (to indicate if there are more records on the next page).

 if you need to paging the result, you can also specify TopCount, PageNumber and PagingCookie.

 You can obtain FetchXML by using Advanced Find tool. As FetchXML query can be multiple lines, use "@ ... @" syntax to speficy the query. 

 .PARAMETER conn
 A connection to your CRM organizatoin. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .PARAMETER Fetch
 A FetchXML query to retrive records. You can obtain FetchXML by using Advanced Find tool. Use "@ ... @" syntax to speficy the query. 

 .PARAMETER TopCount
 Specify how many records you need to retireve at once (up to 5,000 at once).

 .PARAMETER PageNumber
 Specify starting page number for paging. Starting from 1.

 .PARAMETER PageCookie
 Specify cookie string for paging. Keep previous PagingCookie value for next iteration.

 .PARAMETER AllRows
 By default the first 5000 rows are returned, this switch will bring back all results regardless of how many

 .EXAMPLE
 Get-CrmRecordsByFetch -conn $conn -Fetch @"
 <fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="false" no-lock="true">
   <entity name="account">
     <attribute name="name" />
     <attribute name="primarycontactid" />
     <attribute name="telephone1" />
     <attribute name="accountid" />
     <order attribute="name" descending="false" />
   </entity>
 </fetch>
 "@
 Key             Value
 ---             -----                                                                                  
 CrmRecords      {@{name_Property=[name, A. Datum Corporation (sample)]; name=A. Datum Corporation (s...
 Count           5
 PagingCookie   <cookie page="1"><name last="Litware, Inc. (sample)" first="A. Datum Corporation (sa...
 NextPage    False
 
 This example retrieves account records by using FetchXML and results contains CrmRecords, PagingCookie and NextPage.
 Please note that copying and pasting above example may not work due to multiline issue. Please remove all whitespace before last @"

 .EXAMPLE

 $result = Get-CrmRecordsByFetch @"
 <fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="false" no-lock="true">
   <entity name="account">
     <attribute name="name" />
     <attribute name="primarycontactid" />
     <attribute name="telephone1" />
     <attribute name="accountid" />
     <order attribute="name" descending="false" />
   </entity>
 </fetch>
 "@

 PS C:\>$result.CrmRecords
 name_Property             : [name, A. Datum Corporation (sample)]
 name                      : A. Datum Corporation (sample)
 primarycontactid_Property : [primarycontactid, Microsoft.Xrm.Sdk.EntityReference]
 primarycontactid          : Rene Valdes (sample)
 telephone1_Property       : [telephone1, 555-0158]
 telephone1                : 555-0158
 accountid_Property        : [accountid, be02caab-6c16-e511-80d6-c4346bc43dc0]
 accountid                 : be02caab-6c16-e511-80d6-c4346bc43dc0
 returnProperty_EntityName : account
 returnProperty_Id         : be02caab-6c16-e511-80d6-c4346bc43dc0
 original                  : {[name_Property, [name, A. Datum Corporation (sample)]], [name, A. Datum Corporation (sample)], [primarycontactid_Property, [primarycontactid, 
                             Microsoft.Xrm.Sdk.EntityReference]], [primarycontactid, Microsoft.Xrm.Sdk.EntityReference]...}
 logicalname               : account
 
 name_Property             : [name, Adventure Works (sample)]
 name                      : Adventure Works (sample)
 primarycontactid_Property : [primarycontactid, Microsoft.Xrm.Sdk.EntityReference]
 primarycontactid          : Nancy Anderson (sample)
 telephone1_Property       : [telephone1, 555-0152] 
 ....

 This example stores retrieved result into $result variable, then show records by CrmRecords property.
 When ommiting parameter names, you do not provide $conn, cmdlets automatically finds it.

 .EXAMPLE

 $result1 = Get-CrmRecordsByFetch @"
 <fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="false" no-lock="true">
   <entity name="contact">
     <attribute name="fullname" />
   </entity>
 </fetch>
 "@

 PS C:\>$result1.CrmRecords.Count
 5000

 PS C:\>$result1.NextPage
 True

 PS C:\>$result1.PagingCookie
 <cookie page="1"><contactid last="{890FDA88-4217-E511-80DB-C4346BC42D18}" first="{E1D47674-4017-E511-80DB-C4346BC42D18}" /></cookie>

 PS C:\>$result2 = Get-CrmRecordsByFetch -conn $conn -Fetch @"
 <fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="false" no-lock="true">
   <entity name="contact">
     <attribute name="fullname" />
   </entity>
 </fetch>
 "@ -TopCount 5000 -PageNumber 2 -PageCookie $result1.PagingCookie

 PS C:\>$result2.CrmRecords.Count
 453

 PS C:\>$result2.NextPage
 False

 This example stores retrieved result into $result1 variable, then use $result1.PagingCookie data for next iteration.
 FetchXML is exactly same but specifying additional Parameters for second command.

#>

    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1)]
        [string]$Fetch,
        [parameter(Mandatory=$false, Position=2)]
        [int]$TopCount,
        [parameter(Mandatory=$false, Position=3)]
        [int]$PageNumber,
        [parameter(Mandatory=$false, Position=4)]
        [string]$PageCookie,
        [parameter(Mandatory=$false, Position=5)]
        [switch]$AllRows

    )
    
    if($conn -eq $null)
    {
        $connobj = Get-Variable conn -Scope global -ErrorAction SilentlyContinue
        if($connobj.Value -eq $null)
        {
            Write-Warning 'Get-CrmRecordsByFetch(): You need to create Connect to CRM Organization. Use Get-CrmConnection to create it.'
            break;
        }
        else
        {
            $conn = $connobj.Value
        }
    }
    
    #default page number to 1 if not supplied
    if($PageNumber -eq 0)
    {
        $PageNumber = 1
    }
    
    $PagingCookie = ""
    $NextPage = $false
    
    if($PageCookie -eq "")
    {
        $PageCookie = $null
    }

    $recordslist = New-Object 'System.Collections.Generic.List[System.Management.Automation.PSObject]'
    $resultSet = New-Object 'System.Collections.Generic.Dictionary[[System.String],[System.Management.Automation.PSObject]]'

    try
    {
        Write-Debug "Getting data from CRM"
        $records = $conn.GetEntityDataByFetchSearch($Fetch, $TopCount, $PageNumber, $PageCookie, [ref]$PagingCookie, [ref]$NextPage, [Guid]::Empty)
        $xml = [xml]$Fetch
        $logicalname = $xml.SelectSingleNode("/fetch/entity").Name

        #if there are zero results returned 
        if($records.Count -eq 0)
        {
            $error = "No Result" 
            Write-Warning $error
            $resultSet.Add("CrmRecords", $recordslist)
            $resultSet.Add("Count", $recordslist.Count)
            $resultSet.Add("PagingCookie",$null)
            $resultSet.Add("NextPage",$false)
            
            #EXIT
            return $resultSet; 
        }
        #if we have records
        elseif($records.Count -gt 0)
        {
            Write-Debug "Records Found!"
            foreach($record in $records.Values)
            {   
                $psobj = New-Object -TypeName System.Management.Automation.PSObject
        
                foreach($att in $record.GetEnumerator())
                {
                    if($att.Value -is [Microsoft.Xrm.Sdk.EntityReference])
                    {
                        Add-Member -InputObject $psobj -MemberType NoteProperty -Name $att.Key -Value $att.Value.Name
                    }
					elseif($att.Value -is [Microsoft.Xrm.Sdk.AliasedValue])
					{
						Add-Member -InputObject $psobj -MemberType NoteProperty -Name $att.Key -Value $att.Value.Value
					}
                    else
                    {
                        Add-Member -InputObject $psobj -MemberType NoteProperty -Name $att.Key -Value $att.Value
                    }
                }  

                Add-Member -InputObject $psobj -MemberType NoteProperty -Name "original" -Value $record
                Add-Member -InputObject $psobj -MemberType NoteProperty -Name "logicalname" -Value $logicalname

                $recordslist.Add($psobj)
            }
            #IF we have multiple pages!
            if($NextPage -and $AllRows)  
            {
                $PageNumber = $PageNumber + 1
                Write-Debug "Fetching next page #$PageNumber"
                $NextRecordSet = Get-CrmRecordsByFetch -conn $conn -Fetch $Fetch -TopCount $TopCount -PageNumber $PageNumber -PageCookie $PagingCookie -AllRows
                if($NextRecordSet.CrmRecords.Count -gt 0)
                {
                    Write-Debug "Adding data to original results from page#: $PageNumber"
                    $recordslist.AddRange($NextRecordSet.CrmRecords)
                }
            }
        }
    }
    catch
    {
        Write-Error $_.Exception
        return $conn.LastCrmException
    }
    
    $resultSet = New-Object 'System.Collections.Generic.Dictionary[[System.String],[System.Management.Automation.PSObject]]'
    $resultSet.Add("CrmRecords", $recordslist)
    $resultSet.Add("Count", $recordslist.Count)
    $resultSet.Add("PagingCookie",$PagingCookie)
    $resultSet.Add("NextPage",$NextPage)
    $resultSet.Add("FetchXml", $Fetch)

    return $resultSet
}

#GetEntityDisplayName
function Get-CrmEntityDisplayName{

<#
 .SYNOPSIS
 Retrieves Display Name for an Entity.

 .DESCRIPTION
 The Get-CrmEntityDisplayName cmdlet lets you retrieve Display Name for an Entity.

 .PARAMETER conn
 A connection to your CRM organizatoin. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .PARAMETER EntityLogicalName
 A logicalname for Entity. i.e.)accout, contact, lead, etc..

 .EXAMPLE
 Get-CrmEntityDisplayName -conn $conn -EntityLogicalName account
 Account

 This example retrieves Display Name for Account Entity.

 .EXAMPLE
 Get-CrmEntityDisplayName incident
 Case

 This example retrieves Display Name for Incident Entity by ommiting parameter names.
 When ommiting parameter names, you do not provide $conn, cmdlets automatically finds it.
 
#>

 [CmdletBinding()]
    PARAM( 
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1, ParameterSetName="EntityLogicalName")]
        [string]$EntityLogicalName
    )

    if($conn -eq $null)
    {
        $connobj = Get-Variable conn -Scope global -ErrorAction SilentlyContinue
        if($connobj.Value -eq $null)
        {
            Write-Warning 'You need to create Connect to CRM Organization. Use Get-CrmConnection to create it.'
            break;
        }
        else
        {
            $conn = $connobj.Value
        }
    } 

    try
    {
        $result = $conn.GetEntityDisplayName($EntityLogicalName)
    }
    catch
    {
        return $conn.LastCrmException
    }    

    return $result
}

#GetEntityDisplayNamePlural
function Get-CrmEntityDisplayPluralName{

<#
 .SYNOPSIS
 Retrieves Display Plural Name for an Entity.

 .DESCRIPTION
 The Get-CrmEntityDisplayPluralName cmdlet lets you retrieve Display Plural Name for an Entity.

 .PARAMETER conn
 A connection to your CRM organizatoin. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .PARAMETER EntityLogicalName
 A logicalname for Entity. i.e.)accout, contact, lead, etc..

 .EXAMPLE
 Get-CrmEntityDisplayPluralName -conn $conn -EntityLogicalName account
 Accounts

 This example retrieves Display Name for Account Entity.

 .EXAMPLE
 Get-CrmEntityDisplayPluralName incident
 Cases

 This example retrieves Display Name for Incident Entity by ommiting parameter names.
 When ommiting parameter names, you do not provide $conn, cmdlets automatically finds it.
 
#>

 [CmdletBinding()]
    PARAM( 
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1)]
        [string]$EntityLogicalName
    )

    if($conn -eq $null)
    {
        $connobj = Get-Variable conn -Scope global -ErrorAction SilentlyContinue
        if($connobj.Value -eq $null)
        {
            Write-Warning 'You need to create Connect to CRM Organization. Use Get-CrmConnection to create it.'
            break;
        }
        else
        {
            $conn = $connobj.Value
        }
    } 

    try
    {
        $result = $conn.GetEntityDisplayNamePlural($EntityLogicalName)
    }
    catch
    {
        return $conn.LastCrmException
    }     

    return $result
}

#GetEntityMetadata
function Get-CrmEntityMetadata{

<#
 .SYNOPSIS
 Retrieves Metadata for an Entity.

 .DESCRIPTION
 The Get-CrmEntityMetadata cmdlet lets you retrieve Metadata for an Entity. You can specify which type of Metadata you want to retrive by using EntityFilters parameter.

 .PARAMETER conn
 A connection to your CRM organizatoin. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .PARAMETER EntityLogicalName
 A logicalname for Entity. i.e.)accout, contact, lead, etc..

 .PARAMETER EntityFilters
 Specify which type of Metadata you want to retrieve. Valid options are "all", "attributes", "entity", "privileges", "relationships".

 .EXAMPLE
 Get-CrmEntityMetadata -conn $conn -EntityLogicalName account EntityFilters all
 ActivityTypeMask                  : 0
 Attributes                        : {Microsoft.Xrm.Sdk.Metadata.AttributeTypeDisplayName, Microsoft.Xrm.Sdk.Metadata.StringFormatName, 
                                     Microsoft.Xrm.Sdk.Metadata.StringFormatName, Microsoft.Xrm.Sdk.Metadata.StringFormatName...}
 AutoRouteToOwnerQueue             : False
 CanTriggerWorkflow                : True
 Description                       : Microsoft.Xrm.Sdk.Label
 DisplayCollectionName             : Microsoft.Xrm.Sdk.Label
 ...

 This example retrieves all Metadata for Account Entity.

 .EXAMPLE
 Get-CrmEntityMetadata account relationships
 ActivityTypeMask                  : 0
 Attributes                        : {Microsoft.Xrm.Sdk.Metadata.AttributeTypeDisplayName, Microsoft.Xrm.Sdk.Metadata.StringFormatName, 
                                     Microsoft.Xrm.Sdk.Metadata.StringFormatName, Microsoft.Xrm.Sdk.Metadata.StringFormatName...}
 AutoRouteToOwnerQueue             : False
 CanTriggerWorkflow                : True
 Description                       : Microsoft.Xrm.Sdk.Label
 DisplayCollectionName             : Microsoft.Xrm.Sdk.Label
 ...

 This example retrieves all Metadata for Account Entity by ommiting parameter names.
 When ommiting parameter names, you do not provide $conn, cmdlets automatically finds it.

#>

 [CmdletBinding()]
    PARAM( 
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1)]
        [string]$EntityLogicalName,
        [parameter(Mandatory=$false, Position=2)]
        [string]$EntityFilters
    )

    if($conn -eq $null)
    {
        $connobj = Get-Variable conn -Scope global -ErrorAction SilentlyContinue
        if($connobj.Value -eq $null)
        {
            Write-Warning 'You need to create Connect to CRM Organization. Use Get-CrmConnection to create it.'
            break;
        }
        else
        {
            $conn = $connobj.Value
        }
    } 

    switch($EntityFilters.ToLower())
    {
        "all" {
            $filter = [Microsoft.Xrm.Sdk.Metadata.EntityFilters]::All
            break             
        }
        "attributes" {
            $filter = [Microsoft.Xrm.Sdk.Metadata.EntityFilters]::Attributes
            break 
        } 
        "entity" {
            $filter = [Microsoft.Xrm.Sdk.Metadata.EntityFilters]::Entity
            break
        }  
        "privileges" {
            $filter = [Microsoft.Xrm.Sdk.Metadata.EntityFilters]::Privileges
            break
        }  
        "relationships" {
            $filter = [Microsoft.Xrm.Sdk.Metadata.EntityFilters]::Relationships
            break
        }
        default {
            $filter = [Microsoft.Xrm.Sdk.Metadata.EntityFilters]::Default
            break
        }               
    }

    try
    {
        $results = $conn.GetEntityMetadata($EntityLogicalName, $filter)
    }
    catch
    {
        return $conn.LastCrmException
    }    

    return $results
}

#GetEntityName
function Get-CrmEntityName{

<#
 .SYNOPSIS
 Retrieves Entity logicalname for EntityTypeCode.

 .DESCRIPTION
 The Get-CrmEntityName cmdlet lets you retrieve Entity logicalname. You can specify EntityTypeCode.

 .PARAMETER conn
 A connection to your CRM organizatoin. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .PARAMETER EntityTypeCode
 A number for Entity.

 .EXAMPLE
 Get-CrmEntityName -conn $conn -EntityTypeCode 1
 account

 This example retrieves Entity logicalname for EntityTypeCode 1 (Account).

 .EXAMPLE
 Get-CrmEntityName 4200
 activitypointer

 This example retrieves Entity logicalname for EntityTypeCode 4200 (Activity) by ommiting parameter names.
 When ommiting parameter names, you do not provide $conn, cmdlets automatically finds it.

#>

 [CmdletBinding()]
    PARAM( 
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1)]
        [int]$EntityTypeCode
    )

    if($conn -eq $null)
    {
        $connobj = Get-Variable conn -Scope global -ErrorAction SilentlyContinue
        if($connobj.Value -eq $null)
        {
            Write-Warning 'You need to create Connect to CRM Organization. Use Get-CrmConnection to create it.'
            break;
        }
        else
        {
            $conn = $connobj.Value
        }
    }  

    try
    {
        $result = $conn.GetEntityName($EntityTypeCode)
    }
    catch
    {
        return $conn.LastCrmException
    }   

    return $result
}

#GetEntityTypeCode 
function Get-CrmEntityTypeCode{

<#
 .SYNOPSIS
 Retrieves EntityTypeCode for an Entity.

 .DESCRIPTION
 The Get-CrmEntityTypeCode cmdlet lets you retrieve EntityTypeCode. You can specify EntityLogicalName.

 .PARAMETER conn
 A connection to your CRM organizatoin. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .PARAMETER EntityLogicalName
 A logicalname for Entity. i.e.)accout, contact, lead, etc..

 .EXAMPLE
 Get-CrmEntityTypeCode -conn $conn -EntityLogicalName account
 1

 This example retrieves EntityTypeCode for Account Entity.

 .EXAMPLE
 Get-CrmEntityTypeCode lead
 4

 This example retrieves EntityTypeCode for Lead Entity by ommiting parameter names.
 When ommiting parameter names, you do not provide $conn, cmdlets automatically finds it.

#>

 [CmdletBinding()]
    PARAM( 
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1)]
        [string]$EntityLogicalName
    )

    if($conn -eq $null)
    {
        $connobj = Get-Variable conn -Scope global -ErrorAction SilentlyContinue
        if($connobj.Value -eq $null)
        {
            Write-Warning 'You need to create Connect to CRM Organization. Use Get-CrmConnection to create it.'
            break;
        }
        else
        {
            $conn = $connobj.Value
        }
    }  

    try
    {
        $result = $conn.GetEntityTypeCode($EntityLogicalName)
    }
    catch
    {
        return $conn.LastCrmException
    }    

    return $result
}

#GetGlobalOptionSetMetadata  
function Get-CrmGlobalOptionSet{

<#
 .SYNOPSIS
 Retrieves a global OptionSet.

 .DESCRIPTION
 The Get-CrmGlobalOptionSet cmdlet lets you retrieve a global OptionSet.

 .PARAMETER conn
 A connection to your CRM organizatoin. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .PARAMETER OptionSetName
 A logicalname for a global OptionSet.

 .EXAMPLE
 Get-CrmGlobalOptionSet -conn $conn -OptionSetName incident_caseorigincode
 Options           : {, , , ...}
 Description       : Microsoft.Xrm.Sdk.Label
 DisplayName       : Microsoft.Xrm.Sdk.Label
 IsCustomOptionSet : False
 IsGlobal          : True
 IsManaged         : True
 IsCustomizable    : Microsoft.Xrm.Sdk.BooleanManagedProperty
 Name              : incident_caseorigincode
 OptionSetType     : Picklist
 IntroducedVersion : 
 MetadataId        : 08fa2cb2-e3fe-497a-9b5d-ee887f5cc3cd
 HasChanged        : 
 ExtensionData     : System.Runtime.Serialization.ExtensionDataObject

 This example retrieves incident_caseorigincode global OptionSet.

 .EXAMPLE
 Get-CrmGlobalOptionSet incident_caseorigincode
 Options           : {, , , ...}
 Description       : Microsoft.Xrm.Sdk.Label
 DisplayName       : Microsoft.Xrm.Sdk.Label
 IsCustomOptionSet : False
 IsGlobal          : True
 IsManaged         : True
 IsCustomizable    : Microsoft.Xrm.Sdk.BooleanManagedProperty
 Name              : incident_caseorigincode
 OptionSetType     : Picklist
 IntroducedVersion : 
 MetadataId        : 08fa2cb2-e3fe-497a-9b5d-ee887f5cc3cd
 HasChanged        : 
 ExtensionData     : System.Runtime.Serialization.ExtensionDataObject

 This example retrieves incident_caseorigincode global OptionSet by ommiting parameter names.
 When ommiting parameter names, you do not provide $conn, cmdlets automatically finds it.

 .EXAMPLE
 PS C:\>$optionset = Get-CrmGlobalOptionSet incident_caseorigincode
 PS C:\>$optionset.Options | % {[string]$_.Value + ":" + $_.Label.LocalizedLabels.Label}
 1:Phone
 2:Email
 3:Web
 2483:Facebook
 3986:Twitter

 This example retrieves incident_caseorigincode global OptionSet and stores to variable, then query its value and display name.

#>

 [CmdletBinding()]
    PARAM( 
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1)]
        [string]$OptionSetName
    )

    if($conn -eq $null)
    {
        $connobj = Get-Variable conn -Scope global -ErrorAction SilentlyContinue
        if($connobj.Value -eq $null)
        {
            Write-Warning 'You need to create Connect to CRM Organization. Use Get-CrmConnection to create it.'
            break;
        }
        else
        {
            $conn = $connobj.Value
        }
    }  

    try
    {
        $results = $conn.GetGlobalOptionSetMetadata($OptionSetName)
    }
    catch
    {
        return $conn.LastCrmException
    }

    return $results
}

#GetPickListElementFromMetadataEntity   
function Get-CrmEntityOptionSet{

<#
 .SYNOPSIS
 Retrieves a picklist field of an Entity.

 .DESCRIPTION
 The Get-CrmEntityOptionSet cmdlet lets you retrieve a picklist field of an Entity.

 .PARAMETER conn
 A connection to your CRM organizatoin. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .PARAMETER EntityLogicalName
 A logicalname for Entity. i.e.)accout, contact, lead, etc..

 .PARAMETER FieldLogicalName
 A logicalname for a picklist filed.

 .EXAMPLE
 Get-CrmEntityOptionSet -conn $conn -EntityLogicalName account -FieldLogicalName statuscode
 ActualValue   PickListLabel   DisplayValue   Items
 -----------   -------------   ------------   -----
                               Status Reason  {1, 2}  

 This example retrieves statuscode picklist for Account Entity.

 .EXAMPLE
 Get-CrmEntityOptionSet account statuscode
 ActualValue   PickListLabel   DisplayValue   Items
 -----------   -------------   ------------   -----
                               Status Reason  {1, 2} 

 This example retrieves statuscode picklist for Account Entity by ommiting parameter names.
 When ommiting parameter names, you do not provide $conn, cmdlets automatically finds it.

 .EXAMPLE
 Get-CrmEntityOptionSet account statuscode | % {$_.Items}
 DisplayLabel   PickListItemId
 ------------   --------------
 Active         1
 Inactive       2

 This example retrieves statuscode picklist for Account Entity and displays Label and Id.

#>

 [CmdletBinding()]
    PARAM( 
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1)]
        [string]$EntityLogicalName,
        [parameter(Mandatory=$true, Position=2)]
        [string]$FieldLogicalName
    )

    if($conn -eq $null)
    {
        $connobj = Get-Variable conn -Scope global -ErrorAction SilentlyContinue
        if($connobj.Value -eq $null)
        {
            Write-Warning 'You need to create Connect to CRM Organization. Use Get-CrmConnection to create it.'
            break;
        }
        else
        {
            $conn = $connobj.Value
        }
    }

    try
    {
        $result = $conn.GetPickListElementFromMetadataEntity($EntityLogicalName, $FieldLogicalName)
    }
    catch
    {
        return $conn.LastCrmException
    }

    return $result
}

#ImportSolutionToCrm   
function Import-CrmSolution{

<#
 .SYNOPSIS
 Imports solution file to CRM Organization.

 .DESCRIPTION
 The Import-CrmSolution cmdlet lets you import solution file to CRM Organization and returns Job Id. You can use the Job Id to check import progress.

 .PARAMETER conn
 A connection to your CRM organizatoin. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .PARAMETER SolutionFilePath
 A file path to solution zip file.

 .PARAMETER ActivatePlugIns
 Specify the parameter to active plug-ins when importing solution.

 .PARAMETER OverwriteUnManagedCustomizations
 Specify the parameter to overwrite conflicting unmanaged customizations when importing solution.

 .PARAMETER SkipDependancyOnProductUpdateCheckOnInstall
 Specify the parameter to skip dependency check when importing solution.

 .PARAMETER PublishChanges
 Specify the parameter to publish all customizations (applicable only for unmanaged solution)

 .EXAMPLE
 Import-CrmSolution -conn $conn -SolutionFilePath "C:\SampleSolution_1_0_0_0.zip"

 This example imports solution and returns JobId.

 .EXAMPLE
 Import-CrmSolution "C:\SampleSolution_1_0_0_0.zip" $True

 This example imports solution by activating plug-ins and returns JobId by ommiting parameter names.
 When ommiting parameter names, you do not provide $conn, cmdlets automatically finds it.

#>

    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1)]
        [string]$SolutionFilePath,
        [parameter(Mandatory=$false, Position=2)]
        [switch]$ActivatePlugIns,
        [parameter(Mandatory=$false, Position=3)]
        [switch]$OverwriteUnManagedCustomizations,
        [parameter(Mandatory=$false, Position=4)]
        [switch]$SkipDependancyOnProductUpdateCheckOnInstall, 
        [parameter(Mandatory=$false, Position=5)]
        [switch]$PublishChanges
    )

    if($conn -eq $null)
    {
        $connobj = Get-Variable conn -Scope global -ErrorAction SilentlyContinue
        if($connobj.Value -eq $null)
        {
            Write-Error 'You need to create a connection to a CRM Organization using get-CrmConnection or pass the connection as a parameter to use this cmdlet.'
            break;
        }
        else
        {
            $conn = $connobj.Value
        }
    }   

    $importId = [guid]::Empty

    try
    {
        $tmpDest = $conn.CrmConnectOrgUriActual
        Write-Verbose "Importing solution file into: $tmpDest" 
        Write-Verbose "Importing solution file: $SolutionFilePath"
        Write-Verbose "OverwriteCustomizations: $OverwriteUnManagedCustomizations"
        Write-Verbose "SkipDependancyCheck: $SkipDependancyOnProductUpdateCheckOnInstall"
        Write-Verbose "Please wait while importing"
        $result = $conn.ImportSolutionToCrm($SolutionFilePath, [ref]$importId, $ActivatePlugIns,
                $OverwriteUnManagedCustomizations, $SkipDependancyOnProductUpdateCheckOnInstall)
              
        Start-Sleep -Seconds 5;        
        $import = Get-CrmRecord -conn $conn -EntityLogicalName importjob -Id $importId -Fields data,completedon,startedon,progress    
        
        $xml = [xml]($import).data
        $importresult = $xml.importexportxml.solutionManifests.solutionManifest.result
        
        $ProcPercent = $import.progress
        $ProcStart = $import.startedon
        $ProcComplete = $import.completedon

        $stillProcessing = if($ProcComplete -eq $null) {$true} else {$false}; 

        write-verbose "Import of file completed, importId: $importId - Progress: $ProcPercent Start: $ProcStart Complete: $ProcComplete"
        
        $delay = 3;
        $loopCount = 0;  
        write-verbose "ImportJob start time is: $ProcStart - polling job for completion time." 
        while($stillProcessing -and $ProcComplete -eq $null -and $loopCount -lt 80)
        {
            $import = Get-CrmRecord -conn $conn -EntityLogicalName importjob -Id $importId -Fields data,completedon,startedon,progress
            $importManifest = ([xml]($import).data).importexportxml.solutionManifests.solutionManifest
	        $importresult = $importManifest.result;
            $impResult = $importManifest.result.result

            $ProcPercent = $import.progress
            $ProcStart = $import.startedon
            $ProcComplete = $import.completedon

            $stillProcessing = if($ProcComplete -eq $null) {$true} else {$false}; 

            $loopCount = $loopCount+1;
            $waitSeconds = $loopCount * $delay
            write-verbose "Waiting for completion, waiting for $waitSeconds seconds... Progress: $ProcPercent"; 
            Start-Sleep -Seconds $delay;
        }
	    
        write-verbose "Import completed at: $ProcComplete"; 

        if($importresult.result -eq "failure" -or $ProcPercent -lt 100) #Must look at %age instead of this result as the result is usually wrong!
        {
            $importresultresult = $importresult.result
            $importresulterrortext = $importresult.errortext
            Write-Verbose "Import result: $importresultresult"
            Write-Verbose "Import result: $importresulterrortext - job with ID: $importId failed at $ProcPercent complete."
            throw $importresulterrortext
        }
        else
        {
            $managedsolution = $xml.importexportxml.solutionManifests.solutionManifest.Managed
            if($managedsolution -ne 1)
            {
                if($PublishChanges){
                    Write-Verbose "Guid populated and user requested publish changes request..."
                    Write-Verbose "Executing command: Publish-CrmAllCustomization, passing in the same connection"
            
                    Publish-CrmAllCustomization -conn $conn
                }
                else{
                    Write-Verbose "Import Complete don't forget to publish customizations."
                }
            }
            #return $ImportId
        }
    }
    catch
    {
        Write-Error $_.Exception
    }    
}

#InstallSampleDataToCrm    
function Add-CrmSampleData{

<#
 .SYNOPSIS
 Add sample data CRM Organization.

 .DESCRIPTION
 The Add-CrmSampleData cmdlet lets you add sample data to CRM Organization and returns Job Id. You can confirm the status by using Test-CrmSampleDataInstalled cmdlet.

 .PARAMETER conn
 A connection to your CRM organizatoin. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .EXAMPLE
 Add-CrmSampleData -conn $conn
 Guid
 ----
 9f7ff9ec-d2b5-4599-9909-04a072e0a546 

 This example adds sample data.

 .EXAMPLE
 Add-CrmSampleData
 Guid
 ----
 9f7ff9ec-d2b5-4599-9909-04a072e0a546 

 This example adds sample data by ommiting parameter names.
 When ommiting parameter names, you do not provide $conn, cmdlets automatically finds it.

#>

    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false, Position=0)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn
    )

    if($conn -eq $null)
    {
        $connobj = Get-Variable conn -Scope global -ErrorAction SilentlyContinue
        if($connobj.Value -eq $null)
        {
            Write-Warning 'You need to create Connect to CRM Organization. Use Get-CrmConnection to create it.'
            break;
        }
        else
        {
            $conn = $connobj.Value
        }
    } 

    try
    {
        $result = $conn.InstallSampleDataToCrm()
    }
    catch
    {
        return $conn.LastCrmException
    }

    return $result
}

#IsSampleDataInstalled    
function Test-CrmSampleDataInstalled{

<#
 .SYNOPSIS
 Checks if sample data has been installed to CRM Organization.

 .DESCRIPTION
 The Test-CrmSampleDataInstalled cmdlet lets you Check if sample data has been installed to CRM Organization.

 .PARAMETER conn
 A connection to your CRM organizatoin. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .EXAMPLE
 Test-CrmSampleDataInstalled -conn $conn
 Completed

 This example checks if sample data has been installed.

 .EXAMPLE
 Test-CrmSampleDataInstalled 
 Completed

 This example checks if sample data has been installed by ommiting parameter names.
 When ommiting parameter names, you do not provide $conn, cmdlets automatically finds it.

#>

    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn
    )

    if($conn -eq $null)
    {
        $connobj = Get-Variable conn -Scope global -ErrorAction SilentlyContinue
        if($connobj.Value -eq $null)
        {
            Write-Warning 'You need to create Connect to CRM Organization. Use Get-CrmConnection to create it.'
            break;
        }
        else
        {
            $conn = $connobj.Value
        }
    } 

    try
    {
        $result = $conn.IsSampleDataInstalled()
    }
    catch
    {
        return $conn.LastCrmException
    }    

    return $result
}

#PublishEntity  
function Publish-CrmEntity{

<#
 .SYNOPSIS
 Publishes customization for an Entity.

 .DESCRIPTION
 The Publish-CrmEntity cmdlet lets you publish customization for an Entity.

 .PARAMETER conn
 A connection to your CRM organizatoin. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .PARAMETER EntityLogicalName
 A logicalname for Entity. i.e.)accout, contact, lead, etc..

 .EXAMPLE
 Publish-CrmEntity -conn $conn -EntityLogicalName account
 True

 This example publishes customization for Account Entity.

 .EXAMPLE
 Publish-CrmEntity lead
 True

 This example publishes customization for Lead Entity by ommiting parameter names.
 When ommiting parameter names, you do not provide $conn, cmdlets automatically finds it.

#>

 [CmdletBinding()]
    PARAM( 
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1)]
        [string]$EntityLogicalName
    )

    if($conn -eq $null)
    {
        $connobj = Get-Variable conn -Scope global -ErrorAction SilentlyContinue
        if($connobj.Value -eq $null)
        {
            Write-Warning 'You need to create Connect to CRM Organization. Use Get-CrmConnection to create it.'
            break;
        }
        else
        {
            $conn = $connobj.Value
        }
    }  

    try
    {
        $result = $conn.PublishEntity($EntityLogicalName)
    }
    catch
    {
        return $conn.LastCrmException
    }

    return $result
}

#ResetLocalMetadataCache  
function Remove-CrmEntityMetadataCache{

<#
 .SYNOPSIS
 Removes metadata cache for an Entity.

 .DESCRIPTION
 The Remove-CrmEntityMetadataCache cmdlet lets you remove metadata cache for an Entity.

 .PARAMETER conn
 A connection to your CRM organizatoin. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .PARAMETER EntityLogicalName
 A logicalname for Entity. i.e.)accout, contact, lead, etc..

 .EXAMPLE
 Remove-CrmEntityMetadataCache -conn $conn -EntityLogicalName account
 True

 This example removes metadata cache for Account Entity.

 .EXAMPLE
 Remove-CrmEntityMetadataCache lead
 True

 This example removes metadata cache for Lead Entity by ommiting parameter names.
 When ommiting parameter names, you do not provide $conn, cmdlets automatically finds it.

#>

 [CmdletBinding()]
    PARAM( 
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$false, Position=1)]
        [string]$EntityLogicalName
    )

    if($conn -eq $null)
    {
        $connobj = Get-Variable conn -Scope global -ErrorAction SilentlyContinue
        if($connobj.Value -eq $null)
        {
            Write-Warning 'You need to create Connect to CRM Organization. Use Get-CrmConnection to create it.'
            break;
        }
        else
        {
            $conn = $connobj.Value
        }
    }  

    if($EntityLogicalName -eq "")
    {
        $EntityLogicalName = $null
    }
    
    try
    {
        $conn.ResetLocalMetadataCache($EntityLogicalName)
    }
    catch
    {
        return $conn.LastCrmException
    }    
}

#UninstallSampleDataFromCrm     
function Remove-CrmSampleData{

<#
 .SYNOPSIS
 Removes sample data CRM Organization.

 .DESCRIPTION
 The Remove-CrmSampleData cmdlet lets you remove sample data to CRM Organization and returns Job Id. You can confirm the status by using Test-CrmSampleDataInstalled cmdlet.

 .PARAMETER conn
 A connection to your CRM organizatoin. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .EXAMPLE
 Remove-CrmSampleData -conn $conn
 Guid
 ----
 9f7ff9ec-d2b5-4599-9909-04a072e0a546 

 This example removes sample data.

 .EXAMPLE
 Add-CrmSampleData
 Guid
 ----
 9f7ff9ec-d2b5-4599-9909-04a072e0a546 

 This example removes sample data by ommiting parameter names.
 When ommiting parameter names, you do not provide $conn, cmdlets automatically finds it.

#>

    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn
    )

    if($conn -eq $null)
    {
        $connobj = Get-Variable conn -Scope global -ErrorAction SilentlyContinue
        if($connobj.Value -eq $null)
        {
            Write-Warning 'You need to create Connect to CRM Organization. Use Get-CrmConnection to create it.'
            break;
        }
        else
        {
            $conn = $connobj.Value
        }
    } 

    try
    {
        $result = $conn.UninstallSampleDataFromCrm()
    }
    catch
    {
        return $conn.LastCrmException
    }

    return $result
}

#UpdateStateAndStatusForEntity 
function Set-CrmRecordState{

<#
 .SYNOPSIS
 Sets Status/State for a CRM record.

 .DESCRIPTION
 The Set-CrmRecordState cmdlet lets you set Status/State for a CRM record. 
 
 There are two ways to specify a record. 
 1. Pass EntityLogicalName and record's Id.
 2. Get a record object by using Get-CrmRecord/Get-CrmRecords cmdlets, then pass it.

 .PARAMETER conn
 A connection to your CRM organizatoin. Use $conn = Get-CrmConnection <Parameters> to generate it.

  .PARAMETER CrmRecord
 A record object which is obtained via Get-CrmRecord/Get-CrmRecords. When you pass CrmRecord, then you don't use EntityLogicalName/Id.

 .PARAMETER EntityLogicalName
 A logicalname for an Entity to set an owner. i.e.)accout, contact, lead, etc..

 .PARAMETER Id
 An Id (guid) of the record

 .PARAMETER StateCode
 StateCode value for a record. You can retrieve values by using Get-CrmEntityOptionSet <EntityLogicalName> statecode | % {$_.Items}

 .PARAMETER StatusCode
 StatusCode value for a record. You can retrieve values by using Get-CrmEntityOptionSet <EntityLogicalName> statuscode | % {$_.Items}

 .EXAMPLE
 Set-CrmRecordState -conn $conn -EntityLogicalName account -Id 1bf8d93d-1f18-e511-80da-c4346bc43d94 -StateCode Inactive -StatusCode Inactive
 
 This example sets disabled state for an account record.

 .EXAMPLE
 Set-CrmRecordState account 1bf8d93d-1f18-e511-80da-c4346bc43d94 Inactive Inactive
 
 This example sets disabled state for an account record by ommiting parameter names.
 When ommiting parameter names, you do not provide $conn, cmdlets automatically finds it.

 .EXAMPLE 
 PS C:\>$contact = Get-CrmRecord contact 81f8d93d-1f18-e511-80da-c4346bc43d94 fullname
 
 PS C:\>Set-CrmRecordState $contact Inactive Inactive

 This example retrieves and store a contact record, then pass it to Set-CrmRecordState to disable it.
#>

    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1, ValueFromPipeline=$true, ParameterSetName="CrmRecord")]
        [PSObject]$CrmRecord,
        [parameter(Mandatory=$true, Position=1, ParameterSetName="NameWithId")]
        [string]$EntityLogicalName,
        [parameter(Mandatory=$true, Position=2, ParameterSetName="NameWithId")]
        [guid]$Id,
        [parameter(Mandatory=$true, Position=3)]
        [string]$StateCode,
        [parameter(Mandatory=$true, Position=4)]
        [string]$StatusCode
    )

    if($conn -eq $null)
    {
        $connobj = Get-Variable conn -Scope global -ErrorAction SilentlyContinue
        if($connobj.Value -eq $null)
        {
            Write-Warning 'You need to create Connect to CRM Organization. Use Get-CrmConnection to create it.'
            break;
        }
        else
        {
            $conn = $connobj.Value
        }
    }    

    if($CrmRecord -ne $null)
    {
        $EntityLogicalName = $CrmRecord.logicalname
        $Id = $CrmRecord.($EntityLogicalName + "id")
    }

    try
    {
        $result = $conn.UpdateStateAndStatusForEntity($EntityLogicalName, $Id, $stateCode, $statusCode, [Guid]::Empty)
    }
    catch
    {
        return $conn.LastCrmException
    }
}

### Other Cmdlets added by Dynamics CRM PFE ###

function Add-CrmSecurityRoleToTeam{

<#
 .SYNOPSIS
 Assigns a security role to a team.

 .DESCRIPTION
 The Add-CrmSecurityRoleToTeam cmdlet lets you assign a security role to a team. 

 There are two ways to specify records.
 
 1. Pass record's Id for both records.
 2. Get a record object by using Get-CrmRecord/Get-CrmRecords cmdlets, and pass it for both records.

 .PARAMETER conn
 A connection to your CRM organizatoin. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .PARAMETER TeamRecord
 A team record object which is obtained via Get-CrmRecord/Get-CrmRecords. When you pass CrmRecord, then you don't use UserId.

 .PARAMETER SecurityRoleRecord
 A security role record object which is obtained via Get-CrmRecord/Get-CrmRecords. When you pass CrmRecord, then you don't use SecurityRoleId.

 .PARAMETER TeamId
 An Id (guid) of team record

 .PARAMETER SecurityRoleId
 An Id (guid) of security role record

 .PARAMETER SecurityRoleName
 A name of security role record

 .EXAMPLE
 Add-CrmSecurityRoleToTeam -conn $conn -TeamId 00005a70-6317-e511-80da-c4346bc43d94 -SecurityRoleId 66005a70-6317-e511-80da-c4346bc43d94

 This example assigns the security role to the team by using Id.

 .EXAMPLE
 Add-CrmSecurityRoleToTeam 00005a70-6317-e511-80da-c4346bc43d94 66005a70-6317-e511-80da-c4346bc43d94
 
 This example assigns the security role to the team by using Id by ommiting parameters names.
 When ommiting parameter names, you do not provide $conn, cmdlets automatically finds it.

 .EXAMPLE
 PS C:\>$team = Get-CrmRecord team 00005a70-6317-e511-80da-c4346bc43d94 name
 PS C:\>$role = Get-CrmRecord role 66005a70-6317-e511-80da-c4346bc43d94 name
 PS C:\>Add-CrmSecurityRoleToTeam $team $role

 This example assigns the security role to the team by using record objects.

 .EXAMPLE
 Add-CrmSecurityRoleToUser -conn $conn -TeamId 00005a70-6317-e511-80da-c4346bc43d94 -SecurityRoleName "salesperson"
 
 This example assigns the salesperson role to the team by using Id and role name.
#>

    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1, ParameterSetName="CrmRecord")]
        [PSObject]$TeamRecord,
        [parameter(Mandatory=$false, Position=2, ParameterSetName="CrmRecord")]
        [PSObject]$SecurityRoleRecord,
        [parameter(Mandatory=$true, Position=1, ParameterSetName="Id")]
        [string]$TeamId,
        [parameter(Mandatory=$false, Position=2, ParameterSetName="Id")]
        [string]$SecurityRoleId,
        [parameter(Mandatory=$false, Position=2)]
        [string]$SecurityRoleName
    )

    if($conn -eq $null)
    {
        $connobj = Get-Variable conn -Scope global -ErrorAction SilentlyContinue
        if($connobj.Value -eq $null)
        {
            Write-Warning 'You need to create Connect to CRM Organization. Use Get-CrmConnection to create it.'
            break;
        }
        else
        {
            $conn = $connobj.Value
        }
    }

    if($SecurityRoleRecord -eq $null -and $SecurityRoleId -eq "" -and $SecurityRoleName -eq "")
    {
        Write-Warning "You need to specify Security Role information"
        return
    }
    
    if($SecurityRoleName -ne "")
    {
        if($TeamRecord -eq $null -or $TeamRecord.businessunitid -eq $null)
        {
            $TeamRecord = Get-CrmRecord -conn $conn -EntityLogicalName team -Id $TeamId -Fields businessunitid
        }

        $fetch = @"
        <fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="false" no-lock="true">
          <entity name="role">
            <attribute name="businessunitid" />
            <attribute name="roleid" />
            <filter type="and">
              <condition attribute="name" operator="eq" value="{0}" />
              <condition attribute="businessunitid" operator="eq" value="{1}" />
            </filter>
          </entity>
        </fetch>
"@ -F $SecurityRoleName, $TeamRecord.businessunitid_Property.Value.Id
        
        $roles = (Get-CrmRecordsByFetch -conn $conn -Fetch $fetch)
        if($roles.CrmRecords.Count -eq 0)
        {
            Write-Warning "Not Security Role found"
            return
        }
        else
        {
            $role = $roles.CrmRecords[0]
        }
    }

    if($SecurityRoleName -ne "")
    {
        Add-CrmRecordAssociation -conn $conn -CrmRecord1 $TeamRecord -CrmRecord2 $role -RelationshipName teamroles_association
    }
    elseif($TeamRecord -ne $null)
    {
        Add-CrmRecordAssociation -conn $conn -CrmRecord1 $TeamRecord -CrmRecord2 $SecurityRoleRecord -RelationshipName teamroles_association
    }
    else
    {
        Add-CrmRecordAssociation -conn $conn -EntityLogicalName1 team -Id1 $TeamId -EntityLogicalName2 role -Id2 $SecurityRoleId -RelationshipName teamroles_association
    }
}

function Add-CrmSecurityRoleToUser{

<#
 .SYNOPSIS
 Assigns a security role to a user.

 .DESCRIPTION
 The Add-CrmSecurityRoleToUser cmdlet lets you assign a security role to a user. 

 There are two ways to specify records.
 
 1. Pass record's Id for both records.
 2. Get a record object by using Get-CrmRecord/Get-CrmRecords cmdlets, and pass it for both records.

 .PARAMETER conn
 A connection to your CRM organizatoin. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .PARAMETER UserRecord
 A user record object which is obtained via Get-CrmRecord/Get-CrmRecords. When you pass CrmRecord, then you don't use UserId.

 .PARAMETER SecurityRoleRecord
 A security role record object which is obtained via Get-CrmRecord/Get-CrmRecords. When you pass CrmRecord, then you don't use SecurityRoleId.

 .PARAMETER UserId
 An Id (guid) of user record

 .PARAMETER SecurityRoleId
 An Id (guid) of security role record
 
 .PARAMETER SecurityRoleName
 A name of security role record

 .EXAMPLE
 Add-CrmSecurityRoleToUser -conn $conn -UserId 00005a70-6317-e511-80da-c4346bc43d94 -SecurityRoleId 66005a70-6317-e511-80da-c4346bc43d94

 This example assigns the security role to the user by using Id.

 .EXAMPLE
 Add-CrmSecurityRoleToUser 00005a70-6317-e511-80da-c4346bc43d94 66005a70-6317-e511-80da-c4346bc43d94
 
 This example assigns the security role to the user by using Id by ommiting parameters names.
 When ommiting parameter names, you do not provide $conn, cmdlets automatically finds it.

 .EXAMPLE
 PS C:\>$user = Get-CrmRecord sysetmuser 00005a70-6317-e511-80da-c4346bc43d94 fullname

 PS C:\>$role = Get-CrmRecord role 66005a70-6317-e511-80da-c4346bc43d94 name

 PS C:\>Add-CrmSecurityRoleToUser $user $role

 This example assigns the security role to the user by using record objects.

 .EXAMPLE
 Add-CrmSecurityRoleToUser -conn $conn -UserId 00005a70-6317-e511-80da-c4346bc43d94 -SecurityRoleName "salesperson"
 
 This example assigns the salesperson role to the user by using Id and role name.
#>

    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1, ParameterSetName="CrmRecord")]
        [PSObject]$UserRecord,
        [parameter(Mandatory=$false, Position=2, ParameterSetName="CrmRecord")]
        [PSObject]$SecurityRoleRecord,
        [parameter(Mandatory=$true, Position=1, ParameterSetName="Id")]
        [string]$UserId,
        [parameter(Mandatory=$false, Position=2, ParameterSetName="Id")]
        [string]$SecurityRoleId,
        [parameter(Mandatory=$false, Position=2)]
        [string]$SecurityRoleName
    )

    if($conn -eq $null)
    {
        $connobj = Get-Variable conn -Scope global -ErrorAction SilentlyContinue
        if($connobj.Value -eq $null)
        {
            Write-Warning 'You need to create Connect to CRM Organization. Use Get-CrmConnection to create it.'
            break;
        }
        else
        {
            $conn = $connobj.Value
        }
    }

    if($SecurityRoleRecord -eq $null -and $SecurityRoleId -eq "" -and $SecurityRoleName -eq "")
    {
        Write-Warning "You need to specify Security Role information"
        return
    }
    
    if($SecurityRoleName -ne "")
    {
        if($UserRecord -eq $null -or $UserRecord.businessunitid -eq $null)
        {
            $UserRecord = Get-CrmRecord -conn $conn -EntityLogicalName team -Id $UserId -Fields businessunitid
        }

        $fetch = @"
        <fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="false" no-lock="true">
          <entity name="role">
            <attribute name="businessunitid" />
            <attribute name="roleid" />
            <filter type="and">
              <condition attribute="name" operator="eq" value="{0}" />
              <condition attribute="businessunitid" operator="eq" value="{1}" />
            </filter>
          </entity>
        </fetch>
"@ -F $SecurityRoleName, $UserRecord.businessunitid_Property.Value.Id
        
        $roles = (Get-CrmRecordsByFetch -conn $conn -Fetch $fetch)
        if($roles.CrmRecords.Count -eq 0)
        {
            Write-Warning "Not Security Role found"
            return
        }
        else
        {
            $role = $roles.CrmRecords[0]
        }
    }

    if($SecurityRoleName -ne "")
    {
        Add-CrmRecordAssociation -conn $conn -CrmRecord1 $UserRecord -CrmRecord2 $role -RelationshipName systemuserroles_association
    }
    elseif($UserRecord -ne $null)
    {
        Add-CrmRecordAssociation -conn $conn -CrmRecord1 $UserRecord -CrmRecord2 $SecurityRoleRecord -RelationshipName systemuserroles_association
    }
    else
    {
        Add-CrmRecordAssociation -conn $conn -EntityLogicalName1 systemuser -Id1 $UserId -EntityLogicalName2 role -Id2 $SecurityRoleId -RelationshipName systemuserroles_association
    }
}

function Approve-CrmEmailAddress{

<#
 .SYNOPSIS
 Approve email address change of a user or a queue.

 .DESCRIPTION
 The Approve-CrmEmailAddress cmdlet lets you approves email address change of a user or a queue. 

 .PARAMETER conn
 A connection to your CRM organizatoin. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .PARAMETER UserId
 A record Id of User.

 .PARAMETER QueueId
 A record Id of Queue.
  
 .EXAMPLE
 Approve-CrmEmailAddress -conn $conn -UserId 00005a70-6317-e511-80da-c4346bc43d94

 This example approves email address for a user.

 .EXAMPLE
 Approve-CrmEmailAddress -UserId 00005a70-6317-e511-80da-c4346bc43d94
 
 This example approves email address for a user by ommiting parameters names.
 When ommiting parameter names, you do not provide $conn, cmdlets automatically finds it.

 .EXAMPLE
 Approve-CrmEmailAddress -conn $conn -QueueId 00005a70-6317-e511-80da-c4346bc43d94

 This example approves email address for a queue.

#>

    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1, ParameterSetName="UserId")]
        [string]$UserId,
        [parameter(Mandatory=$true, Position=1, ParameterSetName="QueueId")]
        [string]$QueueId
    )

    if($conn -eq $null)
    {
        $connobj = Get-Variable conn -Scope global -ErrorAction SilentlyContinue
        if($connobj.Value -eq $null)
        {
            Write-Warning 'You need to create Connect to CRM Organization. Use Get-CrmConnection to create it.'
            break;
        }
        else
        {
            $conn = $connobj.Value
        }
    }

    if($UserId -ne "")
    {
        Set-CrmRecord systemuser $UserId @{"emailrouteraccessapproval"=(New-CrmOptionSetValue 1)}
    }
    else
    {
        Set-CrmRecord queue $QueueId @{"emailrouteraccessapproval"=(New-CrmOptionSetValue 1)}
    }
}

function Disable-CrmLanguagePack{

<#
 .SYNOPSIS
 Executes DeprovisionLanguageRequest Organization Request.

 .DESCRIPTION
 The Disable-CrmLanguagePack cmdlet lets you deprovision LanguagePack.

 .PARAMETER conn
 A connection to your CRM organizatoin. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .PARAMETER LCID
 A Language ID.

 .EXAMPLE
 Disable-CrmLanguagePack -conn $conn -LCID 1041

 This example deprovisions Japanese Language Pack.

 .EXAMPLE
 Disable-CrmLanguagePack 1041
 
 This example deprovisions Japanese Language Pack by ommiting parameters names.
 When ommiting parameter names, you do not provide $conn, cmdlets automatically finds it.

#>

	[CmdletBinding()]
    PARAM( 
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1)]
        [Int]$LCID
    )

    if($conn -eq $null)
    {
        $connobj = Get-Variable conn -Scope global -ErrorAction SilentlyContinue
        if($connobj.Value -eq $null)
        {
            Write-Warning 'You need to create Connect to CRM Organization. Use Get-CrmConnection to create it.'
            break;
        }
        else
        {
            $conn = $connobj.Value
        }
    }  

    $request = New-Object Microsoft.Crm.Sdk.Messages.DeprovisionLanguageRequest
    $request.Language = $LCID
    
    try
    {
        $response = $conn.ExecuteCrmOrganizationRequest($request, $null)
    }
    catch
    {
        return $conn.LastCrmException
    }    
}

function Enable-CrmLanguagePack{

<#
 .SYNOPSIS
 Executes ProvisionLanguageRequest Organization Request.

 .DESCRIPTION
 The Enable-CrmLanguagePack cmdlet lets you provision LanguagePack. For OnPremise, you need to install corresponding Language Pack inadvance.

 .PARAMETER conn
 A connection to your CRM organizatoin. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .PARAMETER LCID
 A Language ID.

 .EXAMPLE
 Enable-CrmLanguagePack -conn $conn -LCID 1041

 This example provisions Japanese Language Pack.

 .EXAMPLE
 Enable-CrmLanguagePack 1041
 
 This example provisions Japanese Language Pack by ommiting parameters names.
 When ommiting parameter names, you do not provide $conn, cmdlets automatically finds it.

#>

	[CmdletBinding()]
    PARAM( 
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1)]
        [Int]$LCID
    )

    if($conn -eq $null)
    {
        $connobj = Get-Variable conn -Scope global -ErrorAction SilentlyContinue
        if($connobj.Value -eq $null)
        {
            Write-Warning 'You need to create Connect to CRM Organization. Use Get-CrmConnection to create it.'
            break;
        }
        else
        {
            $conn = $connobj.Value
        }
    }  

    $request = New-Object Microsoft.Crm.Sdk.Messages.ProvisionLanguageRequest
    $request.Language = $LCID
    
    try
    {
        $response = $conn.ExecuteCrmOrganizationRequest($request, $null)
    }
    catch
    {
        return $conn.LastCrmException
    }    
}

function Export-CrmSolution{

<#
 .SYNOPSIS
 Exports a solution by Name from a CRM Organization.

 .DESCRIPTION
 The Export-CrmSolution cmdlet lets you export a solution file from CRM.

 .PARAMETER conn
 A connection to your CRM organizatoin. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .PARAMETER SolutionName
 An unique name of the exporting solution.

 .PARAMETER SolutionFilePath
 A path to save exporting solution.

 .PARAMETER SolutionZipFileName
 A file name of exporting solution zip file.

 .PARAMETER Managed
 Specify the parameter to export the solution as managed. if you don't give this parameter, the solution will be exported as unmanaged.

 .PARAMETER TargetVersion
 Specify TargetVersion of exporting solution.

 .PARAMETER ExportAutoNumberingSettings
 Specify the parameter to export auto numbering settings.

 .PARAMETER ExportCalendarSettings
 Specify the parameter to export calendar settings.

 .PARAMETER ExportCustomizationSettings
 Specify the parameter to export customization settings.

 .PARAMETER ExportEmailTrackingSettings
 Specify the parameter to export email tracking settings.

 .PARAMETER ExportGeneralSettings
 Specify the parameter to export general settings.

 .PARAMETER ExportMarketingSettings
 Specify the parameter to export marketing settings.

 .PARAMETER ExportOutlookSynchronizationSettings
 Specify the parameter to export outlook synchronization settings.

 .PARAMETER ExportRelationshipRoles
 Specify the parameter to export relationship roles.

 .PARAMETER ExportIsvConfig
 Specify the parameter to export ISV config.

 .PARAMETER ExportSales
 Specify the parameter to export sales settings.

 .EXAMPLE
 Export-CrmSolution -conn $conn -SolutionName "MySolution"
 ExportSolutionResponse                            SolutionPath                                              
 ----------------------                            ------------                                              
 Microsoft.Crm.Sdk.Messages.ExportSolutionResponse C:\Users\xxxx\Desktop\MySolution_unmanaged_1.0.0.0.zip

 This example exports "MySolution" solution as unmanaged with current path and default name.

 .EXAMPLE
 ExportSolutionResponse                            SolutionPath                                              
 ----------------------                            ------------                                              
 Microsoft.Crm.Sdk.Messages.ExportSolutionResponse C:\Users\xxxx\Desktop\MySolution_unmanaged_1.0.0.0.zip

 This example exports "MySolution" solution as unmanaged with current path and default name by ommiting $conn parameter.
 When ommiting $conn parameter, cmdlets automatically finds it.

 .EXAMPLE
 Export-CrmSolution -conn $conn -SolutionName "MySolution" -Managed -SolutionFilePath "C:\temp" -SolutionZipFileName "MySolution_Managed.zip" 
 ExportSolutionResponse                            SolutionPath                                              
 ----------------------                            ------------                                              
 Microsoft.Crm.Sdk.Messages.ExportSolutionResponse C:\temp\MySolution_Managed.zip

 This example exports "MySolution" solution as managed with specified path and name.
#>

    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1)]
        [string]$SolutionName, 
        [parameter(Mandatory=$false, Position=2)]
        [string]$SolutionFilePath,
        [parameter(Mandatory=$false)]
        [string]$SolutionZipFileName,
        [parameter(Mandatory=$false)]
        [switch]$Managed,
        [parameter(Mandatory=$false)]
        [string]$TargetVersion, 
        [parameter(Mandatory=$false)]
        [switch]$ExportAutoNumberingSettings, 
        [parameter(Mandatory=$false)]
        [switch]$ExportCalendarSettings, 
        [parameter(Mandatory=$false)]
        [switch]$ExportCustomizationSettings, 
        [parameter(Mandatory=$false)]
        [switch]$ExportEmailTrackingSettings, 
        [parameter(Mandatory=$false)]
        [switch]$ExportGeneralSettings, 
        [parameter(Mandatory=$false)]
        [switch]$ExportMarketingSettings, 
        [parameter(Mandatory=$false)]
        [switch]$ExportOutlookSynchronizationSettings, 
        [parameter(Mandatory=$false)]
        [switch]$ExportRelationshipRoles, 
        [parameter(Mandatory=$false)]
        [switch]$ExportIsvConfig, 
        [parameter(Mandatory=$false)]
        [switch]$ExportSales
    )    

    if($conn -eq $null)
    {
        $connobj = Get-Variable conn -Scope global -ErrorAction SilentlyContinue
        if($connobj.Value -eq $null)
        {
            Write-Error 'You need to create a connection to a CRM Organization using get-CrmConnection or pass the connection as a parameter to use this cmdlet.'
            break;
        }
        else
        {
            $conn = $connobj.Value
        }
    }   
    try
    {
        $solutionRecords = (Get-CrmRecords -conn $conn -EntityLogicalName solution -FilterAttribute uniquename -FilterOperator "like" -FilterValue $SolutionName -Fields uniquename,publisherid,version )

        #if we can't find just one solution matching then ERROR
        if($solutionRecords.CrmRecords.Count -ne 1)
        {
            $friendlyName = $conn.ConnectedOrgFriendlyName.ToString(); 

            Write-Error "Solution with name `"$SolutionName`" in CRM Instance: `"$friendlyName`" not found!"
            break; 
        }
        #else PROCEED 
		$crmSolutionRecord = $solutionRecords.CrmRecords[0]; 

        $version = $crmSolutionRecord.version;
		$solutionUniqueName = $crmSolutionRecord.uniquename;

        write-verbose "Solution found with version# $version"
        
        $exportPath = if($SolutionFilePath -ne ""){Get-Item $SolutionFilePath} else {Get-Location}
        
        #if a filename is not given, then we'll default one to [solutionname]_[managed]_[version].zip
        if($SolutionZipFileName.Length -eq 0)
        {
            $version = $version.Replace('.','_')
            $managedFileName = if($Managed) {"_managed_"} else {"_unmanaged_"}
            $solutionZipFileName = "$solutionUniqueName$managedFileName$version.zip"; 
        }
        #now we should have the final path
        $path = Join-Path $exportPath $solutionZipFileName

        Write-Verbose "Solution path: $path"

        #create the export request then set all the properties
        $exportRequest = New-Object Microsoft.Crm.Sdk.Messages.ExportSolutionRequest; 
        $exportRequest.ExportAutoNumberingSettings            =$ExportAutoNumberingSettings 
        $exportRequest.ExportCalendarSettings                 =$ExportCalendarSettings
        $exportRequest.ExportCustomizationSettings            =$ExportCustomizationSettings
        $exportRequest.ExportEmailTrackingSettings            =$ExportEmailTrackingSettings
        $exportRequest.ExportGeneralSettings                  =$ExportGeneralSettings
        $exportRequest.ExportIsvConfig                        =$ExportIsvConfig
        $exportRequest.ExportMarketingSettings                =$ExportMarketingSettings
        $exportRequest.ExportOutlookSynchronizationSettings   =$ExportOutlookSynchronizationSettings
        $exportRequest.ExportRelationshipRoles                =$ExportRelationshipRoles
        $exportRequest.ExportSales                            =$ExportSales
        $exportRequest.Managed                                =$Managed
        $exportRequest.SolutionName                           =$solutionUniqueName
        $exportRequest.TargetVersion                          =$TargetVersion 

        Write-Verbose 'ExportSolutionRequests may take several minutes to complete execution.'
        
        $response = [Microsoft.Crm.Sdk.Messages.ExportSolutionResponse]($conn.ExecuteCrmOrganizationRequest($exportRequest)); 

		Write-Verbose 'Using solution file to path: $path'

        [System.IO.File]::WriteAllBytes($path,$response.ExportSolutionFile);

        Write-Verbose "Successfully wrote file"
        $result = New-Object psObject

        Add-Member -InputObject $result -MemberType NoteProperty -Name "ExportSolutionResponse" -Value $response
        Add-Member -InputObject $result -MemberType NoteProperty -Name "SolutionPath" -Value $path

        return $result; 
    }
    catch
    {
        Write-Error $_.Exception
        #return $conn.LastCrmException
    }

    return $result
}

function Export-CrmSolutionTranslation{

<#
 .SYNOPSIS
 Exports a translation from a solution.

 .DESCRIPTION
 The Export-CrmSolutionTranslation cmdlet lets you export a translation from a solution.

 .PARAMETER conn
 A connection to your CRM organizatoin. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .PARAMETER SolutionName
 An unique name of the exporting solution.

 .PARAMETER TranslationFilePath
 A path to save exporting solution translation.

 .PARAMETER TranslationFileName
 A file name of exporting solution translation zip file.
 
 .EXAMPLE
 Export-CrmSolutionTranslation -conn $conn -SolutionName "MySolution"
 ExportTranslationResponse                            SolutionTranslationPath                                         
 -------------------------                            -----------------------                                         
 Microsoft.Crm.Sdk.Messages.ExportTranslationResponse C:\Users\xxx\Desktop\CrmTranslations_MySolution_1_0_0_0.zip

 This example exports translation file of "MySolution" solution with current path and default name.

 .EXAMPLE
 Export-CrmSolutionTranslation -SolutionName "MySolution"
 ExportTranslationResponse                            SolutionTranslationPath                                         
 -------------------------                            -----------------------                                         
 Microsoft.Crm.Sdk.Messages.ExportTranslationResponse C:\Users\xxx\Desktop\CrmTranslations_MySolution_1_0_0_0.zip

 This example exports translation file of "MySolution" solution with current path and default name by ommiting $conn parameter.
 When ommiting $conn parameter, cmdlets automatically finds it.

 .EXAMPLE
 Export-CrmSolutionTranslation -conn $conn -SolutionName "MySolution" -TranslationFilePath "C:\temp" -TranslationZipFileName "CrmTranslations_MySolution.zip"
 ExportTranslationResponse                            SolutionTranslationPath                                         
 -------------------------                            -----------------------                                         
 Microsoft.Crm.Sdk.Messages.ExportTranslationResponse C:\temp\CrmTranslations_MySolution.zip

 This example exports translation file of "MySolution" solution with specified path and name.
#>

    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1)]
        [string]$SolutionName, 
        [parameter(Mandatory=$false)]
        [string]$TranslationFilePath,
        [parameter(Mandatory=$false)]
        [string]$TranslationZipFileName
    )    

    if($conn -eq $null)
    {
        $connobj = Get-Variable conn -Scope global -ErrorAction SilentlyContinue
        if($connobj.Value -eq $null)
        {
            Write-Error 'You need to create a connection to a CRM Organization using get-CrmConnection or pass the connection as a parameter to use this cmdlet.'
            break;
        }
        else
        {
            $conn = $connobj.Value
        }
    }   
    try
    {
        $solutionRecords = (Get-CrmRecords -conn $conn -EntityLogicalName solution -FilterAttribute uniquename -FilterOperator "like" -FilterValue $SolutionName -Fields publisherid,version )

        #if we can't find just one solution matching then ERROR
        if($solutionRecords.CrmRecords.Count -ne 1)
        {
            $friendlyName = $conn.ConnectedOrgFriendlyName.ToString(); 

            Write-Error "Solution with name `"$SolutionName`" in CRM Instance: `"$friendlyName`" not found!"
            break; 
        }
        #else PROCEED 

        $version = $solutionRecords.CrmRecords[0].version 

        write-verbose "Solution found with version# $version"
       
        $exportPath = if($TranslationFilePath -ne ""){ $TranslationFilePath } else { Get-Location }
        
        #if a filename is not given, then we'll default one to CrmTranslations_[solutionname]_[version].zip
        if($TranslationZipFileName.Length -eq 0)
        {
            $version = $version.Replace('.','_')
            $translationZipFileName = "CrmTranslations_$SolutionName`_$version.zip"; 
        }

        #now we should have the final path
        $path = Join-Path $exportPath $translationZipFileName

        Write-Verbose "Solution path: $path"

        #create the export translation request then set all the properties
        $exportRequest = New-Object Microsoft.Crm.Sdk.Messages.ExportTranslationRequest; 
        $exportRequest.SolutionName = $SolutionName

        Write-Verbose 'ExportTranslationRequest may take several minutes to complete execution.'
        
        $response = [Microsoft.Crm.Sdk.Messages.ExportTranslationResponse]($conn.ExecuteCrmOrganizationRequest($exportRequest)); 

        [System.IO.File]::WriteAllBytes($path,$response.ExportTranslationFile);

        Write-Verbose "Successfully wrote file: $path"
        $result = New-Object psObject

        Add-Member -InputObject $result -MemberType NoteProperty -Name "ExportTranslationResponse" -Value $response
        Add-Member -InputObject $result -MemberType NoteProperty -Name "SolutionTranslationPath" -Value $path

        return $result; 
    }
    catch
    {
        Write-Error $_.Exception
        #return $conn.LastCrmException
    }

    return $result
}

function Get-CrmEntityRecordCount{

<#
 .SYNOPSIS
 Retrieves total record count for an Entity.

 .DESCRIPTION
 The Get-CrmEntityRecordCount cmdlet lets you retrieve total record count for an Entity.

 .PARAMETER conn
 A connection to your CRM organizatoin. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .PARAMETER EntityLogicalName
 A logicalname for an Entity to retrieve. i.e.)accout, contact, lead, etc..

 .EXAMPLE
 Get-CrmEntityRecordCount -conn $conn -EntityLogicalName account
 10

 This example retrieves total record count for Account Entity.

 .EXAMPLE
 Get-CrmEntityRecordCount contact
 8413
 
 This example retrieves total record count for Contact Entity by omitting parameter names.
 When ommiting parameter names, you do not provide $conn, cmdlets automatically finds it.

#>

    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1)]
        [string]$EntityLogicalName
    )

    if($conn -eq $null)
    {
        $connobj = Get-Variable conn -Scope global -ErrorAction SilentlyContinue
        if($connobj.Value -eq $null)
        {
            Write-Warning 'You need to create Connect to CRM Organization. Use Get-CrmConnection to create it.'
            break;
        }
        else
        {
            $conn = $connobj.Value
        }
    }
    
    $count = 0
    $query = New-Object -TypeName 'Microsoft.Xrm.Sdk.Query.QueryExpression'
    $pageInfo = New-Object -TypeName 'Microsoft.Xrm.Sdk.Query.PagingInfo'
    $query.EntityName = $EntityLogicalName
    $pageInfo.Count = 5000
    $pageInfo.PageNumber = 1
    $pageInfo.PagingCookie = $null
    $query.PageInfo = $pageInfo
    
    while($True)
    {
        $request = New-Object -TypeName 'Microsoft.Xrm.Sdk.Messages.RetrieveMultipleRequest'
        $request.Query = $query
        try
        {
            $result = $conn.ExecuteCrmOrganizationRequest($request)
        }
        catch
        {
            return $conn.LastCrmException
        }
        
        $count += $result.EntityCollection.Entities.Count
        if($result.EntityCollection.MoreRecords)
        {
            $pageInfo.PageNumber += 1
            $pageInfo.PagingCookie = $result.EntityCollection.PagingCookie
        }
        else
        {
            break;
        } 
    }
    
    return $count
}

function Get-CrmAllLanguagePacks{

<#
 .SYNOPSIS
 Executes RetrieveAvailableLanguagesRequest Organization Request.

 .DESCRIPTION
 The Get-CrmAllLanguagePacks cmdlet lets you retrieve all available LanguagePack.

 .PARAMETER conn
 A connection to your CRM organizatoin. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .EXAMPLE
 PS C:\>Get-CrmAllLanguagePacks -conn $conn
 1041
 1033
 2052
 ...

 This example retrieves all available Language Pack.

 .EXAMPLE
 PS C:\>Get-CrmAllLanguagePacks
 1041
 1033
 2052
 ...

 This example etrieves all available Language Pack by ommiting parameters names.
 When ommiting parameter names, you do not provide $conn, cmdlets automatically finds it.

#>

	[CmdletBinding()]
    PARAM( 
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn
    )

    if($conn -eq $null)
    {
        $connobj = Get-Variable conn -Scope global -ErrorAction SilentlyContinue
        if($connobj.Value -eq $null)
        {
            Write-Warning 'You need to create Connect to CRM Organization. Use Get-CrmConnection to create it.'
            break;
        }
        else
        {
            $conn = $connobj.Value
        }
    }  

    $request = New-Object Microsoft.Crm.Sdk.Messages.RetrieveAvailableLanguagesRequest

    try
    {
        $response = $conn.ExecuteCrmOrganizationRequest($request, $null)
    }
    catch
    {
        return $conn.LastCrmException
    }    

    return $response.LocaleIds
}

function Get-CrmLicenseSummary{

<#
 .SYNOPSIS
 Displays License assignment and AccessMode/CalType summery.

 .DESCRIPTION
 The Get-CrmLicenseSummery cmdlet lets you display License assignment and AccessMode/CalType summery.

 .PARAMETER conn
 A connection to your CRM organizatoin. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .EXAMPLE
 Get-CrmLicenseSummery -conn $conn
 IsLicensed:
                        Count Name
                        ----- ----                                                        
                            4 Yes
                            2 No
 AccessMode:
                            3 Read-Write
                            1 Non-interactive
                            2 Administrative
 CalType:
                            3 Read-Write
                            1 Non-interactive
                            2 Administrative

 This example displays License assignment and AccessMode/CalType summery.

 .EXAMPLE
 Get-CrmLicenseSummery
 IsLicensed:
                        Count Name
                        ----- ----                                                        
                            4 Yes
                            2 No
 AccessMode:
                            3 Read-Write
                            1 Non-interactive
                            2 Administrative
 CalType:
                            3 Read-Write
                            1 Non-interactive
                            2 Administrative

 This example displays License assignment and AccessMode/CalType summery by omitting parameter names.
 When ommiting parameter names, you do not provide $conn, cmdlets automatically finds it.

#>

    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn
    )

    if($conn -eq $null)
    {
        $connobj = Get-Variable conn -Scope global -ErrorAction SilentlyContinue
        if($connobj.Value -eq $null)
        {
            Write-Warning 'You need to create Connect to CRM Organization. Use Get-CrmConnection to create it.'
            break;
        }
        else
        {
            $conn = $connobj.Value
        }
    }

    $fetch = @"
    <fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="false" no-lock="true">
        <entity name="systemuser">
            <attribute name="islicensed" />
            <attribute name="accessmode" />
            <attribute name="caltype" />
            <filter type='and'>
                <condition attribute='accessmode' operator='ne' value='3' />
                <condition attribute='domainname' operator='ne' value='' />
            </filter>
        </entity>
    </fetch>
"@
    $users = Get-CrmRecordsByFetch -conn $conn -Fetch $fetch
    Write-Output 'IsLicensed:' ($users.CrmRecords | group islicensed | select count, name)
    Write-Output 'AccessMode:' ($users.CrmRecords | group accessmode | select count, name)
    Write-Output 'CalType:' ($users.CrmRecords | group accessmode | select count, name)
}

function Get-CrmOrgDbOrgSettings{

<#
 .SYNOPSIS
 Retrieves CrmOrgDbOrgSettings.

 .DESCRIPTION
 The Get-CrmOrgDbOrgSettings cmdlet lets you retrieve CrmOrgDbOrgSettings.

 .PARAMETER conn
 A connection to your CRM organizatoin. Use $conn = Get-CrmConnection <Parameters> to generate it.
  
 .EXAMPLE
 Get-CrmOrgDbOrgSettings -conn $conn
 ActivateAdditionalRefreshOfWorkflowConditions : false
 ActivityConvertDlgCampaignUnchecked           : true
 ClientUEIPDisabled                            : false
 CreateSPFoldersUsingNameandGuid               : true
 DisableSmartMatching                          : false

 This example retrieves CrmOrgDbOrgSettings.

 .EXAMPLE
 Get-CrmOrgDbOrgSettings contact
 Get-CrmOrgDbOrgSettings -conn $conn
 ActivateAdditionalRefreshOfWorkflowConditions : false
 ActivityConvertDlgCampaignUnchecked           : true
 ClientUEIPDisabled                            : false
 CreateSPFoldersUsingNameandGuid               : true
 DisableSmartMatching                          : false
 
 This example retrieves CrmOrgDbOrgSettings by omitting parameter names.
 When ommiting parameter names, you do not provide $conn, cmdlets automatically finds it.

#>

    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn
    )

    if($conn -eq $null)
    {
        $connobj = Get-Variable conn -Scope global -ErrorAction SilentlyContinue
        if($connobj.Value -eq $null)
        {
            Write-Warning 'You need to create Connect to CRM Organization. Use Get-CrmConnection to create it.'
            break;
        }
        else
        {
            $conn = $connobj.Value
        }
    }
    
    $fetch = @"
    <fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="false" no-lock="true">
      <entity name="organization">
        <attribute name="orgdborgsettings" />
      </entity>
    </fetch>
"@
    $result = Get-CrmRecordsByFetch -conn $conn -Fetch  $fetch
    $record = $result.CrmRecords[0]

    if($record.orgdborgsettings -eq $null)
    {
        Write-Warning 'No settings found.'
    }
    else
    {
        $xml = [xml]$record.orgdborgsettings
        return $xml.SelectSingleNode("/OrgSettings")
    }

}

function Get-CrmRecords{
<#
 .SYNOPSIS
 Retrieves CRM records by using single filter condition.

 .DESCRIPTION
 The Get-CrmRecords cmdlet lets you retrieve CRM records by using single filter condition. It my return more than a record.
 You can specify condition operator by using PowerShell operator like "eq", "ne", "lt", "like", etc.

 You can specify desired fields as fieldname1,fieldname2,fieldname3 syntax or * to retrieve all fields (not recommended for performance reason.)
 You can use Get-CrmEntityAttributes cmdlet to see all fields logicalname. The retrieved data can be passed to several cmdlets like Set-CrmRecord, Removed-CrmRecord to further process it.

 .PARAMETER conn
 A connection to your CRM organizatoin. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .PARAMETER EntityLogicalName
 A logicalname for an Entity to retrieve. i.e.)accout, contact, lead, etc..

 .PARAMETER FilterAttribute
 A field logical name for filtering.

 .PARAMETER FilterOperator
 An condition operator like "eq", "ne", "lt", "like", etc.

 .PARAMETER FilterValue
 A field value for filtering

 .PARAMETER Fields
 A List of field logicalnames. Use "fieldname1, fieldname2, fieldname3" syntax to speficy Fields, or ues "*" to retrieve all fields (not recommended for performance reason.)
 
 .PARAMETER AllRows
 By default the first 5000 rows are returned, this switch will bring back all results regardless of how many

 .EXAMPLE
 Get-CrmRecords -conn $conn -EntityLogicalName account -AttributeName name -Operator "eq" -Value "Adventure Works (sample)" -Fields name,accountnumber
 Key                  Value
 ---                  -----
 CrmRecords           {@{name_Property=[name, Adventure Works (sample)]; name=A...
 Count                5
 PagingCookie        <cookie page="1"><accountid last="{1FF8D93D-1F18-E511-80D...
 NextPage         False

 This example retrieves account(s) which name is "Adventure Works (sample)", with specified fields.

 .EXAMPLE
 Get-CrmRecords account name "like" "%(sample)%" name,accountnumber -AllRows
 Key                  Value
 ---                  -----
 CrmRecords           {@{name_Property=[name, Adventure Works (sample)]; name=A...
 Count                5
 PagingCookie        <cookie page="1"><accountid last="{1FF8D93D-1F18-E511-80D...
 NextPage         False
 
 This example retrieves account(s) which name includes "sample", with specified fields by omitting parameter names.
 When ommiting parameter names, you do not provide $conn, cmdlets automatically finds it.

#>

    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1)][alias("EntityName")]
        [string]$EntityLogicalName,
        [parameter(Mandatory=$false, Position=2)][alias("FieldName")]
        [string]$FilterAttribute,
        [parameter(Mandatory=$false, Position=3)][alias("Op")]
        [string]$FilterOperator,
        [parameter(Mandatory=$false, Position=4)][alias("Value", "FieldValue")]
        [string]$FilterValue,
        [parameter(Mandatory=$false, Position=5)]
        [string[]]$Fields, 
        [parameter(Mandatory=$false, Position=6)]
        [switch]$AllRows,
        [parameter(Mandatory=$false, Position=7)]
        [int]$TopCount
    )
    if($conn -eq $null)
    {
        $connobj = Get-Variable conn -Scope global -ErrorAction SilentlyContinue
        if($connobj.Value -eq $null)
        {
            Write-Warning 'Get-CrmRecords() - You need to create Connect to CRM Organization. Use Get-CrmConnection to create it.'
            break;
        }
        else
        {
            $conn = $connobj.Value
        }
    }

    if($FilterOperator)
    {
        $FilterOperator = $FilterOperator.Replace("-","")
    }
    if( !($EntityLogicalName -cmatch "^[a-z]*$") )
    {
        $EntityLogicalName = $EntityLogicalName.ToLower(); 
        Write-Verbose "EntityLogicalName contains uppercase which isn't possible in CRM, overwritting with ToLower() new value: $EntityLogicalName"
    }

    if( ($Fields -eq "*") -OR ($Fields -eq "%") )
    {
        Write-Warning 'PERFORMANCE: All attributes were requested'
        $fetchAttributes = "<all-attributes/>"
    }
    elseif ($Fields)
    {
        foreach($Field in $Fields)
        {
            if($field -ne $null){
                $fetchAttributes += "<attribute name='{0}' />" -F $Field
            }
        }
    }
    else
    {
        #lookup the primary attribute 
        $primaryAttribute = $conn.GetEntityMetadata($EntityLogicalName.ToLower()).PrimaryIdAttribute;
        $fetchAttributes = "<attribute name='{0}' />" -F $primaryAttribute
    }

    #if any of the values are missing, but they're not *ALL* missing 
    if( (!$FilterAttribute -OR !$FilterOperator -OR !$FilterValue) -AND ($FilterAttribute -Or $FilterOperator -Or $FilterValue) -And !($FilterAttribute -And $FilterOperator -And $FilterValue))
    {
        #TODO: convert this to a parameter set to avoid this extra logic
        Write-Error "One of the `$FilterAttribute `$FilterOperator `$FilterValue parameters is empty, to query all records exclude all filter parameters."
        return; 
    }
    
    if($FilterAttribute -and $FilterOperator -and $FilterValue)
    {
        Write-Verbose "Using the supplied single filter of $FilterAttribute '$FilterOperator' $FilterValue"
        $fetch = 
@"
    <fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="false" no-lock="true">
        <entity name="{0}">
            {1}
            <filter type='and'>
                <condition attribute='{2}' operator='{3}' value='{4}' />
            </filter>
        </entity>
    </fetch>
"@
    }
    else
    {
        Write-Warning "PERFORMANCE: `$FilterAttribute `$FilterOperator `$FilterValue were not supplied, fetching all records with NO filter."
        $fetch = 
@"
    <fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="false" no-lock="true">
        <entity name="{0}">
            {1}
        </entity>
    </fetch>
"@
    }
    $fetch = $fetch -F $EntityLogicalName, $fetchAttributes, $FilterAttribute, $FilterOperator, $FilterValue
    
    if($AllRows)
    {
        Write-Warning "PERFORMANCE: All rows were requested instead of the first 5000"
        $results = Get-CrmRecordsByFetch -conn $conn -Fetch $fetch -AllRows
    }
    else
    {
        $results = Get-CrmRecordsByFetch -conn $conn -Fetch $fetch -TopCount $TopCount
    }
    
    return $results
}

function Get-CrmRecordsByViewName{

<#
 .SYNOPSIS
 Retrieves CRM records by using View Name.

 .DESCRIPTION
 The Get-CrmRecordsByViewName cmdlet lets you retrieve CRM records by using View Name. You can use IsUserView parameter to select SystemView or UserView.

 .PARAMETER conn
 A connection to your CRM organizatoin. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .PARAMETER ViewName
 A name of a view, which contains desired FetchXML.

 .PARAMETER IsUserView
 Speficy $True if the view is User View.
 
 .EXAMPLE
 Get-CrmRecordsByViewName -conn $conn -ViewName "Active Accounts"
 Key            Value                                                                                  
 ---            -----                                                                                  
 CrmRecords     {@{accountid_Property=[accountid, 2bf8d93d-1f18-e511-80da-c4346bc43d94]; accountid=2...
 Count          5
 PagingCookie  <cookie page="1"><name last="account907" first="A. Datum Corporation (sample)" /><ac...
 NextPage   True                                                                                   
 
 This example retrieves account records by using "Active Accounts" system view.

 .EXAMPLE
 Get-CrmRecordsByViewName "My Custom Account View" $True
 Key            Value                                                                                  
 ---            -----                                                                                  
 CrmRecords     {@{accountid_Property=[accountid, 2bf8d93d-1f18-e511-80da-c4346bc43d94]; accountid=2...
 Count          5
 PagingCookie  <cookie page="1"><name last="account907" first="A. Datum Corporation (sample)" /><ac...
 NextPage   True                                                                                   
 

 This example retrieves account records by using "My Custom Account View" user view by omitting parameter names.
 When ommiting parameter names, you do not provide $conn, cmdlets automatically finds it.
  
#>

    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1)]
        [string]$ViewName,
        [parameter(Mandatory=$false, Position=2)]
        [bool]$IsUserView, 
        [parameter(Mandatory=$false, Position=3)]
        [switch]$AllRows,
        [parameter(Mandatory=$false, Position=4)]
        [int]$TopCount
    )

    if($conn -eq $null)
    {
        $connobj = Get-Variable conn -Scope global -ErrorAction SilentlyContinue
        if($connobj.Value -eq $null)
        {
            Write-Warning 'You need to create Connect to CRM Organization. Use Get-CrmConnection to create it.'
            break;
        }
        else
        {
            $conn = $connobj.Value
        }
    }
    
    if($IsUserView)
    {
        $fetch = @"
    <fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="false" no-lock="true">
        <entity name="userquery">
            <attribute name="fetchxml" />
            <filter type='and'>
                <condition attribute='name' operator='eq' value='{0}' />
            </filter>
        </entity>
    </fetch>
"@
    }
    else
    {
        $fetch = @"
    <fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="false" no-lock="true">
        <entity name="savedquery">
            <attribute name="fetchxml" />
            <filter type='and'>
                <condition attribute='name' operator='eq' value='{0}' />
            </filter>
        </entity>
    </fetch>
"@
    }

    $fetch = $fetch -F $ViewName
    #get the views matching the search phrase
    $views = Get-CrmRecordsByFetch -conn $conn -Fetch $fetch

    if($views.CrmRecords.Count -eq 0)
    {
        Write-Warning "Couldn't find the view"
        break
    }
    if($AllRows)
    {
        $results = Get-CrmRecordsByFetch -conn $conn -Fetch $views.CrmRecords[0].fetchxml -AllRows
    }
    else
    {
        $results = Get-CrmRecordsByFetch -conn $conn -Fetch $views.CrmRecords[0].fetchxml -TopCount $TopCount
    }
    return $results
}

function Get-CrmRecordsCount{
<#
 .SYNOPSIS
 Retrieves CRM entity total record counts.

 .DESCRIPTION
 The Get-CrmRecordsCount cmdlet lets you retrieve CRM entity total record counts.

 .PARAMETER conn
 A connection to your CRM organizatoin. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .PARAMETER EntityLogicalName
 A logicalname for an Entity to retrieve. i.e.)accout, contact, lead, etc..

 .EXAMPLE
 Get-CrmRecordsCount -conn $conn -EntityLogicalName account
 5677

 This example retrieves total number of Account entity records.

 .EXAMPLE
 Get-CrmRecordsCount account
 5677

 This example retrieves total number of Account entity records by omitting parameter names.
 When ommiting parameter names, you do not provide $conn, cmdlets automatically finds it.

#>

    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1)][alias("EntityName")]
        [string]$EntityLogicalName
    )
    if($conn -eq $null)
    {
        $connobj = Get-Variable conn -Scope global -ErrorAction SilentlyContinue
        if($connobj.Value -eq $null)
        {
            Write-Warning 'Get-CrmRecords() - You need to create Connect to CRM Organization. Use Get-CrmConnection to create it.'
            break;
        }
        else
        {
            $conn = $connobj.Value
        }
    }
        
    $fetch = 
@"
<fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="false" no-lock="true">
    <entity name="{0}">
        <attribute name='{1}' />
    </entity>
</fetch>
"@
    
    if($EntityLogicalName -eq "usersettings")
    {
        $PrimaryKeyField = "systemuserid"
    }
    else
    {
        $PrimaryKeyField = "$EntityLogicalName`id"
    }
    $fetch = $fetch -F $EntityLogicalName, $PrimaryKeyField
    
    $results = Get-CrmRecordsByFetch -conn $conn -Fetch $fetch -AllRows
        
    return $results.Count
}

function Get-CrmSiteMap{

<#
 .SYNOPSIS
 Retrieves CRM Organization's SiteMap information.

 .DESCRIPTION
 The Get-CrmSiteMap cmdlet lets you retrieve CRM Organization's SiteMap information.

 .PARAMETER conn
 A connection to your CRM organizatoin. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .PARAMETER SiteXml
 Specify SiteXml switch to retrive SiteMapXml xml data.

 .PARAMETER Areas
 Specify Areas switch to retrive AreaIds list.
 
 .PARAMETER GroupsOfArea
 Passing AreaId to GroupsOfArea parameter retrievs GroupIds of the specified Area.

 .PARAMETER SubAreasOfArea
 Passing AreaId to SubAreasOfArea parameter retrievs SubAreaIds of the specified Area.
    
 .EXAMPLE
 Get-CrmSiteMap -conn $conn -SiteXml
 <SiteMap IntroducedVersion="7.0.0.0"><Area Id="SFA" ResourceId="Area_Sales"..

 This example retrieves CRM Organization's SiteMap data.

 .EXAMPLE
 Get-CrmSiteMap -conn $conn -Areas
 SFA
 CS
 MA
 Settings
 HLP

 This example retrieves AreaIds of CRM Organization's SiteMap.

 .EXAMPLE
 Get-CrmSiteMap -conn $conn -GroupsOfArea SFA
 MyWork
 Customers
 SFA
 Collateral
 MA
 Goals
 Tools

 This example retrieves GroupIds of CRM Organization's SiteMap.
 
 .EXAMPLE
 Get-CrmSiteMap -SubAreasOfArea SFA
 nav_dashboards
 nav_personalwall
 nav_activities
 nav_accts
 nav_conts
 nav_leads
 nav_oppts
 ...

 This example retrieves SubAreaIds of CRM Organization's SiteMap by omitting parameter names.
 When ommiting parameter names, you do not provide $conn, cmdlets automatically finds it.

#>

    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$false, ParameterSetName="ShowXml")]
        [switch]$SiteXml,
        [parameter(Mandatory=$false, ParameterSetName="ShowArea")]
        [switch]$Areas,
        [parameter(Mandatory=$false, ParameterSetName="ShowGroupOfArea")]
        [string]$GroupsOfArea,
        [parameter(Mandatory=$false, ParameterSetName="ShowSubAreaOfArea")]
        [string]$SubAreasOfArea
    )

    if($conn -eq $null)
    {
        $connobj = Get-Variable conn -Scope global -ErrorAction SilentlyContinue
        if($connobj.Value -eq $null)
        {
            Write-Warning 'You need to create Connect to CRM Organization. Use Get-CrmConnection to create it.'
            break;
        }
        else
        {
            $conn = $connobj.Value
        }
    }
  
    $fetch = @"
    <fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="false" no-lock="true">
        <entity name="sitemap">
            <attribute name="sitemapxml" />            
        </entity>
    </fetch>
"@

    $record = (Get-CrmRecordsByFetch -conn $conn -Fetch $fetch).CrmRecords[0]

    $sitemap  = [xml]$record.sitemapxml

    if($SiteXml.IsPresent)
    {
        return $sitemap.InnerXml
    }
    if($Areas.IsPresent)
    {
        return $sitemap.SelectNodes("/SiteMap").Area.Id
    }
    if($GroupsOfArea.Length -ne 0)
    {
        return $sitemap.SelectNodes("/SiteMap/Area[@Id='$GroupsOfArea']/Group").Id
    }
    if($SubAreasOfArea.Length -ne 0)
    {
        return $sitemap.SelectNodes("/SiteMap/Area[@Id='$SubAreasOfArea']/Group/SubArea").Id
    }    
}

function Get-CrmSystemSettings{

<#
 .SYNOPSIS
 Retrieves CRM Organization's System Settings.

 .DESCRIPTION
 The Get-CrmSystemSettings cmdlet lets you retrieve CRM Organization's System Settings.

 .PARAMETER conn
 A connection to your CRM organizatoin. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .PARAMETER ShowDisplayName
 When you specify the ShowDisplayName switch, you see DisplayName for all fields, otherwise you see SchemaName.
   
 .EXAMPLE
 Get-CrmSystemSettings -conn $conn
 AllowUsersSeeAppdownloadMessage                  : Yes
 IsPresenceEnabled                                : Yes
 defaultemailsettings:incomingemaildeliverymethod : Server-Side Synchronization or Email Router
 defaultemailsettings:outgoingemaildeliverymethod : Server-Side Synchronization or Email Router
 defaultemailsettings:actdeliverymethod           : Microsoft Dynamics CRM for Outlook
 ...

 This example retrieves CRM Organization's System Settings.

 .EXAMPLE
 Get-CrmSystemSettings -conn $conn -ShowDisplayName
 Allow the showing tablet application notification bars in a browser. : Yes
 Presence Enabled                                                     : Yes
 Default Email Settings:Incoming Email                                : Server-Side Synchronization or Email Router
 Default Email Settings:Outgoing Email                                : Server-Side Synchronization or Email Router
 Default Email Settings:Appointments, Contacts, and Tasks             : Microsoft Dynamics CRM for Outlook
 ...

 This example retrieves CRM Organization's System Settings and show DisplayName for fields.

 .EXAMPLE
 Get-CrmSystemSettings
 AllowUsersSeeAppdownloadMessage                  : Yes
 IsPresenceEnabled                                : Yes
 defaultemailsettings:incomingemaildeliverymethod : Server-Side Synchronization or Email Router
 defaultemailsettings:outgoingemaildeliverymethod : Server-Side Synchronization or Email Router
 defaultemailsettings:actdeliverymethod           : Microsoft Dynamics CRM for Outlook
 ...

 This example retrieves CRM Organization's System Settings by omitting parameter names.
 When ommiting parameter names, you do not provide $conn, cmdlets automatically finds it.

 .EXAMPLE
 Get-CrmSystemSettings -ShowDisplayName
 Allow the showing tablet application notification bars in a browser. : Yes
 Presence Enabled                                                     : Yes
 Default Email Settings:Incoming Email                                : Server-Side Synchronization or Email Router
 Default Email Settings:Outgoing Email                                : Server-Side Synchronization or Email Router
 Default Email Settings:Appointments, Contacts, and Tasks             : Microsoft Dynamics CRM for Outlook
 ...

 This example retrieves CRM Organization's System Settings and show DisplayName for fields by omitting parameter names.
 When ommiting parameter names, you do not provide $conn, cmdlets automatically finds it.

#>

    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$false)]
        [switch]$ShowDisplayName
    )

    if($conn -eq $null)
    {
        $connobj = Get-Variable conn -Scope global -ErrorAction SilentlyContinue
        if($connobj.Value -eq $null)
        {
            Write-Warning 'You need to create Connect to CRM Organization. Use Get-CrmConnection to create it.'
            break;
        }
        else
        {
            $conn = $connobj.Value
        }
    }
  
    $fetch = @"
    <fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="false" no-lock="true">
        <entity name="organization">
            <all-attributes />
        </entity>
    </fetch>
"@
   
    $record = (Get-CrmRecordsByFetch -conn $conn -Fetch $fetch).CrmRecords[0]

    $attributes = Get-CrmEntityAttributes -conn $conn -EntityLogicalName organization

    $psobj = New-Object -TypeName System.Management.Automation.PSObject
        
    foreach($att in $record.original.GetEnumerator())
    {
        if(($att.Key.Contains("Property")) -or ($att.Key -eq "organizationid"))
        {
            continue
        }
        if($att.Key -eq "defaultemailsettings")
        {
            if($ShowDisplayName)
            {
                $name = ($attributes | where {$_.LogicalName -eq $att.Key}).Displayname.UserLocalizedLabel.Label + ":" +((Get-CrmEntityOptionSet mailbox incomingemaildeliverymethod).DisplayValue) 
            }
            else
            {
                $name = "defaultemailsettings:incomingemaildeliverymethod"
            }
            Add-Member -InputObject $psobj -MemberType NoteProperty -Name $name -Value ((Get-CrmEntityOptionSet mailbox incomingemaildeliverymethod).Items.DisplayLabel)[([xml]$att.Value).FirstChild.IncomingEmailDeliveryMethod]
            if($ShowDisplayName)
            {
                $name = ($attributes | where {$_.LogicalName -eq $att.Key}).Displayname.UserLocalizedLabel.Label + ":" +((Get-CrmEntityOptionSet mailbox outgoingemaildeliverymethod).DisplayValue) 
            }
            else
            {
                $name = "defaultemailsettings:outgoingemaildeliverymethod"
            }
            Add-Member -InputObject $psobj -MemberType NoteProperty -Name $name -Value ((Get-CrmEntityOptionSet mailbox outgoingemaildeliverymethod).Items.DisplayLabel)[([xml]$att.Value).FirstChild.OutgoingEmailDeliveryMethod]
            if($ShowDisplayName)
            {
                $name = ($attributes | where {$_.LogicalName -eq $att.Key}).Displayname.UserLocalizedLabel.Label + ":" +((Get-CrmEntityOptionSet mailbox actdeliverymethod).DisplayValue) 
            }
            else
            {
                $name = "defaultemailsettings:actdeliverymethod"
            }
            Add-Member -InputObject $psobj -MemberType NoteProperty -Name $name -Value ((Get-CrmEntityOptionSet mailbox actdeliverymethod).Items.DisplayLabel)[([xml]$att.Value).FirstChild.ACTDeliveryMethod]
            continue
        }
        
        if($ShowDisplayName)
        {
            $name = ($attributes | where {$_.LogicalName -eq $att.Key}).Displayname.UserLocalizedLabel.Label 
            if($name -eq $null)
            {
                $name = ($attributes | where {$_.LogicalName -eq $att.Key}).SchemaName
            }
        }
        else
        {
            $name = ($attributes | where {$_.LogicalName -eq $att.Key}).SchemaName
        }
        Add-Member -InputObject $psobj -MemberType NoteProperty -Name $name -Value $record.($att.Key) 
    }

    return $psobj
}

function Get-CrmTimeZones{
<#
 .SYNOPSIS
 Retrieves CRM Timezone information.

 .DESCRIPTION
 The Get-CrmTimeZones cmdlet lets you retrieve CRM Timezone information.

 .PARAMETER conn
 A connection to your CRM organizatoin. Use $conn = Get-CrmConnection <Parameters> to generate it.
  
 .EXAMPLE
 Get-CrmTimeZones -conn $conn
 Timezone Name                               TimeZoneCode                                                                           
 -------------                               ------------                                                                           
 (GMT-12:00) International Date Line West    0                                                                                      
 (GMT+13:00) Samoa                           1                                                                                      
 (GMT-10:00) Hawaii                          2                                                                                      
 (GMT-09:00) Alaska                          3                                                                                      
 (GMT-08:00) Pacific Time (US & Canada)      4
 ...  

 This example retrieves timezone information.

 .EXAMPLE
 Get-CrmTimeZones
 Timezone Name                               TimeZoneCode                                                                           
 -------------                               ------------                                                                           
 (GMT-12:00) International Date Line West    0                                                                                      
 (GMT+13:00) Samoa                           1                                                                                      
 (GMT-10:00) Hawaii                          2                                                                                      
 (GMT-09:00) Alaska                          3                                                                                      
 (GMT-08:00) Pacific Time (US & Canada)      4
 ...  

 This example retrieves timezone information by omitting parameter names.
 When ommiting parameter names, you do not provide $conn, cmdlets automatically finds it.
  
#>

    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn
    )

    if($conn -eq $null)
    {
        $connobj = Get-Variable conn -Scope global -ErrorAction SilentlyContinue
        if($connobj.Value -eq $null)
        {
            Write-Warning 'You need to create Connect to CRM Organization. Use Get-CrmConnection to create it.'
            break;
        }
        else
        {
            $conn = $connobj.Value
        }
    }
    
    $fetch = @"
    <fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="false" no-lock="true">
        <entity name="timezonedefinition">
            <attribute name="standardname" />
            <attribute name="timezonecode" />
            <attribute name="userinterfacename" />
            <order attribute="timezonecode" descending="false" />
        </entity>
    </fetch>
"@
    $results = Get-CrmRecordsByFetch -conn $conn -Fetch $fetch

    return $results.CrmRecords | select @{name="Timezone Name";expression={$_.userinterfacename}},@{name="TimeZone Code";expression={$_.timezonecode}}
}

function Get-CrmFailedWorkflows{

<#
 .SYNOPSIS
 Retrieves alert notifications from CRM organization.

 .DESCRIPTION
 The Get-CrmFailedWorkflows cmdlet lets you retrieve failed workflows from a CRM organization.

 .PARAMETER conn
 A connection to your CRM organizatoin. Use $conn = Get-CrmConnection <Parameters> to generate it.
  
 .EXAMPLE
 Key            Value
 ---            -----
 CrmRecords     {@{startedon_Property=[startedon, 9/14/2015 8:03:11 AM]; startedon=9/14/2015 3:03 AM;....
 Count          2565
 PagingCookie   <cookie page="1"><modifiedon last="2015-05-10T01:43:22-03:00" first="2015-05-10T01:4...
 NextPage       False
 FetchXml       <fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="false">...

 This example retrieves failed workflow (asyncoperation) records.

 .EXAMPLE
 Get-CrmFailedWorkflows
 Key            Value
 ---            -----
 CrmRecords     {@{startedon_Property=[startedon, 9/14/2015 8:03:11 AM]; startedon=9/14/2015 3:03 AM;....
 Count          2565
 PagingCookie   <cookie page="1"><modifiedon last="2015-05-10T01:43:22-03:00" first="2015-05-10T01:4...
 NextPage       False
 FetchXml       <fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="false">...

 This example retrieves failed workflow records notifications by omitting parameter names.
 When ommiting parameter names, you do not provide $conn, cmdlets automatically finds it.

 .EXAMPLE
 Get-CrmFailedWorkflows | % {$_.CrmRecords} | select message,startedon
startedon           message
---------           -------
9/14/2015 3:03 AM   Unhandled Exception: System.ServiceModel.FaultException`1[[Microsoft.Xrm.Sdk.OrganizationServiceFault...
9/13/2015 3:03 AM   Unhandled Exception: System.ServiceModel.FaultException`1[[Microsoft.Xrm.Sdk.OrganizationServiceFault...
9/12/2015 3:03 AM   Unhandled Exception: System.ServiceModel.FaultException`1[[Microsoft.Xrm.Sdk.OrganizationServiceFault...
9/11/2015 3:03 AM   Unhandled Exception: System.ServiceModel.FaultException`1[[Microsoft.Xrm.Sdk.OrganizationServiceFault...
9/10/2015 3:03 AM   Unhandled Exception: System.ServiceModel.FaultException`1[[Microsoft.Xrm.Sdk.OrganizationServiceFault...
9/9/2015 2:46 PM    Plugin Trace:...
9/9/2015 3:03 AM    Unhandled Exception: System.ServiceModel.FaultException`1[[Microsoft.Xrm.Sdk.OrganizationServiceFault...
9/8/2015 3:03 AM    Unhandled Exception: System.ServiceModel.FaultException`1[[Microsoft.Xrm.Sdk.OrganizationServiceFault...
9/7/2015 3:02 AM    Unhandled Exception: System.ServiceModel.FaultException`1[[Microsoft.Xrm.Sdk.OrganizationServiceFault...
9/6/2015 3:03 AM    Unhandled Exception: System.ServiceModel.FaultException`1[[Microsoft.Xrm.Sdk.OrganizationServiceFault...
9/5/2015 3:02 AM    Unhandled Exception: System.ServiceModel.FaultException`1[[Microsoft.Xrm.Sdk.OrganizationServiceFault...
9/4/2015 7:30 AM    Plugin Trace:...
9/4/2015 3:02 AM    Unhandled Exception: System.ServiceModel.FaultException`1[[Microsoft.Xrm.Sdk.OrganizationServiceFault...
9/3/2015 3:02 AM    Unhandled Exception: System.ServiceModel.FaultException`1[[Microsoft.Xrm.Sdk.OrganizationServiceFault...
9/2/2015 3:02 AM    Unhandled Exception: System.ServiceModel.FaultException`1[[Microsoft.Xrm.Sdk.OrganizationServiceFault...
9/1/2015 3:02 AM    Unhandled Exception: System.ServiceModel.FaultException`1[[Microsoft.Xrm.Sdk.OrganizationServiceFault...
8/31/2015 4:49 PM   Plugin Trace:...
8/31/2015 4:33 PM   Plugin Trace:...
8/31/2015 4:33 PM   Plugin Trace:...
 ...
 
 This example retrieves workflow errors and displays them with the startedon and message attributes.

#>

    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$false, Position=2)]
        [int]$TopCount,
        [parameter(Mandatory=$false, Position=3)]
        [int]$PageNumber,
        [parameter(Mandatory=$false, Position=4)]
        [string]$PageCookie,
        [parameter(Mandatory=$false, Position=5)]
        [switch]$AllRows

    )
    
    if($conn -eq $null)
    {
        $connobj = Get-Variable conn -Scope global -ErrorAction SilentlyContinue
        if($connobj.Value -eq $null)
        {
            Write-Warning 'Get-CrmTraceAlerts(): You need to create Connect to CRM Organization. Use Get-CrmConnection to create it.'
            break;
        }
        else
        {
            $conn = $connobj.Value
        }
    }
    
    $fetch = @"
<fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="false" no-lock="true">
              <entity name="asyncoperation">
                <attribute name="asyncoperationid" />
                <attribute name="name" />
                <attribute name="regardingobjectid" />
                <attribute name="operationtype" />
                <attribute name="statuscode" />
                <attribute name="ownerid" />
                <attribute name="startedon" />
                <attribute name="statecode" />
                <attribute name="workflowstagename" />
                <attribute name="postponeuntil" />
                <attribute name="owningextensionid" />
                <attribute name="modifiedon" />
                <attribute name="modifiedonbehalfby" />
                <attribute name="modifiedby" />
                <attribute name="messagename" />
                <attribute name="message" />
                <attribute name="friendlymessage" />
                <attribute name="errorcode" />
                <attribute name="createdon" />
                <attribute name="createdonbehalfby" />
                <attribute name="createdby" />
                <attribute name="completedon" />
                <order attribute="startedon" descending="true" />
                <filter type="and">
                  <condition attribute="recurrencestarttime" operator="null" />
                  <condition attribute="message" operator="not-null" />
                </filter>
              </entity>
            </fetch>
"@
    
    if($AllRows){
        $results = Get-CrmRecordsByFetch -conn $conn -Fetch $fetch -TopCount $TopCount -PageNumber $PageNumber -PageCookie $PagingCookie -AllRows 
    }
    else{
        $results = Get-CrmRecordsByFetch -conn $conn -Fetch $fetch -TopCount $TopCount -PageNumber $PageNumber -PageCookie $PagingCookie 
    }

    if($results.CrmRecords.Count -eq 0)
    {
        Write-Warning 'No failed worklfows found.'
    }
    else
    {
        return $results
    }
}

function Get-CrmTraceAlerts{

<#
 .SYNOPSIS
 Retrieves alert notifications from CRM organization.

 .DESCRIPTION
 The Get-CrmTraceAlerts cmdlet lets you retrieve alert notifications from CRM organization.

 .PARAMETER conn
 A connection to your CRM organizatoin. Use $conn = Get-CrmConnection <Parameters> to generate it.
  
 .EXAMPLE
 Get-CrmTraceAlerts -conn $conn
 Key            Value
 ---            -----
 CrmRecords     {@{tracecode_Property=[tracecode, 66]; tracecode=66; text_Property=[text, One or mor...
 Count          5
 PagingCookie  <cookie page="1"><modifiedon last="2015-05-10T01:43:22-03:00" first="2015-05-10T01:4...
 NextPage   False

 This example retrieves alert notifications.

 .EXAMPLE
 Get-CrmTraceAlerts
 Key            Value
 ---            -----
 CrmRecords     {@{tracecode_Property=[tracecode, 66]; tracecode=66; text_Property=[text, One or mor...
 Count          5
 PagingCookie  <cookie page="1"><modifiedon last="2015-05-10T01:43:22-03:00" first="2015-05-10T01:4...
 NextPage   False

 This example retrieves alert notifications by omitting parameter names.
 When ommiting parameter names, you do not provide $conn, cmdlets automatically finds it.

 .EXAMPLE
 Get-CrmTraceAlerts | % {$_.CrmRecords} | select text,level
 text                                                                                    level                                                                                  
 ----                                                                                    -----                                                                                  
 One or more mailboxes associated with the email server profile @[9605,7eedeb22-6a30-... Error                                                                                  
 Appointments, contacts and tasks can't be synchronized for mailbox @[9606,d0deee3f-6... Information                                                                            
 Appointments, contacts, and tasks can't be synchronized for your mailbox @[9606,d0de... Information                                                                            
 Email cannnot be be received because the email address of the mailbox @[9606,28c1e1a... Error                                                                                  
 The mailbox @[9606,e453c89b-6417-e511-80dc-c4346bc4fc6c,"Support Queue"] can't recei... Information                                                                            
 Your mailbox @[9606,e453c89b-6417-e511-80dc-c4346bc4fc6c,"Support Queue"] can't rece... Information                                                                            
 The mailbox @[9606,d0deee3f-6a17-e511-80dc-c4346bc4fc6c,"<sample team>"] can't recei... Information                                                                            
 Your mailbox @[9606,d0deee3f-6a17-e511-80dc-c4346bc4fc6c,"<sample team>"] can't rece... Information                                                                            
 Appointments, contacts and tasks can't be synchronized for mailbox @[9606,28c1e1af-6... Information
 ...
 
 This example retrieves alert notifications and display its text and level.

#>

    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn
    )

    if($conn -eq $null)
    {
        $connobj = Get-Variable conn -Scope global -ErrorAction SilentlyContinue
        if($connobj.Value -eq $null)
        {
            Write-Warning 'You need to create Connect to CRM Organization. Use Get-CrmConnection to create it.'
            break;
        }
        else
        {
            $conn = $connobj.Value
        }
    }
    
    $fetch = @"
    <fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="false" no-lock="true">
        <entity name="tracelog">
            <attribute name="level" />
            <attribute name="regardingobjectid" />
            <attribute name="createdon" />
            <attribute name="modifiedon" />
            <attribute name="tracecode" />
            <attribute name="text" />
            <attribute name="modifiedby" />
            <attribute name="createdby" />
            <order attribute="modifiedon" descending="true" />
        </entity>
    </fetch>
"@
    $results = Get-CrmRecordsByFetch -conn $conn -Fetch $fetch

    if($results.CrmRecords.Count -eq 0)
    {
        Write-Warning 'No alert found.'
    }
    else
    {
        return $results
    }
}

function Get-CrmUserMailbox{

<#
 .SYNOPSIS
 Retrieves CRM user's mailbox.

 .DESCRIPTION
 The Get-CrmUserMailbox cmdlet lets you retrieve CRM user's mailbox.

 .PARAMETER conn
 A connection to your CRM organizatoin. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .PARAMETER UserId
 An Id (guid) of CRM User.

 .PARAMETER ShowDisplayName
 When you specify the ShowDisplayName switch, you see DisplayName for all fields, otherwise you see SchemaName.
   
 .EXAMPLE
 Get-CrmUserMailbox -conn $conn -UserId f9d40920-7a43-4f51-9749-0549c4caf67d
 EnabledForOutgoingEmail             : No
 ProcessAndDeleteEmails              : No
 EmailServerProfile                  : Microsoft Exchange Online
 IncomingEmailDeliveryMethod         : Server-Side Synchronization or Email Router
 OwnerId                             : nakamura kenichiro
 IsForwardMailbox                    : No
 ...

 This example retrieves User's mailbox.

 .EXAMPLE
 Get-CrmUserMailbox -conn $conn -UserId f9d40920-7a43-4f51-9749-0549c4caf67d -ShowDisplayName
 Enabled For Outgoing Email                    : No
 Delete Emails after Processing                : No
 Server Profile                                : Microsoft Exchange Online
 Incoming Email                                : Server-Side Synchronization or Email Router
 Owner                                         : nakamura kenichiro
 Is Forward Mailbox                            : No
 ...

 This example retrieves User's mailbox and shows DisplayName for each field.

 .EXAMPLE
 Get-CrmUserMailbox f9d40920-7a43-4f51-9749-0549c4caf67d -ShowDisplayName
 Enabled For Outgoing Email                    : No
 Delete Emails after Processing                : No
 Server Profile                                : Microsoft Exchange Online
 Incoming Email                                : Server-Side Synchronization or Email Router
 Owner                                         : nakamura kenichiro
 Is Forward Mailbox                            : No
 ...

 This example retrieves User's mailbox and shows DisplayName for each field by omitting parameter names.
 When ommiting parameter names, you do not provide $conn, cmdlets automatically finds it.

#>

    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1)]
        [string]$UserId,
        [parameter(Mandatory=$false)]
        [switch]$ShowDisplayName
    )

    if($conn -eq $null)
    {
        $connobj = Get-Variable conn -Scope global -ErrorAction SilentlyContinue
        if($connobj.Value -eq $null)
        {
            Write-Warning 'You need to create Connect to CRM Organization. Use Get-CrmConnection to create it.'
            break;
        }
        else
        {
            $conn = $connobj.Value
        }
    }
    
    $fetch = @"
    <fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="false" no-lock="true">
    <entity name="mailbox">
    <attribute name="name" />
    <attribute name="ownerid" />
    <attribute name="mailboxid" />
    <attribute name="emailserverprofile" />
    <attribute name="processemailreceivedafter" />
    <attribute name="outgoingemailstatus" />
    <attribute name="outgoingemaildeliverymethod" />
    <attribute name="testmailboxaccesscompletedon" />
    <attribute name="isforwardmailbox" />
    <attribute name="incomingemailstatus" />
    <attribute name="incomingemaildeliverymethod" />
    <attribute name="enabledforoutgoingemail" />
    <attribute name="enabledforincomingemail" />
    <attribute name="enabledforact" />
    <attribute name="emailaddress" />
    <attribute name="processanddeleteemails" />
    <attribute name="processinglastattemptedon" />
    <attribute name="actstatus" />
    <attribute name="actdeliverymethod" />
    <attribute name="allowemailconnectortousecredentials" />
    <filter type="and">
      <condition attribute="regardingobjectid" operator="eq" value="{$UserId}" />
    </filter>
  </entity>
</fetch>
"@

    $record = (Get-CrmRecordsByFetch -conn $conn -Fetch $fetch).CrmRecords[0]

    $attributes = Get-CrmEntityAttributes -conn $conn -EntityLogicalName mailbox

    $psobj = New-Object -TypeName System.Management.Automation.PSObject
        
    foreach($att in $record.original.GetEnumerator())
    {
        if(($att.Key.Contains("Property")) -or ($att.Key -eq "mailboxid"))
        {
            continue
        }
        
        $value = $record.($att.Key)

        if($ShowDisplayName -and (($attributes | where {$_.LogicalName -eq $att.Key}).Displayname.UserLocalizedLabel -ne $null))
        {
            $name = ($attributes | where {$_.LogicalName -eq $att.Key}).Displayname.UserLocalizedLabel.Label 
        }
        else
        {
            $name = ($attributes | where {$_.LogicalName -eq $att.Key}).SchemaName
        }
        Add-Member -InputObject $psobj -MemberType NoteProperty -Name $name -Value $value  
    }

    return $psobj
}

function Get-CrmUserPrivileges{

<#
 .SYNOPSIS
 Retrieves privileges a CRM User has.

 .DESCRIPTION
 The Get-CrmUserPrivileges cmdlet lets you retrieve privileges a CRM User has. Result set contains following properties.
 Depth: Accumulated privilege Depth
 PrivilegeId: Privilege ID
 PrivilegeName: Privilege Name
 Origin: Indicate where the privilege comes from. RoleName(TeamName):Depth format 

 .PARAMETER conn
 A connection to your CRM organizatoin. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .PARAMETER UserId
 An Id (guid) of CRM User.
   
 .EXAMPLE
 Get-CrmUserPrivileges -conn $conn -UserId f9d40920-7a43-4f51-9749-0549c4caf67d
 Depth PrivilegeId                          PrivilegeName           Origin                                                                  
 ----- -----------                          -------------           ------                                                                  
 Global 78777c10-09ab-4326-b4c8-cf5729702937 prvAppendActivity      CSR Manager(TeamA):Deep,Salesperson:Local,System Administrator:Global   
 Global 3004684d-5d30-40a8-a8c0-7e92a2c9f326 prvAppendnew_entitya   System Administrator:Global 
 ...

 This example retrieves privileges assigned to the CRM User.

 Get-CrmUserPrivileges f9d40920-7a43-4f51-9749-0549c4caf67d
 Depth PrivilegeId                          PrivilegeName           Origin                                                                  
 ----- -----------                          -------------           ------                                                                  
 Global 78777c10-09ab-4326-b4c8-cf5729702937 prvAppendActivity      CSR Manager(TeamA):Deep,Salesperson:Local,System Administrator:Global   
 Global 3004684d-5d30-40a8-a8c0-7e92a2c9f326 prvAppendnew_entitya   System Administrator:Global 
 ...

 This example retrieves privileges assigned to the CRM User by omitting parameter names.
 When ommiting parameter names, you do not provide $conn, cmdlets automatically finds it.

#>

    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1)]
        [string]$UserId
    )

    if($conn -eq $null)
    {
        $connobj = Get-Variable conn -Scope global -ErrorAction SilentlyContinue
        if($connobj.Value -eq $null)
        {
            Write-Warning 'You need to create Connect to CRM Organization. Use Get-CrmConnection to create it.'
            break;
        }
        else
        {
            $conn = $connobj.Value
        }
    }
    
    # Get User Rolls including Team
    $roles = Get-CrmUserSecurityRoles -conn $conn -UserId $UserId -IncludeTeamRoles
    # Get all privilege records for PrivilegeName
	$privilegeRecords = (Get-CrmRecords -conn $conn -EntityLogicalName privilege -Fields name,privilegeid -WarningAction SilentlyContinue).CrmRecords
    # Create hash for performance reason
    $privileges = @{}
    $privilegeRecords | % {$privileges[$_.privilegeid] = $_.name} 

    # Create Result as hash for performance reason
    $results = @{}
    
	foreach($role in $roles)
	{
        # Get all privileges for the role
	    $request = New-Object Microsoft.Crm.Sdk.Messages.RetrieveRolePrivilegesRoleRequest
	        
	    try
	    {
	        $request.RoleId = $role.RoleId
	        $rolePrivileges = ($conn.ExecuteCrmOrganizationRequest($request, $null)).RolePrivileges            
	    }
	    catch
	    {
	        return $conn.LastCrmException
	    }	    
	    
	    foreach($rolePrivilege in $rolePrivileges)
	    {
            # Create origin as "RoleName(TeamName):Depth" format
            if($role.TeamName -eq $null) 
            {
                $origin = $role.RoleName + ":" + $rolePrivilege.Depth
            }
            else
            {
                $origin = $role.RoleName + "(" + $role.TeamName + "):" + $rolePrivilege.Depth
            }
                        
	        if($results.Contains($rolePrivilege.PrivilegeId))
	        {
                $existingObj = $results[$rolePrivilege.PrivilegeId]

                # Overwrite Depth only if it has higher privilege
                if([Microsoft.Crm.Sdk.Messages.PrivilegeDepth]::($rolePrivilege.Depth) -gt [Microsoft.Crm.Sdk.Messages.PrivilegeDepth]::($existingObj.Depth))
                {
                    $existingObj.Depth = $rolePrivilege.Depth
                }

                $existingObj.Origin += "," + $origin
	        }
	        else
	        {       
                # Create new result object     
	            $psobj = New-Object -TypeName System.Management.Automation.PSObject
	            Add-Member -InputObject $psobj -MemberType NoteProperty -Name "Depth" -Value $rolePrivilege.Depth
	            Add-Member -InputObject $psobj -MemberType NoteProperty -Name "PrivilegeId" -Value $rolePrivilege.PrivilegeId
	            Add-Member -InputObject $psobj -MemberType NoteProperty -Name "PrivilegeName" -Value $privileges[($rolePrivilege.PrivilegeId)]
	            Add-Member -InputObject $psobj -MemberType NoteProperty -Name "Origin" -Value $origin
	            $results[$rolePrivilege.PrivilegeId] = $psobj
	        }
	    }
	}
    
    return $results.Values | sort PrivilegeName
}

function Get-CrmUserSecurityRoles{

<#
 .SYNOPSIS
 Retrieves Security Roles assigned to a CRM User.

 .DESCRIPTION
 The Get-CrmUserSecurityRoles cmdlet lets you retrieve Security Roles assigned to a CRM User.

 .PARAMETER conn
 A connection to your CRM organizatoin. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .PARAMETER UserId
 An Id (guid) of CRM User.

 .PARAMETER IncludeTeamRoles
 When you specify the IncludeTeamRoles switch, Security Roles from teams are also retured.
   
 .EXAMPLE
 Get-CrmUserSecurityRoles -conn $conn -UserId f9d40920-7a43-4f51-9749-0549c4caf67d
 RoleName            
 --------            
 Salesperson         
 System Administrator
 ...

 This example retrieves Security Roles assigned to the CRM User.

 .EXAMPLE
 Get-CrmUserSecurityRoles -conn $conn -UserId f9d40920-7a43-4f51-9749-0549c4caf67d -IncludeTeamRoles
 RoleName             TeamName
 --------             --------
 CSR Manager          TeamA   
 Salesperson                  
 System Administrator 
 ...

 This example retrieves Security Roles assigned to the CRM User and Teams which the CRM User belongs to.

 .EXAMPLE
 Get-CrmUserSecurityRoles f9d40920-7a43-4f51-9749-0549c4caf67d -IncludeTeamRoles
 RoleName             TeamName
 --------             --------
 CSR Manager          TeamA   
 Salesperson                  
 System Administrator 
 ...

 This example retrieves Security Roles assigned to the CRM User and Teams which the CRM User belongs to by omitting parameter names.
 When ommiting parameter names, you do not provide $conn, cmdlets automatically finds it.

#>

    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1)]
        [string]$UserId,
        [parameter(Mandatory=$false)]
        [switch]$IncludeTeamRoles
    )

    if($conn -eq $null)
    {
        $connobj = Get-Variable conn -Scope global -ErrorAction SilentlyContinue
        if($connobj.Value -eq $null)
        {
            Write-Warning 'You need to create Connect to CRM Organization. Use Get-CrmConnection to create it.'
            break;
        }
        else
        {
            $conn = $connobj.Value
        }
    }
    
    $roles = New-Object System.Collections.Generic.List[PSObject]
    
    if($IncludeTeamRoles)
	{
		$fetch = @"
		<fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="true">
		  <entity name="role">
		    <attribute name="name"/>
			<attribute name="roleid" />
		    <link-entity name="teamroles" from="roleid" to="roleid" visible="false" intersect="true">
		      <link-entity name="team" from="teamid" to="teamid" alias="team">
		      <attribute name="name"/>
		        <link-entity name="teammembership" from="teamid" to="teamid" visible="false" intersect="true">
		          <link-entity name="systemuser" from="systemuserid" to="systemuserid" alias="af">
		            <filter type="and">
		              <condition attribute="systemuserid" operator="eq" value="{0}" />
		            </filter>
		          </link-entity>
		        </link-entity>
		      </link-entity>
		    </link-entity>
		  </entity>
		</fetch>
"@ -F $UserId
		
		(Get-CrmRecordsByFetch -conn $conn -Fetch $fetch).CrmRecords | select @{name="RoleId";expression={$_.roleid}}, @{name="RoleName";expression={$_.name}}, @{name="TeamName";expression={$_.'team.name'}} | % {$roles.Add($_)}	
	}

	$fetch = @"
	<fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="true">
	  <entity name="role">
	    <attribute name="name" />
	    <attribute name="roleid" />
	    <order attribute="name" descending="false" />
	    <link-entity name="systemuserroles" from="roleid" to="roleid" visible="false" intersect="true">
	      <link-entity name="systemuser" from="systemuserid" to="systemuserid" alias="ag">
	        <filter type="and">
	          <condition attribute="systemuserid" operator="eq" value="{0}" />
	        </filter>
	      </link-entity>
	    </link-entity>
	  </entity>
	</fetch>
"@ -F $UserId
	
	(Get-CrmRecordsByFetch -conn $conn -Fetch $fetch).CrmRecords | select @{name="RoleId";expression={$_.roleid}}, @{name="RoleName";expression={$_.name}} | % { $roles.Add($_) }
	
	return $roles
}

function Get-CrmUserSettings{

<#
 .SYNOPSIS
 Retrieves CRM user's settings.

 .DESCRIPTION
 The Get-CrmUserSettings cmdlet lets you retrieve CRM user's settings.

 .PARAMETER conn
 A connection to your CRM organizatoin. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .PARAMETER UserId
 An Id (guid) of CRM User.

 .PARAMETER Fields
 A List of field logicalnames. Use "fieldname1, fieldname2, fieldname3" syntax to speficy Fields, or ues "*" to retrieve all fields (not recommended for performance reason.)
   
 .EXAMPLE
 Get-CrmUserSettings -conn $conn -UserId f9d40920-7a43-4f51-9749-0549c4caf67d -Fields *
 timeformatcode_Property                             : [timeformatcode, 0]
 timeformatcode                                      : 0
 timezonestandardminute_Property                     : [timezonestandardminute, 0]
 timezonestandardminute                              : 0
 synccontactcompany_Property                         : [synccontactcompany, True]
 synccontactcompany                                  : Yes
 ...

 This example retrieves all fields from specified User's UserSettings.

 .EXAMPLE
 Get-CrmUserSettings (Get-MyCrmUserId) *
 timeformatcode_Property                             : [timeformatcode, 0]
 timeformatcode                                      : 0
 timezonestandardminute_Property                     : [timezonestandardminute, 0]
 timezonestandardminute                              : 0
 synccontactcompany_Property                         : [synccontactcompany, True]
 synccontactcompany                                  : Yes
 ...

 This example retrieves all fields from login User's UserSettings by omitting parameter names.
 When ommiting parameter names, you do not provide $conn, cmdlets automatically finds it.

#>

    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1)]
        [string]$UserId,
        [parameter(Mandatory=$true, Position=2)]
        [string[]]$Fields
    )

    if($conn -eq $null)
    {
        $connobj = Get-Variable conn -Scope global -ErrorAction SilentlyContinue
        if($connobj.Value -eq $null)
        {
            Write-Warning 'You need to create Connect to CRM Organization. Use Get-CrmConnection to create it.'
            break;
        }
        else
        {
            $conn = $connobj.Value
        }
    }
  
    return Get-CrmRecord -conn $conn -EntityLogicalName usersettings -Id $UserId -Fields $Fields
}

function Import-CrmSolutionTranslation{

<#
 .SYNOPSIS
 Imports a translation to a solution.

 .DESCRIPTION
 The Import-CrmSolutionTranslation cmdlet lets you export a translation to a solution.

 .PARAMETER conn
 A connection to your CRM organizatoin. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .PARAMETER TranslationFileName
 A file name and path of importing solution translation zip file.
 
 .PARAMETER PublishChanges
 Specify the parameter to publish all customizations.

 .EXAMPLE
 Import-CrmSolutionTranslation -conn $conn -TranslationZipFileName "C:\temp\CrmTranslations_MySolution_1_0_0_0.zip"

 This example imports translation file "CrmTranslations_MySolution_1_0_0_0".

 .EXAMPLE
 Import-CrmSolutionTranslation -conn $conn -TranslationZipFileName "C:\temp\CrmTranslations_MySolution_1_0_0_0.zip"

 This example imports translation file "CrmTranslations_MySolution_1_0_0_0" by ommiting $conn parameter.
 When ommiting $conn parameter, cmdlets automatically finds it.

 .EXAMPLE
 Import-CrmSolutionTranslation -conn $conn -TranslationZipFileName "C:\temp\CrmTranslations_MySolution_1_0_0_0.zip" -PublishChanges
 
 This example imports translation file "CrmTranslations_MySolution_1_0_0_0" and publish all customizations.
#>

    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1)]
        [string]$TranslationFileName,
        [parameter(Mandatory=$false, Position=2)]
        [switch]$PublishChanges
    )    

    if($conn -eq $null)
    {
        $connobj = Get-Variable conn -Scope global -ErrorAction SilentlyContinue
        if($connobj.Value -eq $null)
        {
            Write-Error 'You need to create a connection to a CRM Organization using get-CrmConnection or pass the connection as a parameter to use this cmdlet.'
            break;
        }
        else
        {
            $conn = $connobj.Value
        }
    }   
    try
    {
        $importId = [guid]::NewGuid()
        $translationFile = [System.IO.File]::ReadAllBytes($TranslationFileName);

        #create the import translation request then set all the properties
        $importRequest = New-Object Microsoft.Crm.Sdk.Messages.ImportTranslationRequest ; 
        $importRequest.TranslationFile = $translationFile
        $importRequest.ImportJobId = $importId
        
        Write-Verbose 'ImportTranslationRequest may take several minutes to complete execution.'
        $response = [Microsoft.Crm.Sdk.Messages.ImportTranslationResponse]($conn.ExecuteCrmOrganizationRequest($importRequest)); 
                
        Write-Verbose "Confirming the result"
        $xml = [xml](Get-CrmRecord -conn $conn -EntityLogicalName importjob -Id $importId -Fields data).data
        $importresult = $xml.importtranslations
        
        if($importresult.status -ne "Succeeded")
        {
            $importerrordetails = $importresult.errordetails
            Write-Verbose "Import result: $importerrordetails"
            throw $importerrordetails
        }
        else
        {            
            if($PublishChanges){
                Write-Verbose "Guid populated and user requested publish changes request..."
                Write-Verbose "Executing command: Publish-CrmAllCustomization, passing in the same connection"
            
                Publish-CrmAllCustomization -conn $conn
            }
            else{
                Write-Verbose "Import Complete don't forget to publish customizations."
            }
        }
    }
    catch
    {
        Write-Error $_.Exception
        #return $conn.LastCrmException
    }
}

function Invoke-CrmWhoAmI{

<#
 .SYNOPSIS
 Executes WhoAmI Organization Request and returns current user's Id (guid), belonging BusinessUnit Id (guid) and CRM Organization Id (guid).

 .DESCRIPTION
 The Invoke-CrmWhoAmI cmdlet lets you execute WhoAmI request and obtain UserId, BusinessUnitId and OrganizationId.

 .PARAMETER conn
 A connection to your CRM organizatoin. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .EXAMPLE
 Invoke-CrmWhoAmI -conn $conn

 This example executes WhoAmI organization request and returns current user's Id (guid), belonging BusinessUnit Id (guid) and CRM Organization Id (guid).

 .EXAMPLE
 Invoke-CrmWhoAmI
 
 This example executes WhoAmI by ommiting parameters names.
 When ommiting parameter names, you do not provide $conn, cmdlets automatically finds it.

#>

	[CmdletBinding()]
    PARAM( 
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn
    )

    if($conn -eq $null)
    {
        $connobj = Get-Variable conn -Scope global -ErrorAction SilentlyContinue
        if($connobj.Value -eq $null)
        {
            Write-Warning 'You need to create Connect to CRM Organization. Use Get-CrmConnection to create it.'
            break;
        }
        else
        {
            $conn = $connobj.Value
        }
    }  

    $request = New-Object Microsoft.Crm.Sdk.Messages.WhoAmIRequest
    
    try
    {
        $response = $conn.ExecuteCrmOrganizationRequest($request, $null)
    }
    catch
    {
        return $conn.LastCrmException
    }    

    return $response
}

function Publish-CrmAllCustomization{

<#
 .SYNOPSIS
 Publishes all customizations for a CRM Organization.

 .DESCRIPTION
 The Publish-CrmAllCustomization cmdlet lets you publish all customizations.

 .PARAMETER conn
 A connection to your CRM organizatoin. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .EXAMPLE
 Publish-CrmAllCustomization -conn $conn

 This example publishes all customizations.

 .EXAMPLE
 Publish-CrmAllCustomization
 
 This example publishes all customizations by ommiting parameters names.
 When ommiting parameter names, you do not provide $conn, cmdlets automatically finds it.

#>

 [CmdletBinding()]
    PARAM( 
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn
    )

    if($conn -eq $null)
    {
        $connobj = Get-Variable conn -Scope global -ErrorAction SilentlyContinue
        if($connobj.Value -eq $null)
        {
            Write-Warning 'You need to create Connect to CRM Organization. Use Get-CrmConnection to create it.'
            break;
        }
        else
        {
            $conn = $connobj.Value
        }
    }  

    $request = New-Object Microsoft.Crm.Sdk.Messages.PublishAllXmlRequest
    
    try
    {
        $response = $conn.ExecuteCrmOrganizationRequest($request, $null)
    }
    catch
    {
        return $conn.LastCrmException
    }    

    #return $response
}

function Remove-CrmSecurityRoleFromTeam{

<#
 .SYNOPSIS
 Removes a security role from a team.

 .DESCRIPTION
 The Set-CrmSecurityRoleToUser cmdlet lets you remove a security role from a team. 

 There are two ways to specify records.
 
 1. Pass record's Id for both records.
 2. Get a record object by using Get-CrmRecord/Get-CrmRecords cmdlets, and pass it for both records.

 .PARAMETER conn
 A connection to your CRM organizatoin. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .PARAMETER TeamRecord
 A team record object which is obtained via Get-CrmRecord/Get-CrmRecords. When you pass CrmRecord, then you don't use UserId.

 .PARAMETER SecurityRoleRecord
 A security role record object which is obtained via Get-CrmRecord/Get-CrmRecords. When you pass CrmRecord, then you don't use SecurityRoleId.

 .PARAMETER TeamId
 An Id (guid) of team record

 .PARAMETER SecurityRoleId
 An Id (guid) of security role record

 .EXAMPLE
 Remove-CrmSecurityRoleFromTeam -conn $conn -TeamId 00005a70-6317-e511-80da-c4346bc43d94 -SecurityRoleId 66005a70-6317-e511-80da-c4346bc43d94

 This example removes a security role to a team by using Id.

 .EXAMPLE
 Remove-CrmSecurityRoleFromTeam 00005a70-6317-e511-80da-c4346bc43d94 66005a70-6317-e511-80da-c4346bc43d94
 
 This example removes a security role to a team by using Id by ommiting parameters names.
 When ommiting parameter names, you do not provide $conn, cmdlets automatically finds it.

 .EXAMPLE
 PS C:\>$team = Get-CrmRecord team 00005a70-6317-e511-80da-c4346bc43d94 name

 PS C:\>$role = Get-CrmRecord role 66005a70-6317-e511-80da-c4346bc43d94 name

 PS C:\>Remove-CrmSecurityRoleFromTeam $team $role

 This example removes a security role to a team by using record objects.

#>

    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1, ParameterSetName="CrmRecord")]
        [PSObject]$TeamRecord,
        [parameter(Mandatory=$true, Position=2, ParameterSetName="CrmRecord")]
        [PSObject]$SecurityRoleRecord,
        [parameter(Mandatory=$true, Position=1, ParameterSetName="Id")]
        [string]$TeamId,
        [parameter(Mandatory=$true, Position=2, ParameterSetName="Id")]
        [string]$SecurityRoleId
    )

    if($conn -eq $null)
    {
        $connobj = Get-Variable conn -Scope global -ErrorAction SilentlyContinue
        if($connobj.Value -eq $null)
        {
            Write-Warning 'You need to create Connect to CRM Organization. Use Get-CrmConnection to create it.'
            break;
        }
        else
        {
            $conn = $connobj.Value
        }
    }

    if($PrincipalRecord -ne $null)
    {
        Remove-CrmRecordAssociation -conn $conn -CrmRecord1 $UserRecord -CrmRecord2 $SecurityRoleRecord -RelationshipName systemuserroles_association
    }
    else
    {
        Remove-CrmRecordAssociation -conn $conn -EntityLogicalName1 team -Id1 $TeamId -EntityLogicalName2 role -Id2 $SecurityRoleId -RelationshipName teamroles_association
    }
}

function Remove-CrmSecurityRoleFromUser{

<#
 .SYNOPSIS
 Removes a security role to a user.

 .DESCRIPTION
 The Remove-CrmSecurityRoleFromUser cmdlet lets you remove a security role to a user. 

 There are two ways to specify records.
 
 1. Pass record's Id for both records.
 2. Get a record object by using Get-CrmRecord/Get-CrmRecords cmdlets, and pass it for both records.

 .PARAMETER conn
 A connection to your CRM organizatoin. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .PARAMETER UserRecord
 A user record object which is obtained via Get-CrmRecord/Get-CrmRecords. When you pass CrmRecord, then you don't use UserId.

 .PARAMETER SecurityRoleRecord
 A security role record object which is obtained via Get-CrmRecord/Get-CrmRecords. When you pass CrmRecord, then you don't use SecurityRoleId.

 .PARAMETER UserId
 An Id (guid) of user record

 .PARAMETER SecurityRoleId
 An Id (guid) of security role record

 .EXAMPLE
 Remove-CrmSecurityRoleFromUser -conn $conn -UserId 00005a70-6317-e511-80da-c4346bc43d94 -SecurityRoleId 66005a70-6317-e511-80da-c4346bc43d94

 This example removes a security role to a user by using Id.

 .EXAMPLE
 Remove-CrmSecurityRoleFromUser 00005a70-6317-e511-80da-c4346bc43d94 66005a70-6317-e511-80da-c4346bc43d94
 
 This example removes a security role to a user by using Id by ommiting parameters names.
 When ommiting parameter names, you do not provide $conn, cmdlets automatically finds it.

 .EXAMPLE
 PS C:\>$user = Get-CrmRecord sysetmuser 00005a70-6317-e511-80da-c4346bc43d94 fullname

 PS C:\>$role = Get-CrmRecord role 66005a70-6317-e511-80da-c4346bc43d94 name

 PS C:\>Remove-CrmSecurityRoleFromUser $user $role

 This example removes a security role to a user by using record objects.

#>

    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1, ParameterSetName="CrmRecord")]
        [PSObject]$UserRecord,
        [parameter(Mandatory=$true, Position=2, ParameterSetName="CrmRecord")]
        [PSObject]$SecurityRoleRecord,
        [parameter(Mandatory=$true, Position=1, ParameterSetName="Id")]
        [string]$UserId,
        [parameter(Mandatory=$true, Position=2, ParameterSetName="Id")]
        [string]$SecurityRoleId
    )

    if($conn -eq $null)
    {
        $connobj = Get-Variable conn -Scope global -ErrorAction SilentlyContinue
        if($connobj.Value -eq $null)
        {
            Write-Warning 'You need to create Connect to CRM Organization. Use Get-CrmConnection to create it.'
            break;
        }
        else
        {
            $conn = $connobj.Value
        }
    }

    if($PrincipalRecord -ne $null)
    {
        Remove-CrmRecordAssociation -conn $conn -CrmRecord1 $UserRecord -CrmRecord2 $SecurityRoleRecord -RelationshipName systemuserroles_association
    }
    else
    {
        Remove-CrmRecordAssociation -conn $conn -EntityLogicalName1 systemuser -Id1 $UserId -EntityLogicalName2 role -Id2 $SecurityRoleId -RelationshipName systemuserroles_association
    }
}

function Remove-CrmUserManager{

<#
 .SYNOPSIS
 Removes CRM user's manager.

 .DESCRIPTION
 The Remove-CrmUserManager lets you remove CRM user's manager.

 .PARAMETER conn
 A connection to your CRM organizatoin. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .PARAMETER UserId
 An Id (guid) of CRM User.

 .EXAMPLE
 Remove-CrmUserManager -conn $conn -UserId 3772fe6e-8a18-e511-80dc-c4346bc42d48

 This example removes a manager from a CRM user.

 .EXAMPLE
 Remove-CrmUserManager 3772fe6e-8a18-e511-80dc-c4346bc42d48
 
 This example removes a manager from a CRM user by ommiting parameters names.
 When ommiting parameter names, you do not provide $conn, cmdlets automatically finds it.

#>

 [CmdletBinding()]
    PARAM( 
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1)]
        [guid]$UserId
    )

    if($conn -eq $null)
    {
        $connobj = Get-Variable conn -Scope global -ErrorAction SilentlyContinue
        if($connobj.Value -eq $null)
        {
            Write-Warning 'You need to create Connect to CRM Organization. Use Get-CrmConnection to create it.'
            break;
        }
        else
        {
            $conn = $connobj.Value
        }
    }

    $request = New-Object 'Microsoft.Crm.Sdk.Messages.RemoveParentRequest'
    $target = New-CrmEntityReference systemuser $UserId
    $request.Target = $target
    
    try
    {
        $response = $conn.ExecuteCrmOrganizationRequest($request, $null)
    }
    catch
    {
        return $conn.LastCrmException
    } 
}

function Set-CrmConnectionTimeout{

<#
 .SYNOPSIS
 Sets CRM Connection timeout value in seconds.

 .DESCRIPTION
 The Set-CrmConnectionTimeout lets you set CRM Connection timeout value in seconds.

 .PARAMETER conn
 A connection to your CRM organizatoin. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .PARAMETER TimeoutInSeconds
 Timeout value for CRM connection.

 .PARAMETER SetDefault
 Specyfing SetDefault will set default value to the connection. (120 seconds)

 .EXAMPLE
 Set-CrmConnectionTimeout -conn $conn -TimeoutInSeconds 1000

 This example sets CRM Connection timeout to 1000 seconds.

 .EXAMPLE
 Set-CrmConnectionTimeout -conn $conn -SetDefault
 
 This example sets CRM Connection timeout to default. (120 seconds)

#>

 [CmdletBinding()]
    PARAM( 
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$false, position=1)]
        [Int64]$TimeoutInSeconds,
        [parameter(Mandatory=$false, position=1)]
        [switch]$SetDefault
    )

    if($conn -eq $null)
    {
        $connobj = Get-Variable conn -Scope global -ErrorAction SilentlyContinue
        if($connobj.Value -eq $null)
        {
            Write-Warning 'You need to create Connect to CRM Organization. Use Get-CrmConnection to create it.'
            break;
        }
        else
        {
            $conn = $connobj.Value
        }
    }

    if($SetDefault)
    {
        $timeSpan = New-Object System.TimeSpan -ArgumentList 0,0,120
    }
    else
    {
        $currentTimeout = $conn.OrganizationServiceProxy.Timeout.TotalSeconds
        Write-Verbose "Current Timeout is $currentTimeout seconds"
        $timeSpan = New-Object System.TimeSpan -ArgumentList 0,0,$TimeoutInSeconds
    }

    $conn.OrganizationServiceProxy.Timeout = $timeSpan
}

function Set-CrmSystemSettings {

<#
 .SYNOPSIS
 Update CRM Organization's System Settings.

 .DESCRIPTION
 The Set-CrmSystemSettings cmdlet lets you update CRM Organization's System Settings. Use Get-CrmSystemSettings to confirm current settings.

 .PARAMETER conn
 A connection to your CRM organizatoin. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .PARAMETER ACTDeliveryMethod
 Change "Appointments, Contact, and Tasks" setting for Configure default synchronization method. This parameter accepts int. To get all options, use Get-CrmEntityOptionSet mailbox actdeliverymethod | % Items command and use PickListItemId.

 .PARAMETER AllowAddressBookSyncs 
 Change "Users can schedule background address book synchronization" setting.

 .PARAMETER AllowAutoResponseCreation 
 Indicates whether automatic response creation is allowed.

 .PARAMETER AllowAutoUnsubscribe 
 Indicates whether automatic unsubscribe is allowed.

 .PARAMETER AllowAutoUnsubscribeAcknowledgement 
 Change "Send acknowledgement to customers when they unsubscribe" setting.

 .PARAMETER AllowClientMessageBarAd 
 Change "Users see "Get CRM for Outlook" option displayed in the message bar" setting.

 .PARAMETER AllowEntityOnlyAudit 
 Indicates whether auditing of changes to an entity is allowed when no attributes have changed.

 .PARAMETER AllowMarketingEmailExecution 
 Indicates whether marketing emails execution is allowed

 .PARAMETER AllowOfflineScheduledSyncs 
 Change "Users can schedule background local data synchronization" setting.

 .PARAMETER AllowOutlookScheduledSyncs  
 Change "Users can schedule synchronization" setting.

 .PARAMETER AllowUnresolvedPartiesOnEmailSend 
 Change "Allw messages with unresolved email recipients to be sent" setting.

 .PARAMETER AllowUsersSeeAppdownloadMessage
 Change "Users see app download message" setting.

 .PARAMETER AllowWebExcelExport 
 Indicates whether web-based export of grids to Microsoft Office Excel is allowed.

 .PARAMETER BlockedAttachments
 Update "Set blocked file extensions for attachements" list.
  
 .PARAMETER CampaignPrefix
 Prefix used for campaign numbering.

 .PARAMETER CasePrefix
 Prefix to use for all cases throughout Microsoft Dynamics CRM.

 .PARAMETER ContractPrefix
 Prefix to use for all contracts throughout Microsoft Dynamics CRM.

 .PARAMETER CurrencyDisplayOption
 Change "Display currencies by using" setting. This parameter accepts int. To get all options, use Get-CrmEntityOptionSet organization currencydisplayoption | % Items command and use PickListItemId.

 .PARAMETER CurrentCampaignNumber 
 Current campaign number.

 .PARAMETER CurrentCaseNumber 
 First case number to use.

 .PARAMETER CurrentContractNumber  
 First contract number to use.

 .PARAMETER CurrentInvoiceNumber  
 First invoice number to use.

 .PARAMETER CurrentKbNumber  
 First article number to use.
 
 .PARAMETER CurrentOrderNumber  
 First order number to use.

 .PARAMETER CurrentQuoteNumber   
 First quote number to use.

 .PARAMETER DefaultCountryCode
 Chnage "Country/Region Code Prefix" value.

 .PARAMETER DefaultEmailServerProfileId
 Change "Server Profile" setting for Configure default synchronization method. This parameter accepts Email Server Profile record's guid. To get all profiles, use Get-CrmEmailServerProfiles command and use ProfileId.

 .PARAMETER DisableSocialCare
 Change "Prevent feature from receiving social data in CRM" setting.

 .PARAMETER DisplayNavigationTour
 Change "Display welcome screen to users when they sign in" setting.

 .PARAMETER EmailConnectionChannel
 Change "Process Email Using" setting. This parameter accepts int. To get all options, use Get-CrmEntityOptionSet organization emailcommunicationchannel | % Items command and use PickListItemId.

 .PARAMETER EmailCorrelationEnabled 
 Change "User correlation to track email conversations" setting.

 .PARAMETER EnableBingMapsIntegration
 Change "Show Bing Maps on forms" setting.

 .PARAMETER EnableSmartMatching 
 Change "Use Smart Matching" setting.

 .PARAMETER FullNameConventionCode
 Change "Name Format" setting for full-name fields. This parameters accept int. To get all options, use Get-CrmEntityOptionSet organization fullnameconventioncode | % Items command and use PickListItemId.

 .PARAMETER GenerateAlertsForErrors
 Change "Erorr level of Configure Alerts" setting.

 .PARAMETER GenerateAlertsForWarnings
 Change "Warning level of Configure Alerts" setting.

 .PARAMETER GenerateAlertsForInformation
 Change "Information level of Configure Alerts" setting.

 .PARAMETER GlobalHelpUrlEnabled
 Change "Use custom Help for customizable entities" setting.

 .PARAMETER GlobalHelpUrl
 Change "Global custom Help URL" value.

 .PARAMETER GlobalAppendUrlParametersEnabled
 Change "Append parameters to URL" setting.

 .PARAMETER HashDeltaSubjectCount
 Change "Maximum difference allowed between subject keywords" setting.
 
 .PARAMETER HashFilterKeywords 
 Change "Filter subject keywords" setting.

 .PARAMETER HashMaxCount 
 Change "Maximum number of subject keywords or recipients" setting.

 .PARAMETER HashMinAddressCount 
 Change "Minimum number of recipients required to match" setting.

 .PARAMETER IgnoreInternalEmail 
 Change "Track emails sent between CRM users as two activities" setting.

 .PARAMETER IncomingEmailDeliveryMethod
 Change "Incoming Email" setting for Configure default synchronization method. This parameter accepts int. To get all options, use Get-CrmEntityOptionSet mailbox incomingemaildeliverymethod | % Items command and use PickListItemId.

 .PARAMETER InvoicePrefix 
 Prefix to use for all invoice numbers throughout Microsoft Dynamics CRM.

 .PARAMETER IsAutoSaveEnabled
 Change "Enable auto save on all forms" setting.
  
 .PARAMETER IsDefaultCountryCodeCheckEnabled
 Change "Enable country/region code prefixing" setting.

 .PARAMETER IsDuplicateDetectionEnabled
 Indicates whether duplicate detection of records is enabled.

 .PARAMETER IsDuplicateDetectionEnabledForImport 
 Indicates whether duplicate detection of records during import is enabled.

 .PARAMETER IsDuplicateDetectionEnabledForOfflineSync  
 Indicates whether duplicate detection of records during offline synchronization is enabled.

 .PARAMETER IsDuplicateDetectionEnabledForOnlineCreateUpdate   
 Indicates whether duplicate detection during online create or update is enabled.

 .PARAMETER IsFolderBasedTrackingEnabled
 Change "Use folder-level tracking for Exchange folders" setting. Supported at v7.1 or above.

 .PARAMETER IsFullTextSearchEnabled 
 Change "Enable full-text search for Quick Find" setting.

 .PARAMETER IsHierarchicalSecurityModelEnabled 
 Indicates whether the hierarchical security model is enabled.

 .PARAMETER IsPresenceEnabled
 Change "Enable presence for the system" setting.

 .PARAMETER IsUserAccessAuditEnabled
 Change "Audit user access" setting.

 .PARAMETER KbPrefix 
 Prefix to use for all articles in Microsoft Dynamics CRM.
  
 .PARAMETER MaxAppointmentDurationDays 
 Set "Maximum durations of an appointment in days" setting.

 .PARAMETER MaxDepthForHierarchicalSecurityModel 
 Maximum depth for hierarchy security propagation.

 .PARAMETER MaxFolderBasedTrackingMappings 
 Maximum number of Folder Based Tracking mappings user can add.
 
 .PARAMETER MaximumActiveBusinessProcessFlowsAllowedPerEntity 
 Maximum number of active business process flows allowed per entity.

 .PARAMETER MaximumDynamicPropertiesAllowed 
 Maximum number of product properties for a product family or bundle.

 .PARAMETER MaximumTrackingNumber 
 Change "Number of digits for incremental message counter" setting. Specify 999 will set 3 as setting, 9999 will set 4 as setting and you can set up to 999999999. You can disable "Use Tracking Token" by setting 0.

 .PARAMETER MaxProductsInBundle 
 Maximum number of items in a bundle.

 .PARAMETER MaxRecordsForExportToExcel 
 Maximum number of records that will be exported to a static Microsoft Office Excel worksheet when exporting from the grid.

 .PARAMETER MaxRecordsForLookupFilters 
 Maximum number of lookup and picklist records that can be selected by user for filtering.

 .PARAMETER MaxUploadFileSize 
 Maximum allowed size of an attachment.

 .PARAMETER MinAddressBookSyncInterval 
 Change "Minimum time between address book synchronizations" setting.

 .PARAMETER MinOfflineSyncInterval 
 Change "Minimum time between background local data synchronizations" setting.

 .PARAMETER MinOutlookSyncInterval  
 Change "Minimum time between synchronizations" setting.

 .PARAMETER NotifyMailboxOwnerOfEmailServerLevelAlerts
 Change "Notify mailbox owner" setting.

 .PARAMETER OrderPrefix 
 Prefix to use for all orders throughout Microsoft Dynamics CRM.

 .PARAMETER OutgoingEmailDeliveryMethod
 Change "Outgoing Email" setting for Configure default synchronization method. This parameter accepts int. To get all options, use Get-CrmEntityOptionSet mailbox outgoingemaildeliverymethod | % Items command and use PickListItemId.

 .PARAMETER PluginTraceLogSetting
 Change "Enable logging to plug-in trace log" setting. 0:Off 1:Exception 2:All

 .PARAMETER PricingDecimalPrecision
 Change "Pricing Decimal Precision" setting.
 
 .PARAMETER QuickFindRecordLimitEnabled
 Change "Enable Quick Find record limits" setting.
  
 .PARAMETER QuotePrefix 
 Prefix to use for all quotes throughout Microsoft Dynamics CRM.

 .PARAMETER RenderSecureIFrameForEmail 
 Change "Use secure frames to restrict email message content" setting.

 .PARAMETER RequireApprovalForUserEmail
 Change "Process emails only for approved users" setting.

 .PARAMETER RequireApprovalForQueueEmail
 Change "Process emails only for approved queues" setting.

 .PARAMETER ShareToPreviousOwnerOnAssign
 Change "Share reassigned records with original owner" setting.
 
 .PARAMETER TrackingPrefix 
 Change Tracking Token "Prefix" setting.

 .PARAMETER TrackingTokenIdBase
 Change "Deployment base tracking number" setting.

 .PARAMETER TrackingTokenIdDigits  
 Change "Number of digits for user numbers" setting

 .PARAMETER UniqueSpecifierLength 
 Number of characters appended to invoice, quote, and order numbers.

 .PARAMETER UseLegacyRendering
 Change "Use legacy form rendering" setting.

 .PARAMETER UsePositionHierarchy 
 Indicates whether to use position hierarchy.

 .PARAMETER UseSkypeProtocol
 Change "Select provider for Click to call" setting. $true: Skype $false: Lync
 
 .EXAMPLE
 Set-CrmSystemSettings -conn $conn -IsAutoSaveEnabled $false

 This example disables "Enable auto save on all forms" System setting.

 .EXAMPLE
 Set-CrmSystemSettings -conn $conn -FullNameConventionCode 7
 
 This example updates "Name Format" of Set the full-name format to "Last NameFirst Name" System setting.

 .EXAMPLE
 Set-CrmSystemSettings -IsAutoSaveEnabled $false
 
 This example disables "Enable auto save on all forms" System setting by ommiting -conn parameter.
 When ommiting conn parameter, cmdlets automatically finds it.
#>

    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$false)]
        [int]$ACTDeliveryMethod,
        [parameter(Mandatory=$false)]
        [bool]$AllowAddressBookSyncs,
        [parameter(Mandatory=$false)]
        [bool]$AllowAutoResponseCreation,
        [parameter(Mandatory=$false)]
        [bool]$AllowAutoUnsubscribe,
        [parameter(Mandatory=$false)]
        [bool]$AllowAutoUnsubscribeAcknowledgement,
        [parameter(Mandatory=$false)]
        [bool]$AllowClientMessageBarAd,
        [parameter(Mandatory=$false)]
        [bool]$AllowEntityOnlyAudit,
        [parameter(Mandatory=$false)]
        [bool]$AllowMarketingEmailExecution,
        [parameter(Mandatory=$false)]
        [bool]$AllowOfflineScheduledSyncs,
        [parameter(Mandatory=$false)]
        [bool]$AllowOutlookScheduledSyncs,
        [parameter(Mandatory=$false)]
        [bool]$AllowUnresolvedPartiesOnEmailSend,
        [parameter(Mandatory=$false)]
        [bool]$AllowUsersSeeAppdownloadMessage,
        [parameter(Mandatory=$false)]
        [bool]$AllowWebExcelExport,
        [parameter(Mandatory=$false)]
        [string]$BlockedAttachments,
        [parameter(Mandatory=$false)]
        [string]$CampaignPrefix,
        [parameter(Mandatory=$false)]
        [string]$CasePrefix,
        [parameter(Mandatory=$false)]
        [string]$ContractPrefix,
        [parameter(Mandatory=$false)]
        [int]$CurrencyDisplayOption,
        [parameter(Mandatory=$false)]
        [int]$CurrentCampaignNumber,
        [parameter(Mandatory=$false)]
        [int]$CurrentCaseNumber,
        [parameter(Mandatory=$false)]
        [int]$CurrentContractNumber,
        [parameter(Mandatory=$false)]
        [int]$CurrentInvoiceNumber,
        [parameter(Mandatory=$false)]
        [int]$CurrentKbNumber,
        [parameter(Mandatory=$false)]
        [int]$CurrentOrderNumber,
        [parameter(Mandatory=$false)]
        [int]$CurrentQuoteNumber,
        [parameter(Mandatory=$false)]
        [ValidatePattern('\+{1}\d{1,}')]
        [string]$DefaultCountryCode,
        [parameter(Mandatory=$false)]
        [guid]$DefaultEmailServerProfileId,
        [parameter(Mandatory=$false)]
        [bool]$DisableSocialCare,
        [parameter(Mandatory=$false)]
        [bool]$DisplayNavigationTour,
        [parameter(Mandatory=$false)]
        [int]$EmailConnectionChannel, 
        [parameter(Mandatory=$false)]
        [int]$EmailCorrelationEnabled, 
        [parameter(Mandatory=$false)]
        [bool]$EnableBingMapsIntegration,
        [parameter(Mandatory=$false)]
        [bool]$EnableSmartMatching,
        [parameter(Mandatory=$false)]
        [int]$FullNameConventionCode,
        [parameter(Mandatory=$false)]
        [bool]$GenerateAlertsForErrors,
        [parameter(Mandatory=$false)]
        [bool]$GenerateAlertsForWarnings,
        [parameter(Mandatory=$false)]
        [bool]$GenerateAlertsForInformation,
        [parameter(Mandatory=$false)]
        [bool]$GlobalAppendUrlParametersEnabled,
        [parameter(Mandatory=$false)]
        [ValidatePattern('http(s)?://([\w-]+\.)+[\w-]+(/[\w- ./?%&=]*)?')]
        [string]$GlobalHelpUrl,
        [parameter(Mandatory=$false)]
        [bool]$GlobalHelpUrlEnabled,
        [parameter(Mandatory=$false)]
        [int]$HashDeltaSubjectCount,
        [parameter(Mandatory=$false)]
        [string]$HashFilterKeywords,
        [parameter(Mandatory=$false)]
        [int]$HashMaxCount,
        [parameter(Mandatory=$false)]
        [int]$HashMinAddressCount,
        [parameter(Mandatory=$false)]
        [bool]$IgnoreInternalEmail,
        [parameter(Mandatory=$false)]
        [int]$IncomingEmailDeliveryMethod,
        [parameter(Mandatory=$false)]
        [string]$InvoicePrefix,
        [parameter(Mandatory=$false)]
        [bool]$IsAutoSaveEnabled,
        [parameter(Mandatory=$false)]
        [bool]$IsDefaultCountryCodeCheckEnabled,
        [parameter(Mandatory=$false)]
        [bool]$IsDuplicateDetectionEnabled,
        [parameter(Mandatory=$false)]
        [bool]$IsDuplicateDetectionEnabledForImport,
        [parameter(Mandatory=$false)]
        [bool]$IsDuplicateDetectionEnabledForOfflineSync,
        [parameter(Mandatory=$false)]
        [bool]$IsDuplicateDetectionEnabledForOnlineCreateUpdate,
        [parameter(Mandatory=$false)]
        [bool]$IsFolderBasedTrackingEnabled,
        [parameter(Mandatory=$false)]
        [bool]$IsFullTextSearchEnabled,
        [parameter(Mandatory=$false)]
        [bool]$IsHierarchicalSecurityModelEnabled,
        [parameter(Mandatory=$false)]
        [bool]$IsPresenceEnabled,
        [parameter(Mandatory=$false)]
        [bool]$IsUserAccessAuditEnabled,
        [parameter(Mandatory=$false)]
        [string]$KbPrefix,
        [parameter(Mandatory=$false)]
        [int]$MaxAppointmentDurationDays,
        [parameter(Mandatory=$false)]
        [int]$MaxDepthForHierarchicalSecurityModel,
        [parameter(Mandatory=$false)]
        [int]$MaximumActiveBusinessProcessFlowsAllowedPerEntity,
        [parameter(Mandatory=$false)]
        [int]$MaximumDynamicPropertiesAllowed,
        [parameter(Mandatory=$false)]
        [int]$MaximumTrackingNumber,
        [parameter(Mandatory=$false)]
        [int]$MaxProductsInBundle,
        [parameter(Mandatory=$false)]
        [int]$MaxRecordsForExportToExcel,
        [parameter(Mandatory=$false)]
        [int]$MaxRecordsForLookupFilters,
        [parameter(Mandatory=$false)]
        [int]$MaxUploadFileSize,
        [parameter(Mandatory=$false)]
        [int]$MinAddressBookSyncInterval,
        [parameter(Mandatory=$false)]
        [int]$MinOfflineSyncInterval,
        [parameter(Mandatory=$false)]
        [int]$MinOutlookSyncInterval,
        [parameter(Mandatory=$false)]
        [bool]$NotifyMailboxOwnerOfEmailServerLevelAlerts,
        [parameter(Mandatory=$false)]
        [string]$OrderPrefix,
        [parameter(Mandatory=$false)]
        [int]$OutgoingEmailDeliveryMethod,
        [parameter(Mandatory=$false)]
        [ValidateSet(0,1,2)]
        [int]$PluginTraceLogSetting,
        [parameter(Mandatory=$false)]
        [ValidateSet(0,1,2,3,4)]
        [int]$PricingDecimalPrecision,
        [parameter(Mandatory=$false)]
        [bool]$QuickFindRecordLimitEnabled,
        [parameter(Mandatory=$false)]
        [string]$QuotePrefix,
        [parameter(Mandatory=$false)]
        [bool]$RequireApprovalForUserEmail,
        [parameter(Mandatory=$false)]
        [bool]$RequireApprovalForQueueEmail,
        [parameter(Mandatory=$false)]
        [bool]$ShareToPreviousOwnerOnAssign,
        [parameter(Mandatory=$false)]
        [string]$TrackingPrefix,
        [parameter(Mandatory=$false)]
        [int]$TrackingTokenIdBase,
        [parameter(Mandatory=$false)]
        [int]$TrackingTokenIdDigits,
        [parameter(Mandatory=$false)]
        [int]$UniqueSpecifierLength,
        [parameter(Mandatory=$false)]
        [bool]$UseLegacyRendering,
        [parameter(Mandatory=$false)]
        [bool]$UsePositionHierarchy,
        [parameter(Mandatory=$false)]
        [bool]$UseSkypeProtocol

    )

    if($conn -eq $null)
    {
        $connobj = Get-Variable conn -Scope global -ErrorAction SilentlyContinue
        if($connobj.Value -eq $null)
        {
            Write-Warning 'You need to create Connect to CRM Organization. Use Get-CrmConnection to create it.'
            break;
        }
        else
        {
            $conn = $connobj.Value
        }
    }
    
    $updateFields = @{}

    $attributesMetadata = Get-CrmEntityAttributes -EntityLogicalName organization
        
    $defaultEmailSettings = @{}        

    foreach($parameter in $MyInvocation.BoundParameters.GetEnumerator())
    {   
        $attributeMetadata = $attributesMetadata | ? {$_.SchemaName -eq $parameter.Key}

        if($parameter.Key -in ("IncomingEmailDeliveryMethod","OutgoingEmailDeliveryMethod","ACTDeliveryMethod"))
        {
            $defaultEmailSettings.Add($parameter.Key,$parameter.Value)
        }
        elseif($attributeMetadata -eq $null)
        {
            continue
        }
        elseif($attributeMetadata.AttributeType -eq "Picklist")
        {
            $updateFields.Add($parameter.Key.ToLower(), (New-CrmOptionSetValue $parameter.Value))
        }
        elseif($attributeMetadata.AttributeType -eq "Lookup")
        {
            $updateFields.Add($parameter.Key.ToLower(), (New-CrmEntityReference emailserverprofile $parameter.Value))
        }
        else
        {
            $updateFields.Add($parameter.Key.ToLower(), $parameter.Value)
        }
    }
    
    $fetch = @"
    <fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="false" no-lock="true">
        <entity name="organization">
            <attribute name="organizationid" />
            <attribute name="defaultemailsettings" />
        </entity>
    </fetch>
"@

    $systemSettings = (Get-CrmRecordsByFetch -conn $conn -Fetch $fetch).CrmRecords[0]
    $recordid = $systemSettings.organizationid

    if($defaultEmailSettings.Count -ne 0)
    {        
        $emailSettings = [xml]$systemSettings.defaultemailsettings
        if($defaultEmailSettings.ContainsKey("IncomingEmailDeliveryMethod"))
        {
            $emailSettings.SelectSingleNode("/EmailSettings/IncomingEmailDeliveryMethod").InnerText = $defaultEmailSettings["IncomingEmailDeliveryMethod"].Value
        }
        if($defaultEmailSettings.ContainsKey("OutgoingEmailDeliveryMethod"))
        {
            $emailSettings.SelectSingleNode("/EmailSettings/OutgoingEmailDeliveryMethod").InnerText = $defaultEmailSettings["OutgoingEmailDeliveryMethod"].Value
        }
        if($defaultEmailSettings.ContainsKey("ACTDeliveryMethod"))
        {
            $emailSettings.SelectSingleNode("/EmailSettings/ACTDeliveryMethod").InnerText = $defaultEmailSettings["ACTDeliveryMethod"].Value
        }

        $updateFields.Add("defaultemailsettings",$emailSettings.OuterXml);
    }

    Set-CrmRecord -conn $conn -EntityLogicalName organization -Id $recordid -Fields $updateFields
}

function Set-CrmUserMailbox {

<#
 .SYNOPSIS
 Updates CRM user's mailibox.

 .DESCRIPTION
 The Set-CrmUserMailbox cmdlet lets you update CRM user's mailibox. Use Get-CrmUserMailbox to confirm current values.

 .PARAMETER conn
 A connection to your CRM organizatoin. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .PARAMETER UserId
 An Id of CRM user.
 
 .PARAMETER EmailAddress
 An EmailAddress of CRM user. 

 .PARAMETER EmailServerProfile
 Change "Server Profile" setting of Synchronization Method. This parameter accepts Email Server Profile record's guid. To get all profiles, use Get-CrmEmailServerProfiles command and use ProfileId.

 .PARAMETER IncomingEmailDeliveryMethod
 Change "Incoming Email" setting for Configure default synchronization method. This parameter accepts int. To get all options, use Get-CrmEntityOptionSet mailbox incomingemaildeliverymethod | % Items command and use PickListItemId.

 .PARAMETER OutgoingEmailDeliveryMethod
 Change "Outgoing Email" setting for Configure default synchronization method. This parameter accepts int. To get all options, use Get-CrmEntityOptionSet mailbox outgoingemaildeliverymethod | % Items command and use PickListItemId.

 .PARAMETER ACTDeliveryMethod
 Change "Appointments, Contact, and Tasks" setting for Configure default synchronization method. This parameter accepts int. To get all options, use Get-CrmEntityOptionSet mailbox actdeliverymethod | % Items command and use PickListItemId.


 .EXAMPLE
 Set-CrmUserMailbox -conn $conn -UserId f9d40920-7a43-4f51-9749-0549c4caf67d -IncomingEmailDeliveryMethod 0

 This example updates "Incoming Email" setting to "None".

 .EXAMPLE
 Set-CrmUserMailbox -conn $conn -UserId f9d40920-7a43-4f51-9749-0549c4caf67d -EmailServerProfile 1b2d4b03-831e-e511-80e1-c4346bc44d24
 
 This example updates "Server Profile" setting to specified Profile.

 .EXAMPLE
 Set-CrmUserMailbox -conn $conn -UserId f9d40920-7a43-4f51-9749-0549c4caf67d -ApplyDefaultEmailSettings

 This example updates mailbox email settings to default settings, which is in System Settings.

 .EXAMPLE
 Set-CrmUserMailbox -UserId f9d40920-7a43-4f51-9749-0549c4caf67d -IncomingEmailDeliveryMethod 0
 
 This example disables "Incoming Email" setting to "None" by ommiting -conn parameter.
 When ommiting conn parameter, cmdlets automatically finds it.

 .EXAMPLE
 Set-CrmUserMailbox -UserId f9d40920-7a43-4f51-9749-0549c4caf67d -ApplyDefaultEmailSettings

 This example updates mailbox email settings to default settings, which is in System Settings by ommiting -conn parameter.
 When ommiting conn parameter, cmdlets automatically finds it.
#>

    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1)]
        [string]$UserId,
        [parameter(Mandatory=$false)]
        [ValidatePattern('/^([a-zA-Z0-9])+([a-zA-Z0-9\._-])*@([a-zA-Z0-9_-])+([a-zA-Z0-9\._-]+)+$/')]
        [string]$EmailAddress,
        [parameter(Mandatory=$false, ParameterSetName="Custom")]        
        [guid]$EmailServerProfile,
        [parameter(Mandatory=$false, ParameterSetName="Custom")]
        [int]$IncomingEmailDeliveryMethod,
        [parameter(Mandatory=$false, ParameterSetName="Custom")]
        [int]$OutgoingEmailDeliveryMethod,
        [parameter(Mandatory=$false, ParameterSetName="Custom")]
        [int]$ACTDeliveryMethod,
        [parameter(Mandatory=$false, ParameterSetName="Default")]
        [switch]$ApplyDefaultEmailSettings
    )

    if($conn -eq $null)
    {
        $connobj = Get-Variable conn -Scope global -ErrorAction SilentlyContinue
        if($connobj.Value -eq $null)
        {
            Write-Warning 'You need to create Connect to CRM Organization. Use Get-CrmConnection to create it.'
            break;
        }
        else
        {
            $conn = $connobj.Value
        }
    }
    
    $updateFields = @{}

    if($ApplyDefaultEmailSettings)
    {
        $fetch = @"
        <fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="false" no-lock="true">
            <entity name="organization">
                <attribute name="defaultemailserverprofileid" />       
                <attribute name="defaultemailsettings" />            
            </entity>
        </fetch>
"@
    
        $record = (Get-CrmRecordsByFetch -conn $conn -Fetch $fetch).CrmRecords[0]

        $updateFields.Add("emailserverprofile", $record.defaultemailserverprofileid_Property.Value)
        $xml = [xml]$record.defaultemailsettings

        $updateFields.Add("incomingemaildeliverymethod", (New-CrmOptionSetValue $xml.ChildNodes.IncomingEmailDeliveryMethod))
        $updateFields.Add("outgoingemaildeliverymethod", (New-CrmOptionSetValue $xml.ChildNodes.OutgoingEmailDeliveryMethod))
        $updateFields.Add("actdeliverymethod", (New-CrmOptionSetValue $xml.ChildNodes.ACTDeliveryMethod))
    }

    foreach($parameter in $MyInvocation.BoundParameters.GetEnumerator())
    {   
        if($parameter.Key -in ("UserId", "applydefaultemailsettings"))
        {
            continue
        }
        if($parameter.Key -in ("EmailServerProfile"))
        {
            $updateFields.Add($parameter.Key.ToLower(), (New-CrmEntityReference emailserverprofile $parameter.Value))
        }
        elseif($parameter.Key -in ("IncomingEmailDeliveryMethod","OutgoingEmailDeliveryMethod","ACTDeliveryMethod"))
        {
            $updateFields.Add($parameter.Key.ToLower(), (New-CrmOptionSetValue $parameter.Value))
        }
        else
        {
            $updateFields.Add($parameter.Key.ToLower(), $parameter.Value)
        }
    }

    $fetch = @"
    <fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="false" no>
  <entity name="mailbox">
    <attribute name="mailboxid" />
    <filter type="and">
      <condition attribute="regardingobjectid" operator="eq" value="{$UserId}" />
    </filter>
  </entity>
</fetch>
"@
    $Id = (`-conn $conn -Fetch $fetch).CrmRecords[0].MailboxId

    Set-CrmRecord -conn $conn -EntityLogicalName mailbox -Id $Id -Fields $updateFields
}

function Set-CrmUserManager{

<#
 .SYNOPSIS
 Sets CRM user's manager.

 .DESCRIPTION
 The Set-CrmUserManager lets you set CRM user's manager.

 .PARAMETER conn
 A connection to your CRM organizatoin. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .PARAMETER UserId
 An Id (guid) of CRM User.

 .PARAMETER ManagerId
 An Id (guid) of Manager.

 .PARAMETER KeepChildUsers
 Specify if you keep child users for the user.

 .EXAMPLE
 Set-CrmUserManager -conn $conn -UserId 3772fe6e-8a18-e511-80dc-c4346bc42d48 -ManagerId 5a18974c-ae18-e511-80dd-c4346bc44d24 -KeepChildUsers $True

 This example sets a manager to a CRM User and keeps its child users.

 .EXAMPLE
 Set-CrmUserManager 3772fe6e-8a18-e511-80dc-c4346bc42d48 5a18974c-ae18-e511-80dd-c4346bc44d24 $True
 
 This example sets a manager to a CRM User and keeps its child users by ommiting parameters names.
 When ommiting parameter names, you do not provide $conn, cmdlets automatically finds it.

#>

 [CmdletBinding()]
    PARAM( 
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1)]
        [guid]$UserId,
        [parameter(Mandatory=$true, Position=2)]
        [guid]$ManagerId,
        [parameter(Mandatory=$true, Position=3)]
        [bool]$KeepChildUsers
    )

    if($conn -eq $null)
    {
        $connobj = Get-Variable conn -Scope global -ErrorAction SilentlyContinue
        if($connobj.Value -eq $null)
        {
            Write-Warning 'You need to create Connect to CRM Organization. Use Get-CrmConnection to create it.'
            break;
        }
        else
        {
            $conn = $connobj.Value
        }
    }

    $request = New-Object 'Microsoft.Crm.Sdk.Messages.SetParentSystemUserRequest'
    $request.ParentId = $ManagerId
    $request.UserId = $UserId
    $request.KeepChildUsers = $KeepChildUsers

    try
    {
        $response = $conn.ExecuteCrmOrganizationRequest($request, $null)
    }
    catch
    {
        return $conn.LastCrmException
    } 
}

function Set-CrmUserBusinessUnit{

<#
 .SYNOPSIS
 Moves Crm User to another Business Unit.

 .DESCRIPTION
 The Set-CrmUserBusinessUnit lets you move Crm User to another Business Unit. You can specify different CRM UserId to ReassignUserId to update ownership of records as well. 

 .PARAMETER conn
 A connection to your CRM organizatoin. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .PARAMETER UserId
 An Id (guid) of CRM User which moves to another Business Unit.

 .PARAMETER BusinessUnitId
 An Id (guid) of Business Unit.

 .PARAMETER ReassignUserId
 An Id (guid) of CRM User to own records of Moving CRM User. You can specify same Id as UserId if you want to keep records ownership.

 .EXAMPLE
 Set-CrmUserBusinessUnit -conn $conn -UserId 3772fe6e-8a18-e511-80dc-c4346bc42d48 -BusinessUnitId 5a18974c-ae18-e511-80dd-c4346bc44d24 -ReassignUserId 3772fe6e-8a18-e511-80dc-c4346bc42d48

 This example moves a CRM User to specified BusinessUnit, then keeps the records ownership.

 .EXAMPLE
 Set-CrmUserBusinessUnit 3772fe6e-8a18-e511-80dc-c4346bc42d48 5a18974c-ae18-e511-80dd-c4346bc44d24 f9d40920-7a43-4f51-9749-0549c4caf67d
 
 This example moves a CRM User to specified BusinessUnit, then reassign the records ownership by ommiting parameters names.
 When ommiting parameter names, you do not provide $conn, cmdlets automatically finds it.

#>

 [CmdletBinding()]
    PARAM( 
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1)]
        [guid]$UserId,
        [parameter(Mandatory=$true, Position=2)]
        [guid]$BusinessUnitId,
        [parameter(Mandatory=$true, Position=3)]
        [guid]$ReassignUserId
    )

    if($conn -eq $null)
    {
        $connobj = Get-Variable conn -Scope global -ErrorAction SilentlyContinue
        if($connobj.Value -eq $null)
        {
            Write-Warning 'You need to create Connect to CRM Organization. Use Get-CrmConnection to create it.'
            break;
        }
        else
        {
            $conn = $connobj.Value
        }
    }

    $ReassignPrincipal = New-CrmEntityReference -EntityLogicalName systemuser -Id $ReassignUserId

    $request = New-Object 'Microsoft.Crm.Sdk.Messages.SetBusinessSystemUserRequest'
    $request.BusinessId = $BusinessUnitId
    $request.UserId = $UserId
    $request.ReassignPrincipal = $ReassignPrincipal

    try
    {
        $response = $conn.ExecuteCrmOrganizationRequest($request, $null)
    }
    catch
    {
        return $conn.LastCrmException
    }
}

function Set-CrmUserSettings{

<#
 .SYNOPSIS
 Update CRM user's settings.

 .DESCRIPTION
 The Set-CrmUserSettings cmdlet lets you update CRM user's settings.

 .PARAMETER conn
 A connection to your CRM organizatoin. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .PARAMETER CrmRecord
 A CRMUserSettings object. Use Get-CrmUserSettings to retrieve the setting.
   
 .EXAMPLE
 PS C:\>$userSettings = Get-CrmUserSettings -conn $conn -UserId (Get-MyCrmUserId) -Fields *
 PS C:\>$userSettings.timezonecode = 4
 PS C:\>Set-CrmUserSettings -conn $conn -CrmRecord $userSettings

 This example retrieves retrieves all fields from login User's UserSettings update TimeZone to Pacific Time.

 .EXAMPLE
 PS C:\>$userSettings = Get-CrmUserSettings (Get-MyCrmUserId) *
 PS C:\>$userSettings.paginglimit = 100
 PS C:\>Set-CrmUserSettings $userSettings

 This example retrieves all fields from login User's UserSettings and update PagingLimit to 100 by omitting parameter names.
 When ommiting parameter names, you do not provide $conn, cmdlets automatically finds it.

#>

    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1)]
        [Alias("UserSettingsRecord")]
        [PSObject]$CrmRecord
    )

    if($conn -eq $null)
    {
        $connobj = Get-Variable conn -Scope global -ErrorAction SilentlyContinue
        if($connobj.Value -eq $null)
        {
            Write-Warning 'You need to create Connect to CRM Organization. Use Get-CrmConnection to create it.'
            break;
        }
        else
        {
            $conn = $connobj.Value
        }
    }
    
    try
    {
        $result = Set-CrmRecord -conn $conn -CrmRecord $CrmRecord
    }
    catch
    {
        throw
    }    
}

### Get CRM Types object ###
function New-CrmMoney{

<#
 .SYNOPSIS
 Instantiates Money type object.

 .DESCRIPTION
 The New-CrmMoney cmdlet lets you instantiates Money type object. 
 
 .PARAMETER Value
 Money Value.

 .EXAMPLE
 New-CrmMoney -Value 1000
 Value ExtensionData
 ----- -------------
 1000
 
 This example instantiates Money object with Value of 1000.

 .EXAMPLE
 New-CrmMoney 1000.01
 Value ExtensionData
 ----- -------------
 1000.01

 This example instantiates Money object with Value of 1000.01 by ommiting parameter names.

#>

 [CmdletBinding()]
    PARAM(        
        [parameter(Mandatory=$true, Position=0)]
        [double]$Value
    )

    $crmMoney = New-Object -TypeName Microsoft.Xrm.Sdk.Money
    $crmMoney.Value = $Value

    return $crmMoney
}

function New-CrmOptionSetValue{

<#
 .SYNOPSIS
 Instantiates OptionSetValue type object.

 .DESCRIPTION
 The New-CrmOptionSetValue cmdlet lets you instantiates OptionSetValue type object. 
 
 .PARAMETER Value
 OptionSetValue Value.

 .EXAMPLE
 New-CrmOptionSetValue -Value 20
 Value ExtensionData
 ----- -------------
 20
 
 This example instantiates OptionSetValue object with Value of 20.

 .EXAMPLE
 New-CrmOptionSetValue 20
 Value ExtensionData
 ----- -------------
 20

 This example instantiates OptionSetValue object with Value of 20 by ommiting parameter names.

#>

 [CmdletBinding()]
    PARAM(        
        [parameter(Mandatory=$true, Position=0)]
        [int]$Value
    )

    $crmOptionSetValue = New-Object -TypeName Microsoft.Xrm.Sdk.OptionSetValue
    $crmOptionSetValue.Value = $Value

    return $crmOptionSetValue
}

function New-CrmEntityReference{

<#
 .SYNOPSIS
 Instantiates EntityReference type object.

 .DESCRIPTION
 The New-CrmEntityReference cmdlet lets you instantiates EntityReference type object. 
 
 .PARAMETER EntityLogicalName
 A logicalname for an Entity to update. i.e.)accout, contact, lead, etc..

 .PARAMETER Id
 An Id (guid) of the record

 .EXAMPLE
 New-CrmEntityReference -EntityLogicalName account -Id 1df8d93d-1f18-e511-80da-c4346bc43d94
 Id            : 1df8d93d-1f18-e511-80da-c4346bc43d94
 LogicalName   : account
 Name          : 
 KeyAttributes : {}
 RowVersion    : 
 ExtensionData :
 
 This example instantiates CrmEntityReference object for an account record.

 .EXAMPLE
 New-CrmEntityReference account 1df8d93d-1f18-e511-80da-c4346bc43d94
 Id            : 1df8d93d-1f18-e511-80da-c4346bc43d94
 LogicalName   : account
 Name          : 
 KeyAttributes : {}
 RowVersion    : 
 ExtensionData :

 This example instantiates CrmEntityReference object for an account record by ommiting parameter names.

#>

 [CmdletBinding()]
    PARAM(        
        [parameter(Mandatory=$true, Position=0)]
        [string]$EntityLogicalName,
        [parameter(Mandatory=$true, Position=1)]
        [guid]$Id
    )

    $crmEntityReference = New-Object -TypeName Microsoft.Xrm.Sdk.EntityReference

    $crmEntityReference.LogicalName = $EntityLogicalName
    $crmEntityReference.Id = $Id

    return $crmEntityReference
}

### Performance Test cmdlets ###
function Test-XrmTimerStart{

<#
 .SYNOPSIS
 Instantiate timer object and start it.

 .DESCRIPTION
 The Test-XrmTimerStart cmdlet lets you instantiate timer object and start it. Use with Text-XrmTimerStop.
 
 .EXAMPLE
 Test-XrmTimerStart; $result = Get-MyCrmUserId; Test-XrmTimerStop
 The operation took 00:00:00.0816009
 
 This example shows performance result of Get-MyCrmUserId.

#>
 
    $script:crmtimer = New-Object -TypeName 'System.Diagnostics.Stopwatch'
    $script:crmtimer.Start()
}

function Test-XrmTimerStop{

<#
 .SYNOPSIS
 Find timer object which started by Test-XrmTimerStart and stop it. Then display the elapsed time.

 .DESCRIPTION
 The Test-XrmTimerStop cmdlet lets you see the elapsed time after you call Test-XrmTimerStart.
 
 .EXAMPLE
 Test-XrmTimerStart; $result = Get-MyCrmUserId; Test-XrmTimerStop
 The operation took 00:00:00.0816009
 
 This example shows performance result of Get-MyCrmUserId.

#>

    $crmtimerobj = Get-Variable crmtimer -Scope Script
    if($crmtimerobj.Value -ne $null)
    {
        $crmtimer = $crmtimerobj.Value
        $crmtimer.Stop()
        $perf = "The operation took " + $crmtimer.Elapsed.ToString()
        Remove-Variable crmtimer  -Scope Script
        return $perf
    }
}