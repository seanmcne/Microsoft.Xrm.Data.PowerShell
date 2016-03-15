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

function Connect-CrmOnlineDiscovery{
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml
    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false)]
        [PSCredential]$Credential, 
        [Parameter(Mandatory=$false)]
        [switch]$UseCTP,
        [Parameter(Mandatory=$false)]
        [switch]$InteractiveMode
    )
        
    if($InteractiveMode)
    {
        $global:conn = Get-CrmConnection -InteractiveMode -Verbose
        
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
				$crmOrganizations = $crmOrganizations | sort-object FriendlyName
                foreach($crmOrganization in $crmOrganizations)
                {   $friendlyName = $crmOrganization.FriendlyName

                    $message = "[$i] $friendlyName (" + $crmOrganization.WebApplicationUrl + ")"
                    Write-Host $message 
                    $i++
                }
                $orgNumber = Read-Host "`nSelect CRM Organization by index number"
    
                Write-Verbose ($crmOrganizations[$orgNumber]).UniqueName
			}
            $global:conn = Get-CrmConnection -Credential $Credential -DeploymentRegion $crmOrganizations[$orgNumber].DiscoveryServerShortname -OnLineType $onlineType -OrganizationName ($crmOrganizations[$orgNumber]).UniqueName -Verbose

			#yes, we know this isn't recommended BUT this cmdlet is only valid for user interaction in the console and shouldn't be used for non-interactive scenarios
            Write-Host "`nYou are now connected to: $(($crmOrganizations[$orgNumber]).UniqueName)" -foregroundcolor yellow
			Write-Host "For a list of commands run: Get-Command -Module Microsoft.Xrm.Data.Powershell" -foregroundcolor yellow
            return $global:conn    
        }
    }
}

function Connect-CrmOnline{
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml
    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$true)]
        [PSCredential]$Credential, 
        [Parameter(Mandatory=$true)]
        [ValidatePattern('https://([\w-]+).crm([0-9]*).dynamics.com')]
        [string]$ServerUrl
    )
   
    $userName = $Credential.UserName
    $password = $Credential.GetNetworkCredential().Password
    $connectionString = "AuthType=Office365;Username=$userName; Password=$password;Url=$ServerUrl"

    try
    {
        $global:conn =  New-Object Microsoft.Xrm.Tooling.Connector.CrmServiceClient -ArgumentList $connectionString
        return $global:conn
    }
    catch
    {
        throw $_
    }    
}

function Connect-CrmOnPremDiscovery{
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml
    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false, ParameterSetName="ServerUrl")]
        [PSCredential]$Credential, 
		[Parameter(Mandatory=$true, ParameterSetName="ServerUrl")]
        [ValidatePattern('http(s)?://[\w-]+(/[\w- ./?%&=]*)?')]
        [Uri]$ServerUrl,
        [Parameter(Mandatory=$false, ParameterSetName="ServerUrl")]
        [string]$OrganizationName,
        [Parameter(Mandatory=$false, ParameterSetName="ServerUrl")]
        [string]$HomeRealmUrl,
        [Parameter(Mandatory=$false, ParameterSetName="InteractiveMode")]
        [switch]$InteractiveMode
    )
    
    if($InteractiveMode)
    {
        $global:conn = Get-CrmConnection -InteractiveMode -Verbose
        Write-Verbose "You are now connected and may run any of the CRM Commands."
        return $global:conn 
    }
    else
    {
        if($Credential -eq $null -And !$Interactive)
        {
            $Credential = Get-Credential
        }

        # If Organization Name is pased, use it, otherwise retrieve all organizations the user belongs to.
        if($OrganizationName -ne '')
        {
            $organizationName = $OrganizationName
        }
        else
        {
		    $crmOrganizations = Get-CrmOrganizations -Credential $Credential -ServerUrl $ServerUrl -Verbose 
        
            if($crmOrganizations.Count -gt 0)
            {    
		    	if($crmOrganizations.Count -eq 1)
                {
                    $orgNumber = 0
                }
		    	else
                {
                    $i = 0
		    		$crmOrganizations = $crmOrganizations | sort-object FriendlyName
                    foreach($crmOrganization in $crmOrganizations)
                    {   
		    			$friendlyName = $crmOrganization.FriendlyName
                        $message = "[$i] $friendlyName (" + $crmOrganization.WebApplicationUrl + ")"
                        Write-Host $message 
                        $i++
                    }
                    $orgNumber = Read-Host "`nSelect CRM Organization by index number"                                    
		    	}            
                
                # Store the OrganizationName
                Write-Verbose ($crmOrganizations[$orgNumber]).UniqueName    
                $organizationName = ($crmOrganizations[$orgNumber]).UniqueName
            }
            else
            {
                Write-Warning "User belongs to no organization."
                return
            }
        }          

        if($HomeRealmUrl -eq '')
        {
            $global:conn = Get-CrmConnection -Credential $Credential -ServerUrl $ServerUrl -OrganizationName $organizationName -Verbose
        }
        else
        {
            $global:conn = Get-CrmConnection -Credential $Credential -ServerUrl $ServerUrl -OrganizationName $organizationName -HomeRealmUrl $HomeRealmUrl -Verbose
        }
		#yes, we know this isn't recommended BUT this cmdlet is only valid for user interaction in the console and shouldn't be used for non-interactive scenarios
        Write-Host "`nYou are now connected to: $organizationName" -foregroundcolor yellow
		Write-Host "For a list of commands run: Get-Command -Module Microsoft.Xrm.Data.Powershell" -foregroundcolor yellow
        return $global:conn    
    }
}

function New-CrmRecord{
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml
    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1)]
        [string]$EntityLogicalName,
        [parameter(Mandatory=$true, Position=2)]
        [hashtable]$Fields
    )

	$conn = VerifyCrmConnectionParam $conn

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
			"Guid" {
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
        if(!$result -or $result -eq [System.Guid]::Empty)
        {
            throw $conn.LastCrmException
        }
    }
    catch
    {
        throw $conn.LastCrmException        
    }

    return $result
}

function Get-CrmRecord{
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml
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

	$conn = VerifyCrmConnectionParam $conn

    if($Fields -eq "*")
    {
        [Collections.Generic.List[String]]$x = $null
    }
    else
    {
        [Collections.Generic.List[String]]$x = $Fields
    }

    try
    {
        $record = $conn.GetEntityDataById($EntityLogicalName, $Id, $x, [Guid]::Empty)
    }
    catch
    {
        throw $conn.LastCrmException        
    }    
    
    if($record -eq $null)
    {        
        throw $conn.LastCrmException
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

#Alias for Set-CrmRecord
New-Alias -Name Update-CrmRecord -Value Set-CrmRecord

#UpdateEntity 

function Set-CrmRecord{
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml
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

	$conn = VerifyCrmConnectionParam $conn
    
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
            throw $conn.LastCrmException
        }
    }
    catch
    {
        #TODO: Throw Exceptions back to user
        throw $conn.LastCrmException
    }
}

#DeleteEntity 

function Remove-CrmRecord{
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml
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
        $conn = VerifyCrmConnectionParam $conn
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
                throw $conn.LastCrmException
            }
        }
        catch
        {
            throw $conn.LastCrmException
        }
    }
}

### Other Cmdlets from Xrm Tooling ###
#AddEntityToQueue 

function Move-CrmRecordToQueue{
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml
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
	$conn = VerifyCrmConnectionParam $conn  
    if($CrmRecord -ne $null)
    {
        $EntityLogicalName = $CrmRecord.logicalname
        $Id = $CrmRecord.($EntityLogicalName + "id")
    }

    try
    {
        $result = $conn.AddEntityToQueue($Id, $EntityLogicalName, $QueueName, $WorkingUserId, $SetWorkingByUser, [Guid]::Empty)
		if(!$result)
        {
            throw $conn.LastCrmException
        }
    }
    catch
    {
        throw $conn.LastCrmException
    }
}

#AssignEntityToUser

function Set-CrmRecordOwner{
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml
    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,        
        [parameter(Mandatory=$true, Position=1, ParameterSetName="CrmRecord", ValueFromPipeline=$true)]
        [PSObject]$CrmRecord,
        [parameter(Mandatory=$true, Position=1, ParameterSetName="NameWithId")]
        [string]$EntityLogicalName,
        [parameter(Mandatory=$true, Position=2, ParameterSetName="NameWithId")]
        [guid]$Id,
        [parameter(Mandatory=$true, Position=3)][alias("UserId")]
        [guid]$PrincipalId,
		[parameter(Mandatory=$false, Position=4)]
		[switch]$AssignToTeam
    )
    begin
    {
        $conn = VerifyCrmConnectionParam $conn
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
			# As CrmClientService does not have method to assign to team, use Organization Request
			if($AssignToTeam){
				write-verbose "Assigning record with Id: $Id to Team with Id: $PrincipalId"
				
				$req = New-Object Microsoft.Crm.Sdk.Messages.AssignRequest
				$req.target = New-CrmEntityReference -EntityLogicalName $EntityLogicalName -Id $Id
				$req.Assignee = New-CrmEntityReference -EntityLogicalName "team" -Id $PrincipalId
				$result = [Microsoft.Crm.Sdk.Messages.AssignResponse]$conn.ExecuteCrmOrganizationRequest($req, $null)
				# If no result returend, then it had an issue.
				if($result -eq $null)
                {
                    $result = $false
                }
			}
			else{
		        $result = $conn.AssignEntityToUser($PrincipalId, $EntityLogicalName, $Id, [Guid]::Empty)
			}			
			if(!$result)
            {
                throw $conn.LastCrmException
            }

			write-verbose "Completed..."
		}
		catch
		{
		    throw $conn.LastCrmException
		}
	}
}

#CloseActivity 

function Set-CrmActivityRecordToCloseState{
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml
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
	$conn = VerifyCrmConnectionParam $conn   
    if($CrmRecord -ne $null)
    {
        $ActivityEntityType = $CrmRecord.logicalname
        $ActivityId = $CrmRecord.("activityid")
    }
    try
    {
        $result = $conn.CloseActivity($ActivityEntityType, $ActivityId, $StateCode, $StatusCode, [Guid]::Empty)
		if(!$result)
        {
            throw $conn.LastCrmException
        }
    }
    catch
    {
        throw $conn.LastCrmException
    }
}

#CreateAnnotation 

function Add-CrmNoteToCrmRecord{
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml
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
	$conn = VerifyCrmConnectionParam $conn   
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
		if($result -eq $null)
        {
            throw $conn.LastCrmException
        }
    }
    catch
    {
        throw $conn.LastCrmException
    }
}

#CreateEntityAssociation

function Add-CrmRecordAssociation{
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml
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
	$conn = VerifyCrmConnectionParam $conn
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
		if(!$result)
        {
            throw $conn.LastCrmException
        }
    }
    catch
    {
        throw $conn.LastCrmException
    }
}

#CreateMultiEntityAssociation
function Add-CrmMultiRecordAssociation{
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml
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

	$conn = VerifyCrmConnectionParam $conn  

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
            break
        }
    }   

    try
    {
        $result = $conn.CreateMultiEntityAssociation($EntityLogicalName1, $Id1, $EntityLogicalName2, $Id2s, $RelationshipName, [Guid]::Empty, $IsReflexiveRelationship)
		if(!$result)
        {
            throw $conn.LastCrmException
        }
    }
    catch
    {
        throw $conn.LastCrmException
    }
}

#CreateNewActivityEntry 
function Add-CrmActivityToCrmRecord{
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml
    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,        
        [parameter(Mandatory=$true, Position=1, ParameterSetName="CrmRecord", ValueFromPipeline=$true)]
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
    begin
    {
        $conn = VerifyCrmConnectionParam $conn
    }  
	process
	{
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
			if($result -eq $null)
			{
				throw $conn.LastCrmException
			}
		}
		catch
		{
			throw $conn.LastCrmException
		}

		return $result
	}
}

#DeleteEntityAssociation
function Remove-CrmRecordAssociation{
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml
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
	$conn = VerifyCrmConnectionParam $conn    
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
		if(!$result)
        {
            throw $conn.LastCrmException
        }
    }
    catch
    {
        throw $conn.LastCrmException
    }
}

#ExecuteWorkflowOnEntity  
function Invoke-CrmRecordWorkflow{
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml
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
	$conn = VerifyCrmConnectionParam $conn
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
		if($result -eq $null)
        {
            throw $conn.LastCrmException
        }
    }
    catch
    {
        throw $conn.LastCrmException
    }
    return $result
}

#GetMyCrmUserId
function Get-MyCrmUserId{
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml
    [CmdletBinding()]
    PARAM( 
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn
    )

	$conn = VerifyCrmConnectionParam $conn

    try
    {
        $result = $conn.GetMyCrmUserId()
		if($result -eq $null) 
        {
            throw $conn.LastCrmException
        }
    }
    catch
    {
        throw $conn.LastCrmException
    }    

    return $result
}

#GetAllAttributesForEntity
function Get-CrmEntityAttributes{
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml
 [CmdletBinding()]
    PARAM( 
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1)]
        [string]$EntityLogicalName
    )

	$conn = VerifyCrmConnectionParam $conn 
       
    try
    {
        $result = $conn.GetAllAttributesForEntity($EntityLogicalName)
		if($result -eq $null)
        {
            throw $conn.LastCrmException
        }
    }
    catch
    {
        throw $conn.LastCrmException
    }    

    return $result
}

#GetAllEntityMetadata 
function Get-CrmEntityAllMetadata{
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml
 [CmdletBinding()]
    PARAM( 
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$false, Position=1)]
        [bool]$OnlyPublished=$true, 
        [parameter(Mandatory=$false, Position=2)]
        [string]$EntityFilters
    )
	$conn = VerifyCrmConnectionParam $conn 
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
        $result = $conn.GetAllEntityMetadata($OnlyPublished, $filter)
		if($result -eq $null)
        {
            throw $conn.LastCrmException
        }
    }
    catch
    {
        throw $conn.LastCrmException
    }    
    return $result
}

#GetEntityAttributeMetadataForAttribute  
function Get-CrmEntityAttributeMetadata{
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml
 [CmdletBinding()]
    PARAM( 
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1)]
        [string]$EntityLogicalName,
        [parameter(Mandatory=$true, Position=2)]
        [string]$FieldLogicalName
    )

	$conn = VerifyCrmConnectionParam $conn 
    
    try
    {
        $result = $conn.GetEntityAttributeMetadataForAttribute($EntityLogicalName, $FieldLogicalName)
		if($result -ne $null)
        {
            throw $conn.LastCrmException
        }
    }
    catch
    {
        throw $conn.LastCrmException
    }    
    return $result
}

#GetEntityDataByFetchSearch
function Get-CrmRecordsByFetch{
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml
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
    $conn = VerifyCrmConnectionParam $conn
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
            return $resultSet
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
        throw $conn.LastCrmException
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
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml
 [CmdletBinding()]
    PARAM( 
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1, ParameterSetName="EntityLogicalName")]
        [string]$EntityLogicalName
    )
	$conn = VerifyCrmConnectionParam $conn 
    try
    {
        $result = $conn.GetEntityDisplayName($EntityLogicalName)
		if($result -eq $null)
        {
			throw $conn.LastCrmException
        }
    }
    catch
    {
        throw $conn.LastCrmException
    }    
    return $result
}

#GetEntityDisplayNamePlural
function Get-CrmEntityDisplayPluralName{
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml
 [CmdletBinding()]
    PARAM( 
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1)]
        [string]$EntityLogicalName
    )
	$conn = VerifyCrmConnectionParam $conn 
    try
    {
        $result = $conn.GetEntityDisplayNamePlural($EntityLogicalName)
		if($result -eq $null)
        {
            throw $conn.LastCrmException
        }
    }
    catch
    {
        throw $conn.LastCrmException
    }     
    return $result
}

#GetEntityMetadata
function Get-CrmEntityMetadata{
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml
 [CmdletBinding()]
    PARAM( 
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1)]
        [string]$EntityLogicalName,
        [parameter(Mandatory=$false, Position=2)]
        [string]$EntityFilters
    )

	$conn = VerifyCrmConnectionParam $conn 

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
        $result = $conn.GetEntityMetadata($EntityLogicalName, $filter)
		if($result -eq $null)
        {
            throw $conn.LastCrmException
        }
    }
    catch
    {
        throw $conn.LastCrmException
    }    

    return $result
}

#GetEntityName
function Get-CrmEntityName{
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml
 [CmdletBinding()]
    PARAM( 
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1)]
        [int]$EntityTypeCode
    )

	$conn = VerifyCrmConnectionParam $conn  

    try
    {
        $result = $conn.GetEntityName($EntityTypeCode)
		if($result -eq $null)
        {
            throw $conn.LastCrmException
        }
    }
    catch
    {
        throw $conn.LastCrmException
    }   

    return $result
}

#GetEntityTypeCode 
function Get-CrmEntityTypeCode{
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml
 [CmdletBinding()]
    PARAM( 
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1)]
        [string]$EntityLogicalName
    )
	$conn = VerifyCrmConnectionParam $conn  
    try
    {
        $result = $conn.GetEntityTypeCode($EntityLogicalName)
		if($result -eq $null)
        {
            throw $conn.LastCrmException
        }
    }
    catch
    {
        throw $conn.LastCrmException
    }    
    return $result
}

#GetGlobalOptionSetMetadata  
function Get-CrmGlobalOptionSet{
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml
 [CmdletBinding()]
    PARAM( 
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1)]
        [string]$OptionSetName
    )
	$conn = VerifyCrmConnectionParam $conn  
    try
    {
        $result = $conn.GetGlobalOptionSetMetadata($OptionSetName)
		if($result -eq $null)
        {
            throw $conn.LastCrmException
        }
    }
    catch
    {
        throw $conn.LastCrmException
    }
    return $result
}

#GetPickListElementFromMetadataEntity   
function Get-CrmEntityOptionSet{
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml
 [CmdletBinding()]
    PARAM( 
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1)]
        [string]$EntityLogicalName,
        [parameter(Mandatory=$true, Position=2)]
        [string]$FieldLogicalName
    )
	$conn = VerifyCrmConnectionParam $conn
    try
    {
        $result = $conn.GetPickListElementFromMetadataEntity($EntityLogicalName, $FieldLogicalName)
		if($result -eq $null)
        {
            throw $conn.LastCrmException
        }
    }
    catch
    {
        throw $conn.LastCrmException
    }

    return $result
}

#ImportSolutionToCrm   
function Import-CrmSolution{
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml
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
        [switch]$PublishChanges,
		[parameter(Mandatory=$false, Position=6)]
        [int64]$MaxWaitTimeInSeconds = 900
    )
	$conn = VerifyCrmConnectionParam $conn
	$importId = [guid]::Empty
    try
    {
        if (!(Test-Path $SolutionFilePath)) 
        {
            throw [System.IO.FileNotFoundException] "$SolutionFilePath not found."
        }
        $tmpDest = $conn.CrmConnectOrgUriActual
        Write-Host "Importing solution file $SolutionFilePath into: $tmpDest" 
        Write-Verbose "OverwriteCustomizations: $OverwriteUnManagedCustomizations"
        Write-Verbose "SkipDependancyCheck: $SkipDependancyOnProductUpdateCheckOnInstall"
        Write-Verbose "Maximum seconds to poll for successful completion: $MaxWaitTimeInSeconds"
        Write-Verbose "Calling .ImportSolutionToCrm() this process can take minutes..."
        $result = $conn.ImportSolutionToCrm($SolutionFilePath, [ref]$importId, $ActivatePlugIns,
                $OverwriteUnManagedCustomizations, $SkipDependancyOnProductUpdateCheckOnInstall)
        $pollingStart = Get-Date
        $isProcessing = $true
		$secondsSpentPolling = 0
        $pollingDelaySeconds = 5
        Write-Host "Import of file completed, waiting on completion of importId: $importId"
		try{
			while($isProcessing -and $secondsSpentPolling -lt $MaxWaitTimeInSeconds){
				#delay
				Start-Sleep -Seconds $pollingDelaySeconds
				#check the import job for success/fail/inProgress
				$import = Get-CrmRecord -conn $conn -EntityLogicalName importjob -Id $importId -Fields solutionname,data,completedon,startedon,progress
				#Option to use Get-CrmRecords so we can force a no-lock to prevent hangs in the retrieve
				#$import = (Get-CrmRecords -conn $conn -EntityLogicalName importjob -FilterAttribute importjobid -FilterOperator eq -FilterValue $importId -Fields data,completedon,startedon,progress).CrmRecords[0]
				$importManifest = ([xml]($import).data).importexportxml.solutionManifests.solutionManifest
				$ProcPercent = [double](Coalesce $import.progress "0")

				#Check for import completion 
				if($import.completedon -eq $null -and $importManifest.result -ne "success"){
					$isProcessing = $true
					$secondsSpentPolling = ([Int]((Get-Date) - $pollingStart).TotalSeconds)
					Write-Output "$($secondsSPentPolling.ToString("000")) seconds of max: $MaxWaitTimeInSeconds ... ImportJob%: $ProcPercent"
				}
				else{
					Write-Verbose "Processing Completed at: $($import.completedon)" 
					$ProcPercent = 100.0
					$isProcessing = $false
					break
				}
			}
		} Catch {
			Write-Error "ImportJob with ID: $importId has encountered an exception: $_ "
		} Finally{
            $ProcPercent = ([double](Coalesce $ProcPercent 0))
        }
        #User provided timeout and exit function with an error
	    if($secondsSpentPolling -gt $MaxWaitTimeInSeconds){
			throw "Import-CrmSolution halted due to exceeding the maximum timeout of $MaxWaitTimeInSeconds."
		}
		#detect a failure by a failure result OR the percent being less than 100%
        if($importresult.result -eq "failure") #Must look at %age instead of this result as the result is usually wrong!
        {
            Write-Verbose "Import result: $($importManifest.result) - job with ID: $importId failed at $ProcPercent complete."
            throw $importresult.errortext
        }
        elseif($ProcPercent -lt 100){
            try{
                #lets try to dump the failure data as a best effort: 
                ([xml]$import.data).importexportxml.entities.entity|foreach {
                    if($_.result.result -ne $null -and $_.result.result -eq 'failure'){
                        write-output "Name: $($_.LocalizedName) Result: $($_.result.errorcode) Details: $($_.result.errortext)"
                        write-error "Name: $($_.LocalizedName) Result: $($_.result.errorcode) Details: $($_.result.errortext)"
                    }
                }
                #webresource problems
                ([xml]$import.data).importexportxml.webResources.webResource|foreach {
                    if($_.result.result -ne $null -and $_.result.result -eq 'failure'){
                        write-output "Name: $($_.LocalizedName) Result: $($_.result.errorcode) Details: $($_.result.errortext)"
                        write-error "Name: $($_.LocalizedName) Result: $($_.result.errorcode) Details: $($_.result.errortext)"
                    }
                }
                #optionset problems
                ([xml]$import.data).importexportxml.optionSets.optionset|foreach {
                    if($_.result.result -ne $null -and $_.result.result -eq 'failure'){
                        write-output "Name: $($_.LocalizedName) Result: $($_.result.errorcode) Details: $($_.result.errortext)"
                        write-error "Name: $($_.LocalizedName) Result: $($_.result.errorcode) Details: $($_.result.errortext)"
                    }
                }
            }catch{}

            $erroText = "Import result: Job with ID: $importId failed at $ProcPercent percent complete."
            throw $erroText
        }
        else
        {
			#at this point we appear to have imported successfully 
            $managedsolution = $importManifest.Managed
            if($managedsolution -ne 1)
            {
                if($PublishChanges){
                    Write-Verbose "PublishChanges set, executing: Publish-CrmAllCustomization using the same connection."
                    Publish-CrmAllCustomization -conn $conn
                }
                else{
                    Write-Output "Import Complete, don't forget to publish customizations."
                }
            }
			else{
				#managed
                Write-Output "Import of managed solution complete."
			}
        }
    }
    catch
    {
        Write-Error $_.Exception
    }    
}

#InstallSampleDataToCrm    
function Add-CrmSampleData{
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml
    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false, Position=0)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn
    )
	$conn = VerifyCrmConnectionParam $conn 
    try
    {
        $result = $conn.InstallSampleDataToCrm()
		if($result -eq $null)
        {
            throw $conn.LastCrmException
        }
    }
    catch
    {
        throw $conn.LastCrmException
    }
    return $result
}

#IsSampleDataInstalled    
function Test-CrmSampleDataInstalled{
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml
    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn
    )

	$conn = VerifyCrmConnectionParam $conn 

    try
    {
        $result = $conn.IsSampleDataInstalled()
    }
    catch
    {
        throw $conn.LastCrmException
    }    

    return $result
}

#PublishEntity  
function Publish-CrmEntity{
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml
 [CmdletBinding()]
    PARAM( 
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1)]
        [string]$EntityLogicalName
    )

	$conn = VerifyCrmConnectionParam $conn  

    try
    {
        $result = $conn.PublishEntity($EntityLogicalName)
		if(!$result)
        {
            throw $conn.LastCrmException
        }
    }
    catch
    {
        throw $conn.LastCrmException
    }

    return $result
}

#ResetLocalMetadataCache  
function Remove-CrmEntityMetadataCache{
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml
 [CmdletBinding()]
    PARAM( 
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$false, Position=1)]
        [string]$EntityLogicalName
    )

	$conn = VerifyCrmConnectionParam $conn  

    if($EntityLogicalName -eq "")
    {
        $EntityLogicalName = $null
    }
    
    try
    {
        $result = $conn.ResetLocalMetadataCache($EntityLogicalName)
		if($result -eq $null)
        {
            throw $conn.LastCrmException
        }
    }
    catch
    {
        throw $conn.LastCrmException
    }    
}

#UninstallSampleDataFromCrm     
function Remove-CrmSampleData{
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml
    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn
    )

	$conn = VerifyCrmConnectionParam $conn 

    try
    {
        $result = $conn.UninstallSampleDataFromCrm()
    }
    catch
    {
        throw $conn.LastCrmException
    }

    return $result
}

#UpdateStateAndStatusForEntity 
function Set-CrmRecordState{
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml
    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1, ParameterSetName="CrmRecord", ValueFromPipeline=$true)]
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

    begin
    {
        $conn = VerifyCrmConnectionParam $conn
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
			$result = $conn.UpdateStateAndStatusForEntity($EntityLogicalName, $Id, $stateCode, $statusCode, [Guid]::Empty)
			if(!$result)
			{
				throw $conn.LastCrmException
			}
		}
		catch
		{
			throw $conn.LastCrmException
		}
	}
}


<#LEFT OFF HERE#> 
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml
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
 A connection to your CRM organization. Use $conn = Get-CrmConnection <Parameters> to generate it.

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
 
 This example assigns the security role to the team by using Id by omitting parameters names.
 When omitting parameter names, you do not provide $conn, cmdlets automatically finds it.

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

	$conn = VerifyCrmConnectionParam $conn

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
 A connection to your CRM organization. Use $conn = Get-CrmConnection <Parameters> to generate it.

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
 
 This example assigns the security role to the user by using Id by omitting parameters names.
 When omitting parameter names, you do not provide $conn, cmdlets automatically finds it.

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

	$conn = VerifyCrmConnectionParam $conn

    if($SecurityRoleRecord -eq $null -and $SecurityRoleId -eq "" -and $SecurityRoleName -eq "")
    {
        Write-Warning "You need to specify Security Role information"
        return
    }
    
    if($SecurityRoleName -ne "")
    {
        if($UserRecord -eq $null -or $UserRecord.businessunitid -eq $null)
        {
            $UserRecord = Get-CrmRecord -conn $conn -EntityLogicalName systemuser -Id $UserId -Fields businessunitid
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
 A connection to your CRM organization. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .PARAMETER UserId
 A record Id of User.

 .PARAMETER QueueId
 A record Id of Queue.
  
 .EXAMPLE
 Approve-CrmEmailAddress -conn $conn -UserId 00005a70-6317-e511-80da-c4346bc43d94

 This example approves email address for a user.

 .EXAMPLE
 Approve-CrmEmailAddress -UserId 00005a70-6317-e511-80da-c4346bc43d94
 
 This example approves email address for a user by omitting parameters names.
 When omitting parameter names, you do not provide $conn, cmdlets automatically finds it.

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

	$conn = VerifyCrmConnectionParam $conn

    if($UserId -ne "")
    {
        Set-CrmRecord -conn $conn -EntityLogicalName systemuser -Id $UserId -Fields @{"emailrouteraccessapproval"=(New-CrmOptionSetValue 1)}
    }
    else
    {
        Set-CrmRecord -conn $conn -EntityLogicalName queue -Id $QueueId -Fields @{"emailrouteraccessapproval"=(New-CrmOptionSetValue 1)}
    }
}

function Disable-CrmLanguagePack{

<#
 .SYNOPSIS
 Executes DeprovisionLanguageRequest Organization Request.

 .DESCRIPTION
 The Disable-CrmLanguagePack cmdlet lets you deprovision LanguagePack.

 .PARAMETER conn
 A connection to your CRM organization. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .PARAMETER LCID
 A Language ID.

 .EXAMPLE
 Disable-CrmLanguagePack -conn $conn -LCID 1041

 This example deprovisions Japanese Language Pack.

 .EXAMPLE
 Disable-CrmLanguagePack 1041
 
 This example deprovisions Japanese Language Pack by omitting parameters names.
 When omitting parameter names, you do not provide $conn, cmdlets automatically finds it.

#>

	[CmdletBinding()]
    PARAM( 
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1)]
        [Int]$LCID
    )

	$conn = VerifyCrmConnectionParam $conn  

    $request = New-Object Microsoft.Crm.Sdk.Messages.DeprovisionLanguageRequest
    $request.Language = $LCID
    
    try
    {
        $result = $conn.ExecuteCrmOrganizationRequest($request, $null)
		if($result -eq $null)
        {
            throw $conn.LastCrmException
        }
    }
    catch
    {
        throw $conn.LastCrmException
    }    
}

function Enable-CrmLanguagePack{

<#
 .SYNOPSIS
 Executes ProvisionLanguageRequest Organization Request.

 .DESCRIPTION
 The Enable-CrmLanguagePack cmdlet lets you provision LanguagePack. For OnPremise, you need to install corresponding Language Pack inadvance.

 .PARAMETER conn
 A connection to your CRM organization. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .PARAMETER LCID
 A Language ID.

 .EXAMPLE
 Enable-CrmLanguagePack -conn $conn -LCID 1041

 This example provisions Japanese Language Pack.

 .EXAMPLE
 Enable-CrmLanguagePack 1041
 
 This example provisions Japanese Language Pack by omitting parameters names.
 When omitting parameter names, you do not provide $conn, cmdlets automatically finds it.

#>

	[CmdletBinding()]
    PARAM( 
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1)]
        [Int]$LCID
    )

	$conn = VerifyCrmConnectionParam $conn

    $request = New-Object Microsoft.Crm.Sdk.Messages.ProvisionLanguageRequest
    $request.Language = $LCID
    
    try
    {
        $result = $conn.ExecuteCrmOrganizationRequest($request, $null)
		if($result -eq $null)
        {
            throw $conn.LastCrmException
        }
    }
    catch
    {
        throw $conn.LastCrmException
    }    
}

function Export-CrmApplicationRibbonXml {
<#
 .SYNOPSIS
 Retrieves the ribbon definition XML for application and saves it to a file on the file system.

 .DESCRIPTION
 The Export-CrmEntityRibbonXml cmdlet lets you retrieve ribbon definition XML for an Entity. 

 .PARAMETER conn
 A connection to your CRM organization. Use $conn = Get-CrmConnection <Parameters> to generate it.
	
 .PARAMETER RibbonFilePath
 The path to the desired output location. This should be a full file path, e.g.-c:\temp

 .EXAMPLE
 Export-CrmEntityRibbonXml -conn $conn -EntityLogicalName Account -Path d:\temp\accountribbon.xml
 
 This example gets the Ribbon XML for the account entity and outputs it to the specified file path.
#>
 
    [CmdletBinding()]
    PARAM( 
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$false, Position=1)][alias("Path")]		
        [string]$RibbonFilePath
    )

	$conn = VerifyCrmConnectionParam $conn
	
	$exportPath = if($RibbonFilePath -ne ""){Get-Item $RibbonFilePath} else {Get-Location}
	$exportFileName = "applicationRibbon.xml"
	$path = Join-Path $exportPath $exportFileName
	# Instantiate RetrieveEntityRibbonRequest
	$request = New-Object Microsoft.Crm.Sdk.Messages.RetrieveApplicationRibbonRequest
    
	try
    {
        $response = $conn.ExecuteCrmOrganizationRequest($request, $null)
        if ($response.CompressedApplicationRibbonXml -ne $null)
        {
			$ribbonXml = UnzipCrmRibbon -Data $response.CompressedApplicationRibbonXml
			
			Write-Verbose 'Saving ribbon file to path: $path'

			$ribbonXml.Save($path)

			Write-Verbose "Successfully wrote file"
			
			$result = New-Object PSObject
			Add-Member -InputObject $result -MemberType NoteProperty -Name "RetrieveApplicationRibbonRequest" -Value $response
			Add-Member -InputObject $result -MemberType NoteProperty -Name "RibbonFilePath" -Value $path
			return $result
        }

        #Should only get here if there was nothing returned.
        throw $conn.LastCrmException
    }
    catch
    {
	    throw $conn.LastCrmException
    }
}

function Export-CrmEntityRibbonXml {
<#
 .SYNOPSIS
 Retrieves the ribbon definition XML for an Entity and saves it to a file on the file system.

 .DESCRIPTION
 The Export-CrmEntityRibbonXml cmdlet lets you retrieve ribbon definition XML for an Entity. 

 .PARAMETER conn
 A connection to your CRM organization. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .PARAMETER EntityLogicalName
 A logicalname for Entity. e.g.- account, contact, lead, etc..

 .PARAMETER RibbonFilePath
 The path to the desired output location. This should be a full file path, e.g.-c:\temp

 .EXAMPLE
 Export-CrmEntityRibbonXml -conn $conn -EntityLogicalName account -RibbonFilePath c:\temp
 
 This example export the Ribbon XML for the account entity to c:\temp\accountRibbon.xml.
#>
 
    [CmdletBinding()]
    PARAM( 
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1)]
        [string]$EntityLogicalName,
        [parameter(Mandatory=$false, Position=2)][alias("Path")]		
        [string]$RibbonFilePath
    )

	$conn = VerifyCrmConnectionParam $conn
	
	$exportPath = if($RibbonFilePath -ne ""){Get-Item $RibbonFilePath} else {Get-Location}
	$exportFileName = $EntityLogicalName + "Ribbon.xml"
	$path = Join-Path $exportPath $exportFileName
	# Instantiate RetrieveEntityRibbonRequest
	$request = New-Object Microsoft.Crm.Sdk.Messages.RetrieveEntityRibbonRequest
    $request.EntityName = $EntityLogicalName
	$request.RibbonLocationFilter = [Microsoft.Crm.Sdk.Messages.RibbonLocationFilters]::All
    
	try
    {
        $response = $conn.ExecuteCrmOrganizationRequest($request, $null)
        if ($response.CompressedEntityXml -ne $null)
        {
			$ribbonXml = UnzipCrmRibbon -Data $response.CompressedEntityXml
			
			Write-Verbose 'Saving ribbon file to path: $path'

			$ribbonXml.Save($path)

			Write-Verbose "Successfully wrote file"
			
			$result = New-Object PSObject
			Add-Member -InputObject $result -MemberType NoteProperty -Name "RetrieveEntityRibbonResponse" -Value $response
			Add-Member -InputObject $result -MemberType NoteProperty -Name "RibbonFilePath" -Value $path
			return $result
        }

        #Should only get here if there was nothing returned.
        throw $conn.LastCrmException
    }
    catch
    {
	    throw $conn.LastCrmException
    }
}

function Export-CrmSolution{

<#
 .SYNOPSIS
 Exports a solution by Name from a CRM Organization.

 .DESCRIPTION
 The Export-CrmSolution cmdlet lets you export a solution file from CRM.

 .PARAMETER conn
 A connection to your CRM organization. Use $conn = Get-CrmConnection <Parameters> to generate it.

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
 Specify the parameter to export sales settings. Only available CRM 2015+

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

 This example exports "MySolution" solution as unmanaged with current path and default name by omitting $conn parameter.
 When omitting $conn parameter, cmdlets automatically finds it.

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

    $conn = VerifyCrmConnectionParam $conn

    try
    {
        $solutionRecords = (Get-CrmRecords -conn $conn -EntityLogicalName solution -FilterAttribute uniquename -FilterOperator "like" -FilterValue $SolutionName -Fields uniquename,publisherid,version )
        #if we can't find just one solution matching then ERROR
        if($solutionRecords.CrmRecords.Count -ne 1)
        {
            $friendlyName = $conn.ConnectedOrgFriendlyName.ToString()
            Write-Error "Solution with name `"$SolutionName`" in CRM Instance: `"$friendlyName`" not found!"
            break
        }
        #else PROCEED 
		$crmSolutionRecord = $solutionRecords.CrmRecords[0]
        $version = $crmSolutionRecord.version
		$solutionUniqueName = $crmSolutionRecord.uniquename

        write-verbose "Solution found with version# $version"
        $exportPath = if($SolutionFilePath -ne ""){Get-Item $SolutionFilePath} else {Get-Location}
        #if a filename is not given, then we'll default one to [solutionname]_[managed]_[version].zip
        if($SolutionZipFileName.Length -eq 0)
        {
            $version = $version.Replace('.','_')
            $managedFileName = if($Managed) {"_managed_"} else {"_unmanaged_"}
            $solutionZipFileName = "$solutionUniqueName$managedFileName$version.zip"
        }
        #now we should have the final path
        $path = Join-Path $exportPath $solutionZipFileName

        Write-Verbose "Solution path: $path"

        #create the export request then set all the properties
        $exportRequest = New-Object Microsoft.Crm.Sdk.Messages.ExportSolutionRequest
        $exportRequest.ExportAutoNumberingSettings            =$ExportAutoNumberingSettings 
        $exportRequest.ExportCalendarSettings                 =$ExportCalendarSettings
        $exportRequest.ExportCustomizationSettings            =$ExportCustomizationSettings
        $exportRequest.ExportEmailTrackingSettings            =$ExportEmailTrackingSettings
        $exportRequest.ExportGeneralSettings                  =$ExportGeneralSettings
        $exportRequest.ExportIsvConfig                        =$ExportIsvConfig
        $exportRequest.ExportMarketingSettings                =$ExportMarketingSettings
        $exportRequest.ExportOutlookSynchronizationSettings   =$ExportOutlookSynchronizationSettings
        $exportRequest.ExportRelationshipRoles                =$ExportRelationshipRoles
        $exportRequest.Managed                                =$Managed
        $exportRequest.SolutionName                           =$solutionUniqueName
        $exportRequest.TargetVersion                          =$TargetVersion 

		if($conn.ConnectedOrgVersion.Major -ge 7)
		{
			$exportRequest.ExportSales                            =$ExportSales
		}

        Write-Verbose 'ExportSolutionRequests may take several minutes to complete execution.'
        
        $response = [Microsoft.Crm.Sdk.Messages.ExportSolutionResponse]($conn.ExecuteCrmOrganizationRequest($exportRequest))

		Write-Verbose 'Using solution file to path: $path'

        [System.IO.File]::WriteAllBytes($path,$response.ExportSolutionFile)

        Write-Verbose "Successfully wrote file"
        $result = New-Object psObject

        Add-Member -InputObject $result -MemberType NoteProperty -Name "ExportSolutionResponse" -Value $response
        Add-Member -InputObject $result -MemberType NoteProperty -Name "SolutionPath" -Value $path

        return $result
    }
    catch
    {
        Write-Error $_.Exception
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
 A connection to your CRM organization. Use $conn = Get-CrmConnection <Parameters> to generate it.

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

 This example exports translation file of "MySolution" solution with current path and default name by omitting $conn parameter.
 When omitting $conn parameter, cmdlets automatically finds it.

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

	$conn = VerifyCrmConnectionParam $conn

    try
    {
        $solutionRecords = (Get-CrmRecords -conn $conn -EntityLogicalName solution -FilterAttribute uniquename -FilterOperator "like" -FilterValue $SolutionName -Fields publisherid,version )

        #if we can't find just one solution matching then ERROR
        if($solutionRecords.CrmRecords.Count -ne 1)
        {
            $friendlyName = $conn.ConnectedOrgFriendlyName.ToString()

            Write-Error "Solution with name `"$SolutionName`" in CRM Instance: `"$friendlyName`" not found!"
            break
        }
        #else PROCEED 

        $version = $solutionRecords.CrmRecords[0].version 

        write-verbose "Solution found with version# $version"
       
        $exportPath = if($TranslationFilePath -ne ""){ $TranslationFilePath } else { Get-Location }
        
        #if a filename is not given, then we'll default one to CrmTranslations_[solutionname]_[version].zip
        if($TranslationZipFileName.Length -eq 0)
        {
            $version = $version.Replace('.','_')
            $translationZipFileName = "CrmTranslations_$SolutionName`_$version.zip"
        }

        #now we should have the final path
        $path = Join-Path $exportPath $translationZipFileName

        Write-Verbose "Solution path: $path"

        #create the export translation request then set all the properties
        $exportRequest = New-Object Microsoft.Crm.Sdk.Messages.ExportTranslationRequest
        $exportRequest.SolutionName = $SolutionName

        Write-Verbose 'ExportTranslationRequest may take several minutes to complete execution.'
        
        $response = [Microsoft.Crm.Sdk.Messages.ExportTranslationResponse]($conn.ExecuteCrmOrganizationRequest($exportRequest))

        [System.IO.File]::WriteAllBytes($path,$response.ExportTranslationFile)

        Write-Verbose "Successfully wrote file: $path"
        $result = New-Object psObject

        Add-Member -InputObject $result -MemberType NoteProperty -Name "ExportTranslationResponse" -Value $response
        Add-Member -InputObject $result -MemberType NoteProperty -Name "SolutionTranslationPath" -Value $path

        return $result
    }
    catch
    {
        Write-Error $_.Exception
    }

    return $result
}

function Get-CrmAllLanguagePacks{

<#
 .SYNOPSIS
 Executes RetrieveAvailableLanguagesRequest Organization Request.

 .DESCRIPTION
 The Get-CrmAllLanguagePacks cmdlet lets you retrieve all available LanguagePack.

 .PARAMETER conn
 A connection to your CRM organization. Use $conn = Get-CrmConnection <Parameters> to generate it.

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

 This example etrieves all available Language Pack by omitting parameters names.
 When omitting parameter names, you do not provide $conn, cmdlets automatically finds it.

#>

	[CmdletBinding()]
    PARAM( 
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn
    )

	$conn = VerifyCrmConnectionParam $conn  

    $request = New-Object Microsoft.Crm.Sdk.Messages.RetrieveAvailableLanguagesRequest

    try
    {
        $response = $conn.ExecuteCrmOrganizationRequest($request, $null)
    }
    catch
    {
        throw $conn.LastCrmException
    }    

    return $response.LocaleIds
}

function Get-CrmEntityRecordCount{

<#
 .SYNOPSIS
 Retrieves total record count for an Entity.

 .DESCRIPTION
 The Get-CrmEntityRecordCount cmdlet lets you retrieve total record count for an Entity.

 .PARAMETER conn
 A connection to your CRM organization. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .PARAMETER EntityLogicalName
 A logicalname for an Entity to retrieve. i.e.)account, contact, lead, etc..

 .EXAMPLE
 Get-CrmEntityRecordCount -conn $conn -EntityLogicalName account
 10

 This example retrieves total record count for Account Entity.

 .EXAMPLE
 Get-CrmEntityRecordCount contact
 8413
 
 This example retrieves total record count for Contact Entity by omitting parameter names.
 When omitting parameter names, you do not provide $conn, cmdlets automatically finds it.

#>

    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1)]
        [string]$EntityLogicalName
    )

	$conn = VerifyCrmConnectionParam $conn
    
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
			if($result -eq $null)
			{
				throw $conn.LastCrmException
			}
        }
        catch
        {
            throw $conn.LastCrmException
        }
        
        $count += $result.EntityCollection.Entities.Count
        if($result.EntityCollection.MoreRecords)
        {
            $pageInfo.PageNumber += 1
            $pageInfo.PagingCookie = $result.EntityCollection.PagingCookie
        }
        else
        {
            break
        } 
    }
    
    return $count
}

function Get-CrmFailedWorkflows{

<#
 .SYNOPSIS
 Retrieves alert notifications from CRM organization.

 .DESCRIPTION
 The Get-CrmFailedWorkflows cmdlet lets you retrieve failed workflows from a CRM organization.

 .PARAMETER conn
 A connection to your CRM organization. Use $conn = Get-CrmConnection <Parameters> to generate it.
  
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
 When omitting parameter names, you do not provide $conn, cmdlets automatically finds it.

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
    
    $conn = VerifyCrmConnectionParam $conn
    
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

function Get-CrmLicenseSummary{

<#
 .SYNOPSIS
 Displays License assignment and AccessMode/CalType summery.

 .DESCRIPTION
 The Get-CrmLicenseSummery cmdlet lets you display License assignment and AccessMode/CalType summery.

 .PARAMETER conn
 A connection to your CRM organization. Use $conn = Get-CrmConnection <Parameters> to generate it.

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
 When omitting parameter names, you do not provide $conn, cmdlets automatically finds it.

#>

    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn
    )

	$conn = VerifyCrmConnectionParam $conn

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
 A connection to your CRM organization. Use $conn = Get-CrmConnection <Parameters> to generate it.
  
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
 When omitting parameter names, you do not provide $conn, cmdlets automatically finds it.

#>

    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn
    )

	$conn = VerifyCrmConnectionParam $conn
    
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
 A connection to your CRM organization. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .PARAMETER EntityLogicalName
 A logicalname for an Entity to retrieve. i.e.)account, contact, lead, etc..

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
 Get-CrmRecords -conn $conn -EntityLogicalName account -FilterAttribute name -FilterOperator "eq" -FilterValue "Adventure Works (sample)" -Fields name,accountnumber
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
 When omitting parameter names, you do not provide $conn, cmdlets automatically finds it.

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
    $conn = VerifyCrmConnectionParam $conn

    if($FilterOperator -and $FilterOperator.StartsWith("-"))
    {
        $FilterOperator = $FilterOperator.Remove(0, 1)
    }
    if( !($EntityLogicalName -cmatch "^[a-z]*$") )
    {
        $EntityLogicalName = $EntityLogicalName.ToLower()
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
        $primaryAttribute = $conn.GetEntityMetadata($EntityLogicalName.ToLower()).PrimaryIdAttribute
        $fetchAttributes = "<attribute name='{0}' />" -F $primaryAttribute
    }

    #if any of the values are missing, but they're not *ALL* missing 
    if( (!$FilterAttribute -OR !$FilterOperator -OR !$FilterValue) -AND ($FilterAttribute -Or $FilterOperator -Or $FilterValue) -And !($FilterAttribute -And $FilterOperator -And $FilterValue))
    {
        #TODO: convert this to a parameter set to avoid this extra logic
        Write-Error "One of the `$FilterAttribute `$FilterOperator `$FilterValue parameters is empty, to query all records exclude all filter parameters."
        return
    }
    
    if($FilterAttribute -and $FilterOperator -and $FilterValue)
    {
        # Escape XML charactors
        $FilterValue = [System.Security.SecurityElement]::Escape($FilterValue)
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
 A connection to your CRM organization. Use $conn = Get-CrmConnection <Parameters> to generate it.

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
 When omitting parameter names, you do not provide $conn, cmdlets automatically finds it.
  
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

	$conn = VerifyCrmConnectionParam $conn
    
    # Escape XML charactor
    $ViewName = [System.Security.SecurityElement]::Escape($ViewName)

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
 A connection to your CRM organization. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .PARAMETER EntityLogicalName
 A logicalname for an Entity to retrieve. i.e.)account, contact, lead, etc..

 .EXAMPLE
 Get-CrmRecordsCount -conn $conn -EntityLogicalName account
 5677

 This example retrieves total number of Account entity records.

 .EXAMPLE
 Get-CrmRecordsCount account
 5677

 This example retrieves total number of Account entity records by omitting parameter names.
 When omitting parameter names, you do not provide $conn, cmdlets automatically finds it.

#>

    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1)][alias("EntityName")]
        [string]$EntityLogicalName
    )
	$conn = VerifyCrmConnectionParam $conn        
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

function Get-CrmSdkMessageProcessingStepsForPluginAssembly{

<#
 .SYNOPSIS
 Retrieves all registered steps for Plugin.

 .DESCRIPTION
 The Get-CrmSdkMessageProcessingStepsForPluginAssembly cmdlet lets you retrieve all registered steps for Plugin.

 .PARAMETER conn
 A connection to your CRM organization. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .PARAMETER PluginAssemblyName
 A plugin assembly name.

 .PARAMETER OnlyCustomizable
 By specifying this swich returns only customizable steps. 
  
 .EXAMPLE
 Get-CrmSdkMessageProcessingStepsForPluginAssembly -conn $conn -PluginAssemblyName YourPluginAssemblyName
 
 This example retrieves all registered steps for the plugin assembly.

 .EXAMPLE
 Get-CrmSdkMessageProcessingStepsForPluginAssembly -conn $conn -PluginAssemblyName YourPluginAssemblyName -OnlyCustomizable
 
 This example retrieves all registered customizable steps for the plugin assembly.

 .EXAMPLE
 Get-CrmSdkMessageProcessingStepsForPluginAssembly -PluginAssemblyName YourPluginAssemblyName
 
 This example retrieves all registered steps for the plugin assembly by omitting parameter names.
 When omitting parameter names, you do not provide $conn, cmdlets automatically finds it.  
#>
    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1)]
        [string]$PluginAssemblyName,
        [parameter(Mandatory=$false, Position=2)]
        [switch]$OnlyCustomizable
    )

	$conn = VerifyCrmConnectionParam $conn
        
    if($OnlyCustomizable){ $isCustom = "<value>1</value>" } else { $isCustom = "<value>0</value><value>1</value>" }

    $fetch = @"
    <fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="false" no-lock="true">
      <entity name="sdkmessageprocessingstep">
        <all-attributes/>
        <link-entity name="sdkmessagefilter" from="sdkmessagefilterid" to="sdkmessagefilterid" visible="false" link-type="outer" alias="a1">
          <attribute name="secondaryobjecttypecode" />
          <attribute name="primaryobjecttypecode" />
        </link-entity>
        <link-entity name="plugintype" from="plugintypeid" to="plugintypeid" alias="ab">
          <filter type="and">
            <condition attribute="assemblyname" operator="eq" value="{0}" />            
          </filter>
        </link-entity>
        <filter type="and">
            <condition attribute="iscustomizable" operator="in">
                {1}
            </condition> 
        </filter>
      </entity>
    </fetch>
"@ -F $PluginAssemblyName, $isCustom
    
    $results = (Get-CrmRecordsByFetch -conn $conn -Fetch $fetch).CrmRecords
    
    return $results
}

function Get-CrmSiteMap{

<#
 .SYNOPSIS
 Retrieves CRM Organization's SiteMap information.

 .DESCRIPTION
 The Get-CrmSiteMap cmdlet lets you retrieve CRM Organization's SiteMap information.

 .PARAMETER conn
 A connection to your CRM organization. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .PARAMETER SiteXml
 Specify SiteXml switch to retrieve SiteMapXml xml data.

 .PARAMETER Areas
 Specify Areas switch to retrieve AreaIds list.
 
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
 When omitting parameter names, you do not provide $conn, cmdlets automatically finds it.

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

	$conn = VerifyCrmConnectionParam $conn
  
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
 A connection to your CRM organization. Use $conn = Get-CrmConnection <Parameters> to generate it.

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
 When omitting parameter names, you do not provide $conn, cmdlets automatically finds it.

 .EXAMPLE
 Get-CrmSystemSettings -ShowDisplayName
 Allow the showing tablet application notification bars in a browser. : Yes
 Presence Enabled                                                     : Yes
 Default Email Settings:Incoming Email                                : Server-Side Synchronization or Email Router
 Default Email Settings:Outgoing Email                                : Server-Side Synchronization or Email Router
 Default Email Settings:Appointments, Contacts, and Tasks             : Microsoft Dynamics CRM for Outlook
 ...

 This example retrieves CRM Organization's System Settings and show DisplayName for fields by omitting parameter names.
 When omitting parameter names, you do not provide $conn, cmdlets automatically finds it.

#>

    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$false)]
        [switch]$ShowDisplayName
    )

	$conn = VerifyCrmConnectionParam $conn
  
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
 A connection to your CRM organization. Use $conn = Get-CrmConnection <Parameters> to generate it.
  
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
 When omitting parameter names, you do not provide $conn, cmdlets automatically finds it.
  
#>

    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn
    )

	$conn = VerifyCrmConnectionParam $conn
    
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

function Get-CrmTraceAlerts{

<#
 .SYNOPSIS
 Retrieves alert notifications from CRM organization.

 .DESCRIPTION
 The Get-CrmTraceAlerts cmdlet lets you retrieve alert notifications from CRM organization.

 .PARAMETER conn
 A connection to your CRM organization. Use $conn = Get-CrmConnection <Parameters> to generate it.
  
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
 When omitting parameter names, you do not provide $conn, cmdlets automatically finds it.

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

	$conn = VerifyCrmConnectionParam $conn
    
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
 A connection to your CRM organization. Use $conn = Get-CrmConnection <Parameters> to generate it.

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
 When omitting parameter names, you do not provide $conn, cmdlets automatically finds it.

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

	$conn = VerifyCrmConnectionParam $conn
    
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
 Origin: Indicate where the privilege comes from. RoleName:Depth format
 PrincipalType: User or Team
 PrincipalName: User's fullname or Team's name
 BusinessUnitName: User's or Team's BusinessUnitName

 .PARAMETER conn
 A connection to your CRM organization. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .PARAMETER UserId
 An Id (guid) of CRM User.
   
 .EXAMPLE
 Get-CrmUserPrivileges -conn $conn -UserId f9d40920-7a43-4f51-9749-0549c4caf67d
 Depth            : Global
 PrivilegeId      : 59f49e9a-a621-4836-8a35-b44f1f7122fb
 PrivilegeName    : prvAppendToConvertRule
 Origin           : Marketing Manager:Global,System Administrator:Global
 PrincipalType    : User
 PrincipalName    : kenichiro nakamura
 BusinessUnitName : Contoso
 
 Depth            : Global
 PrivilegeId      : 368aff3b-95b1-45c1-bf73-01a7becdedc5
 PrivilegeName    : prvAppendToCustomerOpportunityRole
 Origin           : Salesperson:Global
 PrincipalType    : Team
 PrincipalName    : TeamA
 BusinessUnitName : Contoso 
 ...

 This example retrieves privileges assigned to the CRM User.

 Get-CrmUserPrivileges f9d40920-7a43-4f51-9749-0549c4caf67d
 Depth            : Global
 PrivilegeId      : 59f49e9a-a621-4836-8a35-b44f1f7122fb
 PrivilegeName    : prvAppendToConvertRule
 Origin           : Marketing Manager:Global,System Administrator:Global
 PrincipalType    : User
 PrincipalName    : kenichiro nakamura
 BusinessUnitName : Contoso
 
 Depth            : Global
 PrivilegeId      : 368aff3b-95b1-45c1-bf73-01a7becdedc5
 PrivilegeName    : prvAppendToCustomerOpportunityRole
 Origin           : Salesperson:Global
 PrincipalType    : Team
 PrincipalName    : TeamA
 BusinessUnitName : Contoso 
 ...

 This example retrieves privileges assigned to the CRM User by omitting parameter names.
 When omitting parameter names, you do not provide $conn, cmdlets automatically finds it.

#>

    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1)]
        [string]$UserId
    )

	$conn = VerifyCrmConnectionParam $conn
    
    # Get User Rolls including Team
    $roles = Get-CrmUserSecurityRoles -conn $conn -UserId $UserId -IncludeTeamRoles
    # Get User name
    $user = Get-CrmRecord -conn $conn -EntityLogicalName systemuser -Id $UserId -Fields fullname
    # Get all privilege records for PrivilegeName
	$privilegeRecords = (Get-CrmRecords -conn $conn -EntityLogicalName privilege -Fields name,privilegeid -WarningAction SilentlyContinue).CrmRecords
    # Create hash for performance reason
    $privileges = @{}
    $privilegeRecords | % {$privileges[$_.privilegeid] = $_.name} 

    # Create Result as hash for performance reason
    $results = @{}
    
    $isUserRoleInitialized = $false
	foreach($role in $roles | sort TeamName)
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
	        throw $conn.LastCrmException
	    }	    
	    
	    foreach($rolePrivilege in $rolePrivileges)
	    {
            # Create origin as "RoleName:Depth" format
            $origin = $role.RoleName + ":" + $rolePrivilege.Depth
            
            # If the role is assigned to a team, then add them separately.
            # For roles assign to the user, then accumulate them.
	        if($isUserRoleInitialized -and $results.Contains($rolePrivilege.PrivilegeId))
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
                if($role.TeamName -eq $null)
                {
                    $principalType = "User"
                    $principalName = $user.fullname
                    $key = $rolePrivilege.PrivilegeId
                }
                else
                {
                    $principalType = "Team"
                    $principalName = $role.TeamName
                    $key = $rolePrivilege.PrivilegeId.Guid + $origin
                }
	            Add-Member -InputObject $psobj -MemberType NoteProperty -Name "Depth" -Value $rolePrivilege.Depth
	            Add-Member -InputObject $psobj -MemberType NoteProperty -Name "PrivilegeId" -Value $rolePrivilege.PrivilegeId
	            Add-Member -InputObject $psobj -MemberType NoteProperty -Name "PrivilegeName" -Value $privileges[($rolePrivilege.PrivilegeId)]
	            Add-Member -InputObject $psobj -MemberType NoteProperty -Name "Origin" -Value $origin
                Add-Member -InputObject $psobj -MemberType NoteProperty -Name "PrincipalType" -Value $principalType
                Add-Member -InputObject $psobj -MemberType NoteProperty -Name "PrincipalName" -Value $principalName
                Add-Member -InputObject $psobj -MemberType NoteProperty -Name "BusinessUnitName" -Value $role.BusinessUnitName
	            $results[$key] = $psobj
	        }
	    }

        if($role.TeamName -eq $null) {$isUserRoleInitialized = $true} 
	}
    
    return $results.Values | sort principalName, PrivilegeName
}

function Get-CrmUserSecurityRoles{

<#
 .SYNOPSIS
 Retrieves Security Roles assigned to a CRM User.

 .DESCRIPTION
 The Get-CrmUserSecurityRoles cmdlet lets you retrieve Security Roles assigned to a CRM User.

 .PARAMETER conn
 A connection to your CRM organization. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .PARAMETER UserId
 An Id (guid) of CRM User.

 .PARAMETER IncludeTeamRoles
 When you specify the IncludeTeamRoles switch, Security Roles from teams are also retured.
   
 .EXAMPLE
 Get-CrmUserSecurityRoles -conn $conn -UserId f9d40920-7a43-4f51-9749-0549c4caf67d
 RoleId                               RoleName             TeamName BusinessUnitName 
 ------                               --------             -------- ---------------- 
 cac2d2c7-c6f7-e411-80de-c4346bc520c0 Marketing Manager    TeamA    Contoso
 57d5d2c7-c6f7-e411-80de-c4346bc520c0 Salesperson                   Contoso
 ...

 This example retrieves Security Roles assigned to the CRM User.

 .EXAMPLE
 Get-CrmUserSecurityRoles -conn $conn -UserId f9d40920-7a43-4f51-9749-0549c4caf67d -IncludeTeamRoles
 RoleId                               RoleName             TeamName BusinessUnitName 
 ------                               --------             -------- ---------------- 
 cac2d2c7-c6f7-e411-80de-c4346bc520c0 Marketing Manager    TeamA    Contoso
 57d5d2c7-c6f7-e411-80de-c4346bc520c0 Salesperson                   Contoso
 ...

 This example retrieves Security Roles assigned to the CRM User and Teams which the CRM User belongs to.

 .EXAMPLE
 Get-CrmUserSecurityRoles f9d40920-7a43-4f51-9749-0549c4caf67d -IncludeTeamRoles
 RoleId                               RoleName             TeamName BusinessUnitName 
 ------                               --------             -------- ---------------- 
 cac2d2c7-c6f7-e411-80de-c4346bc520c0 Marketing Manager    TeamA    Contoso
 57d5d2c7-c6f7-e411-80de-c4346bc520c0 Salesperson                   Contoso
 ...

 This example retrieves Security Roles assigned to the CRM User and Teams which the CRM User belongs to by omitting parameter names.
 When omitting parameter names, you do not provide $conn, cmdlets automatically finds it.

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

	$conn = VerifyCrmConnectionParam $conn
    
    $roles = New-Object System.Collections.Generic.List[PSObject]
    
    if($IncludeTeamRoles)
	{
		$fetch = @"
		<fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="true" no-lock="true">
		  <entity name="role">
		    <attribute name="name"/>
			<attribute name="roleid" />
		    <link-entity name="teamroles" from="roleid" to="roleid" visible="false" intersect="true">
		      <link-entity name="team" from="teamid" to="teamid" alias="team">
		      <attribute name="name"/>
              <attribute name="businessunitid"/>
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
		
		(Get-CrmRecordsByFetch -conn $conn -Fetch $fetch).CrmRecords | `
        select @{name="RoleId";expression={$_.roleid}}, @{name="RoleName";expression={$_.name}}, `
        @{name="TeamName";expression={$_.'team.name'}}, @{name="BusinessUnitName";expression={($_.'team.businessunitid').Name}} | `
        % {$roles.Add($_)}	
	}

	$fetch = @"
	<fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="true" no-lock="true">
	  <entity name="role">
	    <attribute name="name" />
	    <attribute name="roleid" />
	    <order attribute="name" descending="false" />
	    <link-entity name="systemuserroles" from="roleid" to="roleid" visible="false" intersect="true">
	      <link-entity name="systemuser" from="systemuserid" to="systemuserid" alias="user">
          <attribute name="businessunitid"/>
	        <filter type="and">
	          <condition attribute="systemuserid" operator="eq" value="{0}" />
	        </filter>
	      </link-entity>
	    </link-entity>
	  </entity>
	</fetch>
"@ -F $UserId
	
	(Get-CrmRecordsByFetch -conn $conn -Fetch $fetch).CrmRecords | `
    select @{name="RoleId";expression={$_.roleid}}, @{name="RoleName";expression={$_.name}}, `
    @{name="BusinessUnitName";expression={($_.'user.businessunitid').Name}}  | % { $roles.Add($_) }
	
	return $roles
}

function Get-CrmUserSettings{

<#
 .SYNOPSIS
 Retrieves CRM user's settings.

 .DESCRIPTION
 The Get-CrmUserSettings cmdlet lets you retrieve CRM user's settings.

 .PARAMETER conn
 A connection to your CRM organization. Use $conn = Get-CrmConnection <Parameters> to generate it.

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
 When omitting parameter names, you do not provide $conn, cmdlets automatically finds it.

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

	$conn = VerifyCrmConnectionParam $conn
  
    return Get-CrmRecord -conn $conn -EntityLogicalName usersettings -Id $UserId -Fields $Fields
}

function Import-CrmSolutionTranslation{

<#
 .SYNOPSIS
 Imports a translation to a solution.

 .DESCRIPTION
 The Import-CrmSolutionTranslation cmdlet lets you export a translation to a solution.

 .PARAMETER conn
 A connection to your CRM organization. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .PARAMETER TranslationFileName
 A file name and path of importing solution translation zip file.
 
 .PARAMETER PublishChanges
 Specify the parameter to publish all customizations.

 .EXAMPLE
 Import-CrmSolutionTranslation -conn $conn -TranslationZipFileName "C:\temp\CrmTranslations_MySolution_1_0_0_0.zip"

 This example imports translation file "CrmTranslations_MySolution_1_0_0_0".

 .EXAMPLE
 Import-CrmSolutionTranslation -conn $conn -TranslationZipFileName "C:\temp\CrmTranslations_MySolution_1_0_0_0.zip"

 This example imports translation file "CrmTranslations_MySolution_1_0_0_0" by omitting $conn parameter.
 When omitting $conn parameter, cmdlets automatically finds it.

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

    $conn = VerifyCrmConnectionParam $conn
	   
    try
    {
        $importId = [guid]::NewGuid()
        $translationFile = [System.IO.File]::ReadAllBytes($TranslationFileName)

        #create the import translation request then set all the properties
        $importRequest = New-Object Microsoft.Crm.Sdk.Messages.ImportTranslationRequest
        $importRequest.TranslationFile = $translationFile
        $importRequest.ImportJobId = $importId
        
        Write-Verbose 'ImportTranslationRequest may take several minutes to complete execution.'
        $response = [Microsoft.Crm.Sdk.Messages.ImportTranslationResponse]($conn.ExecuteCrmOrganizationRequest($importRequest))
                
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
    }
}

function Invoke-CrmWhoAmI{

<#
 .SYNOPSIS
 Executes WhoAmI Organization Request and returns current user's Id (guid), belonging BusinessUnit Id (guid) and CRM Organization Id (guid).

 .DESCRIPTION
 The Invoke-CrmWhoAmI cmdlet lets you execute WhoAmI request and obtain UserId, BusinessUnitId and OrganizationId.

 .PARAMETER conn
 A connection to your CRM organization. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .EXAMPLE
 Invoke-CrmWhoAmI -conn $conn

 This example executes WhoAmI organization request and returns current user's Id (guid), belonging BusinessUnit Id (guid) and CRM Organization Id (guid).

 .EXAMPLE
 Invoke-CrmWhoAmI
 
 This example executes WhoAmI by omitting parameters names.
 When omitting parameter names, you do not provide $conn, cmdlets automatically finds it.

#>

	[CmdletBinding()]
    PARAM( 
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn
    )
	$conn = VerifyCrmConnectionParam $conn

    $request = New-Object Microsoft.Crm.Sdk.Messages.WhoAmIRequest
    
    try
    {
        $result = $conn.ExecuteCrmOrganizationRequest($request, $null)
    }
    catch
    {
        throw $conn.LastCrmException
    }    

    return $result
}

function Publish-CrmAllCustomization{

<#
 .SYNOPSIS
 Publishes all customizations for a CRM Organization.

 .DESCRIPTION
 The Publish-CrmAllCustomization cmdlet lets you publish all customizations.

 .PARAMETER conn
 A connection to your CRM organization. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .EXAMPLE
 Publish-CrmAllCustomization -conn $conn

 This example publishes all customizations.

 .EXAMPLE
 Publish-CrmAllCustomization
 
 This example publishes all customizations by omitting parameters names.
 When omitting parameter names, you do not provide $conn, cmdlets automatically finds it.

#>

 [CmdletBinding()]
    PARAM( 
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn
    )

	$conn = VerifyCrmConnectionParam $conn  

    $request = New-Object Microsoft.Crm.Sdk.Messages.PublishAllXmlRequest
    
    try
    {
        $response = $conn.ExecuteCrmOrganizationRequest($request, $null)
    }
    catch
    {
        throw $conn.LastCrmException
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
 A connection to your CRM organization. Use $conn = Get-CrmConnection <Parameters> to generate it.

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
 
 This example removes a security role to a team by using Id by omitting parameters names.
 When omitting parameter names, you do not provide $conn, cmdlets automatically finds it.

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

	$conn = VerifyCrmConnectionParam $conn

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
 A connection to your CRM organization. Use $conn = Get-CrmConnection <Parameters> to generate it.

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
 
 This example removes a security role to a user by using Id by omitting parameters names.
 When omitting parameter names, you do not provide $conn, cmdlets automatically finds it.

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

	$conn = VerifyCrmConnectionParam $conn

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
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml
 [CmdletBinding()]
    PARAM( 
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1)]
        [guid]$UserId
    )

	$conn = VerifyCrmConnectionParam $conn

    $request = New-Object 'Microsoft.Crm.Sdk.Messages.RemoveParentRequest'
    $target = New-CrmEntityReference systemuser $UserId
    $request.Target = $target
    
    try
    {
        $result = $conn.ExecuteCrmOrganizationRequest($request, $null)
		if($result -eq $null)
        {
            throw $conn.LastCrmException
        }
    }
    catch
    {
        throw $conn.LastCrmException
    } 
}

function Set-CrmConnectionCallerId{
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml
 [CmdletBinding()]
    PARAM( 
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, position=1)][Alias("UserId")]
        [guid]$CallerId
    )

	$conn = VerifyCrmConnectionParam $conn

    # We may need to check if the CallerId exists and enabled.
    $conn.OrganizationServiceProxy.CallerId = $CallerId
}

function Set-CrmConnectionTimeout{
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml
 [CmdletBinding()]
    PARAM( 
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$false, position=1)]
        [Int64]$TimeoutInSeconds,
        [parameter(Mandatory=$false, position=1)]
        [switch]$SetDefault
    )

	$conn = VerifyCrmConnectionParam $conn

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
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml
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

	$conn = VerifyCrmConnectionParam $conn
    
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
            $emailSettings.EmailSettings.IncomingEmailDeliveryMethod = [string]$defaultEmailSettings["IncomingEmailDeliveryMethod"]
        }
        if($defaultEmailSettings.ContainsKey("OutgoingEmailDeliveryMethod"))
        {
            $emailSettings.EmailSettings.OutgoingEmailDeliveryMethod = [string]$defaultEmailSettings["OutgoingEmailDeliveryMethod"]
        }
        if($defaultEmailSettings.ContainsKey("ACTDeliveryMethod"))
        {
            $emailSettings.EmailSettings.ACTDeliveryMethod = [string]$defaultEmailSettings["ACTDeliveryMethod"]
        }

        $updateFields.Add("defaultemailsettings",$emailSettings.OuterXml)
    }

    Set-CrmRecord -conn $conn -EntityLogicalName organization -Id $recordid -Fields $updateFields
}

function Set-CrmUserBusinessUnit{
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml
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

	$conn = VerifyCrmConnectionParam $conn

    $ReassignPrincipal = New-CrmEntityReference -EntityLogicalName systemuser -Id $ReassignUserId

    $request = New-Object 'Microsoft.Crm.Sdk.Messages.SetBusinessSystemUserRequest'
    $request.BusinessId = $BusinessUnitId
    $request.UserId = $UserId
    $request.ReassignPrincipal = $ReassignPrincipal

    try
    {
        $result = $conn.ExecuteCrmOrganizationRequest($request, $null)
		if($result -eq $null)
        {
            throw $conn.LastCrmException
        }
    }
    catch
    {
        throw $conn.LastCrmException
    }
}

function Set-CrmUserMailbox {
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml
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
	$conn = VerifyCrmConnectionParam $conn
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
        if($parameter.Key -in ("UserId", "ApplyDefaultEmailSettings", "conn"))
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
<fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="false" no-lock="true">
  <entity name="mailbox">
    <attribute name="mailboxid" />
    <filter type="and">
      <condition attribute="regardingobjectid" operator="eq" value="{$UserId}" />
    </filter>
  </entity>
</fetch>
"@
    $Id = (Get-CrmRecordsByFetch -conn $conn -Fetch $fetch).CrmRecords[0].MailboxId
    Set-CrmRecord -conn $conn -EntityLogicalName mailbox -Id $Id -Fields $updateFields
}

<############SKIPPED FROM HELP XML####################>
function Set-CrmUserManager{

<#
.SYNOPSIS
Sets CRM user's manager.

.DESCRIPTION
The Set-CrmUserManager lets you set CRM user's manager.

.PARAMETER conn
A connection to your CRM organization. Use $conn = Get-CrmConnection <Parameters> to generate it.

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
 
This example sets a manager to a CRM User and keeps its child users by omitting parameters names.
When omitting parameter names, you do not provide $conn, cmdlets automatically finds it.

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

	$conn = VerifyCrmConnectionParam $conn

    $request = New-Object 'Microsoft.Crm.Sdk.Messages.SetParentSystemUserRequest'
    $request.ParentId = $ManagerId
    $request.UserId = $UserId
    $request.KeepChildUsers = $KeepChildUsers

    try
    {
        $result = $conn.ExecuteCrmOrganizationRequest($request, $null)
		if($result -eq $null)
        {
            throw $conn.LastCrmException
        }
    }
    catch
    {
        throw $conn.LastCrmException
    } 
}

<############SKIPPED FROM HELP XML####################>
function Set-CrmUserSettings{

<#
 .SYNOPSIS
 Update CRM user's settings.

 .DESCRIPTION
 The Set-CrmUserSettings cmdlet lets you update CRM user's settings.

 .PARAMETER conn
 A connection to your CRM organization. Use $conn = Get-CrmConnection <Parameters> to generate it.

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
 When omitting parameter names, you do not provide $conn, cmdlets automatically finds it.

#>

    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1)]
        [Alias("UserSettingsRecord")]
        [PSObject]$CrmRecord
    )

	$conn = VerifyCrmConnectionParam $conn
    
    try
    {
        $result = Set-CrmRecord -conn $conn -CrmRecord $CrmRecord
    }
    catch
    {
        throw
    }    
}

<############SKIPPED FROM HELP XML####################>
function Upsert-CrmRecord{

<#
 .SYNOPSIS
 Creates or Updates CRM record by specifying field name/value set.

 .DESCRIPTION
 The Upsert-CrmRecord cmdlet lets you create new CRM record if no ID match, otherwise update existing one. 
 Use @{"field logical name"="value"} syntax to specify Fields, and make sure you specify correct type of value for the field.
 This is a good example how to use OrganizationService requests which is not exposed to CrmServiceClient yet.

 You can use Get-CrmEntityAttributeMetadata cmdlet and check AttributeType to see the field type. In addition, for CRM specific types, you can use New-CrmMoney, New-CrmOptionSetValue or New-CrmEntityReference cmdlets.

 .PARAMETER conn
 A connection to your CRM organization. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .PARAMETER EntityLogicalName
 A logicalname for an Entity to create. i.e.)account, contact, lead, etc..

 .PARAMETER Id
 An Id of the record

 .PARAMETER Fields
 A List of field name/value pair. Use @{"field logical name"="value"} syntax to create Fields, and make sure you specify correct type of value for the field. 
 You can use Get-CrmEntityAttributeMetadata cmdlet and check AttributeType to see the field type. In addition, for CRM specific types, you can use New-CrmMoney, New-CrmOptionSetValue or New-CrmEntityReference cmdlets.

 .EXAMPLE
 Upsert-CrmRecord -conn $conn -EntityLogicalName account -Id 57bd1c45-2b17-e511-80dc-c4346bc4fc6c -Fields @{"name"="account name";"industrycode"=New-CrmOptionSetValue -Value 1}
 
 This example updates record with id of 57bd1c45-2b17-e511-80dc-c4346bc4fc6c. If no matching record exists, then create an record by 
 assigning the id as it's id.

 .EXAMPLE
 Upsert-CrmRecord account 57bd1c45-2b17-e511-80dc-c4346bc4fc6c @{"name"="account name";"industrycode"=New-CrmOptionSetValue -Value 1}
 
 This example updates record with id of 57bd1c45-2b17-e511-80dc-c4346bc4fc6c. If no matching record exists, then create an record by 
 assigning the id as it's id by omming parameter names. When omitting parameter names, you do not provide $conn, cmdlets automatically finds it.
  
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
        [hashtable]$Fields
    )

	$conn = VerifyCrmConnectionParam $conn
		
	# Check CRM Organization version. Upsert is supported CRM 2015 Update 1 or later
	if($conn.ConnectedOrgVersion -gt (New-Object Version -ArgumentList 8.5))
	{
		# Instantiate Entity object
		$entity = New-Object Microsoft.Xrm.Sdk.Entity -ArgumentList $EntityLogicalName
		$entity.Id = $Id
					
		# Set attributes
		foreach($field in $Fields.GetEnumerator())
		{
			$entity[$field.Name] = $field.Value
		}

		try
		{
			# Instantiate UpsertRequest and set Target
			$request = New-Object Microsoft.Xrm.Sdk.Messages.UpsertRequest
			$request.Target = $entity
			$result = $conn.ExecuteCrmOrganizationRequest($request, $null)
		}
		catch
		{
		    throw $conn.LastCrmException        
		}	
		
		# If something went wrong, then return error.
		if($result -eq $null)
        {
            throw $conn.LastCrmException
        }	
	}
	else
	{
		# Get Primary Id attribute for the entity
		$primaryIdAttribute = (Get-CrmEntityMetadata -conn $conn -EntityLogicalName account).PrimaryIdAttribute
			
		# Get record by using Id first
		$record = Get-CrmRecord -conn $conn -EntityLogicalName $EntityLogicalName -Id $Id -Fields $primaryIdAttribute -WarningAction SilentlyContinue

		# If record exists, then do update, otherwise create new record
		if($record.logicalname -ne $null)
		{
			$result = Set-CrmRecord -conn $conn -EntityLogicalName $EntityLogicalName -Fields $Fields -Id $Id
			if($result -eq $null)
			{
			    throw $conn.LastCrmException
			}	
		}
		else
		{
			# Add the Id field if not included
			if(!$Fields.Contains($primaryIdAttribute))
			{
				$Fields.Add($primaryIdAttribute, $Id)
			}
			$result = New-CrmRecord -conn $conn -EntityLogicalName $EntityLogicalName -Fields $Fields
			
			if(!$result)
			{
			    throw $conn.LastCrmException
			}
		}
	}
  
    return $result
}

### Get CRM Types object ###
function New-CrmMoney{
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml
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
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml
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
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml
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
function Test-CrmViewPerformance{
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml
    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1, ParameterSetName="CrmRecord")][Alias("CrmRecord")]
        [PSObject]$View,
        [parameter(Mandatory=$true, Position=1, ParameterSetName="Id")][Alias("Id")]
        [guid]$ViewId,
        [parameter(Mandatory=$true, Position=1, ParameterSetName="Name")]
        [string]$ViewName,
        [parameter(Mandatory=$false)]
        [switch]$RunAsViewOwner,
        [parameter(Mandatory=$false)]
        [guid]$RunAs,
        [parameter(Mandatory=$false)]
        [switch]$IsUserView       
    )
    
	$conn = VerifyCrmConnectionParam $conn        
 
    if($IsUserView)
    { 
		Write-Verbose "querying userquery"
        $logicalName = "userquery"
        $fields = "name,fetchxml,layoutxml,returnedtypecode,ownerid".Split(",");      
    } 
    else
    {
		Write-Verbose "querying savedquery"
        $logicalName = "savedquery"
        $fields = "name,fetchxml,layoutxml,returnedtypecode".Split(",");
    }
    try
    {
        if($ViewId -ne $null)
        {        
            $View = Get-CrmRecord -conn $conn -EntityLogicalName $logicalName -Id $viewId -Fields $fields
        }
        elseif($viewName -ne "")
        {
            $views = Get-CrmRecords -conn $conn -EntityLogicalName $logicalName -FilterAttribute name -FilterOperator eq -FilterValue $viewName -Fields $fields
            if($views.CrmRecords.Count -eq 0) 
			{ 
				return 
			} 
			else 
			{ 
				$view = $views.CrmRecords[0]
			}
        }
		else{
			throw "ViewID or ViewName is null, input a valid view name or View ID."
		}
        # if the view has ownerid, then its User Defined View
        if($View.ownerid -ne $null)
        {
			if($RunAsViewOwner)
            {
                Set-CrmConnectionCallerId -conn $conn -CallerId $view.ownerid_property.Value.Id
            }
            elseif($RunAs -ne $null)
            {
                Set-CrmConnectionCallerId -conn $conn -CallerId $RunAs
            }
            #else
            #{
            #    Set-CrmConnectionCallerId -conn $conn -CallerId (Get-MyCrmUserId -conn $conn)
            #}
        
            # Get all records by using Fetch
            Test-XrmTimerStart
            $records = Get-CrmRecordsByFetch -conn $conn -Fetch $View.fetchxml -AllRows -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
            $perf = Test-XrmTimerStop
            $owner = $View.ownerid
            $totalCount = $records.Count           
        }
        else
        {            
            if($RunAs -ne $null)
            {
                Set-CrmConnectionCallerId -conn $conn -CallerId $RunAs                
            }
            #else
            #{
            #    Set-CrmConnectionCallerId -conn $conn -CallerId (Get-MyCrmUserId -conn $conn)
            #}
			# Get all records by using Fetch
            Test-XrmTimerStart
            $records = Get-CrmRecordsByFetch -conn $conn -Fetch $View.fetchxml -AllRows -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
            $perf = Test-XrmTimerStop
            $owner = "System"
            $totalCount = $records.Count
        }
        
		# Create result set
        $psobj = New-Object -TypeName System.Management.Automation.PSObject
	    Add-Member -InputObject $psobj -MemberType NoteProperty -Name "ViewName" -Value $View.name 
	    Add-Member -InputObject $psobj -MemberType NoteProperty -Name "FetchXml" -Value $View.fetchxml 
	    Add-Member -InputObject $psobj -MemberType NoteProperty -Name "Entity" -Value $View.returnedtypecode
	    Add-Member -InputObject $psobj -MemberType NoteProperty -Name "Columns" -Value ([xml]$view.layoutxml).grid.row.cell.Count
	    Add-Member -InputObject $psobj -MemberType NoteProperty -Name "LayoutXml" -Value $view.layoutxml
	    Add-Member -InputObject $psobj -MemberType NoteProperty -Name "TotalRecords" -Value $totalCount
	    Add-Member -InputObject $psobj -MemberType NoteProperty -Name "Owner" -Value $owner
        Add-Member -InputObject $psobj -MemberType NoteProperty -Name "Performance" -Value $perf
		
		#before returning always set connection caller id back to ourself: 
        if($RunAs -or $RunAsViewOwner){
			Write-verobse "Setting connection caller id back to current user"
			Set-CrmConnectionCallerId -conn $conn -CallerId $RunAs                
		}

        return $psobj
    }
    catch
    {
        return
    }
}

function Test-XrmTimerStart{
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml
    $script:crmtimer = New-Object -TypeName 'System.Diagnostics.Stopwatch'
    $script:crmtimer.Start()
}

function Test-XrmTimerStop{
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml
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

### Internal Helpers 
function Coalesce {
	foreach($i in $args){
		if($i -ne $null){
			return $i
		}
	}
}
function VerifyCrmConnectionParam {
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
            break
        }
        else
        {
            $conn = $connobj.Value
        }
    }
	return $conn
}
## Taken from CRM SDK sample code
## https://msdn.microsoft.com/en-us/library/microsoft.crm.sdk.messages.retrieveentityribbonresponse.compressedentityxml.aspx
function UnzipCrmRibbon {
    PARAM( 
        [parameter(Mandatory=$true)]
        [Byte[]]$Data
    )

    $memStream = New-Object System.IO.MemoryStream

    $memStream.Write($Data, 0, $Data.Length)
    $package = [System.IO.Packaging.ZipPackage]::Open($memStream, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read)
    $part = $package.GetPart([System.Uri]::new('/RibbonXml.xml', [System.UriKind]::Relative))

    try
    {
        $strm = $part.GetStream()
        $reader = [System.Xml.XmlReader]::Create($strm)

        $xmlDoc = New-Object System.Xml.XmlDocument
        $xmlDoc.Load($reader)

        return $xmlDoc
    }
    finally
    {
        if ($strm -ne $null)
        {
            $strm.Dispose()
            $strm = $null
        }
        if ($reader -ne $null)
        {
            $reader.Dispose()
            $reader = $null
        }
    }
}