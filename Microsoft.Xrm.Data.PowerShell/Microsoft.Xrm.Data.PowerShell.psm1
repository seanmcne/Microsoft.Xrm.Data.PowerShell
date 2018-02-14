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

function Connect-CrmOnlineDiscovery{
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml
    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false)]
        [PSCredential]$Credential, 
        [Parameter(Mandatory=$false)]
        [switch]$InteractiveMode
    )
    AddTls12Support #make sure tls12 is enabled 
    if($InteractiveMode)
    {
        $global:conn = Get-CrmConnection -InteractiveMode -Verbose
        
        Write-Verbose "You are now connected and may run any of the CRM Commands."
        return $global:conn 
    }
    else
    {
        $onlineType = "Office365"
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
        [parameter(Position=1, Mandatory=$true)]
        [PSCredential]$Credential, 
        [Parameter(Position=2,Mandatory=$true)]
        [ValidatePattern('([\w-]+).crm([0-9]*).dynamics.com')]
        [string]$ServerUrl, 
		[Parameter(Position=3,Mandatory=$false)]
        [switch]$ForceDiscovery,
		[Parameter(Position=4,Mandatory=$false)]
        [switch]$ForceOAuth, 
		[Parameter(Position=5,Mandatory=$false)]
        [ValidateScript({
            try {
                [System.Guid]::Parse($_) | Out-Null
                $true
            } catch {
                $false
            }
        })]
        [string]$OAuthClientId,
		[Parameter(Position=6,Mandatory=$false)]
        [string]$OAuthRedirectUri
    )
    AddTls12Support #make sure tls12 is enabled 
	if($ServerUrl.StartsWith("https://","CurrentCultureIgnoreCase") -ne $true){
		Write-Verbose "ServerUrl is missing https, fixing URL: https://$ServerUrl"
		$ServerUrl = "https://" + $ServerUrl
	}
	Write-Verbose "Connecting to: $ServerUrl"
    $cs = "RequireNewInstance=True"
	$cs+= ";Username=$($Credential.UserName)"
	$cs+= ";Password=$($Credential.GetNetworkCredential().Password)"
	$cs+= ";Url=$ServerUrl"
	
	#Default to Office365 Auth, allow oAuth to be used
	if(!$OAuthClientId -and !$ForceOAuth){
		Write-Verbose "Using AuthType=Office365"
		$cs += ";AuthType=Office365"
	}
	else{
		Write-Verbose "Params -> ForceOAuth: {$ForceOAuth} ClientId: {$OAuthClientId} RedirectUri: {$OAuthRedirectUri}"
		#use the clientid if provided, else use a provided clientid 
		if($OAuthClientId){
			Write-Verbose "Using provide "
			$cs += ";AuthType=OAuth;ClientId=$OAuthClientId"
			if($OAuthRedirectUri){
				$cs += ";redirecturi=$OAuthRedirectUri"
			}
		}
		else{
			$cs+=";AuthType=OAuth;ClientId=2ad88395-b77d-4561-9441-d0e40824f9bc"
			$cs+=";redirecturi=app://5d3e90d6-aa8e-48a8-8f2c-58b45cc67315"
		}
	}
	#disable the discovery check by default 
	if($ForceDiscovery){ 
		Write-Verbose "ForceDiscovery: SkipDiscovery=False"
		$cs+=";SkipDiscovery=False" 
	}
	else{ 
		Write-Verbose "Default: SkipDiscovery=True"
		$cs+=";SkipDiscovery=True" 
	}
    try
    {
		if(!$cs -or $cs.Length -eq 0){
			throw "Cannot create the CrmServiceClient, the connection string is null"
		}
        $global:conn = New-Object Microsoft.Xrm.Tooling.Connector.CrmServiceClient -ArgumentList $cs
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
    AddTls12Support #make sure tls12 is enabled 
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
        [parameter(Mandatory=$true, Position=1, ParameterSetName="NameAndFields")]
        [string]$EntityLogicalName,
        [parameter(Mandatory=$true, Position=2, ParameterSetName="NameAndFields")]
        [hashtable]$Fields,
        [parameter(Mandatory=$true, Position=1, ParameterSetName="CrmRecord")]
        [PSObject]$CrmRecord,
	[parameter(Mandatory=$false, Position=2, ParameterSetName="CrmRecord")]
        [switch]$PreserveCrmRecordId
    )

	$conn = VerifyCrmConnectionParam $conn

    $newfields = New-Object 'System.Collections.Generic.Dictionary[[String], [Microsoft.Xrm.Tooling.Connector.CrmDataTypeWrapper]]'
    
    if($CrmRecord -ne $null)
    {
        $EntityLogicalName = $CrmRecord.ReturnProperty_EntityName
        $atts = Get-CrmEntityAttributes -conn $conn -EntityLogicalName $EntityLogicalName
        foreach($crmFieldKey in ($CrmRecord | Get-Member -MemberType NoteProperty).Name)
        {
            if($crmFieldKey.EndsWith("_Property"))
            {
                if($CrmRecord.ReturnProperty_Id -eq $CrmRecord.$crmFieldKey.Value -and !$PreserveCrmRecordId)
                {
                    continue;
                }               
                elseif(($atts | ? logicalname -eq $CrmRecord.$crmFieldKey.Key).IsValidForCreate)
                {
                    # Some fields cannot be created even though it is set as IsValidForCreate
                    if($CrmRecord.$crmFieldKey.Key.Contains("addressid"))
                    {
                        continue;
                    }
                    else
                    {
                        $newfield = New-Object -TypeName 'Microsoft.Xrm.Tooling.Connector.CrmDataTypeWrapper'
            
                        $newfield.Type = MapFieldTypeByFieldValue -Value $CrmRecord.$crmFieldKey.Value                 
                        $newfield.Value = $CrmRecord.$crmFieldKey.Value
                        $newfields.Add($CrmRecord.$crmFieldKey.Key, $newfield)
                    }
                }
            }
        }  
    }
    else
    {
        foreach($field in $Fields.GetEnumerator())
        {  
            $newfield = New-Object -TypeName 'Microsoft.Xrm.Tooling.Connector.CrmDataTypeWrapper'
            
            $newfield.Type = MapFieldTypeByFieldValue -Value $field.Value
            
            $newfield.Value = $field.Value
            $newfields.Add($field.Key, $newfield)
        }
    }
    try
    {        
        $result = $conn.CreateNewRecord($EntityLogicalName, $newfields, $null, $false, [Guid]::Empty)
        if(!$result -or $result -eq [System.Guid]::Empty)
        {
            throw LastCrmConnectorException($conn)
        }
    }
    catch
    {
        throw LastCrmConnectorException($conn)
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
        [string[]]$Fields,
        [parameter(Mandatory=$false, Position=4)]
        [switch]$IncludeNullValue
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
        throw LastCrmConnectorException($conn)        
    }    
    
    if($record -eq $null)
    {        
        throw LastCrmConnectorException($conn)
    }
        
    $psobj = New-Object -TypeName System.Management.Automation.PSObject
    $meta = Get-CrmEntityMetadata -conn $conn -EntityLogicalName $EntityLogicalName -EntityFilters Attributes

    if($IncludeNullValue)
    {
        if($Fields -eq "*")
        {
            # Add all fields first
            foreach($attName in $meta.Attributes | ? IsValidForRead -eq $true | select LogicalName | sort LogicalName)
            {
                Add-Member -InputObject $psobj -MemberType NoteProperty -Name $attName.LogicalName -Value $null
				Add-Member -InputObject $psobj -MemberType NoteProperty -Name ($attName.LogicalName + "_Property") -Value $null
			}
        }
        else
        {
            foreach($attName in $Fields)
            {
                Add-Member -InputObject $psobj -MemberType NoteProperty -Name $attName -Value $null
				Add-Member -InputObject $psobj -MemberType NoteProperty -Name ($attName + "_Property") -Value $null
            }
        }
    }
        
    foreach($att in $record.GetEnumerator())
    {       
	    $keyName = $att.Key
	    
	    if(!($psobj | gm).Name.Contains($keyName))
        {
            Add-Member -InputObject $psobj -MemberType NoteProperty -Name $keyName -Value $null
        }

	    if($att.Value -is [Microsoft.Xrm.Sdk.EntityReference])
        {
	        $psobj.($keyName) = $att.Value.Name
	    }
	    elseif($att.Value -is [Microsoft.Xrm.Sdk.AliasedValue])
        {
	    	$psobj.($keyName) = $att.Value.Value
	    }
	    else
        {
	    	$psobj.($keyName) = $att.Value
	    }                
    }   

    Add-Member -InputObject $psobj -MemberType NoteProperty -Name "original" -Value $record
    Add-Member -InputObject $psobj -MemberType NoteProperty -Name "logicalname" -Value $EntityLogicalName
    Add-Member -InputObject $psobj -MemberType NoteProperty -Name "EntityReference" -Value (New-CrmEntityReference -EntityLogicalName $EntityLogicalName -Id $Id)
    # Add same additional fields to match Get-CrmRecords functions
    Add-Member -InputObject $psobj -MemberType NoteProperty -Name "ReturnProperty_EntityName" -Value $EntityLogicalName
    Add-Member -InputObject $psobj -MemberType NoteProperty -Name "ReturnProperty_Id" -Value $record.($meta.PrimaryIdAttribute)
	
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
        [hashtable]$Fields,
		[parameter(Mandatory=$false)]
        [switch]$Upsert,
        [parameter(Mandatory=$false)]
        [AllowNull()]
        [AllowEmptyString()]
        [string]$PrimaryKeyField
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
    
    # 'PrimaryKeyField' is an options parameter and is used for custom activity entities
    if(-not [string]::IsNullOrEmpty($PrimaryKeyField)) 
    {
         $primaryKeyField = $PrimaryKeyField
    }
    else
    {
        $primaryKeyField = GuessPrimaryKeyField -EntityLogicalName $entityLogicalName
    }

    # If upsert specified
    if($Upsert)
    {
        $retrieveFields = New-Object System.Collections.Generic.List[string]
        if($CrmRecord -ne $null)
        {
            # when CrmRecord passed, assume this comes from other system.
            $id = $CrmRecord.$primaryKeyField
            foreach($crmFieldKey in ($CrmRecord | Get-Member -MemberType NoteProperty).Name)
            {
                if($crmFieldKey.EndsWith("_Property"))
                {
                    $retrieveFields.Add(($CrmRecord.$crmFieldKey).Key)
                }
                elseif(($crmFieldKey -eq "original") -or ($crmFieldKey -eq "logicalname") -or ($crmFieldKey -eq "EntityReference")`
                  -or ($crmFieldKey -like "ReturnProperty_*"))
                {
                    continue
                }
                else
                {
                    # to have original value, rather than formatted value, replace the value from original record.
                    $CrmRecord.$crmFieldKey = $CrmRecord.original[$crmFieldKey+"_Property"].Value
                }
            }            
        }
        else
        {
            foreach($crmFieldKey in $Fields.Keys)
            {
                $retrieveFields.Add($crmFieldKey)
            }           
        }

        $existingRecord = Get-CrmRecord -conn $conn -EntityLogicalName $entityLogicalName -Id $id -Fields $retrieveFields.ToArray() -ErrorAction SilentlyContinue

        if($existingRecord.original -eq $null)
        {
            if($CrmRecord -ne $null)
            {
                $Fields = @{}
                foreach($crmFieldKey in ($CrmRecord | Get-Member -MemberType NoteProperty).Name)
                {
                    if($crmFieldKey.EndsWith("_Property"))
                    {
                        $Fields.Add(($CrmRecord.$crmFieldKey).Key, ($CrmRecord.$crmFieldKey).Value)
                    }
                } 
            }

            if($Fields[$primaryKeyField] -eq $null)
            {
                $Fields.Add($primaryKeyField, $Id)
            }
            # if no record exists, then create new
            $result = New-CrmRecord -conn $conn -EntityLogicalName $entityLogicalName -Fields $Fields

            return $result
        }
        else
        {   
            if($CrmRecord -ne $null)
            {
                # if record exists, then swap original record so that we can compare updated fields
                $CrmRecord.original = $existingRecord.original
            }
        }
    }

    $newfields = New-Object 'System.Collections.Generic.Dictionary[[String], [Microsoft.Xrm.Tooling.Connector.CrmDataTypeWrapper]]'
    
    if($CrmRecord -ne $null)
    {                
        $originalRecord = $CrmRecord.original        
        $Id = $originalRecord[$primaryKeyField]
        
        foreach($crmFieldKey in ($CrmRecord | Get-Member -MemberType NoteProperty).Name)
        {
            $crmFieldValue = $CrmRecord.($crmFieldKey)
            if(($crmFieldKey -eq "original") -or ($crmFieldKey -eq "logicalname") -or ($crmFieldKey -eq "EntityReference")`
              -or ($crmFieldKey -like "*_Property") -or ($crmFieldKey -like "ReturnProperty_*"))
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
                elseif($crmFieldValue -eq $originalRecord[$crmFieldKey])
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
            
            # When value set to null, then just use raw type and set value to $null
            if($crmFieldValue -eq $null)
            {
                $newfield.Type = [Microsoft.Xrm.Tooling.Connector.CrmFieldType]::Raw
                $value = $null
            }
            else
            {
                if($CrmRecord.($crmFieldKey + "_Property") -ne $null)
                {
                    $type = $CrmRecord.($crmFieldKey + "_Property").Value.GetType().Name
                }
                else
                {
                    $type = $crmFieldValue.GetType().Name
                }
                switch($type)
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
                    default {
                        $newfield.Type = [Microsoft.Xrm.Tooling.Connector.CrmFieldType]::Raw
                        $value = $crmFieldValue
                        break
                    }
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
            if($field.value -eq $null)
            {
                $newfield.Type = [Microsoft.Xrm.Tooling.Connector.CrmFieldType]::Raw
            }
            else
            {
                $newfield.Type = MapFieldTypeByFieldValue -Value $field.Value
            }
            $newfield.Value = $field.Value
            $newfields.Add($field.Key, $newfield)
        }
    }
    try
    {
        # if no field has new value, then do nothing.
        if($newfields.Count -eq 0)
        {
            return
        }
        $result = $conn.UpdateEntity($entityLogicalName, $primaryKeyField, $Id, $newfields, $null, $false, [Guid]::Empty)
        if(!$result)
        {
			throw LastCrmConnectorException($conn)
        }
    }
    catch
    {
        #TODO: Throw Exceptions back to user
		throw LastCrmConnectorException($conn)
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
                throw LastCrmConnectorException($conn)
            }
        }
        catch
        {
            throw LastCrmConnectorException($conn)
        }
    }
}

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
            throw LastCrmConnectorException($conn)
        }
    }
    catch
    {
        throw LastCrmConnectorException($conn)
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
                throw LastCrmConnectorException($conn)
            }

			write-verbose "Completed..."
		}
		catch
		{
		    throw LastCrmConnectorException($conn)
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
            throw LastCrmConnectorException($conn)
        }
    }
    catch
    {
        throw LastCrmConnectorException($conn)
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
            throw LastCrmConnectorException($conn)
        }
    }
    catch
    {
        throw LastCrmConnectorException($conn)
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
            throw LastCrmConnectorException($conn)
        }
    }
    catch
    {
        throw LastCrmConnectorException($conn)
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
            throw LastCrmConnectorException($conn)
        }
    }
    catch
    {
        throw LastCrmConnectorException($conn)
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
                $newfield.Type = MapFieldTypeByFieldValue -Value $field.Value
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
				throw LastCrmConnectorException($conn)
			}
		}
		catch
		{
			throw LastCrmConnectorException($conn)
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
            throw LastCrmConnectorException($conn)
        }
    }
    catch
    {
        throw LastCrmConnectorException($conn)
    }
}

#ExecuteWorkflowOnEntity  
function Invoke-CrmRecordWorkflow{
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml
    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,

        [parameter(Mandatory=$true, Position=1, ParameterSetName="CrmRecordWithWorkflowId")]
        [parameter(ParameterSetName="CrmRecordWithWorkflowName")]
        [PSObject]$CrmRecord,

        [parameter(Mandatory=$true, Position=1, ParameterSetName="CrmRecordIdWithWorkflowId")]
        [parameter(ParameterSetName="CrmRecordIdWithWorkflowName")]
        [Alias("Id", "StringId")]
        [string]$EntityId,

        [parameter(Mandatory=$true, Position=2, ParameterSetName="CrmRecordIdWithWorkflowName")]
        [parameter(ParameterSetName="CrmRecordWithWorkflowName")]
        [string]$WorkflowName, 

        [parameter(Mandatory=$true, Position=2, ParameterSetName="CrmRecordIdWithWorkflowId")]
        [parameter(ParameterSetName="CrmRecordWithWorkflowId")]
        [string]$WorkflowId
    )
	$conn = VerifyCrmConnectionParam $conn
    if($CrmRecord -ne $null)
    {        
        $Id = $CrmRecord.($CrmRecord.logicalname + "id")
    }
    elseif($EntityId -ne $null)
    {
        $Id = [guid]::Parse($EntityId)
    }
    try
    {
        $result = $null 
        if($WorkflowName -ne $null){
		        $fetch = @"
<fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="false">
  <entity name="workflow">
    <attribute name="workflowid" />
    <attribute name="name" />
    <attribute name="primaryentity" />
    <attribute name="type" />
    <order attribute="name" descending="false" />
    <filter type="and">
	  <condition attribute="name" operator="eq" value="$WorkflowName" />
      <condition attribute="statecode" operator="eq" value="1" />
      <condition attribute="type" operator="eq" value="1" />
      <condition attribute="rendererobjecttypecode" operator="null" />
      <condition attribute="category" operator="eq" value="0" />
    </filter>
  </entity>
</fetch>
"@
			$workflowResult = (Get-CrmRecordsByFetch -Fetch $fetch -TopCount 1)
			if($workflowResult.NextPage){
				throw "Duplicate workflow detected, try executing the workflow by its ID"
			}
			$WorkflowId = $workflowResult.CrmRecords[0].workflowid
        }
		if($WorkflowId -ne $null){
            $execWFReq = New-Object Microsoft.Crm.Sdk.Messages.ExecuteWorkflowRequest
            $execWFReq.EntityId = $Id
            $execWFReq.WorkflowId=$WorkflowId
            $result = $conn.ExecuteCrmOrganizationRequest($execWFReq) 
        }
		else{ throw "Duplicate workflow detected, try executing the workflow by its ID"}
		if($result -eq $null)
        {
            throw LastCrmConnectorException($conn)
        }
    }
    catch
    {
        throw LastCrmConnectorException($conn)
    }
    return $result
}

#Alias for Get-MyCrmUserId
New-Alias -Name Get-CrmCurrentUserId -Value Get-MyCrmUserId

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
            throw LastCrmConnectorException($conn)
        }
    }
    catch
    {
        throw LastCrmConnectorException($conn)
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
            throw LastCrmConnectorException($conn)
        }
    }
    catch
    {
        throw LastCrmConnectorException($conn)
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
            throw LastCrmConnectorException($conn)
        }
    }
    catch
    {
        throw LastCrmConnectorException($conn)
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
		if($result -eq $null)
        {
            throw LastCrmConnectorException($conn)
        }
    }
    catch
    {
        throw LastCrmConnectorException($conn)
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
		$xml = [xml]$Fetch
		if($xml.fetch.count -ne 0 -and $TopCount -eq 0)
		{
			$TopCount = $xml.fetch.count
		}
        $records = $conn.GetEntityDataByFetchSearch($Fetch, $TopCount, $PageNumber, $PageCookie, [ref]$PagingCookie, [ref]$NextPage, [Guid]::Empty)
        if($conn.LastCrmException -ne $null)
        {
            throw LastCrmConnectorException($conn)
        }
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
            foreach($record in $records.Values){   
                $psobj = New-Object -TypeName System.Management.Automation.PSObject
                if($recordslist.Count -eq 0){
                    $atts = $xml.GetElementsByTagName('attribute')
					foreach($att in $atts){
						if($att.ParentNode.HasAttribute('alias')){
							$attName = $att.ParentNode.GetAttribute('alias') + "." + $att.name
						}
						else{
							$attName = $att.name
						}
						Add-Member -InputObject $psobj -MemberType NoteProperty -Name $attName -Value $null
						Add-Member -InputObject $psobj -MemberType NoteProperty -Name ($attName + "_Property") -Value $null
					}
                    foreach($att in $record.GetEnumerator()){
						#BUG where ReturnProperty_Id is returned as "ReturnProperty_Id " <-- with a trailing space
						$keyName = $att.Key
						if($keyName -eq "ReturnProperty_Id "){
							$keyName = "ReturnProperty_Id"
						}
						if(!($psobj | gm).Name.Contains($keyName)){
							Add-Member -InputObject $psobj -MemberType NoteProperty -Name $keyName -Value $null
						}
						if($att.Value -is [Microsoft.Xrm.Sdk.EntityReference]){
							$psobj.($keyName) = $att.Value.Name
						}
						elseif($att.Value -is [Microsoft.Xrm.Sdk.AliasedValue]){
							$psobj.($keyName) = $att.Value.Value
						}
						else{
							$psobj.($keyName) = $att.Value
						}
					}  
                }
                else{
                    foreach($att in $record.GetEnumerator()){
						#BUG where ReturnProperty_Id is returned as "ReturnProperty_Id " <-- with a trailing space
						$keyName = $att.Key
						if($keyName -eq "ReturnProperty_Id "){
							$keyName = "ReturnProperty_Id"
						}
                        if($att.Value -is [Microsoft.Xrm.Sdk.EntityReference]){
                            Add-Member -InputObject $psobj -MemberType NoteProperty -Name $keyName -Value $att.Value.Name
                        }
				    	elseif($att.Value -is [Microsoft.Xrm.Sdk.AliasedValue]){
				    		Add-Member -InputObject $psobj -MemberType NoteProperty -Name $keyName -Value $att.Value.Value
				    	}
                        else{
                            Add-Member -InputObject $psobj -MemberType NoteProperty -Name $keyName -Value $att.Value
                        }
                    }  
                }
				Add-Member -InputObject $psobj -MemberType NoteProperty -Name "original" -Value $record
				Add-Member -InputObject $psobj -MemberType NoteProperty -Name "logicalname" -Value $logicalname
				#adding Dynamic EntityReference
				if($psobj."ReturnProperty_Id" -ne $null -and $psobj."ReturnProperty_EntityName" -ne $null){
					Add-Member -InputObject $psobj -MemberType NoteProperty -Name "EntityReference" -Value (New-CrmEntityReference -EntityLogicalName $psobj."ReturnProperty_EntityName" -Id $psobj."ReturnProperty_Id")
				}
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
        throw LastCrmConnectorException($conn)
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
			throw LastCrmConnectorException($conn)
        }
    }
    catch
    {
        throw LastCrmConnectorException($conn)
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
            throw LastCrmConnectorException($conn)
        }
    }
    catch
    {
        throw LastCrmConnectorException($conn)
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
            throw LastCrmConnectorException($conn)
        }
    }
    catch
    {
        throw LastCrmConnectorException($conn)
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
            throw LastCrmConnectorException($conn)
        }
    }
    catch
    {
        throw LastCrmConnectorException($conn)
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
            throw LastCrmConnectorException($conn)
        }
    }
    catch
    {
        throw LastCrmConnectorException($conn)
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
            throw LastCrmConnectorException($conn)
        }
    }
    catch
    {
        throw LastCrmConnectorException($conn)
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
            throw LastCrmConnectorException($conn)
        }
    }
    catch
    {
        throw LastCrmConnectorException($conn)
    }

    return $result
}

#ImportSolutionToCrmAsync
function Import-CrmSolutionAsync{
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml
    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1)]
        [string]$SolutionFilePath,
        [parameter(Mandatory=$false, Position=2)]
        [switch]$ActivateWorkflows,
        [parameter(Mandatory=$false, Position=3)]
        [switch]$OverwriteUnManagedCustomizations,
        [parameter(Mandatory=$false, Position=4)]
        [switch]$SkipDependancyOnProductUpdateCheckOnInstall, 
        [parameter(Mandatory=$false, Position=5)]
        [switch]$PublishChanges,
		[parameter(Mandatory=$false, Position=6)]
        [int64]$MaxWaitTimeInSeconds,
		[parameter(Mandatory=$false, Position=7)]
        [switch]$ImportAsHoldingSolution, 
		[parameter(Mandatory=$false, Position=8)]
        [switch]$BlockUntilImportComplete
    )
	$conn = VerifyCrmConnectionParam $conn
	$importId = [guid]::Empty
	$asyncResponse = $null
    try
    {
        if (!(Test-Path $SolutionFilePath)) 
        {
            throw [System.IO.FileNotFoundException] "$SolutionFilePath not found."
        }
		
        Write-Output  "Importing solution file  $SolutionFilePath into: $($conn.CrmConnectOrgUriActual)" 
        Write-Verbose "OverwriteCustomizations: $OverwriteUnManagedCustomizations"
        Write-Verbose "SkipDependancyCheck:     $SkipDependancyOnProductUpdateCheckOnInstall"
		Write-Verbose "ImportAsHoldingSolution: $ImportAsHoldingSolution"
        Write-Verbose "Maximum seconds to poll: $MaxWaitTimeInSeconds"
        Write-Verbose "Block and wait?          $BlockUntilImportComplete"

        if($BlockUntilImportComplete -eq $false -and ($MaxWaitTimeInSeconds -gt 0 -or $MaxWaitTimeInSeconds -eq -1)){
			Write-Warning "MaxWaitTimeInSeconds is $MaxWaitTimeInSeconds, we assume the user wants to block until complete."
			Write-Warning "To avoid this warning in the future please specify the switch: -BlockUntilImportComplete"
			$BlockUntilImportComplete = $true
		}
		if($BlockUntilImportComplete -eq $false -and $PublishChanges -eq $true){
			Write-Warning "PublishChanges will be ignored because BlockUntilImportComplete is $BlockUntilImportComplete"
			$PublishChanges = $false; 
		}

		$data = [System.IO.File]::ReadAllBytes($SolutionFilePath)
		
		$request = New-Object Microsoft.Crm.Sdk.Messages.ImportSolutionRequest
		$request.CustomizationFile = $data  
		$request.PublishWorkflows = $ActivateWorkflows
		$request.OverwriteUnmanagedCustomizations = $OverwriteUnManagedCustomizations
		$request.SkipProductUpdateDependencies = $SkipDependancyOnProductUpdateCheckOnInstall
		$request.HoldingSolution = $ImportAsHoldingSolution

		$asyncRequest = New-Object Microsoft.Xrm.Sdk.Messages.ExecuteAsyncRequest
		$asyncRequest.Request = $request; 

		Write-Verbose "ExecuteCrmOrganizationRequest with ExecuteAsyncRequest containing ImportSolutionRequest() this process can take a while..."
		try
		{
			$asyncResponse = ($conn.ExecuteCrmOrganizationRequest($asyncRequest, "AsyncImportRequest") -as [Microsoft.Xrm.Sdk.Messages.ExecuteAsyncResponse]) 
			$importId = $asyncResponse.AsyncJobId
			
			Write-Verbose "ImportId (asyncoperationid): $importId" 
			if($importId -eq $null -or $importId -eq [Guid]::Empty)
			{
				throw "Import request failed, asyncoperationid is: $importId"
			}
			#if the caller wants to get the ID and does NOT want to wait 
			if($BlockUntilImportComplete -eq $false){
				return $asyncResponse; 
			}
		}
		catch
		{
			throw LastCrmConnectorException($conn)
		}    
        $pollingStart = Get-Date
        $isProcessing = $true
		$secondsSpentPolling = 0
        $pollingDelaySeconds = 5
		$transientFailureCount = 0; 
        Write-Verbose "Import of file completed, waiting on completion of AsyncOperationId: $importId"

		try{
			while(($isProcessing -and $secondsSpentPolling -lt $MaxWaitTimeInSeconds) -or ($isProcessing -and $MaxWaitTimeInSeconds -eq -1)){
				#delay
				Start-Sleep -Seconds $pollingDelaySeconds
				#check the import job for success/fail/inProgress
				try{
					$import = Get-CrmRecord -conn $conn -EntityLogicalName asyncoperation -Id $importId -Fields statuscode
				} catch {
					$transientFailureCount++; 
					Write-Verbose "Import Job status check did not succeed:  $($_.Exception)"
				}
				$status = $import.statuscode_Property.value.Value; 
				#Check for import completion - https://msdn.microsoft.com/en-us/library/gg309288.aspx
				if($status -lt 30){
					$isProcessing = $true
					$secondsSpentPolling = ([Int]((Get-Date) - $pollingStart).TotalSeconds)
					Write-Output "$($secondsSPentPolling.ToString("000")) sec of: $MaxWaitTimeInSeconds - ImportStatus: $($import.statuscode)"
				}
				elseif($status -eq 31 -or $status -eq 32 ){
					$isProcessing = $false
					throw "$($import.statuscode) - AsyncOperation with Id: $importId has been either cancelled or has failed."
					break; 
				}
				elseif($status -eq 30){
					$isProcessing = $false
					Write-Verbose "Processing Completed at: $($import.completedon)" 
					if($PublishChanges){
						Write-Verbose "PublishChanges set, executing: Publish-CrmAllCustomization using the same connection."
						Publish-CrmAllCustomization -conn $conn
						return $asyncResponse
					}
					else{
						Write-Output "Import Complete, don't forget to publish customizations."
						return $asyncResponse
					}
					break; 
				}
			}
			#User provided timeout and exit function with an error
			if($secondsSpentPolling -gt $MaxWaitTimeInSeconds){
				Write-Warning "Import-CrmSolutionAsync exited due to exceeding the maximum timeout of $MaxWaitTimeInSeconds. The import will continue in CRM async until it either succeeds or fails."
			}
			#at this point we appear to have imported successfully 
			return $asyncResponse; 
		} Catch {
			throw "AsyncOperation with ID: $importId has encountered an exception: $_"
		}
    }
    catch
    {
        throw $_.Exception
    }    
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
        [int64]$MaxWaitTimeInSeconds = 900,
		[parameter(Mandatory=$false, Position=7)]
        [switch]$ImportAsHoldingSolution, 
		[parameter(Mandatory=$false, Position=8)]
        [switch]$AsyncOperationImportMethod 
    )
	$conn = VerifyCrmConnectionParam $conn
	$importId = [guid]::Empty
    try
    {
        if (!(Test-Path $SolutionFilePath)) 
        {
            throw [System.IO.FileNotFoundException] "$SolutionFilePath not found."
        }
        Write-Host "Importing solution file $SolutionFilePath into: $($conn.CrmConnectOrgUriActual)" 
        Write-Verbose "OverwriteCustomizations: $OverwriteUnManagedCustomizations"
        Write-Verbose "SkipDependancyCheck: $SkipDependancyOnProductUpdateCheckOnInstall"
		Write-Verbose "ImportAsHoldingSolution: $ImportAsHoldingSolution"
        Write-Verbose "Maximum seconds to poll for successful completion: $MaxWaitTimeInSeconds"

		if($AsyncOperationImportMethod){
			Write-Warning "Option to import using Aync method flagged! With this option you may not be able to debug problems easily."
			Write-Warning "To import async going forward use the Import-CrmSolutionAsync cmdlet instead."
			$result = Import-CrmSolutionAsync `
				-Conn $conn `
				-SolutionFilePath $SolutionFilePath `
				-OverwriteUnManagedCustomizations:$OverwriteUnManagedCustomizations `
				-SkipDependancyOnProductUpdateCheckOnInstall:$SkipDependancyOnProductUpdateCheckOnInstall `
				-MaxWaitTimeInSeconds $MaxWaitTimeInSeconds `
				-ImportAsHoldingSolution:$ImportAsHoldingSolution `
				-BlockUntilImportComplete:$true; 

			Write-Verbose "Solution import using async completed - asyncoperationid = $($result.AsyncJobId)"; 
			return $result; 
		}

        Write-Verbose "Calling .ImportSolutionToCrm() this process can take minutes..."
        $result = $conn.ImportSolutionToCrm($SolutionFilePath, [ref]$importId, $ActivatePlugIns,
                $OverwriteUnManagedCustomizations, $SkipDependancyOnProductUpdateCheckOnInstall,$ImportAsHoldingSolution)
        Write-Verbose "ImportId: $result" 
		if ($result -eq [guid]::Empty) {
             throw LastCrmConnectorException($conn)
        }
        $pollingStart = Get-Date
        $isProcessing = $true
		$secondsSpentPolling = 0
        $pollingDelaySeconds = 5
		$TopPrevProcPercent = [double]0
		$isProcPercentReduced = $false
		#this is for a bug where the service will throw a 401 on retrieve of importjob during an import under certain conditions 
		$transientFailureCount = 0; 
        Write-Host "Import of file completed, waiting on completion of importId: $importId"
		try{
			while($isProcessing -and $secondsSpentPolling -lt $MaxWaitTimeInSeconds){
				#delay
				Start-Sleep -Seconds $pollingDelaySeconds
				#check the import job for success/fail/inProgress
				try{
					$import = Get-CrmRecord -conn $conn -EntityLogicalName importjob -Id $importId -Fields solutionname,data,completedon,startedon,progress
				} catch {
					if($transientFailureCount > 5){
						Write-Error "Import Job status check FAILED 5 times this could be due to a bug where the service returns a 401. Throwing lastException:"; 
						throw  $conn.LastCrmException
					}
					Write-Verbose "Import Job status check FAILED this could be due to a bug where the service returns a 401. We'll allow up to 5 failures before aborting."; 
					$transientFailureCount++; 
				}
				#Option to use Get-CrmRecords so we can force a no-lock to prevent hangs in the retrieve
				#$import = (Get-CrmRecords -conn $conn -EntityLogicalName importjob -FilterAttribute importjobid -FilterOperator eq -FilterValue $importId -Fields data,completedon,startedon,progress).CrmRecords[0]
				$importManifest = ([xml]($import).data).importexportxml.solutionManifests.solutionManifest
				$ProcPercent = [double](Coalesce $import.progress "0")

				#check if processing percentage reduced at any given time
				if($TopPrevProcPercent -gt $ProcPercent)
				{
					$isProcPercentReduced = $true
					Write-Verbose "Processing is reversing... import will fail."
				} else {
					$TopPrevProcPercent = $ProcPercent
				}

				#Check for import completion 
				if($import.completedon -eq $null -and $importManifest.result.result -ne "success"){
					$isProcessing = $true
					$secondsSpentPolling = ([Int]((Get-Date) - $pollingStart).TotalSeconds)
					Write-Host "$($secondsSPentPolling.ToString("000")) seconds of max: $MaxWaitTimeInSeconds ... ImportJob%: $ProcPercent"
				}
				else {
					Write-Verbose "Processing Completed at: $($import.completedon) with ImportJob%: $ProcPercent" 
                    Write-Verbose "Import Manifest Result: $($importManifest.result.result) with ImportJob%: $ProcPercent" 					

                    $solutionImportResults =  Select-Xml -Xml ([xml]$import.data) -XPath "//result"
                    $anyFailuresInImport = $false;
                    $allErrorText = "";

                    foreach($solutionImportResult in $solutionImportResults)
                    {
                        try{
                            $resultParent = ""
                            $itemResult = ""
                            $resultParent = $($solutionImportResult.Node.ParentNode.ParentNode.Name)
                            $itemResult = $($solutionImportResult.Node.result)
                        }catch{}

                        Write-Verbose "Item:$resultParent  result: $itemResult" # write each item result in result data
                        
                        if ($solutionImportResult.Node.result -ne "success")
                        {
                            # if any error in result print more error details
                            try{
                                $errorCode = ""
                                $errorText = ""
                                $moreErrorDetails = ""
                                $errorCode = $($solutionImportResult.Node.errorcode)
                                $errorText = $($solutionImportResult.Node.errortext)
                                $moreErrorDetails = $solutionImportResult.Node.parameters.InnerXml

                            }catch{}

                            Write-Verbose "errorcode: $errorCode errortext: $errorText more details: $moreErrorDetails"
                            if ($solutionImportResult.Node.result -eq "failure") # Fail only on errors, not on warnings
                            {
                                $anyFailuresInImport = $true; # mark if any failures in solution import
                                $allErrorText = $allErrorText + ";" + $errorText;
                            }
                        }
                    }
					if(-not $isProcPercentReduced -and $importManifest.result.result -eq "success" -and (-not $anyFailuresInImport))
					{
                        Write-Verbose "Setting to 100% since all results are success"
						$ProcPercent = 100.0
					}					
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
        if(($importManifest.result.result -eq "failure") -or ($ProcPercent -lt 100) -or $anyFailuresInImport) #Must look at %age instead of this result as the result is usually wrong!
        {
            Write-Verbose "Import result: failed - job with ID: $importId failed at $ProcPercent complete."
            throw $allErrorText
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
            throw LastCrmConnectorException($conn)
        }
    }
    catch
    {
        throw LastCrmConnectorException($conn)
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
        throw LastCrmConnectorException($conn)
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
            throw LastCrmConnectorException($conn)
        }
    }
    catch
    {
        throw LastCrmConnectorException($conn)
    }

    return $result
}

#PublishTheme  
function Publish-CrmTheme{
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml
 [CmdletBinding()]
    PARAM( 
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1, ParameterSetName="ThemeName")]
        [string]$ThemeName,
		[parameter(Mandatory=$true, Position=1, ParameterSetName="ThemeId")]
        [guid]$ThemeId
    )

	$conn = VerifyCrmConnectionParam $conn  

    try
    {
		if($ThemeName -ne "")
		{
			$themes = Get-CrmRecords -conn $conn -EntityLogicalName theme -FilterAttribute name -FilterOperator eq -FilterValue $ThemeName -WarningAction SilentlyContinue
			if($themes.CrmRecords.Count -eq 0)
			{
				Write-Warning "No Theme found"
				return
			}
			else
			{
				$ThemeId = $themes.CrmRecords[0].themeid
			}
		}
        $req = New-Object Microsoft.Crm.Sdk.Messages.PublishThemeRequest
		$req.target = New-CrmEntityReference -EntityLogicalName "theme" -Id $ThemeId
		$result = $conn.ExecuteCrmOrganizationRequest($req, $null)
		if(!$result)
        {
            throw LastCrmConnectorException($conn)
        }
    }
    catch
    {
        throw LastCrmConnectorException($conn)
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
            throw LastCrmConnectorException($conn)
        }
    }
    catch
    {
        throw LastCrmConnectorException($conn)
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
        throw LastCrmConnectorException($conn)
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
		    #$Id = $CrmRecord.($EntityLogicalName + "id")
			$Id = $CrmRecord.'ReturnProperty_Id'
		}

        # Try to parse into int
        $StateCodeInt = 0
        $StatusCodeInt = 0
        
		try
		{
            if([int32]::TryParse($StateCode, [ref]$StateCodeInt) -and [int32]::TryParse($StatusCode, [ref]$StatusCodeInt))
            {
                $result = $conn.UpdateStateAndStatusForEntity($EntityLogicalName, $Id, $StateCodeInt, $statusCodeInt, [Guid]::Empty)
			    if(!$result)
			    {
			    	throw LastCrmConnectorException($conn)
			    }
            }
            else
            {
                $result = $conn.UpdateStateAndStatusForEntity($EntityLogicalName, $Id, $stateCode, $statusCode, [Guid]::Empty)
			    if(!$result)
			    {
			    	throw LastCrmConnectorException($conn)
			    }
            }
		}
		catch
		{
			throw LastCrmConnectorException($conn)
		}
	}
}

function Add-CrmSecurityRoleToTeam{
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml

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

### Other Cmdlets added by Dynamics CRM PFE ###
function Add-CrmSecurityRoleToUser{
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml

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
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml

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
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml

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
            throw LastCrmConnectorException($conn)
        }
    }
    catch
    {
        throw LastCrmConnectorException($conn)
    }    
}

function Enable-CrmLanguagePack{
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml

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
            throw LastCrmConnectorException($conn)
        }
    }
    catch
    {
        throw LastCrmConnectorException($conn)
    }    
}

function Export-CrmApplicationRibbonXml {
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml

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
        throw LastCrmConnectorException($conn)
    }
    catch
    {
	    throw LastCrmConnectorException($conn)
    }
}

function Export-CrmEntityRibbonXml {
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml
 
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
        throw LastCrmConnectorException($conn)
    }
    catch
    {
	    throw LastCrmConnectorException($conn)
    }
}

function Export-CrmSolution{
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml

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
			$exportRequest.ExportSales = $ExportSales
		}

        Write-Verbose 'ExportSolutionRequests may take several minutes to complete execution.'

        $response = [Microsoft.Crm.Sdk.Messages.ExportSolutionResponse]($conn.ExecuteCrmOrganizationRequest($exportRequest))

		if($response -eq $null){
			if($conn.LastCrmException -eq ""){
				throw "The result was null, please double check the command"
			}
			else{
				throw LastCrmConnectorException($conn)
			}
		}

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
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml

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
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml

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
        throw LastCrmConnectorException($conn)
    }    

    return $response.LocaleIds
}

function Get-CrmEntityRecordCount{
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml

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
				throw LastCrmConnectorException($conn)
			}
        }
        catch
        {
            throw LastCrmConnectorException($conn)
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
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml

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
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml

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
"@; 

	$users = Get-CrmRecordsByFetch -conn $conn -Fetch $fetch; 
	Write-Output 'IsLicensed:' ($users.CrmRecords | group islicensed | select count, name); 
	Write-Output 'AccessMode:' ($users.CrmRecords | group accessmode | select count, name); 
	Write-Output 'CalType:' ($users.CrmRecords | group accessmode | select count, name); 
}

function Get-CrmOrgDbOrgSettings{
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml

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
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml

    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1)][alias("EntityName")]
        [string]$EntityLogicalName,
        [parameter(Mandatory=$false, Position=2)][alias("FieldName")]
        [string]$FilterAttribute,
        [parameter(Mandatory=$false, Position=3)][alias("Op")]
        [ValidateSet('eq','neq','ne','gt','ge','le','lt','like','not-like','in','not-in','between','not-between','null','not-null','yesterday','today','tomorrow','last-seven-days','next-seven-days','last-week','this-week','next-week','last-month','this-month','next-month','on','on-or-before','on-or-after','last-year','this-year','next-year','last-x-hours','next-x-hours','last-x-days','next-x-days','last-x-weeks','next-x-weeks','last-x-months','next-x-months','olderthan-x-months','olderthan-x-years','olderthan-x-weeks','olderthan-x-days','olderthan-x-hours','olderthan-x-minutes','last-x-years','next-x-years','eq-userid','ne-userid','eq-userteams','eq-useroruserteams','eq-useroruserhierarchy','eq-useroruserhierarchyandteams','eq-businessid','ne-businessid','eq-userlanguage','this-fiscal-year','this-fiscal-period','next-fiscal-year','next-fiscal-period','last-fiscal-year','last-fiscal-period','last-x-fiscal-years','last-x-fiscal-periods','next-x-fiscal-years','next-x-fiscal-periods','in-fiscal-year','in-fiscal-period','in-fiscal-period-and-year','in-or-before-fiscal-period-and-year','in-or-after-fiscal-period-and-year','begins-with','not-begin-with','ends-with','not-end-with','under','eq-or-under','not-under','above','eq-or-above')]
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
    if( !($EntityLogicalName -cmatch "^[a-z_]*$") )
    {
        $EntityLogicalName = $EntityLogicalName.ToLower()
        Write-Verbose "EntityLogicalName contains uppercase which isn't possible in CRM, overwritting with ToLower() new value: $EntityLogicalName"
    }

    if( ($Fields -eq "*") -OR ($Fields -eq "%") )
    {
        Write-Verbose 'PERFORMANCE: All attributes were requested'
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
	if(
		($FilterAttribute -and !$FilterOperator) -or 
		(!$FilterAttribute -and $FilterOperator) -or
		($FilterValue -and (!$FilterOperator -or !$FilterAttribute)) 
	){
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
    elseif($FilterAttribute -and $FilterOperator){
         # Escape XML charactors
        $FilterValue = [System.Security.SecurityElement]::Escape($FilterValue)
        Write-Verbose "Using the supplied single filter of $FilterAttribute '$FilterOperator' and NO value"
        $fetch = 
@"
    <fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="false" no-lock="true">
        <entity name="{0}">
            {1}
            <filter type='and'>
                <condition attribute='{2}' operator='{3}' />
            </filter>
        </entity>
    </fetch>
"@
    }
    else
    {
        Write-Verbose "PERFORMANCE: `$FilterAttribute `$FilterOperator `$FilterValue were not supplied, fetching all records with NO filter."
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
        Write-Verbose "PERFORMANCE: All rows were requested instead of the first 5000"
        $results = Get-CrmRecordsByFetch -conn $conn -Fetch $fetch -AllRows
    }
    else
    {
        $results = Get-CrmRecordsByFetch -conn $conn -Fetch $fetch -TopCount $TopCount
    }
    
    return $results
}

function Get-CrmRecordsByViewName{
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml

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
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml

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
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml

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
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml

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
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml

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
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml

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
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml

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
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml

    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1)]
        [string]$UserId
    )

	$conn = VerifyCrmConnectionParam $conn
    
    $fetch = @"
    <fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="false" no-lock="true">
    <entity name="mailbox">
	<all-attributes />
    <filter type="and">
      <condition attribute="regardingobjectid" operator="eq" value="{$UserId}" />
    </filter>
  </entity>
</fetch>
"@

    $record = (Get-CrmRecordsByFetch -conn $conn -Fetch $fetch).CrmRecords
	switch ($record.count)
	{
		{$_ -eq 0} {Throw "The user $UserId has no mailbox."}
		{$_ -ge 2} {
						foreach( $id in $record.mailboxid)
						{
							if(!$idString)
							{
								[string]$idString = $id
							}
							else
							{
								[string]$idString = "$idString,$id"
							}
						}
						Throw "The user $UserId has more than one mailbox: $idString"
					}
		Default {
					return $record}
	}
    
}

function Get-CrmUserPrivileges{
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml

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
	        throw LastCrmConnectorException($conn)
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
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml

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
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml

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

function Grant-CrmRecordAccess {
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml
    [OutputType([void])]
    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false, Position=0)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
		[parameter(Mandatory=$true, Position=1, ParameterSetName="CrmRecord", ValueFromPipeline=$true)]
        [PSObject[]]$CrmRecord,
		[parameter(Mandatory=$true, Position=1, ParameterSetName="NameWithId")]
        [string]$EntityLogicalName,
        [parameter(Mandatory=$true, Position=2, ParameterSetName="NameWithId")]
        [guid]$Id,
		[parameter(Mandatory=$true, Position=3)]
        [Microsoft.Xrm.Sdk.EntityReference]$Principal,
        [parameter(Mandatory=$true, Position=4)]
        [Microsoft.Crm.Sdk.Messages.AccessRights]$AccessMask
    )
    begin
    {
        $conn = VerifyCrmConnectionParam $conn

        if ($EntityLogicalName) {
            $CrmRecord += [PSCustomObject] @{
                logicalname = $EntityLogicalName
                "$($EntityLogicalName)id" = $Id
            }
        }
    }
    process
    {
        foreach ($record in $CrmRecord) {
            try {
                $request = [Microsoft.Crm.Sdk.Messages.GrantAccessRequest]::new()
                $request.Target = New-CrmEntityReference -EntityLogicalName $record.logicalname -Id $record.($record.logicalname + "id")
                $principalAccess = [Microsoft.Crm.Sdk.Messages.PrincipalAccess]::new()
                $principalAccess.Principal = $Principal
                $principalAccess.AccessMask = $AccessMask
                $request.PrincipalAccess = $principalAccess
                
                [Microsoft.Crm.Sdk.Messages.GrantAccessResponse]$conn.Execute($request) | Out-Null
            }
            catch {
                Write-Error $_
            }   
        }
    }
}

function Import-CrmSolutionTranslation{
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml

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
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml

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
        throw LastCrmConnectorException($conn)
    }    

    return $result
}

function Invoke-CrmAction {
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml
    [OutputType([hashtable])]
    [OutputType([Microsoft.Xrm.Sdk.OrganizationResponse], ParameterSetName="Raw")]
    param (
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]
        $conn,

        [Parameter(
            Position=1,
            Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Name,

        [Parameter(Position=2)]
        [hashtable]
        $Parameters,

        [Parameter(ValueFromPipeline, Position=3)]
        [ValidateNotNullOrEmpty()]
        [Microsoft.Xrm.Sdk.EntityReference]
        $Target,

        [Parameter(ParameterSetName="Raw")]
        [switch]
        $Raw
    )
    begin
    {
        $conn = VerifyCrmConnectionParam $conn
    }
    process
    {
		$request = new-object Microsoft.Xrm.Sdk.OrganizationRequest
		$request.RequestName = $Name
        if($Target) {
            $request.Parameters.Add("Target", $Target) 
        }

        if($Parameters) {
            foreach($parameter in $Parameters.GetEnumerator()) {
                $request.Parameters.Add($parameter.Name, $parameter.Value)
            }
        }

        try {
            $response = $conn.Execute($request)
        
            if($Raw) {
                Write-Output $response
            } elseif ($response.Results -and $response.Results.Count -gt 0) {
                $outputArguments = @{}
                foreach($outputArgument in $response.Results) {
                    $outputArguments.Add($outputArgument.Key, $outputArgument.Value)
                }
                Write-Output $outputArguments
            } else {
                Write-Output $null
            }
        }
        catch {
            Write-Error $_
        }
    }
}

function Publish-CrmCustomization{
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml

 [CmdletBinding()]
    PARAM( 
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$false)]
        [switch]$Entity,
        [parameter(Mandatory=$false)]
        [string[]]$EntityLogicalNames,
        [parameter(Mandatory=$false)]
        [switch]$Ribbon,
        [parameter(Mandatory=$false)]
        [switch]$SiteMap,
        [parameter(Mandatory=$false)]
        [switch]$Dashbord,
        [parameter(Mandatory=$false)]
        [guid[]]$DashbordIds,
        [parameter(Mandatory=$false)]
        [switch]$OptionSet,
        [parameter(Mandatory=$false)]
        [string[]]$OptionSetNames,
        [parameter(Mandatory=$false)]
        [switch]$WebResource,
        [parameter(Mandatory=$false)]
        [guid[]]$WebResourceIds
    )

	$conn = VerifyCrmConnectionParam $conn  

    $parameterXml = "<importexportxml>"

    if($Entity -and $EntityLogicalNames.Count -ne 0)
    {
        $parameterXml += "<entities>"
        foreach($entityLogicalName in $EntityLogicalNames)
        {
            $parameterXml += "<entity>" + $entityLogicalName + "</entity>"
        }
        $parameterXml += "</entities>"
    }
    if($Ribbon)
    {
        $parameterXml += "<ribbons><ribbon></ribbon></ribbons>"
    }
    if($Dashbord -and $DashbordIds.Count -ne 0)
    {
        $parameterXml += "<dashboards>"
        foreach($dashbordId in $DashbordIds)
        {
            $parameterXml += "<dashboard>{" + $dashbordId + "}</dashboard>"
        }
        $parameterXml += "</dashboards>"
    }
    if($OptionSet -and $OptionSetNames.Count -ne 0)
    {
        $parameterXml += "<optionsets>"
        foreach($optionSetName in $OptionSetNames)
        {
            $parameterXml += "<optionset>{" + $optionSetName + "}</optionset>"
        }
        $parameterXml += "</optionsets>"
    }
    if($SiteMap)
    {
        $parameterXml += "<sitemaps><sitemap></sitemap></sitemaps>"
    }
    if($WebResource -and $WebResourceIds.Count -ne 0)
    {
        $parameterXml += "<webresources>"
        foreach($webResourceId in $WebResourceIds)
        {
            $parameterXml += "<webresource>{" + $webResourceId + "}</webresource>"
        }
        $parameterXml += "</webresources>"
    }

    $parameterXml += "</importexportxml>"

    $request = New-Object Microsoft.Crm.Sdk.Messages.PublishXmlRequest
    $request.ParameterXml = $parameterXml
    try
    {
        $response = $conn.ExecuteCrmOrganizationRequest($request, $null)
        if($response.ResponseName -eq $null)
        {
            throw LastCrmConnectorException($conn)
        }
    }
    catch
    {
        throw LastCrmConnectorException($conn)
    }    
}

function Publish-CrmAllCustomization{
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml

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
        throw LastCrmConnectorException($conn)
    }    

    #return $response
}

function Remove-CrmSecurityRoleFromTeam{
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml

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
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml

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
            throw LastCrmConnectorException($conn)
        }
    }
    catch
    {
        throw LastCrmConnectorException($conn)
    } 
}

function Revoke-CrmEmailAddress{
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml

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
        Set-CrmRecord -conn $conn -EntityLogicalName systemuser -Id $UserId -Fields @{"emailrouteraccessapproval"=(New-CrmOptionSetValue 3)}
    }
    else
    {
        Set-CrmRecord -conn $conn -EntityLogicalName queue -Id $QueueId -Fields @{"emailrouteraccessapproval"=(New-CrmOptionSetValue 3)}
    }
}

function Revoke-CrmRecordAccess {
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml
    [OutputType([void])]
    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false, Position=0)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
		[parameter(Mandatory=$true, Position=1, ParameterSetName="CrmRecord", ValueFromPipeline=$true)]
        [PSObject[]]$CrmRecord,
        [parameter(Mandatory=$true, Position=1, ParameterSetName="NameWithId")]
        [string]$EntityLogicalName,        
        [parameter(Mandatory=$true, Position=2, ParameterSetName="NameWithId")]
        [guid]$Id,
        [parameter(Mandatory=$true, Position=3)]
        [Microsoft.Xrm.Sdk.EntityReference]$Revokee
    )
    begin
    {
        $conn = VerifyCrmConnectionParam $conn

        if ($EntityLogicalName) {
            $CrmRecord += [PSCustomObject] @{
                logicalname = $EntityLogicalName
                "$($EntityLogicalName)id" = $Id
            }
        }
    }
    process
    {
        foreach ($record in $CrmRecord) {
            try {
                $request = [Microsoft.Crm.Sdk.Messages.RevokeAccessRequest]::new()
                $request.Target = New-CrmEntityReference -EntityLogicalName $record.logicalname -Id $record.($record.logicalname + "id")
                $request.Revokee = $Revokee
                
                [Microsoft.Crm.Sdk.Messages.RevokeAccessResponse]$conn.Execute($request) | Out-Null
            }
            catch {
                Write-Error $_
            }   
        }
    }
}

function Set-CrmSolutionVersionNumber {
	[CmdletBinding()]
	PARAM(
		[parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
		[parameter(Mandatory=$true, Position=1)]
        [string]$SolutionName,
		[parameter(Mandatory=$true, Position=2)]
		[ValidatePattern('^(?:[\d]{1,}\.){1,3}[\d]{1,}$')]
		[string]$VersionNumber
	)

	$conn = VerifyCrmConnectionParam $conn

	$solutionRecords = (Get-CrmRecords -conn $conn -EntityLogicalName solution -FilterAttribute uniquename -FilterOperator "like" -FilterValue $SolutionName -Fields uniquename,version )
    #if we can't find just one solution matching then ERROR
    if($solutionRecords.CrmRecords.Count -ne 1)
    {
        $friendlyName = $conn.ConnectedOrgFriendlyName.ToString()
        throw "Solution with name `"$SolutionName`" in CRM Instance: `"$friendlyName`" not found!"
    }
    #else PROCEED 
    
	$crmSolutionRecord = $solutionRecords.CrmRecords[0]
	$oldVersion = $crmSolutionRecord.version
    $crmSolutionRecord.version = $VersionNumber

	try
    {
		Write-Verbose "Updating $($crmSolutionRecord.uniquename) version to $VersionNumber"
		Set-CrmRecord -conn $conn -CrmRecord $crmSolutionRecord

        Write-Verbose "Successfully updated solution record"
        $result = New-Object psObject

        Add-Member -InputObject $result -MemberType NoteProperty -Name "PreviousVersionNumber" -Value $oldVersion
        Add-Member -InputObject $result -MemberType NoteProperty -Name "NewVersionNumber" -Value $VersionNumber
    }
    catch
    {
        throw LastCrmConnectorException($conn)
    }

    $result
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
	#powershell 4.0+ is required for New-TimeSpan -Seconds $TimeoutInSeconds 
	$newTimeout = New-Object System.TimeSpan -ArgumentList 0,0,120
	if(!$SetDefault){
	    $newTimeout = New-Object System.TimeSpan -ArgumentList 0,0,$TimeoutInSeconds
	}
	if($conn.OrganizationServiceProxy -and $conn.OrganizationServiceProxy.Timeout){
	    try{
			Write-Verbose "Updating Timeout on OrganizationServiceProxy"
			$conn.OrganizationServiceProxy.Timeout = $newTimeout
	    }
	    catch{
			Write-Verbose "Failed to set the timeout value"        
	    }
	}
	if($conn.OrganizationWebProxyClient -and $conn.OrganizationWebProxyClient.ChannelFactory.Endpoint.Binding){
	    try{
			Write-Verbose "Updating Timeouts on OrganizationWebProxyClient"
			$conn.OrganizationWebProxyClient.ChannelFactory.Endpoint.Binding.OpenTimeout = $newTimeout
			$conn.OrganizationWebProxyClient.ChannelFactory.Endpoint.Binding.CloseTimeout = $newTimeout
			$conn.OrganizationWebProxyClient.ChannelFactory.Endpoint.Binding.ReceiveTimeout = $newTimeout
			$conn.OrganizationWebProxyClient.ChannelFactory.Endpoint.Binding.SendTimeout = $newTimeout
	    }
	    catch{
			Write-Verbose "Failed to set the timeout values"
	    }
	}
}

function Set-CrmSystemSettings {
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml
    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
		[parameter(Mandatory=$false)]
		[guid]$AcknowledgementTemplateId,
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
		[bool]$AllowUserFormModePreference,
        [parameter(Mandatory=$false)]
        [bool]$AllowUsersSeeAppdownloadMessage,
        [parameter(Mandatory=$false)]
        [bool]$AllowWebExcelExport,
		[parameter(Mandatory=$false)]
        [string]$AMDesignator,
		[parameter(Mandatory=$false)]
        [bool]$AutoApplyDefaultonCaseCreate,
		[parameter(Mandatory=$false)]
        [bool]$AutoApplyDefaultonCaseUpdate,
		[parameter(Mandatory=$false)]
        [bool]$AutoApplySLA,
		[parameter(Mandatory=$false)]
        [string]$BingMapsApiKey,
        [parameter(Mandatory=$false)]
        [string]$BlockedAttachments,
		[parameter(Mandatory=$false)]
        [guid]$BusinessClosureCalendarId,
        [parameter(Mandatory=$false)]
        [string]$CampaignPrefix,
		[parameter(Mandatory=$false)]
        [bool]$CascadeStatusUpdate,
        [parameter(Mandatory=$false)]
        [string]$CasePrefix,
        [parameter(Mandatory=$false)]
        [string]$ContractPrefix,
		[parameter(Mandatory=$false)]
        [bool]$CortanaProactiveExperienceEnabled,
		[parameter(Mandatory=$false)]
        [bool]$CreateProductsWithoutParentInActiveState,
		[parameter(Mandatory=$false)]
        [int]$CurrencyDecimalPrecision,
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

    $attributesMetadata = Get-CrmEntityAttributes -conn $conn -EntityLogicalName organization
        
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
        [parameter(Mandatory=$false, Position=3)]
        [guid]$ReassignUserId
    )

	$conn = VerifyCrmConnectionParam $conn

	# If ReassignUserId is not passed, then assign them to myself  
	if($ReassignUserId -eq $null)  
	{  
		$ReassignUserId = $UserId  
	}  

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
            throw LastCrmConnectorException($conn)
        }
    }
    catch
    {
        throw LastCrmConnectorException($conn)
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
        [switch]$ApplyDefaultEmailSettings,
        [parameter(Mandatory=$false, ParameterSetName="Status")]
        [string]$StateCode,
        [parameter(Mandatory=$false, ParameterSetName="Status")]
        [string]$StatusCode,
        [parameter(Mandatory=$false, ParameterSetName="Status")]
        [switch]$ScheduleTest,
        [parameter(Mandatory=$false, ParameterSetName="Status")]
        [switch]$MarkedAsPrimaryForExchangeSync,
        [parameter(Mandatory=$false, ParameterSetName="Status")]
        [switch]$ApproveEmail
    )
	$conn = VerifyCrmConnectionParam $conn

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
    if($ScheduleTest)
    {
        $updateFields.Add("testemailconfigurationscheduled", $true)
    }
    if($MarkedAsPrimaryForExchangeSync)
    {
        $updateFields.Add("orgmarkedasprimaryforexchangesync", $true)
    }
    if($ApproveEmail)
    {
        Approve-CrmEmailAddress -conn $conn -UserId $UserId
    }
    foreach($parameter in $MyInvocation.BoundParameters.GetEnumerator())
    {   
        if($parameter.Key -in ("EmailServerProfile"))
        {
            $updateFields.Add($parameter.Key.ToLower(), (New-CrmEntityReference emailserverprofile $parameter.Value))
        }
        elseif($parameter.Key -in ("IncomingEmailDeliveryMethod","OutgoingEmailDeliveryMethod","ACTDeliveryMethod"))
        {
            $updateFields.Add($parameter.Key.ToLower(), (New-CrmOptionSetValue $parameter.Value))
        }
        elseif($parameter.Key -in ("StateCode","StatusCode"))
        {
            Set-CrmRecordState -conn $conn -EntityLogicalName mailbox -Id $Id -StateCode $StateCode -StatusCode $StatusCode
        }
        elseif($parameter.Key -in ("EmailAddress"))
        {
            $updateFields.Add($parameter.Key.ToLower(), $parameter.Value)
        }
    }

    Set-CrmRecord -conn $conn -EntityLogicalName mailbox -Id $Id -Fields $updateFields
}

function Set-CrmQueueMailbox {
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml
    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1)]
        [string]$QueueId,
        [parameter(Mandatory=$false)]
        [ValidatePattern('/^([a-zA-Z0-9])+([a-zA-Z0-9\._-])*@([a-zA-Z0-9_-])+([a-zA-Z0-9\._-]+)+$/')]
        [string]$EmailAddress,
        [parameter(Mandatory=$false, ParameterSetName="Custom")]        
        [guid]$EmailServerProfile,
        [parameter(Mandatory=$false, ParameterSetName="Custom")]
        [int]$IncomingEmailDeliveryMethod,
        [parameter(Mandatory=$false, ParameterSetName="Custom")]
        [int]$OutgoingEmailDeliveryMethod,
        [parameter(Mandatory=$false, ParameterSetName="Default")]
        [switch]$ApplyDefaultEmailSettings,
        [parameter(Mandatory=$false, ParameterSetName="Status")]
        [string]$StateCode,
        [parameter(Mandatory=$false, ParameterSetName="Status")]
        [string]$StatusCode,
        [parameter(Mandatory=$false, ParameterSetName="Status")]
        [switch]$ScheduleTest,
        [parameter(Mandatory=$false, ParameterSetName="Status")]
        [switch]$MarkedAsPrimaryForExchangeSync,
        [parameter(Mandatory=$false, ParameterSetName="Status")]
        [switch]$ApproveEmail
    )
	$conn = VerifyCrmConnectionParam $conn
    $fetch = @"
    <fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="false" no-lock="true">
      <entity name="mailbox">
        <attribute name="mailboxid" />
        <filter type="and">
          <condition attribute="regardingobjectid" operator="eq" value="{$QueueId}" />
        </filter>
      </entity>
    </fetch>
"@

    $Id = (Get-CrmRecordsByFetch -conn $conn -Fetch $fetch).CrmRecords[0].MailboxId
    
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
    }
    if($ScheduleTest)
    {
        $updateFields.Add("testemailconfigurationscheduled", $true)
    }
    if($MarkedAsPrimaryForExchangeSync)
    {
        $updateFields.Add("orgmarkedasprimaryforexchangesync", $true)
    }
    if($ApproveEmail)
    {
        Approve-CrmEmailAddress -conn $conn -QueueId $QueueId
    }
	
    foreach($parameter in $MyInvocation.BoundParameters.GetEnumerator())
    {   
        if($parameter.Key -in ("EmailServerProfile"))
        {
            $updateFields.Add($parameter.Key.ToLower(), (New-CrmEntityReference emailserverprofile $parameter.Value))
        }
        elseif($parameter.Key -in ("IncomingEmailDeliveryMethod","OutgoingEmailDeliveryMethod"))
        {
            $updateFields.Add($parameter.Key.ToLower(), (New-CrmOptionSetValue $parameter.Value))
        }
        elseif($parameter.Key -in ("StateCode","StatusCode"))
        {
            Set-CrmRecordState -conn $conn -EntityLogicalName mailbox -Id $Id -StateCode $StateCode -StatusCode $StatusCode
        }
        elseif($parameter.Key -in ("EmailAddress"))
        {
            $updateFields.Add($parameter.Key.ToLower(), $parameter.Value)
        }
    }
    
    Set-CrmRecord -conn $conn -EntityLogicalName mailbox -Id $Id -Fields $updateFields
}

function Set-CrmUserManager{
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml

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
            throw LastCrmConnectorException($conn)
        }
    }
    catch
    {
        throw LastCrmConnectorException($conn)
    } 
}

function Set-CrmUserSettings{
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml

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

function Set-CrmRecordAccess {
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml
    [OutputType([void])]
    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false, Position=0)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1, ParameterSetName="CrmRecord", ValueFromPipeline=$true)]
        [PSObject[]]$CrmRecord,
        [parameter(Mandatory=$true, Position=1, ParameterSetName="NameWithId")]
        [string]$EntityLogicalName,
        [parameter(Mandatory=$true, Position=2, ParameterSetName="NameWithId")]
        [guid]$Id,
        [parameter(Mandatory=$true, Position=3)]
        [Microsoft.Xrm.Sdk.EntityReference]$Principal,
        [parameter(Mandatory=$true, Position=4)]
        [Microsoft.Crm.Sdk.Messages.AccessRights]$AccessMask
    )
    begin
    {
        $conn = VerifyCrmConnectionParam $conn

        if ($EntityLogicalName) {
            $CrmRecord += [PSCustomObject] @{
                logicalname = $EntityLogicalName
                "$($EntityLogicalName)id" = $Id
            }
        }
    }
    process
    {
        foreach ($record in $CrmRecord) {
            try {
                $request = [Microsoft.Crm.Sdk.Messages.ModifyAccessRequest]::new()
                $request.Target = New-CrmEntityReference -EntityLogicalName $record.logicalname -Id $record.($record.logicalname + "id")
                $principalAccess = [Microsoft.Crm.Sdk.Messages.PrincipalAccess]::new()
                $principalAccess.Principal = $Principal
                $principalAccess.AccessMask = $AccessMask
                $request.PrincipalAccess = $principalAccess
                
                [Microsoft.Crm.Sdk.Messages.ModifyAccessResponse]$conn.Execute($request) | Out-Null
            }
            catch {
                Write-Error $_
            }   
        }
    }
}

### CRM Types ###
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

### Performance Tests ###
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
		if($View -eq $null)
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
           
            # Get all records by using Fetch
            Test-CrmTimerStart
            $records = Get-CrmRecordsByFetch -conn $conn -Fetch $View.fetchxml -AllRows -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
            $perf = Test-CrmTimerStop
            $owner = $View.ownerid
            $totalCount = $records.Count           
        }
        else
        {            
            if($RunAs -ne $null)
            {
                Set-CrmConnectionCallerId -conn $conn -CallerId $RunAs                
            }
            
			# Get all records by using Fetch
            Test-CrmTimerStart
            $records = Get-CrmRecordsByFetch -conn $conn -Fetch $View.fetchxml -AllRows -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
            $perf = Test-CrmTimerStop
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
			Write-Verbose "Setting connection caller id back to current user"
			Set-CrmConnectionCallerId -conn $conn -CallerId $RunAs                
		}

        return $psobj
    }
    catch
    {
        throw
    }
}

function Test-CrmTimerStart{
# .ExternalHelp Microsoft.Xrm.Data.PowerShell.Help.xml
    $script:crmtimer = New-Object -TypeName 'System.Diagnostics.Stopwatch'
    $script:crmtimer.Start()
}

function Test-CrmTimerStop{
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
            throw 'A connection to CRM is required, use Get-CrmConnection or one of the other connection functions to connect.'
        }
        else
        {
            $conn = $connobj.Value
        }
    }
	return $conn
}
function MapFieldTypeByFieldValue {
    PARAM(
        [Parameter(Mandatory=$true)]
        [object]$Value
    )

    $valueTypeToCrmTypeMapping = @{
        "Boolean" = [Microsoft.Xrm.Tooling.Connector.CrmFieldType]::CrmBoolean;
        "DateTime" = [Microsoft.Xrm.Tooling.Connector.CrmFieldType]::CrmDateTime;
        "Decimal" = [Microsoft.Xrm.Tooling.Connector.CrmFieldType]::CrmDecimal;
        "Single" = [Microsoft.Xrm.Tooling.Connector.CrmFieldType]::CrmFloat;
        "Money" = [Microsoft.Xrm.Tooling.Connector.CrmFieldType]::Raw;
        "Int32" = [Microsoft.Xrm.Tooling.Connector.CrmFieldType]::CrmNumber;
        "EntityReference" = [Microsoft.Xrm.Tooling.Connector.CrmFieldType]::Raw;
        "OptionSetValue" = [Microsoft.Xrm.Tooling.Connector.CrmFieldType]::Raw;
        "String" = [Microsoft.Xrm.Tooling.Connector.CrmFieldType]::String;
        "Guid" =  [Microsoft.Xrm.Tooling.Connector.CrmFieldType]::UniqueIdentifier;
    }

    # default is RAW
    $crmDataType = [Microsoft.Xrm.Tooling.Connector.CrmFieldType]::Raw

    if($Value -ne $null) {

        $valueType = $Value.GetType().Name
        
        if($valueTypeToCrmTypeMapping.ContainsKey($valueType)) {
            $crmDataType = $valueTypeToCrmTypeMapping[$valueType]
        }   
    }

    return $crmDatatype
}
function GuessPrimaryKeyField() {
    PARAM(
        [Parameter(Mandatory=$true)]
        [object]$EntityLogicalName
    )

    $standardActivityEntities = @(
        "opportunityclose",
        "socialactivity",
        "campaignresponse",
        "letter","orderclose",
        "appointment",
        "recurringappointmentmaster",
        "fax",
        "email",
        "activitypointer",
        "incidentresolution",
        "bulkoperation",
        "quoteclose",
        "task",
        "campaignactivity",
        "serviceappointment",
        "phonecall"
    )

    # Some Entity has different pattern for id name.
    if($EntityLogicalName -eq "usersettings")
    {
        $primaryKeyField = "systemuserid"
    }
    elseif($EntityLogicalName -eq "systemform")
    {
        $primaryKeyField = "formid"
    }
    elseif($EntityLogicalName -in $standardActivityEntities)
    {
        $primaryKeyField = "activityid"
    }
    else 
    {
        # default
        $primaryKeyField = $EntityLogicalName + "id"
    }
    
    $primaryKeyField
}
function LastCrmConnectorException {
	[CmdletBinding()]
    PARAM( 
        [parameter(Mandatory=$true)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn
    )

	return (Coalesce $conn.LastCrmError $conn.LastCrmException) 
}
function AddTls12Support {
	#by default PowerShell will show Ssl3, Tls - since SSL3 is not desirable we will drop it and use Tls + Tls12
	[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls -bor [System.Net.SecurityProtocolType]::Tls12
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
