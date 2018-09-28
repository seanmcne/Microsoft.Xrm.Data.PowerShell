function Publish-CrmDuplicateDetectionRules{
 <#
 .SYNOPSIS
  This function takes the name of duplicate detection rule as an argument, and publish it if it was unpublished.

 .DESCRIPTION
  This function takes the name of duplicate detection rule as an argument, and publish it if it was unpublished.
 
  By default, it publishes the first rule if multiple rules have the same name, but passing PublishAll 1 all 
  rules with the same name will be published.
 
 .PARAMETER conn
 A connection to your CRM organizatoin. Use $conn = Get-CrmConnection <Parameters> to generate it.

 .PARAMETER DuplicateDetectionRule
 String name of duplicate detection rules
 
 .PARAMETER PublishAll
 Optional and false by default. Passing 1 will publish all rules having the same name. 
 Omitting it will just publish the first rule if multiple rules have the same name.   
 
 .EXAMPLE 
 Example 1. Publish the first rule among matching rules:
 Publish-CrmDuplicateDetectionRules -conn $conn  -DuplicateDetectionRule "Accounts with the same Account Name" 

 Example 2. Publish all rules among matching rules:
 Publish-CrmDuplicateDetectionRules -conn $conn -DuplicateDetectionRule "Accounts with the same Account Name" -PublishAll 1

#>
    [CmdletBinding()]
    PARAM(
        [parameter(Mandatory=$false)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$conn,
        [parameter(Mandatory=$true, Position=1)]
        [string]$DuplicateDetectionRule,
        [parameter(Mandatory=$false, Position=2)]
        [bool]$PublishAll
    )
      
      $conn =  VerifyCrmConnectionParam $conn
     
      #$matchingDDRules =Get-CrmRecords  -EntityLogicalName duplicaterule -Fields * -FilterAttribute name -FilterOperator eq -FilterValue $DuplicateDetectionRule  
      
      $fetch = @"
                <fetch>
                  <entity name="duplicaterule" >
                    <all-attributes/>
                    <filter>
                      <condition attribute="statuscode" operator="eq" value="0" />
                      <condition attribute="name" operator="eq" value="Accounts with the same Account Name" />
                    </filter>
                  </entity>
                </fetch>
"@
      $matchingDDRules = Get-CrmRecordsByFetch -conn $conn -Fetch $fetch

      Write-Host $matchingDDRules.Count "rules found"    
     
      if($matchingDDRules.Count -lt 1)
      {      
        throw "Duplicate rule $DuplicateDetectionRule did not exist"
      }

      
      $PublishAll

      if($PublishAll -eq $false)
      {       Write-Host "Publishing one rule"
      $ddRule_toPublish = New-Object Microsoft.Crm.Sdk.Messages.PublishDuplicateRuleRequest
      $ddRule_toPublish.DuplicateRuleId= $matchingDDRules.CrmRecords[0].duplicateruleid
      $conn.ExecuteCrmOrganizationRequest($ddRule_toPublish,$trace)   
      }
      else
      { Write-Host "Publishing rules"
        foreach($rule in $matchingDDRules.CrmRecords)
        {
            write-host "rule is " $rule.duplicateruleid
            $ddRule_toPublish = New-Object Microsoft.Crm.Sdk.Messages.PublishDuplicateRuleRequest            
            $ddRule_toPublish.DuplicateRuleId= $rule.duplicateruleid
            $conn.ExecuteCrmOrganizationRequest($ddRule_toPublish,$trace)  
            Write-Host "Rule Published"       
        }

      }      
}

             
