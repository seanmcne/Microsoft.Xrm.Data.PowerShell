function AddTls12Support {
	#by default PowerShell will show Ssl3, Tls - since SSL3 is not desirable we will drop it and use Tls + Tls12
	[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls -bor [System.Net.SecurityProtocolType]::Tls12
}
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
        [ValidatePattern('([\w-]+).crm([0-9]*).(microsoftdynamics|dynamics|crm[\w-]*).(com|de|us)')]
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
		Write-Verbose ($cs.Replace($Credential.GetNetworkCredential().Password, "")) 

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