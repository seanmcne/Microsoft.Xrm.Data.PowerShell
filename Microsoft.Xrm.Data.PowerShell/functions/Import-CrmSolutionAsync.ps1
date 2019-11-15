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