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