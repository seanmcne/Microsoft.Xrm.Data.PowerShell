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