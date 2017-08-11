# PSake makes variables declared here available in other scriptblocks
# Init some things
Properties {
    # Find the build folder based on build system
    $ProjectRoot = $ENV:BHProjectPath
	Write-Verbose "ProjectRoot: $ProjectRoot"
    if(-not $ProjectRoot)
    {
		$ProjectRoot = Resolve-Path "$PSScriptRoot\.."
        #$ProjectRoot = $PSScriptRoot
    }

    $Timestamp = Get-date -uformat "%Y%m%d-%H%M%S"
    $PSVersion = $PSVersionTable.PSVersion.Major
    $TestFile = "TestResults_PS$PSVersion`_$TimeStamp.xml"
    $lines = '----------------------------------------------------------------------'

    $Verbose = @{}
    if($ENV:BHCommitMessage -match "!verbose")
    {
        $Verbose = @{Verbose = $True}
    }
}

Task Default -Depends Deploy

Task Init {
    $lines
    Set-Location $ProjectRoot
    "Build System Details:"
    Get-Item ENV:BH*
    "`n"
}

Task Test -Depends Init  {
    $lines
    "`n`tSTATUS: Testing with PowerShell $PSVersion"

    # Gather test results. Store them in a variable and file

	$path = "$ProjectRoot\Tests"; 
	if(Test-Path $path){
		$TestResults = Invoke-Pester -Path $path -PassThru -OutputFormat NUnitXml -OutputFile "$ProjectRoot\$TestFile"
		# In Appveyor?  Upload our tests! #Abstract this into a function?
		If($ENV:BHBuildSystem -eq 'AppVeyor')
		{
			(New-Object 'System.Net.WebClient').UploadFile(
				"https://ci.appveyor.com/api/testresults/nunit/$($env:APPVEYOR_JOB_ID)",
				"$ProjectRoot\$TestFile" )
		}

		Remove-Item "$ProjectRoot\$TestFile" -Force -ErrorAction SilentlyContinue

		# Failed tests?
		# Need to tell psake or it will proceed to the deployment. Danger!
		if($TestResults.FailedCount -gt 0)
		{
			Write-Error "Failed '$($TestResults.FailedCount)' tests, build failed"
		}
		"`n"
	}
	else{
		Write-Warning "No tests found at $path - Skipping Tests"
	}
}

Task Build -Depends Test {
    $lines
    
    # Load the module, read the exported functions, update the psd1 FunctionsToExport
    Set-ModuleFunctions

    # Increase the module version
    Try
    {
        #$Version = Get-NextPSGalleryVersion -Name $env:BHProjectName -ErrorAction Stop
        #Update-Metadata -Path $env:BHPSModuleManifest -PropertyName ModuleVersion -Value $Version -ErrorAction stop
	$Copyright = "(C) $((get-date).year) Microsoft Corporation All rights reserved."
	
	#$vNum = (import-powershelldatafile $env:BHPSModuleManifest).ModuleVersion
	#$manifestVersion = [System.Version]::Parse($vNum)
	#$dateVersion = (Get-Date -format yy.Mdd.HHmm)
	#$newBuildNumber = "$($manifestVersion.Major).$dateVersion"
	
	#set the build version in PoweShellGallery identical to the manifest
	$vNum = (import-powershelldatafile $env:BHPSModuleManifest).ModuleVersion
	$manifestVersion = [System.Version]::Parse($vNum)
	$newBuildNumber = "$($manifestVersion.Major).$($manifestVersion.Minor)"
	Update-AppveyorBuild -Version $newBuildNumber
	
	Update-Metadata -Path $env:BHPSModuleManifest -PropertyName ModuleVersion -Value $newBuildNumber -ErrorAction stop
	Update-Metadata -Path $env:BHPSModuleManifest -PropertyName Copyright -Value $Copyright -ErrorAction stop
	
	#remove files that do not belong in the release
	"Removing *.psproj and *.pshproj from $ENV:BHModulePath..."
	try{
		$Files = Get-ChildItem $ENV:BHModulePath -Include *.pssproj,*.pshproj -Recurse
		foreach ($File in $Files){ 
			"Deleting File $File"
			Remove-Item $File | out-null 
		}
	}
	Catch{
		"Failed to cleanup specific file extensions"
	}
    }
    Catch{
        "Failed to update version for '$env:BHProjectName': $_.`nContinuing with existing version"
    }
}

Task Deploy -Depends Build {
    $lines
	#Update-ModuleManifest -CmdletsToExport * -Path $ENV:BHPSModuleManifest

	Write-Verbose "ProjectRoot: $ProjectRoot"
	Write-Verbose "PSScriptRoot: $PSScriptRoot"
	
    $Params = @{
        Path = $PSScriptRoot #$ProjectRoot
        Force = $true
        Recurse = $false # We keep psdeploy artifacts, avoid deploying those : )
    }
    Invoke-PSDeploy @Verbose @Params
}
