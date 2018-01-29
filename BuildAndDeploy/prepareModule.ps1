gci ..\  -Recurse -Verbose

"Setting variables"
$modulePath = Convert-Path ".\Microsoft.Xrm.Data.PowerShell\"
$moduleName = "Microsoft.Xrm.Data.PowerShell.psd1"
$Copyright = "(C) $((get-date).year) Microsoft Corporation All rights reserved."
$datafile = Import-PowerShellDataFile "$modulepath\$moduleName" -Verbose
$vNum = $datafile.ModuleVersion

$manifestVersion = [System.Version]::Parse($vNum)
$newBuildNumber = "$($manifestVersion.Major).$($manifestVersion.Minor)"
$datafile.Copyright = $Copyright
$datafile.ModuleVersion = $newBuildNumber

"Updating module manifest" 

Update-ModuleManifest "$modulepath\$moduleName" -Copyright $Copyright -ModuleVersion $vNum -Verbose

"Removing *.psproj and *.pshproj from $modulepath..."
try{
	$Files = Get-ChildItem $modulepath -Include *.pssproj,*.pshproj -Recurse -verbose
	foreach ($File in $Files){ 
		"Deleting File $File"
		Remove-Item $File | out-null 
	}
}
Catch{
	"Failed to cleanup specific file extensions"
}
