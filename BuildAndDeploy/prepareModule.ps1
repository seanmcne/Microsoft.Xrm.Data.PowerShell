gci ..\  -Recurse -Verbose

$modulePath = Convert-Path .
$moduleName = "Microsoft.Xrm.Data.PowerShell.psd1"
$Copyright = "(C) $((get-date).year) Microsoft Corporation All rights reserved."
$datafile = Import-PowerShellDataFile "$modulepath\$moduleName"
$vNum = $datafile.ModuleVersion

$manifestVersion = [System.Version]::Parse($vNum)
$newBuildNumber = "$($manifestVersion.Major).$($manifestVersion.Minor)"
$datafile.Copyright = $Copyright
$datafile.ModuleVersion = $newBuildNumber
Update-ModuleManifest "$modulepath\$moduleName" -Copyright $Copyright -ModuleVersion $vNum

"Removing *.psproj and *.pshproj from $modulepath..."
try{
	$Files = Get-ChildItem $modulepath -Include *.pssproj,*.pshproj -Recurse
	foreach ($File in $Files){ 
		"Deleting File $File"
		Remove-Item $File | out-null 
	}
}
Catch{
	"Failed to cleanup specific file extensions"
}
