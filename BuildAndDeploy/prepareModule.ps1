gci c:\users\ -Recurse 

$keypass = $args[0]
$keypath = "pshellSigning.pfx"
$Copyright = "(C) $((get-date).year) Microsoft Corporation All rights reserved."
$ModuleName = "Microsoft.Xrm.Data.PowerShell" 
$moduleFileName = "$ModuleName.psd1"

cd $ModuleName

$modulePath = Convert-Path "."

#get, parse, and update Module attributes 
$datafile = Import-PowerShellDataFile "$modulepath\$moduleFileName" -Verbose
$vNum = $datafile.ModuleVersion
$manifestVersion = [System.Version]::Parse($vNum)
$newBuildNumber = "$($manifestVersion.Major).$($manifestVersion.Minor)"
$datafile.Copyright = $Copyright
$datafile.ModuleVersion = $newBuildNumber

"Updating module manifest" 
Update-ModuleManifest "$modulepath\$moduleFileName" -Copyright $Copyright -ModuleVersion $vNum -Verbose

"Clean out *.psproj and *.pshproj from $modulepath..."
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

$Cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($keypath,$keypass)

Set-AuthenticodeSignature -Certificate $Cert -TimeStampServer http://timestamp.verisign.com/scripts/timstamp.dll -FilePath "$ModuleName\$ModuleName.psd1"
Set-AuthenticodeSignature -Certificate $Cert -TimeStampServer http://timestamp.verisign.com/scripts/timstamp.dll -FilePath "$ModuleName\$ModuleName.psm1"
