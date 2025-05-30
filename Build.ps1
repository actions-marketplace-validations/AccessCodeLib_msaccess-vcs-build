param(
    [string]$SourceDir = "\source",
	[string]$TargetDir = "",
    [bool]$Compile = $false, # Default to "false" if not specified
    [string]$vcsUrl = "https://api.github.com/repos/joyfullservice/msaccess-vcs-addin/releases/latest" # empty = don't install msacess-vcs
)

if ($vcsUrl -gt "") {
	Write-Host "Install msaccess-vcs"
	. "$PSScriptRoot/scripts/install-msaccess-vcs.ps1" "${vcsUrl}"
	Write-Host "-----"
}

Write-Host "Build accdb"
$accdbPath = . "$PSScriptRoot/scripts/build-accdb.ps1" -SourceDir "${SourceDir}" -TargetDir "${TargetDir}"
Write-Host "-----"

$accdbPath = "$accdbPath" # simple join if array
$accdbPath = $accdbPath.Trim()

if ([string]::IsNullOrEmpty($accdbPath)) {
    Write-Error "accdbPath is null (missing return value)"
    exit 1
}

if ($Compile) {
	Write-Host "compile accdb"
	$compileResult = . "$PSScriptRoot/scripts/compile-accdb.ps1" -SourceFile "$accdbPath"
	# Write-Host "accdb: $($result.AccdbPath)"
	# Write-Host "accde: $($result.AccdePath)"
	if (-not $result.Success) {
		Write-Error "Failed to create ACCDE file"
		exit 1
	}
	Write-Host "-----"	
}

exit 0
