param(
    [string]$SourceDir = "\source",
    [string]$TargetDir = "",
    [string]$Compile = 'false', # Default to "false" if not specified
    [string]$vcsUrl = "https://api.github.com/repos/josef-poetzl/msaccess-vcs-addin/releases/latest" # empty = don't install msacess-vcs
)

[bool]$CompileBool = $false
if ($Compile -and $Compile.ToLower() -eq "true") {
    $CompileBool = $true
}

if ($vcsUrl -gt "") {
	Write-Host "Install msaccess-vcs"
	. "$PSScriptRoot/scripts/Install-msaccess-vcs.ps1" "${vcsUrl}"
	Write-Host "-----"
}

Write-Host "Build accdb"
$accdbPath = . "$PSScriptRoot/scripts/Build-Accdb.ps1" -SourceDir "${SourceDir}" -TargetDir "${TargetDir}"
Write-Host "-----"

$accdbPath = "$accdbPath" # simple join if array
$accdbPath = $accdbPath.Trim()

if ([string]::IsNullOrEmpty($accdbPath)) {
    Write-Error "accdbPath is null (missing return value)"
    exit 1
}

if ($CompileBool) {
    Write-Host "compile accdb"
    $compileResult = . "$PSScriptRoot/scripts/Compile-Accdb.ps1" -SourceFile "$accdbPath"
    # Write-Host "accdb: $($result.AccdbPath)"
    # Write-Host "accde: $($result.AccdePath)"
    if (-not $result.Success) {
        Write-Error "Failed to create ACCDE file"
        exit 1
    }
    Write-Host "-----"	
}

exit 0
