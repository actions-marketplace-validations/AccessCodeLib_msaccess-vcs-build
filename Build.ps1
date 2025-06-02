param(
    [string]$SourceDir = "\source",
    [string]$TargetDir = "",
    [string]$Compile = 'false', # Default to "false" if not specified
    [string]$AppConfigFile = "", # Default "" => don't change database porperties etc.
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

$accFilePath = $accdbPath

if ($CompileBool) {
    Write-Host "compile accdb"
    $compileResult = . "$PSScriptRoot/scripts/Compile-Accdb.ps1" -SourceFile "$accdbPath"
    # Write-Host "accdb: $($result.AccdbPath)"
    # Write-Host "accde: $($result.AccdePath)"
    if (-not $compileResult.Success) {
        Write-Error "Failed to create ACCDE file"
        exit 1
    }
    $accFilePath = $compileResult.AccdePath
    Write-Host "-----"	
}

if ($AppConfigFile -gt "") {
    Write-Host "Run procedures from config file: $AppConfigFile"
    . "$PSScriptRoot/scripts/Prepare-Application.ps1" -AccessFile "$accFilePath" -ConfigFile "$AppConfigFile"
    Write-Host "-----"
}   

exit 0
