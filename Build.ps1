param(
    [string]$SourceDir = "\source",
    [string]$TargetDir = "",
    [string]$Compile = 'false', # Default to "false" if not specified
    [string]$AppConfigFile = "", # Default "" => don't change database porperties etc.
    [string]$vcsUrl = "https://api.github.com/repos/josef-poetzl/msaccess-vcs-addin/releases/latest", # empty = don't install msacess-vcs
    [string]$SetTrustedLocation = 'true' # set trusted location for current folder
)

[bool]$CompileBool = $false
if ($Compile -and $Compile.ToLower() -eq "true") {
    $CompileBool = $true
}

[bool]$SetTrustedLocationBool = $false
if ($SetTrustedLocation -and $SetTrustedLocation.ToLower() -eq "true") {
    $SetTrustedLocationBool = $true
}

if (-not $TargetDir) {
    $TargetDir = (Get-Location).Path
}
else {
    if (-not ([System.IO.Path]::IsPathRooted($TargetDir))) {
        $TargetDir = Join-Path -Path (Get-Location) -ChildPath $TargetDir.TrimStart('\','/','.')
    }
}

$curDir = $(Get-Location)
$tempTrustedLocationName = "VCS-build-folder_"  + (Get-Date -Format "yyyyMMddHHmmss")
if ($SetTrustedLocationBool)
{
    Write-Host "Set trusted location: $curDir"
    . "$PSScriptRoot/scripts/Set-TrustedLocation.ps1" "$tempTrustedLocationName" "$curDir"
    Write-Host "-----"
}

[string]$vcsAddInPath = ""
if ($vcsUrl -gt "") {
	Write-Host "Install msaccess-vcs"
    $vcsTargetDir = $curDir.Path
	$vcsInstallData = . "$PSScriptRoot/scripts/Install-msaccess-vcs.ps1" -vcsUrl "${vcsUrl}" -TargetDir "$vcsTargetDir" -SetTrustedLocation $false
    $vcsAddInPath = $vcsInstallData.AddInPath
	Write-Host "-----"
}

Write-Host "Build accdb"
$accdbPath = . "$PSScriptRoot/scripts/Build-Accdb.ps1" -SourceDir "${SourceDir}" -TargetDir "${TargetDir}" -VcsAddInPath "$vcsAddInPath" -FileName "VcsBuildTempApp"
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


if ($SetTrustedLocationBool)
{
    Write-Host "Remove trusted location: $curDir"
    . "$PSScriptRoot/scripts/Remove-TrustedLocation.ps1" "$tempTrustedLocationName" 
    Write-Host "-----"
}

exit 0
