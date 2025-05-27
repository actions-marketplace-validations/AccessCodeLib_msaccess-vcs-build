#
# build accdb from source
#
param(
    [string]$SourceDir,
	[string]$TargetDir = "bin",
	[string]$FileName = "" # empty = name from vcs options
)

$tempFileName = "TempApp"
$accdbFileName = $tempFileName
if ($FileName -gt "") {
    $accdbFileName = $FileName
}

$curDir = $(pwd)
$accdbPath = "$curDir\$accdbFileName.accdb"
if (-not (Test-Path $accdbPath)) {
    $access = New-Object -ComObject Access.Application
    $access.Visible = $true
    $access.NewCurrentDatabase($accdbPath)
    $access.CloseCurrentDatabase()
    Start-Sleep -Seconds 2
    $access.Quit(0)
	$access = $null
	Start-Sleep -Seconds 2
}

$access = New-Object -ComObject Access.Application
$access.Visible = $true
$access.OpenCurrentDatabase($accdbPath)

$appdata = $env:APPDATA
$addInFolder = Join-Path $appdata "MSAccessVCS"
$addInProcessPath = Join-Path $addInFolder "Version Control"

if (
    -not ([System.IO.Path]::IsPathRooted($SourceDir)) -or
    ($SourceDir -match "^[\\\/]") # "\source" or "/source"
) {
    $SourceDir = Join-Path -Path (Get-Location) -ChildPath $SourceDir.TrimStart('\','/')
}

Write-Host "add-in path: $addInProcessPath"
Write-Host "current path: $(pwd)"
Write-Host "source: $SourceDir"

$access.Run("$addInProcessPath.SetInteractionMode", [ref] 1)
$access.Run("$addInProcessPath.HandleRibbonCommand", [ref] "btnBuild", [ref] "$SourceDir")

# $vcs = $access.Run("$addInProcessPath.VCS")
# $vcs.Build($SourceDir)
# $vcs = $null

$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
while (($access.Forms.Count -gt 0) -and ($stopwatch.Elapsed.TotalSeconds -lt 30)) {
    Start-Sleep -Seconds 2
    Write-Host "." -NoNewline
}
$stopwatch.Stop()
Start-Sleep -Seconds 3
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
while (($access.Forms.Count -gt 0) -and ($stopwatch.Elapsed.TotalSeconds -lt 30)) {
    Start-Sleep -Seconds 2
    Write-Host "." -NoNewline
}
$stopwatch.Stop()
Start-Sleep -Seconds 3
Write-Host ""

$builtFileName = $access.CurrentProject.Name

$access.Quit(1)
Start-Sleep -Seconds 3

Write-Host "Built: $builtFileName"

if ($FileName -eq "") {
    $FileName = $builtFileName
}

if (
    -not [string]::IsNullOrWhiteSpace($builtFileName) -and
    $builtFileName -ne "$tempFileName.accdb"
) {
	New-Item -Path $TargetDir -ItemType Directory -Force | Out-Null
    Copy-Item -Path ".\$builtFileName" -Destination "$TargetDir\$FileName"
}