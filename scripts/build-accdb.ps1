#
# build accdb from source
#
param(
    [string]$SourceDir,
	[string]$TargetDir = "bin",
	[string]$FileName = "" # empty = name from vcs options
)

$tempFileName = "VcsBuildTempApp"
$accdbFileName = $tempFileName
if ($FileName -gt "") {
    $accdbFileName = $FileName
}

$curDir = $(pwd)
$accdbPath = "$curDir\$accdbFileName.accdb"

# open/create access file
$access = New-Object -ComObject Access.Application
$access.Visible = $true
if (-not (Test-Path $accdbPath)) {    
    $access.NewCurrentDatabase($accdbPath)
} 
else {
	$access.OpenCurrentDatabase($accdbPath)
}

$appdata = $env:APPDATA
$addInFolder = Join-Path $appdata "MSAccessVCS"
$addInProcessPath = Join-Path $addInFolder "Version Control"

if (
    -not ([System.IO.Path]::IsPathRooted($SourceDir)) -or
    ($SourceDir -match "^[\\\/]") # "\source" or "/source"
) {
    $SourceDir = Join-Path -Path (Get-Location) -ChildPath $SourceDir.TrimStart('\','/','.')
}

Write-Host "add-in path: $addInProcessPath"
Write-Host "current path: $(pwd)"
Write-Host "source: $SourceDir"
Write-Host ""

Write-Host "Start msaccess-vcs build " -NoNewline
$access.Run("$addInProcessPath.SetInteractionMode", [ref] 1)
Write-Host "." -NoNewline
$null = $access.Run("$addInProcessPath.HandleRibbonCommand", [ref] "btnBuild", [ref] "$SourceDir")

# VCS Build close tempApp and reopen new accdb => check 2x for Forms.Count
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
Write-Host " completed"

$builtFileName = $access.CurrentProject.Name
Write-Host "Built: $builtFileName"

Start-Sleep -Seconds 1
Write-Host "Close Access " -NoNewline
$access.Quit(2)
Write-Host "." -NoNewline
Start-Sleep -Seconds 1
Write-Host "." -NoNewline
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($access)
Remove-Variable access
[GC]::Collect()
Write-Host "." -NoNewline
[GC]::WaitForPendingFinalizers()
Write-Host " completed"
Write-Host ""

# copy file to TargetDir
if ($FileName -eq "") {
    $FileName = $builtFileName
}

if (
    -not [string]::IsNullOrWhiteSpace($builtFileName) -and
    $builtFileName -ne "$tempFileName.accdb"
) {
	Write-Host "Copy accdb to $TargetDir"
	New-Item -Path $TargetDir -ItemType Directory -Force | Out-Null
    Copy-Item -Path ".\$builtFileName" -Destination "$TargetDir\$FileName"
	Write-Host ""
}
Write-Host "Build accdb completed" -NoNewline
