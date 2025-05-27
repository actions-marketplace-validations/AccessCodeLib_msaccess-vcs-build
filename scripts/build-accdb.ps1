#
# install msaccess-vcs
#
param(
    [string]$SourceDir,
	[string]$TargetDir = ".\bin"
	[string]$FileName = "" # empty = name from vcs options
)

$tempFileName = "TempApp"
if ($FileName -eq "") {
    $FileName = $tempFileName
}

$curDir = $(pwd)
$accdbPath = "$curDir\$FileName.accdb"
if (-not (Test-Path $accdbPath)) {
    $access = New-Object -ComObject Access.Application
    $access.Visible = $true
    $access.NewCurrentDatabase($accdbPath)
    $access.CloseCurrentDatabase()
    Start-Sleep -Seconds 2
    $access.Quit(0)
    $access = New-Object -ComObject Access.Application
    $access.Visible = $true
}

$access = New-Object -ComObject Access.Application
$access.Visible = $true
$access.OpenCurrentDatabase($accdbPath)

$appdata = $env:APPDATA
$addInFolder = Join-Path $appdata "Microsoft\AddIns"
$addInProcessPath = Join-Path $addInFolder "msaccess-vcs"
# $addInFolder = Join-Path $appdata "MSAccessVCS"
# $addInProcessPath = Join-Path $addInFolder "Version Control"

Write-Host "$addInProcessPath"
Write-Host "$(pwd)SourceDir"

$access.Run("$addInProcessPath.SetInteractionMode", [ref] 1)
$access.Run("$addInProcessPath.HandleRibbonCommand", [ref] "btnBuild", [ref] "$SourceDir")

# $vcs = $access.Run("$addInProcessPath.VCS")
# $vcs.Build($sourcePath)

$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
while ((Get-ChildItem -Path . -Filter *.laccdb) -and ($stopwatch.Elapsed.TotalSeconds -lt 30)) {
    Start-Sleep -Seconds 2
    Write-Host "."
}
$stopwatch.Stop()
Start-Sleep -Seconds 2
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
while ((Get-ChildItem -Path . -Filter *.laccdb) -and ($stopwatch.Elapsed.TotalSeconds -lt 10)) {
    Start-Sleep -Seconds 2
    Write-Host "."
}
$stopwatch.Stop()
$access.Quit(1)

$tempFileName = "TempApp"
if ($FileName -eq $tempFileName) {
    Remove-Item -Path "$accdbPath" -ErrorAction SilentlyContinue
}

New-Item -Path "bin" -ItemType Directory -Force
Copy-Item -Path ".\*.accdb" -Destination "$TargetDir\"
