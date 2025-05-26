#
# install msaccess-vcs
#
param(
    [string]$SourceDir,
	[string]$FileName = "TempApp"
)

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

$access.Run("$addInProcessPath.SetInteractionMode", [ref] 1)
$access.Run("$addInProcessPath.HandleRibbonCommand", [ref] "btnBuild", [ref] "$SourceDir")

# $vcs = $access.Run("$addInProcessPath.VCS")
# $vcs.Build($sourcePath)

while (Get-ChildItem -Path . -Filter *.laccdb) {
    Start-Sleep -Seconds 2
	Write-Host "."
}
Start-Sleep -Seconds 2
while (Get-ChildItem -Path . -Filter *.laccdb) {
    Start-Sleep -Seconds 2
	Write-Host "."
}
$access.Quit(1)