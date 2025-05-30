#
# create accde
#
param(
    [string]$SourceFile,
    [string]$DestFile = "" # empty .. use SourceFile with accde extension
)

$SourceFilePath = $SourceFile
if (
    -not ([System.IO.Path]::IsPathRooted($SourceFilePath)) -or
    ($SourceFilePath -match "^[\\\/]") # "\source" or "/source"
) {
    $SourceFilePath = Join-Path -Path (Get-Location) -ChildPath $SourceFilePath.TrimStart('\','/','.')
}

if ($DestFile -gt "") {
	if (
		-not ([System.IO.Path]::IsPathRooted($DestFile)) -or
		($DestFile -match "^[\\\/]") # "\source" or "/source"
	) {
		$DestFilePath = Join-Path -Path (Get-Location) -ChildPath $DestFile.TrimStart('\','/','.')
	}
	else {
		$DestFilePath = $DestFile
	}
} else {
	$DestFilePath = [System.IO.Path]::ChangeExtension($SourceFilePath, "accde")
}

Write-Host "accdb: $SourceFilePath"
Write-Host "accde: $DestFilePath"

if (Test-Path $DestFilePath) {
    Remove-Item $DestFilePath -Force
}

# workaround with vbscript
$vbscript = @"
Dim access
Set access = CreateObject("Access.Application")
access.Visible = True
access.SysCmd 603, "$SourceFilePath", "$DestFilePath"
access.Quit
Set access = Nothing
"@

$vbscriptPath = "$env:TEMP\buildaccde.vbs"
Set-Content -Path $vbscriptPath -Value $vbscript -Encoding ASCII
cscript.exe //nologo $vbscriptPath
Remove-Item $vbscriptPath

$timeout = 10
$success = $false

for ($i = 0; $i -lt $timeout; $i++) {
    if (Test-Path $DestFilePath) {
        $success = $true
        break
    }
    Start-Sleep -Seconds 1
}


if ($success) {
    Write-Host "accde successfully created."
} else {
    Write-Error "Error: accde file was not created."
}

$result = [PSCustomObject]@{
    AccdbPath = "$SourceFilePath"
    AccdePath = "$DestFilePath"
    Success   = $success
}

return $result