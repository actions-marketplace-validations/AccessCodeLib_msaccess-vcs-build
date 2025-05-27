param(
    [string]$SourceDir = "\source",
    [string]$Compile = "false" # Default to "false" if not specified
)

Write-Host "Install msaccess-vcs"
. "$PSScriptRoot/scripts/install-msaccess-vcs.ps1"
Write-Host "-----"
Write-Host "build accda"
. "$PSScriptRoot/scripts/build-accdb.ps1" "${SourceDir}"
Write-Host "-----"
