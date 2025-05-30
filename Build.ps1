param(
    [string]$SourceDir = "\source",
    [string]$Compile = "false", # Default to "false" if not specified
    [string]$vcsUrl = "https://api.github.com/repos/joyfullservice/msaccess-vcs-addin/releases/latest"
)

Write-Host "Install msaccess-vcs"
. "$PSScriptRoot/scripts/install-msaccess-vcs.ps1" "${vcsUrl}"
Write-Host "-----"

Write-Host "Build accdb"
. "$PSScriptRoot/scripts/build-accdb.ps1" "${SourceDir}"
Write-Host "-----"
