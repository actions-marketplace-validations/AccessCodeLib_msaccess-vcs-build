#
# install msaccess-vcs
#
param(
    [string]$vcsUrl = "https://api.github.com/repos/joyfullservice/msaccess-vcs-addin/releases/latest"
)
# URL
# $vcsUrl = "https://api.github.com/repos/joyfullservice/msaccess-vcs-addin/releases/latest"
# $vcsUrl = "https://api.github.com/repos/josef-poetzl/msaccess-vcs-addin/releases/latest"

Write-Host "Download url: $vcsUrl"

$headers = @{
    "User-Agent" = "PowerShell"
}
$release = Invoke-RestMethod -Uri $vcsUrl -Headers $headers

# zip url
$asset = $release.assets | Where-Object { $_.name -like "Version*.zip" } | Select-Object -First 1
$zipUrl = $asset.browser_download_url

# save as
$zipFile = "msaccess-vcs.zip"

# download file
Invoke-WebRequest -Uri $zipUrl -OutFile $zipFile

Write-Host "zip file downloaded from $zipUrl to $zipFile"

# extrat to %appdata%\MSAccessVCS
$appdata = $env:APPDATA
$addInFolder = Join-Path $appdata "MSAccessVCS"
Expand-Archive -Path $zipFile -DestinationPath $addInFolder -Force

$addInFileName = "Version Control.accda"
$addInPath = Join-Path $addInFolder $addInFileName

Write-Host "msaccess-vcs installed: $addInPath"

Write-Host "Set trusted location: $addInFolder"
. "$PSScriptRoot/set-trusted-location.ps1" "VCS-add-in-folder" "$addInFolder"
