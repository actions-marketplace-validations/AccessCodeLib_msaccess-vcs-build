# install msaccess-vcs
#
# URL
# $vcsUrl = "https://api.github.com/repos/joyfullservice/msaccess-vcs-addin/releases/latest"
$vcsUrl = "https://api.github.com/repos/josef-poetzl/msaccess-vcs-addin/releases/latest"
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

# extrat to %appdata%\Microsoft\AddIns
$appdata = $env:APPDATA
$addInFolder = Join-Path $appdata "Microsoft\AddIns"
Expand-Archive -Path $zipFile -DestinationPath $addInFolder -Force

$addInFileName = "msaccess-vcs.accda"
$addInPath = Join-Path $addInFolder $addInFileName
$unzipVcsFile = Join-Path $addInFolder "Version Control.accda"
if (Test-Path $addInPath) {
    Remove-Item $addInPath -Force
}
Rename-Item -Path $unzipVcsFile -NewName $addInFileName -Force

Write-Host "msaccess-vcs installed: $addInPath"

Write-Host "Set trutsted location: $addInFolder"
. "$PSScriptRoot/set-trusted-location.ps1" "Add-in-folder" "$addInFolder"
