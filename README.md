# msaccess-vcs-build
CI/CD - Build accdb/accde from source (msaccess-vcs exports)

Thanks to Martin Leduc (DecimalTurn) for [VBA-Build](https://github.com/DecimalTurn/VBA-Build), which was a great reference.

## Github workflow / Azure devops pipeline

### action.yml
```
inputs:
  source-dir:
    description: 'msaccess-vcs source folder'
    required: false
    default: 'source'
  target-dir:
    description: 'target dir for binary file'
    required: false
    default: ''
  compile:
    description: 'create accde file'
    required: false
    default: false
  vcs-url:
    description: 'msaccess-vcs release url'
    required: false
    default: 'https://api.github.com/repos/josef-poetzl/msaccess-vcs-addin/releases/latest'
    # remove to 'https://api.github.com/repos/joyfullservice/msaccess-vcs-addin/releases/latest' if Commit db07ef2 released
```

Example call:
```
jobs:
  build:
    runs-on: [self-hosted, Windows, Office]

    steps:
    - ...

    - name: "Build Access file (accdb/accde)"
      id: build_access_file
      uses: AccessCodeLib/msaccess-vcs-build@main
      with:
        source-dir: "./Version Control.accda.src"
        target-dir: "bin"
        compile: "false"
        vcs-url: "https://api.github.com/repos/josef-poetzl/msaccess-vcs-addin/releases/tags/v4.1.2-build"
```

### Example YAML files
#### GitHub
* [josef-poetzl/msaccess-vcs-addin: Build-self-hosted (on release)](https://github.com/josef-poetzl/msaccess-vcs-addin/blob/main/.github/workflows/build-for-release.yml)
* [AccessCodeLib/BuildAccdeExample: Build-self-hosted-O64](https://github.com/AccessCodeLib/BuildAccdeExample/blob/main/.github/workflows/Build-self-hosted-O64.yml)
* [AccessCodeLib/BuildAccdeExample: Build-self-hosted-O64-O32](https://github.com/AccessCodeLib/BuildAccdeExample/blob/main/.github/workflows/Build-self-hosted-O64-O32.yml): call 64 and 32 runner to get 32 and 64 bit accde

#### Azure DevOps
* [AccessCodeLib/BuildAccdeExample: Build-self-hosted-MultiBit](https://github.com/AccessCodeLib/BuildAccdeExample/blob/main/.azure-devops/azure-pipelines.yml)

## PowerShell only - Build.ps1
It is also possible to use only the PowerShell scripts to execute the build process locally.

### Parameters
* [string]$SourceDir = "", # empty use parameter SourceFile, don't use msaccess-vcs
* [string]$SourceFile = "", # empty = name from vcs options
* [string]$TargetDir = "", # Folder for output file, default (empty): current folder 
* [string]$Compile = 'false', # Default to "false" if not specified
* [string]$AppConfigFile = "", # Default "" => don't change database properties etc.
* [string]$vcsUrl = "https://api.github.com/repos/josef-poetzl/msaccess-vcs-addin/releases/latest", # empty = don't install msacess-vcs
* [string]$SetTrustedLocation = 'true' # set trusted location for current folder

### Examples

#### Build from source
```powershell
.\Build.ps1 -SourceDir "source" -Compile $true -AppConfigFile ".\Application-Config.json"
```
Steps:
1. download msacesss-vcs
2. build accdb from source (use file name from msaccess-vcs property file)
3. compile accdb to accde
4. config accde with settings from [Application-Config.json](https://github.com/AccessCodeLib/msaccess-vcs-build/blob/main/examples/Prepare-Application-Config.json)

#### Compile accdb file
```powershell
.\Build.ps1 -SourceFile "Test.accdb" -Compile $true -AppConfigFile ".\Application-Config.json"
```
Steps:
1. compile Test.accdb to Test.accde
2. config accde with settings from Application-Config.json
