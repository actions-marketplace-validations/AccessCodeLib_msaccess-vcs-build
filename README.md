# msaccess-vcs-build
CI/CD - Build accdb/accde from source (msaccess-vcs exports)

Thanks to Martin Leduc (DecimalTurn) for [VBA-Build](https://github.com/DecimalTurn/VBA-Build), which was a great reference.

## action.yml
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
    default: 'https://api.github.com/repos/joyfullservice/msaccess-vcs-addin/releases/latest'
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

#### Azure DevOps
* [AccessCodeLib/BuildAccdeExample: Build-self-hosted-O64](https://github.com/AccessCodeLib/BuildAccdeExample/blob/main/.azure-devops/azure-pipelines.yml)
