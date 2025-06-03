param(
    [string]$AccessFile,
    [string]$ConfigFile
)

enum PropertyType {
    Text = 10
    Integer = 3
    Long = 4
    Boolean = 1
    DateTime = 8
}

function Set-DbProperty{
    param (
        [System.Object]$db,
        [string]$PropertyName,
        [int32]$PropertyType,
        [string]$PropertyValue,
        [bool]$RemoveProperty = $false  
    )

    try {
        $db.Properties[$PropertyName].Value = $PropertyValue    
    }
    catch  {
        # Handle the exception if the property does not exist
        # error 3270
        # The property does not exist, so we create it
        $errorCode = $null
        if ($_.Exception.InnerException -and $_.Exception.InnerException.ErrorCode) {
            $errorCode = $_.Exception.InnerException.ErrorCode
        } elseif ($_.Exception.HResult) {
            $errorCode = $_.Exception.HResult
        }
        $errorMsg = $_.Exception.Message

        # Prüfe auf Fehlercode oder auf typische Fehlermeldung
        if ($errorCode -eq -2146825018 -or $errorMsg -like "*Property not found*") {
            Write-Host "Property '$PropertyName' does not exist. Creating it."
            $db.Properties.Append($db.CreateProperty($PropertyName, $PropertyType, $PropertyValue))
        } else {
            Write-Error "An unexpected error occurred: $($errorCode)  $errorMsg"
            return
        }
    }
}

function Invoke-Procedure {
    [CmdletBinding()]   
    param (
        [System.Object]$access,
        [string]$ProcedureName,
        [Parameter(ValueFromRemainingArguments = $true)]
        [object[]]$Arguments
    )

    if (-not $access) {
        Write-Error "Access application object is null."
        return
    }
    if (-not $ProcedureName) {
        Write-Error "Procedure name is null or empty."
        return
    }

    $ArgCount = $Arguments.Count    

    switch ($ArgCount) {   
        0 { 
            $null = $access.Run($ProcedureName)
        }   
        1 { 
            $null = $access.Run($ProcedureName, [ref] $Arguments[0])
        }   
        2 { 
            $null = $access.Run($ProcedureName, [ref] $Arguments[0], [ref] $Arguments[1])
        }
        3 { 
            $null = $access.Run($ProcedureName, [ref] $Arguments[0], [ref] $Arguments[1], [ref] $Arguments[2])
        }
        Default {
            # raise error if more than 3 arguments are passed
            Write-Error "Procedure '$ProcedureName' expects at most 3 arguments, but $ArgCount were provided."
            return
        }
    }

}

# read config file
if (-not $ConfigFile) {
    $ConfigFile = Join-Path -Path (Get-Location) -ChildPath "config.json"
}

if (
    -not ([System.IO.Path]::IsPathRooted($ConfigFile))
) {
    $ConfigFile = Join-Path -Path (Get-Location) -ChildPath $ConfigFile.TrimStart('\','/','.')
}

if (-not (Test-Path -Path $ConfigFile)) {
    Write-Error "Config file not found: $ConfigFile"
    exit 1
}   
$config = Get-Content -Path $ConfigFile | ConvertFrom-Json


[string]$fullPath = $AccessFile
if (-not ([System.IO.Path]::IsPathRooted($fullPath))) {
    $fullPath = Join-Path -Path (Get-Location) -ChildPath $fullPath.TrimStart('\','/','.')
}

[object]$access = $null
[object]$db = $null

Write-Host "Config $fullPath"

try {

# Run procedures from config
    if ($config.Procedures -and $config.Procedures.Count -gt 0) {
        $access = New-Object -ComObject Access.Application
        $access.OpenCurrentDatabase($fullPath)
        foreach ($procedure in $config.Procedures) {
            if (-not $procedure.Name) {
                Write-Error "Procedure name is missing in the configuration."
                continue
            }
            $Parameters = if ($procedure.PSObject.Properties.Match('Parameters')) { $procedure.Parameters } else { @() }
            if (-not $Parameters) {
                $Parameters = @()  # Default to empty array if no parameters are defined
            }
            if ($Parameters -and $Parameters.Count -gt 0) {
                Write-Host "Running procedure '$($procedure.Name)' with parameters: $($Parameters -join ', ')"
            } else {
                Write-Host "Running procedure '$($procedure.Name)'"
            }
            Invoke-Procedure -access $access -ProcedureName $procedure.Name -Arguments $Parameters    
        }
    }
    else {
        Write-Host "No procedures to run."
    }

# Set database properties from config
    if ($config.DatabaseProperties -and $config.DatabaseProperties.Count -gt 0) {
        if ($access) {
            $db = $access.CurrentDb()
        }
        else { # use DAO.Database
            $daoEngine = New-Object -ComObject DAO.DBEngine.120
            $db = $daoEngine.OpenDatabase($fullPath)
        }

        foreach ($property in $config.DatabaseProperties) {
            $propertyName = $property.Name
            $propertyType = [PropertyType]::Parse([PropertyType], $property.Type)
            $propertyValue = $property.Value

            Write-Host "Setting property '$propertyName' of type '$($propertyType)' to '$propertyValue'"
            Set-DbProperty -db $db -PropertyName $propertyName -PropertyType $propertyType -PropertyValue $propertyValue
        }
    }
    else {
        Write-Host "No database properties to set."
    }
}
catch {
    $errorCode = $null
    if ($_.Exception.InnerException -and $_.Exception.InnerException.ErrorCode) {
        $errorCode = $_.Exception.InnerException.ErrorCode
    } elseif ($_.Exception.HResult) {
        $errorCode = $_.Exception.HResult
    }
    $errorMsg = $_.Exception.Message
    Write-Error "An error occurred while setting properties: $($errorCode)  $errorMsg"
    exit 1
}
finally {
    if ($db) {
        if (-not $access) {
            # If we used DAO, we need to close the database
            $db.Close()
        }
        [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($db)
        Remove-Variable -Name db -ErrorAction SilentlyContinue
    }
    if ($daoEngine) {
        [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($daoEngine)
        Remove-Variable -Name daoEngine -ErrorAction SilentlyContinue
    }
    if ($access) {
        $access.CloseCurrentDatabase()
        $access.Quit()
        [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($access)
        Remove-Variable -Name access -ErrorAction SilentlyContinue
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}