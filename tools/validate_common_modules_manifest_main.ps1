param(
    [string]$ManifestPath = (Join-Path (Join-Path (Join-Path (Join-Path $PSScriptRoot '..\CommonModules') 'src') 'CommonModules') 'common-modules-manifest.tsv'),
    [string]$ModulesDirectory = (Join-Path (Join-Path (Join-Path $PSScriptRoot '..\CommonModules') 'src') 'CommonModules'),
    [switch]$RequireAllFiles
)

$ErrorActionPreference = 'Stop'
$InformationPreference = 'Continue'

Import-Module (Join-Path $PSScriptRoot 'VbaDevTool.psm1') -Force

Set-Variable -Name COMMON_MODULE_EXTENSIONS -Value @('.bas', '.cls', '.frm') -Option Constant

function Add-ValidationError {
    param(
        [System.Collections.Generic.List[string]]$Errors,

        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    $Errors.Add($Message)
}

function Get-ModuleFileSet {
    param(
        [Parameter(Mandatory = $true)]
        [string]$DirectoryPath
    )

    $result = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $module_files = @(Get-ChildItem -LiteralPath $DirectoryPath -File | Where-Object { $COMMON_MODULE_EXTENSIONS -contains $_.Extension.ToLowerInvariant() })
    foreach ($module_file in $module_files) {
        [void]$result.Add($module_file.Name)
    }

    return $result
}

try {
    $resolved_manifest_path = (Resolve-Path -LiteralPath $ManifestPath -ErrorAction Stop).ProviderPath
    $resolved_modules_directory = (Resolve-Path -LiteralPath $ModulesDirectory -ErrorAction Stop).ProviderPath
    $modules_directory_info = Get-Item -LiteralPath $resolved_modules_directory -ErrorAction Stop
    if (-not $modules_directory_info.PSIsContainer) {
        throw "ModulesDirectory is not a directory: $resolved_modules_directory"
    }

    $errors = New-Object 'System.Collections.Generic.List[string]'
    $manifest = Read-CommonModulesManifest -ManifestPath $resolved_manifest_path
    $records = @($manifest.Records)
    $records_by_module = $manifest.RecordsByModule
    $module_file_set = Get-ModuleFileSet -DirectoryPath $resolved_modules_directory

    foreach ($record in $records) {
        if (-not $module_file_set.Contains($record.ModuleFile)) {
            Add-ValidationError -Errors $errors -Message "Line $($record.LineNumber) references unknown module file '$($record.ModuleFile)'."
        }
        foreach ($dependency in $record.Dependencies) {
            if (-not $records_by_module.ContainsKey($dependency)) {
                Add-ValidationError -Errors $errors -Message "Line $($record.LineNumber) references unknown dependency '$dependency' from '$($record.ModuleFile)'."
                continue
            }
            if (-not $module_file_set.Contains($dependency)) {
                Add-ValidationError -Errors $errors -Message "Line $($record.LineNumber) references dependency '$dependency' that is not in ModulesDirectory."
            }
        }
    }

    if ($RequireAllFiles) {
        foreach ($module_file in ($module_file_set | Sort-Object)) {
            if ($module_file.StartsWith('Test_', [System.StringComparison]::OrdinalIgnoreCase)) {
                continue
            }
            if (-not $records_by_module.ContainsKey($module_file)) {
                Add-ValidationError -Errors $errors -Message "Module file '$module_file' is missing from the manifest."
            }
        }
    }

    if ($errors.Count -gt 0) {
        foreach ($validation_error in $errors) {
            Write-Error $validation_error -ErrorAction Continue
        }
        exit 1
    }

    Write-Host "CommonModules manifest validation passed. Records: $($records.Count)."
    exit 0
}
catch {
    Write-Error $_ -ErrorAction Continue
    exit 1
}
