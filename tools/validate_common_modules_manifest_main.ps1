param(
    [string]$ManifestPath = (Join-Path (Join-Path (Join-Path (Join-Path $PSScriptRoot '..\CommonModules') 'src') 'CommonModules') 'common-modules-manifest.tsv'),
    [string]$ModulesDirectory = (Join-Path (Join-Path (Join-Path $PSScriptRoot '..\CommonModules') 'src') 'CommonModules'),
    [switch]$RequireAllFiles
)

$ErrorActionPreference = 'Stop'
$InformationPreference = 'Continue'

Set-Variable -Name EXPECTED_HEADER -Value (('ModuleFile', 'Categories', 'Dependencies') -join "`t") -Option Constant
Set-Variable -Name ALLOWED_CATEGORIES -Value @(
    'runtime-baseline',
    'test-foundation',
    'optional',
    'test-double',
    'public-udf'
) -Option Constant
Set-Variable -Name COMMON_MODULE_EXTENSIONS -Value @('.bas', '.cls', '.frm') -Option Constant

function Add-ValidationError {
    param(
        [System.Collections.Generic.List[string]]$Errors,

        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    $Errors.Add($Message)
}

function Split-ManifestList {
    param(
        [string]$Value
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return @()
    }

    return @($Value.Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' })
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

function Read-CommonModulesManifest {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [System.Collections.Generic.List[string]]$Errors
    )

    $records = New-Object 'System.Collections.Generic.Dictionary[string, object]' ([System.StringComparer]::OrdinalIgnoreCase)
    $encoding = [System.Text.Encoding]::GetEncoding(932)
    $text = [System.IO.File]::ReadAllText($Path, $encoding)
    $lines = $text -split "`r?`n"
    $header_found = $false

    for ($i = 0; $i -lt $lines.Count; $i++) {
        $line_number = $i + 1
        $line = $lines[$i]
        if ($i -eq ($lines.Count - 1) -and $line -eq '') {
            continue
        }
        if ([string]::IsNullOrWhiteSpace($line)) {
            continue
        }
        if ($line.TrimStart().StartsWith('#')) {
            continue
        }

        if (-not $header_found) {
            if ($line -ne $EXPECTED_HEADER) {
                Add-ValidationError -Errors $Errors -Message "Line $line_number has an invalid header. Expected '$EXPECTED_HEADER'."
            }
            $header_found = $true
            continue
        }

        $columns = $line.Split("`t")
        if ($columns.Count -ne 3) {
            Add-ValidationError -Errors $Errors -Message "Line $line_number must contain exactly 3 tab-separated columns."
            continue
        }

        $module_file = $columns[0].Trim()
        $categories = @(Split-ManifestList -Value $columns[1])
        $dependencies = @(Split-ManifestList -Value $columns[2])

        if ([string]::IsNullOrWhiteSpace($module_file)) {
            Add-ValidationError -Errors $Errors -Message "Line $line_number has an empty ModuleFile value."
            continue
        }
        if ($records.ContainsKey($module_file)) {
            Add-ValidationError -Errors $Errors -Message "Line $line_number duplicates ModuleFile '$module_file'."
            continue
        }
        if ($categories.Count -eq 0) {
            Add-ValidationError -Errors $Errors -Message "Line $line_number has no category for '$module_file'."
        }

        $seen_categories = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($category in $categories) {
            if ($ALLOWED_CATEGORIES -notcontains $category) {
                Add-ValidationError -Errors $Errors -Message "Line $line_number uses unknown category '$category' for '$module_file'."
            }
            if (-not $seen_categories.Add($category)) {
                Add-ValidationError -Errors $Errors -Message "Line $line_number duplicates category '$category' for '$module_file'."
            }
        }

        $seen_dependencies = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($dependency in $dependencies) {
            if (-not $seen_dependencies.Add($dependency)) {
                Add-ValidationError -Errors $Errors -Message "Line $line_number duplicates dependency '$dependency' for '$module_file'."
            }
            if ([string]::Equals($module_file, $dependency, [System.StringComparison]::OrdinalIgnoreCase)) {
                Add-ValidationError -Errors $Errors -Message "Line $line_number declares a self dependency for '$module_file'."
            }
        }

        $records.Add($module_file, [pscustomobject]@{
            ModuleFile = $module_file
            Categories = $categories
            Dependencies = $dependencies
            LineNumber = $line_number
        })
    }

    if (-not $header_found) {
        Add-ValidationError -Errors $Errors -Message 'Manifest header was not found.'
    }
    if ($records.Count -eq 0) {
        Add-ValidationError -Errors $Errors -Message 'Manifest contains no module records.'
    }

    return $records
}

try {
    $resolved_manifest_path = (Resolve-Path -LiteralPath $ManifestPath -ErrorAction Stop).ProviderPath
    $resolved_modules_directory = (Resolve-Path -LiteralPath $ModulesDirectory -ErrorAction Stop).ProviderPath
    $modules_directory_info = Get-Item -LiteralPath $resolved_modules_directory -ErrorAction Stop
    if (-not $modules_directory_info.PSIsContainer) {
        throw "ModulesDirectory is not a directory: $resolved_modules_directory"
    }

    $errors = New-Object 'System.Collections.Generic.List[string]'
    $records = Read-CommonModulesManifest -Path $resolved_manifest_path -Errors $errors
    $module_file_set = Get-ModuleFileSet -DirectoryPath $resolved_modules_directory

    foreach ($record in $records.Values) {
        if (-not $module_file_set.Contains($record.ModuleFile)) {
            Add-ValidationError -Errors $errors -Message "Line $($record.LineNumber) references unknown module file '$($record.ModuleFile)'."
        }
        foreach ($dependency in $record.Dependencies) {
            if (-not $records.ContainsKey($dependency)) {
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
            if (-not $records.ContainsKey($module_file)) {
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
