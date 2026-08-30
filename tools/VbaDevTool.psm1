function Get-RepositoryRoot {
    return (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).ProviderPath
}

function Resolve-VbaDevExecutable {
    $searched_paths = New-Object 'System.Collections.Generic.List[string]'

    if (-not [string]::IsNullOrWhiteSpace($env:VBA_DEV_EXE)) {
        $searched_paths.Add("VBA_DEV_EXE=$env:VBA_DEV_EXE")
        if (Test-Path -LiteralPath $env:VBA_DEV_EXE -PathType Leaf) {
            return (Resolve-Path -LiteralPath $env:VBA_DEV_EXE).ProviderPath
        }
    }

    $bundled_executable = Join-Path $PSScriptRoot 'vba-dev\vba-dev.exe'
    $searched_paths.Add($bundled_executable)
    if (Test-Path -LiteralPath $bundled_executable -PathType Leaf) {
        return (Resolve-Path -LiteralPath $bundled_executable).ProviderPath
    }

    $path_command = Get-Command 'vba-dev.exe' -ErrorAction SilentlyContinue
    $searched_paths.Add('PATH:vba-dev.exe')
    if ($null -ne $path_command) {
        return $path_command.Source
    }

    throw "vba-dev.exe was not found. Searched locations: $($searched_paths -join '; ')"
}

function Invoke-VbaDev {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Arguments
    )

    $exe_path = Resolve-VbaDevExecutable
    Write-Information "vba-dev $($Arguments -join ' ')"
    & $exe_path @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "vba-dev failed with exit code ${LASTEXITCODE}: $($Arguments -join ' ')"
    }
}

function Get-RequiredFileTarget {
    param(
        [string]$Path,
        [Parameter(Mandatory = $true)]
        [string]$Description
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        throw "$Description is empty."
    }

    $resolved_path = (Resolve-Path -LiteralPath $Path -ErrorAction Stop).ProviderPath
    $item = Get-Item -LiteralPath $resolved_path -ErrorAction Stop
    if ($item.PSIsContainer) {
        throw "$Description must be a file, not a directory: $resolved_path"
    }

    return $item
}

function Get-RequiredDirectoryTarget {
    param(
        [string]$Path,
        [Parameter(Mandatory = $true)]
        [string]$Description
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        throw "$Description is empty."
    }

    $resolved_path = (Resolve-Path -LiteralPath $Path -ErrorAction Stop).ProviderPath
    $item = Get-Item -LiteralPath $resolved_path -Force -ErrorAction Stop
    if (-not $item.PSIsContainer) {
        throw "$Description must be a directory: $resolved_path"
    }

    return $item
}

function Get-WorkbookLocalSourceSetPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$WorkbookPath
    )

    return (Join-Path (Split-Path -Parent $WorkbookPath) 'modules')
}

function Resolve-VbaDevProjectPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProjectRoot,
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if ([System.IO.Path]::IsPathRooted($Path)) {
        return [System.IO.Path]::GetFullPath($Path)
    }

    return [System.IO.Path]::GetFullPath((Join-Path $ProjectRoot ($Path -replace '/', [System.IO.Path]::DirectorySeparatorChar)))
}

function Read-VbaDevProjectManifest {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProjectRoot
    )

    $manifest_path = Join-Path $ProjectRoot 'vba-project.json'
    if (-not (Test-Path -LiteralPath $manifest_path -PathType Leaf)) {
        throw "vba-project.json was not found: $manifest_path"
    }

    return (Get-Content -LiteralPath $manifest_path -Raw | ConvertFrom-Json)
}

function Get-VbaDevDocumentEntries {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Manifest
    )

    foreach ($property in $Manifest.documents.PSObject.Properties) {
        [pscustomobject]@{
            Name = $property.Name
            Document = $property.Value
        }
    }
}

function Get-VbaDevProjectDocumentContext {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProjectRoot,
        [Parameter(Mandatory = $true)]
        [string]$DocumentName
    )

    $resolved_project_root = (Resolve-Path -LiteralPath $ProjectRoot -ErrorAction Stop).ProviderPath
    $manifest = Read-VbaDevProjectManifest -ProjectRoot $resolved_project_root
    $document_entry = Get-VbaDevDocumentEntries -Manifest $manifest | Where-Object {
        [string]::Equals($_.Name, $DocumentName, [System.StringComparison]::OrdinalIgnoreCase)
    } | Select-Object -First 1
    if ($null -eq $document_entry) {
        throw "Document '$DocumentName' is not defined in vba-project.json: $resolved_project_root"
    }

    [pscustomobject]@{
        ProjectRoot = $resolved_project_root
        ProjectName = $manifest.projectName
        DocumentName = $document_entry.Name
        SourceSetPath = Resolve-VbaDevProjectPath -ProjectRoot $resolved_project_root -Path $document_entry.Document.sourcePath
        TemplatePath = Resolve-VbaDevProjectPath -ProjectRoot $resolved_project_root -Path $document_entry.Document.templatePath
    }
}

function Get-VbaDevProjectDocumentContexts {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProjectRoot
    )

    $resolved_project_root = (Resolve-Path -LiteralPath $ProjectRoot -ErrorAction Stop).ProviderPath
    $manifest = Read-VbaDevProjectManifest -ProjectRoot $resolved_project_root
    foreach ($entry in Get-VbaDevDocumentEntries -Manifest $manifest) {
        Get-VbaDevProjectDocumentContext -ProjectRoot $resolved_project_root -DocumentName $entry.Name
    }
}

function Get-VbaDevProjectRoots {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SearchRoot
    )

    $resolved_search_root = (Resolve-Path -LiteralPath $SearchRoot -ErrorAction Stop).ProviderPath
    $excluded_directory_names = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($name in @('.git', 'bin', 'publish', 'node_modules')) {
        [void]$excluded_directory_names.Add($name)
    }

    $pending = New-Object 'System.Collections.Generic.Queue[System.IO.DirectoryInfo]'
    $pending.Enqueue((Get-Item -LiteralPath $resolved_search_root))

    while ($pending.Count -gt 0) {
        $directory = $pending.Dequeue()
        if (Test-Path -LiteralPath (Join-Path $directory.FullName 'vba-project.json') -PathType Leaf) {
            $directory.FullName
        }

        foreach ($child in Get-ChildItem -LiteralPath $directory.FullName -Directory -Force) {
            if ($excluded_directory_names.Contains($child.Name)) {
                continue
            }
            $pending.Enqueue($child)
        }
    }
}

function Test-SamePath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$LeftPath,
        [Parameter(Mandatory = $true)]
        [string]$RightPath
    )

    $left_full_name = [System.IO.Path]::GetFullPath($LeftPath).TrimEnd([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar)
    $right_full_name = [System.IO.Path]::GetFullPath($RightPath).TrimEnd([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar)
    return [string]::Equals($left_full_name, $right_full_name, [System.StringComparison]::OrdinalIgnoreCase)
}

function Test-FileContentEqual {
    param(
        [Parameter(Mandatory = $true)]
        [string]$LeftPath,
        [Parameter(Mandatory = $true)]
        [string]$RightPath
    )

    if (-not (Test-Path -LiteralPath $LeftPath -PathType Leaf) -or -not (Test-Path -LiteralPath $RightPath -PathType Leaf)) {
        return $false
    }

    $left_bytes = [System.IO.File]::ReadAllBytes((Resolve-Path -LiteralPath $LeftPath).ProviderPath)
    $right_bytes = [System.IO.File]::ReadAllBytes((Resolve-Path -LiteralPath $RightPath).ProviderPath)
    if ($left_bytes.Length -ne $right_bytes.Length) {
        return $false
    }

    for ($i = 0; $i -lt $left_bytes.Length; $i++) {
        if ($left_bytes[$i] -ne $right_bytes[$i]) {
            return $false
        }
    }

    return $true
}

function Get-VbaSourceFilePaths {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourceSetPath
    )

    Get-ChildItem -LiteralPath $SourceSetPath -File | Where-Object {
        $_.Extension -in @('.bas', '.cls', '.frm')
    } | Sort-Object -Property Name | ForEach-Object { $_.FullName }
}

function Get-DocumentationOwnerContext {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourceSetPath
    )

    $source_directory = Get-RequiredDirectoryTarget -Path $SourceSetPath -Description 'source set directory'
    $current = $source_directory
    while ($null -ne $current) {
        $manifest_path = Join-Path $current.FullName 'vba-project.json'
        if (Test-Path -LiteralPath $manifest_path -PathType Leaf) {
            $manifest = Read-VbaDevProjectManifest -ProjectRoot $current.FullName
            foreach ($entry in Get-VbaDevDocumentEntries -Manifest $manifest) {
                $entry_source_path = Resolve-VbaDevProjectPath -ProjectRoot $current.FullName -Path $entry.Document.sourcePath
                if (Test-SamePath -LeftPath $source_directory.FullName -RightPath $entry_source_path) {
                    return [pscustomobject]@{
                        SourceSetPath = $source_directory.FullName
                        OwnerDirectory = $current.FullName
                        ProjectName = $manifest.projectName
                    }
                }
            }
        }
        $current = $current.Parent
    }

    $owner_directory = $source_directory.Parent
    if ($null -ne $owner_directory -and [string]::Equals($owner_directory.Name, 'src', [System.StringComparison]::OrdinalIgnoreCase) -and $null -ne $owner_directory.Parent) {
        $owner_directory = $owner_directory.Parent
    }
    if ($null -eq $owner_directory) {
        throw "Documentation owner directory could not be resolved for: $($source_directory.FullName)"
    }

    return [pscustomobject]@{
        SourceSetPath = $source_directory.FullName
        OwnerDirectory = $owner_directory.FullName
        ProjectName = $owner_directory.Name
    }
}

function Test-OrdinalStringInSet {
    param(
        [AllowEmptyString()]
        [string]$Value,

        [Parameter(Mandatory = $true)]
        [string[]]$AllowedValues
    )

    foreach ($allowed_value in $AllowedValues) {
        if ([string]::Equals($Value, $allowed_value, [System.StringComparison]::Ordinal)) {
            return $true
        }
    }

    return $false
}

function Test-JsonEscapedSurrogatePairs {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Value
    )

    $index = 1
    while ($index -lt $Value.Length - 1) {
        if ($Value[$index] -ne '\') {
            $index++
            continue
        }

        if ($Value[$index + 1] -ne 'u') {
            $index += 2
            continue
        }

        $code_unit = [System.Convert]::ToInt32($Value.Substring($index + 2, 4), 16)
        $index += 6
        if ($code_unit -ge 0xd800 -and $code_unit -le 0xdbff) {
            if ($index + 5 -ge $Value.Length -or
                $Value[$index] -ne '\' -or
                $Value[$index + 1] -ne 'u') {
                return $false
            }

            $low_surrogate = [System.Convert]::ToInt32($Value.Substring($index + 2, 4), 16)
            if ($low_surrogate -lt 0xdc00 -or $low_surrogate -gt 0xdfff) {
                return $false
            }
            $index += 6
            continue
        }

        if ($code_unit -ge 0xdc00 -and $code_unit -le 0xdfff) {
            return $false
        }
    }

    return $true
}

function Read-CommonModulesManifest {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ManifestPath
    )

    $resolved_path = (Resolve-Path -LiteralPath $ManifestPath -ErrorAction Stop).ProviderPath
    $bytes = [System.IO.File]::ReadAllBytes($resolved_path)
    if ($bytes.Length -lt 2 -or $bytes[0] -ne 0xFF -or $bytes[1] -ne 0xFE) {
        throw "CommonModules manifest must use UTF-16LE with a BOM: $resolved_path"
    }
    if ((($bytes.Length - 2) % 2) -ne 0) {
        throw "CommonModules manifest contains an invalid UTF-16LE byte count: $resolved_path"
    }

    $encoding = New-Object System.Text.UnicodeEncoding($false, $true, $true)
    try {
        $text = $encoding.GetString($bytes, 2, $bytes.Length - 2)
    }
    catch [System.Text.DecoderFallbackException] {
        throw "CommonModules manifest contains invalid UTF-16LE text: $resolved_path"
    }

    if (-not $text.EndsWith("`r`n", [System.StringComparison]::Ordinal)) {
        throw "CommonModules manifest must end with exactly one CRLF: $resolved_path"
    }
    if ([regex]::IsMatch($text, "(?<!`r)`n|`r(?!`n)")) {
        throw "CommonModules manifest must use CRLF line endings throughout: $resolved_path"
    }

    $body = $text.Substring(0, $text.Length - 2)
    if ($body.EndsWith("`r`n", [System.StringComparison]::Ordinal)) {
        throw "CommonModules manifest must end with exactly one CRLF: $resolved_path"
    }
    $lines = @($body -split "`r`n")

    $line_index = 0
    while ($line_index -lt $lines.Count -and $lines[$line_index].StartsWith('#', [System.StringComparison]::Ordinal)) {
        $comment = $lines[$line_index]
        if ([regex]::IsMatch($comment, '\p{Cc}') -or [char]::IsWhiteSpace($comment[$comment.Length - 1])) {
            throw "CommonModules manifest line $($line_index + 1) contains an invalid comment."
        }
        $line_index++
    }

    $expected_header = "ModuleFile`tCategories`tDependencies`tRequiredReferences"
    if ($line_index -ge $lines.Count -or -not [string]::Equals($lines[$line_index], $expected_header, [System.StringComparison]::Ordinal)) {
        throw "CommonModules manifest has an invalid header: $resolved_path"
    }
    $line_index++
    if ($line_index -ge $lines.Count) {
        throw "CommonModules manifest contains no module rows: $resolved_path"
    }

    $records = New-Object 'System.Collections.Generic.List[object]'
    $records_by_module = New-Object 'System.Collections.Generic.Dictionary[string, object]' ([System.StringComparer]::OrdinalIgnoreCase)
    while ($line_index -lt $lines.Count) {
        $line = $lines[$line_index]
        $line_number = $line_index + 1
        if ([string]::Equals($line, $expected_header, [System.StringComparison]::Ordinal)) {
            throw "CommonModules manifest line $line_number contains a duplicate header."
        }
        $columns = $line.Split("`t")
        if ($columns.Count -ne 4) {
            throw "CommonModules manifest line $line_number must contain exactly 4 tab-separated columns."
        }

        $module_file = $columns[0]
        if ([string]::IsNullOrWhiteSpace($module_file)) {
            throw "CommonModules manifest line $line_number has an empty ModuleFile value."
        }
        if ([regex]::IsMatch($module_file, '\p{Cc}') -or $module_file.IndexOfAny([System.IO.Path]::GetInvalidFileNameChars()) -ge 0 -or -not [string]::Equals($module_file, $module_file.Trim(), [System.StringComparison]::Ordinal) -or [System.IO.Path]::IsPathRooted($module_file) -or $module_file.Contains('\') -or $module_file.Contains('/') -or $module_file.Contains(',')) {
            throw "CommonModules manifest line $line_number has an invalid ModuleFile value '$module_file'."
        }
        $module_extension = [System.IO.Path]::GetExtension($module_file)
        if (-not (Test-OrdinalStringInSet -Value $module_extension -AllowedValues @('.bas', '.cls', '.frm'))) {
            throw "CommonModules manifest line $line_number has an invalid ModuleFile value '$module_file'."
        }
        $module_name = [System.IO.Path]::GetFileNameWithoutExtension($module_file)
        if ([string]::IsNullOrEmpty($module_name) -or -not [string]::Equals($module_name, $module_name.Trim(), [System.StringComparison]::Ordinal) -or $module_name.Contains('.')) {
            throw "CommonModules manifest line $line_number has an invalid ModuleFile value '$module_file'."
        }
        if ($records_by_module.ContainsKey($module_file)) {
            throw "CommonModules manifest line $line_number duplicates ModuleFile '$module_file'."
        }
        $categories = $columns[1]
        if (-not (Test-OrdinalStringInSet -Value $categories -AllowedValues @(
                'runtime-baseline',
                'runtime-baseline,public-udf',
                'test-foundation',
                'optional',
                'optional,public-udf',
                'test-double'
            ))) {
            throw "CommonModules manifest line $line_number has invalid Categories '$categories' for '$module_file'."
        }

        $required_references_text = $columns[3]
        $json_string_pattern = '"(?:[^\x00-\x1F"\\]|\\(?:["\\/bfnrt]|u[0-9A-Fa-f]{4}))*"'
        $json_array_pattern = "^\x20*\[\x20*(?:$json_string_pattern(?:\x20*,\x20*$json_string_pattern)*)?\x20*\]\x20*$"
        if (-not [regex]::IsMatch($required_references_text, $json_array_pattern)) {
            throw "CommonModules manifest line $line_number contains malformed RequiredReferences JSON for '$module_file'."
        }
        $required_references = New-Object 'System.Collections.Generic.List[string]'
        $seen_required_references = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($json_string_match in [regex]::Matches($required_references_text, $json_string_pattern)) {
            if (-not (Test-JsonEscapedSurrogatePairs -Value $json_string_match.Value)) {
                throw "CommonModules manifest line $line_number contains malformed RequiredReferences JSON for '$module_file'."
            }
            try {
                $required_reference = ConvertFrom-Json -InputObject $json_string_match.Value -ErrorAction Stop
            }
            catch {
                throw "CommonModules manifest line $line_number contains malformed RequiredReferences JSON for '$module_file'."
            }
            if ([string]::IsNullOrEmpty($required_reference)) {
                throw "CommonModules manifest line $line_number must contain nonempty RequiredReferences for '$module_file'."
            }
            if (-not [string]::Equals($required_reference, $required_reference.Trim(), [System.StringComparison]::Ordinal)) {
                throw "CommonModules manifest line $line_number must contain already trimmed RequiredReferences for '$module_file'."
            }
            if ([string]::Equals($required_reference, 'Visual Basic For Applications', [System.StringComparison]::OrdinalIgnoreCase)) {
                throw "CommonModules manifest line $line_number must not declare the always-active VBA standard library in RequiredReferences for '$module_file'."
            }
            if (-not $seen_required_references.Add($required_reference)) {
                throw "CommonModules manifest line $line_number contains duplicate RequiredReferences '$required_reference' for '$module_file'."
            }
            $required_references.Add($required_reference)
        }

        $dependencies_text = $columns[2]
        if ([regex]::IsMatch($dependencies_text, '\s')) {
            throw "CommonModules manifest line $line_number must use whitespace-free Dependencies for '$module_file'."
        }
        if ($dependencies_text.Length -eq 0) {
            $dependencies = @()
        }
        else {
            $dependencies = @($dependencies_text.Split(','))
            $seen_dependencies = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
            foreach ($dependency in $dependencies) {
                if ($dependency.Length -eq 0) {
                    throw "CommonModules manifest line $line_number contains an empty dependency for '$module_file'."
                }
                if ([string]::Equals($module_file, $dependency, [System.StringComparison]::OrdinalIgnoreCase)) {
                    throw "CommonModules manifest line $line_number declares a self dependency for '$module_file'."
                }
                if (-not $seen_dependencies.Add($dependency)) {
                    throw "CommonModules manifest line $line_number contains duplicate dependency '$dependency' for '$module_file'."
                }
            }
        }
        $record = [pscustomobject]@{
            ModuleFile = $module_file
            Categories = $categories
            Dependencies = $dependencies
            RequiredReferences = $required_references.ToArray()
            LineNumber = $line_number
        }
        $records.Add($record)
        $records_by_module.Add($module_file, $record)
        $line_index++
    }

    foreach ($record in $records) {
        foreach ($dependency in $record.Dependencies) {
            if (-not $records_by_module.ContainsKey($dependency)) {
                throw "CommonModules manifest line $($record.LineNumber) has unknown dependency '$dependency' for '$($record.ModuleFile)'."
            }
            if (-not [string]::Equals($dependency, $records_by_module[$dependency].ModuleFile, [System.StringComparison]::Ordinal)) {
                throw "CommonModules manifest line $($record.LineNumber) dependency '$dependency' must use the exact ModuleFile spelling '$($records_by_module[$dependency].ModuleFile)'."
            }
            $dependency_record = $records_by_module[$dependency]
            $source_is_runtime = Test-OrdinalStringInSet -Value $record.Categories -AllowedValues @('runtime-baseline', 'runtime-baseline,public-udf', 'optional', 'optional,public-udf')
            $dependency_is_test = Test-OrdinalStringInSet -Value $dependency_record.Categories -AllowedValues @('test-foundation', 'test-double')
            if ($source_is_runtime -and $dependency_is_test) {
                throw "CommonModules manifest line $($record.LineNumber) runtime-role module '$($record.ModuleFile)' cannot depend on test-role module '$dependency'."
            }
        }
    }

    return [pscustomobject]@{
        Path = $resolved_path
        Bytes = $bytes
        Records = $records.ToArray()
        RecordsByModule = $records_by_module
    }
}

function Read-CommonModulesManifestModuleFiles {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ManifestPath
    )

    $manifest = Read-CommonModulesManifest -ManifestPath $ManifestPath
    return @($manifest.Records | ForEach-Object { $_.ModuleFile })
}
