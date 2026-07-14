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
    $item = Get-Item -LiteralPath $resolved_path -ErrorAction Stop
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

    $manifest_path = Join-Path $ProjectRoot 'project.json'
    if (-not (Test-Path -LiteralPath $manifest_path -PathType Leaf)) {
        throw "project.json was not found: $manifest_path"
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
        throw "Document '$DocumentName' is not defined in project.json: $resolved_project_root"
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
        if (Test-Path -LiteralPath (Join-Path $directory.FullName 'project.json') -PathType Leaf) {
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
        $manifest_path = Join-Path $current.FullName 'project.json'
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

function Read-CommonModulesManifestModuleFiles {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ManifestPath
    )

    $encoding = [System.Text.Encoding]::GetEncoding(932)
    $lines = [System.IO.File]::ReadAllLines((Resolve-Path -LiteralPath $ManifestPath).ProviderPath, $encoding)
    $header_found = $false
    $module_files = New-Object 'System.Collections.Generic.List[string]'

    foreach ($line in $lines) {
        if ([string]::IsNullOrWhiteSpace($line) -or $line.TrimStart().StartsWith('#')) {
            continue
        }

        if (-not $header_found) {
            if ($line -ne "ModuleFile`tCategories`tDependencies") {
                throw "CommonModules manifest has an invalid header: $ManifestPath"
            }
            $header_found = $true
            continue
        }

        $columns = $line.Split("`t")
        if ($columns.Count -ne 3) {
            throw "CommonModules manifest row must contain exactly 3 columns: $line"
        }
        if ([string]::IsNullOrWhiteSpace($columns[0])) {
            throw "CommonModules manifest row has an empty ModuleFile value: $line"
        }
        $module_files.Add($columns[0].Trim())
    }

    if (-not $header_found) {
        throw "CommonModules manifest header was not found: $ManifestPath"
    }
    if ($module_files.Count -eq 0) {
        throw "CommonModules manifest contains no module rows: $ManifestPath"
    }

    return $module_files
}
