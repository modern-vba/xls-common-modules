$ErrorActionPreference = 'Stop'
$InformationPreference = 'Continue'

$current_dir = (Get-Location).ProviderPath
Set-Location $PSScriptRoot

$bat_dir = $args[0]
$bat_file = $args[1]
$bat_path = $args[2]
$arg_dir = $args[3]
$arg_file = $args[4]
$arg_path = $args[5]

Import-Module (Join-Path $PSScriptRoot 'VbaDevTool.psm1') -Force

function Get-VbaProjectSourceDirectories {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SearchRoot
    )

    $excluded_directory_names = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($name in @('.backups', '.git', '.out-of-scope', '.tmp', '.venv', '.vs', '.vscode-test', 'artifacts', 'bin', 'common_modules_repo', 'node_modules', 'obj', 'out', 'packages', 'publish', 'temp', 'TestResults')) {
        [void]$excluded_directory_names.Add($name)
    }

    $source_directories = New-Object 'System.Collections.Generic.Dictionary[string, System.IO.DirectoryInfo]' ([System.StringComparer]::OrdinalIgnoreCase)
    $pending = New-Object 'System.Collections.Generic.Queue[System.IO.DirectoryInfo]'
    $pending.Enqueue((Get-Item -LiteralPath $SearchRoot -ErrorAction Stop))

    while ($pending.Count -gt 0) {
        $directory = $pending.Dequeue()
        $manifest_path = Join-Path $directory.FullName 'vba-project.json'
        if (Test-Path -LiteralPath $manifest_path -PathType Leaf) {
            $manifest = Get-Content -LiteralPath $manifest_path -Raw | ConvertFrom-Json
            if ($null -eq $manifest.documents) {
                throw "vba-project.json does not define documents: $manifest_path"
            }

            foreach ($document in $manifest.documents.PSObject.Properties) {
                if ([string]::IsNullOrWhiteSpace($document.Value.sourcePath)) {
                    throw "vba-project.json document '$($document.Name)' does not define sourcePath: $manifest_path"
                }

                $source_path = [System.IO.Path]::GetFullPath((Join-Path $directory.FullName ($document.Value.sourcePath -replace '/', [System.IO.Path]::DirectorySeparatorChar)))
                if (-not (Test-Path -LiteralPath $source_path -PathType Container)) {
                    throw "vba-project.json sourcePath was not found: $source_path"
                }
                $source_item = Get-Item -LiteralPath $source_path
                $source_directories[$source_item.FullName] = $source_item
            }
        }

        foreach ($child in Get-ChildItem -LiteralPath $directory.FullName -Directory -Force) {
            if ($excluded_directory_names.Contains($child.Name)) {
                continue
            }
            if (($child.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
                continue
            }
            $pending.Enqueue($child)
        }
    }

    return @($source_directories.Values | Sort-Object -Property FullName)
}

function Select-NewestProjectSourceFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ModuleFile,

        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo[]]$SourceFiles,

        [Parameter(Mandatory = $true)]
        [string]$CanonicalSourceDirectory
    )

    $candidate_files = @($SourceFiles | Where-Object {
        [string]::Equals($_.Name, $ModuleFile, [System.StringComparison]::OrdinalIgnoreCase)
    })
    if ($candidate_files.Count -eq 0) {
        throw "Manifest-listed CommonModules source file was not found in any vba-project.json sourcePath: $ModuleFile"
    }

    $newest_ticks = ($candidate_files | ForEach-Object { $_.LastWriteTimeUtc.Ticks } | Measure-Object -Maximum).Maximum
    $newest_files = @($candidate_files | Where-Object { $_.LastWriteTimeUtc.Ticks -eq $newest_ticks })
    if ($newest_files.Count -gt 1) {
        $reference_file = $newest_files[0]
        foreach ($candidate_file in $newest_files | Select-Object -Skip 1) {
            if (-not (Test-FileContentEqual -LeftPath $reference_file.FullName -RightPath $candidate_file.FullName)) {
                $paths = ($newest_files.FullName | Sort-Object) -join '; '
                throw "Newest CommonModules source files have the same timestamp but different content for '$ModuleFile': $paths"
            }
        }
    }

    $canonical_path = Join-Path $CanonicalSourceDirectory $ModuleFile
    $canonical_candidate = $newest_files | Where-Object {
        Test-SamePath -LeftPath $_.FullName -RightPath $canonical_path
    } | Select-Object -First 1
    if ($null -ne $canonical_candidate) {
        return $canonical_candidate
    }

    return ($newest_files | Sort-Object -Property FullName | Select-Object -First 1)
}

try {
    $repo_root = Get-RepositoryRoot
    $source_dir = Join-Path (Join-Path (Join-Path $repo_root 'CommonModules') 'src') 'CommonModules'
    $project_search_root = Split-Path -Parent $repo_root
    $manifest_source_path = Join-Path $source_dir 'common-modules-manifest.tsv'
    if (-not (Test-Path -LiteralPath $manifest_source_path -PathType Leaf)) {
        throw "CommonModules manifest was not found: $manifest_source_path"
    }

    if ([string]::IsNullOrWhiteSpace($arg_path)) {
        $target_root = Get-RequiredDirectoryTarget -Path $current_dir -Description 'target root'
    }
    else {
        $target_root = Get-RequiredDirectoryTarget -Path $arg_path -Description 'target root'
    }

    $module_files = @(Read-CommonModulesManifestModuleFiles -ManifestPath $manifest_source_path)
    $source_directories = @(Get-VbaProjectSourceDirectories -SearchRoot $project_search_root)
    if ($source_directories.Count -eq 0) {
        throw "No vba-project.json sourcePath directories were found under: $project_search_root"
    }
    $source_files = @(
        foreach ($source_directory in $source_directories) {
            Get-ChildItem -LiteralPath $source_directory.FullName -Recurse -File | Where-Object {
                $_.Extension -in @('.bas', '.cls', '.frm')
            }
        }
    )

    $selected_sources = @(
        foreach ($module_file in $module_files) {
            Select-NewestProjectSourceFile -ModuleFile $module_file -SourceFiles $source_files -CanonicalSourceDirectory $source_dir
        }
    )

    $target_output_dir = Join-Path $target_root.FullName 'common_modules_repo'
    if (-not (Test-Path -LiteralPath $target_output_dir -PathType Container)) {
        New-Item -ItemType Directory -Path $target_output_dir -Force | Out-Null
    }

    Copy-Item -LiteralPath $manifest_source_path -Destination (Join-Path $target_output_dir 'common-modules-manifest.tsv') -Force
    Write-Host $manifest_source_path

    for ($index = 0; $index -lt $module_files.Count; $index++) {
        $source_file = $selected_sources[$index]
        Write-Host $source_file.FullName
        Copy-Item -LiteralPath $source_file.FullName -Destination (Join-Path $target_output_dir $module_files[$index]) -Force
    }
}
finally {
    Set-Location $current_dir
}
