$ErrorActionPreference = 'Stop'
$InformationPreference = 'Continue'
$WarningPreference = 'Continue'

$invocation_working_directory = (Get-Location).ProviderPath

Import-Module (Join-Path $PSScriptRoot 'VbaDevTool.psm1') -Force

function Get-ExactOrdinaryFileFromInventory {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$Inventory,

        [Parameter(Mandatory = $true)]
        [string]$ExpectedName,

        [Parameter(Mandatory = $true)]
        [string]$DirectoryPath,

        [Parameter(Mandatory = $true)]
        [string]$Description,

        [switch]$AllowAbsent
    )

    $matches = @($Inventory | Where-Object {
        [string]::Equals($_.Name, $ExpectedName, [System.StringComparison]::OrdinalIgnoreCase)
    })
    if ($matches.Count -eq 0) {
        if ($AllowAbsent) {
            return $null
        }
        throw "$Description was not found in '$DirectoryPath': $ExpectedName"
    }
    if ($matches.Count -ne 1) {
        throw "$Description has multiple case-insensitive '$ExpectedName' matches in '$DirectoryPath'."
    }

    $match = $matches[0]
    if (-not [string]::Equals($match.Name, $ExpectedName, [System.StringComparison]::Ordinal)) {
        throw "$Description must use the ordinal-exact $ExpectedName name in '$DirectoryPath': $($match.Name)"
    }
    if ($match.PSIsContainer -or ($match.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
        throw "$Description must be an ordinary non-reparse file: $($match.FullName)"
    }

    return $match
}

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
    $pending.Enqueue((Get-Item -LiteralPath $SearchRoot -Force -ErrorAction Stop))

    while ($pending.Count -gt 0) {
        $directory = $pending.Dequeue()
        $inventory = @(Get-ChildItem -LiteralPath $directory.FullName -Force -ErrorAction Stop)
        $manifest_file = Get-ExactOrdinaryFileFromInventory -Inventory $inventory -ExpectedName 'vba-project.json' -DirectoryPath $directory.FullName -Description 'Project manifest' -AllowAbsent
        if ($null -ne $manifest_file) {
            $manifest_path = $manifest_file.FullName
            $manifest = Get-Content -LiteralPath $manifest_path -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            $documents_property = $manifest.PSObject.Properties['documents']
            if ($null -eq $documents_property -or $documents_property.Value -isnot [System.Management.Automation.PSCustomObject]) {
                throw "vba-project.json must define a nonempty documents object: $manifest_path"
            }
            $documents = @($documents_property.Value.PSObject.Properties)
            if ($documents.Count -eq 0) {
                throw "vba-project.json must define a nonempty documents object: $manifest_path"
            }

            foreach ($document in $documents) {
                if ($document.Value -isnot [System.Management.Automation.PSCustomObject]) {
                    throw "vba-project.json document '$($document.Name)' must be an object with a nonempty string sourcePath: $manifest_path"
                }
                $source_path_property = $document.Value.PSObject.Properties['sourcePath']
                if ($null -eq $source_path_property -or $source_path_property.Value -isnot [string] -or [string]::IsNullOrWhiteSpace($source_path_property.Value)) {
                    throw "vba-project.json document '$($document.Name)' must define a nonempty string sourcePath: $manifest_path"
                }

                $source_path = Resolve-VbaDevProjectPath -ProjectRoot $directory.FullName -Path $source_path_property.Value
                if (-not (Test-Path -LiteralPath $source_path -PathType Container)) {
                    throw "vba-project.json sourcePath was not found: $source_path"
                }
                $source_item = Get-Item -LiteralPath $source_path -Force
                $source_directories[$source_item.FullName] = $source_item
            }
        }

        foreach ($child in $inventory | Where-Object { $_.PSIsContainer }) {
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

function Get-SourceSetFiles {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.DirectoryInfo]$SourceDirectory
    )

    $files = New-Object 'System.Collections.Generic.List[System.IO.FileInfo]'
    $pending = New-Object 'System.Collections.Generic.Queue[System.IO.DirectoryInfo]'
    $pending.Enqueue($SourceDirectory)

    while ($pending.Count -gt 0) {
        $directory = $pending.Dequeue()
        $inventory = @(Get-ChildItem -LiteralPath $directory.FullName -Force -ErrorAction Stop)
        foreach ($item in $inventory) {
            if ($item.PSIsContainer) {
                if (($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -eq 0) {
                    $pending.Enqueue($item)
                }
                continue
            }
            $files.Add($item)
        }
    }

    return $files.ToArray()
}

function Get-CollectionCandidateFileMetadata {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$File,

        [Parameter(Mandatory = $true)]
        [string]$Description
    )

    $File.Refresh()
    if (-not $File.Exists) {
        throw "$Description disappeared: $($File.FullName)"
    }
    if (($File.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
        throw "$Description must be an ordinary non-reparse file: $($File.FullName)"
    }

    return [pscustomobject]@{
        LastWriteTimeUtc = $File.LastWriteTimeUtc
        Length = $File.Length
    }
}

function New-CollectionModuleCandidate {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ModuleFile,

        [Parameter(Mandatory = $true)]
        [object]$SourceSetScan,

        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$SourceFile,

        [Parameter(Mandatory = $true)]
        [bool]$IsAuthoring
    )

    try {
        $module_metadata = Get-CollectionCandidateFileMetadata -File $SourceFile -Description "Candidate '$ModuleFile'"
        $sidecar_file = $null
        $sidecar_metadata = $null
        if ([string]::Equals([System.IO.Path]::GetExtension($ModuleFile), '.frm', [System.StringComparison]::Ordinal)) {
            $sidecar_name = [System.IO.Path]::GetFileNameWithoutExtension($SourceFile.Name) + '.frx'
            $sidecar_matches = @($SourceSetScan.Files | Where-Object {
                [string]::Equals($_.DirectoryName, $SourceFile.DirectoryName, [System.StringComparison]::OrdinalIgnoreCase) -and
                [string]::Equals($_.Name, $sidecar_name, [System.StringComparison]::OrdinalIgnoreCase)
            })
            if ($sidecar_matches.Count -gt 1) {
                throw "Source set '$($SourceSetScan.SourceDirectory.FullName)' has ambiguous sidecars for '$ModuleFile'."
            }
            if ($sidecar_matches.Count -eq 1) {
                $sidecar_file = $sidecar_matches[0]
                $sidecar_metadata = Get-CollectionCandidateFileMetadata -File $sidecar_file -Description "Form sidecar for '$ModuleFile'"
            }
        }

        $selection_time = $module_metadata.LastWriteTimeUtc
        if ($null -ne $sidecar_metadata -and $sidecar_metadata.LastWriteTimeUtc -gt $selection_time) {
            $selection_time = $sidecar_metadata.LastWriteTimeUtc
        }

        return [pscustomobject]@{
            IsUsable = $true
            FailureMessage = $null
            IsAuthoring = $IsAuthoring
            ModuleFile = $SourceFile
            ModuleLastWriteTimeUtc = $module_metadata.LastWriteTimeUtc
            ModuleLength = $module_metadata.Length
            SidecarFile = $sidecar_file
            SidecarLastWriteTimeUtc = if ($null -eq $sidecar_metadata) { $null } else { $sidecar_metadata.LastWriteTimeUtc }
            SidecarLength = if ($null -eq $sidecar_metadata) { $null } else { $sidecar_metadata.Length }
            SelectionTimeUtc = $selection_time
        }
    }
    catch {
        if ($IsAuthoring) {
            throw "CommonModules Authoring Source Set candidate for '$ModuleFile' is invalid: $($_.Exception.Message)"
        }
        return [pscustomobject]@{
            IsUsable = $false
            FailureMessage = $_.Exception.Message
            IsAuthoring = $false
            ModuleFile = $SourceFile
            ModuleLastWriteTimeUtc = $null
            ModuleLength = $null
            SidecarFile = $null
            SidecarLastWriteTimeUtc = $null
            SidecarLength = $null
            SelectionTimeUtc = $null
        }
    }
}

function Test-EquivalentCollectionCandidateShape {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Left,

        [Parameter(Mandatory = $true)]
        [object]$Right
    )

    if ($Left.ModuleLength -ne $Right.ModuleLength) {
        return $false
    }
    $left_has_sidecar = $null -ne $Left.SidecarFile
    $right_has_sidecar = $null -ne $Right.SidecarFile
    if ($left_has_sidecar -ne $right_has_sidecar) {
        return $false
    }
    if ($left_has_sidecar -and $Left.SidecarLength -ne $Right.SidecarLength) {
        return $false
    }

    return $true
}

function Select-NewestProjectSourceCandidate {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ModuleFile,

        [Parameter(Mandatory = $true)]
        [object[]]$Candidates,

        [Parameter(Mandatory = $true)]
        [object]$AuthoringCandidate
    )

    $unusable_candidates = @($Candidates | Where-Object { -not $_.IsUsable })
    if ($unusable_candidates.Count -gt 0) {
        foreach ($candidate in $unusable_candidates) {
            Write-Warning "Candidate for '$ModuleFile' could not be classified and will use the CommonModules Authoring Source Set fallback: $($candidate.FailureMessage)"
        }
        return $AuthoringCandidate
    }

    [long]$newest_ticks = [long]::MinValue
    foreach ($candidate in $Candidates) {
        [long]$candidate_ticks = $candidate.SelectionTimeUtc.Ticks
        if ($candidate_ticks -gt $newest_ticks) {
            $newest_ticks = $candidate_ticks
        }
    }
    $newest_candidates = @($Candidates | Where-Object { $_.SelectionTimeUtc.Ticks -eq $newest_ticks })
    if ($newest_candidates.Count -gt 1) {
        $reference_candidate = $newest_candidates[0]
        foreach ($candidate in $newest_candidates | Select-Object -Skip 1) {
            if (-not (Test-EquivalentCollectionCandidateShape -Left $reference_candidate -Right $candidate)) {
                Write-Warning "Newest candidates have the same timestamp but different lengths or form sidecar shapes for '$ModuleFile'; using the CommonModules Authoring Source Set fallback."
                return $AuthoringCandidate
            }
        }
    }

    $authoring_newest_candidate = $newest_candidates | Where-Object { $_.IsAuthoring } | Select-Object -First 1
    if ($null -ne $authoring_newest_candidate) {
        return $authoring_newest_candidate
    }

    $ordered_paths = @($newest_candidates | ForEach-Object { $_.ModuleFile.FullName })
    [Array]::Sort($ordered_paths, [System.StringComparer]::Ordinal)
    $selected_path = $ordered_paths[0]
    return ($newest_candidates | Where-Object {
        [string]::Equals($_.ModuleFile.FullName, $selected_path, [System.StringComparison]::Ordinal)
    } | Select-Object -First 1)
}

function New-CollectionPackageEntry {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,

        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$SourceFile,

        [Parameter(Mandatory = $true)]
        [datetime]$LastWriteTimeUtc,

        [Parameter(Mandatory = $true)]
        [long]$Length
    )

    return [pscustomobject]@{
        Name = $Name
        SourceFile = $SourceFile
        LastWriteTimeUtc = $LastWriteTimeUtc
        Length = $Length
    }
}

function Test-CollectionPackageMatch {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$OutputInventory,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$ExpectedEntries
    )

    if ($OutputInventory.Count -ne $ExpectedEntries.Count) {
        return $false
    }

    $expected_by_name = New-Object 'System.Collections.Generic.Dictionary[string, object]' ([System.StringComparer]::Ordinal)
    foreach ($entry in $ExpectedEntries) {
        $expected_by_name.Add($entry.Name, $entry)
    }

    foreach ($item in $OutputInventory) {
        if ($item.PSIsContainer -or ($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
            return $false
        }
        if (-not $expected_by_name.ContainsKey($item.Name)) {
            return $false
        }
        $item.Refresh()
        if (-not $item.Exists) {
            throw "Output inventory entry disappeared during preflight: $($item.FullName)"
        }
        $expected = $expected_by_name[$item.Name]
        if ($item.Length -ne $expected.Length -or $item.LastWriteTimeUtc.Ticks -ne $expected.LastWriteTimeUtc.Ticks) {
            return $false
        }
    }

    return $true
}

try {
    if ($args.Count -ne 1) {
        throw 'COLLECT requires exactly one Collection Search Root argument.'
    }
    $search_root_argument = $args[0]
    if ([string]::IsNullOrWhiteSpace($search_root_argument)) {
        throw 'Collection Search Root is empty.'
    }
    if ([System.IO.Path]::IsPathRooted($search_root_argument)) {
        $search_root_path = [System.IO.Path]::GetFullPath($search_root_argument)
    }
    else {
        $search_root_path = [System.IO.Path]::GetFullPath((Join-Path $invocation_working_directory $search_root_argument))
    }
    $search_root = Get-RequiredDirectoryTarget -Path $search_root_path -Description 'Collection Search Root'

    $source_directories = @(Get-VbaProjectSourceDirectories -SearchRoot $search_root.FullName)
    if ($source_directories.Count -eq 0) {
        throw "No vba-project.json sourcePath directories were found under: $($search_root.FullName)"
    }

    $manifest_owners = New-Object 'System.Collections.Generic.List[object]'
    foreach ($source_directory in $source_directories) {
        $source_inventory = @(Get-ChildItem -LiteralPath $source_directory.FullName -Force -ErrorAction Stop)
        $manifest_file = Get-ExactOrdinaryFileFromInventory -Inventory $source_inventory -ExpectedName 'common-modules-manifest.tsv' -DirectoryPath $source_directory.FullName -Description 'CommonModules manifest' -AllowAbsent
        if ($null -ne $manifest_file) {
            $manifest_owners.Add([pscustomobject]@{
                SourceDirectory = $source_directory
                ManifestFile = $manifest_file
            })
        }
    }
    if ($manifest_owners.Count -ne 1) {
        throw "Exactly one CommonModules Authoring Source Set is required, but found $($manifest_owners.Count)."
    }
    $source_dir = $manifest_owners[0].SourceDirectory.FullName
    $manifest_source_path = $manifest_owners[0].ManifestFile.FullName
    $manifest_metadata = Get-CollectionCandidateFileMetadata -File $manifest_owners[0].ManifestFile -Description 'CommonModules manifest'
    $module_files = @(Read-CommonModulesManifestModuleFiles -ManifestPath $manifest_source_path)

    $source_set_scans = @(
        foreach ($source_directory in $source_directories) {
            [pscustomobject]@{
                SourceDirectory = $source_directory
                Files = @(Get-SourceSetFiles -SourceDirectory $source_directory)
            }
        }
    )

    $selected_sources = @(
        foreach ($module_file in $module_files) {
            $module_candidates = New-Object 'System.Collections.Generic.List[object]'
            $authoring_candidate = $null
            foreach ($source_set_scan in $source_set_scans) {
                $source_set_matches = @($source_set_scan.Files | Where-Object {
                    [string]::Equals($_.Name, $module_file, [System.StringComparison]::OrdinalIgnoreCase)
                })
                if ($source_set_matches.Count -gt 1) {
                    $paths = @($source_set_matches | ForEach-Object { $_.FullName })
                    [Array]::Sort($paths, [System.StringComparer]::Ordinal)
                    throw "Source set '$($source_set_scan.SourceDirectory.FullName)' has an ambiguous candidate for '$module_file': $($paths -join '; ')"
                }
                if ($source_set_matches.Count -eq 1) {
                    $is_authoring = Test-SamePath -LeftPath $source_set_scan.SourceDirectory.FullName -RightPath $source_dir
                    $candidate = New-CollectionModuleCandidate -ModuleFile $module_file -SourceSetScan $source_set_scan -SourceFile $source_set_matches[0] -IsAuthoring $is_authoring
                    $module_candidates.Add($candidate)
                    if ($is_authoring) {
                        $authoring_candidate = $candidate
                    }
                }
            }
            if ($null -eq $authoring_candidate) {
                throw "CommonModules Authoring Source Set must provide '$module_file': $source_dir"
            }
            Select-NewestProjectSourceCandidate -ModuleFile $module_file -Candidates $module_candidates.ToArray() -AuthoringCandidate $authoring_candidate
        }
    )

    $package_entries = New-Object 'System.Collections.Generic.List[object]'
    $package_entries.Add((New-CollectionPackageEntry -Name 'common-modules-manifest.tsv' -SourceFile $manifest_owners[0].ManifestFile -LastWriteTimeUtc $manifest_metadata.LastWriteTimeUtc -Length $manifest_metadata.Length))
    for ($index = 0; $index -lt $module_files.Count; $index++) {
        $selected_candidate = $selected_sources[$index]
        $package_entries.Add((New-CollectionPackageEntry -Name $module_files[$index] -SourceFile $selected_candidate.ModuleFile -LastWriteTimeUtc $selected_candidate.ModuleLastWriteTimeUtc -Length $selected_candidate.ModuleLength))
        if ($null -ne $selected_candidate.SidecarFile) {
            $sidecar_output_name = [System.IO.Path]::ChangeExtension($module_files[$index], '.frx')
            $package_entries.Add((New-CollectionPackageEntry -Name $sidecar_output_name -SourceFile $selected_candidate.SidecarFile -LastWriteTimeUtc $selected_candidate.SidecarLastWriteTimeUtc -Length $selected_candidate.SidecarLength))
        }
    }

    $working_inventory = @(Get-ChildItem -LiteralPath $invocation_working_directory -Force -ErrorAction Stop)
    $output_matches = @($working_inventory | Where-Object {
        [string]::Equals($_.Name, 'common_modules_repo', [System.StringComparison]::OrdinalIgnoreCase)
    })
    if ($output_matches.Count -gt 1) {
        throw "Invocation working directory has multiple case-insensitive common_modules_repo entries: $invocation_working_directory"
    }

    $output_directory = $null
    $output_inventory = @()
    if ($output_matches.Count -eq 1) {
        $output_directory = $output_matches[0]
        if (-not [string]::Equals($output_directory.Name, 'common_modules_repo', [System.StringComparison]::Ordinal) -or -not $output_directory.PSIsContainer -or ($output_directory.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
            throw "Output must be an ordinal-exact ordinary non-reparse common_modules_repo directory: $($output_directory.FullName)"
        }
        $output_inventory = @(Get-ChildItem -LiteralPath $output_directory.FullName -Force -ErrorAction Stop)
        if (Test-CollectionPackageMatch -OutputInventory $output_inventory -ExpectedEntries $package_entries.ToArray()) {
            Write-Host "UNCHANGED: $($output_directory.FullName)"
            return
        }
    }

    $target_output_dir = Join-Path $invocation_working_directory 'common_modules_repo'
    if ($null -eq $output_directory) {
        $output_directory = New-Item -ItemType Directory -Path $target_output_dir -ErrorAction Stop
    }
    else {
        foreach ($item in $output_inventory) {
            Remove-Item -LiteralPath $item.FullName -Recurse -Force -ErrorAction Stop
        }
    }

    foreach ($entry in $package_entries) {
        Write-Host $entry.SourceFile.FullName
        Copy-Item -LiteralPath $entry.SourceFile.FullName -Destination (Join-Path $target_output_dir $entry.Name) -Force -ErrorAction Stop
    }
}
catch {
    Write-Error -ErrorRecord $_
    exit 1
}
