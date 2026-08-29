$ErrorActionPreference = 'Stop'
$InformationPreference = 'Continue'
$WarningPreference = 'Continue'

$invocation_working_directory = (Get-Location).ProviderPath

Import-Module (Join-Path $PSScriptRoot 'VbaDevTool.psm1') -Force

Set-Variable -Name COMMON_MODULES_REPO_DIR_NAME -Value 'common_modules_repo' -Option Constant
Set-Variable -Name COMMON_MODULES_MANIFEST_FILE_NAME -Value 'common-modules-manifest.tsv' -Option Constant

function Test-ReparsePoint {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileSystemInfo]$Item
    )

    return (($Item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0)
}

function Get-NormalizedLexicalPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $full_path = [System.IO.Path]::GetFullPath($Path)
    $root_path = [System.IO.Path]::GetPathRoot($full_path)
    while ($full_path.Length -gt $root_path.Length -and
        ($full_path.EndsWith([System.IO.Path]::DirectorySeparatorChar) -or
            $full_path.EndsWith([System.IO.Path]::AltDirectorySeparatorChar))) {
        $full_path = $full_path.Substring(0, $full_path.Length - 1)
    }

    return $full_path
}

function Get-OrdinalSortedItems {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$Items,

        [Parameter(Mandatory = $true)]
        [string]$PropertyName
    )

    $items_by_value = New-Object 'System.Collections.Generic.Dictionary[string, object]' ([System.StringComparer]::Ordinal)
    foreach ($item in $Items) {
        $value = [string]$item.$PropertyName
        if ($items_by_value.ContainsKey($value)) {
            throw "Cannot order duplicate '$PropertyName' value: $value"
        }
        $items_by_value.Add($value, $item)
    }

    [string[]]$ordered_values = @($items_by_value.Keys)
    [Array]::Sort($ordered_values, [System.StringComparer]::Ordinal)
    foreach ($value in $ordered_values) {
        $items_by_value[$value]
    }
}

function Resolve-DistributionSearchRoot {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$Arguments,

        [Parameter(Mandatory = $true)]
        [string]$InvocationWorkingDirectory
    )

    if ($Arguments.Count -ne 1) {
        throw 'DIST requires exactly one Distribution Search Root argument.'
    }

    $argument_path = [string]$Arguments[0]
    if ([string]::IsNullOrWhiteSpace($argument_path)) {
        throw 'Distribution Search Root is empty.'
    }

    if ([System.IO.Path]::IsPathRooted($argument_path)) {
        $candidate_path = $argument_path
    }
    else {
        $candidate_path = Join-Path $InvocationWorkingDirectory $argument_path
    }

    $resolved_path = (Resolve-Path -LiteralPath $candidate_path -ErrorAction Stop).ProviderPath
    $search_root = Get-Item -LiteralPath $resolved_path -Force -ErrorAction Stop
    if (-not $search_root.PSIsContainer) {
        throw "Distribution Search Root must be a directory: $resolved_path"
    }

    return $search_root
}

function Get-WrapperSourceRepository {
    param(
        [Parameter(Mandatory = $true)]
        [string]$InvocationWorkingDirectory
    )

    $parent_inventory = @(Get-ChildItem -LiteralPath $InvocationWorkingDirectory -Force -ErrorAction Stop)
    $matches = @($parent_inventory | Where-Object {
        [string]::Equals($_.Name, $COMMON_MODULES_REPO_DIR_NAME, [System.StringComparison]::OrdinalIgnoreCase)
    })
    if ($matches.Count -eq 0) {
        throw "Source $COMMON_MODULES_REPO_DIR_NAME was not found directly under the invocation working directory: $InvocationWorkingDirectory"
    }
    if ($matches.Count -ne 1) {
        throw "Source has multiple case-insensitive '$COMMON_MODULES_REPO_DIR_NAME' matches directly under the invocation working directory: $InvocationWorkingDirectory"
    }

    $source_repository = $matches[0]
    if (-not [string]::Equals($source_repository.Name, $COMMON_MODULES_REPO_DIR_NAME, [System.StringComparison]::Ordinal)) {
        throw "Source repository must use the ordinal-exact '$COMMON_MODULES_REPO_DIR_NAME' name: $($source_repository.FullName)"
    }
    if (-not $source_repository.PSIsContainer -or $source_repository -isnot [System.IO.DirectoryInfo] -or (Test-ReparsePoint -Item $source_repository)) {
        throw "Source repository must be an ordinary non-reparse directory: $($source_repository.FullName)"
    }

    return $source_repository
}

function Get-ExactPackageFile {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$Inventory,

        [Parameter(Mandatory = $true)]
        [string]$ExpectedName,

        [Parameter(Mandatory = $true)]
        [string]$SourceRepositoryPath,

        [switch]$AllowAbsent
    )

    $matches = @($Inventory | Where-Object {
        [string]::Equals($_.Name, $ExpectedName, [System.StringComparison]::OrdinalIgnoreCase)
    })
    if ($matches.Count -eq 0) {
        if ($AllowAbsent) {
            return $null
        }
        throw "Source package entry was not found in '$SourceRepositoryPath': $ExpectedName"
    }
    if ($matches.Count -ne 1) {
        throw "Source package has multiple case-insensitive '$ExpectedName' matches in '$SourceRepositoryPath'."
    }

    $match = $matches[0]
    if (-not [string]::Equals($match.Name, $ExpectedName, [System.StringComparison]::Ordinal)) {
        throw "Source package entry must use the ordinal-exact '$ExpectedName' name: $($match.FullName)"
    }
    if ($match.PSIsContainer -or $match -isnot [System.IO.FileInfo] -or (Test-ReparsePoint -Item $match)) {
        throw "Source package entry must be an ordinary non-reparse file: $($match.FullName)"
    }

    return $match
}

function Assert-FileReadable {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$File,

        [Parameter(Mandatory = $true)]
        [string]$Description
    )

    $stream = $null
    try {
        $file_share = [System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete
        $stream = [System.IO.File]::Open($File.FullName, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, $file_share)
    }
    catch {
        throw "$Description is unreadable: $($File.FullName). $($_.Exception.Message)"
    }
    finally {
        if ($null -ne $stream) {
            $stream.Dispose()
        }
    }
}

function New-SourcePackageEntry {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$File
    )

    $File.Refresh()
    if (-not $File.Exists -or (Test-ReparsePoint -Item $File)) {
        throw "Source package entry must remain an ordinary non-reparse file: $($File.FullName)"
    }
    Assert-FileReadable -File $File -Description 'Source package entry'

    return [pscustomobject]@{
        Name = $File.Name
        FullName = $File.FullName
        LastWriteTimeUtc = $File.LastWriteTimeUtc
        Length = $File.Length
    }
}

function Read-DistributionSourcePackage {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.DirectoryInfo]$SourceRepository
    )

    $inventory = @(Get-ChildItem -LiteralPath $SourceRepository.FullName -Force -ErrorAction Stop)
    foreach ($item in $inventory) {
        if ($item.PSIsContainer -or $item -isnot [System.IO.FileInfo] -or (Test-ReparsePoint -Item $item)) {
            throw "Source package must be flat and contain only ordinary non-reparse files: $($item.FullName)"
        }
    }

    $manifest_file = Get-ExactPackageFile -Inventory $inventory -ExpectedName $COMMON_MODULES_MANIFEST_FILE_NAME -SourceRepositoryPath $SourceRepository.FullName
    $manifest = Read-CommonModulesManifest -ManifestPath $manifest_file.FullName

    $expected_names = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::Ordinal)
    $files_by_name = New-Object 'System.Collections.Generic.Dictionary[string, System.IO.FileInfo]' ([System.StringComparer]::Ordinal)
    [void]$expected_names.Add($COMMON_MODULES_MANIFEST_FILE_NAME)
    $files_by_name.Add($COMMON_MODULES_MANIFEST_FILE_NAME, $manifest_file)

    foreach ($record in $manifest.Records) {
        $module_file = Get-ExactPackageFile -Inventory $inventory -ExpectedName $record.ModuleFile -SourceRepositoryPath $SourceRepository.FullName
        [void]$expected_names.Add($record.ModuleFile)
        $files_by_name.Add($record.ModuleFile, $module_file)

        if ([string]::Equals([System.IO.Path]::GetExtension($record.ModuleFile), '.frm', [System.StringComparison]::Ordinal)) {
            $sidecar_name = [System.IO.Path]::GetFileNameWithoutExtension($record.ModuleFile) + '.frx'
            $sidecar_file = Get-ExactPackageFile -Inventory $inventory -ExpectedName $sidecar_name -SourceRepositoryPath $SourceRepository.FullName -AllowAbsent
            if ($null -ne $sidecar_file) {
                [void]$expected_names.Add($sidecar_name)
                $files_by_name.Add($sidecar_name, $sidecar_file)
            }
        }
    }

    foreach ($item in $inventory) {
        if (-not $expected_names.Contains($item.Name)) {
            throw "Source package contains an unexpected entry: $($item.FullName)"
        }
    }

    [string[]]$ordered_names = @($files_by_name.Keys)
    [Array]::Sort($ordered_names, [System.StringComparer]::Ordinal)
    $entries = New-Object 'System.Collections.Generic.List[object]'
    foreach ($name in $ordered_names) {
        $entries.Add((New-SourcePackageEntry -File $files_by_name[$name]))
    }

    return [pscustomobject]@{
        SourceRepository = $SourceRepository
        Entries = $entries.ToArray()
    }
}

function Assert-DistributionSourceSnapshot {
    param(
        [Parameter(Mandatory = $true)]
        [object]$SourcePackage
    )

    $source_repository = Get-Item -LiteralPath $SourcePackage.SourceRepository.FullName -Force -ErrorAction Stop
    if (-not $source_repository.PSIsContainer -or $source_repository -isnot [System.IO.DirectoryInfo] -or
        -not [string]::Equals($source_repository.Name, $COMMON_MODULES_REPO_DIR_NAME, [System.StringComparison]::Ordinal) -or
        (Test-ReparsePoint -Item $source_repository)) {
        throw "Source repository is no longer the validated ordinary '$COMMON_MODULES_REPO_DIR_NAME' directory: $($SourcePackage.SourceRepository.FullName)"
    }

    $inventory = @(Get-ChildItem -LiteralPath $source_repository.FullName -Force -ErrorAction Stop)
    if ($inventory.Count -ne $SourcePackage.Entries.Count) {
        throw "Source package inventory changed after validation: $($source_repository.FullName)"
    }

    foreach ($item in $inventory) {
        if ($item.PSIsContainer -or $item -isnot [System.IO.FileInfo] -or (Test-ReparsePoint -Item $item)) {
            throw "Source package became invalid after validation: $($item.FullName)"
        }
    }

    foreach ($entry in $SourcePackage.Entries) {
        $matches = @($inventory | Where-Object {
            [string]::Equals($_.Name, $entry.Name, [System.StringComparison]::Ordinal)
        })
        if ($matches.Count -ne 1) {
            throw "Source package inventory changed after validation: $($entry.FullName)"
        }

        $file = $matches[0]
        $file.Refresh()
        if (-not $file.Exists -or (Test-ReparsePoint -Item $file) -or
            $file.LastWriteTimeUtc.Ticks -ne $entry.LastWriteTimeUtc.Ticks -or
            $file.Length -ne $entry.Length) {
            throw "Source package metadata changed after validation: $($entry.FullName)"
        }
        Assert-FileReadable -File $file -Description 'Source package entry'
    }
}

function New-DistributionCandidateFailure {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SortPath,

        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    return [pscustomobject]@{
        Kind = 'Failure'
        SortPath = $SortPath
        TargetPath = $SortPath
        TargetDirectory = $null
        Message = $Message
    }
}

function New-DistributionTargetCandidate {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.DirectoryInfo]$TargetDirectory
    )

    return [pscustomobject]@{
        Kind = 'Target'
        SortPath = (Get-NormalizedLexicalPath -Path $TargetDirectory.FullName)
        TargetPath = $TargetDirectory.FullName
        TargetDirectory = $TargetDirectory
        Message = $null
    }
}

function Find-DistributionCandidates {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.DirectoryInfo]$SearchRoot,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$SearchRootInventory,

        [Parameter(Mandatory = $true)]
        [System.IO.DirectoryInfo]$SourceRepository
    )

    $candidates = New-Object 'System.Collections.Generic.List[object]'
    $target_count = 0
    $failure_count = 0
    $source_path = Get-NormalizedLexicalPath -Path $SourceRepository.FullName

    $project_directories = @($SearchRootInventory | Where-Object { $_.PSIsContainer })
    $ordered_projects = @(Get-OrdinalSortedItems -Items $project_directories -PropertyName 'FullName')
    foreach ($project_directory in $ordered_projects) {
        if (Test-ReparsePoint -Item $project_directory) {
            Write-Warning "Distribution project child is a reparse point and was skipped: $($project_directory.FullName)"
            continue
        }

        try {
            $project_inventory = @(Get-ChildItem -LiteralPath $project_directory.FullName -Force -ErrorAction Stop)
            $matches = @($project_inventory | Where-Object {
                [string]::Equals($_.Name, $COMMON_MODULES_REPO_DIR_NAME, [System.StringComparison]::OrdinalIgnoreCase)
            })
        }
        catch {
            Write-Warning "Could not determine whether project child opted in; it was skipped untouched: $($project_directory.FullName). $($_.Exception.Message)"
            continue
        }

        if ($matches.Count -eq 0) {
            continue
        }

        $expected_target_path = Get-NormalizedLexicalPath -Path (Join-Path $project_directory.FullName $COMMON_MODULES_REPO_DIR_NAME)
        if ($matches.Count -ne 1) {
            $candidates.Add((New-DistributionCandidateFailure -SortPath $expected_target_path -Message "Multiple case-insensitive '$COMMON_MODULES_REPO_DIR_NAME' entries were found in '$($project_directory.FullName)'."))
            $failure_count++
            continue
        }

        $match = $matches[0]
        if (-not [string]::Equals($match.Name, $COMMON_MODULES_REPO_DIR_NAME, [System.StringComparison]::Ordinal)) {
            $candidates.Add((New-DistributionCandidateFailure -SortPath $expected_target_path -Message "Opt-in repository must use the ordinal-exact '$COMMON_MODULES_REPO_DIR_NAME' name: $($match.FullName)"))
            $failure_count++
            continue
        }
        if (-not $match.PSIsContainer -or $match -isnot [System.IO.DirectoryInfo] -or (Test-ReparsePoint -Item $match)) {
            $candidates.Add((New-DistributionCandidateFailure -SortPath $expected_target_path -Message "Opt-in repository must be an ordinary non-reparse directory: $($match.FullName)"))
            $failure_count++
            continue
        }

        $target_path = Get-NormalizedLexicalPath -Path $match.FullName
        if ([string]::Equals($target_path, $source_path, [System.StringComparison]::OrdinalIgnoreCase)) {
            continue
        }

        $candidates.Add((New-DistributionTargetCandidate -TargetDirectory $match))
        $target_count++
    }

    $ordered_candidates = @(Get-OrdinalSortedItems -Items $candidates.ToArray() -PropertyName 'SortPath')
    return [pscustomobject]@{
        Candidates = $ordered_candidates
        TargetCount = $target_count
        FailureCount = $failure_count
    }
}

function Get-CurrentDistributionTarget {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TargetPath
    )

    try {
        $target = Get-Item -LiteralPath $TargetPath -Force -ErrorAction Stop
    }
    catch {
        if ($_.CategoryInfo.Category -eq [System.Management.Automation.ErrorCategory]::ObjectNotFound) {
            return $null
        }
        throw
    }

    if (-not $target.PSIsContainer -or $target -isnot [System.IO.DirectoryInfo] -or
        -not [string]::Equals($target.Name, $COMMON_MODULES_REPO_DIR_NAME, [System.StringComparison]::Ordinal) -or
        (Test-ReparsePoint -Item $target)) {
        throw "Admitted target is no longer an ordinary non-reparse '$COMMON_MODULES_REPO_DIR_NAME' directory: $TargetPath"
    }

    return $target
}

function Get-DistributionTargetPreflight {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.DirectoryInfo]$TargetDirectory
    )

    $pending = New-Object 'System.Collections.Generic.Queue[System.IO.DirectoryInfo]'
    $pending.Enqueue($TargetDirectory)
    $root_inventory = @()
    $is_first = $true

    while ($pending.Count -gt 0) {
        $directory = $pending.Dequeue()
        $inventory = @(Get-ChildItem -LiteralPath $directory.FullName -Force -ErrorAction Stop)
        if ($is_first) {
            $root_inventory = $inventory
            $is_first = $false
        }

        foreach ($item in $inventory) {
            if (Test-ReparsePoint -Item $item) {
                throw "Target contains a reparse entry and cannot be changed safely: $($item.FullName)"
            }
            if ($item.PSIsContainer) {
                if ($item -isnot [System.IO.DirectoryInfo]) {
                    throw "Target contains a nonordinary directory entry: $($item.FullName)"
                }
                $pending.Enqueue($item)
            }
            elseif ($item -isnot [System.IO.FileInfo]) {
                throw "Target contains a nonordinary file entry: $($item.FullName)"
            }
        }
    }

    return [pscustomobject]@{
        RootInventory = $root_inventory
    }
}

function Test-DistributionPackageMatch {
    param(
        [Parameter(Mandatory = $true)]
        [object]$SourcePackage,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$TargetInventory
    )

    if ($TargetInventory.Count -ne $SourcePackage.Entries.Count) {
        return $false
    }

    foreach ($source_entry in $SourcePackage.Entries) {
        $matches = @($TargetInventory | Where-Object {
            [string]::Equals($_.Name, $source_entry.Name, [System.StringComparison]::Ordinal)
        })
        if ($matches.Count -ne 1 -or $matches[0].PSIsContainer -or $matches[0] -isnot [System.IO.FileInfo]) {
            return $false
        }

        $target_file = $matches[0]
        $target_file.Refresh()
        if (-not $target_file.Exists -or
            $target_file.LastWriteTimeUtc.Ticks -ne $source_entry.LastWriteTimeUtc.Ticks -or
            $target_file.Length -ne $source_entry.Length) {
            return $false
        }
    }

    return $true
}

function Clear-DistributionTarget {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$TargetInventory
    )

    $ordered_items = @(Get-OrdinalSortedItems -Items $TargetInventory -PropertyName 'FullName')
    foreach ($item in $ordered_items) {
        Remove-Item -LiteralPath $item.FullName -Recurse -Force -ErrorAction Stop
    }
}

function Copy-DistributionPackage {
    param(
        [Parameter(Mandatory = $true)]
        [object]$SourcePackage,

        [Parameter(Mandatory = $true)]
        [System.IO.DirectoryInfo]$TargetDirectory
    )

    foreach ($source_entry in $SourcePackage.Entries) {
        $destination_path = Join-Path $TargetDirectory.FullName $source_entry.Name
        Copy-Item -LiteralPath $source_entry.FullName -Destination $destination_path -Force -ErrorAction Stop
        Write-Information "Copied '$($source_entry.Name)' to '$($TargetDirectory.FullName)'."
    }
}

try {
    $search_root = Resolve-DistributionSearchRoot -Arguments $args -InvocationWorkingDirectory $invocation_working_directory
    $search_root_inventory = @(Get-ChildItem -LiteralPath $search_root.FullName -Force -ErrorAction Stop)
    $source_repository = Get-WrapperSourceRepository -InvocationWorkingDirectory $invocation_working_directory
    $source_package = Read-DistributionSourcePackage -SourceRepository $source_repository
    $discovery = Find-DistributionCandidates -SearchRoot $search_root -SearchRootInventory $search_root_inventory -SourceRepository $source_repository
}
catch {
    throw "Distribution global failure: $($_.Exception.Message)"
}

Write-Host "Source: $($source_repository.FullName)"
Write-Host "Distribution Search Root: $($search_root.FullName)"

if ($discovery.TargetCount -eq 0 -and $discovery.FailureCount -eq 0) {
    throw "No eligible distribution target was found: $($search_root.FullName)"
}

$has_candidate_failure = $false
foreach ($candidate in $discovery.Candidates) {
    if ([string]::Equals($candidate.Kind, 'Failure', [System.StringComparison]::Ordinal)) {
        $has_candidate_failure = $true
        Write-Error -Message "Distribution candidate failure for '$($candidate.TargetPath)': $($candidate.Message)" -ErrorAction Continue
        continue
    }

    try {
        $target_directory = Get-CurrentDistributionTarget -TargetPath $candidate.TargetPath
    }
    catch {
        $has_candidate_failure = $true
        Write-Error -Message "Distribution candidate failure for '$($candidate.TargetPath)': $($_.Exception.Message)" -ErrorAction Continue
        continue
    }
    if ($null -eq $target_directory) {
        Write-Warning "Admitted distribution target disappeared before its turn and was skipped: $($candidate.TargetPath)"
        continue
    }

    try {
        Assert-DistributionSourceSnapshot -SourcePackage $source_package
    }
    catch {
        throw "Distribution global failure: the source package changed or became unreadable after validation. $($_.Exception.Message)"
    }

    try {
        $target_preflight = Get-DistributionTargetPreflight -TargetDirectory $target_directory
        $is_unchanged = Test-DistributionPackageMatch -SourcePackage $source_package -TargetInventory $target_preflight.RootInventory
    }
    catch {
        $has_candidate_failure = $true
        Write-Error -Message "Distribution candidate failure for '$($candidate.TargetPath)': $($_.Exception.Message)" -ErrorAction Continue
        continue
    }

    if ($is_unchanged) {
        Write-Host "UNCHANGED: $($target_directory.FullName)"
        continue
    }

    Write-Host "Updating distribution target: $($target_directory.FullName)"
    try {
        Clear-DistributionTarget -TargetInventory $target_preflight.RootInventory
    }
    catch {
        $has_candidate_failure = $true
        Write-Error -Message "Distribution candidate failure while clearing '$($candidate.TargetPath)': $($_.Exception.Message)" -ErrorAction Continue
        continue
    }

    try {
        Copy-DistributionPackage -SourcePackage $source_package -TargetDirectory $target_directory
    }
    catch {
        $copy_error = $_
        try {
            Assert-DistributionSourceSnapshot -SourcePackage $source_package
        }
        catch {
            throw "Distribution global failure: the source package changed or became unreadable during copy. $($_.Exception.Message)"
        }

        $has_candidate_failure = $true
        Write-Error -Message "Distribution candidate failure while copying to '$($candidate.TargetPath)': $($copy_error.Exception.Message)" -ErrorAction Continue
        continue
    }
}

if ($has_candidate_failure) {
    throw 'One or more distribution candidates failed.'
}

Write-Host 'CommonModules package distribution completed.'
