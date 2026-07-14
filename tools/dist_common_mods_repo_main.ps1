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

Set-Variable -Name COMMON_MODULES_REPO_DIR_NAME -Value 'common_modules_repo' -Option Constant

function Get-DirectoryInfo {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$Description
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        throw "$Description is empty."
    }

    $resolved_path = (Resolve-Path -LiteralPath $Path -ErrorAction Stop).ProviderPath
    $dir_info = Get-Item -LiteralPath $resolved_path -ErrorAction Stop

    if (-not $dir_info.PSIsContainer) {
        throw "$Description is not a directory: $resolved_path"
    }

    return $dir_info
}

function Test-DirectoryContainsPath {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.DirectoryInfo]$ParentDirectory,

        [Parameter(Mandatory = $true)]
        [string]$ChildPath
    )

    $parent_path = $ParentDirectory.FullName.TrimEnd([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar) + [System.IO.Path]::DirectorySeparatorChar
    $child_full_name = (Get-Item -LiteralPath $ChildPath -ErrorAction Stop).FullName

    return $child_full_name.StartsWith($parent_path, [System.StringComparison]::OrdinalIgnoreCase)
}

function Test-SamePath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$LeftPath,

        [Parameter(Mandatory = $true)]
        [string]$RightPath
    )

    $left_full_name = (Get-Item -LiteralPath $LeftPath -ErrorAction Stop).FullName.TrimEnd([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar)
    $right_full_name = (Get-Item -LiteralPath $RightPath -ErrorAction Stop).FullName.TrimEnd([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar)

    return [string]::Equals($left_full_name, $right_full_name, [System.StringComparison]::OrdinalIgnoreCase)
}

function Get-SourceDirectory {
    if ([string]::IsNullOrWhiteSpace($arg_path)) {
        $source_path = Join-Path $current_dir $COMMON_MODULES_REPO_DIR_NAME
        return Get-DirectoryInfo -Path $source_path -Description 'source common_modules_repo'
    }

    return Get-DirectoryInfo -Path $arg_path -Description 'source directory'
}

function Get-ParentDirectoryInfo {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.DirectoryInfo]$Directory,

        [Parameter(Mandatory = $true)]
        [string]$Description
    )

    $parent_path = Split-Path -Parent $Directory.FullName
    return Get-DirectoryInfo -Path $parent_path -Description $Description
}

function Get-DistributionSearchRoot {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.DirectoryInfo]$SourceDirectory
    )

    $source_parent = Get-ParentDirectoryInfo -Directory $SourceDirectory -Description 'parent directory of the source directory'
    return Get-ParentDirectoryInfo -Directory $source_parent -Description 'distribution target search root'
}

function Assert-CommonModulesRepoPath {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.DirectoryInfo]$ProjectRoot,

        [Parameter(Mandatory = $true)]
        [System.IO.DirectoryInfo]$CommonModulesRepoDirectory
    )

    if ($CommonModulesRepoDirectory.Name -ne $COMMON_MODULES_REPO_DIR_NAME) {
        throw "Unexpected distribution target directory name: $($CommonModulesRepoDirectory.FullName)"
    }

    if (-not (Test-DirectoryContainsPath -ParentDirectory $ProjectRoot -ChildPath $CommonModulesRepoDirectory.FullName)) {
        throw "Distribution target directory is not under the project root: $($CommonModulesRepoDirectory.FullName)"
    }
}

function Get-TargetCommonModulesRepoDirectories {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.DirectoryInfo]$SearchRoot,

        [Parameter(Mandatory = $true)]
        [System.IO.DirectoryInfo]$SourceDirectory
    )

    $target_directories = @()
    $project_roots = @(Get-ChildItem -LiteralPath $SearchRoot.FullName -Directory -Force | Sort-Object -Property FullName)
    foreach ($project_root in $project_roots) {
        $target_common_modules_repo_path = Join-Path $project_root.FullName $COMMON_MODULES_REPO_DIR_NAME
        if (-not (Test-Path -LiteralPath $target_common_modules_repo_path -PathType Container)) {
            continue
        }

        $target_common_modules_repo = Get-DirectoryInfo -Path $target_common_modules_repo_path -Description "target $($project_root.Name) common_modules_repo"
        Assert-CommonModulesRepoPath -ProjectRoot $project_root -CommonModulesRepoDirectory $target_common_modules_repo
        if (Test-SamePath -LeftPath $SourceDirectory.FullName -RightPath $target_common_modules_repo.FullName) {
            continue
        }

        $target_directories += $target_common_modules_repo
    }

    if ($target_directories.Count -eq 0) {
        throw "Target $COMMON_MODULES_REPO_DIR_NAME was not found: $($SearchRoot.FullName)"
    }

    return $target_directories
}

function Assert-SourceDirectoryNotEmpty {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.DirectoryInfo]$SourceDirectory
    )

    $source_items = @(Get-ChildItem -LiteralPath $SourceDirectory.FullName -Force)
    if ($source_items.Count -eq 0) {
        throw "The source directory contains no files: $($SourceDirectory.FullName)"
    }
}

function Clear-DirectoryContents {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.DirectoryInfo]$Directory
    )

    $directory_path = $Directory.FullName
    foreach ($item in Get-ChildItem -LiteralPath $directory_path -Force) {
        if (-not (Test-DirectoryContainsPath -ParentDirectory $Directory -ChildPath $item.FullName)) {
            throw "The item to delete is outside the target directory: $($item.FullName)"
        }

        Remove-Item -LiteralPath $item.FullName -Recurse -Force
    }
}

function Copy-DirectoryContents {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.DirectoryInfo]$SourceDirectory,

        [Parameter(Mandatory = $true)]
        [System.IO.DirectoryInfo]$DestinationDirectory
    )

    $source_items = @(Get-ChildItem -LiteralPath $SourceDirectory.FullName -Force)
    if ($source_items.Count -eq 0) {
        throw "The source directory contains no files: $($SourceDirectory.FullName)"
    }

    foreach ($source_item in $source_items) {
        Copy-Item -LiteralPath $source_item.FullName -Destination $DestinationDirectory.FullName -Recurse -Force
        Write-Information "Copied '$($source_item.Name)' to '$($DestinationDirectory.FullName)'."
    }
}

try {
    $source_directory = Get-SourceDirectory
    Assert-SourceDirectoryNotEmpty -SourceDirectory $source_directory
    $search_root = Get-DistributionSearchRoot -SourceDirectory $source_directory
    $target_common_modules_repos = @(Get-TargetCommonModulesRepoDirectories -SearchRoot $search_root -SourceDirectory $source_directory)

    Write-Host "Source: $($source_directory.FullName)"
    Write-Host "Distribution target search root: $($search_root.FullName)"

    foreach ($target_common_modules_repo in $target_common_modules_repos) {
        Write-Host "Updating distribution target: $($target_common_modules_repo.FullName)"
        Clear-DirectoryContents -Directory $target_common_modules_repo
        Copy-DirectoryContents -SourceDirectory $source_directory -DestinationDirectory $target_common_modules_repo
    }

    Write-Host 'Common modules repo distribution completed.'
}
finally {
    Set-Location $current_dir
}
