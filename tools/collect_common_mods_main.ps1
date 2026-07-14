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

try {
    $repo_root = Get-RepositoryRoot
    $source_dir = Join-Path (Join-Path (Join-Path $repo_root 'CommonModules') 'src') 'CommonModules'
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

    $target_output_dir = Join-Path $target_root.FullName 'common_modules_repo'
    if (-not (Test-Path -LiteralPath $target_output_dir -PathType Container)) {
        New-Item -ItemType Directory -Path $target_output_dir -Force | Out-Null
    }

    $module_files = @(Read-CommonModulesManifestModuleFiles -ManifestPath $manifest_source_path)
    Copy-Item -LiteralPath $manifest_source_path -Destination (Join-Path $target_output_dir 'common-modules-manifest.tsv') -Force
    Write-Host $manifest_source_path

    foreach ($module_file in $module_files) {
        $source_path = Join-Path $source_dir $module_file
        if (-not (Test-Path -LiteralPath $source_path -PathType Leaf)) {
            throw "Manifest-listed CommonModules source file was not found: $source_path"
        }

        Write-Host $source_path
        Copy-Item -LiteralPath $source_path -Destination (Join-Path $target_output_dir $module_file) -Force
    }
}
finally {
    Set-Location $current_dir
}
