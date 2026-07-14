$ErrorActionPreference = 'Stop'
$InformationPreference = 'Continue'

$validator_path = Join-Path $PSScriptRoot 'validate_common_modules_manifest_main.ps1'
$repo_root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).ProviderPath

function Invoke-ManifestValidator {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ManifestPath,

        [Parameter(Mandatory = $true)]
        [string]$ModulesDirectory,

        [switch]$RequireAllFiles,

        [int]$ExpectedExitCode = 0
    )

    $arguments = @(
        '-NoProfile',
        '-ExecutionPolicy',
        'Bypass',
        '-File',
        $validator_path,
        '-ManifestPath',
        $ManifestPath,
        '-ModulesDirectory',
        $ModulesDirectory
    )
    if ($RequireAllFiles) {
        $arguments += '-RequireAllFiles'
    }

    $previous_error_action_preference = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    $output = & powershell @arguments 2>&1
    $exit_code = $LASTEXITCODE
    $ErrorActionPreference = $previous_error_action_preference
    if ($exit_code -ne $ExpectedExitCode) {
        $output | ForEach-Object { Write-Host $_ }
        throw "Expected exit code $ExpectedExitCode but got $exit_code for manifest '$ManifestPath'."
    }
    if ($ExpectedExitCode -eq 0) {
        $output | ForEach-Object { Write-Host $_ }
    }
}

$source_modules_directory = Join-Path (Join-Path (Join-Path $repo_root 'CommonModules') 'src') 'CommonModules'
$source_manifest_path = Join-Path $source_modules_directory 'common-modules-manifest.tsv'
$distributed_manifest_path = Join-Path (Join-Path $repo_root 'common_modules_repo') 'common-modules-manifest.tsv'
$distributed_modules_directory = Join-Path $repo_root 'common_modules_repo'

Invoke-ManifestValidator -ManifestPath $source_manifest_path -ModulesDirectory $source_modules_directory -RequireAllFiles
Invoke-ManifestValidator -ManifestPath $distributed_manifest_path -ModulesDirectory $distributed_modules_directory -RequireAllFiles

$temp_root = Join-Path ([System.IO.Path]::GetTempPath()) ('common-modules-manifest-test-' + [System.Guid]::NewGuid().ToString('N'))
$temp_modules = Join-Path $temp_root 'modules'
$encoding = [System.Text.Encoding]::GetEncoding(932)

try {
    New-Item -ItemType Directory -Path $temp_modules -Force | Out-Null
    Copy-Item -LiteralPath (Join-Path $source_modules_directory 'Lib_Common.bas') -Destination $temp_modules -Force

    $unknown_dependency_manifest = Join-Path $temp_root 'unknown-dependency.tsv'
    $unknown_dependency_text = "ModuleFile`tCategories`tDependencies`r`nLib_Common.bas`truntime-baseline`tMissingDependency.cls`r`n"
    [System.IO.File]::WriteAllText($unknown_dependency_manifest, $unknown_dependency_text, $encoding)
    Invoke-ManifestValidator -ManifestPath $unknown_dependency_manifest -ModulesDirectory $temp_modules -ExpectedExitCode 1

    $malformed_manifest = Join-Path $temp_root 'malformed.tsv'
    $malformed_text = "ModuleFile`tCategories`r`nLib_Common.bas`truntime-baseline`r`n"
    [System.IO.File]::WriteAllText($malformed_manifest, $malformed_text, $encoding)
    Invoke-ManifestValidator -ManifestPath $malformed_manifest -ModulesDirectory $temp_modules -ExpectedExitCode 1
}
finally {
    if (Test-Path -LiteralPath $temp_root) {
        Remove-Item -LiteralPath $temp_root -Recurse -Force
    }
}

Write-Host 'CommonModules manifest tests passed.'
