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
    if ([string]::IsNullOrWhiteSpace($arg_path)) {
        $target_root = Get-RequiredDirectoryTarget -Path $current_dir -Description 'target root'
    }
    else {
        $target_root = Get-RequiredDirectoryTarget -Path $arg_path -Description 'target root'
    }

    $project_roots = @(Get-VbaDevProjectRoots -SearchRoot $target_root.FullName | Sort-Object)
    if ($project_roots.Count -eq 0) {
        throw "No vba-dev projects were found under: $($target_root.FullName)"
    }

    foreach ($project_root in $project_roots) {
        Write-Host "Updating CommonModules for project: $project_root"
        Invoke-VbaDev -Arguments @('common-module', 'update', '--project', $project_root)

        $document_contexts = @(Get-VbaDevProjectDocumentContexts -ProjectRoot $project_root | Sort-Object -Property DocumentName)
        foreach ($document_context in $document_contexts) {
            Write-Host "Building document: $($document_context.DocumentName)"
            Invoke-VbaDev -Arguments @('build', '--project', $project_root, '--document', $document_context.DocumentName)
        }
    }
}
finally {
    Set-Location $current_dir
}
