[CmdletBinding(PositionalBinding = $false)]
param (
    [ValidateNotNullOrEmpty()]
    [string]$BatPath,

    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$ScriptArguments = @()
)
$ErrorActionPreference = 'Stop'
$InformationPreference = 'Continue'

$current_dir = (Get-Location).ProviderPath
Import-Module (Join-Path $PSScriptRoot 'CommonTools')
#Import-Module (Join-Path $PSScriptRoot 'ExcelModuleManager')

if ($PSBoundParameters.ContainsKey('BatPath')) {
    $bat_file_info = Resolve-FileSystemPath -Path $BatPath -BasePath $current_dir
    $bat_dir = $bat_file_info.Directory
    $bat_file = $bat_file_info.File
    $bat_path = $bat_file_info.Path
}
else {
    $bat_dir = $current_dir
    $bat_file = ''
    $bat_path = $current_dir
}

# /nowait belongs to the BAT launcher, not the target paths.
$arg_list = @($ScriptArguments)
if ($arg_list.Count -gt 0 -and $arg_list[0] -eq '/nowait') {
    $arg_list = @($arg_list | Select-Object -Skip 1)
}
# Add script-specific processing below.

if ($arg_list.Count -gt 1) {
    throw 'update_common_mods accepts at most one path argument.'
}
$arg_file_infos = @(foreach ($arg_item in $arg_list) {
    Resolve-FileSystemPath -Path $arg_item -BasePath $current_dir
})
$arg_path = ''
$arg_dir = ''
$arg_file = ''
if ($arg_file_infos.Count -gt 0) {
    $arg_path = $arg_file_infos[0].Path
    $arg_dir = $arg_file_infos[0].Directory
    $arg_file = $arg_file_infos[0].File
}

Import-Module (Join-Path $PSScriptRoot 'VbaDevTool.psm1') -Force
$exit_code = 0

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

    $exe_path = Resolve-SourceWorkbookVbaDevExecutable

    foreach ($project_root in $project_roots) {
        Write-Host "Updating CommonModules for project: $project_root"
        Invoke-VbaDev -ExecutablePath $exe_path -Arguments @('common-module', 'update', '--project', $project_root)

        $document_contexts = @(Get-VbaDevProjectDocumentContexts -ProjectRoot $project_root | Sort-Object -Property DocumentName)
        foreach ($document_context in $document_contexts) {
            Write-Host "Building document: $($document_context.DocumentName)"
            Invoke-VbaDev -ExecutablePath $exe_path -Arguments @('build', '--project', $project_root, '--document', $document_context.DocumentName, '--interactive', 'false')
        }
    }
}
catch {
    $exit_code = 1
    $child_exit_code = $_.Exception.Data['VbaDevExitCode']
    if ($child_exit_code -is [int] -and $child_exit_code -ne 0) { $exit_code = $child_exit_code }
    Write-Error -ErrorRecord $_ -ErrorAction Continue
}
finally {
    Set-Location -LiteralPath $current_dir
}
if ($exit_code -ne 0) { exit $exit_code }
