[CmdletBinding(PositionalBinding = $false)]
param (
    [ValidateNotNullOrEmpty()]
    [string]$BatPath,

    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$Path = @()
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
$path_arguments = @($Path)
if ($path_arguments.Count -gt 0 -and $path_arguments[0] -eq '/nowait') {
    $path_arguments = @($path_arguments | Select-Object -Skip 1)
}

if ($path_arguments.Count -ne 1) {
    throw 'imp_mods requires exactly one path argument.'
}
$arg_file_infos = @(foreach ($argument in $path_arguments) {
    Resolve-FileSystemPath -Path $argument -BasePath $current_dir
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

try {
    $workbook = Get-RequiredFileTarget -Path $arg_path -Description 'target workbook'
    $source_set_path = Get-WorkbookLocalSourceSetPath -WorkbookPath $workbook.FullName
    if (-not (Test-Path -LiteralPath $source_set_path -PathType Container)) {
        throw "workbook-local source set was not found: $source_set_path"
    }

    Invoke-VbaDev -Arguments @('import', '--from', $source_set_path, '--to', $workbook.FullName)
}
finally {
    Set-Location -LiteralPath $current_dir
}
