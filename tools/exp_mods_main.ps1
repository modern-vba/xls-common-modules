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
if ($arg_list.Count -ne 1) {
    throw 'exp_mods requires exactly one path argument.'
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

try {
    $workbook = Get-RequiredFileTarget -Path $arg_path -Description 'target workbook'
    $source_set_path = Get-WorkbookLocalSourceSetPath -WorkbookPath $workbook.FullName
    if (-not (Test-Path -LiteralPath $source_set_path -PathType Container)) {
        New-Item -ItemType Directory -Path $source_set_path -Force | Out-Null
    }

    Invoke-VbaDev -Arguments @('export', '--from', $workbook.FullName, '--to', $source_set_path)
}
finally {
    Set-Location -LiteralPath $current_dir
}
