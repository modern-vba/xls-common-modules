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
    $workbook = Get-RequiredFileTarget -Path $arg_path -Description 'target workbook'
    $source_set_path = Get-WorkbookLocalSourceSetPath -WorkbookPath $workbook.FullName
    if (-not (Test-Path -LiteralPath $source_set_path -PathType Container)) {
        New-Item -ItemType Directory -Path $source_set_path -Force | Out-Null
    }

    Invoke-VbaDev -Arguments @('export', '--from', $workbook.FullName, '--to', $source_set_path)
}
finally {
    Set-Location $current_dir
}
