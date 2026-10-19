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
} else {
    $bat_dir = $current_dir
    $bat_file = ""
    $bat_path = $current_dir
}

# /nowait controls the BAT wrapper's wait, and is not a target path.
$arg_list = @($ScriptArguments)
if ($arg_list.Count -gt 0 -and $arg_list[0] -eq '/nowait') {
    $arg_list = @($arg_list | Select-Object -Skip 1)
}

# Add script-specific processing below.
$arg_file_infos = @(foreach ($arg in $arg_list) {
    Resolve-FileSystemPath -Path $arg -BasePath $current_dir
})
"bat_dir: " + $bat_dir
"bat_file: " + $bat_file
"bat_path: " + $bat_path

foreach ($arg_file_info in $arg_file_infos) {
    "arg_dir: " + $arg_file_info.Directory
    "arg_file: " + $arg_file_info.File
    "arg_path: " + $arg_file_info.Path
}
