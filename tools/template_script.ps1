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
} else {
    $bat_dir = $current_dir
    $bat_file = ""
    $bat_path = $current_dir
}

# /nowait controls the BAT wrapper's wait, and is not a target path.
$path_arguments = @($Path)
if ($path_arguments.Count -gt 0 -and $path_arguments[0] -eq '/nowait') {
    $path_arguments = @($path_arguments | Select-Object -Skip 1)
}

# Add script-specific processing below.
$arg_file_infos = @(foreach ($arg in $path_arguments) {
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
