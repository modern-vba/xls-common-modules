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

function Get-SyncConfigPath {
    if (-not [string]::IsNullOrWhiteSpace($arg_path)) {
        $item = Get-Item -LiteralPath $arg_path -ErrorAction Stop
        if ($item.PSIsContainer) {
            return (Join-Path $item.FullName 'sync.json')
        }
        return $item.FullName
    }

    $current_config = Join-Path $current_dir 'sync.json'
    if (Test-Path -LiteralPath $current_config -PathType Leaf) {
        return $current_config
    }

    return (Join-Path $bat_dir 'sync.json')
}

function Get-RelativeOrAbsolutePath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$BaseDirectory,
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if ([System.IO.Path]::IsPathRooted($Path)) {
        return [System.IO.Path]::GetFullPath($Path)
    }

    return [System.IO.Path]::GetFullPath((Join-Path $BaseDirectory $Path))
}

try {
    $config_path = Get-SyncConfigPath
    if (-not (Test-Path -LiteralPath $config_path -PathType Leaf)) {
        throw "sync.json was not found: $config_path"
    }

    $config_directory = Split-Path -Parent (Resolve-Path -LiteralPath $config_path).ProviderPath
    $script_config = Get-Content -LiteralPath $config_path -Raw | ConvertFrom-Json
    if ($null -eq $script_config.targets) {
        throw 'sync.json must define targets entries with project and document.'
    }
    if ($null -eq $script_config.modules) {
        throw 'sync.json must define modules.'
    }

    $target_contexts = @()
    foreach ($target in @($script_config.targets)) {
        if ([string]::IsNullOrWhiteSpace($target.project) -or [string]::IsNullOrWhiteSpace($target.document)) {
            throw 'Each sync target must define project and document.'
        }

        $project_root = Get-RelativeOrAbsolutePath -BaseDirectory $config_directory -Path $target.project
        $target_contexts += Get-VbaDevProjectDocumentContext -ProjectRoot $project_root -DocumentName $target.document
    }
    if ($target_contexts.Count -eq 0) {
        throw 'sync.json targets must not be empty.'
    }

    $changed_workbooks = New-Object 'System.Collections.Generic.Dictionary[string, object]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($module_name in @($script_config.modules)) {
        if ([string]::IsNullOrWhiteSpace($module_name)) {
            throw 'sync.json modules must not contain empty entries.'
        }

        $candidate_files = @()
        foreach ($context in $target_contexts) {
            $candidate_path = Join-Path $context.SourceSetPath $module_name
            if (Test-Path -LiteralPath $candidate_path -PathType Leaf) {
                $candidate_files += Get-Item -LiteralPath $candidate_path
            }
        }
        if ($candidate_files.Count -eq 0) {
            throw "Synchronized module was not found in any target source set: $module_name"
        }

        $newest_file = $candidate_files | Sort-Object -Property LastWriteTimeUtc, FullName -Descending | Select-Object -First 1
        foreach ($context in $target_contexts) {
            $destination_path = Join-Path $context.SourceSetPath $module_name
            if (Test-SamePath -LeftPath $newest_file.FullName -RightPath $destination_path) {
                Write-Information "Skip $($context.DocumentName): $module_name is already newest."
                continue
            }

            if (Test-FileContentEqual -LeftPath $newest_file.FullName -RightPath $destination_path) {
                Write-Information "Skip $($context.DocumentName): $module_name content is unchanged."
                continue
            }

            $destination_directory = Split-Path -Parent $destination_path
            if (-not (Test-Path -LiteralPath $destination_directory -PathType Container)) {
                New-Item -ItemType Directory -Path $destination_directory -Force | Out-Null
            }
            Copy-Item -LiteralPath $newest_file.FullName -Destination $destination_path -Force
            Write-Information "Copied $module_name to $($context.SourceSetPath)."
            $changed_workbooks[$context.TemplatePath] = $context.SourceSetPath
        }
    }

    foreach ($workbook_path in ($changed_workbooks.Keys | Sort-Object)) {
        Invoke-VbaDev -Arguments @('import', '--from', $changed_workbooks[$workbook_path], '--to', $workbook_path)
    }
}
finally {
    Set-Location $current_dir
}
