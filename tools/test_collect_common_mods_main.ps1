$ErrorActionPreference = 'Stop'
$InformationPreference = 'Continue'

$collect_script = Join-Path $PSScriptRoot 'collect_common_mods_main.ps1'
$test_host_executable = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName

function Assert-True {
    param(
        [Parameter(Mandatory = $true)]
        [bool]$Condition,

        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    if (-not $Condition) {
        throw $Message
    }
}

function Assert-Equal {
    param(
        [AllowNull()]
        [object]$Expected,

        [AllowNull()]
        [object]$Actual,

        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    if ($Expected -ne $Actual) {
        throw "$Message Expected '$Expected' but got '$Actual'."
    }
}

function Write-TestFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$Content,

        [Parameter(Mandatory = $true)]
        [System.Text.Encoding]$Encoding
    )

    $parent = Split-Path -Parent $Path
    if (-not (Test-Path -LiteralPath $parent -PathType Container)) {
        New-Item -ItemType Directory -Path $parent -Force | Out-Null
    }
    [System.IO.File]::WriteAllText($Path, $Content, $Encoding)
}

function New-TestProject {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProjectRoot,

        [Parameter(Mandatory = $true)]
        [string]$DocumentName
    )

    $source_set = Join-Path $ProjectRoot "src\$DocumentName"
    New-Item -ItemType Directory -Path $source_set -Force | Out-Null
    $project = [ordered]@{
        schemaVersion = 1
        projectName = (Split-Path -Leaf $ProjectRoot)
        primaryDocument = $DocumentName
        documents = [ordered]@{
            $DocumentName = [ordered]@{
                kind = 'excel'
                sourcePath = "src/$DocumentName"
                templatePath = "src/$DocumentName/$DocumentName.xlsm"
                binPath = "bin/$DocumentName.xlsm"
                publishPath = "publish/$DocumentName.xlsm"
                commonModules = @()
                references = @()
            }
        }
    }
    Write-TestFile -Path (Join-Path $ProjectRoot 'vba-project.json') -Content ($project | ConvertTo-Json -Depth 8) -Encoding (New-Object System.Text.UTF8Encoding($false))
    return $source_set
}

function New-TestJunction {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$Target
    )

    $parent = Split-Path -Parent $Path
    if (-not (Test-Path -LiteralPath $parent -PathType Container)) {
        New-Item -ItemType Directory -Path $parent -Force | Out-Null
    }
    New-Item -ItemType Junction -Path $Path -Target $Target | Out-Null
}

function New-TestFileSymbolicLink {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$Target
    )

    $parent = Split-Path -Parent $Path
    if (-not (Test-Path -LiteralPath $parent -PathType Container)) {
        New-Item -ItemType Directory -Path $parent -Force | Out-Null
    }
    $mklink_command = "mklink `"$Path`" `"$Target`""
    $mklink_output = & $env:ComSpec /d /c $mklink_command 2>&1
    if ($LASTEXITCODE -ne 0) {
        throw "Could not create test file symbolic link '$Path': $($mklink_output | Out-String)"
    }
}

function Invoke-Collect {
    param(
        [Parameter(Mandatory = $true)]
        [string]$WorkingDirectory,

        [string[]]$Arguments = @()
    )

    Push-Location $WorkingDirectory
    $previous_error_action_preference = $ErrorActionPreference
    try {
        $ErrorActionPreference = 'Continue'
        $output = & $test_host_executable -NoProfile -ExecutionPolicy Bypass -File $collect_script @Arguments 2>&1
        $exit_code = $LASTEXITCODE
    }
    finally {
        $ErrorActionPreference = $previous_error_action_preference
        Pop-Location
    }

    return [pscustomobject]@{
        ExitCode = $exit_code
        Output = ($output | Out-String)
    }
}

function Assert-CollectFailure {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Result,

        [Parameter(Mandatory = $true)]
        [string]$ExpectedMessage
    )

    Assert-Equal -Expected 1 -Actual $Result.ExitCode -Message "COLLECT should fail. Output: $($Result.Output)"
    Assert-True -Condition ($Result.Output -like "*$ExpectedMessage*") -Message "COLLECT failure did not contain '$ExpectedMessage'. Output: $($Result.Output)"
}

$temp_root = Join-Path ([System.IO.Path]::GetTempPath()) ('collect-common-modules-test-' + [System.Guid]::NewGuid().ToString('N'))
$utf16_le = New-Object System.Text.UnicodeEncoding($false, $true, $true)
$shift_jis = [System.Text.Encoding]::GetEncoding(932)

try {
    $search_root = Join-Path $temp_root 'explicit-search-root'
    $working_directory = Join-Path $temp_root 'invocation-working-directory'
    New-Item -ItemType Directory -Path $search_root, $working_directory -Force | Out-Null
    $authoring_source_set = New-TestProject -ProjectRoot (Join-Path $search_root 'AuthorProject') -DocumentName 'AuthorBook'
    $manifest_text = "ModuleFile`tCategories`tDependencies`tRequiredReferences`r`nShared.bas`truntime-baseline`t`t[]`r`n"
    Write-TestFile -Path (Join-Path $authoring_source_set 'common-modules-manifest.tsv') -Content $manifest_text -Encoding $utf16_le
    Write-TestFile -Path (Join-Path $authoring_source_set 'Shared.bas') -Content "Attribute VB_Name = `"Shared`"`r`n'explicit search root`r`n" -Encoding $shift_jis
    Write-TestFile -Path (Join-Path $authoring_source_set 'AuthorOnly.bas') -Content "Attribute VB_Name = `"AuthorOnly`"`r`n" -Encoding $shift_jis

    $result = Invoke-Collect -WorkingDirectory $working_directory -Arguments @($search_root)
    Assert-Equal -Expected 0 -Actual $result.ExitCode -Message "COLLECT failed. Output: $($result.Output)"
    $output_repository = Join-Path $working_directory 'common_modules_repo'
    Assert-True -Condition (Test-Path -LiteralPath $output_repository -PathType Container) -Message 'COLLECT did not create common_modules_repo directly below the invocation working directory.'
    Assert-True -Condition (-not (Test-Path -LiteralPath (Join-Path $search_root 'common_modules_repo'))) -Message 'COLLECT wrote output below the Collection Search Root.'
    Assert-True -Condition (Test-Path -LiteralPath (Join-Path $output_repository 'Shared.bas') -PathType Leaf) -Message 'COLLECT ignored the explicit Collection Search Root.'
    $inventory = @(Get-ChildItem -LiteralPath $output_repository -Force | ForEach-Object { $_.Name } | Sort-Object)
    Assert-Equal -Expected 2 -Actual $inventory.Count -Message 'COLLECT did not create a closed one-module package.'
    Assert-True -Condition (-not (Test-Path -LiteralPath (Join-Path $output_repository 'AuthorOnly.bas'))) -Message 'COLLECT copied an unlisted authoring source into the closed package.'
    Assert-True -Condition ((Get-Content -LiteralPath (Join-Path $output_repository 'Shared.bas') -Raw) -like '*explicit search root*') -Message 'COLLECT selected unexpected module bytes.'

    Write-TestFile -Path (Join-Path $output_repository 'obsolete.bas') -Content "Attribute VB_Name = `"obsolete`"`r`n" -Encoding $shift_jis
    Write-TestFile -Path (Join-Path $output_repository 'unexpected\nested.txt') -Content 'unexpected' -Encoding (New-Object System.Text.UTF8Encoding($false))
    $cleanup_link_target = Join-Path $temp_root 'cleanup-link-target'
    New-Item -ItemType Directory -Path $cleanup_link_target -Force | Out-Null
    Write-TestFile -Path (Join-Path $cleanup_link_target 'target-sentinel.txt') -Content 'preserve target' -Encoding $shift_jis
    New-TestJunction -Path (Join-Path $output_repository 'unexpected-link') -Target $cleanup_link_target
    $cleanup_result = Invoke-Collect -WorkingDirectory $working_directory -Arguments @($search_root)
    Assert-Equal -Expected 0 -Actual $cleanup_result.ExitCode -Message "COLLECT failed to replace a stale package. Output: $($cleanup_result.Output)"
    $cleaned_inventory = @(Get-ChildItem -LiteralPath $output_repository -Force | ForEach-Object { $_.Name } | Sort-Object)
    Assert-Equal -Expected 2 -Actual $cleaned_inventory.Count -Message 'COLLECT did not clear every obsolete or unexpected output entry.'
    Assert-True -Condition (-not (Test-Path -LiteralPath (Join-Path $output_repository 'obsolete.bas'))) -Message 'COLLECT left an obsolete source unit in the closed package.'
    Assert-True -Condition (-not (Test-Path -LiteralPath (Join-Path $output_repository 'unexpected'))) -Message 'COLLECT left an unexpected directory in the closed package.'
    Assert-True -Condition (-not (Test-Path -LiteralPath (Join-Path $output_repository 'unexpected-link'))) -Message 'COLLECT left an unexpected reparse child in the closed package.'
    Assert-True -Condition (Test-Path -LiteralPath (Join-Path $cleanup_link_target 'target-sentinel.txt') -PathType Leaf) -Message 'COLLECT followed an output reparse child and mutated its target.'
    $unchanged_module_path = Join-Path $output_repository 'Shared.bas'
    $unchanged_module_time = [System.IO.File]::GetLastWriteTimeUtc($unchanged_module_path)
    $exclusive_output_stream = [System.IO.File]::Open($unchanged_module_path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::None)
    try {
        $unchanged_result = Invoke-Collect -WorkingDirectory $working_directory -Arguments @($search_root)
    }
    finally {
        $exclusive_output_stream.Dispose()
    }
    Assert-Equal -Expected 0 -Actual $unchanged_result.ExitCode -Message "COLLECT rewrote a metadata-identical package instead of reporting UNCHANGED. Output: $($unchanged_result.Output)"
    Assert-True -Condition ($unchanged_result.Output -like '*UNCHANGED*') -Message "COLLECT did not report UNCHANGED for an exact metadata match. Output: $($unchanged_result.Output)"
    Assert-Equal -Expected $unchanged_module_time -Actual ([System.IO.File]::GetLastWriteTimeUtc($unchanged_module_path)) -Message 'COLLECT changed an UNCHANGED package timestamp.'

    $relative_working_directory = Join-Path $temp_root 'relative-invocation-working-directory'
    New-Item -ItemType Directory -Path $relative_working_directory -Force | Out-Null
    $relative_search_root = '..\explicit-search-root'
    $relative_result = Invoke-Collect -WorkingDirectory $relative_working_directory -Arguments @($relative_search_root)
    Assert-Equal -Expected 0 -Actual $relative_result.ExitCode -Message "COLLECT did not resolve a relative Search Root against the invocation working directory. Output: $($relative_result.Output)"
    Assert-True -Condition (Test-Path -LiteralPath (Join-Path $relative_working_directory 'common_modules_repo\Shared.bas') -PathType Leaf) -Message 'Relative-root COLLECT wrote to an unexpected output location.'

    Assert-CollectFailure -Result (Invoke-Collect -WorkingDirectory $working_directory) -ExpectedMessage 'exactly one Collection Search Root'
    Assert-CollectFailure -Result (Invoke-Collect -WorkingDirectory $working_directory -Arguments @($search_root, $search_root)) -ExpectedMessage 'exactly one Collection Search Root'
    Assert-CollectFailure -Result (Invoke-Collect -WorkingDirectory $working_directory -Arguments @('')) -ExpectedMessage 'Collection Search Root'
    Assert-CollectFailure -Result (Invoke-Collect -WorkingDirectory $working_directory -Arguments @((Join-Path $temp_root 'absent-search-root'))) -ExpectedMessage 'absent-search-root'
    $search_root_file = Join-Path $temp_root 'search-root-file.txt'
    Write-TestFile -Path $search_root_file -Content 'not a directory' -Encoding (New-Object System.Text.UTF8Encoding($false))
    Assert-CollectFailure -Result (Invoke-Collect -WorkingDirectory $working_directory -Arguments @($search_root_file)) -ExpectedMessage 'must be a directory'

    $case_only_search_root = Join-Path $temp_root 'case-only-project-manifest'
    $case_only_project_root = Join-Path $case_only_search_root 'Project'
    $case_only_source_set = New-TestProject -ProjectRoot $case_only_project_root -DocumentName 'Book'
    $exact_project_manifest = Join-Path $case_only_project_root 'vba-project.json'
    $temporary_project_manifest = Join-Path $case_only_project_root 'manifest.tmp'
    Move-Item -LiteralPath $exact_project_manifest -Destination $temporary_project_manifest
    Move-Item -LiteralPath $temporary_project_manifest -Destination (Join-Path $case_only_project_root 'VBA-PROJECT.JSON')
    Write-TestFile -Path (Join-Path $case_only_source_set 'common-modules-manifest.tsv') -Content $manifest_text -Encoding $utf16_le
    Write-TestFile -Path (Join-Path $case_only_source_set 'Shared.bas') -Content "Attribute VB_Name = `"Shared`"`r`n'case-only authority`r`n" -Encoding $shift_jis
    $case_only_working_directory = Join-Path $temp_root 'case-only-working-directory'
    $case_only_output = Join-Path $case_only_working_directory 'common_modules_repo'
    New-Item -ItemType Directory -Path $case_only_output -Force | Out-Null
    Write-TestFile -Path (Join-Path $case_only_output 'sentinel.txt') -Content 'preserve before preflight' -Encoding (New-Object System.Text.UTF8Encoding($false))
    $case_only_result = Invoke-Collect -WorkingDirectory $case_only_working_directory -Arguments @($case_only_search_root)
    Assert-CollectFailure -Result $case_only_result -ExpectedMessage 'ordinal-exact vba-project.json'
    Assert-True -Condition (Test-Path -LiteralPath (Join-Path $case_only_output 'sentinel.txt') -PathType Leaf) -Message 'COLLECT mutated output before rejecting a case-only project manifest.'

    $boundary_real_root = Join-Path $temp_root 'boundary-real-root'
    $boundary_authoring_project_root = Join-Path $boundary_real_root 'AuthorProject'
    $boundary_authoring_source_set = New-TestProject -ProjectRoot $boundary_authoring_project_root -DocumentName 'AuthorBook'
    $boundary_authoring_project_path = Join-Path $boundary_authoring_project_root 'vba-project.json'
    $boundary_authoring_project = Get-Content -LiteralPath $boundary_authoring_project_path -Raw | ConvertFrom-Json
    $boundary_authoring_project.documents | Add-Member -MemberType NoteProperty -Name 'AuthorBookDuplicate' -Value ([pscustomobject]@{ sourcePath = 'src/AuthorBook/.' })
    Write-TestFile -Path $boundary_authoring_project_path -Content ($boundary_authoring_project | ConvertTo-Json -Depth 8) -Encoding (New-Object System.Text.UTF8Encoding($false))
    Write-TestFile -Path (Join-Path $boundary_authoring_source_set 'common-modules-manifest.tsv') -Content $manifest_text -Encoding $utf16_le
    $boundary_authoring_module = Join-Path $boundary_authoring_source_set 'Shared.bas'
    Write-TestFile -Path $boundary_authoring_module -Content "Attribute VB_Name = `"Shared`"`r`n'boundary authoring`r`n" -Encoding $shift_jis
    [System.IO.File]::SetLastWriteTimeUtc($boundary_authoring_module, [datetime]'2025-01-01T00:00:00Z')

    $boundary_external_source = Join-Path $temp_root 'boundary-external-source'
    New-Item -ItemType Directory -Path $boundary_external_source -Force | Out-Null
    $boundary_external_module = Join-Path $boundary_external_source 'shared.BAS'
    Write-TestFile -Path $boundary_external_module -Content "Attribute VB_Name = `"Shared`"`r`n'junction source root`r`n" -Encoding $shift_jis
    [System.IO.File]::SetLastWriteTimeUtc($boundary_external_module, [datetime]'2025-01-03T00:00:00Z')

    $boundary_ignored_source = Join-Path $temp_root 'boundary-ignored-source'
    New-Item -ItemType Directory -Path $boundary_ignored_source -Force | Out-Null
    Write-TestFile -Path (Join-Path $boundary_ignored_source 'Shared.bas') -Content "Attribute VB_Name = `"Shared`"`r`n'ignored reparse child`r`n" -Encoding $shift_jis
    New-TestJunction -Path (Join-Path $boundary_external_source 'nested-link') -Target $boundary_ignored_source

    $boundary_linked_project = Join-Path $boundary_real_root 'LinkedProject'
    $boundary_linked_source = Join-Path $boundary_linked_project 'src\LinkedBook'
    New-TestJunction -Path $boundary_linked_source -Target $boundary_external_source
    $boundary_project = [ordered]@{
        schemaVersion = 1
        projectName = 'LinkedProject'
        primaryDocument = 'First'
        documents = [ordered]@{
            First = [ordered]@{ kind = 'excel'; sourcePath = 'src/LinkedBook' }
            Second = [ordered]@{ kind = 'excel'; sourcePath = 'src/LinkedBook/.' }
        }
    }
    Write-TestFile -Path (Join-Path $boundary_linked_project 'vba-project.json') -Content ($boundary_project | ConvertTo-Json -Depth 8) -Encoding (New-Object System.Text.UTF8Encoding($false))

    $excluded_project_root = Join-Path $boundary_real_root '.GiT\IgnoredProject'
    New-Item -ItemType Directory -Path $excluded_project_root -Force | Out-Null
    Write-TestFile -Path (Join-Path $excluded_project_root 'vba-project.json') -Content '{ invalid excluded json' -Encoding (New-Object System.Text.UTF8Encoding($false))
    $boundary_ignored_discovery = Join-Path $temp_root 'boundary-ignored-discovery'
    New-Item -ItemType Directory -Path $boundary_ignored_discovery -Force | Out-Null
    Write-TestFile -Path (Join-Path $boundary_ignored_discovery 'vba-project.json') -Content '{ invalid linked json' -Encoding (New-Object System.Text.UTF8Encoding($false))
    New-TestJunction -Path (Join-Path $boundary_real_root 'linked-discovery') -Target $boundary_ignored_discovery

    $boundary_search_root = Join-Path $temp_root 'boundary-search-root-junction'
    New-TestJunction -Path $boundary_search_root -Target $boundary_real_root
    $boundary_working_directory = Join-Path $temp_root 'boundary-working-directory'
    New-Item -ItemType Directory -Path $boundary_working_directory -Force | Out-Null
    $boundary_result = Invoke-Collect -WorkingDirectory $boundary_working_directory -Arguments @($boundary_search_root)
    Assert-Equal -Expected 0 -Actual $boundary_result.ExitCode -Message "COLLECT did not honor discovery and source-set reparse boundaries. Output: $($boundary_result.Output)"
    $boundary_output = Join-Path $boundary_working_directory 'common_modules_repo'
    $boundary_inventory = @(Get-ChildItem -LiteralPath $boundary_output -Force)
    Assert-Equal -Expected 2 -Actual $boundary_inventory.Count -Message 'COLLECT included an excluded or reparse-child entry in the boundary package.'
    Assert-True -Condition (Test-Path -LiteralPath (Join-Path $boundary_output 'Shared.bas') -PathType Leaf) -Message 'COLLECT did not restore the canonical manifest basename spelling.'
    Assert-True -Condition ((Get-Content -LiteralPath (Join-Path $boundary_output 'Shared.bas') -Raw) -like '*junction source root*') -Message 'COLLECT did not scan the explicitly resolved reparse source-set root.'

    $absolute_source_search_root = Join-Path $temp_root 'absolute-source-path'
    $absolute_source_authoring = New-TestProject -ProjectRoot (Join-Path $absolute_source_search_root 'AuthorProject') -DocumentName 'AuthorBook'
    Write-TestFile -Path (Join-Path $absolute_source_authoring 'common-modules-manifest.tsv') -Content $manifest_text -Encoding $utf16_le
    $absolute_source_authoring_module = Join-Path $absolute_source_authoring 'Shared.bas'
    Write-TestFile -Path $absolute_source_authoring_module -Content "Attribute VB_Name = `"Shared`"`r`n'absolute authoring`r`n" -Encoding $shift_jis
    [System.IO.File]::SetLastWriteTimeUtc($absolute_source_authoring_module, [datetime]'2025-01-01T00:00:00Z')
    $absolute_external_source = Join-Path $temp_root 'absolute-external-source'
    New-Item -ItemType Directory -Path $absolute_external_source -Force | Out-Null
    $absolute_external_module = Join-Path $absolute_external_source 'Shared.bas'
    Write-TestFile -Path $absolute_external_module -Content "Attribute VB_Name = `"Shared`"`r`n'absolute external source`r`n" -Encoding $shift_jis
    [System.IO.File]::SetLastWriteTimeUtc($absolute_external_module, [datetime]'2025-01-04T00:00:00Z')
    $absolute_source_project = Join-Path $absolute_source_search_root 'AbsoluteProject'
    New-Item -ItemType Directory -Path $absolute_source_project -Force | Out-Null
    $absolute_source_project_json = [ordered]@{
        documents = [ordered]@{
            AbsoluteBook = [ordered]@{ sourcePath = $absolute_external_source }
        }
    }
    Write-TestFile -Path (Join-Path $absolute_source_project 'vba-project.json') -Content ($absolute_source_project_json | ConvertTo-Json -Depth 5) -Encoding (New-Object System.Text.UTF8Encoding($false))
    $absolute_source_working_directory = Join-Path $temp_root 'absolute-source-working-directory'
    New-Item -ItemType Directory -Path $absolute_source_working_directory -Force | Out-Null
    $absolute_source_result = Invoke-Collect -WorkingDirectory $absolute_source_working_directory -Arguments @($absolute_source_search_root)
    Assert-Equal -Expected 0 -Actual $absolute_source_result.ExitCode -Message "COLLECT did not accept an absolute document sourcePath. Output: $($absolute_source_result.Output)"
    Assert-True -Condition ((Get-Content -LiteralPath (Join-Path $absolute_source_working_directory 'common_modules_repo\Shared.bas') -Raw) -like '*absolute external source*') -Message 'COLLECT did not include the absolute source-set candidate.'

    $hidden_search_root = Join-Path $temp_root 'hidden-search-root'
    $hidden_source_set = New-TestProject -ProjectRoot (Join-Path $hidden_search_root 'AuthorProject') -DocumentName 'AuthorBook'
    Write-TestFile -Path (Join-Path $hidden_source_set 'common-modules-manifest.tsv') -Content $manifest_text -Encoding $utf16_le
    Write-TestFile -Path (Join-Path $hidden_source_set 'Shared.bas') -Content "Attribute VB_Name = `"Shared`"`r`n'hidden roots`r`n" -Encoding $shift_jis
    $hidden_source_item = Get-Item -LiteralPath $hidden_source_set
    $hidden_source_item.Attributes = $hidden_source_item.Attributes -bor [System.IO.FileAttributes]::Hidden
    $hidden_search_item = Get-Item -LiteralPath $hidden_search_root
    $hidden_search_item.Attributes = $hidden_search_item.Attributes -bor [System.IO.FileAttributes]::Hidden
    $hidden_working_directory = Join-Path $temp_root 'hidden-root-working-directory'
    New-Item -ItemType Directory -Path $hidden_working_directory -Force | Out-Null
    $hidden_result = Invoke-Collect -WorkingDirectory $hidden_working_directory -Arguments @($hidden_search_root)
    Assert-Equal -Expected 0 -Actual $hidden_result.ExitCode -Message "COLLECT rejected readable hidden Search Root or source-set roots. Output: $($hidden_result.Output)"
    Assert-True -Condition ((Get-Content -LiteralPath (Join-Path $hidden_working_directory 'common_modules_repo\Shared.bas') -Raw) -like '*hidden roots*') -Message 'COLLECT did not package a readable hidden source set.'

    $invalid_project_search_root = Join-Path $temp_root 'invalid-project-json'
    $invalid_project_root = Join-Path $invalid_project_search_root 'InvalidProject'
    New-Item -ItemType Directory -Path $invalid_project_root -Force | Out-Null
    Write-TestFile -Path (Join-Path $invalid_project_root 'vba-project.json') -Content '{ invalid project json' -Encoding (New-Object System.Text.UTF8Encoding($false))
    $invalid_project_working_directory = Join-Path $temp_root 'invalid-project-working-directory'
    $invalid_project_output = Join-Path $invalid_project_working_directory 'common_modules_repo'
    New-Item -ItemType Directory -Path $invalid_project_output -Force | Out-Null
    Write-TestFile -Path (Join-Path $invalid_project_output 'sentinel.txt') -Content 'preserve' -Encoding $shift_jis
    $invalid_project_result = Invoke-Collect -WorkingDirectory $invalid_project_working_directory -Arguments @($invalid_project_search_root)
    Assert-Equal -Expected 1 -Actual $invalid_project_result.ExitCode -Message "COLLECT accepted an invalid project manifest. Output: $($invalid_project_result.Output)"
    Assert-True -Condition (Test-Path -LiteralPath (Join-Path $invalid_project_output 'sentinel.txt') -PathType Leaf) -Message 'COLLECT mutated output before rejecting an invalid project manifest.'

    $absent_source_search_root = Join-Path $temp_root 'absent-project-source'
    $absent_source_project = Join-Path $absent_source_search_root 'AbsentSourceProject'
    New-Item -ItemType Directory -Path $absent_source_project -Force | Out-Null
    $absent_source_project_json = [ordered]@{
        documents = [ordered]@{
            MissingBook = [ordered]@{ sourcePath = 'missing/source' }
        }
    }
    Write-TestFile -Path (Join-Path $absent_source_project 'vba-project.json') -Content ($absent_source_project_json | ConvertTo-Json -Depth 5) -Encoding (New-Object System.Text.UTF8Encoding($false))
    $absent_source_working_directory = Join-Path $temp_root 'absent-source-working-directory'
    $absent_source_output = Join-Path $absent_source_working_directory 'common_modules_repo'
    New-Item -ItemType Directory -Path $absent_source_output -Force | Out-Null
    Write-TestFile -Path (Join-Path $absent_source_output 'sentinel.txt') -Content 'preserve' -Encoding $shift_jis
    $absent_source_result = Invoke-Collect -WorkingDirectory $absent_source_working_directory -Arguments @($absent_source_search_root)
    Assert-CollectFailure -Result $absent_source_result -ExpectedMessage 'sourcePath was not found'
    Assert-True -Condition (Test-Path -LiteralPath (Join-Path $absent_source_output 'sentinel.txt') -PathType Leaf) -Message 'COLLECT mutated output before rejecting an absent project source directory.'

    $invalid_source_type_search_root = Join-Path $temp_root 'invalid-source-path-type'
    $invalid_source_type_authoring = New-TestProject -ProjectRoot (Join-Path $invalid_source_type_search_root 'AuthorProject') -DocumentName 'AuthorBook'
    Write-TestFile -Path (Join-Path $invalid_source_type_authoring 'common-modules-manifest.tsv') -Content $manifest_text -Encoding $utf16_le
    Write-TestFile -Path (Join-Path $invalid_source_type_authoring 'Shared.bas') -Content "Attribute VB_Name = `"Shared`"`r`n" -Encoding $shift_jis
    $invalid_source_type_project = Join-Path $invalid_source_type_search_root 'InvalidProject'
    New-Item -ItemType Directory -Path (Join-Path $invalid_source_type_project '123') -Force | Out-Null
    $invalid_source_type_json = [ordered]@{
        documents = [ordered]@{
            InvalidBook = [ordered]@{ sourcePath = 123 }
        }
    }
    Write-TestFile -Path (Join-Path $invalid_source_type_project 'vba-project.json') -Content ($invalid_source_type_json | ConvertTo-Json -Depth 5) -Encoding (New-Object System.Text.UTF8Encoding($false))
    $invalid_source_type_working_directory = Join-Path $temp_root 'invalid-source-type-working-directory'
    $invalid_source_type_output = Join-Path $invalid_source_type_working_directory 'common_modules_repo'
    New-Item -ItemType Directory -Path $invalid_source_type_output -Force | Out-Null
    Write-TestFile -Path (Join-Path $invalid_source_type_output 'sentinel.txt') -Content 'preserve' -Encoding $shift_jis
    $invalid_source_type_result = Invoke-Collect -WorkingDirectory $invalid_source_type_working_directory -Arguments @($invalid_source_type_search_root)
    Assert-CollectFailure -Result $invalid_source_type_result -ExpectedMessage 'nonempty string sourcePath'
    Assert-True -Condition (Test-Path -LiteralPath (Join-Path $invalid_source_type_output 'sentinel.txt') -PathType Leaf) -Message 'COLLECT mutated output before rejecting a non-string sourcePath.'

    $empty_documents_search_root = Join-Path $temp_root 'empty-project-documents'
    $empty_documents_authoring = New-TestProject -ProjectRoot (Join-Path $empty_documents_search_root 'AuthorProject') -DocumentName 'AuthorBook'
    Write-TestFile -Path (Join-Path $empty_documents_authoring 'common-modules-manifest.tsv') -Content $manifest_text -Encoding $utf16_le
    Write-TestFile -Path (Join-Path $empty_documents_authoring 'Shared.bas') -Content "Attribute VB_Name = `"Shared`"`r`n" -Encoding $shift_jis
    $empty_documents_project = Join-Path $empty_documents_search_root 'EmptyProject'
    New-Item -ItemType Directory -Path $empty_documents_project -Force | Out-Null
    Write-TestFile -Path (Join-Path $empty_documents_project 'vba-project.json') -Content '{"documents":{}}' -Encoding (New-Object System.Text.UTF8Encoding($false))
    $empty_documents_working_directory = Join-Path $temp_root 'empty-documents-working-directory'
    $empty_documents_output = Join-Path $empty_documents_working_directory 'common_modules_repo'
    New-Item -ItemType Directory -Path $empty_documents_output -Force | Out-Null
    Write-TestFile -Path (Join-Path $empty_documents_output 'sentinel.txt') -Content 'preserve' -Encoding $shift_jis
    $empty_documents_result = Invoke-Collect -WorkingDirectory $empty_documents_working_directory -Arguments @($empty_documents_search_root)
    Assert-CollectFailure -Result $empty_documents_result -ExpectedMessage 'nonempty documents object'
    Assert-True -Condition (Test-Path -LiteralPath (Join-Path $empty_documents_output 'sentinel.txt') -PathType Leaf) -Message 'COLLECT mutated output before rejecting an empty documents object.'

    $reparse_project_search_root = Join-Path $temp_root 'reparse-project-manifest'
    $reparse_project_root = Join-Path $reparse_project_search_root 'ReparseProject'
    $reparse_project_target = Join-Path $temp_root 'reparse-project-target'
    New-Item -ItemType Directory -Path $reparse_project_root, $reparse_project_target -Force | Out-Null
    New-TestJunction -Path (Join-Path $reparse_project_root 'vba-project.json') -Target $reparse_project_target
    $reparse_project_working_directory = Join-Path $temp_root 'reparse-project-working-directory'
    $reparse_project_output = Join-Path $reparse_project_working_directory 'common_modules_repo'
    New-Item -ItemType Directory -Path $reparse_project_output -Force | Out-Null
    Write-TestFile -Path (Join-Path $reparse_project_output 'sentinel.txt') -Content 'preserve' -Encoding $shift_jis
    $reparse_project_result = Invoke-Collect -WorkingDirectory $reparse_project_working_directory -Arguments @($reparse_project_search_root)
    Assert-CollectFailure -Result $reparse_project_result -ExpectedMessage 'ordinary non-reparse file'
    Assert-True -Condition (Test-Path -LiteralPath (Join-Path $reparse_project_output 'sentinel.txt') -PathType Leaf) -Message 'COLLECT mutated output before rejecting a reparse project manifest.'

    $zero_authority_search_root = Join-Path $temp_root 'zero-manifest-authority'
    $zero_authority_source_set = New-TestProject -ProjectRoot (Join-Path $zero_authority_search_root 'Project') -DocumentName 'Book'
    Write-TestFile -Path (Join-Path $zero_authority_source_set 'Shared.bas') -Content "Attribute VB_Name = `"Shared`"`r`n" -Encoding $shift_jis
    $zero_authority_working_directory = Join-Path $temp_root 'zero-authority-working-directory'
    $zero_authority_output = Join-Path $zero_authority_working_directory 'common_modules_repo'
    New-Item -ItemType Directory -Path $zero_authority_output -Force | Out-Null
    Write-TestFile -Path (Join-Path $zero_authority_output 'sentinel.txt') -Content 'preserve' -Encoding $shift_jis
    $zero_authority_result = Invoke-Collect -WorkingDirectory $zero_authority_working_directory -Arguments @($zero_authority_search_root)
    Assert-CollectFailure -Result $zero_authority_result -ExpectedMessage 'Exactly one CommonModules Authoring Source Set'
    Assert-True -Condition (Test-Path -LiteralPath (Join-Path $zero_authority_output 'sentinel.txt') -PathType Leaf) -Message 'COLLECT mutated output before rejecting zero manifest authorities.'

    $multiple_authority_search_root = Join-Path $temp_root 'multiple-manifest-authorities'
    $multiple_authority_a = New-TestProject -ProjectRoot (Join-Path $multiple_authority_search_root 'ProjectA') -DocumentName 'BookA'
    $multiple_authority_b = New-TestProject -ProjectRoot (Join-Path $multiple_authority_search_root 'ProjectB') -DocumentName 'BookB'
    foreach ($multiple_authority_source_set in @($multiple_authority_a, $multiple_authority_b)) {
        Write-TestFile -Path (Join-Path $multiple_authority_source_set 'common-modules-manifest.tsv') -Content $manifest_text -Encoding $utf16_le
        Write-TestFile -Path (Join-Path $multiple_authority_source_set 'Shared.bas') -Content "Attribute VB_Name = `"Shared`"`r`n" -Encoding $shift_jis
    }
    $multiple_authority_working_directory = Join-Path $temp_root 'multiple-authority-working-directory'
    $multiple_authority_output = Join-Path $multiple_authority_working_directory 'common_modules_repo'
    New-Item -ItemType Directory -Path $multiple_authority_output -Force | Out-Null
    Write-TestFile -Path (Join-Path $multiple_authority_output 'sentinel.txt') -Content 'preserve' -Encoding $shift_jis
    $multiple_authority_result = Invoke-Collect -WorkingDirectory $multiple_authority_working_directory -Arguments @($multiple_authority_search_root)
    Assert-CollectFailure -Result $multiple_authority_result -ExpectedMessage 'Exactly one CommonModules Authoring Source Set'
    Assert-True -Condition (Test-Path -LiteralPath (Join-Path $multiple_authority_output 'sentinel.txt') -PathType Leaf) -Message 'COLLECT mutated output before rejecting multiple manifest authorities.'

    $case_manifest_search_root = Join-Path $temp_root 'case-only-common-manifest'
    $case_manifest_source_set = New-TestProject -ProjectRoot (Join-Path $case_manifest_search_root 'Project') -DocumentName 'Book'
    Write-TestFile -Path (Join-Path $case_manifest_source_set 'COMMON-MODULES-MANIFEST.TSV') -Content $manifest_text -Encoding $utf16_le
    Write-TestFile -Path (Join-Path $case_manifest_source_set 'Shared.bas') -Content "Attribute VB_Name = `"Shared`"`r`n" -Encoding $shift_jis
    $case_manifest_working_directory = Join-Path $temp_root 'case-manifest-working-directory'
    $case_manifest_output = Join-Path $case_manifest_working_directory 'common_modules_repo'
    New-Item -ItemType Directory -Path $case_manifest_output -Force | Out-Null
    Write-TestFile -Path (Join-Path $case_manifest_output 'sentinel.txt') -Content 'preserve' -Encoding $shift_jis
    $case_manifest_result = Invoke-Collect -WorkingDirectory $case_manifest_working_directory -Arguments @($case_manifest_search_root)
    Assert-CollectFailure -Result $case_manifest_result -ExpectedMessage 'ordinal-exact common-modules-manifest.tsv'
    Assert-True -Condition (Test-Path -LiteralPath (Join-Path $case_manifest_output 'sentinel.txt') -PathType Leaf) -Message 'COLLECT mutated output before rejecting a case-only CommonModules manifest.'

    $reparse_manifest_search_root = Join-Path $temp_root 'reparse-common-manifest'
    $reparse_manifest_source_set = New-TestProject -ProjectRoot (Join-Path $reparse_manifest_search_root 'Project') -DocumentName 'Book'
    $reparse_manifest_target = Join-Path $temp_root 'reparse-common-manifest-target.tsv'
    Write-TestFile -Path $reparse_manifest_target -Content $manifest_text -Encoding $utf16_le
    New-TestFileSymbolicLink -Path (Join-Path $reparse_manifest_source_set 'common-modules-manifest.tsv') -Target $reparse_manifest_target
    Write-TestFile -Path (Join-Path $reparse_manifest_source_set 'Shared.bas') -Content "Attribute VB_Name = `"Shared`"`r`n" -Encoding $shift_jis
    $reparse_manifest_working_directory = Join-Path $temp_root 'reparse-manifest-working-directory'
    $reparse_manifest_output = Join-Path $reparse_manifest_working_directory 'common_modules_repo'
    New-Item -ItemType Directory -Path $reparse_manifest_output -Force | Out-Null
    Write-TestFile -Path (Join-Path $reparse_manifest_output 'sentinel.txt') -Content 'preserve' -Encoding $shift_jis
    $reparse_manifest_result = Invoke-Collect -WorkingDirectory $reparse_manifest_working_directory -Arguments @($reparse_manifest_search_root)
    Assert-CollectFailure -Result $reparse_manifest_result -ExpectedMessage 'ordinary non-reparse file'
    Assert-True -Condition (Test-Path -LiteralPath (Join-Path $reparse_manifest_output 'sentinel.txt') -PathType Leaf) -Message 'COLLECT mutated output before rejecting a reparse CommonModules manifest.'

    $invalid_manifest_search_root = Join-Path $temp_root 'invalid-canonical-manifest'
    $invalid_manifest_source_set = New-TestProject -ProjectRoot (Join-Path $invalid_manifest_search_root 'Project') -DocumentName 'Book'
    $invalid_manifest_text = "WrongHeader`tCategories`tDependencies`tRequiredReferences`r`nShared.bas`truntime-baseline`t`t[]`r`n"
    Write-TestFile -Path (Join-Path $invalid_manifest_source_set 'common-modules-manifest.tsv') -Content $invalid_manifest_text -Encoding $utf16_le
    Write-TestFile -Path (Join-Path $invalid_manifest_source_set 'Shared.bas') -Content "Attribute VB_Name = `"Shared`"`r`n" -Encoding $shift_jis
    $invalid_manifest_working_directory = Join-Path $temp_root 'invalid-manifest-working-directory'
    $invalid_manifest_output = Join-Path $invalid_manifest_working_directory 'common_modules_repo'
    New-Item -ItemType Directory -Path $invalid_manifest_output -Force | Out-Null
    Write-TestFile -Path (Join-Path $invalid_manifest_output 'sentinel.txt') -Content 'preserve' -Encoding $shift_jis
    $invalid_manifest_result = Invoke-Collect -WorkingDirectory $invalid_manifest_working_directory -Arguments @($invalid_manifest_search_root)
    Assert-CollectFailure -Result $invalid_manifest_result -ExpectedMessage 'invalid header'
    Assert-True -Condition (Test-Path -LiteralPath (Join-Path $invalid_manifest_output 'sentinel.txt') -PathType Leaf) -Message 'COLLECT mutated output before rejecting the canonical manifest contract.'

    $ambiguous_search_root = Join-Path $temp_root 'ambiguous-source-set'
    $ambiguous_source_set = New-TestProject -ProjectRoot (Join-Path $ambiguous_search_root 'Project') -DocumentName 'Book'
    Write-TestFile -Path (Join-Path $ambiguous_source_set 'common-modules-manifest.tsv') -Content $manifest_text -Encoding $utf16_le
    $ambiguous_direct_module = Join-Path $ambiguous_source_set 'Shared.bas'
    $ambiguous_nested_module = Join-Path $ambiguous_source_set 'nested\Shared.bas'
    Write-TestFile -Path $ambiguous_direct_module -Content "Attribute VB_Name = `"Shared`"`r`n'direct`r`n" -Encoding $shift_jis
    Write-TestFile -Path $ambiguous_nested_module -Content "Attribute VB_Name = `"Shared`"`r`n'nested`r`n" -Encoding $shift_jis
    [System.IO.File]::SetLastWriteTimeUtc($ambiguous_direct_module, [datetime]'2025-01-01T00:00:00Z')
    [System.IO.File]::SetLastWriteTimeUtc($ambiguous_nested_module, [datetime]'2025-01-02T00:00:00Z')
    $ambiguous_working_directory = Join-Path $temp_root 'ambiguous-working-directory'
    $ambiguous_output = Join-Path $ambiguous_working_directory 'common_modules_repo'
    New-Item -ItemType Directory -Path $ambiguous_output -Force | Out-Null
    Write-TestFile -Path (Join-Path $ambiguous_output 'sentinel.txt') -Content 'preserve before ambiguity preflight' -Encoding (New-Object System.Text.UTF8Encoding($false))
    $ambiguous_result = Invoke-Collect -WorkingDirectory $ambiguous_working_directory -Arguments @($ambiguous_search_root)
    Assert-CollectFailure -Result $ambiguous_result -ExpectedMessage 'ambiguous candidate'
    Assert-True -Condition (Test-Path -LiteralPath (Join-Path $ambiguous_output 'sentinel.txt') -PathType Leaf) -Message 'COLLECT mutated output before rejecting an ambiguous source-set basename.'

    $missing_authoring_search_root = Join-Path $temp_root 'missing-authoring-candidate'
    $missing_authoring_source_set = New-TestProject -ProjectRoot (Join-Path $missing_authoring_search_root 'AuthorProject') -DocumentName 'AuthorBook'
    Write-TestFile -Path (Join-Path $missing_authoring_source_set 'common-modules-manifest.tsv') -Content $manifest_text -Encoding $utf16_le
    $fallback_only_source_set = New-TestProject -ProjectRoot (Join-Path $missing_authoring_search_root 'OtherProject') -DocumentName 'OtherBook'
    Write-TestFile -Path (Join-Path $fallback_only_source_set 'Shared.bas') -Content "Attribute VB_Name = `"Shared`"`r`n'other project only`r`n" -Encoding $shift_jis
    $missing_authoring_working_directory = Join-Path $temp_root 'missing-authoring-working-directory'
    $missing_authoring_output = Join-Path $missing_authoring_working_directory 'common_modules_repo'
    New-Item -ItemType Directory -Path $missing_authoring_output -Force | Out-Null
    Write-TestFile -Path (Join-Path $missing_authoring_output 'sentinel.txt') -Content 'preserve before mandatory-candidate preflight' -Encoding (New-Object System.Text.UTF8Encoding($false))
    $missing_authoring_result = Invoke-Collect -WorkingDirectory $missing_authoring_working_directory -Arguments @($missing_authoring_search_root)
    Assert-CollectFailure -Result $missing_authoring_result -ExpectedMessage 'Authoring Source Set must provide'
    Assert-True -Condition (Test-Path -LiteralPath (Join-Path $missing_authoring_output 'sentinel.txt') -PathType Leaf) -Message 'COLLECT mutated output before rejecting a missing mandatory authoring candidate.'

    $conflicting_tie_search_root = Join-Path $temp_root 'conflicting-size-tie'
    $tie_authoring_source_set = New-TestProject -ProjectRoot (Join-Path $conflicting_tie_search_root 'AuthorProject') -DocumentName 'AuthorBook'
    $tie_other_a_source_set = New-TestProject -ProjectRoot (Join-Path $conflicting_tie_search_root 'OtherA') -DocumentName 'OtherBookA'
    $tie_other_b_source_set = New-TestProject -ProjectRoot (Join-Path $conflicting_tie_search_root 'OtherB') -DocumentName 'OtherBookB'
    Write-TestFile -Path (Join-Path $tie_authoring_source_set 'common-modules-manifest.tsv') -Content $manifest_text -Encoding $utf16_le
    $tie_authoring_module = Join-Path $tie_authoring_source_set 'Shared.bas'
    $tie_other_a_module = Join-Path $tie_other_a_source_set 'Shared.bas'
    $tie_other_b_module = Join-Path $tie_other_b_source_set 'Shared.bas'
    Write-TestFile -Path $tie_authoring_module -Content "Attribute VB_Name = `"Shared`"`r`n'authoring fallback`r`n" -Encoding $shift_jis
    Write-TestFile -Path $tie_other_a_module -Content "Attribute VB_Name = `"Shared`"`r`n'A`r`n" -Encoding $shift_jis
    Write-TestFile -Path $tie_other_b_module -Content "Attribute VB_Name = `"Shared`"`r`n'BBBBBBBB`r`n" -Encoding $shift_jis
    [System.IO.File]::SetLastWriteTimeUtc($tie_authoring_module, [datetime]'2025-01-01T00:00:00Z')
    [System.IO.File]::SetLastWriteTimeUtc($tie_other_a_module, [datetime]'2025-01-03T00:00:00Z')
    [System.IO.File]::SetLastWriteTimeUtc($tie_other_b_module, [datetime]'2025-01-03T00:00:00Z')
    $conflicting_tie_working_directory = Join-Path $temp_root 'conflicting-size-tie-working-directory'
    New-Item -ItemType Directory -Path $conflicting_tie_working_directory -Force | Out-Null
    $conflicting_tie_result = Invoke-Collect -WorkingDirectory $conflicting_tie_working_directory -Arguments @($conflicting_tie_search_root)
    Assert-Equal -Expected 0 -Actual $conflicting_tie_result.ExitCode -Message "COLLECT should warn and use the authoring fallback for a conflicting newest-size tie. Output: $($conflicting_tie_result.Output)"
    Assert-True -Condition ($conflicting_tie_result.Output -like '*WARNING*Shared.bas*') -Message "COLLECT did not warn about a conflicting newest-size tie. Output: $($conflicting_tie_result.Output)"
    Assert-True -Condition ((Get-Content -LiteralPath (Join-Path $conflicting_tie_working_directory 'common_modules_repo\Shared.bas') -Raw) -like '*authoring fallback*') -Message 'COLLECT did not select the authoring fallback for a conflicting newest-size tie.'

    $author_tie_search_root = Join-Path $temp_root 'authoring-equivalent-tie'
    $author_tie_authoring_source_set = New-TestProject -ProjectRoot (Join-Path $author_tie_search_root 'AuthorProject') -DocumentName 'AuthorBook'
    $author_tie_other_source_set = New-TestProject -ProjectRoot (Join-Path $author_tie_search_root 'OtherProject') -DocumentName 'OtherBook'
    Write-TestFile -Path (Join-Path $author_tie_authoring_source_set 'common-modules-manifest.tsv') -Content $manifest_text -Encoding $utf16_le
    $author_tie_authoring_module = Join-Path $author_tie_authoring_source_set 'Shared.bas'
    $author_tie_other_module = Join-Path $author_tie_other_source_set 'Shared.bas'
    Write-TestFile -Path $author_tie_authoring_module -Content "Attribute VB_Name = `"Shared`"`r`n'AAAA`r`n" -Encoding $shift_jis
    Write-TestFile -Path $author_tie_other_module -Content "Attribute VB_Name = `"Shared`"`r`n'BBBB`r`n" -Encoding $shift_jis
    [System.IO.File]::SetLastWriteTimeUtc($author_tie_authoring_module, [datetime]'2025-02-01T00:00:00Z')
    [System.IO.File]::SetLastWriteTimeUtc($author_tie_other_module, [datetime]'2025-02-01T00:00:00Z')
    $author_tie_working_directory = Join-Path $temp_root 'author-tie-working-directory'
    New-Item -ItemType Directory -Path $author_tie_working_directory -Force | Out-Null
    $author_tie_lock = [System.IO.File]::Open($author_tie_other_module, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
    try {
        $author_tie_result = Invoke-Collect -WorkingDirectory $author_tie_working_directory -Arguments @($author_tie_search_root)
    }
    finally {
        $author_tie_lock.Dispose()
    }
    Assert-Equal -Expected 0 -Actual $author_tie_result.ExitCode -Message "COLLECT read equivalent candidate contents instead of preferring the authoring candidate by metadata. Output: $($author_tie_result.Output)"
    Assert-True -Condition ((Get-Content -LiteralPath (Join-Path $author_tie_working_directory 'common_modules_repo\Shared.bas') -Raw) -like '*AAAA*') -Message 'COLLECT did not prefer the authoring candidate in an equivalent newest metadata tie.'

    $ordinal_tie_search_root = Join-Path $temp_root 'ordinal-equivalent-tie'
    $ordinal_tie_authoring_source_set = New-TestProject -ProjectRoot (Join-Path $ordinal_tie_search_root 'AuthorProject') -DocumentName 'AuthorBook'
    $ordinal_tie_a_source_set = New-TestProject -ProjectRoot (Join-Path $ordinal_tie_search_root 'AProject') -DocumentName 'ABook'
    $ordinal_tie_b_source_set = New-TestProject -ProjectRoot (Join-Path $ordinal_tie_search_root 'BProject') -DocumentName 'BBook'
    Write-TestFile -Path (Join-Path $ordinal_tie_authoring_source_set 'common-modules-manifest.tsv') -Content $manifest_text -Encoding $utf16_le
    $ordinal_tie_authoring_module = Join-Path $ordinal_tie_authoring_source_set 'Shared.bas'
    $ordinal_tie_a_module = Join-Path $ordinal_tie_a_source_set 'Shared.bas'
    $ordinal_tie_b_module = Join-Path $ordinal_tie_b_source_set 'Shared.bas'
    Write-TestFile -Path $ordinal_tie_authoring_module -Content "Attribute VB_Name = `"Shared`"`r`n'old!`r`n" -Encoding $shift_jis
    Write-TestFile -Path $ordinal_tie_a_module -Content "Attribute VB_Name = `"Shared`"`r`n'AAAA`r`n" -Encoding $shift_jis
    Write-TestFile -Path $ordinal_tie_b_module -Content "Attribute VB_Name = `"Shared`"`r`n'BBBB`r`n" -Encoding $shift_jis
    [System.IO.File]::SetLastWriteTimeUtc($ordinal_tie_authoring_module, [datetime]'2025-02-01T00:00:00Z')
    [System.IO.File]::SetLastWriteTimeUtc($ordinal_tie_a_module, [datetime]'2025-02-03T00:00:00Z')
    [System.IO.File]::SetLastWriteTimeUtc($ordinal_tie_b_module, [datetime]'2025-02-03T00:00:00Z')
    $ordinal_tie_working_directory = Join-Path $temp_root 'ordinal-tie-working-directory'
    New-Item -ItemType Directory -Path $ordinal_tie_working_directory -Force | Out-Null
    $ordinal_tie_lock = [System.IO.File]::Open($ordinal_tie_b_module, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
    try {
        $ordinal_tie_result = Invoke-Collect -WorkingDirectory $ordinal_tie_working_directory -Arguments @($ordinal_tie_search_root)
    }
    finally {
        $ordinal_tie_lock.Dispose()
    }
    Assert-Equal -Expected 0 -Actual $ordinal_tie_result.ExitCode -Message "COLLECT did not select an ordinal-minimum path from an equivalent newest metadata tie. Output: $($ordinal_tie_result.Output)"
    Assert-True -Condition ((Get-Content -LiteralPath (Join-Path $ordinal_tie_working_directory 'common_modules_repo\Shared.bas') -Raw) -like '*AAAA*') -Message 'COLLECT did not select the ordinal-minimum normalized source path.'

    $tick_precision_search_root = Join-Path $temp_root 'tick-precision-selection'
    $tick_precision_authoring_source_set = New-TestProject -ProjectRoot (Join-Path $tick_precision_search_root 'AuthorProject') -DocumentName 'AuthorBook'
    $tick_precision_earlier_source_set = New-TestProject -ProjectRoot (Join-Path $tick_precision_search_root 'EarlierProject') -DocumentName 'EarlierBook'
    $tick_precision_later_source_set = New-TestProject -ProjectRoot (Join-Path $tick_precision_search_root 'LaterProject') -DocumentName 'LaterBook'
    Write-TestFile -Path (Join-Path $tick_precision_authoring_source_set 'common-modules-manifest.tsv') -Content $manifest_text -Encoding $utf16_le
    $tick_precision_authoring_module = Join-Path $tick_precision_authoring_source_set 'Shared.bas'
    $tick_precision_earlier_module = Join-Path $tick_precision_earlier_source_set 'Shared.bas'
    $tick_precision_later_module = Join-Path $tick_precision_later_source_set 'Shared.bas'
    Write-TestFile -Path $tick_precision_authoring_module -Content "Attribute VB_Name = `"Shared`"`r`n'authoring fallback`r`n" -Encoding $shift_jis
    Write-TestFile -Path $tick_precision_earlier_module -Content "Attribute VB_Name = `"Shared`"`r`n'earlier and longer`r`n" -Encoding $shift_jis
    Write-TestFile -Path $tick_precision_later_module -Content "Attribute VB_Name = `"Shared`"`r`n'later`r`n" -Encoding $shift_jis
    $tick_precision_earlier_time = [datetime]::SpecifyKind([datetime]'2025-03-01T00:00:00', [System.DateTimeKind]::Utc)
    $tick_precision_later_time = $tick_precision_earlier_time.AddTicks(1)
    [System.IO.File]::SetLastWriteTimeUtc($tick_precision_authoring_module, [datetime]'2025-02-01T00:00:00Z')
    [System.IO.File]::SetLastWriteTimeUtc($tick_precision_earlier_module, $tick_precision_earlier_time)
    [System.IO.File]::SetLastWriteTimeUtc($tick_precision_later_module, $tick_precision_later_time)
    Assert-Equal -Expected $tick_precision_later_time.Ticks -Actual ([System.IO.File]::GetLastWriteTimeUtc($tick_precision_later_module).Ticks) -Message 'The test filesystem did not preserve the required one-tick timestamp difference.'
    $tick_precision_working_directory = Join-Path $temp_root 'tick-precision-working-directory'
    New-Item -ItemType Directory -Path $tick_precision_working_directory -Force | Out-Null
    $tick_precision_result = Invoke-Collect -WorkingDirectory $tick_precision_working_directory -Arguments @($tick_precision_search_root)
    Assert-Equal -Expected 0 -Actual $tick_precision_result.ExitCode -Message "COLLECT lost LastWriteTimeUtc tick precision. Output: $($tick_precision_result.Output)"
    Assert-True -Condition ($tick_precision_result.Output -notlike '*WARNING*') -Message "COLLECT misclassified distinct one-tick timestamps as a conflicting tie. Output: $($tick_precision_result.Output)"
    Assert-True -Condition ((Get-Content -LiteralPath (Join-Path $tick_precision_working_directory 'common_modules_repo\Shared.bas') -Raw) -like '*later*') -Message 'COLLECT did not select the candidate that was newer by one LastWriteTimeUtc tick.'

    $reparse_candidate_search_root = Join-Path $temp_root 'reparse-fallback-candidate'
    $reparse_candidate_authoring_source_set = New-TestProject -ProjectRoot (Join-Path $reparse_candidate_search_root 'AuthorProject') -DocumentName 'AuthorBook'
    $reparse_candidate_other_source_set = New-TestProject -ProjectRoot (Join-Path $reparse_candidate_search_root 'OtherProject') -DocumentName 'OtherBook'
    Write-TestFile -Path (Join-Path $reparse_candidate_authoring_source_set 'common-modules-manifest.tsv') -Content $manifest_text -Encoding $utf16_le
    Write-TestFile -Path (Join-Path $reparse_candidate_authoring_source_set 'Shared.bas') -Content "Attribute VB_Name = `"Shared`"`r`n'ordinary authoring`r`n" -Encoding $shift_jis
    $reparse_candidate_target = Join-Path $temp_root 'reparse-candidate-target.bas'
    Write-TestFile -Path $reparse_candidate_target -Content "Attribute VB_Name = `"Shared`"`r`n'reparse target`r`n" -Encoding $shift_jis
    New-TestFileSymbolicLink -Path (Join-Path $reparse_candidate_other_source_set 'Shared.bas') -Target $reparse_candidate_target
    $reparse_candidate_working_directory = Join-Path $temp_root 'reparse-candidate-working-directory'
    New-Item -ItemType Directory -Path $reparse_candidate_working_directory -Force | Out-Null
    $reparse_candidate_result = Invoke-Collect -WorkingDirectory $reparse_candidate_working_directory -Arguments @($reparse_candidate_search_root)
    Assert-Equal -Expected 0 -Actual $reparse_candidate_result.ExitCode -Message "COLLECT failed instead of warning and using the authoring fallback for a non-authoring reparse candidate. Output: $($reparse_candidate_result.Output)"
    Assert-True -Condition ($reparse_candidate_result.Output -like '*WARNING*ordinary non-reparse file*') -Message "COLLECT did not warn about a non-authoring reparse candidate. Output: $($reparse_candidate_result.Output)"
    Assert-True -Condition ((Get-Content -LiteralPath (Join-Path $reparse_candidate_working_directory 'common_modules_repo\Shared.bas') -Raw) -like '*ordinary authoring*') -Message 'COLLECT did not use the authoring fallback for a non-authoring reparse candidate.'

    $reparse_authoring_search_root = Join-Path $temp_root 'reparse-authoring-candidate'
    $reparse_authoring_source_set = New-TestProject -ProjectRoot (Join-Path $reparse_authoring_search_root 'AuthorProject') -DocumentName 'AuthorBook'
    Write-TestFile -Path (Join-Path $reparse_authoring_source_set 'common-modules-manifest.tsv') -Content $manifest_text -Encoding $utf16_le
    $reparse_authoring_target = Join-Path $temp_root 'reparse-authoring-target.bas'
    Write-TestFile -Path $reparse_authoring_target -Content "Attribute VB_Name = `"Shared`"`r`n'reparse authoring target`r`n" -Encoding $shift_jis
    New-TestFileSymbolicLink -Path (Join-Path $reparse_authoring_source_set 'Shared.bas') -Target $reparse_authoring_target
    $reparse_authoring_working_directory = Join-Path $temp_root 'reparse-authoring-working-directory'
    $reparse_authoring_output = Join-Path $reparse_authoring_working_directory 'common_modules_repo'
    New-Item -ItemType Directory -Path $reparse_authoring_output -Force | Out-Null
    Write-TestFile -Path (Join-Path $reparse_authoring_output 'sentinel.txt') -Content 'preserve' -Encoding $shift_jis
    $reparse_authoring_result = Invoke-Collect -WorkingDirectory $reparse_authoring_working_directory -Arguments @($reparse_authoring_search_root)
    Assert-CollectFailure -Result $reparse_authoring_result -ExpectedMessage 'Authoring Source Set candidate'
    Assert-True -Condition (Test-Path -LiteralPath (Join-Path $reparse_authoring_output 'sentinel.txt') -PathType Leaf) -Message 'COLLECT mutated output before rejecting a reparse authoring candidate.'

    $form_search_root = Join-Path $temp_root 'form-sidecar-selection'
    $form_authoring_source_set = New-TestProject -ProjectRoot (Join-Path $form_search_root 'AuthorProject') -DocumentName 'AuthorBook'
    $form_pair_source_set = New-TestProject -ProjectRoot (Join-Path $form_search_root 'PairProject') -DocumentName 'PairBook'
    $form_solo_source_set = New-TestProject -ProjectRoot (Join-Path $form_search_root 'SoloProject') -DocumentName 'SoloBook'
    $form_manifest_text = "ModuleFile`tCategories`tDependencies`tRequiredReferences`r`nDialog.frm`toptional`t`t[]`r`n"
    Write-TestFile -Path (Join-Path $form_authoring_source_set 'common-modules-manifest.tsv') -Content $form_manifest_text -Encoding $utf16_le
    $form_authoring_file = Join-Path $form_authoring_source_set 'Dialog.frm'
    $form_pair_file = Join-Path $form_pair_source_set 'Dialog.frm'
    $form_pair_sidecar = Join-Path $form_pair_source_set 'Dialog.frx'
    $form_solo_file = Join-Path $form_solo_source_set 'Dialog.frm'
    Write-TestFile -Path $form_authoring_file -Content "VERSION 5.00`r`n'authoring form`r`n" -Encoding $shift_jis
    Write-TestFile -Path $form_pair_file -Content "VERSION 5.00`r`n'paired form`r`n" -Encoding $shift_jis
    Write-TestFile -Path $form_pair_sidecar -Content 'paired sidecar' -Encoding $shift_jis
    Write-TestFile -Path $form_solo_file -Content "VERSION 5.00`r`n'solo form`r`n" -Encoding $shift_jis
    [System.IO.File]::SetLastWriteTimeUtc($form_authoring_file, [datetime]'2025-01-01T00:00:00Z')
    [System.IO.File]::SetLastWriteTimeUtc($form_pair_file, [datetime]'2025-01-02T00:00:00Z')
    [System.IO.File]::SetLastWriteTimeUtc($form_pair_sidecar, [datetime]'2025-01-04T00:00:00Z')
    [System.IO.File]::SetLastWriteTimeUtc($form_solo_file, [datetime]'2025-01-03T00:00:00Z')
    $form_working_directory = Join-Path $temp_root 'form-working-directory'
    New-Item -ItemType Directory -Path $form_working_directory -Force | Out-Null
    $form_result = Invoke-Collect -WorkingDirectory $form_working_directory -Arguments @($form_search_root)
    Assert-Equal -Expected 0 -Actual $form_result.ExitCode -Message "COLLECT failed to select a form pair. Output: $($form_result.Output)"
    $form_output = Join-Path $form_working_directory 'common_modules_repo'
    Assert-True -Condition ((Get-Content -LiteralPath (Join-Path $form_output 'Dialog.frm') -Raw) -like '*paired form*') -Message 'COLLECT did not use the later `.frx` timestamp when selecting a form candidate.'
    Assert-True -Condition (Test-Path -LiteralPath (Join-Path $form_output 'Dialog.frx') -PathType Leaf) -Message 'COLLECT did not copy the selected form sidecar.'
    Assert-Equal -Expected 'paired sidecar' -Actual (Get-Content -LiteralPath (Join-Path $form_output 'Dialog.frx') -Raw) -Message 'COLLECT copied unexpected form sidecar bytes.'

    $form_tie_search_root = Join-Path $temp_root 'form-sidecar-shape-tie'
    $form_tie_authoring_source_set = New-TestProject -ProjectRoot (Join-Path $form_tie_search_root 'AuthorProject') -DocumentName 'AuthorBook'
    $form_tie_solo_source_set = New-TestProject -ProjectRoot (Join-Path $form_tie_search_root 'SoloProject') -DocumentName 'SoloBook'
    $form_tie_pair_source_set = New-TestProject -ProjectRoot (Join-Path $form_tie_search_root 'PairProject') -DocumentName 'PairBook'
    $form_tie_orphan_source_set = New-TestProject -ProjectRoot (Join-Path $form_tie_search_root 'OrphanProject') -DocumentName 'OrphanBook'
    Write-TestFile -Path (Join-Path $form_tie_authoring_source_set 'common-modules-manifest.tsv') -Content $form_manifest_text -Encoding $utf16_le
    $form_tie_authoring_file = Join-Path $form_tie_authoring_source_set 'Dialog.frm'
    $form_tie_authoring_sidecar = Join-Path $form_tie_authoring_source_set 'Dialog.frx'
    $form_tie_solo_file = Join-Path $form_tie_solo_source_set 'Dialog.frm'
    $form_tie_pair_file = Join-Path $form_tie_pair_source_set 'Dialog.frm'
    $form_tie_pair_sidecar = Join-Path $form_tie_pair_source_set 'Dialog.frx'
    $form_tie_orphan_sidecar = Join-Path $form_tie_orphan_source_set 'Dialog.frx'
    Write-TestFile -Path $form_tie_authoring_file -Content "VERSION 5.00`r`n'authoring pair`r`n" -Encoding $shift_jis
    Write-TestFile -Path $form_tie_authoring_sidecar -Content 'authoring sidecar' -Encoding $shift_jis
    Write-TestFile -Path $form_tie_solo_file -Content "VERSION 5.00`r`n'AAAA`r`n" -Encoding $shift_jis
    Write-TestFile -Path $form_tie_pair_file -Content "VERSION 5.00`r`n'BBBB`r`n" -Encoding $shift_jis
    Write-TestFile -Path $form_tie_pair_sidecar -Content 'paired sidecar' -Encoding $shift_jis
    Write-TestFile -Path $form_tie_orphan_sidecar -Content 'orphan sidecar' -Encoding $shift_jis
    [System.IO.File]::SetLastWriteTimeUtc($form_tie_authoring_file, [datetime]'2025-01-01T00:00:00Z')
    [System.IO.File]::SetLastWriteTimeUtc($form_tie_authoring_sidecar, [datetime]'2025-01-01T00:00:00Z')
    [System.IO.File]::SetLastWriteTimeUtc($form_tie_solo_file, [datetime]'2025-01-05T00:00:00Z')
    [System.IO.File]::SetLastWriteTimeUtc($form_tie_pair_file, [datetime]'2025-01-02T00:00:00Z')
    [System.IO.File]::SetLastWriteTimeUtc($form_tie_pair_sidecar, [datetime]'2025-01-05T00:00:00Z')
    [System.IO.File]::SetLastWriteTimeUtc($form_tie_orphan_sidecar, [datetime]'2025-01-06T00:00:00Z')
    $form_tie_working_directory = Join-Path $temp_root 'form-tie-working-directory'
    New-Item -ItemType Directory -Path $form_tie_working_directory -Force | Out-Null
    $form_tie_result = Invoke-Collect -WorkingDirectory $form_tie_working_directory -Arguments @($form_tie_search_root)
    Assert-Equal -Expected 0 -Actual $form_tie_result.ExitCode -Message "COLLECT failed instead of using the authoring form pair for a newest shape tie. Output: $($form_tie_result.Output)"
    Assert-True -Condition ($form_tie_result.Output -like '*WARNING*form sidecar shapes*') -Message "COLLECT did not warn about differing newest form sidecar shapes. Output: $($form_tie_result.Output)"
    Assert-True -Condition ((Get-Content -LiteralPath (Join-Path $form_tie_working_directory 'common_modules_repo\Dialog.frm') -Raw) -like '*authoring pair*') -Message 'COLLECT did not use the authoring form fallback for a sidecar-shape tie.'
    Assert-Equal -Expected 'authoring sidecar' -Actual (Get-Content -LiteralPath (Join-Path $form_tie_working_directory 'common_modules_repo\Dialog.frx') -Raw) -Message 'COLLECT did not preserve the authoring fallback sidecar.'

    $reparse_sidecar_search_root = Join-Path $temp_root 'reparse-sidecar-fallback'
    $reparse_sidecar_authoring_source_set = New-TestProject -ProjectRoot (Join-Path $reparse_sidecar_search_root 'AuthorProject') -DocumentName 'AuthorBook'
    $reparse_sidecar_other_source_set = New-TestProject -ProjectRoot (Join-Path $reparse_sidecar_search_root 'OtherProject') -DocumentName 'OtherBook'
    Write-TestFile -Path (Join-Path $reparse_sidecar_authoring_source_set 'common-modules-manifest.tsv') -Content $form_manifest_text -Encoding $utf16_le
    Write-TestFile -Path (Join-Path $reparse_sidecar_authoring_source_set 'Dialog.frm') -Content "VERSION 5.00`r`n'authoring form fallback`r`n" -Encoding $shift_jis
    Write-TestFile -Path (Join-Path $reparse_sidecar_other_source_set 'Dialog.frm') -Content "VERSION 5.00`r`n'other form`r`n" -Encoding $shift_jis
    $reparse_sidecar_target = Join-Path $temp_root 'reparse-sidecar-target.frx'
    Write-TestFile -Path $reparse_sidecar_target -Content 'reparse sidecar target' -Encoding $shift_jis
    New-TestFileSymbolicLink -Path (Join-Path $reparse_sidecar_other_source_set 'Dialog.frx') -Target $reparse_sidecar_target
    $reparse_sidecar_working_directory = Join-Path $temp_root 'reparse-sidecar-working-directory'
    New-Item -ItemType Directory -Path $reparse_sidecar_working_directory -Force | Out-Null
    $reparse_sidecar_result = Invoke-Collect -WorkingDirectory $reparse_sidecar_working_directory -Arguments @($reparse_sidecar_search_root)
    Assert-Equal -Expected 0 -Actual $reparse_sidecar_result.ExitCode -Message "COLLECT failed instead of using the authoring fallback for a reparse form sidecar. Output: $($reparse_sidecar_result.Output)"
    Assert-True -Condition ($reparse_sidecar_result.Output -like '*WARNING*Form sidecar*ordinary non-reparse file*') -Message "COLLECT did not warn about a non-authoring reparse form sidecar. Output: $($reparse_sidecar_result.Output)"
    Assert-True -Condition (-not (Test-Path -LiteralPath (Join-Path $reparse_sidecar_working_directory 'common_modules_repo\Dialog.frx'))) -Message 'COLLECT copied a reparse fallback sidecar that was not part of the authoring form candidate.'

    $reparse_authoring_sidecar_search_root = Join-Path $temp_root 'reparse-authoring-sidecar'
    $reparse_authoring_sidecar_source_set = New-TestProject -ProjectRoot (Join-Path $reparse_authoring_sidecar_search_root 'AuthorProject') -DocumentName 'AuthorBook'
    Write-TestFile -Path (Join-Path $reparse_authoring_sidecar_source_set 'common-modules-manifest.tsv') -Content $form_manifest_text -Encoding $utf16_le
    Write-TestFile -Path (Join-Path $reparse_authoring_sidecar_source_set 'Dialog.frm') -Content "VERSION 5.00`r`n'authoring form`r`n" -Encoding $shift_jis
    New-TestFileSymbolicLink -Path (Join-Path $reparse_authoring_sidecar_source_set 'Dialog.frx') -Target $reparse_sidecar_target
    $reparse_authoring_sidecar_working_directory = Join-Path $temp_root 'reparse-authoring-sidecar-working-directory'
    $reparse_authoring_sidecar_output = Join-Path $reparse_authoring_sidecar_working_directory 'common_modules_repo'
    New-Item -ItemType Directory -Path $reparse_authoring_sidecar_output -Force | Out-Null
    Write-TestFile -Path (Join-Path $reparse_authoring_sidecar_output 'sentinel.txt') -Content 'preserve' -Encoding $shift_jis
    $reparse_authoring_sidecar_result = Invoke-Collect -WorkingDirectory $reparse_authoring_sidecar_working_directory -Arguments @($reparse_authoring_sidecar_search_root)
    Assert-CollectFailure -Result $reparse_authoring_sidecar_result -ExpectedMessage 'Authoring Source Set candidate'
    Assert-True -Condition (Test-Path -LiteralPath (Join-Path $reparse_authoring_sidecar_output 'sentinel.txt') -PathType Leaf) -Message 'COLLECT mutated output before rejecting a reparse authoring form sidecar.'

    $file_output_working_directory = Join-Path $temp_root 'file-output-working-directory'
    New-Item -ItemType Directory -Path $file_output_working_directory -Force | Out-Null
    $file_output_path = Join-Path $file_output_working_directory 'common_modules_repo'
    Write-TestFile -Path $file_output_path -Content 'preserve file output' -Encoding $shift_jis
    $file_output_result = Invoke-Collect -WorkingDirectory $file_output_working_directory -Arguments @($search_root)
    Assert-CollectFailure -Result $file_output_result -ExpectedMessage 'ordinary non-reparse common_modules_repo directory'
    Assert-Equal -Expected 'preserve file output' -Actual (Get-Content -LiteralPath $file_output_path -Raw) -Message 'COLLECT mutated a file occupying the output repository name.'

    $case_output_working_directory = Join-Path $temp_root 'case-output-working-directory'
    $case_output_path = Join-Path $case_output_working_directory 'Common_Modules_Repo'
    New-Item -ItemType Directory -Path $case_output_path -Force | Out-Null
    Write-TestFile -Path (Join-Path $case_output_path 'sentinel.txt') -Content 'preserve case output' -Encoding $shift_jis
    $case_output_result = Invoke-Collect -WorkingDirectory $case_output_working_directory -Arguments @($search_root)
    Assert-CollectFailure -Result $case_output_result -ExpectedMessage 'ordinal-exact ordinary non-reparse common_modules_repo directory'
    Assert-True -Condition (Test-Path -LiteralPath (Join-Path $case_output_path 'sentinel.txt') -PathType Leaf) -Message 'COLLECT mutated a case-only output repository.'

    $reparse_output_working_directory = Join-Path $temp_root 'reparse-output-working-directory'
    $reparse_output_target = Join-Path $temp_root 'reparse-output-target'
    New-Item -ItemType Directory -Path $reparse_output_working_directory, $reparse_output_target -Force | Out-Null
    Write-TestFile -Path (Join-Path $reparse_output_target 'target-sentinel.txt') -Content 'preserve reparse output target' -Encoding $shift_jis
    New-TestJunction -Path (Join-Path $reparse_output_working_directory 'common_modules_repo') -Target $reparse_output_target
    $reparse_output_result = Invoke-Collect -WorkingDirectory $reparse_output_working_directory -Arguments @($search_root)
    Assert-CollectFailure -Result $reparse_output_result -ExpectedMessage 'ordinary non-reparse common_modules_repo directory'
    Assert-True -Condition (Test-Path -LiteralPath (Join-Path $reparse_output_target 'target-sentinel.txt') -PathType Leaf) -Message 'COLLECT followed or mutated a reparse output repository.'

    $partial_copy_search_root = Join-Path $temp_root 'partial-copy-failure'
    $partial_copy_source_set = New-TestProject -ProjectRoot (Join-Path $partial_copy_search_root 'AuthorProject') -DocumentName 'AuthorBook'
    Write-TestFile -Path (Join-Path $partial_copy_source_set 'common-modules-manifest.tsv') -Content $manifest_text -Encoding $utf16_le
    $partial_copy_module = Join-Path $partial_copy_source_set 'Shared.bas'
    Write-TestFile -Path $partial_copy_module -Content "Attribute VB_Name = `"Shared`"`r`n'locked copy source`r`n" -Encoding $shift_jis
    $partial_copy_working_directory = Join-Path $temp_root 'partial-copy-working-directory'
    $partial_copy_output = Join-Path $partial_copy_working_directory 'common_modules_repo'
    New-Item -ItemType Directory -Path $partial_copy_output -Force | Out-Null
    Write-TestFile -Path (Join-Path $partial_copy_output 'sentinel.txt') -Content 'remove before copy' -Encoding $shift_jis
    $partial_copy_lock = [System.IO.File]::Open($partial_copy_module, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
    try {
        $partial_copy_result = Invoke-Collect -WorkingDirectory $partial_copy_working_directory -Arguments @($partial_copy_search_root)
    }
    finally {
        $partial_copy_lock.Dispose()
    }
    Assert-Equal -Expected 1 -Actual $partial_copy_result.ExitCode -Message "COLLECT did not fail when Copy-Item could not read a selected source. Output: $($partial_copy_result.Output)"
    Assert-True -Condition (-not (Test-Path -LiteralPath (Join-Path $partial_copy_output 'sentinel.txt'))) -Message 'COLLECT did not clear the existing output before the selected-source copy failure.'
    Assert-True -Condition (Test-Path -LiteralPath (Join-Path $partial_copy_output 'common-modules-manifest.tsv') -PathType Leaf) -Message 'COLLECT did not leave the sequentially copied manifest in the permitted partial output.'
    Assert-True -Condition (-not (Test-Path -LiteralPath (Join-Path $partial_copy_output 'Shared.bas'))) -Message 'COLLECT unexpectedly copied a source held with FileShare.None.'

    $bat_fixture = Join-Path $temp_root 'bat-fixture'
    $bat_search_root = Join-Path $bat_fixture 'search root'
    $bat_bin = Join-Path $bat_fixture 'bin'
    New-Item -ItemType Directory -Path $bat_fixture, $bat_search_root, $bat_bin -Force | Out-Null
    $test_bat_path = Join-Path $bat_fixture 'COLLECT_COMMON_MODS.BAT'
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'COLLECT_COMMON_MODS.BAT') -Destination $test_bat_path -Force
    $bat_stub_script = @'
[System.IO.File]::WriteAllLines($env:COLLECT_TEST_ARGS_LOG, [string[]]$args, (New-Object System.Text.UTF8Encoding($false)))
exit 37
'@
    Write-TestFile -Path (Join-Path $bat_fixture 'collect_common_mods_main.ps1') -Content $bat_stub_script -Encoding (New-Object System.Text.UTF8Encoding($true))
    $timeout_stub = "@ECHO off`r`nECHO timeout>>`"%COLLECT_TEST_TIMEOUT_LOG%`"`r`nEXIT /B 0`r`n"
    Write-TestFile -Path (Join-Path $bat_bin 'TIMEOUT.CMD') -Content $timeout_stub -Encoding $shift_jis
    $bat_args_log = Join-Path $bat_fixture 'args.log'
    $bat_timeout_log = Join-Path $bat_fixture 'timeout.log'
    $previous_collect_test_args_log = $env:COLLECT_TEST_ARGS_LOG
    $previous_collect_test_timeout_log = $env:COLLECT_TEST_TIMEOUT_LOG
    $previous_test_path = $env:PATH
    $previous_pathext = $env:PATHEXT
    try {
        $env:COLLECT_TEST_ARGS_LOG = $bat_args_log
        $env:COLLECT_TEST_TIMEOUT_LOG = $bat_timeout_log
        $env:PATH = "$bat_bin;$env:PATH"
        $env:PATHEXT = '.CMD;.EXE;.COM;.BAT'
        $bat_command = "`"$test_bat_path`" `"$bat_search_root`""
        $bat_output = & $env:ComSpec /d /c $bat_command 2>&1
        $bat_exit_code = $LASTEXITCODE
    }
    finally {
        $env:COLLECT_TEST_ARGS_LOG = $previous_collect_test_args_log
        $env:COLLECT_TEST_TIMEOUT_LOG = $previous_collect_test_timeout_log
        $env:PATH = $previous_test_path
        $env:PATHEXT = $previous_pathext
    }
    Assert-Equal -Expected 37 -Actual $bat_exit_code -Message "COLLECT_COMMON_MODS.BAT did not preserve the PowerShell exit code. Output: $($bat_output | Out-String)"
    $bat_arguments = @(Get-Content -LiteralPath $bat_args_log)
    Assert-Equal -Expected 1 -Actual $bat_arguments.Count -Message 'COLLECT_COMMON_MODS.BAT did not pass exactly one public argument to PowerShell.'
    Assert-Equal -Expected $bat_search_root -Actual $bat_arguments[0] -Message 'COLLECT_COMMON_MODS.BAT changed the Collection Search Root argument.'
    Assert-True -Condition (Test-Path -LiteralPath $bat_timeout_log -PathType Leaf) -Message 'BAT exit preservation test did not execute the informational TIMEOUT stub.'
    $bat_text = [System.IO.File]::ReadAllText((Join-Path $PSScriptRoot 'COLLECT_COMMON_MODS.BAT'), $shift_jis)
    Assert-True -Condition ($bat_text -match '(?im)^powershell\s+-NoProfile\b') -Message 'COLLECT_COMMON_MODS.BAT does not invoke Windows PowerShell with -NoProfile.'

    Write-Host 'COLLECT CommonModules tests passed.'
}
finally {
    if (Test-Path -LiteralPath $temp_root) {
        Remove-Item -LiteralPath $temp_root -Recurse -Force
    }
}
