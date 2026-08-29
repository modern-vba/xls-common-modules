$ErrorActionPreference = 'Stop'
$InformationPreference = 'Continue'

$dist_script = Join-Path $PSScriptRoot 'dist_common_mods_repo_main.ps1'
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

function Assert-Like {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Actual,

        [Parameter(Mandatory = $true)]
        [string]$Pattern,

        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    if ($Actual -notlike $Pattern) {
        throw "$Message Expected output matching '$Pattern' but got:`r`n$Actual"
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

function Invoke-Dist {
    param(
        [Parameter(Mandatory = $true)]
        [string]$WorkingDirectory,

        [string[]]$Arguments = @()
    )

    Push-Location -LiteralPath $WorkingDirectory
    $previous_error_action_preference = $ErrorActionPreference
    try {
        $ErrorActionPreference = 'Continue'
        $output = & $test_host_executable -NoProfile -ExecutionPolicy Bypass -File $dist_script @Arguments 2>&1
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

function Write-ValidPackage {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepositoryPath,

        [string]$ModuleName = 'Shared.bas',

        [string]$ModuleContent = "Attribute VB_Name = `"Shared`"`r`n'package source`r`n"
    )

    New-Item -ItemType Directory -Path $RepositoryPath -Force | Out-Null
    $manifest_text = "ModuleFile`tCategories`tDependencies`tRequiredReferences`r`n$ModuleName`truntime-baseline`t`t[]`r`n"
    Write-TestFile -Path (Join-Path $RepositoryPath 'common-modules-manifest.tsv') -Content $manifest_text -Encoding (New-Object System.Text.UnicodeEncoding($false, $true, $true))
    Write-TestFile -Path (Join-Path $RepositoryPath $ModuleName) -Content $ModuleContent -Encoding ([System.Text.Encoding]::GetEncoding(932))
}

function New-DistFixture {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FixtureRoot
    )

    $working_directory = Join-Path $FixtureRoot 'wrapper'
    $source_repository = Join-Path $working_directory 'common_modules_repo'
    $search_root = Join-Path $FixtureRoot 'search-root'
    $target_repository = Join-Path $search_root 'ProjectA\common_modules_repo'
    Write-ValidPackage -RepositoryPath $source_repository
    New-Item -ItemType Directory -Path $target_repository -Force | Out-Null
    Write-TestFile -Path (Join-Path $target_repository 'Sentinel.bas') -Content "Attribute VB_Name = `"Sentinel`"`r`n" -Encoding ([System.Text.Encoding]::GetEncoding(932))

    return [pscustomobject]@{
        WorkingDirectory = $working_directory
        SourceRepository = $source_repository
        SearchRoot = $search_root
        TargetRepository = $target_repository
    }
}

function Copy-PackageFiles {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourceRepository,

        [Parameter(Mandatory = $true)]
        [string]$TargetRepository
    )

    New-Item -ItemType Directory -Path $TargetRepository -Force | Out-Null
    foreach ($file in Get-ChildItem -LiteralPath $SourceRepository -File -Force) {
        Copy-Item -LiteralPath $file.FullName -Destination (Join-Path $TargetRepository $file.Name) -Force
    }
}

function Invoke-ExpectedDistFailure {
    param(
        [Parameter(Mandatory = $true)]
        [string]$WorkingDirectory,

        [string[]]$Arguments = @(),

        [Parameter(Mandatory = $true)]
        [string]$ExpectedPattern,

        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    $result = Invoke-Dist -WorkingDirectory $WorkingDirectory -Arguments $Arguments
    Assert-Equal -Expected 1 -Actual $result.ExitCode -Message "$Message Output: $($result.Output)"
    Assert-Like -Actual $result.Output -Pattern $ExpectedPattern -Message $Message
    return $result
}

function Write-LargeTestPackage {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepositoryPath,

        [Parameter(Mandatory = $true)]
        [string]$TailModuleName,

        [int]$IntermediateCount = 256
    )

    New-Item -ItemType Directory -Path $RepositoryPath -Force | Out-Null
    $module_names = New-Object 'System.Collections.Generic.List[string]'
    $module_names.Add('A_Trigger.bas')
    for ($index = 0; $index -lt $IntermediateCount; $index++) {
        $module_names.Add(('M{0:D3}.bas' -f $index))
    }
    $module_names.Add($TailModuleName)

    $manifest_lines = New-Object 'System.Collections.Generic.List[string]'
    $manifest_lines.Add("ModuleFile`tCategories`tDependencies`tRequiredReferences")
    foreach ($module_name in $module_names) {
        $manifest_lines.Add("$module_name`truntime-baseline`t`t[]")
        $vba_name = [System.IO.Path]::GetFileNameWithoutExtension($module_name)
        Write-TestFile -Path (Join-Path $RepositoryPath $module_name) -Content "Attribute VB_Name = `"$vba_name`"`r`n'$module_name`r`n" -Encoding ([System.Text.Encoding]::GetEncoding(932))
    }
    $manifest_text = ($manifest_lines -join "`r`n") + "`r`n"
    Write-TestFile -Path (Join-Path $RepositoryPath 'common-modules-manifest.tsv') -Content $manifest_text -Encoding (New-Object System.Text.UnicodeEncoding($false, $true, $true))
}

function Start-TestHelperProcess {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ScriptPath,

        [Parameter(Mandatory = $true)]
        [string[]]$Arguments
    )

    $process_arguments = New-Object 'System.Collections.Generic.List[string]'
    $process_arguments.Add('-NoProfile')
    $process_arguments.Add('-ExecutionPolicy')
    $process_arguments.Add('Bypass')
    $process_arguments.Add('-File')
    $process_arguments.Add(('"{0}"' -f $ScriptPath))
    foreach ($argument in $Arguments) {
        $process_arguments.Add(('"{0}"' -f $argument))
    }

    return Start-Process -FilePath $test_host_executable -ArgumentList $process_arguments.ToArray() -PassThru -WindowStyle Hidden
}

function Wait-TestHelperProcess {
    param(
        [Parameter(Mandatory = $true)]
        [System.Diagnostics.Process]$Process,

        [Parameter(Mandatory = $true)]
        [string]$Description
    )

    if (-not $Process.WaitForExit(10000)) {
        $Process.Kill()
        throw "$Description did not finish within 10 seconds."
    }
    Assert-Equal -Expected 0 -Actual $Process.ExitCode -Message "$Description failed."
    $Process.Dispose()
}

$temp_root = Join-Path ([System.IO.Path]::GetTempPath()) ('dist-common-modules-test-' + [System.Guid]::NewGuid().ToString('N'))
$utf16_le = New-Object System.Text.UnicodeEncoding($false, $true, $true)
$shift_jis = [System.Text.Encoding]::GetEncoding(932)

try {
    $working_directory = Join-Path $temp_root 'invocation-working-directory'
    $source_repository = Join-Path $working_directory 'common_modules_repo'
    $search_root = Join-Path $temp_root 'distribution-search-root'
    $target_repository = Join-Path $search_root 'ProjectA\common_modules_repo'
    $nested_repository = Join-Path $search_root 'Container\NestedProject\common_modules_repo'
    $non_target_project = Join-Path $search_root 'ProjectWithoutRepository'
    New-Item -ItemType Directory -Path $source_repository, $target_repository, $nested_repository, $non_target_project -Force | Out-Null

    $manifest_text = "ModuleFile`tCategories`tDependencies`tRequiredReferences`r`nShared.bas`truntime-baseline`t`t[]`r`n"
    Write-TestFile -Path (Join-Path $source_repository 'common-modules-manifest.tsv') -Content $manifest_text -Encoding $utf16_le
    Write-TestFile -Path (Join-Path $source_repository 'Shared.bas') -Content "Attribute VB_Name = `"Shared`"`r`n'distributed source`r`n" -Encoding $shift_jis
    Write-TestFile -Path (Join-Path $target_repository 'Stale.bas') -Content "Attribute VB_Name = `"Stale`"`r`n" -Encoding $shift_jis
    Write-TestFile -Path (Join-Path $nested_repository 'NestedSentinel.bas') -Content "Attribute VB_Name = `"NestedSentinel`"`r`n" -Encoding $shift_jis

    $absolute_result = Invoke-Dist -WorkingDirectory $working_directory -Arguments @($search_root)
    Assert-Equal -Expected 0 -Actual $absolute_result.ExitCode -Message "DIST failed with an absolute independent Distribution Search Root. Output: $($absolute_result.Output)"
    Assert-True -Condition (Test-Path -LiteralPath (Join-Path $target_repository 'Shared.bas') -PathType Leaf) -Message 'DIST did not update the opted-in immediate-child target.'
    Assert-True -Condition (-not (Test-Path -LiteralPath (Join-Path $target_repository 'Stale.bas'))) -Message 'DIST did not clear stale target content.'
    Assert-True -Condition (Test-Path -LiteralPath (Join-Path $nested_repository 'NestedSentinel.bas') -PathType Leaf) -Message 'DIST recursively discovered and changed a nested repository.'
    Assert-True -Condition (-not (Test-Path -LiteralPath (Join-Path $non_target_project 'common_modules_repo'))) -Message 'DIST created a missing opt-in repository.'
    Assert-True -Condition ((Get-Content -LiteralPath (Join-Path $source_repository 'Shared.bas') -Raw) -like '*distributed source*') -Message 'DIST modified the central source repository.'

    Write-TestFile -Path (Join-Path $target_repository 'StaleAgain.bas') -Content "Attribute VB_Name = `"StaleAgain`"`r`n" -Encoding $shift_jis
    $relative_result = Invoke-Dist -WorkingDirectory $working_directory -Arguments @('..\distribution-search-root')
    Assert-Equal -Expected 0 -Actual $relative_result.ExitCode -Message "DIST did not resolve a relative Search Root against the invocation working directory. Output: $($relative_result.Output)"
    Assert-True -Condition (-not (Test-Path -LiteralPath (Join-Path $target_repository 'StaleAgain.bas'))) -Message 'Relative-root DIST did not update the expected target.'

    $argument_fixture = New-DistFixture -FixtureRoot (Join-Path $temp_root 'argument-contract')
    [void](Invoke-ExpectedDistFailure -WorkingDirectory $argument_fixture.WorkingDirectory -ExpectedPattern '*exactly one Distribution Search Root*' -Message 'DIST accepted a missing Search Root argument.')
    [void](Invoke-ExpectedDistFailure -WorkingDirectory $argument_fixture.WorkingDirectory -Arguments @('') -ExpectedPattern '*Distribution global failure*' -Message 'DIST accepted an empty Search Root argument.')
    [void](Invoke-ExpectedDistFailure -WorkingDirectory $argument_fixture.WorkingDirectory -Arguments @($argument_fixture.SearchRoot, $argument_fixture.SearchRoot) -ExpectedPattern '*exactly one Distribution Search Root*' -Message 'DIST accepted more than one public argument.')
    [void](Invoke-ExpectedDistFailure -WorkingDirectory $argument_fixture.WorkingDirectory -Arguments @((Join-Path $argument_fixture.WorkingDirectory 'missing-root')) -ExpectedPattern '*Distribution global failure*' -Message 'DIST accepted an absent Search Root.')
    $search_root_file = Join-Path $argument_fixture.WorkingDirectory 'search-root.txt'
    Write-TestFile -Path $search_root_file -Content 'not a directory' -Encoding $shift_jis
    [void](Invoke-ExpectedDistFailure -WorkingDirectory $argument_fixture.WorkingDirectory -Arguments @($search_root_file) -ExpectedPattern '*must be a directory*' -Message 'DIST accepted a file as its Search Root.')
    Assert-True -Condition (Test-Path -LiteralPath (Join-Path $argument_fixture.TargetRepository 'Sentinel.bas') -PathType Leaf) -Message 'Argument validation mutated a target.'

    $search_alias_fixture = New-DistFixture -FixtureRoot (Join-Path $temp_root 'search-root-alias')
    $search_root_alias = Join-Path (Split-Path -Parent $search_alias_fixture.SearchRoot) 'search-root-junction'
    New-Item -ItemType Junction -Path $search_root_alias -Target $search_alias_fixture.SearchRoot | Out-Null
    $search_alias_result = Invoke-Dist -WorkingDirectory $search_alias_fixture.WorkingDirectory -Arguments @($search_root_alias)
    Assert-Equal -Expected 0 -Actual $search_alias_result.ExitCode -Message "DIST rejected an enumerable explicit reparse Search Root. Output: $($search_alias_result.Output)"
    Assert-True -Condition (Test-Path -LiteralPath (Join-Path $search_alias_fixture.TargetRepository 'Shared.bas') -PathType Leaf) -Message 'DIST did not use the explicit reparse Search Root.'

    $missing_source_root = Join-Path $temp_root 'missing-source'
    $missing_source_working = Join-Path $missing_source_root 'wrapper'
    $missing_source_search = Join-Path $missing_source_root 'search\ProjectA\common_modules_repo'
    New-Item -ItemType Directory -Path $missing_source_working, $missing_source_search -Force | Out-Null
    Write-TestFile -Path (Join-Path $missing_source_search 'Sentinel.bas') -Content 'keep' -Encoding $shift_jis
    [void](Invoke-ExpectedDistFailure -WorkingDirectory $missing_source_working -Arguments @((Split-Path -Parent (Split-Path -Parent $missing_source_search))) -ExpectedPattern '*Source common_modules_repo was not found*' -Message 'DIST searched for or created a missing CWD source repository.')
    Assert-True -Condition (Test-Path -LiteralPath (Join-Path $missing_source_search 'Sentinel.bas') -PathType Leaf) -Message 'Missing-source validation mutated a target.'

    $wrong_type_source_root = Join-Path $temp_root 'wrong-type-source'
    $wrong_type_source_working = Join-Path $wrong_type_source_root 'wrapper'
    $wrong_type_source_search = Join-Path $wrong_type_source_root 'search'
    New-Item -ItemType Directory -Path $wrong_type_source_working, (Join-Path $wrong_type_source_search 'ProjectA\common_modules_repo') -Force | Out-Null
    Write-TestFile -Path (Join-Path $wrong_type_source_working 'common_modules_repo') -Content 'not a directory' -Encoding $shift_jis
    [void](Invoke-ExpectedDistFailure -WorkingDirectory $wrong_type_source_working -Arguments @($wrong_type_source_search) -ExpectedPattern '*ordinary non-reparse directory*' -Message 'DIST accepted a file as the CWD source repository.')

    $case_source_root = Join-Path $temp_root 'case-source'
    $case_source_working = Join-Path $case_source_root 'wrapper'
    $case_source_repository = Join-Path $case_source_working 'Common_Modules_Repo'
    $case_source_search = Join-Path $case_source_root 'search'
    Write-ValidPackage -RepositoryPath $case_source_repository
    New-Item -ItemType Directory -Path (Join-Path $case_source_search 'ProjectA\common_modules_repo') -Force | Out-Null
    [void](Invoke-ExpectedDistFailure -WorkingDirectory $case_source_working -Arguments @($case_source_search) -ExpectedPattern '*ordinal-exact*' -Message 'DIST accepted a case-only CWD source repository name.')

    $source_link_root = Join-Path $temp_root 'source-link'
    $source_link_working = Join-Path $source_link_root 'wrapper'
    $source_link_backing = Join-Path $source_link_root 'backing-package'
    $source_link_search = Join-Path $source_link_root 'search'
    Write-ValidPackage -RepositoryPath $source_link_backing
    New-Item -ItemType Directory -Path $source_link_working, (Join-Path $source_link_search 'ProjectA\common_modules_repo') -Force | Out-Null
    New-Item -ItemType Junction -Path (Join-Path $source_link_working 'common_modules_repo') -Target $source_link_backing | Out-Null
    [void](Invoke-ExpectedDistFailure -WorkingDirectory $source_link_working -Arguments @($source_link_search) -ExpectedPattern '*ordinary non-reparse directory*' -Message 'DIST followed a reparse CWD source repository.')

    foreach ($source_defect in @('extra-file', 'nested-directory', 'missing-module', 'invalid-manifest', 'orphan-sidecar')) {
        $defect_fixture = New-DistFixture -FixtureRoot (Join-Path $temp_root ('source-defect-' + $source_defect))
        switch ($source_defect) {
            'extra-file' {
                Write-TestFile -Path (Join-Path $defect_fixture.SourceRepository 'Extra.txt') -Content 'extra' -Encoding $shift_jis
            }
            'nested-directory' {
                New-Item -ItemType Directory -Path (Join-Path $defect_fixture.SourceRepository 'Nested') -Force | Out-Null
            }
            'missing-module' {
                Remove-Item -LiteralPath (Join-Path $defect_fixture.SourceRepository 'Shared.bas') -Force
            }
            'invalid-manifest' {
                Write-TestFile -Path (Join-Path $defect_fixture.SourceRepository 'common-modules-manifest.tsv') -Content "invalid`r`n" -Encoding (New-Object System.Text.UTF8Encoding($true))
            }
            'orphan-sidecar' {
                Write-TestFile -Path (Join-Path $defect_fixture.SourceRepository 'Orphan.frx') -Content 'orphan' -Encoding $shift_jis
            }
        }

        [void](Invoke-ExpectedDistFailure -WorkingDirectory $defect_fixture.WorkingDirectory -Arguments @($defect_fixture.SearchRoot) -ExpectedPattern '*Distribution global failure*' -Message "DIST accepted source defect '$source_defect'.")
        Assert-True -Condition (Test-Path -LiteralPath (Join-Path $defect_fixture.TargetRepository 'Sentinel.bas') -PathType Leaf) -Message "Source defect '$source_defect' mutated a target before global validation completed."
    }

    $locked_source_fixture = New-DistFixture -FixtureRoot (Join-Path $temp_root 'locked-source')
    $locked_manifest_path = Join-Path $locked_source_fixture.SourceRepository 'common-modules-manifest.tsv'
    $locked_manifest_stream = [System.IO.File]::Open($locked_manifest_path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::None)
    try {
        [void](Invoke-ExpectedDistFailure -WorkingDirectory $locked_source_fixture.WorkingDirectory -Arguments @($locked_source_fixture.SearchRoot) -ExpectedPattern '*Distribution global failure*' -Message 'DIST accepted an unreadable source manifest.')
    }
    finally {
        $locked_manifest_stream.Dispose()
    }
    Assert-True -Condition (Test-Path -LiteralPath (Join-Path $locked_source_fixture.TargetRepository 'Sentinel.bas') -PathType Leaf) -Message 'Unreadable source validation mutated a target.'

    $project_link_fixture = New-DistFixture -FixtureRoot (Join-Path $temp_root 'project-link-warning')
    $linked_project_backing = Join-Path $temp_root 'linked-project-backing'
    $linked_target = Join-Path $linked_project_backing 'common_modules_repo'
    New-Item -ItemType Directory -Path $linked_target -Force | Out-Null
    Write-TestFile -Path (Join-Path $linked_target 'LinkedSentinel.bas') -Content 'keep' -Encoding $shift_jis
    New-Item -ItemType Junction -Path (Join-Path $project_link_fixture.SearchRoot 'A_LinkedProject') -Target $linked_project_backing | Out-Null
    $project_link_result = Invoke-Dist -WorkingDirectory $project_link_fixture.WorkingDirectory -Arguments @($project_link_fixture.SearchRoot)
    Assert-Equal -Expected 0 -Actual $project_link_result.ExitCode -Message "A reparse project child should be a warning skip while later targets continue. Output: $($project_link_result.Output)"
    Assert-Like -Actual $project_link_result.Output -Pattern '*reparse point and was skipped*' -Message 'DIST did not report a reparse project child as a warning skip.'
    Assert-True -Condition (Test-Path -LiteralPath (Join-Path $linked_target 'LinkedSentinel.bas') -PathType Leaf) -Message 'DIST followed and changed a reparse project child.'
    Assert-True -Condition (Test-Path -LiteralPath (Join-Path $project_link_fixture.TargetRepository 'Shared.bas') -PathType Leaf) -Message 'DIST did not continue after a reparse project warning skip.'

    $classification_root = Join-Path $temp_root 'candidate-classification'
    $classification_working = Join-Path $classification_root 'wrapper'
    $classification_source = Join-Path $classification_working 'common_modules_repo'
    $classification_search = Join-Path $classification_root 'search-root'
    $case_only_target = Join-Path $classification_search 'A_CaseOnly\Common_Modules_Repo'
    $wrong_type_project = Join-Path $classification_search 'B_WrongType'
    $valid_classification_target = Join-Path $classification_search 'C_Valid\common_modules_repo'
    Write-ValidPackage -RepositoryPath $classification_source
    New-Item -ItemType Directory -Path $case_only_target, $wrong_type_project, $valid_classification_target -Force | Out-Null
    Write-TestFile -Path (Join-Path $case_only_target 'CaseSentinel.bas') -Content 'keep' -Encoding $shift_jis
    Write-TestFile -Path (Join-Path $wrong_type_project 'common_modules_repo') -Content 'wrong type' -Encoding $shift_jis
    Write-TestFile -Path (Join-Path $valid_classification_target 'Stale.bas') -Content 'replace' -Encoding $shift_jis
    $classification_result = Invoke-Dist -WorkingDirectory $classification_working -Arguments @($classification_search)
    Assert-Equal -Expected 1 -Actual $classification_result.ExitCode -Message "Candidate defects should fail only after later targets are attempted. Output: $($classification_result.Output)"
    Assert-Like -Actual $classification_result.Output -Pattern '*ordinal-exact*' -Message 'DIST did not report a case-only target as a candidate failure.'
    Assert-Like -Actual $classification_result.Output -Pattern '*ordinary non-reparse*directory*' -Message 'DIST did not report a wrong-type target as a candidate failure.'
    Assert-True -Condition ($classification_result.Output -notlike '*No eligible distribution target*') -Message 'DIST incorrectly added a zero-target error when known candidate failures existed.'
    Assert-True -Condition (Test-Path -LiteralPath (Join-Path $case_only_target 'CaseSentinel.bas') -PathType Leaf) -Message 'DIST mutated a case-only candidate failure.'
    Assert-True -Condition (Test-Path -LiteralPath (Join-Path $valid_classification_target 'Shared.bas') -PathType Leaf) -Message 'DIST did not continue to the valid target after candidate failures.'

    $source_exclusion_root = Join-Path $temp_root 'source-exclusion'
    $source_owner = Join-Path $source_exclusion_root 'B_SourceOwner'
    $excluded_source = Join-Path $source_owner 'common_modules_repo'
    $source_exclusion_target = Join-Path $source_exclusion_root 'A_Target\common_modules_repo'
    Write-ValidPackage -RepositoryPath $excluded_source -ModuleContent "Attribute VB_Name = `"Shared`"`r`n'exclusion source`r`n"
    New-Item -ItemType Directory -Path $source_exclusion_target -Force | Out-Null
    Write-TestFile -Path (Join-Path $source_exclusion_target 'Stale.bas') -Content 'replace' -Encoding $shift_jis
    $source_exclusion_result = Invoke-Dist -WorkingDirectory $source_owner -Arguments @($source_exclusion_root)
    Assert-Equal -Expected 0 -Actual $source_exclusion_result.ExitCode -Message "DIST did not exclude the normalized lexical source path. Output: $($source_exclusion_result.Output)"
    Assert-True -Condition ((Get-Content -LiteralPath (Join-Path $excluded_source 'Shared.bas') -Raw) -like '*exclusion source*') -Message 'DIST modified the central source while it appeared inside the Search Root.'
    Assert-True -Condition (Test-Path -LiteralPath (Join-Path $source_exclusion_target 'Shared.bas') -PathType Leaf) -Message 'DIST did not update the non-source candidate.'

    $unchanged_fixture = New-DistFixture -FixtureRoot (Join-Path $temp_root 'metadata-unchanged')
    Remove-Item -LiteralPath (Join-Path $unchanged_fixture.TargetRepository 'Sentinel.bas') -Force
    Copy-PackageFiles -SourceRepository $unchanged_fixture.SourceRepository -TargetRepository $unchanged_fixture.TargetRepository
    $unchanged_module_path = Join-Path $unchanged_fixture.TargetRepository 'Shared.bas'
    $unchanged_time = [System.IO.File]::GetLastWriteTimeUtc($unchanged_module_path)
    $unchanged_lock = [System.IO.File]::Open($unchanged_module_path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::None)
    try {
        $unchanged_result = Invoke-Dist -WorkingDirectory $unchanged_fixture.WorkingDirectory -Arguments @($unchanged_fixture.SearchRoot)
    }
    finally {
        $unchanged_lock.Dispose()
    }
    Assert-Equal -Expected 0 -Actual $unchanged_result.ExitCode -Message "Metadata-identical target was not treated as UNCHANGED without opening file bytes. Output: $($unchanged_result.Output)"
    Assert-Like -Actual $unchanged_result.Output -Pattern '*UNCHANGED:*' -Message 'DIST did not report an exact metadata package match as UNCHANGED.'
    Assert-Equal -Expected $unchanged_time -Actual ([System.IO.File]::GetLastWriteTimeUtc($unchanged_module_path)) -Message 'DIST rewrote an UNCHANGED target.'

    $replacement_root = Join-Path $temp_root 'replacement-triggers'
    $replacement_working = Join-Path $replacement_root 'wrapper'
    $replacement_source = Join-Path $replacement_working 'common_modules_repo'
    $replacement_search = Join-Path $replacement_root 'search-root'
    Write-ValidPackage -RepositoryPath $replacement_source
    $replacement_targets = @{}
    foreach ($project_name in @('A_Timestamp', 'B_Length', 'C_Extra', 'D_Nested', 'E_CaseOnly')) {
        $replacement_target = Join-Path $replacement_search "$project_name\common_modules_repo"
        Copy-PackageFiles -SourceRepository $replacement_source -TargetRepository $replacement_target
        $replacement_targets[$project_name] = $replacement_target
    }
    [System.IO.File]::SetLastWriteTimeUtc((Join-Path $replacement_targets['A_Timestamp'] 'Shared.bas'), [datetime]'2020-01-01T00:00:00Z')
    Write-TestFile -Path (Join-Path $replacement_targets['B_Length'] 'Shared.bas') -Content 'different length' -Encoding $shift_jis
    Write-TestFile -Path (Join-Path $replacement_targets['C_Extra'] 'Extra.bas') -Content 'extra' -Encoding $shift_jis
    New-Item -ItemType Directory -Path (Join-Path $replacement_targets['D_Nested'] 'Nested') -Force | Out-Null
    Rename-Item -LiteralPath (Join-Path $replacement_targets['E_CaseOnly'] 'Shared.bas') -NewName 'shared.bas'
    $replacement_result = Invoke-Dist -WorkingDirectory $replacement_working -Arguments @($replacement_search)
    Assert-Equal -Expected 0 -Actual $replacement_result.ExitCode -Message "A metadata or inventory difference did not trigger full replacement. Output: $($replacement_result.Output)"
    foreach ($replacement_target in $replacement_targets.Values) {
        $actual_names = @((Get-ChildItem -LiteralPath $replacement_target -Force).Name)
        Assert-Equal -Expected 2 -Actual $actual_names.Count -Message "DIST did not produce the complete closed package in '$replacement_target'."
        Assert-True -Condition ($actual_names -ccontains 'Shared.bas') -Message "DIST did not restore ordinal-exact package spelling in '$replacement_target'."
        Assert-True -Condition ($actual_names -ccontains 'common-modules-manifest.tsv') -Message "DIST did not restore the exact manifest in '$replacement_target'."
        Assert-Equal -Expected ([System.IO.File]::GetLastWriteTimeUtc((Join-Path $replacement_source 'Shared.bas'))) -Actual ([System.IO.File]::GetLastWriteTimeUtc((Join-Path $replacement_target 'Shared.bas'))) -Message "DIST did not preserve source metadata in '$replacement_target'."
    }

    $form_root = Join-Path $temp_root 'form-package'
    $form_working = Join-Path $form_root 'wrapper'
    $form_source = Join-Path $form_working 'common_modules_repo'
    $form_search = Join-Path $form_root 'search-root'
    $form_target = Join-Path $form_search 'ProjectA\common_modules_repo'
    New-Item -ItemType Directory -Path $form_source, $form_target -Force | Out-Null
    Write-TestFile -Path (Join-Path $form_source 'common-modules-manifest.tsv') -Content "ModuleFile`tCategories`tDependencies`tRequiredReferences`r`nDialog.frm`truntime-baseline`t`t[]`r`n" -Encoding $utf16_le
    Write-TestFile -Path (Join-Path $form_source 'Dialog.frm') -Content "VERSION 5.00`r`nAttribute VB_Name = `"Dialog`"`r`n" -Encoding $shift_jis
    Write-TestFile -Path (Join-Path $form_source 'Dialog.frx') -Content 'binary-sidecar-fixture' -Encoding $shift_jis
    Write-TestFile -Path (Join-Path $form_target 'Stale.bas') -Content 'stale' -Encoding $shift_jis
    $form_result = Invoke-Dist -WorkingDirectory $form_working -Arguments @($form_search)
    Assert-Equal -Expected 0 -Actual $form_result.ExitCode -Message "DIST did not distribute an optional exact form sidecar. Output: $($form_result.Output)"
    Assert-True -Condition (Test-Path -LiteralPath (Join-Path $form_target 'Dialog.frm') -PathType Leaf) -Message 'DIST did not copy a manifest-listed form.'
    Assert-True -Condition (Test-Path -LiteralPath (Join-Path $form_target 'Dialog.frx') -PathType Leaf) -Message 'DIST did not copy the matching optional form sidecar.'

    $target_reparse_root = Join-Path $temp_root 'target-reparse-safety'
    $target_reparse_working = Join-Path $target_reparse_root 'wrapper'
    $target_reparse_source = Join-Path $target_reparse_working 'common_modules_repo'
    $target_reparse_search = Join-Path $target_reparse_root 'search-root'
    $target_root_backing = Join-Path $target_reparse_root 'target-root-backing'
    $descendant_backing = Join-Path $target_reparse_root 'descendant-backing'
    $target_root_project = Join-Path $target_reparse_search 'A_TargetRootLink'
    $descendant_target = Join-Path $target_reparse_search 'B_DescendantLink\common_modules_repo'
    $valid_reparse_target = Join-Path $target_reparse_search 'C_Valid\common_modules_repo'
    Write-ValidPackage -RepositoryPath $target_reparse_source
    New-Item -ItemType Directory -Path $target_root_backing, $target_root_project, $descendant_target, $descendant_backing, $valid_reparse_target -Force | Out-Null
    Write-TestFile -Path (Join-Path $target_root_backing 'RootLinkSentinel.bas') -Content 'keep' -Encoding $shift_jis
    Write-TestFile -Path (Join-Path $descendant_backing 'DescendantSentinel.bas') -Content 'keep' -Encoding $shift_jis
    Write-TestFile -Path (Join-Path $descendant_target 'TargetSentinel.bas') -Content 'keep' -Encoding $shift_jis
    Write-TestFile -Path (Join-Path $valid_reparse_target 'Stale.bas') -Content 'replace' -Encoding $shift_jis
    New-Item -ItemType Junction -Path (Join-Path $target_root_project 'common_modules_repo') -Target $target_root_backing | Out-Null
    New-Item -ItemType Junction -Path (Join-Path $descendant_target 'NestedLink') -Target $descendant_backing | Out-Null
    $target_reparse_result = Invoke-Dist -WorkingDirectory $target_reparse_working -Arguments @($target_reparse_search)
    Assert-Equal -Expected 1 -Actual $target_reparse_result.ExitCode -Message "Unsafe target reparses should be isolated failures while later targets continue. Output: $($target_reparse_result.Output)"
    Assert-Like -Actual $target_reparse_result.Output -Pattern '*reparse*' -Message 'DIST did not report unsafe target reparses.'
    Assert-True -Condition (Test-Path -LiteralPath (Join-Path $target_root_backing 'RootLinkSentinel.bas') -PathType Leaf) -Message 'DIST followed and changed a reparse target root.'
    Assert-True -Condition (Test-Path -LiteralPath (Join-Path $descendant_target 'TargetSentinel.bas') -PathType Leaf) -Message 'DIST mutated a target containing a descendant reparse entry.'
    Assert-True -Condition (Test-Path -LiteralPath (Join-Path $descendant_backing 'DescendantSentinel.bas') -PathType Leaf) -Message 'DIST followed a descendant reparse entry.'
    Assert-True -Condition (Test-Path -LiteralPath (Join-Path $valid_reparse_target 'Shared.bas') -PathType Leaf) -Message 'DIST did not continue after target reparse failures.'

    $delete_failure_root = Join-Path $temp_root 'delete-failure-continuation'
    $delete_failure_working = Join-Path $delete_failure_root 'wrapper'
    $delete_failure_source = Join-Path $delete_failure_working 'common_modules_repo'
    $delete_failure_search = Join-Path $delete_failure_root 'search-root'
    $locked_target = Join-Path $delete_failure_search 'A_Locked\common_modules_repo'
    $later_target = Join-Path $delete_failure_search 'B_Later\common_modules_repo'
    Write-ValidPackage -RepositoryPath $delete_failure_source
    New-Item -ItemType Directory -Path $locked_target, $later_target -Force | Out-Null
    $locked_target_file = Join-Path $locked_target 'Locked.bas'
    Write-TestFile -Path $locked_target_file -Content 'locked' -Encoding $shift_jis
    Write-TestFile -Path (Join-Path $later_target 'Stale.bas') -Content 'replace' -Encoding $shift_jis
    $target_lock = [System.IO.File]::Open($locked_target_file, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::None)
    try {
        $delete_failure_result = Invoke-Dist -WorkingDirectory $delete_failure_working -Arguments @($delete_failure_search)
    }
    finally {
        $target_lock.Dispose()
    }
    Assert-Equal -Expected 1 -Actual $delete_failure_result.ExitCode -Message "A target delete failure should produce final failure after later targets are attempted. Output: $($delete_failure_result.Output)"
    Assert-Like -Actual $delete_failure_result.Output -Pattern '*candidate failure while clearing*' -Message 'DIST did not classify a target deletion error as a candidate failure.'
    Assert-True -Condition (Test-Path -LiteralPath $locked_target_file -PathType Leaf) -Message 'The locked target fixture unexpectedly disappeared.'
    Assert-True -Condition (Test-Path -LiteralPath (Join-Path $later_target 'Shared.bas') -PathType Leaf) -Message 'DIST did not continue after a target deletion failure.'

    $zero_fixture_root = Join-Path $temp_root 'zero-target'
    $zero_working = Join-Path $zero_fixture_root 'wrapper'
    $zero_source = Join-Path $zero_working 'common_modules_repo'
    $zero_search = Join-Path $zero_fixture_root 'search-root'
    Write-ValidPackage -RepositoryPath $zero_source
    New-Item -ItemType Directory -Path (Join-Path $zero_search 'ProjectWithoutOptIn') -Force | Out-Null
    [void](Invoke-ExpectedDistFailure -WorkingDirectory $zero_working -Arguments @($zero_search) -ExpectedPattern '*No eligible distribution target*' -Message 'DIST did not fail when initial discovery found no target or candidate failure.')

    $candidate_only_root = Join-Path $temp_root 'candidate-only'
    $candidate_only_working = Join-Path $candidate_only_root 'wrapper'
    $candidate_only_source = Join-Path $candidate_only_working 'common_modules_repo'
    $candidate_only_search = Join-Path $candidate_only_root 'search-root'
    Write-ValidPackage -RepositoryPath $candidate_only_source
    New-Item -ItemType Directory -Path (Join-Path $candidate_only_search 'ProjectA\Common_Modules_Repo') -Force | Out-Null
    $candidate_only_result = Invoke-Dist -WorkingDirectory $candidate_only_working -Arguments @($candidate_only_search)
    Assert-Equal -Expected 1 -Actual $candidate_only_result.ExitCode -Message 'A known candidate failure did not produce final failure.'
    Assert-Like -Actual $candidate_only_result.Output -Pattern '*candidate failure for*' -Message 'DIST did not report the candidate-only defect.'
    Assert-True -Condition ($candidate_only_result.Output -notlike '*No eligible distribution target*') -Message 'DIST added a zero-target error when a known candidate failure already explained the failure.'

    $warning_only_root = Join-Path $temp_root 'warning-only-zero'
    $warning_only_working = Join-Path $warning_only_root 'wrapper'
    $warning_only_source = Join-Path $warning_only_working 'common_modules_repo'
    $warning_only_search = Join-Path $warning_only_root 'search-root'
    $warning_only_backing = Join-Path $warning_only_root 'linked-project'
    Write-ValidPackage -RepositoryPath $warning_only_source
    New-Item -ItemType Directory -Path $warning_only_search, (Join-Path $warning_only_backing 'common_modules_repo') -Force | Out-Null
    New-Item -ItemType Junction -Path (Join-Path $warning_only_search 'ProjectLink') -Target $warning_only_backing | Out-Null
    $warning_only_result = Invoke-Dist -WorkingDirectory $warning_only_working -Arguments @($warning_only_search)
    Assert-Equal -Expected 1 -Actual $warning_only_result.ExitCode -Message 'A warning-only uncertain scan with no admitted target did not produce zero-target failure.'
    Assert-Like -Actual $warning_only_result.Output -Pattern '*reparse point and was skipped*' -Message 'DIST did not report the warning-only uncertain project.'
    Assert-Like -Actual $warning_only_result.Output -Pattern '*No eligible distribution target*' -Message 'Warning-only uncertainty incorrectly suppressed the zero-target error.'

    $fixed_set_root = Join-Path $temp_root 'fixed-candidate-set'
    $fixed_set_working = Join-Path $fixed_set_root 'wrapper'
    $fixed_set_source = Join-Path $fixed_set_working 'common_modules_repo'
    $fixed_set_search = Join-Path $fixed_set_root 'search-root'
    $fixed_target_a = Join-Path $fixed_set_search 'A_First\common_modules_repo'
    $fixed_target_b = Join-Path $fixed_set_search 'B_Disappears\common_modules_repo'
    $fixed_target_c = Join-Path $fixed_set_search 'C_Later\common_modules_repo'
    $late_project = Join-Path $fixed_set_search 'D_LateArrival'
    $late_repository = Join-Path $fixed_set_root 'late-repository'
    $late_destination = Join-Path $late_project 'common_modules_repo'
    $disappeared_holding = Join-Path $fixed_set_root 'disappeared-holding'
    Write-LargeTestPackage -RepositoryPath $fixed_set_source -TailModuleName 'Z_Tail.bas'
    New-Item -ItemType Directory -Path $fixed_target_a, $fixed_target_b, $fixed_target_c, $late_project, $late_repository -Force | Out-Null
    Write-TestFile -Path (Join-Path $fixed_target_a 'Stale.bas') -Content 'stale' -Encoding $shift_jis
    Write-TestFile -Path (Join-Path $fixed_target_b 'DisappearingSentinel.bas') -Content 'move' -Encoding $shift_jis
    Write-TestFile -Path (Join-Path $fixed_target_c 'Stale.bas') -Content 'stale' -Encoding $shift_jis
    Write-TestFile -Path (Join-Path $late_repository 'LateSentinel.bas') -Content 'late' -Encoding $shift_jis
    $fixed_helper_script = Join-Path $fixed_set_root 'move-targets.ps1'
    Write-TestFile -Path $fixed_helper_script -Content @'
param($TriggerPath, $DisappearingPath, $HoldingPath, $LatePath, $LateDestination)
$deadline = [datetime]::UtcNow.AddSeconds(30)
while (-not (Test-Path -LiteralPath $TriggerPath -PathType Leaf) -and [datetime]::UtcNow -lt $deadline) {
    Start-Sleep -Milliseconds 5
}
if (-not (Test-Path -LiteralPath $TriggerPath -PathType Leaf)) {
    exit 2
}
[System.IO.Directory]::Move($DisappearingPath, $HoldingPath)
[System.IO.Directory]::Move($LatePath, $LateDestination)
exit 0
'@ -Encoding (New-Object System.Text.UTF8Encoding($true))
    $fixed_helper = Start-TestHelperProcess -ScriptPath $fixed_helper_script -Arguments @(
        (Join-Path $fixed_target_a 'A_Trigger.bas'),
        $fixed_target_b,
        $disappeared_holding,
        $late_repository,
        $late_destination
    )
    $fixed_set_result = Invoke-Dist -WorkingDirectory $fixed_set_working -Arguments @($fixed_set_search)
    Wait-TestHelperProcess -Process $fixed_helper -Description 'Fixed candidate-set helper'
    Assert-Equal -Expected 0 -Actual $fixed_set_result.ExitCode -Message "An admitted target disappearance should be a warning-only skip. Output: $($fixed_set_result.Output)"
    Assert-Like -Actual $fixed_set_result.Output -Pattern '*disappeared before its turn*' -Message 'DIST did not warning-skip an admitted target that disappeared before its turn.'
    Assert-True -Condition (Test-Path -LiteralPath (Join-Path $fixed_target_a 'Z_Tail.bas') -PathType Leaf) -Message 'DIST did not finish the first fixed candidate.'
    Assert-True -Condition (Test-Path -LiteralPath (Join-Path $fixed_target_c 'Z_Tail.bas') -PathType Leaf) -Message 'DIST did not continue after an admitted target disappeared.'
    Assert-True -Condition (Test-Path -LiteralPath (Join-Path $disappeared_holding 'DisappearingSentinel.bas') -PathType Leaf) -Message 'The disappearance helper did not preserve the moved target fixture.'
    Assert-True -Condition (Test-Path -LiteralPath (Join-Path $late_destination 'LateSentinel.bas') -PathType Leaf) -Message 'The late-arrival fixture was not installed.'
    Assert-True -Condition (-not (Test-Path -LiteralPath (Join-Path $late_destination 'A_Trigger.bas'))) -Message 'DIST rescanned and processed a target that appeared after initial discovery.'

    $partial_root = Join-Path $temp_root 'partial-copy-continuation'
    $partial_working = Join-Path $partial_root 'wrapper'
    $partial_source = Join-Path $partial_working 'common_modules_repo'
    $partial_search = Join-Path $partial_root 'search-root'
    $partial_target_a = Join-Path $partial_search 'A_Partial\common_modules_repo'
    $partial_target_b = Join-Path $partial_search 'B_Later\common_modules_repo'
    $partial_lock_path = Join-Path $partial_target_a 'Z_Fail.bas'
    $partial_ready_path = Join-Path $partial_root 'lock-ready.txt'
    $partial_stop_path = Join-Path $partial_root 'stop-lock.txt'
    Write-LargeTestPackage -RepositoryPath $partial_source -TailModuleName 'Z_Fail.bas'
    New-Item -ItemType Directory -Path $partial_target_a, $partial_target_b -Force | Out-Null
    Write-TestFile -Path (Join-Path $partial_target_a 'Stale.bas') -Content 'stale' -Encoding $shift_jis
    Write-TestFile -Path (Join-Path $partial_target_b 'Stale.bas') -Content 'stale' -Encoding $shift_jis
    $partial_helper_script = Join-Path $partial_root 'lock-copy-target.ps1'
    Write-TestFile -Path $partial_helper_script -Content @'
param($TriggerPath, $LockPath, $ReadyPath, $StopPath)
$deadline = [datetime]::UtcNow.AddSeconds(30)
while (-not (Test-Path -LiteralPath $TriggerPath -PathType Leaf) -and [datetime]::UtcNow -lt $deadline) {
    Start-Sleep -Milliseconds 5
}
if (-not (Test-Path -LiteralPath $TriggerPath -PathType Leaf)) {
    exit 2
}
$stream = [System.IO.File]::Open($LockPath, [System.IO.FileMode]::OpenOrCreate, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
try {
    [System.IO.File]::WriteAllText($ReadyPath, 'ready')
    $stop_deadline = [datetime]::UtcNow.AddSeconds(30)
    while (-not (Test-Path -LiteralPath $StopPath -PathType Leaf) -and [datetime]::UtcNow -lt $stop_deadline) {
        Start-Sleep -Milliseconds 5
    }
    if (-not (Test-Path -LiteralPath $StopPath -PathType Leaf)) {
        exit 3
    }
}
finally {
    $stream.Dispose()
}
exit 0
'@ -Encoding (New-Object System.Text.UTF8Encoding($true))
    $partial_helper = Start-TestHelperProcess -ScriptPath $partial_helper_script -Arguments @(
        (Join-Path $partial_target_a 'A_Trigger.bas'),
        $partial_lock_path,
        $partial_ready_path,
        $partial_stop_path
    )
    try {
        $partial_result = Invoke-Dist -WorkingDirectory $partial_working -Arguments @($partial_search)
    }
    finally {
        Write-TestFile -Path $partial_stop_path -Content 'stop' -Encoding $shift_jis
    }
    Wait-TestHelperProcess -Process $partial_helper -Description 'Partial-copy helper'
    Assert-Equal -Expected 1 -Actual $partial_result.ExitCode -Message "A target-local copy failure should permit later candidates and produce final failure. Output: $($partial_result.Output)"
    Assert-Like -Actual $partial_result.Output -Pattern '*candidate failure while copying*' -Message 'DIST did not classify a target write failure as candidate-local.'
    Assert-True -Condition (Test-Path -LiteralPath (Join-Path $partial_target_a 'A_Trigger.bas') -PathType Leaf) -Message 'The failed target did not retain an allowed partial copy.'
    Assert-True -Condition (-not (Test-Path -LiteralPath (Join-Path $partial_target_a 'common-modules-manifest.tsv'))) -Message 'The failed target unexpectedly received the complete package.'
    Assert-True -Condition (Test-Path -LiteralPath (Join-Path $partial_target_b 'common-modules-manifest.tsv') -PathType Leaf) -Message 'DIST did not complete a later target after a target-local copy failure.'

    $source_change_root = Join-Path $temp_root 'source-change-global'
    $source_change_working = Join-Path $source_change_root 'wrapper'
    $source_change_source = Join-Path $source_change_working 'common_modules_repo'
    $source_change_search = Join-Path $source_change_root 'search-root'
    $source_change_target_a = Join-Path $source_change_search 'A_First\common_modules_repo'
    $source_change_target_b = Join-Path $source_change_search 'B_MustRemain\common_modules_repo'
    $source_tail_path = Join-Path $source_change_source 'Z_SourceGone.bas'
    $source_tail_holding = Join-Path $source_change_root 'Z_SourceGone.bas'
    Write-LargeTestPackage -RepositoryPath $source_change_source -TailModuleName 'Z_SourceGone.bas'
    New-Item -ItemType Directory -Path $source_change_target_a, $source_change_target_b -Force | Out-Null
    Write-TestFile -Path (Join-Path $source_change_target_a 'Stale.bas') -Content 'stale' -Encoding $shift_jis
    Write-TestFile -Path (Join-Path $source_change_target_b 'MustRemain.bas') -Content 'keep' -Encoding $shift_jis
    $source_change_helper_script = Join-Path $source_change_root 'remove-source.ps1'
    Write-TestFile -Path $source_change_helper_script -Content @'
param($TriggerPath, $SourcePath, $HoldingPath)
$deadline = [datetime]::UtcNow.AddSeconds(30)
while (-not (Test-Path -LiteralPath $TriggerPath -PathType Leaf) -and [datetime]::UtcNow -lt $deadline) {
    Start-Sleep -Milliseconds 5
}
if (-not (Test-Path -LiteralPath $TriggerPath -PathType Leaf)) {
    exit 2
}
[System.IO.File]::Move($SourcePath, $HoldingPath)
exit 0
'@ -Encoding (New-Object System.Text.UTF8Encoding($true))
    $source_change_helper = Start-TestHelperProcess -ScriptPath $source_change_helper_script -Arguments @(
        (Join-Path $source_change_target_a 'A_Trigger.bas'),
        $source_tail_path,
        $source_tail_holding
    )
    $source_change_result = Invoke-Dist -WorkingDirectory $source_change_working -Arguments @($source_change_search)
    Wait-TestHelperProcess -Process $source_change_helper -Description 'Source-change helper'
    Assert-Equal -Expected 1 -Actual $source_change_result.ExitCode -Message 'A newly invalid source package should stop distribution globally.'
    Assert-Like -Actual $source_change_result.Output -Pattern '*Distribution global failure*source package*' -Message 'DIST did not classify a source change during copy as a global failure.'
    Assert-True -Condition (Test-Path -LiteralPath (Join-Path $source_change_target_b 'MustRemain.bas') -PathType Leaf) -Message 'DIST continued to a later target after a global source failure.'
    Assert-True -Condition (-not (Test-Path -LiteralPath (Join-Path $source_change_target_b 'A_Trigger.bas'))) -Message 'DIST mutated a later target after a global source failure.'

    $bat_fixture = Join-Path $temp_root 'bat-wrapper'
    New-Item -ItemType Directory -Path $bat_fixture -Force | Out-Null
    $bat_path = Join-Path $bat_fixture 'DIST_COMMON_MODS_REPO.BAT'
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'DIST_COMMON_MODS_REPO.BAT') -Destination $bat_path -Force
    $bat_record_path = Join-Path $bat_fixture 'arguments.txt'
    $timeout_record_path = Join-Path $bat_fixture 'timeout.txt'
    Write-TestFile -Path (Join-Path $bat_fixture 'dist_common_mods_repo_main.ps1') -Content @"
`$args | Set-Content -LiteralPath `$env:DIST_TEST_RECORD_PATH -Encoding UTF8
exit 37
"@ -Encoding (New-Object System.Text.UTF8Encoding($true))
    [System.IO.File]::WriteAllText(
        (Join-Path $bat_fixture 'TIMEOUT.CMD'),
        "@echo off`r`n>`"%DIST_TIMEOUT_RECORD_PATH%`" echo called`r`nexit /b 0`r`n",
        [System.Text.Encoding]::GetEncoding(932)
    )

    $previous_path = $env:PATH
    $previous_record_path = $env:DIST_TEST_RECORD_PATH
    $previous_timeout_record_path = $env:DIST_TIMEOUT_RECORD_PATH
    try {
        $env:PATH = "$bat_fixture;$env:PATH"
        $env:DIST_TEST_RECORD_PATH = $bat_record_path
        $env:DIST_TIMEOUT_RECORD_PATH = $timeout_record_path
        $bat_output = & $env:ComSpec /d /c "`"$bat_path`" `"$search_root`"" 2>&1
        $bat_exit_code = $LASTEXITCODE
    }
    finally {
        $env:PATH = $previous_path
        $env:DIST_TEST_RECORD_PATH = $previous_record_path
        $env:DIST_TIMEOUT_RECORD_PATH = $previous_timeout_record_path
    }

    Assert-Equal -Expected 37 -Actual $bat_exit_code -Message "DIST BAT did not preserve the PowerShell exit code after TIMEOUT. Output: $($bat_output | Out-String)"
    Assert-True -Condition (Test-Path -LiteralPath $timeout_record_path -PathType Leaf) -Message 'DIST BAT did not run its informational TIMEOUT command.'
    $bat_arguments = @(Get-Content -LiteralPath $bat_record_path)
    Assert-Equal -Expected 1 -Actual $bat_arguments.Count -Message 'DIST BAT did not forward exactly one public argument to PowerShell.'
    Assert-Equal -Expected $search_root -Actual $bat_arguments[0] -Message 'DIST BAT changed the Distribution Search Root argument.'
    $bat_text = Get-Content -LiteralPath $bat_path -Raw
    Assert-True -Condition ($bat_text -match 'powershell\s+-NoProfile') -Message 'DIST BAT did not invoke PowerShell with -NoProfile.'

    Write-Host 'DIST CommonModules tests passed.'
}
finally {
    if (Test-Path -LiteralPath $temp_root) {
        Remove-Item -LiteralPath $temp_root -Recurse -Force
    }
}
