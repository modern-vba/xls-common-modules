$ErrorActionPreference = 'Stop'
$InformationPreference = 'Continue'

$validator_path = Join-Path $PSScriptRoot 'validate_common_modules_manifest_main.ps1'
$repo_root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).ProviderPath
Import-Module (Join-Path $PSScriptRoot 'VbaDevTool.psm1') -Force

function Assert-StringSequence {
    param(
        [string[]]$Actual,

        [string[]]$Expected,

        [Parameter(Mandatory = $true)]
        [string]$Context
    )

    if ($Actual.Count -ne $Expected.Count) {
        throw "$Context expected $($Expected.Count) values but got $($Actual.Count)."
    }
    for ($index = 0; $index -lt $Expected.Count; $index++) {
        if (-not [string]::Equals($Actual[$index], $Expected[$index], [System.StringComparison]::Ordinal)) {
            throw "$Context expected '$($Expected[$index])' at index $index but got '$($Actual[$index])'."
        }
    }
}

function Invoke-ManifestValidator {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ManifestPath,

        [Parameter(Mandatory = $true)]
        [string]$ModulesDirectory,

        [switch]$RequireAllFiles,

        [int]$ExpectedExitCode = 0
    )

    $arguments = @(
        '-NoProfile',
        '-ExecutionPolicy',
        'Bypass',
        '-File',
        $validator_path,
        '-ManifestPath',
        $ManifestPath,
        '-ModulesDirectory',
        $ModulesDirectory
    )
    if ($RequireAllFiles) {
        $arguments += '-RequireAllFiles'
    }

    $previous_error_action_preference = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    $output = & powershell @arguments 2>&1
    $exit_code = $LASTEXITCODE
    $ErrorActionPreference = $previous_error_action_preference
    if ($exit_code -ne $ExpectedExitCode) {
        $output | ForEach-Object { Write-Host $_ }
        throw "Expected exit code $ExpectedExitCode but got $exit_code for manifest '$ManifestPath'."
    }
    if ($ExpectedExitCode -eq 0) {
        $output | ForEach-Object { Write-Host $_ }
    }
}

function Assert-ManifestReadFails {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ManifestPath,

        [Parameter(Mandatory = $true)]
        [string]$ExpectedMessage
    )

    try {
        [void](Read-CommonModulesManifest -ManifestPath $ManifestPath)
    }
    catch {
        if ($_.Exception.Message -notlike "*$ExpectedMessage*") {
            throw "Expected manifest reader error containing '$ExpectedMessage' but got '$($_.Exception.Message)'."
        }
        return
    }

    throw "Expected manifest reader to reject '$ManifestPath'."
}

$source_modules_directory = Join-Path (Join-Path (Join-Path $repo_root 'CommonModules') 'src') 'CommonModules'
$source_manifest_path = Join-Path $source_modules_directory 'common-modules-manifest.tsv'
$distributed_manifest_path = Join-Path (Join-Path $repo_root 'common_modules_repo') 'common-modules-manifest.tsv'
$distributed_modules_directory = Join-Path $repo_root 'common_modules_repo'

$tracer_root = Join-Path ([System.IO.Path]::GetTempPath()) ('common-modules-manifest-tracer-' + [System.Guid]::NewGuid().ToString('N'))
$tracer_modules = Join-Path $tracer_root 'modules'
try {
    New-Item -ItemType Directory -Path $tracer_modules -Force | Out-Null
    Copy-Item -LiteralPath (Join-Path $source_modules_directory 'Lib_Common.bas') -Destination $tracer_modules -Force

    $minimal_manifest = Join-Path $tracer_root 'minimal.tsv'
    $minimal_text = "ModuleFile`tCategories`tDependencies`tRequiredReferences`r`nLib_Common.bas`truntime-baseline`t`t[]`r`n"
    $utf16_le = New-Object System.Text.UnicodeEncoding($false, $true, $true)
    [System.IO.File]::WriteAllText($minimal_manifest, $minimal_text, $utf16_le)
    Invoke-ManifestValidator -ManifestPath $minimal_manifest -ModulesDirectory $tracer_modules

    $json_module_files = @(
        'ArrayObject.cls',
        'Counter.cls',
        'CounterSet.cls',
        'DebugInformation.cls',
        'Enumerator.cls',
        'Fx_Common.bas',
        'IComparable.cls'
    )
    foreach ($json_module_file in $json_module_files) {
        Copy-Item -LiteralPath (Join-Path $source_modules_directory $json_module_file) -Destination $tracer_modules -Force
    }

    $json_manifest = Join-Path $tracer_root 'required-references-json.tsv'
    $json_lines = @(
        "ModuleFile`tCategories`tDependencies`tRequiredReferences",
        "ArrayObject.cls`toptional`t`t[]",
        "Counter.cls`toptional`t`t[`"Single Reference`"]",
        "CounterSet.cls`toptional`t`t[`"First Reference`",`"Second Reference`"]",
        "DebugInformation.cls`toptional`t`t[`"Unicode 参照`"]",
        "Enumerator.cls`toptional`t`t[`"Vendor, Library`"]",
        "Fx_Common.bas`toptional`t`t[`"Quoted \`"Reference\`"`"]",
        "IComparable.cls`toptional`t`t[`"C:\\Vendor\\Library`"]"
    )
    [System.IO.File]::WriteAllText($json_manifest, (($json_lines -join "`r`n") + "`r`n"), $utf16_le)
    Invoke-ManifestValidator -ManifestPath $json_manifest -ModulesDirectory $tracer_modules

    $json_records = (Read-CommonModulesManifest -ManifestPath $json_manifest).RecordsByModule
    Assert-StringSequence -Actual @($json_records['ArrayObject.cls'].RequiredReferences) -Expected @() -Context 'Empty RequiredReferences'
    Assert-StringSequence -Actual @($json_records['Counter.cls'].RequiredReferences) -Expected @('Single Reference') -Context 'Single RequiredReferences'
    Assert-StringSequence -Actual @($json_records['CounterSet.cls'].RequiredReferences) -Expected @('First Reference', 'Second Reference') -Context 'Multiple RequiredReferences'
    Assert-StringSequence -Actual @($json_records['DebugInformation.cls'].RequiredReferences) -Expected @('Unicode 参照') -Context 'Unicode RequiredReferences'
    Assert-StringSequence -Actual @($json_records['Enumerator.cls'].RequiredReferences) -Expected @('Vendor, Library') -Context 'Comma RequiredReferences'
    Assert-StringSequence -Actual @($json_records['Fx_Common.bas'].RequiredReferences) -Expected @('Quoted "Reference"') -Context 'Quoted RequiredReferences'
    Assert-StringSequence -Actual @($json_records['IComparable.cls'].RequiredReferences) -Expected @('C:\Vendor\Library') -Context 'Backslash RequiredReferences'
}
finally {
    if (Test-Path -LiteralPath $tracer_root) {
        Remove-Item -LiteralPath $tracer_root -Recurse -Force
    }
}

Invoke-ManifestValidator -ManifestPath $source_manifest_path -ModulesDirectory $source_modules_directory -RequireAllFiles
Invoke-ManifestValidator -ManifestPath $distributed_manifest_path -ModulesDirectory $distributed_modules_directory -RequireAllFiles

$expected_direct_references = @{
    'ApplicationScreenUpdateManager.cls' = @('Microsoft Excel 16.0 Object Library')
    'CounterSet.cls' = @('Microsoft Scripting Runtime')
    'FileSystemService.cls' = @('Microsoft Excel 16.0 Object Library', 'Microsoft Scripting Runtime')
    'IWorkbookService.cls' = @('Microsoft Excel 16.0 Object Library')
    'IWorksheetService.cls' = @('Microsoft Excel 16.0 Object Library')
    'Lib_Common.bas' = @('Microsoft Excel 16.0 Object Library', 'Microsoft Office 16.0 Object Library')
    'Lib_IPv4.bas' = @('Microsoft VBScript Regular Expressions 5.5')
    'Lib_UnitTest.bas' = @('Microsoft Excel 16.0 Object Library', 'Microsoft VBScript Regular Expressions 5.5')
    'ObjectDictionary.cls' = @('OLE Automation')
    'ObjectList.cls' = @('OLE Automation', 'Microsoft Scripting Runtime')
    'ObjectSet.cls' = @('OLE Automation', 'Microsoft Scripting Runtime')
    'ProgressStatus.cls' = @('Microsoft Excel 16.0 Object Library')
    'TestDoubleBehaviorStore.cls' = @('Microsoft Scripting Runtime')
    'TestDoubleCallRecord.cls' = @('Microsoft Scripting Runtime')
    'WorkbookService.cls' = @('Microsoft Excel 16.0 Object Library')
    'WorkbookServiceTestDouble.cls' = @('Microsoft Excel 16.0 Object Library')
    'WorksheetRangeBounds.cls' = @('OLE Automation')
    'WorksheetService.cls' = @('Microsoft Excel 16.0 Object Library')
    'WorksheetServiceTestDouble.cls' = @('Microsoft Excel 16.0 Object Library')
    'WorksheetVirtualTable.cls' = @('OLE Automation')
}
$source_records = (Read-CommonModulesManifest -ManifestPath $source_manifest_path).Records
foreach ($source_record in $source_records) {
    $expected_references = @()
    if ($expected_direct_references.ContainsKey($source_record.ModuleFile)) {
        $expected_references = @($expected_direct_references[$source_record.ModuleFile])
    }
    Assert-StringSequence -Actual @($source_record.RequiredReferences) -Expected $expected_references -Context "Direct RequiredReferences for $($source_record.ModuleFile)"
}
if (-not (Test-FileContentEqual -LeftPath $source_manifest_path -RightPath $distributed_manifest_path)) {
    throw 'Canonical and distributed CommonModules manifests must contain identical bytes.'
}

$temp_root = Join-Path ([System.IO.Path]::GetTempPath()) ('common-modules-manifest-test-' + [System.Guid]::NewGuid().ToString('N'))
$temp_modules = Join-Path $temp_root 'modules'
$utf16_le = New-Object System.Text.UnicodeEncoding($false, $true, $true)

try {
    New-Item -ItemType Directory -Path $temp_modules -Force | Out-Null
    Copy-Item -LiteralPath (Join-Path $source_modules_directory 'Lib_Common.bas') -Destination $temp_modules -Force

    $unknown_dependency_manifest = Join-Path $temp_root 'unknown-dependency.tsv'
    $unknown_dependency_text = "ModuleFile`tCategories`tDependencies`tRequiredReferences`r`nLib_Common.bas`truntime-baseline`tMissingDependency.cls`t[]`r`n"
    [System.IO.File]::WriteAllText($unknown_dependency_manifest, $unknown_dependency_text, $utf16_le)
    Invoke-ManifestValidator -ManifestPath $unknown_dependency_manifest -ModulesDirectory $temp_modules -ExpectedExitCode 1
    Assert-ManifestReadFails -ManifestPath $unknown_dependency_manifest -ExpectedMessage 'unknown dependency'

    $malformed_manifest = Join-Path $temp_root 'malformed.tsv'
    $malformed_text = "ModuleFile`tCategories`tDependencies`tRequiredReferences`r`nLib_Common.bas`truntime-baseline`t`r`n"
    [System.IO.File]::WriteAllText($malformed_manifest, $malformed_text, $utf16_le)
    Invoke-ManifestValidator -ManifestPath $malformed_manifest -ModulesDirectory $temp_modules -ExpectedExitCode 1

    $extra_column_manifest = Join-Path $temp_root 'extra-column.tsv'
    $extra_column_text = "ModuleFile`tCategories`tDependencies`tRequiredReferences`r`nLib_Common.bas`truntime-baseline`t`t[]`textra`r`n"
    [System.IO.File]::WriteAllText($extra_column_manifest, $extra_column_text, $utf16_le)
    Assert-ManifestReadFails -ManifestPath $extra_column_manifest -ExpectedMessage 'exactly 4 tab-separated columns'

    $tabbed_comment_manifest = Join-Path $temp_root 'tabbed-comment.tsv'
    $tabbed_comment_text = "# invalid`tcomment`r`nModuleFile`tCategories`tDependencies`tRequiredReferences`r`nLib_Common.bas`truntime-baseline`t`t[]`r`n"
    [System.IO.File]::WriteAllText($tabbed_comment_manifest, $tabbed_comment_text, $utf16_le)
    Invoke-ManifestValidator -ManifestPath $tabbed_comment_manifest -ModulesDirectory $temp_modules -ExpectedExitCode 1

    $trailing_comment_whitespace_manifest = Join-Path $temp_root 'trailing-comment-whitespace.tsv'
    $trailing_comment_whitespace_text = "# invalid comment `r`nModuleFile`tCategories`tDependencies`tRequiredReferences`r`nLib_Common.bas`truntime-baseline`t`t[]`r`n"
    [System.IO.File]::WriteAllText($trailing_comment_whitespace_manifest, $trailing_comment_whitespace_text, $utf16_le)
    Invoke-ManifestValidator -ManifestPath $trailing_comment_whitespace_manifest -ModulesDirectory $temp_modules -ExpectedExitCode 1

    $rooted_module_manifest = Join-Path $temp_root 'rooted-module.tsv'
    $rooted_module_text = "ModuleFile`tCategories`tDependencies`tRequiredReferences`r`nC:\Lib_Common.bas`truntime-baseline`t`t[]`r`n"
    [System.IO.File]::WriteAllText($rooted_module_manifest, $rooted_module_text, $utf16_le)
    Assert-ManifestReadFails -ManifestPath $rooted_module_manifest -ExpectedMessage 'invalid ModuleFile'

    $nested_module_manifest = Join-Path $temp_root 'nested-module.tsv'
    $nested_module_text = "ModuleFile`tCategories`tDependencies`tRequiredReferences`r`nsub/Lib_Common.bas`truntime-baseline`t`t[]`r`n"
    [System.IO.File]::WriteAllText($nested_module_manifest, $nested_module_text, $utf16_le)
    Assert-ManifestReadFails -ManifestPath $nested_module_manifest -ExpectedMessage 'invalid ModuleFile'

    $backslash_module_manifest = Join-Path $temp_root 'backslash-module.tsv'
    $backslash_module_text = "ModuleFile`tCategories`tDependencies`tRequiredReferences`r`nsub\Lib_Common.bas`truntime-baseline`t`t[]`r`n"
    [System.IO.File]::WriteAllText($backslash_module_manifest, $backslash_module_text, $utf16_le)
    Assert-ManifestReadFails -ManifestPath $backslash_module_manifest -ExpectedMessage 'invalid ModuleFile'

    $uppercase_extension_manifest = Join-Path $temp_root 'uppercase-extension.tsv'
    $uppercase_extension_text = "ModuleFile`tCategories`tDependencies`tRequiredReferences`r`nLib_Common.BAS`truntime-baseline`t`t[]`r`n"
    [System.IO.File]::WriteAllText($uppercase_extension_manifest, $uppercase_extension_text, $utf16_le)
    Assert-ManifestReadFails -ManifestPath $uppercase_extension_manifest -ExpectedMessage 'invalid ModuleFile'

    $ignorable_extension_manifest = Join-Path $temp_root 'ignorable-extension.tsv'
    $ignorable_extension_text = "ModuleFile`tCategories`tDependencies`tRequiredReferences`r`nLib_Common.b$([char]0x00AD)as`truntime-baseline`t`t[]`r`n"
    [System.IO.File]::WriteAllText($ignorable_extension_manifest, $ignorable_extension_text, $utf16_le)
    Assert-ManifestReadFails -ManifestPath $ignorable_extension_manifest -ExpectedMessage 'invalid ModuleFile'

    $padded_module_manifest = Join-Path $temp_root 'padded-module.tsv'
    $padded_module_text = "ModuleFile`tCategories`tDependencies`tRequiredReferences`r`n Lib_Common.bas`truntime-baseline`t`t[]`r`n"
    [System.IO.File]::WriteAllText($padded_module_manifest, $padded_module_text, $utf16_le)
    Assert-ManifestReadFails -ManifestPath $padded_module_manifest -ExpectedMessage 'invalid ModuleFile'

    $trailing_padded_module_manifest = Join-Path $temp_root 'trailing-padded-module.tsv'
    $trailing_padded_module_text = "ModuleFile`tCategories`tDependencies`tRequiredReferences`r`nLib_Common.bas `truntime-baseline`t`t[]`r`n"
    [System.IO.File]::WriteAllText($trailing_padded_module_manifest, $trailing_padded_module_text, $utf16_le)
    Assert-ManifestReadFails -ManifestPath $trailing_padded_module_manifest -ExpectedMessage 'invalid ModuleFile'

    $comma_module_manifest = Join-Path $temp_root 'comma-module.tsv'
    $comma_module_text = "ModuleFile`tCategories`tDependencies`tRequiredReferences`r`nLib,Common.bas`truntime-baseline`t`t[]`r`n"
    [System.IO.File]::WriteAllText($comma_module_manifest, $comma_module_text, $utf16_le)
    Assert-ManifestReadFails -ManifestPath $comma_module_manifest -ExpectedMessage 'invalid ModuleFile'

    $control_module_manifest = Join-Path $temp_root 'control-module.tsv'
    $control_module_text = "ModuleFile`tCategories`tDependencies`tRequiredReferences`r`nLib_$([char]0)Common.bas`truntime-baseline`t`t[]`r`n"
    [System.IO.File]::WriteAllText($control_module_manifest, $control_module_text, $utf16_le)
    Assert-ManifestReadFails -ManifestPath $control_module_manifest -ExpectedMessage 'invalid ModuleFile'

    $invalid_filename_character_manifest = Join-Path $temp_root 'invalid-filename-character.tsv'
    $invalid_filename_character_text = "ModuleFile`tCategories`tDependencies`tRequiredReferences`r`nLib:Common.bas`truntime-baseline`t`t[]`r`n"
    [System.IO.File]::WriteAllText($invalid_filename_character_manifest, $invalid_filename_character_text, $utf16_le)
    Assert-ManifestReadFails -ManifestPath $invalid_filename_character_manifest -ExpectedMessage 'invalid ModuleFile'

    $dot_segment_manifest = Join-Path $temp_root 'dot-segment.tsv'
    $dot_segment_text = "ModuleFile`tCategories`tDependencies`tRequiredReferences`r`n..bas`truntime-baseline`t`t[]`r`n"
    [System.IO.File]::WriteAllText($dot_segment_manifest, $dot_segment_text, $utf16_le)
    Assert-ManifestReadFails -ManifestPath $dot_segment_manifest -ExpectedMessage 'invalid ModuleFile'

    $padded_module_name_manifest = Join-Path $temp_root 'padded-module-name.tsv'
    $padded_module_name_text = "ModuleFile`tCategories`tDependencies`tRequiredReferences`r`nLib_Common .bas`truntime-baseline`t`t[]`r`n"
    [System.IO.File]::WriteAllText($padded_module_name_manifest, $padded_module_name_text, $utf16_le)
    Assert-ManifestReadFails -ManifestPath $padded_module_name_manifest -ExpectedMessage 'invalid ModuleFile'

    $embedded_dot_manifest = Join-Path $temp_root 'embedded-dot-module.tsv'
    $embedded_dot_text = "ModuleFile`tCategories`tDependencies`tRequiredReferences`r`nLib..Common.bas`truntime-baseline`t`t[]`r`n"
    [System.IO.File]::WriteAllText($embedded_dot_manifest, $embedded_dot_text, $utf16_le)
    Assert-ManifestReadFails -ManifestPath $embedded_dot_manifest -ExpectedMessage 'invalid ModuleFile'

    $invalid_category_manifest = Join-Path $temp_root 'invalid-category.tsv'
    $invalid_category_text = "ModuleFile`tCategories`tDependencies`tRequiredReferences`r`nLib_Common.bas`tRuntime-baseline`t`t[]`r`n"
    [System.IO.File]::WriteAllText($invalid_category_manifest, $invalid_category_text, $utf16_le)
    Assert-ManifestReadFails -ManifestPath $invalid_category_manifest -ExpectedMessage 'invalid Categories'

    $ignorable_category_manifest = Join-Path $temp_root 'ignorable-category.tsv'
    $ignorable_category_text = "ModuleFile`tCategories`tDependencies`tRequiredReferences`r`nLib_Common.bas`truntime-$([char]0)baseline`t`t[]`r`n"
    [System.IO.File]::WriteAllText($ignorable_category_manifest, $ignorable_category_text, $utf16_le)
    Assert-ManifestReadFails -ManifestPath $ignorable_category_manifest -ExpectedMessage 'invalid Categories'

    $invalid_category_values = @(
        ' runtime-baseline',
        'runtime-baseline ',
        'public-udf,optional',
        'runtime-baseline,optional',
        'public-udf',
        'optional,runtime-baseline,public-udf'
    )
    for ($category_index = 0; $category_index -lt $invalid_category_values.Count; $category_index++) {
        $invalid_category_value_manifest = Join-Path $temp_root "invalid-category-$category_index.tsv"
        $invalid_category_value_text = "ModuleFile`tCategories`tDependencies`tRequiredReferences`r`nLib_Common.bas`t$($invalid_category_values[$category_index])`t`t[]`r`n"
        [System.IO.File]::WriteAllText($invalid_category_value_manifest, $invalid_category_value_text, $utf16_le)
        Assert-ManifestReadFails -ManifestPath $invalid_category_value_manifest -ExpectedMessage 'invalid Categories'
    }

    $dependency_case_manifest = Join-Path $temp_root 'dependency-case.tsv'
    $dependency_case_text = "ModuleFile`tCategories`tDependencies`tRequiredReferences`r`nLib_Common.bas`truntime-baseline`tarrayobject.cls`t[]`r`nArrayObject.cls`truntime-baseline`t`t[]`r`n"
    [System.IO.File]::WriteAllText($dependency_case_manifest, $dependency_case_text, $utf16_le)
    Assert-ManifestReadFails -ManifestPath $dependency_case_manifest -ExpectedMessage 'exact ModuleFile spelling'

    $dependency_whitespace_manifest = Join-Path $temp_root 'dependency-whitespace.tsv'
    $dependency_whitespace_text = "ModuleFile`tCategories`tDependencies`tRequiredReferences`r`nLib_Common.bas`truntime-baseline`tArrayObject.cls, Counter.cls`t[]`r`nArrayObject.cls`truntime-baseline`t`t[]`r`nCounter.cls`truntime-baseline`t`t[]`r`n"
    [System.IO.File]::WriteAllText($dependency_whitespace_manifest, $dependency_whitespace_text, $utf16_le)
    Assert-ManifestReadFails -ManifestPath $dependency_whitespace_manifest -ExpectedMessage 'whitespace-free Dependencies'

    $ignorable_dependency_manifest = Join-Path $temp_root 'ignorable-dependency.tsv'
    $ignorable_dependency_text = "ModuleFile`tCategories`tDependencies`tRequiredReferences`r`nLib_Common.bas`truntime-baseline`t$([char]0x00AD)`t[]`r`n"
    [System.IO.File]::WriteAllText($ignorable_dependency_manifest, $ignorable_dependency_text, $utf16_le)
    Assert-ManifestReadFails -ManifestPath $ignorable_dependency_manifest -ExpectedMessage 'unknown dependency'

    $empty_dependency_item_manifest = Join-Path $temp_root 'empty-dependency-item.tsv'
    $empty_dependency_item_text = "ModuleFile`tCategories`tDependencies`tRequiredReferences`r`nLib_Common.bas`truntime-baseline`tArrayObject.cls,,Counter.cls`t[]`r`nArrayObject.cls`truntime-baseline`t`t[]`r`nCounter.cls`truntime-baseline`t`t[]`r`n"
    [System.IO.File]::WriteAllText($empty_dependency_item_manifest, $empty_dependency_item_text, $utf16_le)
    Assert-ManifestReadFails -ManifestPath $empty_dependency_item_manifest -ExpectedMessage 'empty dependency'

    $duplicate_dependency_manifest = Join-Path $temp_root 'duplicate-dependency.tsv'
    $duplicate_dependency_text = "ModuleFile`tCategories`tDependencies`tRequiredReferences`r`nLib_Common.bas`truntime-baseline`tArrayObject.cls,ArrayObject.cls`t[]`r`nArrayObject.cls`truntime-baseline`t`t[]`r`n"
    [System.IO.File]::WriteAllText($duplicate_dependency_manifest, $duplicate_dependency_text, $utf16_le)
    Assert-ManifestReadFails -ManifestPath $duplicate_dependency_manifest -ExpectedMessage 'duplicate dependency'

    $case_duplicate_dependency_manifest = Join-Path $temp_root 'case-duplicate-dependency.tsv'
    $case_duplicate_dependency_text = "ModuleFile`tCategories`tDependencies`tRequiredReferences`r`nLib_Common.bas`truntime-baseline`tArrayObject.cls,arrayobject.cls`t[]`r`nArrayObject.cls`truntime-baseline`t`t[]`r`n"
    [System.IO.File]::WriteAllText($case_duplicate_dependency_manifest, $case_duplicate_dependency_text, $utf16_le)
    Assert-ManifestReadFails -ManifestPath $case_duplicate_dependency_manifest -ExpectedMessage 'duplicate dependency'

    $self_dependency_manifest = Join-Path $temp_root 'self-dependency.tsv'
    $self_dependency_text = "ModuleFile`tCategories`tDependencies`tRequiredReferences`r`nLib_Common.bas`truntime-baseline`tLib_Common.bas`t[]`r`n"
    [System.IO.File]::WriteAllText($self_dependency_manifest, $self_dependency_text, $utf16_le)
    Assert-ManifestReadFails -ManifestPath $self_dependency_manifest -ExpectedMessage 'self dependency'

    $runtime_to_test_manifest = Join-Path $temp_root 'runtime-to-test.tsv'
    $runtime_to_test_text = "ModuleFile`tCategories`tDependencies`tRequiredReferences`r`nLib_Common.bas`truntime-baseline`tLib_UnitTest.bas`t[]`r`nLib_UnitTest.bas`ttest-foundation`t`t[]`r`n"
    [System.IO.File]::WriteAllText($runtime_to_test_manifest, $runtime_to_test_text, $utf16_le)
    Assert-ManifestReadFails -ManifestPath $runtime_to_test_manifest -ExpectedMessage 'runtime-role'

    $additional_runtime_categories = @('runtime-baseline,public-udf', 'optional', 'optional,public-udf')
    for ($runtime_category_index = 0; $runtime_category_index -lt $additional_runtime_categories.Count; $runtime_category_index++) {
        $runtime_role_manifest = Join-Path $temp_root "runtime-role-to-test-$runtime_category_index.tsv"
        $target_test_category = if ($runtime_category_index -eq 1) { 'test-double' } else { 'test-foundation' }
        $runtime_role_text = "ModuleFile`tCategories`tDependencies`tRequiredReferences`r`nLib_Common.bas`t$($additional_runtime_categories[$runtime_category_index])`tLib_UnitTest.bas`t[]`r`nLib_UnitTest.bas`t$target_test_category`t`t[]`r`n"
        [System.IO.File]::WriteAllText($runtime_role_manifest, $runtime_role_text, $utf16_le)
        Assert-ManifestReadFails -ManifestPath $runtime_role_manifest -ExpectedMessage 'runtime-role'
    }

    $trailing_comma_json_manifest = Join-Path $temp_root 'trailing-comma-json.tsv'
    $trailing_comma_json_text = "ModuleFile`tCategories`tDependencies`tRequiredReferences`r`nLib_Common.bas`truntime-baseline`t`t[`"Reference`",]`r`n"
    [System.IO.File]::WriteAllText($trailing_comma_json_manifest, $trailing_comma_json_text, $utf16_le)
    Assert-ManifestReadFails -ManifestPath $trailing_comma_json_manifest -ExpectedMessage 'malformed RequiredReferences JSON'

    $commented_json_manifest = Join-Path $temp_root 'commented-json.tsv'
    $commented_json_text = "ModuleFile`tCategories`tDependencies`tRequiredReferences`r`nLib_Common.bas`truntime-baseline`t`t[/*invalid*/`"Reference`"]`r`n"
    [System.IO.File]::WriteAllText($commented_json_manifest, $commented_json_text, $utf16_le)
    Assert-ManifestReadFails -ManifestPath $commented_json_manifest -ExpectedMessage 'malformed RequiredReferences JSON'

    $empty_reference_manifest = Join-Path $temp_root 'empty-reference.tsv'
    $empty_reference_text = "ModuleFile`tCategories`tDependencies`tRequiredReferences`r`nLib_Common.bas`truntime-baseline`t`t[`"`"]`r`n"
    [System.IO.File]::WriteAllText($empty_reference_manifest, $empty_reference_text, $utf16_le)
    Assert-ManifestReadFails -ManifestPath $empty_reference_manifest -ExpectedMessage 'nonempty RequiredReferences'

    $padded_reference_manifest = Join-Path $temp_root 'padded-reference.tsv'
    $padded_reference_text = "ModuleFile`tCategories`tDependencies`tRequiredReferences`r`nLib_Common.bas`truntime-baseline`t`t[`" Reference`"]`r`n"
    [System.IO.File]::WriteAllText($padded_reference_manifest, $padded_reference_text, $utf16_le)
    Assert-ManifestReadFails -ManifestPath $padded_reference_manifest -ExpectedMessage 'already trimmed RequiredReferences'

    $duplicate_reference_manifest = Join-Path $temp_root 'duplicate-reference.tsv'
    $duplicate_reference_text = "ModuleFile`tCategories`tDependencies`tRequiredReferences`r`nLib_Common.bas`truntime-baseline`t`t[`"Reference`",`"reference`"]`r`n"
    [System.IO.File]::WriteAllText($duplicate_reference_manifest, $duplicate_reference_text, $utf16_le)
    Assert-ManifestReadFails -ManifestPath $duplicate_reference_manifest -ExpectedMessage 'duplicate RequiredReferences'

    $vba_reference_manifest = Join-Path $temp_root 'vba-reference.tsv'
    $vba_reference_text = "ModuleFile`tCategories`tDependencies`tRequiredReferences`r`nLib_Common.bas`truntime-baseline`t`t[`"visual basic for applications`"]`r`n"
    [System.IO.File]::WriteAllText($vba_reference_manifest, $vba_reference_text, $utf16_le)
    Assert-ManifestReadFails -ManifestPath $vba_reference_manifest -ExpectedMessage 'always-active VBA standard library'

    $valid_manifest_text = "ModuleFile`tCategories`tDependencies`tRequiredReferences`r`nLib_Common.bas`truntime-baseline`t`t[]`r`n"

    $utf8_manifest = Join-Path $temp_root 'utf8.tsv'
    [System.IO.File]::WriteAllText($utf8_manifest, $valid_manifest_text, (New-Object System.Text.UTF8Encoding($true)))
    Assert-ManifestReadFails -ManifestPath $utf8_manifest -ExpectedMessage 'UTF-16LE with a BOM'

    $utf16_be_manifest = Join-Path $temp_root 'utf16-be.tsv'
    [System.IO.File]::WriteAllText($utf16_be_manifest, $valid_manifest_text, (New-Object System.Text.UnicodeEncoding($true, $true, $true)))
    Assert-ManifestReadFails -ManifestPath $utf16_be_manifest -ExpectedMessage 'UTF-16LE with a BOM'

    $missing_bom_manifest = Join-Path $temp_root 'missing-bom.tsv'
    [System.IO.File]::WriteAllText($missing_bom_manifest, $valid_manifest_text, (New-Object System.Text.UnicodeEncoding($false, $false, $true)))
    Assert-ManifestReadFails -ManifestPath $missing_bom_manifest -ExpectedMessage 'UTF-16LE with a BOM'

    $utf8_without_bom_manifest = Join-Path $temp_root 'utf8-without-bom.tsv'
    [System.IO.File]::WriteAllText($utf8_without_bom_manifest, $valid_manifest_text, (New-Object System.Text.UTF8Encoding($false)))
    Assert-ManifestReadFails -ManifestPath $utf8_without_bom_manifest -ExpectedMessage 'UTF-16LE with a BOM'

    $shift_jis_manifest = Join-Path $temp_root 'shift-jis.tsv'
    [System.IO.File]::WriteAllText($shift_jis_manifest, $valid_manifest_text, [System.Text.Encoding]::GetEncoding(932))
    Assert-ManifestReadFails -ManifestPath $shift_jis_manifest -ExpectedMessage 'UTF-16LE with a BOM'

    $invalid_utf16_manifest = Join-Path $temp_root 'invalid-utf16.tsv'
    [System.IO.File]::WriteAllBytes($invalid_utf16_manifest, [byte[]](0xFF, 0xFE, 0x00, 0xD8, 0x0D, 0x00, 0x0A, 0x00))
    Assert-ManifestReadFails -ManifestPath $invalid_utf16_manifest -ExpectedMessage 'invalid UTF-16LE text'

    $odd_utf16_byte_count_manifest = Join-Path $temp_root 'odd-utf16-byte-count.tsv'
    $valid_manifest_bytes = [byte[]]($utf16_le.GetPreamble() + $utf16_le.GetBytes($valid_manifest_text))
    [System.IO.File]::WriteAllBytes($odd_utf16_byte_count_manifest, $valid_manifest_bytes[0..($valid_manifest_bytes.Length - 2)])
    Assert-ManifestReadFails -ManifestPath $odd_utf16_byte_count_manifest -ExpectedMessage 'invalid UTF-16LE byte count'

    $invalid_physical_grammar_cases = @(
        [pscustomobject]@{
            Name = 'mixed-line-endings.tsv'
            Text = "ModuleFile`tCategories`tDependencies`tRequiredReferences`nLib_Common.bas`truntime-baseline`t`t[]`r`n"
            Error = 'CRLF line endings throughout'
        },
        [pscustomobject]@{
            Name = 'lone-cr-line-ending.tsv'
            Text = "ModuleFile`tCategories`tDependencies`tRequiredReferences`rLib_Common.bas`truntime-baseline`t`t[]`r`n"
            Error = 'CRLF line endings throughout'
        },
        [pscustomobject]@{
            Name = 'missing-final-crlf.tsv'
            Text = "ModuleFile`tCategories`tDependencies`tRequiredReferences`r`nLib_Common.bas`truntime-baseline`t`t[]"
            Error = 'exactly one CRLF'
        },
        [pscustomobject]@{
            Name = 'double-final-crlf.tsv'
            Text = "ModuleFile`tCategories`tDependencies`tRequiredReferences`r`nLib_Common.bas`truntime-baseline`t`t[]`r`n`r`n"
            Error = 'exactly one CRLF'
        },
        [pscustomobject]@{
            Name = 'blank-line.tsv'
            Text = "ModuleFile`tCategories`tDependencies`tRequiredReferences`r`n`r`nLib_Common.bas`truntime-baseline`t`t[]`r`n"
            Error = 'exactly 4 tab-separated columns'
        },
        [pscustomobject]@{
            Name = 'indented-comment.tsv'
            Text = " # invalid`r`nModuleFile`tCategories`tDependencies`tRequiredReferences`r`nLib_Common.bas`truntime-baseline`t`t[]`r`n"
            Error = 'invalid header'
        },
        [pscustomobject]@{
            Name = 'ignorable-prefix-comment.tsv'
            Text = "$([char]0x00AD)# invalid`r`nModuleFile`tCategories`tDependencies`tRequiredReferences`r`nLib_Common.bas`truntime-baseline`t`t[]`r`n"
            Error = 'invalid header'
        },
        [pscustomobject]@{
            Name = 'comment-after-header.tsv'
            Text = "ModuleFile`tCategories`tDependencies`tRequiredReferences`r`n# invalid placement`r`nLib_Common.bas`truntime-baseline`t`t[]`r`n"
            Error = 'exactly 4 tab-separated columns'
        },
        [pscustomobject]@{
            Name = 'different-case-header.tsv'
            Text = "moduleFile`tCategories`tDependencies`tRequiredReferences`r`nLib_Common.bas`truntime-baseline`t`t[]`r`n"
            Error = 'invalid header'
        },
        [pscustomobject]@{
            Name = 'ignorable-code-unit-header.tsv'
            Text = "Module$([char]0x00AD)File`tCategories`tDependencies`tRequiredReferences`r`nLib_Common.bas`truntime-baseline`t`t[]`r`n"
            Error = 'invalid header'
        },
        [pscustomobject]@{
            Name = 'duplicate-header.tsv'
            Text = "ModuleFile`tCategories`tDependencies`tRequiredReferences`r`nLib_Common.bas`truntime-baseline`t`t[]`r`nModuleFile`tCategories`tDependencies`tRequiredReferences`r`n"
            Error = 'duplicate header'
        },
        [pscustomobject]@{
            Name = 'inline-comment.tsv'
            Text = "ModuleFile`tCategories`tDependencies`tRequiredReferences`r`nLib_Common.bas`truntime-baseline`t`t[] # invalid`r`n"
            Error = 'malformed RequiredReferences JSON'
        }
    )
    foreach ($invalid_case in $invalid_physical_grammar_cases) {
        $invalid_case_path = Join-Path $temp_root $invalid_case.Name
        [System.IO.File]::WriteAllText($invalid_case_path, $invalid_case.Text, $utf16_le)
        Assert-ManifestReadFails -ManifestPath $invalid_case_path -ExpectedMessage $invalid_case.Error
    }

    $duplicate_module_manifest = Join-Path $temp_root 'duplicate-module.tsv'
    $duplicate_module_text = "ModuleFile`tCategories`tDependencies`tRequiredReferences`r`nLib_Common.bas`truntime-baseline`t`t[]`r`nlib_common.bas`truntime-baseline`t`t[]`r`n"
    [System.IO.File]::WriteAllText($duplicate_module_manifest, $duplicate_module_text, $utf16_le)
    Assert-ManifestReadFails -ManifestPath $duplicate_module_manifest -ExpectedMessage 'duplicates ModuleFile'

    $category_contract_manifest = Join-Path $temp_root 'category-contract.tsv'
    $category_contract_lines = @(
        "ModuleFile`tCategories`tDependencies`tRequiredReferences",
        "ArrayObject.cls`truntime-baseline`t`t[]",
        "Fx_Common.bas`truntime-baseline,public-udf`t`t[]",
        "Lib_UnitTest.bas`ttest-foundation`t`t[]",
        "Counter.cls`toptional`t`t[]",
        "Lib_IPv4.bas`toptional,public-udf`t`t[]",
        "FileSystemServiceTestDouble.cls`ttest-double`t`t[]"
    )
    [System.IO.File]::WriteAllText($category_contract_manifest, (($category_contract_lines -join "`r`n") + "`r`n"), $utf16_le)
    $category_records = (Read-CommonModulesManifest -ManifestPath $category_contract_manifest).Records
    Assert-StringSequence -Actual @($category_records | ForEach-Object { $_.Categories }) -Expected @(
        'runtime-baseline',
        'runtime-baseline,public-udf',
        'test-foundation',
        'optional',
        'optional,public-udf',
        'test-double'
    ) -Context 'Canonical Categories'

    $unicode_form_file = "Caf$([char]0x00E9).frm"
    $valid_dependency_contract_manifest = Join-Path $temp_root 'valid-dependency-contract.tsv'
    $valid_dependency_contract_lines = @(
        '# Unicode comment: 依存関係',
        '# Declaration order is significant',
        "ModuleFile`tCategories`tDependencies`tRequiredReferences",
        "TestDoubleBehaviorStore.cls`ttest-double`tTestDoubleCallRecord.cls,Counter.cls`t[]",
        "Counter.cls`toptional`t`t[]",
        "TestDoubleCallRecord.cls`ttest-foundation`t$unicode_form_file`t[]",
        "$unicode_form_file`truntime-baseline`t`t[]"
    )
    [System.IO.File]::WriteAllText($valid_dependency_contract_manifest, (($valid_dependency_contract_lines -join "`r`n") + "`r`n"), $utf16_le)
    $valid_dependency_contract = Read-CommonModulesManifest -ManifestPath $valid_dependency_contract_manifest
    Assert-StringSequence -Actual @($valid_dependency_contract.Records | ForEach-Object { $_.ModuleFile }) -Expected @(
        'TestDoubleBehaviorStore.cls',
        'Counter.cls',
        'TestDoubleCallRecord.cls',
        $unicode_form_file
    ) -Context 'ModuleFile declaration order and ordinal spelling'
    Assert-StringSequence -Actual @($valid_dependency_contract.RecordsByModule['TestDoubleBehaviorStore.cls'].Dependencies) -Expected @(
        'TestDoubleCallRecord.cls',
        'Counter.cls'
    ) -Context 'Dependency declaration order'
    Assert-StringSequence -Actual @($valid_dependency_contract.RecordsByModule['TestDoubleCallRecord.cls'].Dependencies) -Expected @($unicode_form_file) -Context 'Test-to-runtime dependency'

    $dependency_cycle_manifest = Join-Path $temp_root 'dependency-cycle.tsv'
    $dependency_cycle_text = "ModuleFile`tCategories`tDependencies`tRequiredReferences`r`nCounter.cls`toptional`tCounterSet.cls`t[]`r`nCounterSet.cls`toptional`tCounter.cls`t[]`r`n"
    [System.IO.File]::WriteAllText($dependency_cycle_manifest, $dependency_cycle_text, $utf16_le)
    $dependency_cycle_records = (Read-CommonModulesManifest -ManifestPath $dependency_cycle_manifest).RecordsByModule
    Assert-StringSequence -Actual @($dependency_cycle_records['Counter.cls'].Dependencies) -Expected @('CounterSet.cls') -Context 'Dependency cycle first edge'
    Assert-StringSequence -Actual @($dependency_cycle_records['CounterSet.cls'].Dependencies) -Expected @('Counter.cls') -Context 'Dependency cycle second edge'
}
finally {
    if (Test-Path -LiteralPath $temp_root) {
        Remove-Item -LiteralPath $temp_root -Recurse -Force
    }
}

Write-Host 'CommonModules manifest tests passed.'
