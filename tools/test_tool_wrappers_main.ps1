$ErrorActionPreference = 'Stop'
$InformationPreference = 'Continue'

$repo_root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).ProviderPath
$tools_root = Join-Path $repo_root 'tools'

function Write-TestFileUtf8 {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [Parameter(Mandatory = $true)]
        [string]$Content
    )

    $directory = Split-Path -Parent $Path
    if (-not (Test-Path -LiteralPath $directory -PathType Container)) {
        New-Item -ItemType Directory -Path $directory -Force | Out-Null
    }
    [System.IO.File]::WriteAllText($Path, $Content, [System.Text.UTF8Encoding]::new($false))
}

function Write-TestFileSjis {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [Parameter(Mandatory = $true)]
        [string]$Content
    )

    $directory = Split-Path -Parent $Path
    if (-not (Test-Path -LiteralPath $directory -PathType Container)) {
        New-Item -ItemType Directory -Path $directory -Force | Out-Null
    }
    [System.IO.File]::WriteAllText($Path, $Content, [System.Text.Encoding]::GetEncoding(932))
}

function Test-True {
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

function Test-Equal {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Expected,
        [Parameter(Mandatory = $true)]
        [object]$Actual,
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    if ($Expected -ne $Actual) {
        throw "$Message Expected '$Expected' but got '$Actual'."
    }
}

function Test-PathEqual {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Expected,
        [Parameter(Mandatory = $true)]
        [string]$Actual,
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    $expected_full_name = [System.IO.Path]::GetFullPath($Expected).TrimEnd([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar)
    $actual_full_name = [System.IO.Path]::GetFullPath($Actual).TrimEnd([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar)
    Test-True -Condition ([string]::Equals($expected_full_name, $actual_full_name, [System.StringComparison]::OrdinalIgnoreCase)) -Message "$Message Expected '$expected_full_name' but got '$actual_full_name'."
}

function Invoke-ExpectedFailure {
    param(
        [Parameter(Mandatory = $true)]
        [scriptblock]$Script,
        [Parameter(Mandatory = $true)]
        [string]$ExpectedMessage
    )

    try {
        & $Script
    }
    catch {
        if ($_.Exception.Message -notlike "*$ExpectedMessage*") {
            throw "Expected failure containing '$ExpectedMessage' but got '$($_.Exception.Message)'."
        }
        return
    }

    throw "Expected failure containing '$ExpectedMessage' but the command succeeded."
}

function New-TestProject {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Root,
        [Parameter(Mandatory = $true)]
        [string]$ProjectName,
        [Parameter(Mandatory = $true)]
        [string]$DocumentName
    )

    $source_dir = Join-Path (Join-Path $Root 'src') $DocumentName
    New-Item -ItemType Directory -Path $source_dir -Force | Out-Null
    Write-TestFileUtf8 -Path (Join-Path $source_dir "$DocumentName.xlsm") -Content 'workbook'
    $manifest = [ordered]@{
        schemaVersion = 1
        projectName = $ProjectName
        primaryDocument = $DocumentName
        documents = [ordered]@{
            $DocumentName = [ordered]@{
                kind = 'excel'
                sourcePath = "src/$DocumentName"
                templatePath = "src/$DocumentName/$DocumentName.xlsm"
                binPath = "bin/$DocumentName/$DocumentName.xlsm"
                publishPath = "publish/$DocumentName/$DocumentName.xlsm"
                commonModules = @()
                references = @()
            }
        }
        commonModulesRepository = '../common_modules_repo'
    }
    $manifest | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $Root 'project.json') -Encoding UTF8

    return [pscustomobject]@{
        Root = $Root
        DocumentName = $DocumentName
        SourceSetPath = $source_dir
        WorkbookPath = Join-Path $source_dir "$DocumentName.xlsm"
    }
}

function New-TestVbaDevToolLayout {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Root,
        [switch]$WithBundledExecutable
    )

    $repository_root = Join-Path $Root 'xls-common-modules'
    $layout_tools_root = Join-Path $repository_root 'tools'
    New-Item -ItemType Directory -Path $layout_tools_root -Force | Out-Null
    $module_path = Join-Path $layout_tools_root 'VbaDevTool.psm1'
    Copy-Item -LiteralPath (Join-Path $tools_root 'VbaDevTool.psm1') -Destination $module_path -Force

    $bundled_executable = Join-Path $layout_tools_root 'vba-dev\vba-dev.exe'
    if ($WithBundledExecutable) {
        Write-TestFileUtf8 -Path $bundled_executable -Content 'fake bundled vba-dev'
    }

    return [pscustomobject]@{
        WorkspaceRoot = $Root
        RepositoryRoot = $repository_root
        ToolsRoot = $layout_tools_root
        ModulePath = $module_path
        BundledExecutable = $bundled_executable
    }
}

function Import-TestVbaDevToolLayout {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Layout
    )

    Remove-Module -Name VbaDevTool -Force -ErrorAction SilentlyContinue
    Import-Module $Layout.ModulePath -Force
}

$temp_root = Join-Path ([System.IO.Path]::GetTempPath()) ('xls-common-tools-test-' + [System.Guid]::NewGuid().ToString('N'))
$previous_vba_dev_exe = $env:VBA_DEV_EXE
$previous_vba_dev_calls = $env:VBA_DEV_CALLS
$previous_path = $env:PATH

try {
    New-Item -ItemType Directory -Path $temp_root -Force | Out-Null

    $resolver_root = Join-Path $temp_root 'resolver'
    $env_executable = Join-Path $resolver_root 'env\vba-dev.exe'
    Write-TestFileUtf8 -Path $env_executable -Content 'fake environment vba-dev'
    $path_root = Join-Path $resolver_root 'path'
    Write-TestFileUtf8 -Path (Join-Path $path_root 'vba-dev.exe') -Content 'fake PATH vba-dev'
    $empty_path_root = Join-Path $resolver_root 'empty-path'
    New-Item -ItemType Directory -Path $empty_path_root -Force | Out-Null

    $env_wins_layout = New-TestVbaDevToolLayout -Root (Join-Path $resolver_root 'env-wins') -WithBundledExecutable
    Import-TestVbaDevToolLayout -Layout $env_wins_layout
    $env:VBA_DEV_EXE = $env_executable
    Test-PathEqual -Expected $env_executable -Actual (Resolve-VbaDevExecutable) -Message 'Resolve-VbaDevExecutable should prefer a valid VBA_DEV_EXE.'

    $invalid_env_layout = New-TestVbaDevToolLayout -Root (Join-Path $resolver_root 'invalid-env-falls-back') -WithBundledExecutable
    Import-TestVbaDevToolLayout -Layout $invalid_env_layout
    $missing_env_executable = Join-Path $resolver_root 'missing-env\vba-dev.exe'
    $env:VBA_DEV_EXE = $missing_env_executable
    Test-PathEqual -Expected $invalid_env_layout.BundledExecutable -Actual (Resolve-VbaDevExecutable) -Message 'Resolve-VbaDevExecutable should fall back when VBA_DEV_EXE is missing.'

    $bundled_wins_layout = New-TestVbaDevToolLayout -Root (Join-Path $resolver_root 'bundled-wins') -WithBundledExecutable
    Import-TestVbaDevToolLayout -Layout $bundled_wins_layout
    $env:VBA_DEV_EXE = $null
    $env:PATH = $path_root
    Test-PathEqual -Expected $bundled_wins_layout.BundledExecutable -Actual (Resolve-VbaDevExecutable) -Message 'Resolve-VbaDevExecutable should prefer the bundled executable over PATH.'

    $no_development_build_layout = New-TestVbaDevToolLayout -Root (Join-Path $resolver_root 'no-development-build')
    $old_development_build = Join-Path $no_development_build_layout.WorkspaceRoot 'vba-tools\tools\vba-dev\src\VbaDev.Cli\bin\Release\net10.0\win-x64\publish\vba-dev.exe'
    Write-TestFileUtf8 -Path $old_development_build -Content 'fake development build vba-dev'
    Import-TestVbaDevToolLayout -Layout $no_development_build_layout
    $env:VBA_DEV_EXE = $missing_env_executable
    $env:PATH = $empty_path_root
    $not_found_message = $null
    try {
        [void](Resolve-VbaDevExecutable)
    }
    catch {
        $not_found_message = $_.Exception.Message
    }

    Test-True -Condition (-not [string]::IsNullOrWhiteSpace($not_found_message)) -Message 'Resolve-VbaDevExecutable should fail when only the removed development build exists.'
    Test-True -Condition ($not_found_message -like "*VBA_DEV_EXE=$missing_env_executable*") -Message 'Resolve-VbaDevExecutable should include VBA_DEV_EXE in the not-found error.'
    Test-True -Condition ($not_found_message -like "*$($no_development_build_layout.BundledExecutable)*") -Message 'Resolve-VbaDevExecutable should include the bundled path in the not-found error.'
    Test-True -Condition ($not_found_message -like '*PATH:vba-dev.exe*') -Message 'Resolve-VbaDevExecutable should include PATH in the not-found error.'
    Test-True -Condition ($not_found_message -notlike '*vba-tools\tools\vba-dev\src\VbaDev.Cli\bin\Release*') -Message 'Resolve-VbaDevExecutable should not search the vba-tools development build.'

    Remove-Module -Name VbaDevTool -Force -ErrorAction SilentlyContinue
    $env:VBA_DEV_EXE = $previous_vba_dev_exe
    $env:PATH = $previous_path

    $call_log = Join-Path $temp_root 'vba-dev-calls.txt'
    $fake_vba_dev = Join-Path $temp_root 'vba-dev.cmd'
    [System.IO.File]::WriteAllText(
        $fake_vba_dev,
        "@echo off`r`necho %*>> `"%VBA_DEV_CALLS%`"`r`nexit /b 0`r`n",
        [System.Text.Encoding]::GetEncoding(932))
    $env:VBA_DEV_EXE = $fake_vba_dev
    $env:VBA_DEV_CALLS = $call_log

    $workbook_dir = Join-Path $temp_root 'WorkbookLocal'
    New-Item -ItemType Directory -Path $workbook_dir -Force | Out-Null
    $workbook = Join-Path $workbook_dir 'Book.xlsm'
    Write-TestFileUtf8 -Path $workbook -Content 'workbook'

    & (Join-Path $tools_root 'exp_mods_main.ps1') '' '' '' '' '' $workbook
    Test-True -Condition (Test-Path -LiteralPath (Join-Path $workbook_dir 'modules') -PathType Container) -Message 'EXP_MODS did not create the workbook-local source set.'
    $calls = @(Get-Content -LiteralPath $call_log)
    Test-Equal -Expected 1 -Actual $calls.Count -Message 'EXP_MODS should invoke vba-dev once.'
    Test-True -Condition ($calls[0] -like "export --from *Book.xlsm --to *modules") -Message "EXP_MODS called unexpected vba-dev arguments: $($calls[0])"
    Invoke-ExpectedFailure -ExpectedMessage 'target workbook must be a file' -Script {
        & (Join-Path $tools_root 'exp_mods_main.ps1') '' '' '' '' '' $workbook_dir
    }

    Clear-Content -LiteralPath $call_log
    $module_dir = Join-Path $workbook_dir 'modules'
    Write-TestFileSjis -Path (Join-Path $module_dir 'Module1.bas') -Content "Attribute VB_Name = `"Module1`"`r`n"
    & (Join-Path $tools_root 'imp_mods_main.ps1') '' '' '' '' '' $workbook
    $calls = @(Get-Content -LiteralPath $call_log)
    Test-Equal -Expected 1 -Actual $calls.Count -Message 'IMP_MODS should invoke vba-dev once.'
    Test-True -Condition ($calls[0] -like "import --from *modules --to *Book.xlsm") -Message "IMP_MODS called unexpected vba-dev arguments: $($calls[0])"
    Invoke-ExpectedFailure -ExpectedMessage 'workbook-local source set was not found' -Script {
        $other_workbook = Join-Path $temp_root 'Other.xlsm'
        Write-TestFileUtf8 -Path $other_workbook -Content 'workbook'
        & (Join-Path $tools_root 'imp_mods_main.ps1') '' '' '' '' '' $other_workbook
    }

    Clear-Content -LiteralPath $call_log
    $project_a = New-TestProject -Root (Join-Path $temp_root 'ProjectA') -ProjectName 'ProjectA' -DocumentName 'BookA'
    $project_b = New-TestProject -Root (Join-Path $temp_root 'ProjectB') -ProjectName 'ProjectB' -DocumentName 'BookB'
    $ignored_project = New-TestProject -Root (Join-Path (Join-Path $temp_root 'bin') 'Ignored') -ProjectName 'Ignored' -DocumentName 'IgnoredBook'
    & (Join-Path $tools_root 'update_common_mods_main.ps1') '' '' '' '' '' $temp_root
    $calls = @(Get-Content -LiteralPath $call_log)
    Test-True -Condition ($calls -contains "common-module update --project $($project_a.Root)") -Message 'UPDATE_COMMON_MODS did not update ProjectA.'
    Test-True -Condition ($calls -contains "common-module update --project $($project_b.Root)") -Message 'UPDATE_COMMON_MODS did not update ProjectB.'
    Test-True -Condition ($calls -contains "build --project $($project_a.Root) --document BookA") -Message 'UPDATE_COMMON_MODS did not build BookA.'
    Test-True -Condition (-not ($calls -contains "common-module update --project $($ignored_project.Root)")) -Message 'UPDATE_COMMON_MODS should ignore bin projects.'

    Clear-Content -LiteralPath $call_log
    $sync_root = Join-Path $temp_root 'sync'
    New-Item -ItemType Directory -Path $sync_root -Force | Out-Null
    $sync_project_a = New-TestProject -Root (Join-Path $sync_root 'SyncA') -ProjectName 'SyncA' -DocumentName 'SyncBookA'
    $sync_project_b = New-TestProject -Root (Join-Path $sync_root 'SyncB') -ProjectName 'SyncB' -DocumentName 'SyncBookB'
    Write-TestFileSjis -Path (Join-Path $sync_project_a.SourceSetPath 'Shared.bas') -Content "Attribute VB_Name = `"Shared`"`r`n'new`r`n"
    Write-TestFileSjis -Path (Join-Path $sync_project_b.SourceSetPath 'Shared.bas') -Content "Attribute VB_Name = `"Shared`"`r`n'old`r`n"
    (Get-Item -LiteralPath (Join-Path $sync_project_a.SourceSetPath 'Shared.bas')).LastWriteTime = (Get-Date).AddMinutes(1)
    (Get-Item -LiteralPath (Join-Path $sync_project_b.SourceSetPath 'Shared.bas')).LastWriteTime = (Get-Date).AddMinutes(-1)
    $sync_config = [ordered]@{
        modules = @('Shared.bas')
        targets = @(
            [ordered]@{ project = 'SyncA'; document = 'SyncBookA' },
            [ordered]@{ project = 'SyncB'; document = 'SyncBookB' }
        )
    }
    $sync_config | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $sync_root 'sync.json') -Encoding UTF8
    & (Join-Path $tools_root 'sync_mods_main.ps1') '' '' '' '' '' $sync_root
    Test-Equal -Expected (Get-Content -LiteralPath (Join-Path $sync_project_a.SourceSetPath 'Shared.bas') -Raw) -Actual (Get-Content -LiteralPath (Join-Path $sync_project_b.SourceSetPath 'Shared.bas') -Raw) -Message 'SYNC_MODS did not copy the newest module content.'
    $calls = @(Get-Content -LiteralPath $call_log)
    Test-Equal -Expected 1 -Actual $calls.Count -Message 'SYNC_MODS should import only changed target workbooks.'
    Test-True -Condition ($calls[0] -like "import --from *SyncBookB --to *SyncBookB.xlsm") -Message "SYNC_MODS called unexpected vba-dev arguments: $($calls[0])"

    $collect_root = Join-Path $temp_root 'collect-target'
    New-Item -ItemType Directory -Path $collect_root -Force | Out-Null
    & (Join-Path $tools_root 'collect_common_mods_main.ps1') '' '' '' '' '' $collect_root
    Test-True -Condition (Test-Path -LiteralPath (Join-Path (Join-Path $collect_root 'common_modules_repo') 'common-modules-manifest.tsv') -PathType Leaf) -Message 'COLLECT_COMMON_MODS did not copy the manifest.'

    $fake_doxygen_script = Join-Path $temp_root 'fake-doxygen.ps1'
    Write-TestFileUtf8 -Path $fake_doxygen_script -Content @"
function Get-DoxyfileSetting {
    param(
        [Parameter(Mandatory = `$true)]
        [string]`$ConfigPath,
        [Parameter(Mandatory = `$true)]
        [string]`$Name
    )

    `$line = Get-Content -LiteralPath `$ConfigPath | Where-Object {
        `$_ -match ('^' + [System.Text.RegularExpressions.Regex]::Escape(`$Name) + '\s*=')
    } | Select-Object -First 1
    if (`$null -eq `$line) {
        throw "Missing Doxyfile setting: `$Name"
    }

    return (`$line -replace ('^' + [System.Text.RegularExpressions.Regex]::Escape(`$Name) + '\s*=\s*'), '').Trim()
}

function ConvertFrom-DoxyfileValue {
    param(
        [Parameter(Mandatory = `$true)]
        [string]`$Value
    )

    if (`$Value.StartsWith('"') -and `$Value.EndsWith('"')) {
        return `$Value.Substring(1, `$Value.Length - 2)
    }

    return `$Value
}

`$config_path = `$args[0]
`$output_dir = ConvertFrom-DoxyfileValue -Value (Get-DoxyfileSetting -ConfigPath `$config_path -Name 'OUTPUT_DIRECTORY')
`$input_dir = ConvertFrom-DoxyfileValue -Value (Get-DoxyfileSetting -ConfigPath `$config_path -Name 'INPUT')
`$project_name = ConvertFrom-DoxyfileValue -Value (Get-DoxyfileSetting -ConfigPath `$config_path -Name 'PROJECT_NAME')
New-Item -ItemType Directory -Path `$output_dir -Force | Out-Null
Set-Content -LiteralPath (Join-Path `$output_dir 'index.html') -Value 'fake doxygen output' -Encoding UTF8
Set-Content -LiteralPath (Join-Path `$output_dir 'project-name.txt') -Value `$project_name -Encoding UTF8
Get-ChildItem -LiteralPath `$input_dir -File | Sort-Object -Property Name | ForEach-Object { `$_.Name } | Set-Content -LiteralPath (Join-Path `$output_dir 'source-files.txt') -Encoding UTF8
"@
    $fake_doxygen = Join-Path $temp_root 'doxygen.cmd'
    [System.IO.File]::WriteAllText($fake_doxygen, "@echo off`r`npowershell -NoProfile -ExecutionPolicy Bypass -File `"%~dp0fake-doxygen.ps1`" %*`r`nexit /b %ERRORLEVEL%`r`n", [System.Text.Encoding]::GetEncoding(932))
    $env:PATH = "$temp_root;$env:PATH"

    $doc_root = Join-Path $temp_root 'DocOwner'
    $doc_source = Join-Path $doc_root 'modules'
    New-Item -ItemType Directory -Path $doc_source -Force | Out-Null
    Write-TestFileSjis -Path (Join-Path $doc_source 'DocModule.bas') -Content "Attribute VB_Name = `"DocModule`"`r`n"
    Write-TestFileSjis -Path (Join-Path $doc_source 'DocForm.frx') -Content 'binary sidecar'
    $nested_doc_source = Join-Path $doc_source 'Nested'
    New-Item -ItemType Directory -Path $nested_doc_source -Force | Out-Null
    Write-TestFileSjis -Path (Join-Path $nested_doc_source 'NestedModule.bas') -Content "Attribute VB_Name = `"NestedModule`"`r`n"
    $preserved_doc_target = Join-Path (Join-Path (Join-Path $doc_root 'docs') 'OtherTarget') 'api-reference'
    New-Item -ItemType Directory -Path $preserved_doc_target -Force | Out-Null
    Write-TestFileUtf8 -Path (Join-Path $preserved_doc_target 'keep.txt') -Content 'keep'
    & (Join-Path $tools_root 'gen_doc_main.ps1') '' '' '' $doc_source '' $doc_source
    $doc_api = Join-Path (Join-Path $doc_root 'docs') 'api-reference'
    Test-True -Condition (Test-Path -LiteralPath $doc_api -PathType Container) -Message 'GEN_DOC did not create docs/api-reference under the source-set owner.'
    Test-True -Condition (Test-Path -LiteralPath (Join-Path $doc_root 'docs\api-reference.zip') -PathType Leaf) -Message 'GEN_DOC did not create docs/api-reference.zip.'
    Test-Equal -Expected 'DocOwner' -Actual ((Get-Content -LiteralPath (Join-Path $doc_api 'project-name.txt') -Raw).Trim()) -Message 'GEN_DOC used the wrong owner-level project name.'
    $doc_input_files = @(Get-Content -LiteralPath (Join-Path $doc_api 'source-files.txt'))
    Test-Equal -Expected 1 -Actual $doc_input_files.Count -Message 'GEN_DOC should pass only direct VBA source files to Doxygen.'
    Test-Equal -Expected 'DocModule.bas' -Actual $doc_input_files[0] -Message 'GEN_DOC passed an unexpected owner-level source file.'
    Test-True -Condition (Test-Path -LiteralPath (Join-Path $preserved_doc_target 'keep.txt') -PathType Leaf) -Message 'GEN_DOC removed an unrelated documentation target.'

    $target_owner = Join-Path $temp_root 'TargetOwner'
    $target_name = 'Target Source'
    $target_source = Join-Path (Join-Path $target_owner 'src') $target_name
    New-Item -ItemType Directory -Path $target_source -Force | Out-Null
    Write-TestFileSjis -Path (Join-Path $target_source 'TargetClass.cls') -Content "VERSION 1.0 CLASS`r`nAttribute VB_Name = `"TargetClass`"`r`n"
    $owner_level_api = Join-Path (Join-Path $target_owner 'docs') 'api-reference'
    New-Item -ItemType Directory -Path $owner_level_api -Force | Out-Null
    Write-TestFileUtf8 -Path (Join-Path $owner_level_api 'keep.txt') -Content 'keep'
    & (Join-Path $tools_root 'gen_doc_main.ps1') '' '' '' $target_source '' $target_source
    $target_api = Join-Path (Join-Path (Join-Path $target_owner 'docs') $target_name) 'api-reference'
    Test-True -Condition (Test-Path -LiteralPath $target_api -PathType Container) -Message 'GEN_DOC did not create docs/<target>/api-reference for a src child source set.'
    Test-True -Condition (Test-Path -LiteralPath ($target_api + '.zip') -PathType Leaf) -Message 'GEN_DOC did not create docs/<target>/api-reference.zip.'
    Test-Equal -Expected $target_name -Actual ((Get-Content -LiteralPath (Join-Path $target_api 'project-name.txt') -Raw).Trim()) -Message 'GEN_DOC used the wrong src child project name.'
    Test-True -Condition (Test-Path -LiteralPath (Join-Path $owner_level_api 'keep.txt') -PathType Leaf) -Message 'GEN_DOC removed owner-level documentation while generating src child documentation.'

    $src_owner = Join-Path $temp_root 'SrcOwner'
    $src_source = Join-Path $src_owner 'src'
    New-Item -ItemType Directory -Path $src_source -Force | Out-Null
    Write-TestFileSjis -Path (Join-Path $src_source 'RootForm.frm') -Content "VERSION 5.00`r`nAttribute VB_Name = `"RootForm`"`r`n"
    $warnings = & (Join-Path $tools_root 'gen_doc_main.ps1') '' '' '' $src_source '' $src_source 3>&1
    Test-True -Condition (($warnings | Out-String) -like '*normally a container directory*') -Message 'GEN_DOC did not warn when the source directory itself was named src.'
    $src_api = Join-Path (Join-Path $src_owner 'docs') 'api-reference'
    Test-True -Condition (Test-Path -LiteralPath $src_api -PathType Container) -Message 'GEN_DOC did not create owner-level docs/api-reference for a src source directory.'
    Test-Equal -Expected 'SrcOwner' -Actual ((Get-Content -LiteralPath (Join-Path $src_api 'project-name.txt') -Raw).Trim()) -Message 'GEN_DOC used the wrong src fallback project name.'

    $empty_root = Join-Path $temp_root 'EmptyOwner'
    $empty_source = Join-Path $empty_root 'modules'
    New-Item -ItemType Directory -Path (Join-Path $empty_source 'Nested') -Force | Out-Null
    Write-TestFileSjis -Path (Join-Path (Join-Path $empty_source 'Nested') 'NestedOnly.bas') -Content "Attribute VB_Name = `"NestedOnly`"`r`n"
    Invoke-ExpectedFailure -ExpectedMessage 'No VBA source files were found in source directory' -Script {
        & (Join-Path $tools_root 'gen_doc_main.ps1') '' '' '' $empty_source '' $empty_source
    }

    Write-Host 'Tool wrapper tests passed.'
}
finally {
    $env:VBA_DEV_EXE = $previous_vba_dev_exe
    $env:VBA_DEV_CALLS = $previous_vba_dev_calls
    $env:PATH = $previous_path
    if (Test-Path -LiteralPath $temp_root) {
        Remove-Item -LiteralPath $temp_root -Recurse -Force
    }
}
