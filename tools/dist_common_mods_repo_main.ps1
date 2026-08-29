$ErrorActionPreference = 'Stop'
$InformationPreference = 'Continue'
$WarningPreference = 'Continue'

$invocation_working_directory = (Get-Location).ProviderPath

Import-Module (Join-Path $PSScriptRoot 'VbaDevTool.psm1') -Force

Set-Variable -Name COMMON_MODULES_REPO_DIR_NAME -Value 'common_modules_repo' -Option Constant
Set-Variable -Name COMMON_MODULES_MANIFEST_FILE_NAME -Value 'common-modules-manifest.tsv' -Option Constant
Set-Variable -Name VBA_WSC_PATTERN -Value '[\u0009\u0019\u0020\u1680\u180e\u2000-\u200a\u202f\u205f\u3000]' -Option Constant
Set-Variable -Name VBA_RESERVED_IDENTIFIERS -Value ([System.Collections.Generic.HashSet[string]]::new(
    [string[]]@(
        'Abs', 'AddressOf', 'And', 'Any', 'Array', 'As', 'Attribute',
        'Boolean', 'Byte', 'ByRef', 'ByVal', 'Call', 'Case', 'CBool',
        'CByte', 'CCur', 'CDate', 'CDecl', 'CDec', 'CDbl', 'Circle',
        'CInt', 'CLng', 'CLngLng', 'CLngPtr', 'Close', 'Const', 'CSng',
        'CStr', 'Currency', 'CVar', 'CVErr', 'Date', 'Debug', 'Decimal',
        'Declare', 'DefBool', 'DefByte', 'DefCur', 'DefDate', 'DefDbl',
        'DefDec', 'DefInt', 'DefLng', 'DefLngLng', 'DefLngPtr', 'DefObj',
        'DefSng', 'DefStr', 'DefVar', 'Dim', 'Do', 'DoEvents', 'Double',
        'Each', 'Else', 'ElseIf', 'Empty', 'End', 'EndIf', 'Enum', 'Eqv',
        'Erase', 'Event', 'Exit', 'False', 'Fix', 'For', 'Friend',
        'Function', 'Get', 'Global', 'GoSub', 'GoTo', 'If', 'Imp',
        'Implements', 'In', 'Input', 'InputB', 'Int', 'Integer', 'Is',
        'LBound', 'Len', 'LenB', 'Let', 'Like', 'LINEINPUT', 'Lock',
        'Long', 'LongLong', 'LongPtr', 'Loop', 'LSet', 'Me', 'Mod', 'New',
        'Next', 'Not', 'Nothing', 'Null', 'On', 'Open', 'Option', 'Optional',
        'Or', 'ParamArray', 'Preserve', 'Print', 'Private', 'PSet', 'Public',
        'Put', 'RaiseEvent', 'ReDim', 'Rem', 'Resume', 'Return', 'RSet',
        'Scale', 'Seek', 'Select', 'Set', 'Sgn', 'Shared', 'Single', 'Spc',
        'Static', 'Stop', 'String', 'Sub', 'Tab', 'Then', 'To', 'True',
        'Type', 'TypeOf', 'UBound', 'Unlock', 'Until', 'Variant', 'VB_Base',
        'VB_Control', 'VB_Creatable', 'VB_Customizable', 'VB_Description',
        'VB_Exposed', 'VB_Ext_KEY', 'VB_GlobalNameSpace', 'VB_HelpID',
        'VB_Invoke_Func', 'VB_Invoke_Property', 'VB_Invoke_PropertyPut',
        'VB_Invoke_PropertyPutRef', 'VB_MemberFlags', 'VB_Name',
        'VB_PredeclaredId', 'VB_ProcData', 'VB_TemplateDerived',
        'VB_UserMemId', 'VB_VarDescription', 'VB_VarHelpID',
        'VB_VarMemberFlags', 'VB_VarProcData', 'VB_VarUserMemId', 'Wend',
        'While', 'With', 'WithEvents', 'Write', 'Xor'
    ),
    [System.StringComparer]::OrdinalIgnoreCase)) -Option Constant

function Test-ReparsePoint {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileSystemInfo]$Item
    )

    return (($Item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0)
}

function Get-NormalizedLexicalPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $full_path = [System.IO.Path]::GetFullPath($Path)
    $root_path = [System.IO.Path]::GetPathRoot($full_path)
    while ($full_path.Length -gt $root_path.Length -and
        ($full_path.EndsWith([System.IO.Path]::DirectorySeparatorChar) -or
            $full_path.EndsWith([System.IO.Path]::AltDirectorySeparatorChar))) {
        $full_path = $full_path.Substring(0, $full_path.Length - 1)
    }

    return $full_path
}

function Get-OrdinalSortedItems {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$Items,

        [Parameter(Mandatory = $true)]
        [string]$PropertyName
    )

    $items_by_value = New-Object 'System.Collections.Generic.Dictionary[string, object]' ([System.StringComparer]::Ordinal)
    foreach ($item in $Items) {
        $value = [string]$item.$PropertyName
        if ($items_by_value.ContainsKey($value)) {
            throw "Cannot order duplicate '$PropertyName' value: $value"
        }
        $items_by_value.Add($value, $item)
    }

    [string[]]$ordered_values = @($items_by_value.Keys)
    [Array]::Sort($ordered_values, [System.StringComparer]::Ordinal)
    foreach ($value in $ordered_values) {
        $items_by_value[$value]
    }
}

function Resolve-DistributionSearchRoot {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$Arguments,

        [Parameter(Mandatory = $true)]
        [string]$InvocationWorkingDirectory
    )

    if ($Arguments.Count -ne 1) {
        throw 'DIST requires exactly one Distribution Search Root argument.'
    }

    $argument_path = [string]$Arguments[0]
    if ([string]::IsNullOrWhiteSpace($argument_path)) {
        throw 'Distribution Search Root is empty.'
    }

    if ([System.IO.Path]::IsPathRooted($argument_path)) {
        $candidate_path = $argument_path
    }
    else {
        $candidate_path = Join-Path $InvocationWorkingDirectory $argument_path
    }

    $resolved_path = (Resolve-Path -LiteralPath $candidate_path -ErrorAction Stop).ProviderPath
    $search_root = Get-Item -LiteralPath $resolved_path -Force -ErrorAction Stop
    if (-not $search_root.PSIsContainer) {
        throw "Distribution Search Root must be a directory: $resolved_path"
    }

    return $search_root
}

function Get-WrapperSourceRepository {
    param(
        [Parameter(Mandatory = $true)]
        [string]$InvocationWorkingDirectory
    )

    $parent_inventory = @(Get-ChildItem -LiteralPath $InvocationWorkingDirectory -Force -ErrorAction Stop)
    $matches = @($parent_inventory | Where-Object {
        [string]::Equals($_.Name, $COMMON_MODULES_REPO_DIR_NAME, [System.StringComparison]::OrdinalIgnoreCase)
    })
    if ($matches.Count -eq 0) {
        throw "Source $COMMON_MODULES_REPO_DIR_NAME was not found directly under the invocation working directory: $InvocationWorkingDirectory"
    }
    if ($matches.Count -ne 1) {
        throw "Source has multiple case-insensitive '$COMMON_MODULES_REPO_DIR_NAME' matches directly under the invocation working directory: $InvocationWorkingDirectory"
    }

    $source_repository = $matches[0]
    if (-not [string]::Equals($source_repository.Name, $COMMON_MODULES_REPO_DIR_NAME, [System.StringComparison]::Ordinal)) {
        throw "Source repository must use the ordinal-exact '$COMMON_MODULES_REPO_DIR_NAME' name: $($source_repository.FullName)"
    }
    if (-not $source_repository.PSIsContainer -or $source_repository -isnot [System.IO.DirectoryInfo] -or (Test-ReparsePoint -Item $source_repository)) {
        throw "Source repository must be an ordinary non-reparse directory: $($source_repository.FullName)"
    }

    return $source_repository
}

function Get-ExactPackageFile {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$Inventory,

        [Parameter(Mandatory = $true)]
        [string]$ExpectedName,

        [Parameter(Mandatory = $true)]
        [string]$SourceRepositoryPath,

        [switch]$AllowAbsent
    )

    $matches = @($Inventory | Where-Object {
        [string]::Equals($_.Name, $ExpectedName, [System.StringComparison]::OrdinalIgnoreCase)
    })
    if ($matches.Count -eq 0) {
        if ($AllowAbsent) {
            return $null
        }
        throw "Source package entry was not found in '$SourceRepositoryPath': $ExpectedName"
    }
    if ($matches.Count -ne 1) {
        throw "Source package has multiple case-insensitive '$ExpectedName' matches in '$SourceRepositoryPath'."
    }

    $match = $matches[0]
    if (-not [string]::Equals($match.Name, $ExpectedName, [System.StringComparison]::Ordinal)) {
        throw "Source package entry must use the ordinal-exact '$ExpectedName' name: $($match.FullName)"
    }
    if ($match.PSIsContainer -or $match -isnot [System.IO.FileInfo] -or (Test-ReparsePoint -Item $match)) {
        throw "Source package entry must be an ordinary non-reparse file: $($match.FullName)"
    }

    return $match
}

function Assert-FileReadable {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$File,

        [Parameter(Mandatory = $true)]
        [string]$Description
    )

    $stream = $null
    try {
        $stream = [System.IO.File]::Open(
            $File.FullName,
            [System.IO.FileMode]::Open,
            [System.IO.FileAccess]::Read,
            [System.IO.FileShare]::Read)
        $buffer = New-Object byte[] 81920
        while ($stream.Read($buffer, 0, $buffer.Length) -gt 0) {
        }
    }
    catch {
        throw "$Description is unreadable: $($File.FullName). $($_.Exception.Message)"
    }
    finally {
        if ($null -ne $stream) {
            $stream.Dispose()
        }
    }
}

function Test-VbaModuleIdentity {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Value
    )

    $code_point_count = 0
    for ($index = 0; $index -lt $Value.Length; $index++) {
        if ([char]::IsHighSurrogate($Value[$index])) {
            if ($index + 1 -ge $Value.Length -or -not [char]::IsLowSurrogate($Value[$index + 1])) {
                return $false
            }
            $index++
        }
        elseif ([char]::IsLowSurrogate($Value[$index])) {
            return $false
        }
        $code_point_count++
    }

    return $code_point_count -le 31 -and
        -not $VBA_RESERVED_IDENTIFIERS.Contains($Value) -and
        [regex]::IsMatch(
        $Value,
        '^\p{L}[\p{L}\p{Nd}_]*$',
        [System.Text.RegularExpressions.RegexOptions]::CultureInvariant)
}

function Test-VbaWhitespaceOnly {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Value
    )

    return [regex]::IsMatch(
        $Value,
        "^(?:$VBA_WSC_PATTERN)*$",
        [System.Text.RegularExpressions.RegexOptions]::CultureInvariant)
}

function Get-VbaWhitespaceTrimmedText {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Value
    )

    $trimmed_start = [regex]::Replace(
        $Value,
        "^(?:$VBA_WSC_PATTERN)+",
        '',
        [System.Text.RegularExpressions.RegexOptions]::CultureInvariant)
    return [regex]::Replace(
        $trimmed_start,
        "(?:$VBA_WSC_PATTERN)+$",
        '',
        [System.Text.RegularExpressions.RegexOptions]::CultureInvariant)
}

function Test-VbaIdentifierWordCharacter {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Value
    )

    if ($Value.Length -ne 1) {
        return $false
    }

    $code_point = [int]$Value[0]
    if ($code_point -le 0x7f) {
        return [regex]::IsMatch($Value, '^[A-Za-z0-9_]$', [System.Text.RegularExpressions.RegexOptions]::CultureInvariant)
    }

    # The source has already passed strict Windows-932 byte round-trip validation.
    # This 32-code-point complement is derived from the union of the parser-owned
    # Cp2 and all code-page SubsequentRanges over every canonical Windows-932 char.
    return -not (
        $code_point -eq 0x3000 -or
        $code_point -eq 0xff01 -or
        ($code_point -ge 0xff03 -and $code_point -le 0xff06) -or
        ($code_point -ge 0xff08 -and $code_point -le 0xff0f) -or
        ($code_point -ge 0xff1a -and $code_point -le 0xff20) -or
        ($code_point -ge 0xff3b -and $code_point -le 0xff3e) -or
        $code_point -eq 0xff40 -or
        ($code_point -ge 0xff5b -and $code_point -le 0xff5e) -or
        $code_point -eq 0xffe3 -or
        $code_point -eq 0xffe5)
}

function Test-ObjectModuleIdentityPlacement {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourceText
    )

    $lines = @([regex]::Split($SourceText, "`r`n|`r|`n"))
    $first_nonempty_index = -1
    for ($index = 0; $index -lt $lines.Count; $index++) {
        if (-not (Test-VbaWhitespaceOnly -Value $lines[$index])) {
            $first_nonempty_index = $index
            break
        }
    }
    if ($first_nonempty_index -lt 0) {
        return $false
    }

    $designer_blocks = New-Object 'System.Collections.Generic.Stack[string]'
    $saw_designer_block = $false
    $header_started = $false
    for ($index = $first_nonempty_index + 1; $index -lt $lines.Count; $index++) {
        $line = Get-VbaWhitespaceTrimmedText -Value $lines[$index]
        if ($line.Length -eq 0) {
            if ($header_started) {
                return $false
            }
            continue
        }

        if ([regex]::IsMatch(
                $line,
                "^Attribute${VBA_WSC_PATTERN}+VB_Name${VBA_WSC_PATTERN}*=${VBA_WSC_PATTERN}*`"[^`"]+`"${VBA_WSC_PATTERN}*$",
                [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [System.Text.RegularExpressions.RegexOptions]::CultureInvariant)) {
            return $designer_blocks.Count -eq 0
        }

        if ([regex]::IsMatch(
                $line,
                "^BeginProperty(?=$VBA_WSC_PATTERN|$)",
                [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [System.Text.RegularExpressions.RegexOptions]::CultureInvariant)) {
            if ($designer_blocks.Count -eq 0 -and $header_started) {
                return $false
            }
            $saw_designer_block = $true
            $designer_blocks.Push('Property')
            continue
        }
        if ([regex]::IsMatch(
                $line,
                "^Begin(?=$VBA_WSC_PATTERN|$)",
                [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [System.Text.RegularExpressions.RegexOptions]::CultureInvariant)) {
            if ($designer_blocks.Count -eq 0 -and $header_started) {
                return $false
            }
            $saw_designer_block = $true
            $designer_blocks.Push('Component')
            continue
        }

        if ([string]::Equals($line, 'EndProperty', [System.StringComparison]::OrdinalIgnoreCase)) {
            if ($designer_blocks.Count -eq 0 -or -not [string]::Equals($designer_blocks.Peek(), 'Property', [System.StringComparison]::Ordinal)) {
                return $false
            }
            [void]$designer_blocks.Pop()
            continue
        }
        if ([string]::Equals($line, 'End', [System.StringComparison]::OrdinalIgnoreCase)) {
            if ($designer_blocks.Count -eq 0 -or -not [string]::Equals($designer_blocks.Peek(), 'Component', [System.StringComparison]::Ordinal)) {
                return $false
            }
            [void]$designer_blocks.Pop()
            continue
        }

        if ($designer_blocks.Count -gt 0) {
            continue
        }
        if ([regex]::IsMatch(
                $line,
                "^Object${VBA_WSC_PATTERN}*=",
                [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [System.Text.RegularExpressions.RegexOptions]::CultureInvariant)) {
            if ($saw_designer_block -or $header_started) {
                return $false
            }
            continue
        }
        if ([regex]::IsMatch(
                $line,
                "^Attribute${VBA_WSC_PATTERN}+(?:VB_GlobalNameSpace|VB_Creatable)${VBA_WSC_PATTERN}*=${VBA_WSC_PATTERN}*False${VBA_WSC_PATTERN}*$",
                [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [System.Text.RegularExpressions.RegexOptions]::CultureInvariant)) {
            $header_started = $true
            continue
        }
        if ([regex]::IsMatch(
                $line,
                "^Attribute${VBA_WSC_PATTERN}+(?:VB_PredeclaredId|VB_Exposed|VB_Customizable)${VBA_WSC_PATTERN}*=${VBA_WSC_PATTERN}*(?:True|False)${VBA_WSC_PATTERN}*$",
                [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [System.Text.RegularExpressions.RegexOptions]::CultureInvariant)) {
            $header_started = $true
            continue
        }

        return $false
    }

    return $false
}

function Confirm-SourceManifestMetadata {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$SourceFile,

        [Parameter(Mandatory = $true)]
        [string]$ManifestIdentity
    )

    try {
        $strict_shift_jis = [System.Text.Encoding]::GetEncoding(
            932,
            [System.Text.EncoderExceptionFallback]::new(),
            [System.Text.DecoderExceptionFallback]::new())
        $source_bytes = [System.IO.File]::ReadAllBytes($SourceFile.FullName)
        $source_text = $strict_shift_jis.GetString($source_bytes)
        $round_trip_bytes = $strict_shift_jis.GetBytes($source_text)
        if ($source_bytes.Length -ne $round_trip_bytes.Length) {
            throw 'The decoded source could not reproduce its canonical Windows-932 bytes.'
        }
        for ($index = 0; $index -lt $source_bytes.Length; $index++) {
            if ($source_bytes[$index] -ne $round_trip_bytes[$index]) {
                throw 'The decoded source could not reproduce its canonical Windows-932 bytes.'
            }
        }
    }
    catch {
        throw "Source package entry could not be decoded as Windows-932 text: $($SourceFile.FullName). $($_.Exception.Message)"
    }

    $source_lines = @([regex]::Split($source_text, "`r`n|`r|`n"))
    $first_nonempty_line = $null
    foreach ($source_line in $source_lines) {
        if (-not (Test-VbaWhitespaceOnly -Value $source_line)) {
            $first_nonempty_line = Get-VbaWhitespaceTrimmedText -Value $source_line
            break
        }
    }
    if ($null -eq $first_nonempty_line) {
        $actual_kind = 'StandardModule'
    }
    elseif ([string]::Equals($first_nonempty_line, 'VERSION 1.0 CLASS', [System.StringComparison]::OrdinalIgnoreCase)) {
        $actual_kind = 'ClassModule'
    }
    elseif ([string]::Equals($first_nonempty_line, 'VERSION 5.00', [System.StringComparison]::OrdinalIgnoreCase)) {
        $actual_kind = 'FormModule'
    }
    else {
        $actual_kind = 'StandardModule'
    }

    switch ([System.IO.Path]::GetExtension($SourceFile.Name)) {
        '.bas' { $expected_kind = 'StandardModule' }
        '.cls' { $expected_kind = 'ClassModule' }
        '.frm' { $expected_kind = 'FormModule' }
        default { throw "Source package entry has an unsupported source extension: $($SourceFile.FullName)" }
    }
    if (-not [string]::Equals($actual_kind, $expected_kind, [System.StringComparison]::Ordinal)) {
        throw "Source package entry '$($SourceFile.Name)' declares source kind '$actual_kind' instead of '$expected_kind'."
    }

    $regex_options = [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [System.Text.RegularExpressions.RegexOptions]::CultureInvariant
    $identity_pattern = "^${VBA_WSC_PATTERN}*Attribute${VBA_WSC_PATTERN}+VB_Name${VBA_WSC_PATTERN}*=${VBA_WSC_PATTERN}*`"(?<name>[^`"]+)`"${VBA_WSC_PATTERN}*$"
    $identity_prefix_pattern = "^${VBA_WSC_PATTERN}*Attribute${VBA_WSC_PATTERN}+(?<keyword>VB_Name)"
    $identity_records = New-Object 'System.Collections.Generic.List[object]'
    $identity_like_count = 0
    for ($line_index = 0; $line_index -lt $source_lines.Count; $line_index++) {
        $source_line = $source_lines[$line_index]
        $prefix_match = [regex]::Match($source_line, $identity_prefix_pattern, $regex_options)
        if (-not $prefix_match.Success) {
            continue
        }

        $is_identity_like = $prefix_match.Length -eq $source_line.Length
        if (-not $is_identity_like) {
            $next_character = $source_line.Substring($prefix_match.Length, 1)
            $is_identity_like = -not (Test-VbaIdentifierWordCharacter -Value $next_character)
        }
        if (-not $is_identity_like) {
            continue
        }

        $identity_like_count++
        $identity_match = [regex]::Match($source_line, $identity_pattern, $regex_options)
        if ($identity_match.Success) {
            $identity_records.Add([pscustomobject]@{
                    LineIndex = $line_index
                    Name = $identity_match.Groups['name'].Value
                })
        }
    }

    if ($identity_like_count -ne $identity_records.Count) {
        throw "Source package entry contains invalid ModuleIdentity metadata: $($SourceFile.FullName)"
    }
    if ($identity_records.Count -ne 1) {
        throw "Source package entry must contain exactly one authoritative ModuleIdentity record: $($SourceFile.FullName)"
    }
    if ([string]::Equals($expected_kind, 'StandardModule', [System.StringComparison]::Ordinal) -and
        $identity_records[0].LineIndex -ne 0) {
        throw "Standard-module identity metadata must begin on physical line 1: $($SourceFile.FullName)"
    }
    if (-not [string]::Equals($expected_kind, 'StandardModule', [System.StringComparison]::Ordinal) -and
        -not (Test-ObjectModuleIdentityPlacement -SourceText $source_text)) {
        throw "Object-module identity metadata must appear in the exported object-module header: $($SourceFile.FullName)"
    }

    $source_identity = $identity_records[0].Name
    if (-not (Test-VbaModuleIdentity -Value $source_identity)) {
        throw "Source package entry '$($SourceFile.Name)' declares invalid ModuleIdentity '$source_identity'."
    }
    if (-not [string]::Equals($source_identity, $ManifestIdentity, [System.StringComparison]::Ordinal)) {
        throw "Source package entry '$($SourceFile.Name)' declares ModuleIdentity '$source_identity' instead of exact manifest identity '$ManifestIdentity'."
    }
}

function New-SourcePackageEntry {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$File
    )

    $File.Refresh()
    if (-not $File.Exists -or (Test-ReparsePoint -Item $File)) {
        throw "Source package entry must remain an ordinary non-reparse file: $($File.FullName)"
    }
    Assert-FileReadable -File $File -Description 'Source package entry'

    return [pscustomobject]@{
        Name = $File.Name
        FullName = $File.FullName
        LastWriteTimeUtc = $File.LastWriteTimeUtc
        Length = $File.Length
    }
}

function Read-DistributionSourcePackage {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.DirectoryInfo]$SourceRepository
    )

    $inventory = @(Get-ChildItem -LiteralPath $SourceRepository.FullName -Force -ErrorAction Stop)
    foreach ($item in $inventory) {
        if ($item.PSIsContainer -or $item -isnot [System.IO.FileInfo] -or (Test-ReparsePoint -Item $item)) {
            throw "Source package must be flat and contain only ordinary non-reparse files: $($item.FullName)"
        }
    }

    $manifest_file = Get-ExactPackageFile -Inventory $inventory -ExpectedName $COMMON_MODULES_MANIFEST_FILE_NAME -SourceRepositoryPath $SourceRepository.FullName
    $manifest = Read-CommonModulesManifest -ManifestPath $manifest_file.FullName

    $expected_names = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::Ordinal)
    $files_by_name = New-Object 'System.Collections.Generic.Dictionary[string, System.IO.FileInfo]' ([System.StringComparer]::Ordinal)
    $module_files_by_identity = New-Object 'System.Collections.Generic.Dictionary[string, string]' ([System.StringComparer]::OrdinalIgnoreCase)
    [void]$expected_names.Add($COMMON_MODULES_MANIFEST_FILE_NAME)
    $files_by_name.Add($COMMON_MODULES_MANIFEST_FILE_NAME, $manifest_file)

    foreach ($record in $manifest.Records) {
        $manifest_identity = [System.IO.Path]::GetFileNameWithoutExtension($record.ModuleFile)
        if ($module_files_by_identity.ContainsKey($manifest_identity)) {
            throw "Source package contains duplicate CommonModuleName '$manifest_identity': '$($module_files_by_identity[$manifest_identity])' and '$($record.ModuleFile)'."
        }
        $module_files_by_identity.Add($manifest_identity, $record.ModuleFile)

        $module_file = Get-ExactPackageFile -Inventory $inventory -ExpectedName $record.ModuleFile -SourceRepositoryPath $SourceRepository.FullName
        Confirm-SourceManifestMetadata -SourceFile $module_file -ManifestIdentity $manifest_identity
        [void]$expected_names.Add($record.ModuleFile)
        $files_by_name.Add($record.ModuleFile, $module_file)

        if ([string]::Equals([System.IO.Path]::GetExtension($record.ModuleFile), '.frm', [System.StringComparison]::Ordinal)) {
            $sidecar_name = [System.IO.Path]::GetFileNameWithoutExtension($record.ModuleFile) + '.frx'
            $sidecar_file = Get-ExactPackageFile -Inventory $inventory -ExpectedName $sidecar_name -SourceRepositoryPath $SourceRepository.FullName -AllowAbsent
            if ($null -ne $sidecar_file) {
                [void]$expected_names.Add($sidecar_name)
                $files_by_name.Add($sidecar_name, $sidecar_file)
            }
        }
    }

    foreach ($item in $inventory) {
        if (-not $expected_names.Contains($item.Name)) {
            throw "Source package contains an unexpected entry: $($item.FullName)"
        }
    }

    [string[]]$ordered_names = @($files_by_name.Keys)
    [Array]::Sort($ordered_names, [System.StringComparer]::Ordinal)
    $entries = New-Object 'System.Collections.Generic.List[object]'
    foreach ($name in $ordered_names) {
        $entries.Add((New-SourcePackageEntry -File $files_by_name[$name]))
    }

    return [pscustomobject]@{
        SourceRepository = $SourceRepository
        Entries = $entries.ToArray()
    }
}

function Assert-DistributionSourceSnapshot {
    param(
        [Parameter(Mandatory = $true)]
        [object]$SourcePackage
    )

    $source_repository = Get-Item -LiteralPath $SourcePackage.SourceRepository.FullName -Force -ErrorAction Stop
    if (-not $source_repository.PSIsContainer -or $source_repository -isnot [System.IO.DirectoryInfo] -or
        -not [string]::Equals($source_repository.Name, $COMMON_MODULES_REPO_DIR_NAME, [System.StringComparison]::Ordinal) -or
        (Test-ReparsePoint -Item $source_repository)) {
        throw "Source repository is no longer the validated ordinary '$COMMON_MODULES_REPO_DIR_NAME' directory: $($SourcePackage.SourceRepository.FullName)"
    }

    $inventory = @(Get-ChildItem -LiteralPath $source_repository.FullName -Force -ErrorAction Stop)
    if ($inventory.Count -ne $SourcePackage.Entries.Count) {
        throw "Source package inventory changed after validation: $($source_repository.FullName)"
    }

    foreach ($item in $inventory) {
        if ($item.PSIsContainer -or $item -isnot [System.IO.FileInfo] -or (Test-ReparsePoint -Item $item)) {
            throw "Source package became invalid after validation: $($item.FullName)"
        }
    }

    foreach ($entry in $SourcePackage.Entries) {
        $matches = @($inventory | Where-Object {
            [string]::Equals($_.Name, $entry.Name, [System.StringComparison]::Ordinal)
        })
        if ($matches.Count -ne 1) {
            throw "Source package inventory changed after validation: $($entry.FullName)"
        }

        $file = $matches[0]
        $file.Refresh()
        if (-not $file.Exists -or (Test-ReparsePoint -Item $file) -or
            $file.LastWriteTimeUtc.Ticks -ne $entry.LastWriteTimeUtc.Ticks -or
            $file.Length -ne $entry.Length) {
            throw "Source package metadata changed after validation: $($entry.FullName)"
        }
        Assert-FileReadable -File $file -Description 'Source package entry'
    }
}

function New-DistributionCandidateFailure {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SortPath,

        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    return [pscustomobject]@{
        Kind = 'Failure'
        SortPath = $SortPath
        TargetPath = $SortPath
        TargetDirectory = $null
        Message = $Message
    }
}

function New-DistributionTargetCandidate {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.DirectoryInfo]$TargetDirectory
    )

    return [pscustomobject]@{
        Kind = 'Target'
        SortPath = (Get-NormalizedLexicalPath -Path $TargetDirectory.FullName)
        TargetPath = $TargetDirectory.FullName
        TargetDirectory = $TargetDirectory
        Message = $null
    }
}

function Find-DistributionCandidates {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.DirectoryInfo]$SearchRoot,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$SearchRootInventory,

        [Parameter(Mandatory = $true)]
        [System.IO.DirectoryInfo]$SourceRepository
    )

    $candidates = New-Object 'System.Collections.Generic.List[object]'
    $target_count = 0
    $failure_count = 0
    $source_path = Get-NormalizedLexicalPath -Path $SourceRepository.FullName

    $project_directories = @($SearchRootInventory | Where-Object { $_.PSIsContainer })
    $ordered_projects = @(Get-OrdinalSortedItems -Items $project_directories -PropertyName 'FullName')
    foreach ($project_directory in $ordered_projects) {
        if (Test-ReparsePoint -Item $project_directory) {
            Write-Warning "Distribution project child is a reparse point and was skipped: $($project_directory.FullName)"
            continue
        }

        try {
            $project_inventory = @(Get-ChildItem -LiteralPath $project_directory.FullName -Force -ErrorAction Stop)
            $matches = @($project_inventory | Where-Object {
                [string]::Equals($_.Name, $COMMON_MODULES_REPO_DIR_NAME, [System.StringComparison]::OrdinalIgnoreCase)
            })
        }
        catch {
            Write-Warning "Could not determine whether project child opted in; it was skipped untouched: $($project_directory.FullName). $($_.Exception.Message)"
            continue
        }

        if ($matches.Count -eq 0) {
            continue
        }

        $expected_target_path = Get-NormalizedLexicalPath -Path (Join-Path $project_directory.FullName $COMMON_MODULES_REPO_DIR_NAME)
        if ($matches.Count -ne 1) {
            $candidates.Add((New-DistributionCandidateFailure -SortPath $expected_target_path -Message "Multiple case-insensitive '$COMMON_MODULES_REPO_DIR_NAME' entries were found in '$($project_directory.FullName)'."))
            $failure_count++
            continue
        }

        $match = $matches[0]
        if (-not [string]::Equals($match.Name, $COMMON_MODULES_REPO_DIR_NAME, [System.StringComparison]::Ordinal)) {
            $candidates.Add((New-DistributionCandidateFailure -SortPath $expected_target_path -Message "Opt-in repository must use the ordinal-exact '$COMMON_MODULES_REPO_DIR_NAME' name: $($match.FullName)"))
            $failure_count++
            continue
        }
        if (-not $match.PSIsContainer -or $match -isnot [System.IO.DirectoryInfo] -or (Test-ReparsePoint -Item $match)) {
            $candidates.Add((New-DistributionCandidateFailure -SortPath $expected_target_path -Message "Opt-in repository must be an ordinary non-reparse directory: $($match.FullName)"))
            $failure_count++
            continue
        }

        $target_path = Get-NormalizedLexicalPath -Path $match.FullName
        if ([string]::Equals($target_path, $source_path, [System.StringComparison]::OrdinalIgnoreCase)) {
            continue
        }

        $candidates.Add((New-DistributionTargetCandidate -TargetDirectory $match))
        $target_count++
    }

    $ordered_candidates = @(Get-OrdinalSortedItems -Items $candidates.ToArray() -PropertyName 'SortPath')
    return [pscustomobject]@{
        Candidates = $ordered_candidates
        TargetCount = $target_count
        FailureCount = $failure_count
    }
}

function Get-CurrentDistributionTarget {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TargetPath
    )

    try {
        $target = Get-Item -LiteralPath $TargetPath -Force -ErrorAction Stop
    }
    catch {
        if ($_.CategoryInfo.Category -eq [System.Management.Automation.ErrorCategory]::ObjectNotFound) {
            return $null
        }
        throw
    }

    if (-not $target.PSIsContainer -or $target -isnot [System.IO.DirectoryInfo] -or
        -not [string]::Equals($target.Name, $COMMON_MODULES_REPO_DIR_NAME, [System.StringComparison]::Ordinal) -or
        (Test-ReparsePoint -Item $target)) {
        throw "Admitted target is no longer an ordinary non-reparse '$COMMON_MODULES_REPO_DIR_NAME' directory: $TargetPath"
    }

    return $target
}

function Get-DistributionTargetPreflight {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.DirectoryInfo]$TargetDirectory
    )

    $pending = New-Object 'System.Collections.Generic.Queue[System.IO.DirectoryInfo]'
    $pending.Enqueue($TargetDirectory)
    $root_inventory = @()
    $is_first = $true

    while ($pending.Count -gt 0) {
        $directory = $pending.Dequeue()
        $inventory = @(Get-ChildItem -LiteralPath $directory.FullName -Force -ErrorAction Stop)
        if ($is_first) {
            $root_inventory = $inventory
            $is_first = $false
        }

        foreach ($item in $inventory) {
            if (Test-ReparsePoint -Item $item) {
                throw "Target contains a reparse entry and cannot be changed safely: $($item.FullName)"
            }
            if ($item.PSIsContainer) {
                if ($item -isnot [System.IO.DirectoryInfo]) {
                    throw "Target contains a nonordinary directory entry: $($item.FullName)"
                }
                $pending.Enqueue($item)
            }
            elseif ($item -isnot [System.IO.FileInfo]) {
                throw "Target contains a nonordinary file entry: $($item.FullName)"
            }
        }
    }

    return [pscustomobject]@{
        RootInventory = $root_inventory
    }
}

function Test-DistributionPackageMatch {
    param(
        [Parameter(Mandatory = $true)]
        [object]$SourcePackage,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$TargetInventory
    )

    if ($TargetInventory.Count -ne $SourcePackage.Entries.Count) {
        return $false
    }

    foreach ($source_entry in $SourcePackage.Entries) {
        $matches = @($TargetInventory | Where-Object {
            [string]::Equals($_.Name, $source_entry.Name, [System.StringComparison]::Ordinal)
        })
        if ($matches.Count -ne 1 -or $matches[0].PSIsContainer -or $matches[0] -isnot [System.IO.FileInfo]) {
            return $false
        }

        $target_file = $matches[0]
        $target_file.Refresh()
        if (-not $target_file.Exists -or
            $target_file.LastWriteTimeUtc.Ticks -ne $source_entry.LastWriteTimeUtc.Ticks -or
            $target_file.Length -ne $source_entry.Length) {
            return $false
        }
    }

    return $true
}

function Clear-DistributionTarget {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$TargetInventory
    )

    $ordered_items = @(Get-OrdinalSortedItems -Items $TargetInventory -PropertyName 'FullName')
    foreach ($item in $ordered_items) {
        Remove-Item -LiteralPath $item.FullName -Recurse -Force -ErrorAction Stop
    }
}

function Copy-DistributionPackage {
    param(
        [Parameter(Mandatory = $true)]
        [object]$SourcePackage,

        [Parameter(Mandatory = $true)]
        [System.IO.DirectoryInfo]$TargetDirectory
    )

    foreach ($source_entry in $SourcePackage.Entries) {
        $destination_path = Join-Path $TargetDirectory.FullName $source_entry.Name
        Copy-Item -LiteralPath $source_entry.FullName -Destination $destination_path -Force -ErrorAction Stop
        Write-Information "Copied '$($source_entry.Name)' to '$($TargetDirectory.FullName)'."
    }
}

try {
    $search_root = Resolve-DistributionSearchRoot -Arguments $args -InvocationWorkingDirectory $invocation_working_directory
    $search_root_inventory = @(Get-ChildItem -LiteralPath $search_root.FullName -Force -ErrorAction Stop)
    $source_repository = Get-WrapperSourceRepository -InvocationWorkingDirectory $invocation_working_directory
    $source_package = Read-DistributionSourcePackage -SourceRepository $source_repository
    $discovery = Find-DistributionCandidates -SearchRoot $search_root -SearchRootInventory $search_root_inventory -SourceRepository $source_repository
}
catch {
    throw "Distribution global failure: $($_.Exception.Message)"
}

Write-Host "Source: $($source_repository.FullName)"
Write-Host "Distribution Search Root: $($search_root.FullName)"

if ($discovery.TargetCount -eq 0 -and $discovery.FailureCount -eq 0) {
    throw "No eligible distribution target was found: $($search_root.FullName)"
}

$has_candidate_failure = $false
foreach ($candidate in $discovery.Candidates) {
    if ([string]::Equals($candidate.Kind, 'Failure', [System.StringComparison]::Ordinal)) {
        $has_candidate_failure = $true
        Write-Error -Message "Distribution candidate failure for '$($candidate.TargetPath)': $($candidate.Message)" -ErrorAction Continue
        continue
    }

    try {
        $target_directory = Get-CurrentDistributionTarget -TargetPath $candidate.TargetPath
    }
    catch {
        $has_candidate_failure = $true
        Write-Error -Message "Distribution candidate failure for '$($candidate.TargetPath)': $($_.Exception.Message)" -ErrorAction Continue
        continue
    }
    if ($null -eq $target_directory) {
        Write-Warning "Admitted distribution target disappeared before its turn and was skipped: $($candidate.TargetPath)"
        continue
    }

    try {
        Assert-DistributionSourceSnapshot -SourcePackage $source_package
    }
    catch {
        throw "Distribution global failure: the source package changed or became unreadable after validation. $($_.Exception.Message)"
    }

    try {
        $target_preflight = Get-DistributionTargetPreflight -TargetDirectory $target_directory
        $is_unchanged = Test-DistributionPackageMatch -SourcePackage $source_package -TargetInventory $target_preflight.RootInventory
    }
    catch {
        $has_candidate_failure = $true
        Write-Error -Message "Distribution candidate failure for '$($candidate.TargetPath)': $($_.Exception.Message)" -ErrorAction Continue
        continue
    }

    if ($is_unchanged) {
        Write-Host "UNCHANGED: $($target_directory.FullName)"
        continue
    }

    Write-Host "Updating distribution target: $($target_directory.FullName)"
    try {
        Clear-DistributionTarget -TargetInventory $target_preflight.RootInventory
    }
    catch {
        $has_candidate_failure = $true
        Write-Error -Message "Distribution candidate failure while clearing '$($candidate.TargetPath)': $($_.Exception.Message)" -ErrorAction Continue
        continue
    }

    try {
        Copy-DistributionPackage -SourcePackage $source_package -TargetDirectory $target_directory
    }
    catch {
        $copy_error = $_
        try {
            Assert-DistributionSourceSnapshot -SourcePackage $source_package
        }
        catch {
            throw "Distribution global failure: the source package changed or became unreadable during copy. $($_.Exception.Message)"
        }

        $has_candidate_failure = $true
        Write-Error -Message "Distribution candidate failure while copying to '$($candidate.TargetPath)': $($copy_error.Exception.Message)" -ErrorAction Continue
        continue
    }
}

if ($has_candidate_failure) {
    throw 'One or more distribution candidates failed.'
}

Write-Host 'CommonModules package distribution completed.'
