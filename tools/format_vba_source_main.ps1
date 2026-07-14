param(
    [Parameter(Mandatory = $true, Position = 0, ValueFromRemainingArguments = $true)]
    [string[]] $Path,

    [switch] $Recurse,

    [switch] $Check
)

$ErrorActionPreference = 'Stop'
$InformationPreference = 'Continue'

Set-Variable -Name C_VBA_ENCODING_CODEPAGE -Value 932 -Option Constant

function Get-VbaSourceFiles {
    param(
        [string[]] $InputPath,
        [bool] $ScanRecurse
    )

    $result = @()
    foreach ($path_item in $InputPath) {
        $resolved_items = Get-Item -LiteralPath $path_item -ErrorAction Stop
        foreach ($resolved_item in $resolved_items) {
            if ($resolved_item.PSIsContainer) {
                $child_items = Get-ChildItem -LiteralPath $resolved_item.FullName -File -Recurse:$ScanRecurse
                $result += $child_items | Where-Object { $_.Extension -in @('.bas', '.cls') }
            }
            elseif ($resolved_item.Extension -in @('.bas', '.cls')) {
                $result += $resolved_item
            }
        }
    }

    $result | Sort-Object FullName -Unique
}

function Get-MemberStartInfo {
    param([string] $Line)

    $pattern = '^\s*(?:(Public|Private|Friend)\s+)?(?:(Static)\s+)?(Sub|Function|Property\s+(Get|Let|Set))\s+([A-Za-z_][A-Za-z0-9_]*)\b'
    $match = [regex]::Match($Line, $pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if (-not $match.Success) { return $null }

    $kind_text = $match.Groups[3].Value
    $member_kind = $kind_text
    $accessor = ''
    if ($kind_text -match '^Property\s+(Get|Let|Set)$') {
        $member_kind = 'Property'
        $accessor = $Matches[1]
    }

    $scope = $match.Groups[1].Value
    if ([string]::IsNullOrWhiteSpace($scope)) { $scope = 'Public' }

    [pscustomobject]@{
        Scope = $scope
        Kind = $member_kind
        Accessor = $accessor
        Name = $match.Groups[5].Value
    }
}

function Get-MemberEndPattern {
    param([string] $MemberKind)

    if ($MemberKind -eq 'Property') {
        return '^\s*End\s+Property\s*$'
    }

    '^\s*End\s+' + [regex]::Escape($MemberKind) + '\s*$'
}

function Get-FirstMemberLineIndex {
    param([string[]] $Lines)

    for ($line_idx = 0; $line_idx -lt $Lines.Count; $line_idx++) {
        if ($null -ne (Get-MemberStartInfo $Lines[$line_idx])) {
            return $line_idx
        }
    }

    -1
}

function Test-ModuleOverviewBoundaryLine {
    param([string] $Line)

    $Line -match "^\s*'\s*#{5,}\s*$"
}

function Get-ModuleOverviewBlockInfo {
    param([string[]] $Lines)

    for ($line_idx = 0; $line_idx -lt $Lines.Count; $line_idx++) {
        if (-not (Test-ModuleOverviewBoundaryLine $Lines[$line_idx])) { continue }

        $has_brief = $false
        for ($end_idx = $line_idx; $end_idx -lt $Lines.Count; $end_idx++) {
            if ($Lines[$end_idx] -match "^\s*'!\s*@brief\s*$") {
                $has_brief = $true
            }

            if ($end_idx -gt $line_idx -and (Test-ModuleOverviewBoundaryLine $Lines[$end_idx])) {
                if ($has_brief) {
                    return [pscustomobject]@{
                        StartIndex = $line_idx
                        EndIndex = $end_idx
                    }
                }
                break
            }
        }
    }

    $null
}

function Remove-LineRange {
    param(
        [string[]] $Lines,
        [int] $StartIndex,
        [int] $EndIndex
    )

    $result = New-Object 'System.Collections.Generic.List[string]'
    for ($line_idx = 0; $line_idx -lt $Lines.Count; $line_idx++) {
        if ($line_idx -ge $StartIndex -and $line_idx -le $EndIndex) { continue }
        $result.Add($Lines[$line_idx])
    }

    while ($StartIndex -gt 0 -and $StartIndex -lt $result.Count -and $result[$StartIndex - 1].Trim() -eq '' -and $result[$StartIndex].Trim() -eq '') {
        $result.RemoveAt($StartIndex)
    }

    $result.ToArray()
}

function Get-AttachedLeadingStartIndex {
    param(
        [string[]] $Lines,
        [int] $MemberIndex
    )

    if ($MemberIndex -le 0) { return $MemberIndex }
    if (-not (Test-AttachedCommentLine $Lines[$MemberIndex - 1])) { return $MemberIndex }

    $line_idx = $MemberIndex - 1
    while ($line_idx -ge 0 -and (Test-AttachedCommentLine $Lines[$line_idx])) {
        $line_idx--
    }

    $line_idx + 1
}

function Get-ClassModuleOverviewInsertIndex {
    param([string[]] $Lines)

    $line_idx = 0
    for ($scan_idx = 0; $scan_idx -lt $Lines.Count; $scan_idx++) {
        if ($Lines[$scan_idx] -match '^\s*Option\s+') {
            $line_idx = $scan_idx + 1
        }
    }

    if ($line_idx -eq 0) {
        for ($scan_idx = 0; $scan_idx -lt $Lines.Count; $scan_idx++) {
            if ($Lines[$scan_idx] -match '^\s*Attribute\s+') {
                $line_idx = $scan_idx + 1
            }
        }
    }

    while ($line_idx -lt $Lines.Count) {
        if ($Lines[$line_idx].Trim() -eq '') {
            $line_idx++
            continue
        }
        if ($Lines[$line_idx] -match "^\s*'\#Interface\s*$") {
            $line_idx++
            continue
        }
        if ($Lines[$line_idx] -match '^\s*Implements\s+') {
            $line_idx++
            continue
        }
        break
    }

    $line_idx
}

function Get-StandardModuleOverviewInsertIndex {
    param([string[]] $Lines)

    $line_idx = 0
    while ($line_idx -lt $Lines.Count) {
        if ($Lines[$line_idx] -match '^\s*Attribute\s+') {
            $line_idx++
            continue
        }
        if ($Lines[$line_idx] -match '^\s*Option\s+') {
            $line_idx++
            continue
        }
        if ($Lines[$line_idx].Trim() -eq '') {
            $line_idx++
            continue
        }
        break
    }

    $line_idx
}

function Get-ModuleOverviewInsertIndex {
    param(
        [string[]] $Lines,
        [bool] $IsClassModule
    )

    if ($IsClassModule) {
        return Get-ClassModuleOverviewInsertIndex -Lines $Lines
    }

    Get-StandardModuleOverviewInsertIndex -Lines $Lines
}

function Insert-ModuleOverviewBlock {
    param(
        [string[]] $Lines,
        [string[]] $OverviewLines,
        [int] $InsertIndex
    )

    $result = New-Object 'System.Collections.Generic.List[string]'
    foreach ($line in $Lines) {
        $result.Add($line)
    }

    while ($InsertIndex -gt 0 -and $result[$InsertIndex - 1].Trim() -eq '') {
        $result.RemoveAt($InsertIndex - 1)
        $InsertIndex--
    }
    while ($InsertIndex -lt $result.Count -and $result[$InsertIndex].Trim() -eq '') {
        $result.RemoveAt($InsertIndex)
    }

    $insert_lines = New-Object 'System.Collections.Generic.List[string]'
    if ($InsertIndex -gt 0) {
        $insert_lines.Add('')
    }
    foreach ($line in $OverviewLines) {
        $insert_lines.Add($line)
    }
    if ($InsertIndex -lt $result.Count) {
        $insert_lines.Add('')
    }

    $result.InsertRange($InsertIndex, $insert_lines)
    $result.ToArray()
}

function Get-OverviewPrecedingAttachedStartIndex {
    param(
        [string[]] $Lines,
        [int] $OverviewStartIndex
    )

    $line_idx = $OverviewStartIndex - 1
    while ($line_idx -ge 0 -and $Lines[$line_idx].Trim() -eq '') {
        $line_idx--
    }
    if ($line_idx -lt 0 -or $Lines[$line_idx] -notmatch "^\s*'\*") { return -1 }

    while ($line_idx -ge 0 -and (Test-AttachedCommentLine $Lines[$line_idx])) {
        $line_idx--
    }

    $line_idx + 1
}

function Remove-BlankBetweenAttachedCommentAndMember {
    param([string[]] $Lines)

    $result = New-Object 'System.Collections.Generic.List[string]'
    $line_idx = 0
    while ($line_idx -lt $Lines.Count) {
        if ($Lines[$line_idx].Trim() -eq '' -and $result.Count -gt 0 -and $result[$result.Count - 1] -match "^\s*'\*") {
            $next_idx = $line_idx + 1
            while ($next_idx -lt $Lines.Count -and $Lines[$next_idx].Trim() -eq '') {
                $next_idx++
            }
            if ($next_idx -lt $Lines.Count -and $null -ne (Get-MemberStartInfo $Lines[$next_idx])) {
                $line_idx = $next_idx
                continue
            }
        }

        $result.Add($Lines[$line_idx])
        $line_idx++
    }

    $result.ToArray()
}

function Move-MisplacedModuleOverviewBlock {
    param(
        [string[]] $Lines,
        [bool] $IsClassModule
    )

    $overview_info = Get-ModuleOverviewBlockInfo -Lines $Lines
    if ($null -eq $overview_info) { return $Lines }

    $overview_lines = $Lines[$overview_info.StartIndex..$overview_info.EndIndex]
    $remaining_lines = Remove-LineRange -Lines $Lines -StartIndex $overview_info.StartIndex -EndIndex $overview_info.EndIndex
    $insert_idx = Get-ModuleOverviewInsertIndex -Lines $remaining_lines -IsClassModule $IsClassModule
    $result_lines = Insert-ModuleOverviewBlock -Lines $remaining_lines -OverviewLines $overview_lines -InsertIndex $insert_idx

    Remove-BlankBetweenAttachedCommentAndMember -Lines $result_lines
}

function Test-AttachedCommentLine {
    param([string] $Line)

    $Line -match "^\s*'"
}

function Pop-AttachedLeadingLines {
    param([System.Collections.Generic.List[string]] $PendingLines)

    $attached_lines = New-Object 'System.Collections.Generic.List[string]'
    if ($PendingLines.Count -eq 0) { return $attached_lines }
    if (-not (Test-AttachedCommentLine $PendingLines[$PendingLines.Count - 1])) { return $attached_lines }

    $start_idx = $PendingLines.Count - 1
    while ($start_idx -ge 0 -and (Test-AttachedCommentLine $PendingLines[$start_idx])) {
        $start_idx--
    }
    $start_idx++

    for ($line_idx = $start_idx; $line_idx -lt $PendingLines.Count; $line_idx++) {
        $attached_lines.Add($PendingLines[$line_idx])
    }

    $PendingLines.RemoveRange($start_idx, $PendingLines.Count - $start_idx)
    $attached_lines
}

function Test-WhitespaceBlock {
    param([object] $Block)

    if ($Block.Kind -ne 'Text') { return $false }
    foreach ($line in $Block.Lines) {
        if ($line.Trim() -ne '') { return $false }
    }
    $true
}

function Get-InterfaceMemberTargetName {
    param([string] $MemberName)

    $underscore_idx = $MemberName.IndexOf('_')
    if ($underscore_idx -lt 0 -or $underscore_idx -eq $MemberName.Length - 1) { return '' }
    $MemberName.Substring($underscore_idx + 1)
}

function Test-InterfaceImplementationMember {
    param([object] $Member)

    if ($Member.Scope -ne 'Private') { return $false }
    $underscore_idx = $Member.Name.IndexOf('_')
    if ($underscore_idx -lt 1 -or $underscore_idx -eq $Member.Name.Length - 1) { return $false }

    $interface_name = $Member.Name.Substring(0, $underscore_idx)
    if ($interface_name -notmatch '^I[A-Za-z0-9]+$') { return $false }

    $true
}

function Test-SuppressBlankBetweenMembers {
    param(
        [object] $CurrentMember,
        [object] $NextMember
    )

    if ($CurrentMember.Kind -eq 'Property' -and $NextMember.Kind -eq 'Property' -and $CurrentMember.Name -eq $NextMember.Name) {
        return $true
    }

    if ($CurrentMember.Scope -ne 'Private' -and (Test-InterfaceImplementationMember $NextMember)) {
        $target_name = Get-InterfaceMemberTargetName $NextMember.Name
        if ($CurrentMember.Kind -eq $NextMember.Kind -and $CurrentMember.Name -eq $target_name) {
            return $true
        }
    }

    $false
}

function New-TextBlock {
    param([string[]] $Lines)

    [pscustomobject]@{
        Kind = 'Text'
        Lines = @($Lines)
    }
}

function New-MemberBlock {
    param(
        [string[]] $Lines,
        [object] $MemberInfo
    )

    [pscustomobject]@{
        Kind = 'Member'
        Lines = @($Lines)
        Scope = $MemberInfo.Scope
        MemberKind = $MemberInfo.Kind
        Accessor = $MemberInfo.Accessor
        Name = $MemberInfo.Name
    }
}

function ConvertTo-MemberShape {
    param([object] $Block)

    [pscustomobject]@{
        Scope = $Block.Scope
        Kind = $Block.MemberKind
        Accessor = $Block.Accessor
        Name = $Block.Name
    }
}

function Test-InterfaceClassModule {
    param([object[]] $Blocks)

    foreach ($block in $Blocks) {
        if ($block.Kind -ne 'Text') { continue }
        foreach ($line in $block.Lines) {
            if ($line -match "^\s*'\#Interface\s*$") { return $true }
        }
    }

    $false
}

function Sort-PropertyMemberBlocks {
    param([object[]] $MemberBlocks)

    $ordered_blocks = New-Object 'System.Collections.Generic.List[object]'
    foreach ($accessor in @('Get', 'Let', 'Set')) {
        foreach ($block in $MemberBlocks) {
            if ($block.Accessor -eq $accessor) {
                $ordered_blocks.Add($block)
            }
        }
    }

    $ordered_blocks.ToArray()
}

function New-MemberGroup {
    param(
        [object[]] $MemberBlocks,
        [int] $Ordinal
    )

    if ($MemberBlocks[0].MemberKind -eq 'Property') {
        $MemberBlocks = Sort-PropertyMemberBlocks -MemberBlocks $MemberBlocks
    }

    $first_block = $MemberBlocks[0]
    [pscustomobject]@{
        Blocks = @($MemberBlocks)
        AttachedGroups = @()
        Ordinal = $Ordinal
        Scope = $first_block.Scope
        MemberKind = $first_block.MemberKind
        Name = $first_block.Name
        SortKey = 0
    }
}

function New-MemberGroups {
    param([object[]] $MemberBlocks)

    $groups = New-Object 'System.Collections.Generic.List[object]'
    $block_idx = 0
    $ordinal = 0
    while ($block_idx -lt $MemberBlocks.Count) {
        $group_blocks = New-Object 'System.Collections.Generic.List[object]'
        $current_block = $MemberBlocks[$block_idx]
        $group_blocks.Add($current_block)
        $block_idx++

        if ($current_block.MemberKind -eq 'Property') {
            while ($block_idx -lt $MemberBlocks.Count) {
                $next_block = $MemberBlocks[$block_idx]
                if ($next_block.MemberKind -ne 'Property' -or $next_block.Name -ne $current_block.Name) { break }
                $group_blocks.Add($next_block)
                $block_idx++
            }
        }

        $groups.Add((New-MemberGroup -MemberBlocks $group_blocks.ToArray() -Ordinal $ordinal))
        $ordinal++
    }

    $groups.ToArray()
}

function ConvertTo-GroupShape {
    param([object] $Group)

    [pscustomobject]@{
        Scope = $Group.Scope
        Kind = $Group.MemberKind
        Accessor = ''
        Name = $Group.Name
    }
}

function Test-InterfaceImplementationGroup {
    param([object] $Group)

    Test-InterfaceImplementationMember (ConvertTo-GroupShape $Group)
}

function Attach-InterfaceImplementationGroups {
    param([object[]] $Groups)

    $used_ordinals = New-Object 'System.Collections.Generic.HashSet[int]'
    foreach ($group in $Groups) {
        if ($group.Scope -eq 'Private') { continue }

        foreach ($candidate in $Groups) {
            if (-not (Test-InterfaceImplementationGroup $candidate)) { continue }
            if ($candidate.MemberKind -ne $group.MemberKind) { continue }

            $target_name = Get-InterfaceMemberTargetName $candidate.Name
            if ($target_name -ne $group.Name) { continue }

            $group.AttachedGroups = @($group.AttachedGroups + $candidate)
            [void]$used_ordinals.Add($candidate.Ordinal)
        }
    }

    $result = New-Object 'System.Collections.Generic.List[object]'
    foreach ($group in $Groups) {
        if ($used_ordinals.Contains($group.Ordinal)) { continue }
        $result.Add($group)
    }

    $result.ToArray()
}

function Get-MemberGroupText {
    param([object] $Group)

    $lines = New-Object 'System.Collections.Generic.List[string]'
    foreach ($block in $Group.Blocks) {
        foreach ($line in $block.Lines) {
            $lines.Add($line)
        }
    }

    $lines -join "`n"
}

function Test-MemberGroupReferencesName {
    param(
        [object] $Group,
        [string] $Name
    )

    $pattern = '(?i)(^|[^A-Za-z0-9_])' + [regex]::Escape($Name) + '([^A-Za-z0-9_]|$)'
    (Get-MemberGroupText $Group) -match $pattern
}

function Test-PrivateHelperGroup {
    param([object] $Group)

    if ($Group.Scope -ne 'Private') { return $false }
    if ($Group.MemberKind -notin @('Function', 'Sub')) { return $false }
    if (Test-InterfaceImplementationGroup $Group) { return $false }

    $Group.Name -match '^p[A-Z]'
}

function Test-InitializeGroup {
    param([object] $Group)

    $Group.Name -match '^Initialize'
}

function Test-PublicCallableGroup {
    param([object] $Group)

    if ($Group.Scope -eq 'Private') { return $false }
    $Group.MemberKind -in @('Function', 'Sub')
}

function Test-HelperAttachRootGroup {
    param([object] $Group)

    if ($Group.Name -eq 'Class_Initialize') { return $true }
    if ($Group.Name -eq 'Class_Terminate') { return $true }
    if (Test-InitializeGroup $Group) { return $true }
    if (Test-PublicCallableGroup $Group) { return $true }

    $false
}

function Attach-PrivateHelperGroups {
    param([object[]] $Groups)

    $root_groups = @($Groups | Where-Object { Test-HelperAttachRootGroup $_ })
    $used_ordinals = New-Object 'System.Collections.Generic.HashSet[int]'

    foreach ($helper_group in @($Groups | Where-Object { Test-PrivateHelperGroup $_ })) {
        $used_by = @($root_groups | Where-Object {
            $_.Ordinal -ne $helper_group.Ordinal -and (Test-MemberGroupReferencesName $_ $helper_group.Name)
        })

        if ($used_by.Count -eq 0) { continue }

        $target_group = $null
        $init_users = @($used_by | Where-Object { Test-InitializeGroup $_ })
        if ($init_users.Count -eq $used_by.Count) {
            $target_group = @($Groups | Where-Object { Test-InitializeGroup $_ } | Sort-Object Ordinal | Select-Object -Last 1)[0]
        }
        elseif ($used_by.Count -eq 1) {
            $target_group = $used_by[0]
        }

        if ($null -eq $target_group) { continue }

        $target_group.AttachedGroups = @($target_group.AttachedGroups + $helper_group)
        [void]$used_ordinals.Add($helper_group.Ordinal)
    }

    $result = New-Object 'System.Collections.Generic.List[object]'
    foreach ($group in $Groups) {
        if ($used_ordinals.Contains($group.Ordinal)) { continue }
        $result.Add($group)
    }

    $result.ToArray()
}

function Get-ClassMemberGroupSortKey {
    param([object] $Group)

    if ($Group.Name -eq 'Item') { return 20 }
    if ($Group.MemberKind -eq 'Property') { return 10 }
    if ($Group.Name -eq 'Class_Initialize') { return 30 }
    if (Test-InitializeGroup $Group) { return 40 }
    if ($Group.Name -eq 'Class_Terminate') { return 50 }
    if (Test-PublicCallableGroup $Group) { return 60 }

    90
}

function Expand-MemberGroups {
    param([object[]] $Groups)

    $blocks = New-Object 'System.Collections.Generic.List[object]'
    foreach ($group in $Groups) {
        foreach ($block in $group.Blocks) {
            $blocks.Add($block)
        }
        foreach ($attached_group in $group.AttachedGroups) {
            foreach ($block in $attached_group.Blocks) {
                $blocks.Add($block)
            }
        }
    }

    $blocks.ToArray()
}

function Format-ClassMemberSegment {
    param([object[]] $SegmentBlocks)

    $first_member_idx = -1
    $last_member_idx = -1
    for ($block_idx = 0; $block_idx -lt $SegmentBlocks.Count; $block_idx++) {
        if ($SegmentBlocks[$block_idx].Kind -ne 'Member') { continue }
        if ($first_member_idx -lt 0) { $first_member_idx = $block_idx }
        $last_member_idx = $block_idx
    }

    if ($first_member_idx -lt 0) { return $SegmentBlocks }

    $member_blocks = @($SegmentBlocks | Where-Object { $_.Kind -eq 'Member' })
    if ($member_blocks.Count -le 1) { return $SegmentBlocks }

    $groups = New-MemberGroups -MemberBlocks $member_blocks
    $groups = Attach-InterfaceImplementationGroups -Groups $groups
    $groups = Attach-PrivateHelperGroups -Groups $groups

    foreach ($group in $groups) {
        $group.SortKey = Get-ClassMemberGroupSortKey $group
    }

    $ordered_member_blocks = Expand-MemberGroups (@($groups | Sort-Object SortKey, Ordinal))
    $result = New-Object 'System.Collections.Generic.List[object]'

    for ($block_idx = 0; $block_idx -lt $first_member_idx; $block_idx++) {
        $result.Add($SegmentBlocks[$block_idx])
    }
    foreach ($block in $ordered_member_blocks) {
        $result.Add($block)
    }
    for ($block_idx = $last_member_idx + 1; $block_idx -lt $SegmentBlocks.Count; $block_idx++) {
        $result.Add($SegmentBlocks[$block_idx])
    }

    $result.ToArray()
}

function Reorder-ClassBlocks {
    param([object[]] $Blocks)

    if (Test-InterfaceClassModule $Blocks) { return $Blocks }

    $result = New-Object 'System.Collections.Generic.List[object]'
    $segment = New-Object 'System.Collections.Generic.List[object]'

    foreach ($block in $Blocks) {
        if ($block.Kind -eq 'Member' -or (Test-WhitespaceBlock $block)) {
            $segment.Add($block)
            continue
        }

        if ($segment.Count -gt 0) {
            foreach ($segment_block in (Format-ClassMemberSegment $segment.ToArray())) {
                $result.Add($segment_block)
            }
            $segment.Clear()
        }
        $result.Add($block)
    }

    if ($segment.Count -gt 0) {
        foreach ($segment_block in (Format-ClassMemberSegment $segment.ToArray())) {
            $result.Add($segment_block)
        }
    }

    $result.ToArray()
}

function Split-VbaSourceIntoBlocks {
    param([string[]] $Lines)

    $blocks = New-Object 'System.Collections.Generic.List[object]'
    $pending_lines = New-Object 'System.Collections.Generic.List[string]'
    $line_idx = 0

    while ($line_idx -lt $Lines.Count) {
        $member_info = Get-MemberStartInfo $Lines[$line_idx]
        if ($null -eq $member_info) {
            $pending_lines.Add($Lines[$line_idx])
            $line_idx++
            continue
        }

        $attached_lines = Pop-AttachedLeadingLines $pending_lines
        if ($pending_lines.Count -gt 0) {
            $blocks.Add((New-TextBlock -Lines $pending_lines.ToArray()))
            $pending_lines.Clear()
        }

        $member_lines = New-Object 'System.Collections.Generic.List[string]'
        foreach ($attached_line in $attached_lines) {
            $member_lines.Add($attached_line)
        }

        $end_pattern = Get-MemberEndPattern $member_info.Kind
        do {
            $member_lines.Add($Lines[$line_idx])
            $is_end_line = $Lines[$line_idx] -match $end_pattern
            $line_idx++
        } while (-not $is_end_line -and $line_idx -lt $Lines.Count)

        $blocks.Add((New-MemberBlock -Lines $member_lines.ToArray() -MemberInfo $member_info))
    }

    if ($pending_lines.Count -gt 0) {
        $blocks.Add((New-TextBlock -Lines $pending_lines.ToArray()))
    }

    $blocks
}

function Join-FormattedBlocks {
    param([object[]] $Blocks)

    $output_lines = New-Object 'System.Collections.Generic.List[string]'
    $block_idx = 0

    while ($block_idx -lt $Blocks.Count) {
        $block = $Blocks[$block_idx]
        foreach ($line in $block.Lines) {
            $output_lines.Add($line)
        }

        if ($block.Kind -eq 'Member') {
            $next_member_idx = -1
            $skip_next_whitespace = $false

            if ($block_idx + 1 -lt $Blocks.Count -and $Blocks[$block_idx + 1].Kind -eq 'Member') {
                $next_member_idx = $block_idx + 1
            }
            elseif ($block_idx + 2 -lt $Blocks.Count -and (Test-WhitespaceBlock $Blocks[$block_idx + 1]) -and $Blocks[$block_idx + 2].Kind -eq 'Member') {
                $next_member_idx = $block_idx + 2
                $skip_next_whitespace = $true
            }

            if ($next_member_idx -gt 0) {
                $current_member = ConvertTo-MemberShape $block
                $next_member = ConvertTo-MemberShape $Blocks[$next_member_idx]
                if (-not (Test-SuppressBlankBetweenMembers $current_member $next_member)) {
                    $output_lines.Add('')
                }
                if ($skip_next_whitespace) {
                    $block_idx++
                }
            }
        }

        $block_idx++
    }

    $output_lines.ToArray()
}

function Format-VbaSourceText {
    param(
        [string] $Text,
        [bool] $IsClassModule
    )

    $normalized_text = $Text -replace "`r`n", "`n" -replace "`r", "`n"
    $lines = @($normalized_text -split "`n", -1)
    if ($lines.Count -gt 0 -and $lines[$lines.Count - 1] -eq '') {
        if ($lines.Count -eq 1) {
            $lines = @()
        }
        else {
            $lines = $lines[0..($lines.Count - 2)]
        }
    }
    $lines = @($lines | ForEach-Object {
        $_ -replace '\s+$', ''
    })
    $lines = Move-MisplacedModuleOverviewBlock -Lines $lines -IsClassModule $IsClassModule
    $lines = Remove-BlankBetweenAttachedCommentAndMember -Lines $lines

    $blocks = Split-VbaSourceIntoBlocks -Lines $lines
    if ($IsClassModule) {
        $blocks = Reorder-ClassBlocks -Blocks $blocks
    }

    $formatted_lines = Join-FormattedBlocks -Blocks $blocks
    ($formatted_lines -join "`r`n") + "`r`n"
}

function Write-VbaSourceText {
    param(
        [string] $FilePath,
        [string] $Text,
        [System.Text.Encoding] $Encoding
    )

    try {
        [System.IO.File]::WriteAllText($FilePath, $Text, $Encoding)
    }
    catch [System.UnauthorizedAccessException] {
        $stream = [System.IO.File]::Open($FilePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::Read)
        try {
            $stream.SetLength(0)
            $writer = New-Object System.IO.StreamWriter($stream, $Encoding)
            try {
                $writer.Write($Text)
            }
            finally {
                $writer.Dispose()
            }
        }
        finally {
            if ($null -ne $stream) { $stream.Dispose() }
        }
    }
}

$encoding = [System.Text.Encoding]::GetEncoding($C_VBA_ENCODING_CODEPAGE)
$target_files = Get-VbaSourceFiles -InputPath $Path -ScanRecurse $Recurse.IsPresent
$changed_files = @()

foreach ($target_file in $target_files) {
    $original_text = [System.IO.File]::ReadAllText($target_file.FullName, $encoding)
    $formatted_text = Format-VbaSourceText -Text $original_text -IsClassModule ($target_file.Extension -eq '.cls')

    if ($formatted_text -ne $original_text) {
        $changed_files += $target_file.FullName
        if ($Check) {
            Write-Information "Needs format: $($target_file.FullName)"
        }
        else {
            Write-VbaSourceText -FilePath $target_file.FullName -Text $formatted_text -Encoding $encoding
            Write-Information "Formatted: $($target_file.FullName)"
        }
    }
}

if ($Check -and $changed_files.Count -gt 0) {
    throw "VBA source format check failed: $($changed_files.Count) file(s)."
}

Write-Information "Checked $($target_files.Count) VBA source file(s). Changed $($changed_files.Count) file(s)."
