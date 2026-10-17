<#
.SYNOPSIS
Replaces forbidden Windows file name punctuation with fullwidth characters.

.DESCRIPTION
Converts the characters \ / : * ? " < > | to their fullwidth equivalents.
Accepts one file or directory name, rather than an entire path.
Other characters, including control characters, reserved device names, trailing
spaces, and trailing periods, are left unchanged.

.PARAMETER Name
The file or directory name to convert. Accepts empty strings and pipeline input.

.OUTPUTS
System.String

.EXAMPLE
ConvertTo-FileSystemSafeName -Name 'Report: Q1?.txt'
Returns 'Report： Q1？.txt'.

.EXAMPLE
'one/two', 'three|four' | ConvertTo-FileSystemSafeName
Returns 'one／two' and 'three｜four'.
#>
function ConvertTo-FileSystemSafeName {
    [CmdletBinding()]
    [OutputType([string])]
    param (
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
        [AllowEmptyString()]
        [string]$Name
    )

    begin {
        $replacements = [ordered]@{
            '\' = '＼'
            '/' = '／'
            ':' = '：'
            '*' = '＊'
            '?' = '？'
            '"' = '＂'
            '<' = '＜'
            '>' = '＞'
            '|' = '｜'
        }
    }

    process {
        $result = $Name
        foreach ($character in $replacements.Keys) {
            $result = $result.Replace($character, $replacements[$character])
        }
        $result
    }
}
