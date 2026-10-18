<#
.SYNOPSIS
Replaces forbidden Windows file name punctuation with fullwidth characters.

.DESCRIPTION
Converts the characters \ / : * ? " < > | to their fullwidth equivalents.
Accepts one file or directory name, rather than an entire path, and returns one
converted string for each input. An empty input string produces an empty string.
Other characters, including control characters, reserved device names, trailing
spaces, and trailing periods, are left unchanged. This is a punctuation conversion,
not a complete validation of a Windows file name.

.PARAMETER Name
The file or directory name to convert. Required, at position 0. Accepts empty
strings and string values from the pipeline.

.INPUTS
System.String
You can pipe file or directory names to this function.

.OUTPUTS
System.String
The input name with the nine supported punctuation characters replaced.

.EXAMPLE
ConvertTo-FileSystemSafeName -Name 'Report: Q1?.txt'
Returns 'Report： Q1？.txt'.

.EXAMPLE
'one/two', 'three|four' | ConvertTo-FileSystemSafeName
Returns 'one／two' and 'three｜four'.

.NOTES
Does not create, rename, or inspect any file or directory. Passing an entire path
also converts its drive colon and directory separators.

.LINK
Resolve-FileSystemPath
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
