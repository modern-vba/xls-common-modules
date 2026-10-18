<#
.SYNOPSIS
Resolves a path string into an absolute directory path, file name, and full path.

.DESCRIPTION
Expands environment variables in %VAR% format and resolves relative paths against
BasePath, or the current PowerShell location when BasePath is omitted.
Undefined environment variables are preserved as %VAR%.
Existing paths are classified by their actual type. Nonexistent paths are treated
as directories if the expanded path ends with \ or /, and as files otherwise.
Paths are treated literally; wildcards and custom escape sequences are not interpreted.
Only paths provided by the FileSystem provider are supported. No items are created.

.PARAMETER Path
The path string to resolve. Required, at position 0. Accepts pipeline input,
absolute paths, relative paths, and environment variables in %VAR% format.
Null, empty, and whitespace-only values are rejected.

.PARAMETER BasePath
The base directory for relative paths, at position 1. Defaults to the current
PowerShell location. Supports relative paths, %VAR% expansion, and absolute paths
to nonexistent directories or drives. A relative BasePath is resolved against the
current PowerShell location. Existing files and non-FileSystem paths are rejected.
Does not affect absolute paths, drive-qualified paths, root-relative paths, or
home-relative paths starting with ~. Drive-relative paths such as D:report.csv
fail when the drive is unavailable.

.INPUTS
System.String
You can pipe path strings to this function.

.OUTPUTS
System.Management.Automation.PSCustomObject
One object per input, with Directory, File, and Path string properties.
Directory is the absolute directory path. File is the file name, or an empty
string for a directory. Path is the full absolute path and equals Directory for
a directory. Trailing directory separators are removed except at a root.

.EXAMPLE
Resolve-FileSystemPath -Path '%TEMP%\report.csv'
Expands TEMP and returns the absolute directory, file name, and full path.

.EXAMPLE
Resolve-FileSystemPath -Path '.\new-folder\'
Returns an absolute directory path even if the directory does not exist.

.EXAMPLE
Resolve-FileSystemPath -Path '.\report.csv' -BasePath 'C:\work'
Resolves the path as C:\work\report.csv, independently of the current directory.

.EXAMPLE
'.\first.csv', '.\second.csv' | Resolve-FileSystemPath -BasePath 'C:\work'
Returns one path-information object for each input string.

.NOTES
Requires a current FileSystem location when relative paths need that location.
Path resolution, provider access, and invalid path errors are not suppressed.

.LINK
ConvertTo-FileSystemSafeName
#>
function Resolve-FileSystemPath {
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param (
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Path,

        [Parameter(Position = 1)]
        [ValidateNotNullOrEmpty()]
        [string]$BasePath
    )

    begin {
        $resolvedBasePath = $null
        if ($PSBoundParameters.ContainsKey('BasePath')) {
            if ([string]::IsNullOrWhiteSpace($BasePath)) {
                throw [ArgumentException]::new('BasePath must not contain only whitespace.')
            }
            $baseProviderInfo = $null
            $baseDriveInfo = $null
            $expandedBasePath = [Environment]::ExpandEnvironmentVariables($BasePath)
            try {
                $resolvedBasePath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath(
                    $expandedBasePath, [ref]$baseProviderInfo, [ref]$baseDriveInfo
                )
            }
            catch [System.Management.Automation.DriveNotFoundException] {
                if ($expandedBasePath -notmatch '^[a-zA-Z]:[\\/]') { throw }
                # Resolve absolute paths on unavailable drives without requiring a PSDrive.
                $resolvedBasePath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath(
                    ('FileSystem::' + $expandedBasePath), [ref]$baseProviderInfo, [ref]$baseDriveInfo
                )
            }
            if ($baseProviderInfo.Name -ne 'FileSystem') {
                throw [ArgumentException]::new('BasePath must be a file system path.')
            }
            $resolvedBasePath = [IO.Path]::GetFullPath($resolvedBasePath)
            if (Test-Path -LiteralPath $resolvedBasePath -ErrorAction Stop) {
                $baseItem = Get-Item -LiteralPath $resolvedBasePath -Force -ErrorAction Stop
                if (-not $baseItem.PSIsContainer) {
                    throw [ArgumentException]::new('BasePath must be a directory, not a file.')
                }
            }
        }
    }

    process {
        if ([string]::IsNullOrWhiteSpace($Path)) {
            throw [ArgumentException]::new('Path must not contain only whitespace.')
        }

        $expandedPath = [Environment]::ExpandEnvironmentVariables($Path)
        $pathToResolve = $expandedPath
        if ($null -ne $resolvedBasePath) {
            $pathDriveName = $null
            $isAbsolute = $ExecutionContext.SessionState.Path.IsPSAbsolute($expandedPath, [ref]$pathDriveName)
            $isHomeRelative = $expandedPath -match '^~([\\/]|$)'
            if (-not $isAbsolute -and -not [IO.Path]::IsPathRooted($expandedPath) -and -not $isHomeRelative) {
                $pathToResolve = [IO.Path]::Combine($resolvedBasePath, $expandedPath)
            }
        }
        $providerInfo = $null
        $driveInfo = $null
        try {
            $providerPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath(
                $pathToResolve, [ref]$providerInfo, [ref]$driveInfo
            )
        }
        catch [System.Management.Automation.DriveNotFoundException] {
            if ($pathToResolve -notmatch '^[a-zA-Z]:[\\/]') { throw }
            $providerPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath(
                ('FileSystem::' + $pathToResolve), [ref]$providerInfo, [ref]$driveInfo
            )
        }
        if ($providerInfo.Name -ne 'FileSystem') {
            throw [ArgumentException]::new('Path must be a file system path.')
        }

        $absolutePath = [IO.Path]::GetFullPath($providerPath)
        if (Test-Path -LiteralPath $absolutePath -ErrorAction Stop) {
            $item = Get-Item -LiteralPath $absolutePath -Force -ErrorAction Stop
            $isDirectory = $item.PSIsContainer
        }
        else {
            $isDirectory = $expandedPath.EndsWith('\') -or $expandedPath.EndsWith('/')
        }

        $rootPath = [IO.Path]::GetPathRoot($absolutePath)
        if ($absolutePath.Length -gt $rootPath.Length) {
            $absolutePath = $absolutePath.TrimEnd([char[]]'\/')
        }

        if ($isDirectory) {
            $directory = $absolutePath
            $file = ''
        }
        else {
            $directory = [IO.Path]::GetDirectoryName($absolutePath)
            $file = [IO.Path]::GetFileName($absolutePath)
        }

        [pscustomobject][ordered]@{
            Directory = $directory
            File = $file
            Path = $absolutePath
        }
    }
}
