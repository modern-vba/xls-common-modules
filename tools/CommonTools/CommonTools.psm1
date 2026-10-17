# Load the public functions.
. "$PSScriptRoot\Resolve-FileSystemPath.ps1"
. "$PSScriptRoot\ConvertTo-FileSystemSafeName.ps1"

Export-ModuleMember -Function `
        Resolve-FileSystemPath, `
        ConvertTo-FileSystemSafeName
