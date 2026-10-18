<#
.SYNOPSIS
Exports one named VBA component from an Excel workbook.

.DESCRIPTION
Starts a hidden Excel instance, opens InputFile, and finds a component whose name
matches ModuleName. Standard modules are exported as .bas, class and document
modules as .cls, and UserForms as .frm with any companion file produced by VBIDE.
The workbook is closed without saving, and the Excel instance is quit.

.PARAMETER ModuleName
The VBA component name, without its file extension. Required, at position 0.
The name comparison is case-insensitive.

.PARAMETER InputFile
The path to the existing Excel workbook. Required, at position 1.

.PARAMETER OutputDirectory
The export directory, at position 2. A missing directory is created. Although the
parameter is optional in the signature, no default directory is implemented;
supply a directory explicitly.

.PARAMETER Force
Removes an existing main export file with the same name before exporting.
Without this switch, an existing main export file produces an error.
This does not pre-delete a UserForm's companion .frx file.

.INPUTS
None.
This function does not accept pipeline input.

.OUTPUTS
None.
Writes component files to OutputDirectory. Diagnostics use the error and debug streams.

.EXAMPLE
Export-ModuleFromExcelFile -ModuleName 'Common_Mod' -InputFile 'C:\work\Book.xlsm' -OutputDirectory 'C:\exports'
Exports Common_Mod to the specified directory.

.EXAMPLE
Export-ModuleFromExcelFile -ModuleName 'ThisWorkbook' -InputFile 'C:\work\Book.xlsm' -OutputDirectory 'C:\exports' -Force
Replaces an existing ThisWorkbook.cls export. Exporting a document module does not
make it directly importable as the workbook's document module.

.NOTES
Legacy Excel COM helper. Normal workspace import/export workflows use vba-dev.
Requires installed Excel, an accessible VBA project, and Trust access to the VBA
project object model. Close the target workbook before calling this function.
The current loop also writes errors for components other than the requested one,
so a successful export can be accompanied by unrelated-component errors.
Supports WhatIf and Confirm, but Excel instance creation and output-directory
creation occur outside ShouldProcess. WhatIf is not a side-effect-free preview.

.LINK
Export-AllModuleFromExcelFile

.LINK
Import-ModuleToExcelFile
#>
function Export-ModuleFromExcelFile {
    [CmdletBinding(SupportsShouldProcess)]
    Param (
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$ModuleName,

        [Parameter(Mandatory = $true, Position = 1)]
        [string]$InputFile,

        [Parameter(Mandatory = $false, Position = 2)]
        [string]$OutputDirectory,

        [switch]$Force
    )

    if (-not (Test-Path $InputFile)) {
        Write-Error "Input File '$($InputFile)' does not exist."
        return
    }
    
    # Internal Functions

    function _get_fileext($comp_type, $export_doc) {
        if ($comp_type -eq 1) { return ".bas" }
        if ($comp_type -eq 2) { return ".cls" }
        if ($comp_type -eq 3) { return ".frm" }
        if ($component.Type -eq 100) { return ".cls" }
        return $null
    }

    # Main
    # Create an instance of Excel Application
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $False
    #$excel.AutomationSecurity = [Microsoft.Office.Core.MsoAutomationSecurity]::msoAutomationSecurityLow # Adjust macro settings if necessary

    $inputFilePath = Resolve-Path $InputFile
    
    if (-not (Test-Path $OutputDirectory)) {
        New-Item $OutputDirectory -ItemType Directory -Force | Out-Null
    }

    $outputDirPath = Resolve-Path $OutputDirectory

    try {
        # Open the workbook
        if ($PSCmdlet.ShouldProcess($InputFile, "Open Excel file $($InputFile)")) {
            $workbook = $excel.Workbooks.Open($inputFilePath)
        }
        Write-Debug "File opened."
        try {
            # Loop through all VBComponents in the VBA project
            if ($PSCmdlet.ShouldProcess($OutputDirectory, "Export VBA modules in Excel file '$($InputFile)' to directory '$($OutputDirectory)'")) {
                [bool]$is_found = $false
                foreach ($component in $workbook.VBProject.VBComponents) {
                    # Export only standard modules and class modules
                    $fileext = _get_fileext($component.Type)
                    if (-not $null -eq $fileext -and $component.Name -eq $ModuleName) {
                        # Construct the file name
                        $filename = (Join-Path $outputDirPath $component.Name) + $fileext
                        Write-Debug $filename

                        # Check if the file already exists
                        if (Test-Path $filename) {
                            if ($Force) {
                                Remove-Item $filename -Force
                            }
                            else {
                                Write-Error "File name '$($filename)' already exists."
                                return
                            }
                        }

                        # Export the module
                        $component.Export($filename)
                        $is_found = $true
                    }
                    else {
                        Write-Error "Module '$($ModuleName)' is ignore type. (type:$($component.Type))"
                    }
                }

                if (-not $is_found) {
                    Write-Error "Module name '$($ModuleName)' was not found."
                }
            }
        }
        finally {
            if ($PSCmdlet.ShouldProcess($InputFile, "Close Excel file $($InputFile)")) {
                # Close the workbook
                $workbook.Close($False)
                $excel.Quit()
            }
        }
    }
    finally { 
        # Reset macro security settings if they were changed
        #$excel.AutomationSecurity = [Microsoft.Office.Core.MsoAutomationSecurity]::msoAutomationSecurityByUI

        # Release COM objects
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) > $null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}
