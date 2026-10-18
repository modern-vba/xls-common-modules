<#
.SYNOPSIS
Exports VBA components from an Excel workbook to a directory.

.DESCRIPTION
Starts a hidden Excel instance and opens InputFile. With ExportDocumentModule
omitted, exports standard modules as .bas, class modules as .cls, and UserForms
as .frm with any companion file produced by VBIDE. Document modules are skipped.
Closes the workbook without saving and quits the Excel instance after export.

.PARAMETER InputFile
The path to the existing Excel workbook. Required, at position 0.

.PARAMETER OutputDirectory
The export directory. Required, at position 1. Use an existing empty directory,
or specify Force to request removal of its existing entries. The implementation
can create a missing directory, but its earlier directory enumeration may emit
an error before creation.

.PARAMETER ExportDocumentModule
Requests worksheet and ThisWorkbook module export. Known implementation limitation:
when this switch is set, the current extension helper classifies every component
as .bas, including class modules, UserForms, and document modules. Do not use this
switch to produce correctly typed component files; use vba-dev export instead.

.PARAMETER Force
Removes existing entries matched by OutputDirectory\* before export. Without
this switch, a nonempty output directory produces an error. Removal uses
Remove-Item -Force without Recurse; nonempty subdirectories may prompt or fail.

.INPUTS
None.
This function does not accept pipeline input.

.OUTPUTS
None.
Writes exported component files and informational messages.

.EXAMPLE
Export-AllModuleFromExcelFile -InputFile 'C:\work\Book.xlsm' -OutputDirectory 'C:\exports'
Exports standard modules, classes, and UserForms to an existing empty directory.

.EXAMPLE
Export-AllModuleFromExcelFile -InputFile 'C:\work\Book.xlsm' -OutputDirectory 'C:\exports' -Force
Removes existing matching entries from the output directory before exporting.

.NOTES
Legacy Excel COM helper. Normal workspace import/export workflows use vba-dev.
Requires installed Excel, an accessible VBA project, and Trust access to the VBA
project object model. Close the target workbook before calling this function.
Supports WhatIf and Confirm, but Excel instance creation, directory creation,
and Force cleanup occur outside ShouldProcess. WhatIf can still remove files.

.LINK
Export-ModuleFromExcelFile

.LINK
Import-ModuleToExcelFile
#>
function Export-AllModuleFromExcelFile {
    [CmdletBinding(SupportsShouldProcess)]
    Param (
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$InputFile,

        [Parameter(Mandatory = $true, Position = 1)]
        [string]$OutputDirectory,

        [switch]$ExportDocumentModule,
        [switch]$Force
    )

    if (-not (Test-Path $InputFile)) {
        Write-Error "Input File '$($InputFile)' does not exist."
        return
    }

    if ((Get-ChildItem -Path $OutputDirectory).Count -ne 0 -and -not $Force) {
        Write-Error "Output directory '$($OutputDirectory)' is not empty. Use the -Force switch to remove existing files."
        return
    }

    # Internal Functions

    function _get_fileext($comp_type, $export_doc) {
        if ($comp_type -eq 1) { return ".bas" }
        if ($comp_type -eq 2) { return ".cls" }
        if ($comp_type -eq 3) { return ".frm" }
        if ($export_doc -and $component.Type -eq 100) { return ".cls" }
        return $null
    }

    # Main

    # Create an instance of Excel Application
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $False
    #$excel.AutomationSecurity = [Microsoft.Office.Core.MsoAutomationSecurity]::msoAutomationSecurityLow # Change macro settings if needed

    $inputFilePath = Resolve-Path $InputFile
    
    if (-not (Test-Path $OutputDirectory)) {
        New-Item $OutputDirectory -ItemType Directory -Force | Out-Null
    }
    elseif ($Force) {
        Remove-Item -Path (Join-Path $OutputDirectory "*") -Force
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
                foreach ($component in $workbook.VBProject.VBComponents) {
                    # Export only standard modules and class modules
                    $fileext = _get_fileext($component.Type, $ExportDocumentModule)
                    if (-not $null -eq $fileext) {
                        $filename = (Join-Path $outputDirPath $component.Name) + $fileext
                        Write-Information "Exported $($filename)"
                        $component.Export($filename)
                    }
                    else {
                        Write-Debug "Ignoring type '$($component.Name)($($component.Type))'"
                    }
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
        # Reset macro security if changed
        #$excel.AutomationSecurity = [Microsoft.Office.Core.MsoAutomationSecurity]::msoAutomationSecurityByUI

        # Release COM objects
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) > $null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}
