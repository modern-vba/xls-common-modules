<#
.SYNOPSIS
Imports selected VBA source files into an existing Excel workbook and saves it.

.DESCRIPTION
Starts a hidden Excel instance and opens ExcelFile. For each source file, reads
Attribute VB_Name to determine its component name, falling back to the file name.
Recognizes .bas, .cls, and .frm as standard modules, class modules, and UserForms.
By default, removes existing components with matching names and types, then
imports every supplied source file. Unrelated components are retained.
The same workbook is saved in place, closed, and its Excel instance is quit.

.PARAMETER ExcelFile
The workbook path passed to Excel.Workbooks.Open. Required, at position 0.
Use an absolute path to an existing workbook. There is no separate output path.

.PARAMETER ModuleFiles
An array of paths to individual .bas, .cls, or .frm source files. Required, at
position 1. Use explicit file paths; the function does not enumerate directories
or expand this list into a recursive search. For forms, retain the companion
resource (.frx) file next to the .frm file when the exported form requires it.

.PARAMETER Flush
Removes all existing standard modules, class modules, and UserForms before
importing every supplied source file. Document modules are retained.
Cannot be combined with Update.

.PARAMETER Update
Reimports only source files whose component name and type already exist in the
workbook. New components are skipped. Cannot be combined with Flush.

.INPUTS
None.
This function does not accept pipeline input.

.OUTPUTS
System.Collections.Hashtable
Maps each imported source-file path to $true. Skipped files have no entry.
An error before completion can prevent the result from being returned.

.EXAMPLE
$result = Import-ModuleToExcelFile -ExcelFile 'C:\work\Book.xlsm' -ModuleFiles 'C:\exports\Common_Mod.bas', 'C:\exports\Worker.cls'
Replaces matching components, adds new ones, and saves Book.xlsm in place.

.EXAMPLE
Import-ModuleToExcelFile -ExcelFile 'C:\work\Book.xlsm' -ModuleFiles 'C:\exports\Worker.cls' -Update
Reimports Worker only if a class module with that name already exists.

.EXAMPLE
Import-ModuleToExcelFile -ExcelFile 'C:\work\Book.xlsm' -ModuleFiles 'C:\exports\Common_Mod.bas' -Flush
Removes existing standard modules, classes, and forms, then imports Common_Mod.

.NOTES
Legacy Excel COM helper. Normal workspace import/export workflows use vba-dev.
Requires installed Excel, an accessible VBA project, and Trust access to the VBA
project object model. Close the target workbook before calling this function.
Source names are read with Get-Content -Encoding oem. This command does not replace
the code inside existing worksheet or ThisWorkbook document modules.
Supports WhatIf and Confirm, but Excel creation and final workbook.Close($true)
occur outside ShouldProcess. The final close requests saving even after an error;
WhatIf must not be treated as a side-effect-free preview or rollback mechanism.

.LINK
Export-ModuleFromExcelFile

.LINK
Export-AllModuleFromExcelFile
#>
function Import-ModuleToExcelFile {
    [CmdletBinding(SupportsShouldProcess)]
    Param (
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$ExcelFile,

        [Parameter(Mandatory = $true, Position = 1)]
        [string[]]$ModuleFiles,

        [switch]$Flush,

        [switch]$Update
    )

    if ($Flush -and $Update) {
        Write-Error "The 'Update' and 'Flush' switches cannot be specified at the same time."
        return
    }

    # Local Functions
    function _get_modname($mod_file) {
        $mod_content = Get-Content -Path $mod_file -Encoding oem
        $vb_name_pattern = '^\s*Attribute\s+VB_Name\s+=\s+"(.+)"\s*$'
        foreach ($mod_line in $mod_content) {
            if ($mod_line -match $vb_name_pattern) {
                return $matches[1]
            }
        }
        return (Split-Path $mod_file -Leaf) -replace '^(\.?(?:[^.]+\.?$|[^.]+))(\.[^\.]+\.?$)?', '$1'
    }

    function _get_modtype($mod_file) {
        $mod_ext = (Split-Path $mod_file -Leaf) -replace '^(\.?(?:[^.]+\.?$|[^.]+))(\.[^\.]+\.?$)?', '$2'
        if ($mod_ext -eq ".bas") {
            return 1
        }
        elseif ($mod_ext -eq ".cls") {
            return 2
        }
        elseif ($mod_ext -eq ".frm") {
            return 3
        }
        else {
            return -1
        }
    }

    function _get_modinfo($check_mod_name, $check_mod_type, [System.Collections.ArrayList]$check_mod_list) {
        foreach ($check_mod_item in $check_mod_list) {
            if ($check_mod_item.Name -eq $check_mod_name -and $check_mod_item.Type -eq $check_mod_type) {
                return $check_mod_item
            }
        }
        return $null
    }

    function _remove_comps($workbook_obj, $vb_comps) {
        $remove_after = [System.Collections.ArrayList]::new()
        foreach ($vb_comp in $vb_comps) {
            [string]$vb_comp_name = $vb_comp.Name
            if ($vb_comp_name.StartsWith("I")) {
                $remove_after.Add($vb_comp) > $null
            }
            else {
                Write-Debug "Removing the old module '$($vb_comp_name)'."
                $workbook_obj.VBProject.VBComponents.Remove($vb_comp)
                Write-Information "Removed the old module '$($vb_comp_name)'."
            }
        }
        foreach ($vb_comp in $remove_after) {
            $vb_comp_name = $vb_comp.Name
            Write-Debug "Removing the old module '$($vb_comp_name)'."
            $workbook_obj.VBProject.VBComponents.Remove($vb_comp)
            Write-Information "Removed the old module '$($vb_comp_name)'."
        }
    }

    # Main

    # Create an instance of the Excel application
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $False
    try {
        # Open the workbook
        if ($PSCmdlet.ShouldProcess($ExcelFile, "Open Excel file $($ExcelFile)")) {
            Write-Debug "Opening '$($ExcelFile)'."
            $workbook = $excel.Workbooks.Open($ExcelFile)
            Write-Information "Opened '$($ExcelFile)'."
        }
        
        # Prepare the results array
        $result = @{}

        try {
            # Build list of target modules.
            $target_modules = [System.Collections.ArrayList]::new()
            foreach ($file in $ModuleFiles) {
                $mod_name = _get_modname $file
                $mod_type = _get_modtype $file

                $mod_item = [ExcelModuleInformation]::new()
                $mod_item.Name = $mod_name
                $mod_item.Type = $mod_type
                $mod_item.File = $file
                    
                $target_modules.Add($mod_item) > $null
            }
            
            $to_be_removed = [System.Collections.ArrayList]::new()
            if ($Flush) {
                # Remove all modules.
                foreach ($component in $workbook.VBProject.VBComponents) {
                    if ($component.Type -eq 1 -or $component.Type -eq 2 -or $component.Type -eq 3) {
                        Write-Debug "Marking module '$($component.Name)' for removal."
                        $to_be_removed.Add($component) > $null
                    }
                    else {
                        Write-Debug "Skipped module '$($component.Name)' because its type is $($component.Type)."
                    }
                }
            }
            else {
                # Remove old modules.
                $removed_list = [System.Collections.ArrayList]::new()
                foreach ($component in $workbook.VBProject.VBComponents) {
                    $mod_item = _get_modinfo $component.Name $component.Type $target_modules
                    if ($null -ne $mod_item) {
                        Write-Debug "Marking module '$($component.Name)' for removal."
                        $to_be_removed.Add($component) > $null
                        $removed_list.Add($mod_item) > $null
                    }
                    else {
                        if ($component.Type -eq 1 -or $component.Type -eq 2 -or $component.Type -eq 3) {
                            Write-Information "Skipped module '$($component.Name)' because it does not to be update."
                        }
                        else {
                            Write-Debug "Skipped module '$($component.Name)' because its type is $($component.Type)."
                        }
                    }
                }
            }

            # Execute removal.
            if ($PSCmdlet.ShouldProcess($ExcelFile, "Remove listed old modules in Excel file '$($ExcelFile)'")) {
                _remove_comps $workbook $to_be_removed
            }

            if ($Update) {
                $target_modules = $removed_list
            }

            # Import files.
            foreach ($mod_item in $target_modules) {
                if ($PSCmdlet.ShouldProcess($ExcelFile, "Import module file '$($file)' to Excel file '$($ExcelFile)'")) {
                    Write-Debug "Importing '$($mod_item.File)'."
                    $workbook.VBProject.VBComponents.Import($mod_item.File) > $null
                    $result[$mod_item.File] = $true
                    Write-Information "Imported '$($mod_item.Name)'."
                }
            }
            
            if ($PSCmdlet.ShouldProcess($ExcelFile, "Save and close Excel file $($ExcelFile)")) {
                Write-Debug "Saving '$($ExcelFile)'."
                $workbook.Save()
                Write-Information "Saved '$($ExcelFile)'."
            }
        }
        finally {
            # Save and close the workbook.
            Write-Debug "Closing '$($ExcelFile)'."
            $workbook.Close($True)
            Write-Information "Closed."
            $excel.Quit()
        }

        # Return results.
        return $result
    }
    finally {
        # Release COM objects.
        if ($null -ne $workbook) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) > $null }
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) > $null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

class ExcelModuleInformation {
    [string]$Name
    [int]$Type
    $File
}
