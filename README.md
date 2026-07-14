# xls-common-modules

`xls-common-modules` is a development repository for `CommonModules`, which are shared by multiple Excel VBA tools, and for supporting workbook import, extraction, distribution, formatting, and documentation generation.

The source of truth for common modules is `CommonModules/src/CommonModules`. Common modules under an individual tool's `common_modules_repo`, document source set, or workbook-local source set are treated as distributed copies.

## Features

This repository lets Excel VBA macros be treated as development assets that can be source-controlled, reused, and tested, rather than as logic confined inside a workbook.

One way to share common processing is to place it in an `.xlsb` or add-in. However, when deliverables depend on external workbooks or add-in references, they are no longer self-contained single files and become less portable. Another approach is to manually copy common processing into each `.xlsm`; that keeps each `.xlsm` easy to distribute, but makes updates across multiple workbooks easy to miss.

`xls-common-modules` centrally manages common modules as VBA source in `CommonModules/src/CommonModules`, collects manifest-listed files into `common_modules_repo`, and updates vba-dev projects through `UPDATE_COMMON_MODS`. Distribution keeps the portability of a single `.xlsm`, while development can keep common modules current across multiple individual tools.

Excel macros have very few widely established unit-test frameworks. `xls-common-modules` provides a lightweight test runner, assertions, and a test double foundation that run inside Excel workbooks. The normal automated path is `vba-dev test`, which builds the selected document and runs `UnitTestMain` inside the generated workbook.

- Externalizes VBA source as `.bas` / `.cls` files so Git can manage diffs.
- Centrally manages common modules and distributes the same implementation to multiple individual tools.
- Allows each individual tool to be distributed as a single `.xlsm` with the necessary VBA source included.
- Wraps Excel workbook, worksheet, file, and text I/O operations as services so individual modules are easier to test.
- Provides typed sets, keyed references, and virtual tables over input ranges through `ObjectList`, `ObjectSet`, `ObjectDictionary`, and `WorksheetVirtualTable`.
- Provides regression testing for VBA through `Lib_UnitTest`, `UnitTestAssert`, and each `*TestDouble`.
- Generates API reference documentation for VBA source through Doxygen and `DoxyVB6`.

## Core Capabilities

`CommonModules` centralizes the following kinds of processing.

| Category | Main VBA sources |
| --- | --- |
| Excel workbook/worksheet operations | `WorkbookService`, `WorksheetService`, `WorksheetRangeBounds`, `WorksheetVirtualTable`, `Lib_Common`, `Lib_CommonConstructor` |
| Collections/enumeration | `ObjectList`, `ObjectSet`, `ObjectDictionary`, `Counter`, `CounterSet`, `Enumerator`, `ArrayObject`, `IElementTypeProvider` |
| Input sheets | `Lib_InputSheet`, `IUserInputSheet`, `UserInputSheet`, `UserInputSheetTestDouble` |
| IPv4 | `Lib_IPv4` |
| Files/text | `FileSystemService`, `TextFileService`, `TextFileEntity` |
| Test support | `Lib_UnitTest`, `UnitTestAssert`, `TestDoubleBehaviorStore`, each `*TestDouble` |
| Run state/diagnostics | `ApplicationScreenUpdateManager`, `CommonRunStateManager`, `DebugInformation`, `ProgressStatus` |

Normal GUI entry points call `InitializeCommonService(Force:=True)` and use `WbSrv`, `WsSrv`, `New_RangeBounds`, and related APIs. Public UDF entry points used as Excel worksheet functions use `InitializeUdfCommonService` to avoid side effects during cell recalculation.

For range processing, `WorksheetRangeBounds.GetRows()` / `GetColumns()` and `WorksheetVirtualTable` allow worksheet rectangles to be handled as rows, columns, or headed records. `ObjectDictionary` is a typed dictionary that supports key references, and its standard `For Each` enumerates keys. `FileSystemService.GetAbsolutePath` expands OS environment variables such as `%LOCALAPPDATA%` in non-URL local paths before converting them to absolute paths.

```vb
Call InitializeCommonService(Force:=True)

Dim target_cell As WorksheetRangeBounds
Set target_cell = New_RangeBounds(Row:=1, Column:=1, Sheet:="INPUT")

Call WsSrv.WriteCell(target_cell, "value", TypeConvert:=False)
```

## Requirements

- Windows and Microsoft Excel are required.
- Workbook import, extraction, vba-dev build/test/publish, and `UnitTestMain` require Excel's "Trust access to the VBA project object model" setting.
- Do not leave the target `.xlsm` open manually while importing, extracting, building, testing, or publishing.
- API reference generation requires Doxygen. `tools/DoxyVB6/DoxyVB6.exe` is used as the VBA input filter.
- Tool wrappers resolve `vba-dev.exe` in this order: `VBA_DEV_EXE`, the known workspace development build, then `PATH`.

## CommonModules vba-dev Project

`CommonModules/` is a workbook-backed `vba-dev` project root. Its document source set is `CommonModules/src/CommonModules`, the template workbook is `CommonModules/src/CommonModules/CommonModules.xlsm`, and generated workbooks are written under `CommonModules/bin/CommonModules` and `CommonModules/publish/CommonModules`.

`CommonModules/project.json` keeps the `vba-dev new excel` default CommonModules repository relationship for this repository layout: `commonModulesRepository` is `../common_modules_repo`. All manifest-listed CommonModules entries are installed as requested modules, equivalent to adding every common module to the project.

```powershell
vba-dev doctor --project .\CommonModules
vba-dev build --project .\CommonModules
vba-dev test --project .\CommonModules
vba-dev publish --project .\CommonModules
```

## Development Workflow

For vba-dev projects, individual tool developers edit the document source set declared by `project.json` `sourcePath`, build the workbook, and test it. Workbook-local workflows remain available for explicit workbook import/extract through a `modules` folder next to the target workbook.

This section uses `xls-web-tools/SampleWebTool` as an example.

1. Edit the individual modules under the tool's document source set, such as `xls-web-tools/SampleWebTool/src/SampleWebTool`.
2. If you need to change a copy of a common module, first decide whether the work should be handled as a `CommonModules` change instead of as an individual module change.
3. Update `Test_*.bas` or test `.cls` files in the same source set as needed.
4. Run the formatting check for the source set with `tools/format_vba_source_main.ps1`.
5. Run `vba-dev test --project <project> --document <document>` for the target document. Use `vba-dev build` when the document has no tests and only build output is required.
6. For workbook-local workflows only, import `modules` into the workbook with `tools/IMP_MODS.BAT`.
7. If worksheet buttons or GUI entry points changed, also verify the real operation path.
8. If you had to edit in VBE, extract VBA source from `.xlsm` with `tools/EXP_MODS.BAT`; this writes to the workbook-local source set next to that workbook.
9. When taking in common-module distribution updates, run `tools/UPDATE_COMMON_MODS.BAT` on the parent folder that contains vba-dev projects, then rerun `vba-dev test` for the target individual tools.

## Unit Testing

Tests are written as `Test_...(ByVal Assert As UnitTestAssert)` subprocedures in `Test_*.bas`, then executed through `vba-dev test`. Internally the built workbook runs `UnitTestMain`, and results are output to `UNIT_TEST_SHEET`.

For example, a document source set may contain test modules such as:

- `Test_WebDriverSessionClient.bas`
- `Test_WebDriverClient.bas`
- `Test_ToolSettings.bas`
- `Test_OutputSheetWriter.bas`

The basic form is Arrange to prepare the subject and test doubles, Act to call the target processing, and Assert to verify return values, state, and call history.

```vb
Public Sub Test_Sample(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim actual_value As String

    ' --- Act ---
    actual_value = "expected"

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "expected", actual_value
End Sub
```

Actual tests substitute dependencies such as `WorksheetServiceTestDouble`, `WorkbookServiceTestDouble`, `FileSystemServiceTestDouble`, and `UserInputSheetTestDouble` for `WbSrv` / `WsSrv` / `FsSrv` or input sheet dependencies. This allows reads, writes, and issued calls to be verified without mutating real Excel sheets or files.

## Tooling

Use the `.bat` files or the `.lnk` files placed in individual tool folders when those entry points exist. The usual operation is to drag and drop the target folder, source set directory, or `.xlsm` onto the `.bat` / `.lnk`. When running from the command line, call the `.bat` files.

| Purpose | Recommended entry point | Target example |
| --- | --- | --- |
| Workbook import from workbook-local source set into `SampleWebTool.xlsm` | `tools/IMP_MODS.BAT` | `xls-web-tools/SampleWebTool/SampleWebTool.xlsm` |
| Extract VBA source from `SampleWebTool.xlsm` into workbook-local source set | `tools/EXP_MODS.BAT` | `xls-web-tools/SampleWebTool/SampleWebTool.xlsm` |
| Generate the API reference for a source set | `tools/GEN_DOC.BAT` | `xls-web-tools/SampleWebTool/src/SampleWebTool` |
| Update CommonModules and build all vba-dev documents under a folder | `tools/UPDATE_COMMON_MODS.BAT` | `xls-web-tools` |
| Collect manifest-listed common modules into `common_modules_repo` | `tools/COLLECT_COMMON_MODS.BAT` | `xls-common-modules` |
| Distribute to each individual tool repository's `common_modules_repo` | `tools/DIST_COMMON_MODS_REPO.BAT` | `xls-common-modules/common_modules_repo` |

`GEN_DOC.BAT` treats the dropped path as a source directory and reads only direct `.bas`, `.cls`, and `.frm` files. For `<owner>/src/<target>`, it writes HTML to `<owner>/docs/<target>/api-reference` and creates `<owner>/docs/<target>/api-reference.zip`; the Doxygen project name is `<target>`. For `<owner>/modules`, `<owner>/<name>`, or `<owner>/src`, it writes to `<owner>/docs/api-reference` and creates `<owner>/docs/api-reference.zip`; `<owner>/src` emits a warning because `src` is normally a container directory.

### Command-Line Examples

The following examples are run from the `xls-common-modules` root.

```powershell
# Workbook import from workbook-local source set into SampleWebTool.xlsm
.\tools\IMP_MODS.BAT ..\xls-web-tools\SampleWebTool\SampleWebTool.xlsm

# Extract VBA source from SampleWebTool.xlsm into the adjacent modules folder
.\tools\EXP_MODS.BAT ..\xls-web-tools\SampleWebTool\SampleWebTool.xlsm

# Generate the API reference for a source set
.\tools\GEN_DOC.BAT ..\xls-web-tools\SampleWebTool\src\SampleWebTool

# Update CommonModules and build all vba-dev documents under xls-web-tools
.\tools\UPDATE_COMMON_MODS.BAT ..\xls-web-tools

# Check VBA source formatting
powershell -ExecutionPolicy bypass -NoLogo -NonInteractive -File .\tools\format_vba_source_main.ps1 ..\xls-web-tools\SampleWebTool\src\SampleWebTool -Recurse -Check

# Apply VBA source formatting
powershell -ExecutionPolicy bypass -NoLogo -NonInteractive -File .\tools\format_vba_source_main.ps1 ..\xls-web-tools\SampleWebTool\src\SampleWebTool -Recurse
```

## Documentation

- CommonModules specification: `docs/product-spec.md`
- API reference: `docs/api-reference.zip` for owner-level source directories, or `docs/<target>/api-reference.zip` for `<owner>/src/<target>` source directories
- Glossary: `CONTEXT.md`
