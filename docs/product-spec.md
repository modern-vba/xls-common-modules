# CommonModules Library Specification

## Positioning

This document is the library specification for the common modules under `xls-common-modules\CommonModules\src\CommonModules`.

Detailed API reference:

- Entry point: `xls-common-modules\CommonModules\doc\html\index.html`
- Member lists by class and standard module: `xls-common-modules\CommonModules\doc\html\annotated.html`, `namespace_class.html`, `namespace_standard.html`

## Overview

CommonModules is a foundation library shared by the Excel VBA tool set. A typical tool implementation calls `InitializeCommonService` in its entry point and uses APIs such as `WbSrv`, `WsSrv`, `New_RangeBounds`, `New_ObjectList`, `New_ObjectSet`, `New_ObjectDictionary`, and `WorksheetVirtualTable`. Public UDFs are placed in `Fx_*.bas`; non-UDF public APIs are placed in standard modules that have `Option Private Module`.

| Category | Main modules | Main responsibilities | Typical usage |
| --- | --- | --- | --- |
| General | `Lib_Common`, `Lib_CommonConstructor`, `Fx_Common`, `WorkbookService`, `WorksheetService`, `WorksheetRangeBounds`, `WorksheetRangeBoundsEnumerator`, `WorksheetVirtualTable`, `WorksheetVirtualTableEnumerator`, `ApplicationScreenUpdateManager`, `CommonRunStateManager`, `DebugInformation`, `ProgressStatus` | Excel workbook, worksheet, and cell-range operations; range value objects; virtual tables; common factories; common service initialization; string, array, and Excel address helpers; public UDFs; progress and diagnostics | Ordinary business macros, worksheet read/write, data conversion, range enumeration, worksheet formulas, GUI run-state management |
| Collections and enumeration | `ObjectList`, `ObjectSet`, `ObjectDictionary`, `Counter`, `CounterSet`, `Enumerator`, `ArrayObject`, `IEnumerator`, `IElementTypeProvider`, `IComparable`, `IEquatable`, `IDuplicateCheckable`, `IStringable` | Typed lists, sets, and dictionaries; counters; enumerators; internal arrays; element type self-declaration; comparison, duplicate checking, sorting, and stringification helpers | Holding data collections, duplicate management, keyed references, enumeration, sorting, key generation |
| Input screens | `Lib_InputSheet`, `IUserInputSheet`, `UserInputSheet`, `UserInputSheetTestDouble` | Input sheet creation, header search, input value range resolution, test substitution | Processing that retrieves settings values or target ranges from headed input sheets |
| IPv4 | `Lib_IPv4` | Conversion and validation of IPv4 addresses, networks, mask lengths, and mask values | Address normalization and network splitting/aggregation in firewall and network-related tools |
| File operations | `Lib_FileSystem`, `FileSystemService`, `Lib_TextFile`, `TextFileService`, `TextFileEntity` | File system access, path resolution, file and directory operations, text file I/O | Settings file generation, backup, file listing, text output |
| Test support | `Lib_UnitTest`, `UnitTestAssert`, each `*TestDouble`, `TestDoubleBehaviorStore`, `TestDoubleVariantKeyBuilder` | Unit tests for Excel VBA, assertions, service substitution, stub return values, call history, error injection | Unit tests for CommonModules and each tool |

## Basic Usage

### Common Service Initialization

Entry-point processing calls `InitializeCommonService`. Specify `Force:=True` when existing services should be regenerated as production services.

```vb
Call InitializeCommonService

Dim target_cell As WorksheetRangeBounds
Set target_cell = New_RangeBounds(Row:=1, Column:=1, Sheet:="INPUT")

Call WsSrv.WriteCell(target_cell, "123")
```

`InitializeCommonService` behaves as follows.

| Target | Behavior |
| --- | --- |
| `Lib_Common.WbSrv` | With `Force:=True`, regenerates `WorkbookService`. With `Force:=False`, generates `WorkbookService` only when the value is `Nothing`. |
| `Lib_Common.WsSrv` | With `Force:=True`, regenerates `WorksheetService`. With `Force:=False`, generates `WorksheetService` only when the value is `Nothing`. |
| `Lib_FileSystem.FsSrv` | Calls `InitializeFileSystemService` with `Force` only when that procedure exists in the same workbook. With `Force:=True`, regenerates `FileSystemService`; with `Force:=False`, generates `FileSystemService` only when the value is `Nothing`. |
| `Lib_TextFile.TfSrv` | Calls `InitializeTextFileService` with `Force` only when that procedure exists in the same workbook. With `Force:=True`, regenerates `TextFileService`; with `Force:=False`, generates `TextFileService` only when the value is `Nothing`. |

`Force` defaults to `False`. Services that have already been substituted with test doubles or other instances are not overwritten when `Force:=False`. With `Force:=True`, all target services are regenerated as production services. If `InitializeFileSystemService` or `InitializeTextFileService` does not exist in the same workbook, it is ignored as before even with `Force:=True`.

```vb
Set Lib_Common.WsSrv = New WorksheetServiceTestDouble
Call InitializeCommonService  ' WsSrv remains the substituted instance

Call InitializeCommonService(Force:=True)  ' Regenerates all target services as production services
```

UDF entry points executed as Excel worksheet functions do not call `InitializeCommonService` directly; they call `InitializeUdfCommonService`. `InitializeCommonService` uses `Application.Run` to initialize optional services, so during cell recalculation it can raise `Err.Number = 1004`, `HRESULT = 0x800A03EC`, and `'Run' method of '_Application' object failed`.

```vb
' UDF entry points normally omit Force
Call InitializeUdfCommonService
```

`InitializeUdfCommonService` targets only `WbSrv` and `WsSrv`; it does not initialize `FsSrv` or `TfSrv`. Use `InitializeWorkbookService` or `InitializeWorksheetService` when individual initialization is needed. These APIs also preserve substituted services with `Force:=False` and regenerate production services with `Force:=True`. Public UDFs such as `DIFFSTR` are placed in `Fx_Common.bas`, and `Fx_*.bas` does not have `Option Private Module`. If public UDFs grow enough to separate responsibilities, split them into `Fx_...` modules. Standard modules other than `Fx_*.bas` and `Test_*.bas` have `Option Private Module` to suppress non-UDF `Public Function` procedures from Excel formula suggestions. Non-UDF return-value APIs are not changed to `Sub` plus `ByRef` solely for this purpose.

### Basic Range Operations

Excel cells and ranges are represented by `WorksheetRangeBounds`.

```vb
Call InitializeCommonService

Dim input_cell As WorksheetRangeBounds
Set input_cell = New_RangeBounds(Row:=2, Column:=3, Sheet:="INPUT")

Dim value_text As String
value_text = WsSrv.ReadCell(input_cell)

Call WsSrv.WriteCell(input_cell.Shift(Column:=1), value_text, NumberFormat:="@")
```

### Used Range Enumeration

```vb
Call InitializeCommonService

Dim used_range As WorksheetRangeBounds
Set used_range = WsSrv.GetUsedRangeBounds( _
        New_RangeBounds(Row:=1, Column:=1, FinishRow:=G_ROW_MAX, FinishColumn:=G_COL_MAX, Sheet:="DATA"))

used_range.EnumerationMode = G_RANGE_ENUM_MODE_ROWS

Dim row_bounds As WorksheetRangeBounds
For Each row_bounds In used_range
    Debug.Print row_bounds.ToString(CellOnly:=True)
Next row_bounds
```

`ObjectList`, `ObjectSet`, `WorksheetRangeBounds`, and `WorksheetVirtualTable` support standard `For Each`. Standard `For Each` on `ObjectDictionary` enumerates keys; use `Items` to enumerate values. `For Each` enumerates a snapshot from the start of enumeration. If the original collection or range enumeration settings are changed during a standard `For Each`, continuation of the running loop is not guaranteed. For enumeration with update or deletion, `ObjectList`, `ObjectSet`, `WorksheetRangeBounds`, and `WorksheetVirtualTable` use `GetEnumerator`; `ObjectDictionary` uses copies from `Keys` or `Items` plus explicit APIs.

`WorksheetRangeBounds.GetRows()` and `GetColumns()` return row-wise or column-wise `WorksheetRangeBounds` values as `ObjectList("WorksheetRangeBounds")`. Use `ExpandRangeBoundsToMax` when multiple range shapes need to be aligned.

### Virtual Tables

```vb
Dim table_bounds As WorksheetRangeBounds
Set table_bounds = New_RangeBounds(Row:=1, Column:=1, FinishRow:=10, FinishColumn:=3, Sheet:="INPUT")

Dim vtable As WorksheetVirtualTable
Set vtable = New_WorksheetVirtualTableFromRangeBounds( _
        TableRange:=table_bounds, _
        TreatFirstRowAsHeader:=True)

Dim row_dict As ObjectDictionary
For Each row_dict In vtable
    Debug.Print row_dict.Item("Name").ToString(CellOnly:=True)
Next row_dict
```

`WorksheetVirtualTable` associates multiple `WorksheetRangeBounds` values by header and relative row. Each row is returned as an `ObjectDictionary` whose keys are headers and whose values are the corresponding one-row `WorksheetRangeBounds` values. Missing rows in shorter fields are returned as empty ranges with `RowCount = 0` while preserving the original field's column width.

### Duplicate-Free Sets

```vb
Dim names As ObjectSet
Set names = New ObjectSet

Call names.Add("alpha")
Call names.Add("beta")
Call names.Add("alpha", ErrorIfExists:=False)

If names.Exists("alpha") Then
    Debug.Print JoinStringSet(names, ",")
End If
```

### Keyed Typed Dictionaries

```vb
Dim row_values As ObjectDictionary
Set row_values = New_ObjectDictionary("WorksheetRangeBounds")
row_values.CompareMode = vbTextCompare

Call row_values.Add("Name", New_RangeBounds(Row:=2, Column:=1, Sheet:="INPUT"))
Call row_values.AddOrUpdate("Value", New_RangeBounds(Row:=2, Column:=2, Sheet:="INPUT"))

If row_values.Exists("name") Then
    Debug.Print row_values.Item("name").ToString(CellOnly:=True)
End If
```

The getter for `ObjectDictionary.Item(Key)` raises an error instead of implicitly adding a missing key. Use `Add` to add, `Update` to update, and `AddOrUpdate` or assignment to `Item(Key)` to add or update.

### Text File Output

```vb
Call InitializeCommonService
Call InitializeTextFileService

Dim entity As ITextFileEntity
Set entity = TfSrv.GetTextFileEntity("output.txt")

Call entity.OpenFile(AsWrite:=True, Force:=True)
Call entity.WriteLine("first line")
Call entity.CloseFile
```

### Service Substitution in Tests

```vb
Dim ws_double As WorksheetServiceTestDouble
Set ws_double = New WorksheetServiceTestDouble

Set Lib_Common.WsSrv = ws_double
Call InitializeCommonService

Dim target As WorksheetRangeBounds
Set target = New_RangeBounds(Row:=1, Column:=1, Sheet:="TEST")

Dim return_value As Variant
return_value = "stubbed"
Call ws_double.Store.SetReturn("ReadCell", return_value, target, "", "", False)

Debug.Print WsSrv.ReadCell(target)
```

## Specifications by Category

### General

- Common services
  - Main modules: `Lib_Common`, `Lib_CommonConstructor`
  - Specification summary: Holds `WbSrv` and `WsSrv` as global services, lazily initializing them with `InitializeCommonService` for normal entry points and `InitializeUdfCommonService` for UDF entry points. Also provides individual initialization through `InitializeWorkbookService` and `InitializeWorksheetService`. With `Force:=True`, target services are regenerated as production services. Provides `New_RangeBounds`, `New_RangeBoundsFromAddress`, `New_ObjectList`, `New_ObjectSet`, `New_ObjectDictionary`, `New_WorksheetVirtualTable`, and `New_WorksheetVirtualTableFromRangeBounds`.
  - Important limitations and side effects: Because global state is used, tests need explicit substitution and restoration. Substituted services are not overwritten by initialization with `Force:=False`, but are replaced by production services with `Force:=True`. VBIDE is not used to detect optional services.
- Excel workbooks
  - Main modules: `WorkbookService`, `IWorkbookService`
  - Specification summary: Handles workbook existence checks, list retrieval, open, save, close, worksheet add/delete/copy, worksheet search, and VBA component removal.
  - Important limitations and side effects: `SaveWorkbook` temporarily changes `Application.Calculation`, `DisplayAlerts`, and window visibility. `RemoveVBComponents` requires VBIDE access permission.
- Excel worksheets
  - Main modules: `WorksheetService`, `IWorksheetService`
  - Specification summary: Handles reading and writing cells/ranges, search, sort, copy, UsedRange calculation, duplicate removal, row insertion/deletion, and display/format/color/alignment changes. `WriteCell` supports type conversion, formulas, empty-value handling, and number format specification.
  - Important limitations and side effects: Mutates the target cells or ranges. String writes in `WriteCell` preserve the existing `WrapText`, and on write failure the implementation attempts to restore cell contents, number format, and `WrapText`. `ActivateRange` changes the active workbook, sheet, and cell. Copy operations may affect `CutCopyMode` or the clipboard.
- Range value objects
  - Main modules: `WorksheetRangeBounds`, `WorksheetRangeBoundsEnumerator`
  - Specification summary: Holds workbook name, worksheet name, and start/end rows and columns, representing cells, rectangles, full rows, full columns, whole sheets, and empty ranges. Supports range extraction, row/column/cell retrieval, list conversion through `GetRows` and `GetColumns`, transformation, shifting, intersection, stringification, and enumeration. Standard enumeration mode is switched through `EnumerationMode` and `EnumerationDescending`.
  - Important limitations and side effects: Initialization is one-time only. `Sheet` cannot be an empty string. Multi-area ranges are errors in `InitializeFromAddress`. Row and column limits depend on Excel limits.
- Virtual tables
  - Main modules: `WorksheetVirtualTable`, `WorksheetVirtualTableEnumerator`
  - Specification summary: Builds a virtual table associated by header and relative row from `ObjectList("WorksheetRangeBounds")` or a single `WorksheetRangeBounds`. Each row can be retrieved and enumerated as an `ObjectDictionary` whose keys are headers and whose values are the corresponding `WorksheetRangeBounds`.
  - Important limitations and side effects: `ColumnRangeList` must have the `WorksheetRangeBounds` element type contract. Header count mismatches and duplicate headers are initialization errors. Missing rows in shorter fields are returned as empty ranges with `RowCount = 0`. Business validation such as required headers, blank-row exclusion, required key columns, and duplicate-key detection is performed by the caller.
- General-purpose functions
  - Main modules: `Lib_Common`, `Fx_Common`
  - Specification summary: `Lib_Common` provides path joining/splitting, array conversion, string processing, key generation, range-shape expansion, Excel address generation/parsing, bit operations, and Excel error value conversion. `Fx_Common` provides public UDFs intended for use from worksheet formulas.
  - Important limitations and side effects: Some functions depend on Excel objects or `WsSrv`. Public UDFs call `InitializeUdfCommonService` without `Force` at the entry point even when the current implementation does not directly use services. `Fx_Common.DIFFSTR` is in this category and does not initialize `FsSrv` or `TfSrv`. The full specification of `DIFFSTR` `ExtractType` is not yet confirmed. Standard modules other than `Fx_*.bas` and `Test_*.bas` have `Option Private Module` and are treated as public APIs for use inside the same VBA project.
- UI and state helpers
  - Main modules: `ApplicationScreenUpdateManager`, `CommonRunStateManager`, `DebugInformation`, `ProgressStatus`
  - Specification summary: Provides save/restore of Excel application settings, creation/destruction of `DbgInfo` and `ProgStat` associated with one GUI run, a debug task stack, and StatusBar progress display.
  - Important limitations and side effects: `ApplicationScreenUpdateManager.Restore` errors if nothing has been saved. `CommonRunStateManager` creates new `DbgInfo` and `ProgStat` instances on construction or `Initialize`, and returns them to `Nothing` on destruction or `Clear`. `ProgressStatus` changes `Application.StatusBar`.

### Collections and Enumeration

- Typed element collections
  - Main modules: `ObjectList`, `ObjectSet`, `ObjectDictionary`, `Enumerator`, `ArrayObject`, `IElementTypeProvider`
  - Specification summary: Provides type-fixed lists, sets, dictionaries, enumerators, and internal arrays. Supports standard `For Each` and explicit `IEnumerator` enumeration, element type self-declaration through `IElementTypeProvider`, and comparison, duplicate checking, sorting, and stringification through `IEquatable`, `IDuplicateCheckable`, `IComparable`, and `IStringable`.
  - Important limitations and side effects: `Empty` and `Null` cannot be elements. The element type contract is fixed by the first non-`Nothing` element or by the specification in `Initialize` / `New_Object...`; later values of a different type error. `Nothing` is treated as a null reference for object type contracts. Performance limits with large element counts are not yet confirmed.
- Counters
  - Main modules: `Counter`, `CounterSet`
  - Specification summary: Provides a single counter with count, step, and maximum value, and a named counter collection.
  - Important limitations and side effects: Whether progress is allowed depends on `StopWhenMax` and `MaxCount`. `CounterSet` holds `Counter` instances internally.

### Input Screens

- Input sheets
  - Main modules: `Lib_InputSheet`, `IUserInputSheet`, `UserInputSheet`, `UserInputSheetTestDouble`
  - Specification summary: `Lib_InputSheet` provides `New_InputSheet`. `UserInputSheet` searches headers on an input sheet and resolves corresponding cells and ranges. `IUserInputSheet` and `UserInputSheetTestDouble` allow input sheet reads to be substituted in tests.
  - Important limitations and side effects: `UserInputSheet` checks that the target workbook and worksheet exist during initialization. Read targets depend on header placement and `WbSrv` / `WsSrv`.

### IPv4

- Address conversion
  - Main APIs: `ConvertFromIpAddress`, `ConvertToIpAddress`
  - Specification summary: Converts between dotted-decimal IPv4 strings and 32-bit values.
  - Important limitation: Because VBA `Long` is signed, values at or above `128.0.0.0` become negative.
- Mask conversion
  - Main APIs: `ConvertFromMaskLength`, `ConvertToMaskLength`, `InvertMaskValue`, `GetMaskValue`
  - Specification summary: Converts mask lengths, mask values, and bit-range masks.
  - Important limitation: Mask length is 0-32. Non-contiguous mask values are errors.
- Network judgment
  - Main APIs: `IsNetwork`, `IsValidMaskValue`, `GetHostAddress`, `GetNetworkAddress`
  - Specification summary: Determines and extracts network and host portions from IP values and mask values.
  - Important limitation: Exhaustive behavior for invalid mask values is not yet confirmed for all functions.
- Network transformation
  - Main APIs: `ExpandNetwork`, `NarrowNetwork`
  - Specification summary: Returns a parent network with mask length shortened by 1, or two subnets with mask length lengthened by 1.
  - Important limitation: `ExpandNetwork` errors on `/0`. The implementation treats `/30` through `/32` as "cannot narrow further" and raises an error.
- Parsing and formatting
  - Main APIs: `WellFormedAddress`, `ParseIpAddressAndMask`, `G_IPV4_*_RE`
  - Specification summary: Replaces `_` with `/` and appends `/32` to standalone IPv4 values. `IP/mask` accepts both CIDR length and dotted-decimal masks.
  - Important limitation: `ParseIpAddressAndMask` is presumed not to accept omitted-mask forms. The external compatibility guarantee for regular expression constants is not yet confirmed.

### File System Access

- File service initialization
  - Main module: `Lib_FileSystem`
  - Specification summary: Holds `FsSrv` and generates `FileSystemService` only when unset through `InitializeFileSystemService`. With `Force:=True`, regenerates `FileSystemService` regardless of the existing value.
  - Important limitations and side effects: Because global state is used, tests must account for substitution.
- Paths and file system
  - Main modules: `FileSystemService`, `IFileSystemService`
  - Specification summary: Handles relative-path absolutization, OS environment variable expansion for local paths, OS user temporary folder retrieval, temporary directory creation, file/directory listing, existence checks, last modified timestamps, creation, move, copy, delete, and backup creation.
  - Important limitations and side effects: Local paths that are not URLs are expanded with `WScript.Shell.ExpandEnvironmentStrings` before absolute-path judgment in `GetAbsolutePath`. URL strings do not receive environment variable expansion. Relative paths are based on `WbSrv.GetThisWorkbookDirectoryPath`. `GetTemporaryDirectoryPath` returns an absolute Windows path without a trailing `\`. `CreateTemporaryDirectory` creates a real directory directly under the OS user temporary folder. It mutates the real file system. Exhaustive behavior for network paths, long paths, ACLs, and read-only attributes is not yet confirmed.
- Text file initialization
  - Main module: `Lib_TextFile`
  - Specification summary: Holds `TfSrv` and generates `TextFileService` only when unset through `InitializeTextFileService`. With `Force:=True`, regenerates `TextFileService` regardless of the existing value.
  - Important limitations and side effects: Because global state is used, tests must account for substitution.
- Text files
  - Main modules: `TextFileService`, `TextFileEntity`, `ITextFileService`, `ITextFileEntity`
  - Specification summary: Creates `TextFileEntity` for a file path and handles read, overwrite, append, locking, and line-based I/O.
  - Important limitations and side effects: There is no character encoding selection API. Behavior depends on VBA defaults for `Open`, `Line Input`, and `Print`. `WriteLine` appends a line break.

### Tests

- Test runner
  - Main module: `Lib_UnitTest`
  - Specification summary: `UnitTestMain` detects `Test_...` procedures in `Test_*.bas` and outputs results to `UNIT_TEST_SHEET`.
  - Important limitations and side effects: Requires VBIDE access. Creates and deletes temporary VBA modules for execution. Creates or updates the result sheet.
- Assertions
  - Main module: `UnitTestAssert`
  - Specification summary: Provides assertions for scalars, numbers, booleans, types, `Nothing`, `Empty`, raised errors, and arrays.
  - Important limitations and side effects: Assertion failures are held in `IsFailed` and `ResultMessage` rather than treated as ordinary errors. After the first failure, additional assertions are not processed.
- Test double foundation
  - Main modules: `TestDoubleBehaviorStore`, `TestDoubleCallRecord`, `TestDoubleVariantKeyBuilder`
  - Specification summary: Manages return values, ByRef outputs, errors, and call history by method name and argument key.
  - Important limitations and side effects: Argument keys depend on `GetTypedValueKey`-style APIs. There are cases where `DefaultObjectKeyMode` and method-specific keying modes cannot be changed after key generation.
- Service test doubles
  - Main modules: `WorkbookServiceTestDouble`, `WorksheetServiceTestDouble`, `FileSystemServiceTestDouble`, `TextFileServiceTestDouble`, `TextFileEntityTestDouble`, `UserInputSheetTestDouble`
  - Specification summary: Implement each service or input sheet interface, return stub values or errors from `Store`, and record calls.
  - Important limitations and side effects: Default behavior varies by method. Classification follows "Default Behavior Classification for Service Test Doubles" below.

#### Default Behavior Classification for Service Test Doubles

| Category | Contract | Call history |
| --- | --- | --- |
| `strict stub` | Return values, ByRef outputs, or empty results must be explicitly registered in `Store`. If not registered, raise an error as a test setup defect. | Record only calls that complete normally. Do not record unregistered errors or explicit errors from `Store.SetError`. |
| `safe default` | Return a conservative default when unregistered. Existence checks generally return `False`, and ThisWorkbook-related methods return test defaults. | Record only calls that complete normally. |
| `spy only` | Do not execute operations without return values; only record calls. Use `Store.SetError` when testing failure cases. | If there is a payload to verify, record it as the value. Record `True` only when there is no particular payload. |
| `lightweight fake` | Do not touch external resources; implement only the lightweight state transitions needed by tests. | Record only calls that complete normally. Do not record state-violation errors. |

Classification for `WorkbookServiceTestDouble`:

| Category | Methods |
| --- | --- |
| `safe default` | `GetThisWorkbookName`, `GetThisWorkbookDirectoryPath`, `ExistsWorkbook`, `GetAllWorkbook`, `GetOtherWorkbook`, `ExistsWorksheet`, `IsSaved`, `HasPath` |
| `strict stub` | `GetAllWorksheet`, `GetOtherWorksheet`, `OpenWorkbook`, `SaveWorkbook`, `Find`, `AddWorksheet`, `CopyWorksheet` |
| `spy only` | `CloseWorkbook`, `RemoveWorksheet`, `ActivateWorksheet`, `RemoveVBComponents` |

Classification for `WorksheetServiceTestDouble`:

| Category | Methods |
| --- | --- |
| `strict stub` | `Find`, `ReadCell`, `ReadRange`, `EvaluateFormula`, `XLookup`, `IsEmptyCell`, `HasFormula`, `GetUsedRangeBounds` |
| `spy only` | `Sort`, `ActivateRange`, `WriteCell`, `WriteRange`, `WriteArrayFormula`, `CopyCell`, `CopyRange`, `ClearRange`, `RemoveDuplicates`, `InsertRows`, `DeleteRows`, `SetAllDataVisible`, `SetSheetOutlineLevel`, `SetSheetTabColor`, `SetRangeColor`, `SetWrapText`, `SetShrinkToFit`, `SetAlignment` |

Classification for `FileSystemServiceTestDouble`:

| Category | Methods |
| --- | --- |
| `safe default` | `PathExists`, `IsFile`, `IsDirectory` |
| `strict stub` | `GetAbsolutePath`, `GetTemporaryDirectoryPath`, `CreateTemporaryDirectory`, `GetFileList`, `GetDirectoryList`, `GetLastModified`, `CreateDirectory`, `RemoveDirectory`, `RemoveFile`, `GetNewestFile`, `CreateBackupFile` |
| `spy only` | `MoveDirectory`, `CopyDirectory`, `MoveFile`, `CopyFile` |

Classification for `TextFileServiceTestDouble` / `TextFileEntityTestDouble`:

| Class | Category | Methods |
| --- | --- | --- |
| `TextFileServiceTestDouble` | `lightweight fake` | `GetTextFileEntity` |
| `TextFileEntityTestDouble` | `lightweight fake` | `OpenFile`, `CloseFile` |
| `TextFileEntityTestDouble` | `strict stub` | `ReadLine` |
| `TextFileEntityTestDouble` | `spy only` | `Initialize`, `WriteLine` |

`FilePath`, `AsRead`, `AsWrite`, `AsAppend`, `GetReadLock`, `GetWriteLock`, `IsOpen`, and `IsEndOfFile` on `TextFileEntityTestDouble` are read-only properties that query fake state, and are outside the method classification table above.

## Important Limitations

### Side Effects on Excel State

The following APIs mutate Excel application or workbook state.

- `ApplicationScreenUpdateManager`
  - Side effects: Changes and restores `ScreenUpdating`, `EnableEvents`, `DisplayAlerts`, and `Calculation`.
- `WorkbookService.OpenWorkbook` / `SaveWorkbook` / `CloseWorkbook`
  - Side effects: Affects workbooks, window visibility, active workbook, and destination files.
- Write/copy/formatting APIs in `WorksheetService`
  - Side effects: Affect cell contents, formulas, number formats, colors, rows, duplicate rows, filters, outlines, and sheet tab colors.
- `WorksheetService.ActivateRange`
  - Side effects: Changes the active workbook, active worksheet, and selected cell.
- `ProgressStatus`
  - Side effects: Changes `Application.StatusBar`.
- `SetClipboard` / `GetClipboard` / `PasteFormulas`
  - Side effects: Affect the clipboard or paste destination.
- `Lib_UnitTest.UnitTestMain`
  - Side effects: Affects `UNIT_TEST_SHEET`, temporary VBA modules, and screen update state.

Some implementations attempt to restore state, but complete restoration is not yet confirmed for Excel-originated errors, protected state, and external factors.

### `WorksheetRangeBounds` Constraints

- Initialization
  - Constraint: Possible only once. Reinitialization is an error.
- Worksheet name
  - Constraint: Cannot be an empty string.
- Workbook name
  - Constraint: If empty, uses the default value from `WbSrv.GetThisWorkbookName`.
- Rows and columns
  - Constraint: Start row/column must be at least 1 and within Excel limits. End row/column cannot be negative.
- Empty ranges
  - Constraint: If the start exceeds the end, the range is treated as an empty range rather than an error.
- Multi-area ranges
  - Constraint: `InitializeFromAddress` treats multi-area ranges as errors.
- Comparison
  - Constraint: Workbook and worksheet name case differences are treated as the same. `EnumerationMode` and `EnumerationDescending` do not affect identity.
- Standard enumeration
  - Constraint: `EnumerationMode` is one of `G_RANGE_ENUM_MODE_ROWS`, `G_RANGE_ENUM_MODE_COLUMNS`, `G_RANGE_ENUM_MODE_CELLS_HORIZONTAL`, or `G_RANGE_ENUM_MODE_CELLS_VERTICAL`. `For Each` enumerates a snapshot from the start of enumeration, but if enumeration settings are changed during enumeration, continuation of the running loop is not guaranteed.

### `ObjectList` / `ObjectSet` / `ObjectDictionary` Constraints

- Types
  - Constraint: The element type contract is fixed by the first non-`Nothing` element or by explicit initialization. Later values of a different type error.
- Element type self-declaration
  - Constraint: If an object implements `IElementTypeProvider`, `ElementTypeKey` is used as the element type contract name. `ElementTypeKey` must have a form valid for class module names.
- Required capabilities
  - Constraint: `RequireComparable=True` requires an `IComparable` implementation. Identity/duplicate checking is determined by one of `G_OBJECT_KEY_MODE_REFERENCE`, `G_OBJECT_KEY_MODE_I_EQUATABLE`, or `G_OBJECT_KEY_MODE_DUPLICATE_CHECKABLE`.
- `Empty` / `Null`
  - Constraint: Cannot be used as elements.
- Excel error values
  - Constraint: `CVErr` can be treated as an element.
- Object comparison
  - Constraint: In legacy inference when `ObjectKeyMode` is not specified, element capability is inferred in the order `IDuplicateCheckable`, `IEquatable`, then reference equality. Duplicate checking in `ObjectSet` can prioritize `IDuplicateCheckable`.
- Array elements
  - Constraint: Compared by typed value keys.
- `ObjectDictionary.Item(Key)`
  - Constraint: The getter raises an error instead of implicitly adding a missing key. The setter is treated as add or update.
- `ObjectDictionary.CompareMode`
  - Constraint: String key comparison mode can be changed only before the first add. The value element type contract and `CompareMode` are preserved after `RemoveAll`.
- Standard enumeration
  - Constraint: `For Each` on `ObjectList` and `ObjectSet` enumerates elements. `For Each` on `ObjectDictionary` enumerates keys. It enumerates a snapshot from the start of enumeration, but if the original collection is updated or deleted during enumeration, continuation of the running loop is not guaranteed. Because `ObjectDictionary` does not expose `GetEnumerator`, use `Keys` for keys and `Items` or `ConvertToArray` for values, and use explicit APIs for update and deletion.

### File/Text API Constraints

- Environment variables
  - Constraint: `FileSystemService.GetAbsolutePath` applies OS environment variable expansion before absolute-path judgment, only for local paths that are not URLs. Handling of undefined variables and strings containing `%` follows the OS `ExpandEnvironmentStrings` result.
- Relative paths
  - Constraint: `FileSystemService` absolutizes them relative to `WbSrv.GetThisWorkbookDirectoryPath`.
- Drive-relative paths
  - Constraint: Forms such as `C:foo` are errors because they depend on the current directory.
- Temporary folders
  - Constraint: `GetTemporaryDirectoryPath` returns the OS user temporary folder without a trailing `\`. `CreateTemporaryDirectory` creates a directory directly under that folder based on `Scripting.FileSystemObject.GetTempName`, and the caller is responsible for deletion.
- Destructive operations
  - Constraint: `Move*`, `Copy*`, `Remove*`, `CreateTemporaryDirectory`, and `CreateBackupFile(Move:=True)` mutate the real file system.
- Text encoding
  - Constraint: `TextFileEntity` has no character encoding selection API.
- Writes
  - Constraint: `OpenFile(AsWrite:=True, AsAppend:=False, Force:=False)` errors if the file already exists.

### IPv4 Constraints

- Value representation
  - Constraint: IPv4 values are held in signed `Long`. Addresses with the most significant bit set become negative.
- Mask length
  - Constraint: `ConvertFromMaskLength` accepts only 0-32.
- Mask value
  - Constraint: `ConvertToMaskLength` accepts only masks with contiguous 1 bits.
- Network narrowing
  - Constraint: The implementation of `NarrowNetwork` errors on `/30` through `/32`. The intended specification is not yet confirmed.

## Dependencies

### External

| Dependency | Usage |
| --- | --- |
| Excel Object Model | Workbooks, worksheets, ranges, formats, shapes, application state |
| VBIDE / VBA Extensibility | `UnitTestMain`, `RemoveVBComponents`, temporary execution modules |
| `Scripting.FileSystemObject` | File and directory operations |
| `Scripting.Dictionary` | `ObjectSet`, `ObjectDictionary`, key management, test double records |
| `WScript.Shell` | Environment variable expansion in local paths |
| `VBScript.RegExp` | Regular expressions for paths, IPv4, test detection, and related processing |
| `Forms.TextBox.1` | Clipboard operations |
| Excel `WorksheetFunction.XLookup` | `WorksheetService.XLookup` |

### Internal

`CommonModules/src/CommonModules/common-modules-manifest.tsv` is the canonical machine-readable contract for CommonModules package membership, classifications, direct CommonModule dependencies, and direct external VBA reference requirements. COLLECT locates that manifest through the unique authoring document source set discovered below its explicit Search Root, preserves its validated bytes, and selects each listed source unit by filesystem timestamp and length rather than content or hash. It materializes the manifest, selected source units, and matching form sidecars as a closed flat package in `common_modules_repo` directly below the invocation-start working directory; that package is generated output and never an authoring source. Downstream tooling reads the distributed manifest rather than scraping this Markdown section or VBA source.

The manifest uses strict UTF-16LE with a BOM, CRLF throughout, and exactly one final CRLF. Its physical line grammar is zero or more whole-line comments beginning with `#` in the first code unit, the exact header `ModuleFile<TAB>Categories<TAB>Dependencies<TAB>RequiredReferences`, and one or more contiguous data rows. Comments cannot contain tabs, control characters, or trailing whitespace. Readers reject blank lines, comments after the header, normalization, encoding fallback, and malformed rows.

| Column | Contract |
| --- | --- |
| `ModuleFile` | Ordinally preserved flat basename `<CommonModuleName>.bas`, `.cls`, or `.frm`, with a lowercase extension and a case-insensitively unique identity. |
| `Categories` | Exactly one primary role: `runtime-baseline`, `test-foundation`, `optional`, or `test-double`. The `public-udf` modifier may appear only in the canonical whole-cell spellings `runtime-baseline,public-udf` or `optional,public-udf`; it is not another primary role. |
| `Dependencies` | Empty or an exact, whitespace-free, declaration-ordered comma-separated list of target `ModuleFile` spellings. Multi-module dependency components and cycles are valid, but self-dependencies are not. Runtime-role modules may depend only on runtime-role modules; test-role modules may depend on either role family. |
| `RequiredReferences` | A standalone JSON array of unique, nonempty, already-trimmed human-visible external reference names, in declaration order. Each row lists only references used directly by that source unit, never transitive requirements from its CommonModule dependencies. `Visual Basic For Applications` is omitted because the standard library is always active. |

The following are treated as standard modules to import.

- `Lib_Common`
- `Lib_CommonConstructor`
- `Fx_Common`
- `ApplicationScreenUpdateManager`
- `CommonRunStateManager`
- `DebugInformation`
- `ArrayObject`
- `ObjectList`
- `ObjectSet`
- `ObjectDictionary`
- `Enumerator`
- `IEnumerator`
- `IElementTypeProvider`
- `IEquatable`
- `IStringable`
- `IDuplicateCheckable`
- `IComparable`
- `IWorkbookService`
- `WorkbookService`
- `IWorksheetService`
- `WorksheetService`
- `WorksheetRangeBounds`
- `WorksheetRangeBoundsEnumerator`
- `WorksheetVirtualTable`
- `WorksheetVirtualTableEnumerator`

Dependencies of modules imported as needed are as follows.

| Area | Source module | Dependency modules |
| --- | --- | --- |
| Input screens | `Lib_InputSheet` | `IUserInputSheet`, `UserInputSheet` |
| Input screens | `UserInputSheet` | `IUserInputSheet` |
| Collections and enumeration | `CounterSet` | `Counter` |
| File operations | `Lib_FileSystem` | `IFileSystemService`, `FileSystemService` |
| File operations | `FileSystemService` | `IFileSystemService` |
| File operations | `Lib_TextFile` | `ITextFileService`, `TextFileService` |
| File operations | `TextFileService` | `ITextFileService`, `ITextFileEntity`, `TextFileEntity` |
| File operations | `TextFileEntity` | `ITextFileEntity` |
| File operations | `ITextFileService` | `ITextFileEntity` |

For unit tests, the following are treated as standard modules to import.

- `Lib_UnitTest`
- `UnitTestAssert`
- `TestDoubleBehaviorStore`
- `TestDoubleCallRecord`
- `TestDoubleVariantKeyBuilder`

Other modules imported as needed are as follows.

| Area | Source module | Dependency modules |
| --- | --- | --- |
| Test support | `FileSystemServiceTestDouble` | `IFileSystemService` |
| Test support | `TextFileServiceTestDouble` | `ITextFileService`, `ITextFileEntity`, `TextFileEntity`, `TextFileEntityTestDouble` |
| Test support | `TextFileEntityTestDouble` | `ITextFileEntity` |
| Test support | `UserInputSheetTestDouble` | `IUserInputSheet` |

## Internal Implementation Summary

- Interfaces are defined as classes for VBA `Implements`, and they raise an error when instantiated directly.
- Major services and input sheets have concrete classes and test doubles that implement the same interfaces.
- Comparison, duplicate checking, and stringification are extended through `IComparable`, `IEquatable`, `IDuplicateCheckable`, and `IStringable`.
- Element type contracts are determined by `TypeName` or `IElementTypeProvider.ElementTypeKey`; key generation is centralized in `GetTypedValueKey` / `TestDoubleVariantKeyBuilder`, distinguishing special values, arrays, object references, and interface implementations.
- Excel operation APIs are implemented to restore active workbook and application state before and after operations as much as possible.
