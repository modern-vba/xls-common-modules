# VBA Coding Conventions

## 1. Differences From Typical VB/VBA Style

In this workspace, consistency with `vba-dev` build/test/import/export, DoxyVB6 comment parsing, common-module distribution, and the unit-test runtime takes priority over typical VB/VBA style. Some patterns that are often tolerated in ordinary VBA should be avoided in this repository.

The following points are especially easy to get wrong. See the later sections for the detailed rules.

| Item | Rule in this repository | Details |
| --- | --- | --- |
| Implicit visibility | In standard modules, explicitly write `Public` or `Private`. | [4. Naming, Scope, and Types](#4-naming-scope-and-types) |
| Public functions in standard modules | Add `Option Private Module` except in `Fx_*.bas` and `Test_*.bas`. | [3. Module Structure](#3-module-structure) |
| Function return values | Assign the return value immediately before returning. Do not rely on default values in the middle of the function. | [5. Procedures, Classes, and Interfaces](#5-procedures-classes-and-interfaces) |
| Arguments | Do not omit `ByVal` / `ByRef` or the type. Always specify a default value for `Optional` arguments. | [4. Naming, Scope, and Types](#4-naming-scope-and-types) |
| `Variant` / `Object` | Do not specify these types unless truly unavoidable; prefer concrete types, existing interfaces, typed wrappers, or dedicated classes. | [4. Naming, Scope, and Types](#4-naming-scope-and-types) |
| Local variable names | Use `snake_case`, not Hungarian notation. Use at least two words and include at least one underscore. | [4. Naming, Scope, and Types](#4-naming-scope-and-types) |
| Excel object access | Avoid direct dependency on `ActiveSheet` or `Selection`; prefer service layers and `WorksheetRangeBounds`. | [6. Excel Operations, UI Entrypoints, and Error Handling](#6-excel-operations-ui-entrypoints-and-error-handling) |
| Common service initialization | Do not initialize implicitly from deep logic. Initialize at the entrypoint or in the test Arrange step. | [6. Excel Operations, UI Entrypoints, and Error Handling](#6-excel-operations-ui-entrypoints-and-error-handling) |
| Comments | Public APIs must have comments in the `'* ` form that DoxyVB6 can parse. | [7. Comments and Documentation](#7-comments-and-documentation) |
| Tests | Use `Public Sub Test_...(ByVal Assert As UnitTestAssert)` so `UnitTestMain` can discover the test. | [8. Tests](#8-tests) |
| Common-module changes | Edit the source of truth in `xls-common-modules\CommonModules\src\CommonModules`, not distributed copies. | [9. Common Modules, Synchronization, and External Dependencies](#9-common-modules-synchronization-and-external-dependencies) |

When in doubt, prioritize reproducibility after `vba-dev` build/test/import/export, ease of replacing dependencies in tests, and safe synchronization across multiple workbooks over superficial readability.

## 2. Core Principles

- Separate tool-specific logic from shared logic. Logic used by multiple tools should live in modules from `common_modules_repo` or in synchronized modules.
- Understand the functionality already provided by common modules, and do not independently reimplement duplicate functionality.
- Localize direct access to Excel objects. In normal logic, use `WorkbookService`, `WorksheetService`, and `WorksheetRangeBounds`.
- Localize direct access to the file system and text files. In normal logic, use `FileSystemService` and `TextFileService`.
- Represent complex values with classes, and prefer `ObjectList`, `ObjectSet`, or typed wrappers around them for collections.
- In tool-specific code, initialize common infrastructure services at the entrypoint or in the test Arrange step. Do not scatter initialization inside `Mod_...`, tool-specific classes, or helper functions.
- In common modules, implicit initialization is allowed only for the minimum services needed to avoid breaking public APIs that may be used standalone.
- Entrypoints called from worksheet buttons must be centralized in `GUIHandler.bas`.
- Do not swallow runtime errors. Display them at the entrypoint, and inside the implementation re-raise with a traceable `Source` and `Description`.
- Always assign variables explicitly before using them or before returning a value; do not rely on default values.

## 3. Module Structure

### 3.1 File Placement and Synchronization Boundaries

- For vba-dev projects, VBA sources go under the document source set declared by `vba-project.json` `documents.*.sourcePath`.
- Workbook-local `modules` directories are used only for explicit workbook-level `vba-dev import --from <book-dir>\modules --to <book.xlsm>` or `vba-dev export --from <book.xlsm> --to <book-dir>\modules` workflows.
- Shared modules are copied from each project root's `common_modules_repo` into document source sets by `vba-dev common-module update`. The selected files are determined by `vba-project.json` `commonModules` entries and the CommonModules manifest.
- When updating common modules, first edit `xls-common-modules\CommonModules\src\CommonModules`, verify with `vba-dev test`, collect into `common_modules_repo`, distribute that repo, and run `UPDATE_COMMON_MODS` for vba-dev project consumers when needed.
- Do not edit distributed `common_modules_repo` copies, document source set copies, or workbook-local source set copies first.
- Excel workbooks are managed as `.xlsm`; VBA sources are managed as `.bas` / `.cls`.
- The exported file's `Attribute VB_Name` must match the file name.
- Preserve existing class export headers.
- Module names must fit within Excel's 31-character limit. Considering the `Test_` prefix for tests, keep module and class names to 26 characters or fewer when practical.
- For foundational types where meaning is more important, names may exceed 26 characters as long as they fit within 31 characters. If adding `Test_` would exceed 31 characters, shorten from the end without breaking the meaning.

Standard class export header:

```vb
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ObjectList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
```

### 3.2 Module Naming

| Kind | Name | Purpose |
| --- | --- | --- |
| UI entrypoint | `GUIHandler.bas` | Logic called from worksheet buttons or click events |
| General-purpose library | `Lib_Xxx.bas` | Functions that can be shared by multiple tools |
| Tool-specific logic | `Mod_Xxx.bas` | Business logic for a specific tool |
| Worksheet functions | `Fx_Xxx.bas` | Public functions used as Excel worksheet functions |
| Factory | `Constructor.bas` | `New_Xxx` functions that act like argument-taking constructors |
| Global constants/state | `Grobal.bas` | Preserve the existing name. Consider moving to `Global.bas` for new code |
| Interface | `IXxx.cls` | Contract implemented with `Implements` |
| Test | `Test_Xxx.bas` | Tests discovered and run by `UnitTestMain` |
| Test double | `XxxTestDouble.cls` | Stubs, spies, and fakes |

### 3.3 Module Preamble

- Put `Option Explicit` in every `.bas` / `.cls` file.
- Add `Option Base 0` in modules containing logic that depends on the lower bound of arrays.
- In standard modules, use this order: `Attribute VB_Name`, `Option Explicit`, `Option Base 0` when needed, `Option Private Module` when needed, then the module comment.
- Add `Option Private Module` to standard modules except `Fx_*.bas` and `Test_*.bas`. This prevents non-UDF `Public Function` members from appearing in Excel's formula suggestions; it is not a reason to convert public APIs that return values into `Sub` procedures with `ByRef` outputs.
- Put UDFs exposed as Excel worksheet functions in `Fx_*.bas`, and do not add `Option Private Module`. Start by collecting public UDFs in `Fx_Common.bas`, then split into `Fx_...` modules once responsibilities grow.
- In class modules, put `Attribute VB_Name`, `Option Explicit`, and `Option Base 0` when needed at the top.
- `Implements`, module constants, and Private fields may appear before the module comment when needed for DoxyVB6 comment parsing or VBA ordering constraints.
- If the class is an interface, place a blank line after the `Option` declarations and then write `'#Interface`.

Example standard-module preamble:

```vb
Attribute VB_Name = "Mod_Sample"
Option Explicit
Option Base 0
Option Private Module

' #############################################################################
'!
'! @brief
'! Standard module containing sample processing.
'!
' #############################################################################
```

## 4. Naming, Scope, and Types

### 4.1 Scope

- Only procedures that need to be public should be `Public`.
- Procedures, helpers, and constants used only within a module should be `Private`.
- Avoid omitting `Public` in standard modules when declaring a public member.
- Class internal state should use `Private pXxx` fields, exposing only what is needed through `Property Get` / `Let` / `Set`.
- If both read access through `Get` and write access through `Set` / `Let` are accepted, exposing the field directly may be allowed.
- Global variables are limited to replaceable common infrastructure service instances such as `WbSrv`, `WsSrv`, `FsSrv`, `TfSrv`, `DbgInfo`, and `ProgStat`.
- Avoid adding new global variables. If one is necessary, put it in `Lib_Xxx` with its initialization procedure, and allow short public names only for frequently used infrastructure services.

### 4.2 Naming

VBA is case-insensitive, so avoid names that can collide easily.

- Do not use Hungarian notation.
- Use English names.
- Modules, classes, public procedures, and public properties use PascalCase.
- Outside local variables, do not abbreviate words like local variables. Exceptions are allowed for short, conventional infrastructure names such as `WsSrv`, `FsSrv`, and `DbgInfo`.
- Acronyms and initialisms that are common or fixed should be treated as one word. Examples: `IpAddress`, `PanDevice`.
- Public procedures used from worksheets should use uppercase letters and digits only.
- Local variables should use `snake_case` with at least two words and at least one underscore. Abbreviations such as `cfg_lst` for `configuration_list` are encouraged.
- Class Private fields use `p` + PascalCase.
- Private helper procedures use `p` + PascalCase.
- Global constants use `G_` + uppercase snake case.
- Module constants use `C_` + uppercase snake case.
- Avoid collisions with reserved words or Excel VBA built-in objects. Allow an exception only when you are confident it is safe and it makes the code more intuitive.

```vb
Private Const C_COL_START As Long = 1
Public Const G_ROW_MAX As Long = 1048576

Private pIsInitialized As Boolean
Private pWorksheetName As String

Public Property Get WorksheetName() As String
    Call pCheckInit
    WorksheetName = pWorksheetName
End Property

Private Sub pCheckInit()
    If Not pIsInitialized Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class Sample", Description:="Object is not initialized."
    End If
End Sub
```

### 4.3 Types and Declarations

- Explicitly specify types for variables, arguments, and return values.
- Do not use implicit type-declaration characters, untyped `Const`, or untyped `Function`.
- Row numbers, column numbers, array indexes, and counts should generally be `Long`.
- `Integer` may be used for values whose range is small and fixed by specification, such as IPv4 mask lengths, bit positions, progress percentages, and small chunk-size constants used only internally.
- Excel dates use `Date`; an unset date uses the existing `G_DATE_NULL`.
- Always use `Set` for object assignment.
- Arguments should generally be `ByVal`. Use `ByRef` only when needed, such as for return-value arguments or array updates.
- Explicitly write `ByVal` and `ByRef`.
- Put `ByRef` arguments used as outputs at the beginning of the argument list.
- Always specify default values for `Optional` arguments.
- Do not specify `Variant` / `Object` types unless it is truly unavoidable. Prefer concrete types, existing interfaces, typed wrappers, or dedicated classes when they can express the value.

## 5. Procedures, Classes, and Interfaces

### 5.1 Procedure Design

- Keep entrypoint procedures short, and delegate real work to `Mod_...` modules or classes.
- In functions that return objects, use `Set New_Xxx = obj`.
- Assign function and property return values immediately before `Exit Function` / `End Function` / `Exit Property` / `End Property`.
- Use `DbgInfo.StartTask` / `FinishTask` when a whole function needs tracing.
- Constructor-like logic should generally be placed in `New_Xxx` functions in `Constructor.bas`.
- Constructor-like functions for foundational common types such as `WorksheetRangeBounds` go in `New_Xxx` functions in `Lib_CommonConstructor.bas`.
- `New_Xxx` factories should be limited to object creation and `Initialize...` calls; real work should be delegated to the target class.

```vb
Public Function New_ZoneInformation( _
        ByVal DeviceName As String, _
        ByVal ZoneName As String) As ZoneInformation

    Dim zone_info As ZoneInformation
    Set zone_info = New ZoneInformation

    zone_info.DeviceName = DeviceName
    zone_info.ZoneName = ZoneName

    Set New_ZoneInformation = zone_info
End Function
```

### 5.2 Class Design

- Classes that require initialization should have `pIsInitialized`, and public members should call `pCheckInit` at the start.
- VBA's lack of argument-taking constructors is handled with `New_Xxx` factories.
- Classes that need value comparison, string representation, or duplicate detection should implement the existing `IEquatable`, `IComparable`, `IStringable`, or `IDuplicateCheckable`.
- Do not expose collections directly; provide typed add, remove, and enumeration methods.
- `Class_Initialize` should only initialize internal fields. Avoid heavy operations on external workbooks or worksheets.
- In `Class_Terminate`, perform only cleanup that is safe if it fails. If `On Error Resume Next` is needed, keep its scope minimal.

### 5.3 Interface Definitions

- Interfaces use `I` + PascalCase `.cls` names and contain only the contract implemented with `Implements`.
- Add an interface when polymorphism is needed, such as multiple implementations, test doubles, or swappable services.
- `Attribute VB_Name` must match the file name. Place `'#Interface` after a blank line following `Option Explicit` / `Option Base 0` when needed.
- In `Class_Initialize`, prevent direct instantiation by raising with `Source:="Interface IXxx"` and `Description:="This class is for interface use."`
- Interfaces must not have internal state, Private fields, implementation helpers, or default behavior.
- Except for `Class_Initialize`, include only empty definitions of public `Sub`, `Function`, and `Property` members.
- Contract members must explicitly write `Public`, and their argument names, types, `ByVal` / `ByRef`, `Optional` default values, and return values must be in a form that implementations can match exactly.
- In implementation classes, implement members as `Private Function IXxx_MemberName(...)`, using the `InterfaceName_MemberName` pattern.
- Contract members exposed through `Implements` should also be exposed as same-named `Public` members on the concrete class.
- Interface implementation members should generally delegate to the same-named public member.
- Public members in interfaces should also have Doxygen-style comments that describe the behavior callers may expect, return values, and error conditions in `@details`.

```vb
Attribute VB_Name = "ISampleService"
Option Explicit
Option Base 0

'#Interface

' #############################################################################
'!
'! @brief
'! Interface that makes sample processing replaceable.
'!
' #############################################################################

Private Sub Class_Initialize()
    Err.Raise Number:=vbObjectError + 1, Source:="Interface ISampleService", Description:="This class is for interface use."
End Sub

'* Reads a sample value.
'*
'* @param Key Key to read.
'* @return Sample value.
'*
'* @details
'* Returns the value for Key. If the value cannot be retrieved, the implementation re-raises the implementation-specific error.
Public Function ReadValue(ByVal Key As String) As String
End Function
```

### 5.4 Member Placement and Formatting

- Put at least one blank line immediately after normal `End Property`, `End Sub`, and `End Function` lines.
- Do not put a blank line between `Property Get` / `Property Let` / `Property Set` members of the same property.
- An `Implements` member that only delegates to a same-named public member should be placed immediately after the delegated-to public member with no blank line between them.
- In class modules, unless there is a special reason, order members as: properties, collection `Item`, `Class_Initialize`, `Initialize...`, `Class_Terminate`, public methods, then shared Private helpers.
- A Private helper used by only one public `Function` / `Sub` should be placed immediately after that public member.
- A Private helper used only by `Class_Initialize`, `Class_Terminate`, or `Initialize...` should be placed immediately after the corresponding initialization or termination logic.
- In standard modules, preserve the broad ordering of existing public APIs and move only clearly related Private helpers close to their callers.
- Private helpers used by multiple public members of the same kind or processing group should be placed after that public member group.
- Doxygen-style comments immediately before a member are treated as part of that member's block.
- Member attributes such as `Attribute NewEnum.VB_UserMemId = -4` are treated as part of that member's block.
- When moving a member, move its attached comments and attributes with it.
- `.cls` files with `'#Interface` are excluded from normal class member reordering. Apply only blank-line rules and comment attachment handling.
- Contract members in interfaces and the corresponding public members and `Implements` members in concrete classes should appear in the same order.
- The VBA formatter lives under `xls-common-modules\tools`, and formatting targets exported `.bas` / `.cls` VBA sources in document source sets or workbook-local source sets.

## 6. Excel Operations, UI Entrypoints, and Error Handling

### 6.1 Excel Operations

- Prefer `WorksheetRangeBounds` and `New_RangeBounds` over string concatenation for range specification.
- Workbook and worksheet operations should go through `WbSrv` / `WsSrv` where practical, so they can be replaced with test doubles.
- Direct access to `ThisWorkbook`, `ActiveWorkbook`, `ActiveSheet`, and `Selection` is limited to UI entrypoints, service implementations, test setup, and unavoidable Excel-specific behavior not provided by services.
- When stopping screen updates, events, or calculation, use `ApplicationScreenUpdateManager` and always restore the settings.
- In write operations, make handling of values, formulas, number formats, and type conversion explicit through argument names.

```vb
Dim range_bounds As WorksheetRangeBounds
Set range_bounds = New_RangeBounds(Row:=2, Column:=2, Sheet:="InputSheet")

Call WsSrv.WriteCell(range_bounds, "12345", TypeConvert:=False)
```

### 6.2 UI Entrypoints

- Procedures called directly from buttons belong in `GUIHandler.bas`.
- In UI entrypoints, initialize in this order: `InitializeCommonService(Force:=True)`, `CommonRunStateManager`, `ApplicationScreenUpdateManager.DisableUpdates`, and then `On Error GoTo ON_ERROR`.
- Common infrastructure services used by tool-specific logic should be initialized explicitly in `GUIHandler.bas`; do not scatter `InitializeCommonService` into called `Mod_...` modules or classes.
- In standard-form UI entrypoints, use `GoTo ON_EXIT` instead of `Exit Sub` for early exits.
- Put `Exit Sub` at the end of both `ON_EXIT` and `ON_ERROR`.
- In standard-form UI entrypoints, place `' ==== Main process ========` immediately before the real work and `' ==== Main process ends ========` immediately after it.
- On normal completion or cancellation, do not explicitly dispose `ApplicationScreenUpdateManager` or `CommonRunStateManager`; let local variable destruction handle it.

```vb
Public Sub Xxx_Click()
    Call InitializeCommonService(Force:=True)

    Dim RunState As CommonRunStateManager
    Set RunState = New CommonRunStateManager

    Dim ASUM As ApplicationScreenUpdateManager
    Set ASUM = New ApplicationScreenUpdateManager
    Call ASUM.DisableUpdates

    On Error GoTo ON_ERROR

    ' ==== Main process ========

    ' For early exit:
    ' GoTo ON_EXIT

    ' ==== Main process ends ========
ON_EXIT:
    On Error GoTo 0
    Exit Sub
ON_ERROR:
    Dim err_desc As String
    Dim err_num As Long
    Dim err_source As Variant
    Dim debug_lines As String

    err_desc = Err.Description
    err_num = Err.Number
    err_source = Err.Source

    On Error Resume Next
    If Not ASUM Is Nothing Then Call ASUM.Restore
    If Not DbgInfo Is Nothing Then debug_lines = vbCrLf & DbgInfo.BuildMessageLines()
    On Error GoTo 0

    MsgBoxPage err_desc & " (0x" & Hex(err_num) & ") @" & err_source & debug_lines
    Exit Sub
End Sub
```

### 6.3 Worksheet Function Entrypoints

- Public UDF entrypoints executed as Excel worksheet functions must not call `InitializeCommonService`.
- Even if the current logic does not directly use common services, call `InitializeUdfCommonService` without `Force`.
- `InitializeUdfCommonService` initializes only `WbSrv` / `WsSrv`, avoiding optional service initialization through `Application.Run` during cell recalculation.

### 6.4 Error Handling

- Do not use `On Error Resume Next` routinely in internal functions. Use it only for short ranges where exceptions are expected, such as existence checks, and immediately return to `On Error GoTo 0`.
- In `Err.Raise`, specify `Number`, `Source`, and `Description`; put `Class Xxx` / `Function Xxx` / `Sub Xxx` in `Source`.
- When re-raising an error, preserve the original `Err.Number`, `Err.Source`, `Err.Description`, `Err.HelpFile`, and `Err.HelpContext`.
- In UI entrypoint error handling, explicitly restore Excel settings immediately before displaying the error message, and append debug information only when `DbgInfo` exists.

## 7. Comments and Documentation

- Add Doxygen-style comments to modules, classes, public procedures, and public properties.
- Comments must be written in English.
- Module and class descriptions use only the implicit `@brief`.
- Module documentation comments are placed immediately after the declarations described in [3.3 Module Preamble](#33-module-preamble).
- Module documentation comments use a block enclosed by `' #############################################################################`, with the body written on `'!` lines.
- A module description should describe only the module's responsibility in one or two sentences.
- Public property descriptions are written once immediately before the property, not before each accessor.
- Public property descriptions do not specify `@brief`; describe what the property is with a concise noun phrase. Do not write anything other than the implicit `@brief`.
- For public procedures, write a summary first, then always write `@param` when arguments exist and `@return` when a return value exists.
- Public procedures must always include `@details`.
- For `ByRef` arguments used as outputs, prefix the description with `[Output] `.
- Comments should focus on why the logic is necessary and on Excel-specific or business-rule caveats.
- Do not add comments that only explain simple assignments or loops.
- To make public APIs available to DoxyVB6 documentation generation, public API comments must start with `'* `.
- Do not write documentation comments for test subprocedures such as `Public Sub Test_...`.

Module documentation comment example:

```vb
' #############################################################################
'!
'! @brief
'! Standard module containing application-information reading logic.
'!
' #############################################################################
```

Documentation comment examples for a public property and public procedure:

```vb
'* Target worksheet name.
Public Property Get TargetSheetName() As String
    TargetSheetName = pTargetSheetName
End Property

'* Reads the requester name.
'*
'* @param MultiRowSheetName Sheet name to read. If omitted, the default sheet is used.
'* @return Requester name.
'*
'* @details
'* Reads the requester name from the target sheet and returns it as a string.
Public Function ReadRequesterName(Optional ByVal MultiRowSheetName As String = "") As String
    ReadRequesterName = "Requester Name"
End Function
```

## 8. Tests

### 8.1 Test Files and Procedures

- Local variable names follow the same naming rules as normal code.
- Test modules use `Test_Xxx.bas`, and `Attribute VB_Name` must match the file name.
- Unit test files should generally have a one-to-one relationship with the target module.
- Tests for standard module `Mod_Xxx.bas` use `Test_Xxx.bas`, and tests for class module `Xxx.cls` use `Test_Xxx.bas`. Remove `Lib_` and `Mod_` from the test file name.
- Do not mix tests for multiple standard modules or class modules in a single `Test_*.bas`.
- If existing tests are mixed, split them into target-specific `Test_*.bas` files.
- When testing an individual function in a standard module, the test procedure name may include the function name, but the file name must correspond to the target standard module.
- If adding `Test_` would exceed Excel's 31-character module-name limit, shorten from the end without breaking the meaning.
- `UnitTestMain` discovers and runs procedures in standard modules whose names start with `Test_` and whose argument is `UnitTestAssert`.
- Run tests for vba-dev projects through `vba-dev test --project <project> --document <document>`. Use `--module` and `--procedure` for targeted test runs.
- To make a test discoverable, use `Public Sub Test_...(ByVal Assert As UnitTestAssert)` as the basic form.
- Test procedure names should generally use the `Test_Target_Condition_ExpectedResult` form. Examples: `Test_Initialize_NegativeUnitCount_RaisesError`, `Test_ValidateWorkingHour_FullDayOffWithoutInput_ReturnsEmptyMessage`.

### 8.2 Test Body

- Organize test bodies in Arrange, Act, Assert order.
- Each procedure should verify one condition and one expected result.
- Prefer test granularity where a failure can be identified from the `Test Item` on `UNIT_TEST_SHEET`, rather than packing multiple patterns into a loop.
- Set `On Error Resume Next` at the start of the test procedure, and after Act, pass `Err.Number`, `Err.Source`, and `Err.Description` to `UnitTestAssert` for verification.
- In tests that do not expect an error, before continuing to additional verification, write `If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub`.
- In tests that expect an error, use `Assert.ErrorRaised expected_error_number, Err.Number, Err.Source, Err.Description`.
- For existing specifications that do not fix the error number, pass `0` and verify only that an error was raised. If the business message is part of the specification, also verify `Err.Description` with `Assert.Equals`.
- Use the main assertions `Assert.Equals`, `Assert.NotEquals`, `Assert.EqualsNumeric`, `Assert.IsTrue`, `Assert.IsFalse`, `Assert.IsNothing`, `Assert.IsEmpty`, and `Assert.EqualsArray`.
- Prefer `EqualsNumeric` for values whose numeric type can vary, such as values from Excel functions or Range values.

Basic test module form:

```vb
Attribute VB_Name = "Test_Xxx"
Option Explicit

' #############################################################################
'!
'! @brief
'! Unit tests for Xxx.
'! These tests are run by Lib_UnitTest.UnitTestMain().
'!
' #############################################################################

Public Sub Test_Target_Condition_ExpectedResult(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim expected_val As Variant
    expected_val = "expected"

    Dim target_obj As Xxx
    Set target_obj = New Xxx

    ' --- Act ---
    Dim actual_val As Variant
    actual_val = target_obj.SomeFunction()

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub

    Assert.Equals expected_val, actual_val
    Assert.ErrorNotRaised 0, Err.Number, Err.Source, Err.Description
End Sub

Public Sub Test_Target_InvalidCondition_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim target_obj As Xxx
    Set target_obj = New Xxx

    ' --- Act ---
    Call target_obj.SomeProcedure("invalid")

    ' --- Assert ---
    Assert.ErrorRaised 0, Err.Number, Err.Source, Err.Description
    Assert.Equals "Expected error message", Err.Description
End Sub
```

### 8.3 Test Doubles and External Dependencies

- Replace external dependencies with test doubles such as `WorksheetServiceTestDouble`, `WorkbookServiceTestDouble`, and `FileSystemServiceTestDouble`.
- If a test unavoidably includes direct Excel operations, prepare a dedicated test sheet and avoid dependency on existing state.
- Set up any required state in Arrange so repeated runs produce the same result.
- Test sheet names start with `tmp_test_`.
- In logic that uses global services, call `InitializeCommonService` in Arrange so the test does not depend on prior test state.
- Do not create unit tests that generate real files; verify such behavior manually.

## 9. Common Modules, Synchronization, and External Dependencies

### 9.1 Common Modules and Synchronization

- When changing common modules, add or update the corresponding `Test_...` under `xls-common-modules\CommonModules\src\CommonModules`, then verify with `vba-dev test --project xls-common-modules\CommonModules`.
- When changing target tool-specific modules, add or update tests under that project's document source set or workbook-local source set, then verify the target project/document with `vba-dev test`.
- When editing common modules, check diffs in `xls-common-modules\CommonModules\src\CommonModules`, each `common_modules_repo`, the relevant document source set copies, and generated `.xlsm` outputs.
- When changing modules listed in `sync.json`, run `SYNC_MODS` and also verify behavior in synchronization targets with `vba-dev test` when tests exist.
- Before sharing logic, confirm it does not contain tool-specific sheet names, column numbers, business messages, or file paths.
- Shared modules may contain functions that current tools do not use. Before deleting a function, search references across all tools.

### 9.2 References and External Dependencies

- When using a type that requires a reference, confirm the reference is enabled in the target workbook.
- Main dependencies used by existing code are `Scripting.Dictionary`, `VBScript.RegExp`, `VBIDE`, `Forms.TextBox.1`, and the Excel Object Model.
- Do not add new external COM dependencies. If one is necessary, consider late binding or isolating it behind an existing service.
- Use `Application.VBE` only in necessary places such as the test runtime, and document in README or equivalent documentation that users must enable "Trust access to the VBA project object model."

### 9.3 Encoding and Line Endings

- For exported `.bas` / `.cls`, prioritize compatibility with existing files.
- New documentation and PowerShell scripts use UTF-8 by default.
- When handling Japanese comments or strings in VBA sources, verify that text is not corrupted after Excel import/export.
- Preserve existing Windows CRLF line endings.
- If a repository-specific `AGENTS.md` specifies encoding or line endings, follow that specification.
