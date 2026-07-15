# CommonModules

CommonModules is the VBA foundation shared by multiple Excel macro workbooks.
Terms for common modules are used in meanings that do not depend on any distribution target tool.

## Language

**CommonModules**:
The whole library shared by multiple Excel VBA tools.
_Avoid_: an individual common module, an individual tool

**Common module**:
Shared VBA source whose source of truth is `CommonModules/src/CommonModules`. Same-named files in an individual tool's `common_modules_repo`, document source set, or workbook-local source set are treated as distributed or reflected copies.
_Avoid_: individual module, tool-side copy

**Individual tool**:
An Excel VBA tool that has its own business-specific `.xlsm` and source set, such as a tool under `xls-web-tools`.
_Avoid_: CommonModules, common module

**Document source set**:
The Git-managed VBA source and template workbook belonging to one individual tool document in a vba-dev project.
_Avoid_: generated workbook output, common_modules_repo, workbook-local source set

**Workbook-local source set**:
A project-independent `modules` folder next to a workbook, used when extracting from or importing into that workbook directly.
_Avoid_: document source set, `vba-dev` `sourcePath`

**Individual module**:
Tool-specific VBA source in an individual tool's source set. This does not include copies of common modules.
_Avoid_: common module, whole individual tool

**VBA source**:
Git-managed `.bas` / `.cls` files.
_Avoid_: unexported code inside `.xlsm`, Excel workbook

**Original workbook file**:
An existing Excel workbook file identified as the workbook file itself, not a VBA source file or a generated copy.
_Avoid_: source file, template copy

**Reflect**:
The operation of writing a document source set into a generated workbook used for development or test execution.
_Avoid_: extract, distribute, publish

**Extract**:
The operation of writing VBA source from a workbook into a source set.
_Avoid_: reflect, distribute

**Workbook import**:
The operation of writing VBA source from a source set into an existing workbook in place.
_Avoid_: build, reflect, publish

**Test diagnostic information**:
Information checked to identify causes when a unit test fails. This includes expected and actual values, case names, raised errors, and test-double call history.
_Avoid_: test result itself

**Call arguments**:
For a test-double call that completed successfully, all arguments passed to the actual method, arranged in the original function definition order. These are taken from call history to verify the actual operation that was issued.
_Avoid_: recorded arguments

**Match arguments**:
Key arguments used with the method name to match a test double's return value, output value, error, or call history. These are not necessarily all arguments of the actual method and are treated separately from call arguments.
_Avoid_: call arguments, recorded arguments

**Input sheet**:
A worksheet where users enter processing conditions and settings. This refers to the whole sheet, not only the rectangular range searched by logic.
_Avoid_: input area

**Input area**:
A rectangular range inside an input sheet that is treated as the target for reading item names and values. This refers to the area containing the user-input table, not the entire sheet.
_Avoid_: whole input sheet

**Virtual table**:
A logical table that associates multiple columns read from an input area by common headers and relative rows. It is not the actual rectangular range on the worksheet.
_Avoid_: input table, real table, input area

**Button**:
A worksheet shape responsible for running a macro through user operation. This does not refer only to UnitTest rerun buttons or shapes created by `AddButton`.
_Avoid_: UnitTest button, AddButton-created button

**Common service**:
The shared access foundation that CommonModules provides to each tool. This refers to Excel workbooks, worksheets, the file system, and text files; it does not include debug information or progress display.
_Avoid_: common run state

**Common run state**:
Debug information and progress-display state held only during one GUI run. This is not a common service; it is state tied to a run scope.
_Avoid_: common service

**UDF-safe**:
A property of not requiring side effects that are inappropriate for a recalculation context when called as an Excel worksheet function during cell recalculation. Treat this separately from being safe to use from GUI or batch entrypoints.
_Avoid_: safety during macro execution

**Public UDF**:
A standard-module function that CommonModules intentionally exposes for use from Excel worksheet formulas. This refers only to functions intended for use as cell formulas, not all `Public Function` members callable from VBA.
_Avoid_: all public standard functions, public class methods

**Public API**:
The callable surface CommonModules exposes externally as contracts with distribution target tools and test doubles. Separate from public UDFs, this includes public members of services, value objects, and collections used from VBA, plus their interface and test-double contracts.
_Avoid_: public UDF, internal helper, names used only in comments

**Range shape**:
The dimensions represented by `WorksheetRangeBounds`, consisting of row count and column count. This refers to a size that can be handled with the same relative row and column numbers, not absolute start or end coordinates.
_Avoid_: end position, absolute coordinates

**Typed element collection**:
A collection that has an element type contract and holds only elements that follow the same contract. Elements include primitive values, arrays, Excel error values, object references, and Nothing.
_Avoid_: object-only collection, value collection

**Keyed typed element collection**:
A typed element collection whose elements can be referenced by key. It treats the key contract and value element type contract separately, and applies the same element type contract as typed element collections to values.
_Avoid_: unkeyed typed element collection, object-only dictionary, arbitrary-type dictionary

**Key contract**:
The range of keys accepted by a keyed typed element collection and the rules for considering keys equal. Treat this separately from the value element type contract.
_Avoid_: value element type contract, element capability contract

**Key comparison mode**:
The mode in a keyed typed element collection that determines whether string-key matching is case-sensitive. Treat this separately from the element type contract.
_Avoid_: element type contract, identity/duplicate detection mode

**Explicit element type contract**:
A contract where a typed element collection has the target element type before waiting for the first element to be added.
_Avoid_: implicit type inference only

**Assignable type acceptance**:
The behavior where a typed element collection accepts an element based on assignability to the specified element type, rather than exact concrete class name.
_Avoid_: exact concrete type name match only

**Element type self-reporting**:
The behavior where an object stored in a typed element collection returns which element type contract it should be treated as. This is used when CommonModules does not directly know a caller-specific interface type.
_Avoid_: type contract based only on concrete class name

**Element capability contract**:
A contract in a typed element collection requiring elements with the same element type contract to have the same capabilities for comparison, identity, and duplicate detection. This does not include display or string-conversion capability.
_Avoid_: display capability contract, string-conversion capability contract

**Required capability**:
Behavior a typed element collection requires from elements stored in the same collection. Behavior that is not required is not included in the condition for accepting an element.
_Avoid_: optional capability, display string conversion

**Identity/duplicate detection mode**:
The exclusive criterion a typed element collection uses to consider elements the same or duplicate. Treat this separately from ordering requirements.
_Avoid_: ordering, display string conversion

## Example Dialogue

Developer: Does `ClearButton` remove only the UnitTest rerun button?
Domain expert: No. A button is a worksheet shape responsible for running a macro; it is not UnitTest-specific.
