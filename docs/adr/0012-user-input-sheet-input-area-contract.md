# ADR 0012: UserInputSheet Accepts the Input Area as WorksheetRangeBounds

- Status: Accepted
- Date: 2026-05-27

## Context

`UserInputSheet` is used to retrieve value ranges that correspond to item names on an input worksheet. The previous initialization API accepted a worksheet name and workbook name, and treated the entire used range of that worksheet as the input target.

When descriptions, history sections, auxiliary tables, or other content exist on the same worksheet, callers need to search only the input area rather than the whole worksheet. Splitting the worksheet name, workbook name, start row, start column, end row, and end column into separate arguments would duplicate the information already represented by `WorksheetRangeBounds` and would make inconsistencies more likely.

Also, when `IUserInputSheet` exposes `WorksheetName` and `WorkbookName`, callers tend to keep depending on worksheet-level concepts instead of the input area. The public `WbSrv` and `WsSrv` variables on `UserInputSheet` are no longer needed as test-substitution APIs now that `UserInputSheetTestDouble` exists.

## Decision

`UserInputSheet.Initialize` and `New_InputSheet` accept a `WorksheetRangeBounds` and treat that range as the input area. Callers that want the whole worksheet as the input area pass a `WorksheetRangeBounds` that contains only the worksheet name.

`IUserInputSheet` exposes `InputArea` and `GetItemRange`. It does not expose `WorksheetName` or `WorkbookName`. Callers that need the worksheet name or workbook name refer to `InputArea.WorksheetName` and `InputArea.WorkbookName`.

`InputArea` is returned as a copy, so external references cannot mutate the input area held inside `UserInputSheet`. Initializing with `Nothing` or an empty input area is an error.

`UserInputSheet` does not expose `WbSrv` or `WsSrv`. Its internal processing uses `Lib_Common.WbSrv` and `Lib_Common.WsSrv` after common services have been initialized.

`GetItemRange` treats the leftmost column of the input area as the first item-name column. The search target is trimmed to the lower-right edge of the used range inside the input area, but it never references rows above or columns left of the input area's upper-left corner, nor rows below or columns right of its lower-right corner.

## Consequences

- Calls such as `New_InputSheet("input")` migrate to forms such as `New_InputSheet(New_RangeBounds(Sheet:="input"))`.
- Callers that need to restrict the input area pass a `WorksheetRangeBounds` with row, column, end row, and end column set.
- Code that depends on `IUserInputSheet.WorksheetName` or `IUserInputSheet.WorkbookName` migrates to references through `InputArea`.
- Service substitution through `UserInputSheet.WbSrv` and `UserInputSheet.WsSrv` is removed. Tests use `UserInputSheetTestDouble` or common-service substitution instead.
