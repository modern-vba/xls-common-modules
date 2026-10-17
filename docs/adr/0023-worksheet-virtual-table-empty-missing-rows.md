# ADR 0023: WorksheetVirtualTable Returns Missing Rows as Empty Ranges

- Status: Accepted
- Date: 2026-06-13

## Context

Callers that read tabular settings want to group multiple `WorksheetRangeBounds` values by header and relative row, treating the same relative row as one record. However, if each field range is physically expanded to the maximum row count, it becomes easy to lose the distinction between input that exists in the original range and rows supplemented for the virtual table.

## Decision

`WorksheetVirtualTable` is not a dedicated API for `IUserInputSheet`; it is provided as a virtual table composed from `ObjectList("WorksheetRangeBounds")` and headers.

The actual range of each field is not expanded with `ExpandRangeBoundsToMax`. The row count of the virtual table is determined from the longest field. For missing rows in shorter fields, the row `ObjectDictionary` stores an empty range with `RowCount = 0` that preserves the original field's column width and has the target relative row position.

## Consequences

- `WorksheetVirtualTable` does not depend on #110 `ExpandRangeBoundsToMax`.
- Callers can distinguish real input rows from missing rows supplemented by the virtual table by checking the returned `WorksheetRangeBounds.IsEmpty`.
- Even for multi-column fields, an empty range for a missing row preserves the original field's column width.
- Business validation such as required headers, blank-row exclusion, required key columns, and duplicate-key detection remains with the caller.
