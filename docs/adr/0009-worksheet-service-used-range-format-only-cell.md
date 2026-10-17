# ADR 0009: WorksheetService.GetUsedRangeBounds May Treat Format-Only Cells Without Values as Empty

- Status: Accepted
- Date: 2026-05-23

## Context

When UsedRange has only one cell, `WorksheetService.pGetRawUsedRange` treats it as an empty range if the value is empty and there are no borders on the four sides.

It does not check non-value usage such as fill, font, number format, comments, hyperlinks, or validation. As a result, a format-only cell without a value or borders may be treated as empty by `GetUsedRangeBounds(GetRawRange:=True)`.

## Decision

Cells that contain only formatting other than borders and no value are specified as not being retained as used range.

We will not add support for treating formatting alone as used range.

## Consequences

- Logic that depends on UsedRange, such as `CopyRange(CopyNumberFormat:=True)`, may exclude format-only cells.
- Cells that should remain in the used range must have a value or an element that the current detection treats as used, such as borders.
- If a requirement appears to retain format-only cells, consider it as another mode of `GetUsedRangeBounds` or as a separate API.
