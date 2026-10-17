# ADR 0006: Preserve Existing WorksheetService Excel Error Value String Conversion

- Status: Accepted
- Date: 2026-05-23

## Context

`WorksheetService.pConvertErrorToString` determines Excel error values with `Select Case ErrValue` and `Case CVErr(...)`. The `####` fallback path in `ReadCell(GetText:=True)` uses string conversion through `"" & TargetCell.Value`.

There was a concern that this direct comparison or string concatenation could cause type mismatch errors for `Variant` values that contain Excel error values.

Unit tests with the existing seven Excel error values and narrow-column reads did not reproduce type mismatch on the current paths.

## Decision

Keep the current implementation for string conversion of existing Excel error values. Do not replace it with a different stable-detection method.

This ADR records the decision for the type mismatch concern. Treat version-dependent handling of new Excel error constants as a separate issue.

## Consequences

- Existing error-value conversion paths such as `ReadCell`, `XLookup`, and `pGetFormulaLiteral` do not change.
- Decisions about decoupling Excel version-dependent constants are handled as separate specification issues.
- If new conditions reproduce type mismatch, create a new decision with the reproduction condition.
