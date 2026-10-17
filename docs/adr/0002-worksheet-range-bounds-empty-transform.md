# ADR 0002: Empty WorksheetRangeBounds Can Become Non-Empty Through Transform

- Status: Accepted
- Date: 2026-05-23

## Context

An empty `WorksheetRangeBounds` is not just a marker that there are no target cells. It is represented as a zero-size range that still retains the start position and one side of the row or column range.

If an empty range created by a non-intersecting column operation is `Row:=1, Column:=3, FinishRow:=10, FinishColumn:=0`, calling `Transform(AddColumn:=1)` uses the retained position and row range and turns it back into the non-empty range `C1:C10`.

## Decision

An empty `WorksheetRangeBounds` is treated as a zero-size range with positional information.

We will keep the behavior where `Transform` can turn an empty range back into a non-empty range by assigning a row count or column count again.

## Consequences

- `Transform` is not an API that always keeps empty ranges empty.
- Callers that want to freeze an empty range as "no target" must check `IsEmpty` and not continue into transformation logic.
- Existing unit tests that expect empty ranges to revive remain valid.
