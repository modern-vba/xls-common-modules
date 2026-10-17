# ADR 0007: WorksheetService Array-Formula Copy Is Not Treated as a Source-Destruction Bug

- Status: Accepted
- Date: 2026-05-23

## Context

When copying an array formula, `WorksheetService.pCopyCellCore` temporarily sets a normal formula on the source cell before continuing copy processing. This raised a concern that if the destination operation fails, the source array formula might be broken.

Failure-path unit tests confirmed that the source array formula is preserved even when the destination operation fails.

## Decision

Source destruction during array-formula copy is not treated as a bug in the current implementation.

We will not make an implementation fix, and we will preserve current behavior based on the verified failure-path unit tests.

## Consequences

- No additional fix is made solely to protect the copy source.
- Cleanup that makes restoration responsibility in array-formula copy easier to read is treated as refactoring, not a bug fix.
- If a new source-destruction reproduction condition is found, handle it as a separate decision based on the reproduction steps.
