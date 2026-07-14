# ADR 0003: WorksheetService.CopyRange Does Not Treat Copying From Non-Anchor Array-Formula Cells as a Bug

- Status: Accepted
- Date: 2026-05-23

## Context

There was a concern that if `WorksheetService.CopyRange` starts processing from a cell that is not the anchor of an array formula, temporary writes to non-anchor cells could trigger Excel array-formula modification errors.

In runtime verification, copying `E2:F3`, which does not include the anchor of the array formula in `D2:F3`, did not raise an error. The destination received the array formula correctly, and the source array formula was preserved.

## Decision

Copying a range that starts from a non-anchor cell of an array formula is not treated as a bug in the current implementation.

Until real harm can be reproduced when copying ranges that contain non-anchor cells, we will not change `CopyRange` for this concern.

## Consequences

- Current behavior is preserved based on the verified partial-array-formula copy case.
- If a new reproduction condition is found, handle it as a new decision that supersedes this ADR.
- Implementation cleanup that makes array-formula copy restoration easier to read may be handled as a separate refactoring issue.
