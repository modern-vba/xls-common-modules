# ADR 0004: Set UnitTestAssert Failure State Through pSetResultMessage

- Status: Accepted
- Date: 2026-05-23

## Context

There was a concern that a failure message could be set in `UnitTestAssert` while `IsFailed` remained `False`, causing `Lib_UnitTest` to treat the test as `OK`.

The checked paths were `pSetUnsupportedComparisonResult`, `IsTypeOf`, and non-array arguments to `EqualsArray` / `NotEqualsArray`. In the current implementation, all of these paths go through `pSetResultMessage`, where `pIsFailed = True` is set.

## Decision

The reported failure paths already become failures in the current implementation, so no additional fix is needed.

`UnitTestAssert` will continue to set failure state through `pSetResultMessage` by setting `pIsFailed`.

## Consequences

- Consistency between failure messages and `IsFailed` is centralized in `pSetResultMessage`.
- New assertion failure paths should use the same path instead of manipulating fields directly.
- Existing tests for array arguments, non-numeric arguments, and non-array arguments are treated as evidence.
