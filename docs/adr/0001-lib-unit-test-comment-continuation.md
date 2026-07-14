# ADR 0001: Treat Comment-Line Continuation as Actual VBA Behavior in Lib_UnitTest Logical-Line Reading

- Status: Accepted
- Date: 2026-05-23

## Context

`Lib_UnitTest`'s `pReadLogicalLine` concatenates physical lines that end with whitespace + `_` as continuation lines. If a `Public Sub Test_...` declaration appears immediately after an apostrophe comment line or a `Rem` comment line, it is treated as a continuation of the comment line and is not detected as a test declaration.

This can look like a missed test, but in VBA, line continuation is valid even on comment lines.

## Decision

`pReadLogicalLine` will keep the current behavior of concatenating logical VBA lines, including comment lines.

A `Public Sub Test_...` swallowed by line continuation at the end of a comment line is not a valid independent test declaration. We will not treat this as a bug in test detection logic.

## Consequences

- `Lib_UnitTest` does not special-case comment lines.
- Future test-detection improvements must not introduce syntax interpretation that differs from the VBA runtime.
