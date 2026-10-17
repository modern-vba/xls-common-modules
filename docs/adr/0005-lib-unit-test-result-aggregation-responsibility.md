# ADR 0005: Lib_UnitTest Delegates Overall Result Judgment to External Verification Steps

- Status: Accepted
- Date: 2026-05-23

## Context

`UnitTestMain` writes assertion failures as `NG`, runtime errors, runner errors, and missing assertions as `ERR` to `UNIT_TEST_SHEET`. It does not return an overall result to the caller and does not return counts other than `OK`.

If an external automated verification step only checks `NG`, tests with `ERR` can remain and still be treated as success.

## Decision

Treat `Lib_UnitTest` as a runner that writes `OK` / `NG` / `ERR` to the result sheet.

Overall judgment is the responsibility of the external verification step. We will not change `Lib_UnitTest` to aggregate `ERR` and return it to the caller. Automated verification steps must treat anything other than `OK` as failure.

## Consequences

- The `UnitTestMain` API contract does not change.
- Verification steps run through Excel COM read result rows from the result sheet and treat anything other than `OK` as failure.
- The distinction between `NG` and `ERR` remains on the result sheet and is used to inspect failure types.
