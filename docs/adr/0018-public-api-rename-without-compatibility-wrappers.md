# ADR 0018: Public API Renames Do Not Keep Old-Name Compatibility Wrappers

- Status: Accepted
- Date: 2026-05-31

## Context

`CommonModules` still contains public API names that are hard to read, such as the spelling of `Unsign`-style names, the word order of `ExistsWorkbook`-style names, and retrieval APIs that return arrays while using singular names. Because these APIs are distributed to `xls-bfw-tools` and `xls-ces-timesheet`, adding compatibility wrappers for old names would make migration easier, but it would also leave old and new APIs coexisting on the shared foundation's public surface for the long term.

## Decision

When public API names are cleaned up, do not keep compatibility wrappers for the old names. The `CommonModules` source, `CommonModules.xlsm`, and distributable `common_modules_repo` are updated in the CommonModules issue. Migration of callers in destination tool-specific code is handled in migration issues for each destination repository.

## Consequences

- The public surface of `CommonModules` does not retain duplicate APIs, and future users can reference only the new names.
- `xls-bfw-tools` and `xls-ces-timesheet` must migrate old API calls to the new names after the CommonModules update.
- Until destination migration is complete, tool-specific code that still contains old API calls may fail compilation or tests.
- If a temporary compatibility wrapper for an old name becomes necessary, this ADR must be explicitly revisited.
