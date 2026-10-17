# ADR 0011: Treat ApplicationScreenUpdateManager as a One-Shot Scope Manager

- Status: Accepted
- Date: 2026-05-25

## Context

`ApplicationScreenUpdateManager` is used to temporarily change Excel's `ScreenUpdating`, `EnableEvents`, `DisplayAlerts`, and `Calculation`, then restore them after processing. The existing implementation saved settings when an instance was created and discarded the saved state after `Restore`, but reuse, nesting, and `StatusBar` handling were not explicitly documented.

## Decision

Treat `ApplicationScreenUpdateManager` as a one-shot scope-management object. It saves settings at instance creation time, and `Restore` returns to the creation-time state even if `DisableUpdates` was never called. Reusing the same instance after `Restore` or restoring twice is an error.

Nesting is allowed, but the caller's contract is to call `Restore` in the reverse LIFO order of creation. Ordering mistakes are not detected.

`Application.StatusBar` is the responsibility of `ProgressStatus`; `ApplicationScreenUpdateManager` does not save or restore it. `Class_Terminate` tries to call `Restore` as a safety net if restoration was not performed, but restoration errors during destruction are swallowed.

## Consequences

- Create a new `ApplicationScreenUpdateManager` instance for each usage scope.
- If an instance has already been restored, create a new instance instead of reusing it.
- If the `StatusBar` restoration contract changes, consider it separately as a `ProgressStatus` responsibility.
