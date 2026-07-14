# ADR 0014: Common Services Are Explicitly Initialized at Entry Points

- Status: Accepted
- Date: 2026-05-28

## Context

Common services (`WbSrv`, `WsSrv`, `FsSrv`, `TfSrv`) are the foundation that lets shared processing across multiple workbooks substitute workbook, worksheet, file-system, and text-file operations. Previously, some public APIs in common modules implicitly called initialization procedures such as `InitializeCommonService` or `InitializeFileSystemService` immediately before using the services.

This implicit initialization is convenient for standalone use, but it makes the caller's initialization responsibility hard to see. In tests especially, initialization inside an API can overwrite test doubles arranged by the test, and the boundary becomes unclear between processing that assumes initialized services and processing that takes responsibility for initialization.

## Decision

Public APIs in common modules that use common services assume the required common services have already been initialized before the call. Inside common modules, except for the service-initialization APIs themselves, common-service use does not implicitly call `InitializeCommonService`, `InitializeFileSystemService`, or `InitializeTextFileService`.

Any remaining implicit initialization calls inside existing common modules are also removal targets. Explicit initialization responsibility at the entry point and in test Arrange code takes priority over convenience for standalone use.

Debug information and progress display are not common services. They are treated as common run state associated with one GUI execution. Common run state is scoped by `CommonRunStateManager` from start to finish.

The public `DbgInfo` and `ProgStat` variables remain in `Lib_Common` as the common run-state entry points that existing code can reference directly. However, their creation and disposal are centralized in the scope manager and are not declared in each tool's `GUIHandler`.

`CommonRunStateManager` does not call `InitializeCommonService`; it only creates and destroys `DbgInfo` and `ProgStat`. Common service initialization is kept separate as the responsibility of GUI entry points or test Arrange code.

`CommonRunStateManager` has an explicit termination method, `Clear`. Normally `Class_Terminate` performs cleanup when the manager has not been ended, but callers can explicitly call `Clear` when they need to control termination order, such as after showing an error.

`Clear` is idempotent and does nothing on a second call. Safe termination takes priority even when both error handling and `Class_Terminate` call it.

When `CommonRunStateManager` initializes, it always creates new `DbgInfo` and `ProgStat` instances and overwrites any existing values. This prevents state left by an abnormal previous run from carrying over to the next run.

Using tools explicitly call `InitializeCommonService` at entry points such as GUI entry points and batch entry points. `GUIHandler` entry points use `InitializeCommonService Force:=True` so each execution restores common services to production instances. Testability of GUI entry points is preserved by keeping complex logic out of `GUIHandler` and delegating the actual processing to `Mod_...` modules or classes. Unit tests explicitly initialize required services or substitute test doubles in Arrange code. Tests that directly use only one service are also responsible for initializing that service in Arrange code.

## Consequences

- Callers are responsible for initialization even when they call common-module public APIs standalone.
- Callers that depended on existing implicit initialization initialize the required common services at the entry point or in test Arrange code.
- The chance that implicit initialization inside common modules overwrites substituted test doubles is reduced.
- Debug information and progress display for GUI entry points are scoped as common run state rather than common services.
- Entry points that need specific termination order explicitly call `CommonRunStateManager.Clear`.
- The reference shape for `DbgInfo` and `ProgStat` is preserved, while their declaration source is centralized in `Lib_Common`.
- Existing consuming code must verify, as a distribution impact, that `InitializeCommonService` is called at entry points or that required services are already initialized before individual service use.
