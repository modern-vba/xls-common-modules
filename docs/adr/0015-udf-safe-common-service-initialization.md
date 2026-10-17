# ADR 0015: UDF Common Service Initialization Is Separate from Normal Entry Points

- Status: Accepted
- Date: 2026-05-30

## Context

For UDFs executed as Excel worksheet functions, calling initialization through `Application.Run` during cell recalculation can raise `Err.Number = 1004`, `HRESULT = 0x800A03EC`, and `'Run' method of '_Application' object failed`. The existing `InitializeCommonService` attempts to initialize the optional `FileSystemService` and `TextFileService` through `Application.Run`, so it is valid for normal GUI and batch entry points but is not a safe entry point in the UDF recalculation context.

## Decision

Do not make `InitializeCommonService` UDF-safe. UDFs use `InitializeUdfCommonService`. `InitializeUdfCommonService` initializes only `WbSrv` and `WsSrv`, and does not use `Application.Run`, VBIDE, external I/O, or cell changes. Individual initialization APIs, `InitializeWorkbookService` and `InitializeWorksheetService`, are also exposed. Like the existing `InitializeFileSystemService` and `InitializeTextFileService`, they preserve already substituted services when `Force:=False` and regenerate production services when `Force:=True`.

Public UDFs call `InitializeUdfCommonService` at the beginning with `Force` omitted, even when the current body does not directly use `WbSrv` or `WsSrv`. This treats public UDFs as entry points for the cell-recalculation context and keeps a consistent UDF-safe common-service boundary. Because `DIFFSTR` is a public UDF, it follows this standard form.

Normal GUI and batch entry points continue to use `InitializeCommonService`. `InitializeCommonService` remains the initialization entry point that includes optional services, preserving its existing behavior for normal macro entry points.

## Consequences

- UDFs do not call `InitializeCommonService` directly; when needed, they call `InitializeUdfCommonService` or the necessary individual service-initialization API.
- Public UDF entry points call `InitializeUdfCommonService` with `Force` omitted, so uninitialized `WbSrv` and `WsSrv` are created while already substituted services are preserved.
- `InitializeUdfCommonService` does not initialize `FsSrv` or `TfSrv`, so UDFs do not assume file-system or text-file common services.
- Existing normal macro entry points can continue using `InitializeCommonService Force:=True`.
