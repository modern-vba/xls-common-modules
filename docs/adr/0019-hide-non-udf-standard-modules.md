# ADR 0019: Hide Non-UDF Standard Modules with Option Private Module

- Status: Accepted
- Date: 2026-06-03

## Context

Standard modules in `CommonModules` contain many `Public Function` procedures intended for use from VBA. These functions intentionally return values as public APIs, but they also appear in Excel's formula suggestions when a user types `=` in a cell, where they are mixed with public UDFs that are intended to be used as worksheet functions.

Converting non-UDF `Function` procedures to `Sub` plus `ByRef` only to avoid formula suggestions would harm caller readability, composability as expressions, and existing public API contracts. `CommonModules` assumes its modules are reflected into each destination workbook and used from the same VBA project; it does not define a contract for referencing standard modules in `CommonModules` from another VBA project.

## Decision

Public UDFs are placed in `Fx_*.bas`, and `Fx_*.bas` does not use `Option Private Module`. For now, public UDFs are collected in `Fx_Common.bas`; if they grow enough to separate responsibilities, `Fx_...` modules will be split.

Standard modules other than `Fx_*.bas` and `Test_*.bas` use `Option Private Module`. This keeps non-UDF `Public Function` procedures as VBA public APIs while making them less likely to appear in Excel formula suggestions.

Non-UDF return-value APIs are not changed to `Sub` plus `ByRef` solely to address formula suggestions. Because `DIFFSTR` is a public UDF, it moves from `Lib_Common.bas` to `Fx_Common.bas` and delegates the actual processing to the existing `DiffStringArray`.

## Consequences

- Non-UDF public APIs such as `JoinPath`, `StartsWith`, and IPv4 conversion functions remain `Function` procedures.
- `Fx_Common.bas` is included in distribution, and `DIFFSTR` remains available from cell formulas as before.
- Because of `Option Private Module`, non-UDF standard modules are not treated as reference targets from other VBA projects.
- When common modules are reflected to destinations, additions of `Fx_Common.bas` and `Option Private Module` are expanded into `common_modules_repo` and each tool-side `modules` directory.
- The same policy must also be applied to tool-specific standard modules.
