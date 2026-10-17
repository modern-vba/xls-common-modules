# ADR 0013: Common Service Initialization Does Not Use VBIDE Optional-Service Detection

- Status: Accepted
- Date: 2026-05-28

## Context

`InitializeCommonService` treats `FileSystemService` and `TextFileService` as optional services and calls them only when their initialization procedures exist in the same workbook. After adding the `Force` argument, it becomes useful to distinguish strictly between optional services that have not been imported and a mixed state where old and new APIs coexist. However, detecting procedures through VBIDE requires the user's Excel environment to allow access to the VBA project object model.

Developers who run unit tests are expected to make VBIDE available, but ordinary tool users do not enable that setting. Common service initialization also runs from normal tool entry points, so depending on VBIDE would make the startup requirements for business macros unnecessarily strict.

## Decision

Common service initialization does not use VBIDE to detect optional-service presence. `InitializeCommonService` calls `InitializeFileSystemService` and `InitializeTextFileService` through `Application.Run`, passing `Force`; only when the initialization procedure does not exist does it ignore the optional service as before.

Common modules do not absorb mixed old/new API states at runtime through VBIDE detection. Common modules are expected to be synchronized as a set, and synchronization drift is detected and resolved through the import/export, distribution, and unit-test workflow.

## Consequences

- Ordinary users can run common service initialization without VBIDE access permission.
- In workbooks where `InitializeFileSystemService` or `InitializeTextFileService` has not been imported, the optional service is ignored as before, even with `Force:=True`.
- Initialization does not give special treatment to synchronization drift where only part of the common modules remains on an old API.
