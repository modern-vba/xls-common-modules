# Source Workbook Tooling

## Maintained workbook and outputs

`CommonModules/vba-project.json` selects
`CommonModules/src/CommonModules/CommonModules.xlsm` through `templatePath`.
This is the maintained source workbook, not a template copied to a development
bin directory. The maintained project omits `binPath`; ordinary Build, Test,
Debug, and project Export neither read nor create a bin-workbook target.
Existing legacy `binPath` settings remain accepted only through the upstream
deprecated compatibility contract and emit a removal warning.

Publish is unchanged: it creates the distributable workbook at
`CommonModules/publish/CommonModules.xlsm`. Explicit snapshot-output Build also
retains its separate output contract; it is not ordinary in-place Build.

## Workbook lifetime and Save policy

| Operation | Exact source workbook already open | Source workbook closed | Tool Save |
| --- | --- | --- | --- |
| Ordinary Build | Reuse its file/process/window, import, save, and leave open | Open hidden, import, save, and close the command-owned session | Save source workbook in place |
| Test | Reuse its file/process/window and leave open with unsaved state | Open hidden, run, and close without saving | None |
| Debug F5/Restart | Reuse its file/process/window and retain it after Stop | Open visibly and retain it for native VBE debugging | None |

Borrowed Build/Test sessions preserve their displayed window and application
settings. A basename, arbitrary active workbook, or another workbook in the
same Excel process cannot substitute for the selected workbook. Unrelated
workbooks and Excel processes are not Save/Close/Quit/termination targets.

Normal Test imports saved disk source; Test Explorer's snapshot profile can
include participating unsaved source-editor content without saving editors.
Normal/snapshot Test retains source admission, analysis, import, and verification.
`--no-build` instead runs current workbook VBA without importing external source.
Closed Test discards remaining unsaved workbook changes, so retain the command's
text or NDJSON output instead of expecting a saved results worksheet.

Test and Debug do not initiate `Workbook.Save`. This does not prevent the VBA
procedure from explicitly saving, changing other state, or causing external
side effects. Neither command promises whole-workbook rollback of macro effects.

## Unsaved changes: Build/Test versus Debug

Direct CLI Build/Test defaults to `--interactive true` and asks for terminal
`[y/N]` consent when the exact already-open workbook has unsaved changes. Build's
warning covers replacing live VBA and saving other workbook edits; Test's warning
covers replacing VBA and running tests, or execution-only consent for `--no-build`.
EOF, an empty response, or a nonaffirmative answer declines; there is no GUI
message window or automatic detection of interactive callers.

For batch operation, pass `--interactive false`. A dirty selected workbook then
fails immediately without waiting, import, execution, or automatic consent.
PowerShell's `-NonInteractive` controls the shell, not the companion's consent
policy. UPDATE always supplies `--interactive false` to its Build calls. Save or
close a dirty source workbook and rerun a refused update; do not treat a failed
Build as successful installation into the workbook.

Debug F5 and Restart never ask for dirty-workbook confirmation. The accepted
Debug request authorizes replacing live VBA with the complete captured source in
the same selected workbook and Excel process/window. This can overwrite unsaved
VBE-direct edits; existing cell edits remain. Debug does not auto-save source
editors or the workbook before preparation, after execution, on Restart, or on
Stop. Its independent syntax/source-error analysis precheck is omitted; native
VBE compile/runtime behavior remains authoritative. This exemption does not apply
to Build or normal/snapshot Test.

Removing Debug consent does not remove immutable source capture, strict encoding,
target/breakpoint identity checks, pre-replacement module/reference recovery
capture, generation-bound readiness, verified import, or bounded cancellation.
An import/preparation failure attempts code/reference recovery without Save;
incomplete recovery reports retained evidence and manual guidance. It is not a
rollback guarantee for cells or effects of VBA that has already run.

`prepare-debug --interactive` remains accepted only for CLI compatibility:
neither true nor false prompts or refuses solely because the workbook is dirty.
Build and Test retain their normal meaning for this option.

## Wrapper executable selection and compatibility

Wrappers select a valid `VBA_DEV_EXE` override first, then the repository bundle
at `tools/vba-dev/vba-dev.exe`, then a PATH provider. They check the actual
selected executable with `capabilities --format json`; `--version` alone cannot
prove source-workbook behavior. An existing chosen incompatible provider is
reported as failure rather than silently retried through a different provider.

Feature versions describe behavior independently of command output schemas:

| Needed behavior | Capability feature and version |
| --- | --- |
| Manifest without legacy bin output | `projectManifest.optionalBinPath: 1.0` |
| Ordinary in-place source Build | `build.sourceWorkbook: 1.0` |
| Source-workbook Test | `test.sourceWorkbook: 1.0` |
| No-confirmation source Debug | `debug.sourceWorkbookPreparation: 2.0` |

Before modifying project sources, UPDATE requires the selected CLI's contract
`1.0`, optional-bin and source-Build features, Build output schema `3.0`, and
CommonModule Update output schema `1.0`. UPDATE does not require Debug because it
does not invoke it. Separate bundle verification additionally proves source-Test
feature `1.0`/output schema `1.2` and current Debug preparation feature `2.0`/output
schema `1.0`. A full companion bundle for this workflow supports both behaviors;
Debug's older preparation feature `1.0` can still prompt and is not compatible
with the no-confirmation requirement. Build schema `3.0`, Test schema `1.2`, and
`prepare-debug` result schema `1.0` remain separate contracts; an unchanged output
schema does not imply the newer source-workbook or no-prompt behavior.

PowerShell/BAT entry points preserve a failed companion invocation as failure;
informational pauses must not replace the captured exit code with success. They
do not auto-approve dirty workbooks or retry failed workbook mutations.

## Unchanged operations

Standalone Import/Export and `IMP_MODS`, `EXP_MODS`, and `SYNC_MODS` keep their
existing closed-target behavior, arguments, and lifetime rules. This source
workflow does not add dirty-session reuse or new confirmation flags to them.
Publish remains unchanged. COLLECT/DIST retain their independent Search Roots,
closed-package metadata selection, ordinary copy, failure scopes, and rerun
policy; no transaction, rollback, or retry framework is introduced.

## Upstream behavior authority

The source-workbook changes follow `modern-vba/vba-tools` issues #447-#450.
Debug's no-confirmation change follows issue #451 and
[ADR 0064](https://github.com/modern-vba/vba-tools/blob/8919d7e46d62b463cd5716ed02e91c2296402c1f/docs/adr/0064-debug-dirty-source-workbooks-without-confirmation.md).
[PR #452](https://github.com/modern-vba/vba-tools/pull/452) was squash merged into
upstream `main` as `8919d7e46d62b463cd5716ed02e91c2296402c1f`, with the same source
tree as the tested candidate `f1baf12201fe6f011cf09ff93bac0966dcc9ca26`.
Upstream integration does not imply a published release, tag, or version.
Use the bundled tool's recorded provenance and actual capability response when
checking this repository's companion artifact.

## Bundled companion provenance

The repository's `tools/vba-dev/vba-dev.exe` was refreshed on 2026-10-10 from the
standalone Windows x64 artifact built by `npm run package:devtool` on the clean
upstream `main` source tree at `8919d7e46d62b463cd5716ed02e91c2296402c1f`.
The archive entry and upstream published executable had identical SHA-256 hashes;
the copied repository executable was checked again after replacement.

- Archive: `vba-dev-win-x64-0.1.0.zip`.
- Archive SHA-256: `BE8306E1ADC217751AEB98EEB99DA05ABF739F4D0587B8958134019CE900294E`.
- Executable SHA-256: `688A2C8A15C6B1E6E4441924E0F0C2D627A0D86E38C329E397C7897B11CAD1A1`.
- Actual capability response: optional-bin, source-Build and source-Test features
  `1.0`; no-confirmation Debug preparation feature `2.0`; Build/Test/preparation
  output schemas `3.0`/`1.2`/`1.0` respectively.

The CLI reports `vba-dev 0.1.0`; this label alone is not the compatibility proof.
This is a repository companion refresh, not a new upstream release or publication.
