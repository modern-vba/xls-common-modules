# CommonModules

CommonModules is the VBA foundation shared by multiple Excel macro workbooks.
Terms for common modules are used in meanings that do not depend on any distribution target tool.

## Wrapper implementation scope

`COLLECT_COMMON_MODS` and `DIST_COMMON_MODS_REPO` remain plain Windows PowerShell workflows. Prefer small, readable improvements to the current scripts; do not introduce managed helper state machines, native lease protocols, transactional recovery frameworks, or exhaustive failure taxonomies unless a later implementation task explicitly requires them. For current behavior, the script and its focused tests take precedence over unimplemented speculative detail elsewhere in this file.

## Language

**CommonModules**:
The whole library shared by multiple Excel VBA tools.
_Avoid_: an individual common module, an individual tool

**Common module**:
A shared VBA source unit listed by the canonical CommonModules manifest. COLLECT chooses its bytes from the newest eligible matching file in project source sets discovered below the explicit `Collection Search Root`; a `common_modules_repo` is generated output rather than a source candidate.
_Avoid_: individual module, tool-side copy

**Common module primary role**:
The exactly one repository classification that places a common module in `runtime-baseline`, `test-foundation`, `optional`, or `test-double`. It determines whether the module belongs to the initial foundation and whether downstream runtime publication treats it as test-only.
_Avoid_: arbitrary category bag, category precedence

**Public UDF modifier**:
The optional `public-udf` classification attached only to a runtime common module primary role. It describes worksheet exposure without changing foundation selection or test-only classification.
_Avoid_: primary role, test-only marker

**Common module dependency**:
A common module that must accompany another common module. A module with a runtime primary role may depend only on runtime-role modules, while a test-role module may depend on either role family.
_Avoid_: optional module, copied file list

**Common module dependency component**:
A maximal set of common modules that are mutually reachable through common module dependencies. Its members remain distinct modules but form one indivisible unit when dependency closure and canonical order are derived.
_Avoid_: invalid cycle, merged common module

**Common module required reference**:
A human-visible external VBA reference name that one common module directly requires downstream projects to select. It excludes the always-active Visual Basic for Applications standard library.
_Avoid_: inferred reference, transitive dependency copy, standard library

**Individual tool**:
An Excel VBA tool that has its own business-specific `.xlsm` and source set, such as a tool under `xls-web-tools`.
_Avoid_: CommonModules, common module

**CommonModulesRepository**:
A generated closed flat package directory named `common_modules_repo`. It contains the canonical manifest, exactly one root-level source unit for every manifest row, and only each listed form's optional matching `.frx`; every other entry is outside the package. COLLECT writes this output, while mandatory baseline and fallback candidates come only from `CommonModules Authoring Source Set`.
_Avoid_: authoring source set, transaction workspace, package cache

**Collection Search Root**:
The one required directory argument supplied to COLLECT. A relative value is resolved against the `Wrapper Repository Parent`. COLLECT recursively discovers `vba-project.json` files without descending into reparse child directories; the explicit root itself may be a reparse path, while neither the script location nor a discovered package repository substitutes for this input.
_Avoid_: output root, script repository, distribution search root

**CommonModules Authoring Source Set**:
The unique discovered project document source set whose root directly contains `common-modules-manifest.tsv`. It establishes the canonical manifest and mandatory fallback sources, but its project manifest does not locate the wrapper's repository; zero or multiple matches establish no collection authority.
_Avoid_: generated common_modules_repo, folder-name convention, arbitrary manifest match

**Wrapper Repository Parent**:
The working directory captured when COLLECT or DIST starts. Its direct `common_modules_repo` child is the COLLECT output or DIST source, independently of either Search Root and any project manifest's `commonModulesRepository` value.
_Avoid_: collection search root, distribution search root, project manifest directory

**Distribution Search Root**:
The workspace directory explicitly supplied to DIST, whose immediate child project directories may contain opted-in `CommonModulesRepository` targets. It is independent of the source repository under the `Wrapper Repository Parent` and is never inferred from that repository's path.
_Avoid_: source repository, source parent, distribution target

**Distribution Target**:
An existing exact `common_modules_repo` beneath an immediate child project directory of the `Distribution Search Root`, excluding the central source repository. Project manifests, repository names, nested directories, and missing repositories do not create distribution authority.
_Avoid_: inferred target, nested project, newly created repository

**Distribution Candidate Set**:
The path-ordered target set fixed by one initial scan of the `Distribution Search Root`. A target appearing later waits for another invocation; a selected target that disappears before its turn becomes a `Distribution Warning Skip` without widening or rebuilding the set.
_Avoid_: live target query, retry set, recursive discovery

**Distribution Package Match**:
Two closed flat `CommonModulesRepository` inventories with the same ordinal-exact root-level names and the same per-file `LastWriteTimeUtc` and length values. DIST does not read file bytes or hashes to establish this match.
_Avoid_: content equality, directory timestamp, hash comparison

**Distribution Candidate Failure**:
A conclusive defect or operation failure isolated to one distribution candidate. The candidate is not retried or rolled back, later candidates remain eligible, and the DIST invocation ultimately fails.
_Avoid_: warning skip, global failure, automatic recovery

**Distribution Warning Skip**:
A candidate whose opt-in state cannot be established before mutation, or which disappears before its turn begins. DIST leaves it untouched, warns, and continues; this classification alone does not fail the invocation.
_Avoid_: candidate failure, silent exclusion, retry

**Distribution Global Failure**:
An invalid or unreadable central source, `Distribution Search Root`, or complete discovery operation. It prevents a trustworthy candidate set and stops the invocation with failure.
_Avoid_: candidate failure, warning skip

**Document source set**:
The Git-managed VBA source and template workbook belonging to one individual tool document in a vba-dev project.
_Avoid_: generated workbook output, common_modules_repo, workbook-local source set

**Workbook-local source set**:
A project-independent `modules` folder next to a workbook, used when extracting from or importing into that workbook directly.
_Avoid_: document source set, `vba-dev` `sourcePath`

**Individual module**:
Tool-specific VBA source in an individual tool's source set. This does not include copies of common modules.
_Avoid_: common module, whole individual tool

**VBA source**:
Git-managed `.bas` / `.cls` files.
_Avoid_: unexported code inside `.xlsm`, Excel workbook

**Original workbook file**:
An existing Excel workbook file identified as the workbook file itself, not a VBA source file or a generated copy.
_Avoid_: source file, template copy

**Reflect**:
The operation of writing a document source set into a generated workbook used for development or test execution.
_Avoid_: extract, distribute, publish

**Extract**:
The operation of writing VBA source from a workbook into a source set.
_Avoid_: reflect, distribute

**Workbook import**:
The operation of writing VBA source from a source set into an existing workbook in place.
_Avoid_: build, reflect, publish

**Test diagnostic information**:
Information checked to identify causes when a unit test fails. This includes expected and actual values, case names, raised errors, and test-double call history.
_Avoid_: test result itself

**Call arguments**:
For a test-double call that completed successfully, all arguments passed to the actual method, arranged in the original function definition order. These are taken from call history to verify the actual operation that was issued.
_Avoid_: recorded arguments

**Match arguments**:
Key arguments used with the method name to match a test double's return value, output value, error, or call history. These are not necessarily all arguments of the actual method and are treated separately from call arguments.
_Avoid_: call arguments, recorded arguments

**Input sheet**:
A worksheet where users enter processing conditions and settings. This refers to the whole sheet, not only the rectangular range searched by logic.
_Avoid_: input area

**Input area**:
A rectangular range inside an input sheet that is treated as the target for reading item names and values. This refers to the area containing the user-input table, not the entire sheet.
_Avoid_: whole input sheet

**Virtual table**:
A logical table that associates multiple columns read from an input area by common headers and relative rows. It is not the actual rectangular range on the worksheet.
_Avoid_: input table, real table, input area

**Button**:
A worksheet shape responsible for running a macro through user operation. This does not refer only to UnitTest rerun buttons or shapes created by `AddButton`.
_Avoid_: UnitTest button, AddButton-created button

**Common service**:
The shared access foundation that CommonModules provides to each tool. This refers to Excel workbooks, worksheets, the file system, and text files; it does not include debug information or progress display.
_Avoid_: common run state

**Common run state**:
Debug information and progress-display state held only during one GUI run. This is not a common service; it is state tied to a run scope.
_Avoid_: common service

**UDF-safe**:
A property of not requiring side effects that are inappropriate for a recalculation context when called as an Excel worksheet function during cell recalculation. Treat this separately from being safe to use from GUI or batch entrypoints.
_Avoid_: safety during macro execution

**Public UDF**:
A standard-module function that CommonModules intentionally exposes for use from Excel worksheet formulas. This refers only to functions intended for use as cell formulas, not all `Public Function` members callable from VBA.
_Avoid_: all public standard functions, public class methods

**Public API**:
The callable surface CommonModules exposes externally as contracts with distribution target tools and test doubles. Separate from public UDFs, this includes public members of services, value objects, and collections used from VBA, plus their interface and test-double contracts.
_Avoid_: public UDF, internal helper, names used only in comments

**Range shape**:
The dimensions represented by `WorksheetRangeBounds`, consisting of row count and column count. This refers to a size that can be handled with the same relative row and column numbers, not absolute start or end coordinates.
_Avoid_: end position, absolute coordinates

**Typed element collection**:
A collection that has an element type contract and holds only elements that follow the same contract. Elements include primitive values, arrays, Excel error values, object references, and Nothing.
_Avoid_: object-only collection, value collection

**Keyed typed element collection**:
A typed element collection whose elements can be referenced by key. It treats the key contract and value element type contract separately, and applies the same element type contract as typed element collections to values.
_Avoid_: unkeyed typed element collection, object-only dictionary, arbitrary-type dictionary

**Key contract**:
The range of keys accepted by a keyed typed element collection and the rules for considering keys equal. Treat this separately from the value element type contract.
_Avoid_: value element type contract, element capability contract

**Key comparison mode**:
The mode in a keyed typed element collection that determines whether string-key matching is case-sensitive. Treat this separately from the element type contract.
_Avoid_: element type contract, identity/duplicate detection mode

**Explicit element type contract**:
A contract where a typed element collection has the target element type before waiting for the first element to be added.
_Avoid_: implicit type inference only

**Assignable type acceptance**:
The behavior where a typed element collection accepts an element based on assignability to the specified element type, rather than exact concrete class name.
_Avoid_: exact concrete type name match only

**Element type self-reporting**:
The behavior where an object stored in a typed element collection returns which element type contract it should be treated as. This is used when CommonModules does not directly know a caller-specific interface type.
_Avoid_: type contract based only on concrete class name

**Element capability contract**:
A contract in a typed element collection requiring elements with the same element type contract to have the same capabilities for comparison, identity, and duplicate detection. This does not include display or string-conversion capability.
_Avoid_: display capability contract, string-conversion capability contract

**Required capability**:
Behavior a typed element collection requires from elements stored in the same collection. Behavior that is not required is not included in the condition for accepting an element.
_Avoid_: optional capability, display string conversion

**Identity/duplicate detection mode**:
The exclusive criterion a typed element collection uses to consider elements the same or duplicate. Treat this separately from ordering requirements.
_Avoid_: ordering, display string conversion

## Example Dialogue

Developer: Does `ClearButton` remove only the UnitTest rerun button?
Domain expert: No. A button is a worksheet shape responsible for running a macro; it is not UnitTest-specific.

Developer: May a repository reader normalize whitespace, casing, empty items, or order in `Categories` and `Dependencies`?
Domain expert: No. A category cell is exactly one of the six canonical whole-cell spellings. A dependency cell is empty or a whitespace-free comma-separated sequence that exactly matches target-row `ModuleFile` spellings; preserve its order and reject rather than normalize empty items, casing differences, duplicates, or self-dependencies.

Developer: May comments or blank lines appear between manifest data rows?
Domain expert: No. Allow only a contiguous whole-line comment prologue whose `#` begins in column one, then the exact header, one or more contiguous data rows, and the final CRLF. Blank or whitespace-only lines, comments after the header, indented or inline comments, and duplicate headers are invalid rather than ignored.

Developer: May `ModuleFile` point into a subdirectory and be flattened when distributed?
Domain expert: No. `ModuleFile` is the exact flat basename `<common module name>.bas`, `.cls`, or `.frm`, never a relative path. COLLECT finds every exact ordinary `vba-project.json` under the explicit `Collection Search Root`, resolves each document `sourcePath`, searches those source sets recursively for one case-insensitive basename match, and selects the greatest `LastWriteTimeUtc`. At equal newest timestamps it compares only `Length`: equal lengths are treated as equivalent, preferring the CommonModules candidate when it is tied and otherwise using ordinal path order; different lengths use the CommonModules fallback. Collection writes the selected source unit with the manifest's exact casing at the package root, so source subdirectories do not change the flat output layout.

Developer: How does COLLECT distinguish the canonical manifest from manifests in generated repositories?
Domain expert: The canonical manifest is the one directly contained by the unique `CommonModules Authoring Source Set` discovered through a project document `sourcePath`. COLLECT writes to `common_modules_repo` directly under the `Wrapper Repository Parent`; neither the containing project's `commonModulesRepository` nor a generated repository manifest establishes wrapper output authority.

Developer: Does a `tests` directory decide which authoring source may be absent from the distribution manifest?
Domain expert: No. Every file absent from the canonical manifest remains authoring-only regardless of directory and is not copied. Multiple project-source files matching one listed basename are normal candidates and use only their `LastWriteTimeUtc` and `Length` under the collection selection rule. The distributed repository remains a flat manifest-listed package.

Developer: Does successful collection certify the complete CommonModules authoring project?
Domain expert: No. It certifies package selection and the exact manifest-listed package inputs. All unlisted source remains authoring-only and is not a collection health target; `vba-dev build`, `vba-dev test`, and Doctor own its source validity.

Developer: May COLLECT retain an old file in `common_modules_repo`, or may DIST distribute an extra entry not named by the manifest?
Domain expert: No. A `CommonModulesRepository` is a closed flat package. COLLECT removes obsolete package entries, and DIST treats a missing listed entry, an unexpected entry, or a non-flat inventory as a `Distribution Global Failure` before touching any target.

Developer: May collection combine manifest-listed files read before and after an authoring edit, or retry silently onto a newer generation?
Domain expert: The current PowerShell implementation performs one ordinary discovery and newest-file selection pass before it starts copying. It does not provide a cross-file atomic snapshot or retry while files are being edited, so do not edit project sources during COLLECT; rerun COLLECT after an interrupted or concurrent edit.

Developer: What determines the scope of distribution?
Domain expert: The caller explicitly supplies the `Distribution Search Root`; DIST never infers it from the source repository. DIST reads its source from `common_modules_repo` directly under the `Wrapper Repository Parent`, while only existing repositories in immediate child project directories are opted-in targets.

Developer: May COLLECT create a missing source repository, or may DIST search for one elsewhere?
Domain expert: COLLECT may create the missing `common_modules_repo` under the `Wrapper Repository Parent`, but only after every preflight succeeds. DIST requires that same path to contain its existing source repository; if it is missing, DIST reports a `Distribution Global Failure` with exit `1` and neither searches for nor creates a substitute.

Developer: What is the public argument surface of the COLLECT and DIST wrappers?
Domain expert: Each `.ps1` and `.BAT` accepts exactly one first argument: its explicit Collection or Distribution Search Root. Resolve a relative argument against the `Wrapper Repository Parent`. A missing, empty, absent, non-directory, or unreadable argument is a global error with exit `1`; the explicit root itself may be a reparse path. Neither command infers its Search Root from the working directory, script location, source path, or output path, and neither accepts source or output overrides.

Developer: How do the BAT wrappers preserve the PowerShell result?
Domain expert: Invoke PowerShell with `-NoProfile`, capture `%ERRORLEVEL%` immediately after the child returns, and finish with `EXIT /B` using that captured value even when informational UI or `TIMEOUT` follows. Success and warning-only completion return `0`; every COLLECT global or copy failure and every DIST global or candidate failure return `1`, and the BAT layer never rewrites failure as success.

Developer: How should focused tests cover the simplified COLLECT and DIST wrappers?
Domain expert: Use temporary-directory PowerShell fixtures for required Search Roots and exit codes, COLLECT discovery, exclusions, manifest validation, newest-file selection and CommonModules fallback, closed-package update, DIST opt-in, warning skip, candidate continuation, and BAT exit preservation. An ordinary `UNCHANGED` fixture may use identical contents, timestamps, and lengths, but tests do not pin a deliberately different-content file with equal time and length. Do not add Excel, native API, lease, transaction, rollback, or state-machine test machinery.

Developer: How does one immediate child project opt in to distribution?
Domain expert: Zero case-insensitive `common_modules_repo` matches means no opt-in. Exactly one ordinal-exact ordinary directory is a `Distribution Target`; a case-only match, multiple matches, or a wrong entry type is a `Distribution Candidate Failure`, not absence or a target to repair.

Developer: May the caller use the fixed central repository to distribute into an unrelated directory tree?
Domain expert: Yes. The source repository under the `Wrapper Repository Parent` and the explicit `Distribution Search Root` are independent authorities. DIST excludes the source itself if it appears in that scope, but does not require its parent to be one of the root's children.

Developer: Must the `Distribution Search Root` be a physical non-reparse directory?
Domain expert: No. An explicitly supplied junction or symbolic-link path is a valid search authority when it resolves and can be enumerated. DIST does not promise physical-identity continuity if that link is retargeted during the invocation; repository roots and package contents retain their separate no-reparse rule.

Developer: Does DIST descend into an immediate child project that is a junction or symbolic link?
Domain expert: No. As with recursive COLLECT discovery, the explicitly supplied Search Root may itself be an alias, but a reparse child encountered during discovery is a `Distribution Warning Skip`. Pass that link itself as another explicit Search Root when its target should be processed.

Developer: May source files, targets, or reparse links be changed concurrently while DIST runs?
Domain expert: No continuity guarantee is provided for concurrent changes. Run DIST against a stable source and search tree; an alias of the central source ordinarily remains `UNCHANGED` through its matching names, lengths, and timestamps, but link retargeting or filesystem edits belong to a later invocation.

Developer: Does recursive COLLECT discovery follow a junction or symbolic-link child found below its Search Root?
Domain expert: No. The explicit `Collection Search Root` itself may be an alias, but discovery does not descend into a reparse child. To include the linked tree, invoke COLLECT separately with that link as the explicit root.

Developer: Does COLLECT search generated or administrative directory trees for projects?
Domain expert: No. Discovery excludes `.backups`, `.git`, `.out-of-scope`, `.tmp`, `.venv`, `.vs`, `.vscode-test`, `artifacts`, `bin`, `common_modules_repo`, `node_modules`, `obj`, `out`, `packages`, `publish`, `temp`, `TestResults`, and their descendants. `TestResults` is ordinary VSTest output, not a project source location.

Developer: May COLLECT warning-skip a discovered project manifest or source set that it cannot read or validate?
Domain expert: No. An unreadable or invalid `vba-project.json`, an invalid or unresolvable document `sourcePath`, or an absent or unreadable resolved source directory is a collection-global failure before output mutation. Skipping one could hide the newest module candidate and produce a stale central package.

Developer: May COLLECT follow a reparse-point `vba-project.json` or canonical `common-modules-manifest.tsv`?
Domain expert: No. Both authority files must be ordinary non-reparse files. Discovering either canonical basename as a symbolic link or another reparse file is a collection-global failure before output mutation, not a warning skip or authority to follow the link.

Developer: May authority-file discovery silently ignore case-only basename differences?
Domain expert: No. Require the ordinal-exact basenames `vba-project.json` and `common-modules-manifest.tsv`, while using a case-insensitive directory check to detect case-only or multiple matches. Either defect is a collection-global failure rather than an absent project or manifest.

Developer: May one document source set contain several files with the same common-module basename?
Domain expert: No. A case-insensitive basename duplicate within one resolved `sourcePath` is an ambiguous source-set identity and makes collection fail before output mutation. Equal basenames in different source sets are normal cross-project candidates and remain eligible for newest-`LastWriteTimeUtc` selection.

Developer: Must a module candidate's basename casing exactly match the manifest `ModuleFile`?
Domain expert: No. Accept exactly one case-insensitive basename match per source set and write the selected unit using the manifest's ordinal-exact `ModuleFile` casing. A second case variant in the same source set remains the prohibited duplicate rather than another candidate.

Developer: How does COLLECT select a manifest-listed form that has an `.frx` sidecar?
Domain expert: Treat the `.frm` and its optional same-directory, same-basename `.frx` as one candidate. Its selection time is the later `LastWriteTimeUtc` of the pair, and the winning pair is copied together; an orphan `.frx` is not a candidate. At equal newest times, equal sidecar presence and corresponding `.frm` and `.frx` lengths are treated as equivalent; any difference uses the CommonModules fallback. No form content or hash is read.

Developer: Does COLLECT need staging, a transaction, or rollback for the central package?
Domain expert: No. Complete manifest validation, project and source-set discovery, candidate selection, and tie validation before changing the central repository. If the package is not `UNCHANGED`, clear every existing repository entry, then copy the canonical manifest and every selected source unit and `.frx` sidecar sequentially with ordinary PowerShell. A deletion or copy error returns exit `1` and may leave a partial package, without differential patching, staging, retry, transaction, or rollback.

Developer: Does COLLECT lock or recheck source files while materializing the selected package?
Domain expert: No. Discover projects and read candidate `LastWriteTimeUtc` and `Length` once, then use ordinary sequential PowerShell copy without a source lock, snapshot, post-copy verification, or concurrency guarantee. A selected source that disappears or causes copy failure returns exit `1` and may leave a partial package; a normal copy return succeeds, and later edits belong to a later invocation against a stable tree.

Developer: Does COLLECT compare file contents before updating the central package?
Domain expert: No. Candidate selection and package matching use only exact ordinal filenames, `LastWriteTimeUtc`, and `Length`; they never read file contents or hashes. If the central package has the selected package's exact filename inventory and every corresponding time and length, report `UNCHANGED` and perform no write; otherwise update it.

Developer: Must a document source set contain the canonical CommonModules manifest to participate in collection?
Domain expert: No. Every document `sourcePath` resolved from a discovered valid `vba-project.json` is a candidate source set. Exactly one of those source-set roots must directly contain `common-modules-manifest.tsv`; that source set is `CommonModules Authoring Source Set`, while zero or multiple manifest-owning source sets make collection fail before output mutation.

Developer: How does COLLECT resolve and deduplicate document `sourcePath` values?
Domain expert: Resolve each relative value against the directory containing its `vba-project.json`, normalize it to an absolute full path, and compare those paths with Windows case-insensitive semantics. Scan equal normalized paths once even when several documents reference them. Do not add physical-file identity, native handles, or alias resolution merely to deduplicate source sets.

Developer: May an explicitly resolved document `sourcePath` itself be a junction or symbolic link?
Domain expert: Yes. Treat that manifest-resolved root like the explicit Collection Search Root and scan it, but do not descend into a reparse child directory encountered below it. Do not resolve physical alias identity, and require a stable link target for the duration of the invocation.

Developer: May COLLECT use a reparse-point `.bas`, `.cls`, `.frm`, or `.frx` as a module candidate?
Domain expert: No. Source units and form sidecars must be ordinary non-reparse files. A required CommonModules source unit or sidecar that is a reparse point is a collection-global error; a matching non-CommonModules reparse candidate produces a warning and uses the CommonModules fallback without following the link.

Developer: What happens when ordinary newest-candidate selection cannot be completed?
Domain expert: Every file listed by `common-modules-manifest.tsv` is mandatory in `CommonModules Authoring Source Set`: a missing file or unavailable `LastWriteTimeUtc` or `Length` makes collection fail before output mutation. A missing candidate in another project is ordinary. If another project's `LastWriteTimeUtc` or `Length` cannot be read, or equal newest timestamps have different lengths, select the CommonModules candidate instead of failing and ignore the uncertain non-CommonModules candidates. Equal newest timestamps and lengths are treated as equivalent, preferring CommonModules when it is tied and otherwise using ordinal path order.

Developer: Does using the CommonModules fallback produce a warning or an error?
Domain expert: A missing candidate in another project produces no warning. An unreadable `LastWriteTimeUtc` or `Length` in another project, or equal newest timestamps with different lengths, produces a warning naming the affected module and then uses the CommonModules candidate; fallback warnings alone retain exit `0`. A missing CommonModules candidate or unreadable required CommonModules attribute is a collection-global error with exit `1`.

Developer: Where does COLLECT create or update its `CommonModulesRepository`?
Domain expert: Capture the invocation-start working directory once as the `Wrapper Repository Parent` and use its direct `common_modules_repo` child. Do not consult any project manifest's `commonModulesRepository` value or retarget the output after an internal location change. A missing path may be created as an ordinary directory only after every preflight succeeds. An existing path must be an ordinary non-reparse directory whose entries, `LastWriteTimeUtc` values, and lengths are readable; a file, reparse point, enumeration failure, or unavailable required attribute is a collection-global failure before output mutation. The CommonModules project source set, not the generated repository, supplies the mandatory baseline and fallback candidates.

Developer: Does DIST keep discovering targets while distribution is running?
Domain expert: No. It fixes one path-ordered `Distribution Candidate Set` at invocation start. A later target waits for the next invocation, while a selected target that disappears before mutation is warning-skipped without rediscovery, retry, or retargeting.

Developer: Must DIST compare file content before deciding that a target is unchanged?
Domain expert: No. A `Distribution Package Match` uses exact flat inventory names and each corresponding file's exact `LastWriteTimeUtc` and length. Matching metadata is accepted without reading bytes or hashes; any inventory, timestamp, or length difference makes the target eligible for replacement.

Developer: Must DIST reread copied files to prove their timestamp and length after `Copy-Item` returns normally?
Domain expert: No. A normal copy return completes that candidate without an additional verification pass. A later DIST invocation applies `Distribution Package Match` again and replaces the target if its metadata does not match.

Developer: Does distribution require a native lease, managed state machine, transactional workspace, rollback protocol, or custom cancellation channel?
Domain expert: No. Keep the wrapper as a plain Windows PowerShell file-copy workflow. A `Distribution Candidate Failure` leaves that candidate unchanged or possibly partial, records final failure, and continues to later candidates; a `Distribution Global Failure` stops the invocation. Correct the reported problem and rerun rather than attempting automatic rollback or recovery.

Developer: If an I/O error occurs while DIST is copying one target, does it always have the same scope?
Domain expert: No. Failure to enumerate or read the central source is a `Distribution Global Failure` because no later copy can trust that package. Failure to remove or write one target is a `Distribution Candidate Failure`, so later targets remain eligible.

Developer: May DIST follow a junction, symbolic link, or other reparse point while inspecting or replacing a repository?
Domain expert: No. A reparse point at the central source root or in its package inventory is a `Distribution Global Failure`. A reparse point at a target root or anywhere below it is a `Distribution Candidate Failure`; DIST leaves that target untouched and continues without following the entry.

Developer: Is every access failure a global distribution failure?
Domain expert: No. Failure to enumerate the `Distribution Search Root` is global. Inability to determine whether one immediate child project opted in is a `Distribution Warning Skip`; once an exact target is admitted, inability to enumerate or compare its contents is a `Distribution Candidate Failure`.

Developer: Does DIST need result rows, counts, summaries, or a JSON result envelope to report these classifications?
Domain expert: No. Report updated and unchanged targets as ordinary messages, `Distribution Warning Skip` as a warning, and candidate or global failures as errors. Warnings alone retain exit `0`; any failure produces exit `1` without adding another result protocol.

Developer: When is an invocation with no updated or unchanged target a zero-target error?
Domain expert: It is a zero-target error only when initial discovery finds neither an eligible target nor a known candidate failure. Known candidate failures already make the invocation fail without another zero-target error; a target admitted initially and warning-skipped after disappearing does not retroactively create one. Warning-only uncertainty that admits no target still accompanies the zero-target error.

Developer: What focused behavior must the distribution tests preserve?
Domain expert: A nonempty source is copied to every eligible immediate-child target, stale target contents are removed, the source itself is not modified, and missing target directories are not created. Tests also distinguish candidate-local failure, pre-mutation warning skip, global failure, and the initial zero-target error without adding transaction or rollback behavior.

Developer: Does a case-only repository rename create a new common module or leave installed spelling unchanged?
Domain expert: Neither. It is the same case-insensitive common module name with refreshed canonical spelling. Repository-backed reconciliation adopts that spelling for the installed manifest, source basename, optional form sidecar, and source bytes while preserving the containing directory and ordinary installation intent.
