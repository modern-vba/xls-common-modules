# Tool launchers

The seven BAT/PowerShell pairs use the shared script template. Run a BAT normally
for a 60-second informational wait, or place `/nowait` first to skip the wait:

```powershell
.\EXP_MODS.BAT /nowait .\Book.xlsm
.\UPDATE_COMMON_MODS.BAT /nowait .\MyRepository
```

| Launcher | Target paths | Default when omitted |
| --- | --- | --- |
| COLLECT_COMMON_MODS | Exactly one Collection Search Root directory | Error |
| DIST_COMMON_MODS_REPO | Exactly one Distribution Search Root directory | Error |
| UPDATE_COMMON_MODS | Zero or one repository/search directory | Invocation working directory |
| SYNC_MODS | Zero or one directory or sync.json file | sync.json in the invocation working directory, then the BAT directory |
| EXP_MODS | Exactly one workbook file | Error |
| IMP_MODS | Exactly one workbook file | Error |
| GEN_DOC | Exactly one VBA source directory | Error |

BAT files forward their own absolute path through `-BatPath` and forward all
user arguments unchanged. PowerShell consumes a leading `/nowait`, checks the
number of target paths, and resolves them through CommonTools. Relative paths
use the invocation working directory; `%VAR%` environment variables are expanded.
File existence and required file/directory types remain tool-specific checks.
BAT files preserve the PowerShell exit code, including after the wait.

PowerShell scripts also support direct invocation:

```powershell
.\exp_mods_main.ps1 -Path .\Book.xlsm
.\sync_mods_main.ps1 -Path .\sync.json
```

The old six positional arguments containing pre-split BAT and target paths are
no longer supported. `-BatPath` is optional metadata for direct invocation.
The scripts leave the caller's current directory unchanged.

COLLECT and DIST still require one explicit Search Root. Their central
`common_modules_repo` remains directly under the invocation working directory,
independently of the Search Root and BAT/PowerShell locations. Launcher metadata
does not change collection, distribution, validation, or failure semantics.

Run `test_tool_wrappers_main.ps1`, `test_collect_common_mods_main.ps1`, and
`test_dist_common_mods_repo_main.ps1` for the existing tool contracts.
