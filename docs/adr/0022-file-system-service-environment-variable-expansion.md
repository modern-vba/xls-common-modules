# ADR 0022: FileSystemService Expands Environment Variables in Local Paths Using OS Behavior

- Status: Accepted
- Date: 2026-06-11

## Context

Users may want to specify local paths that include Windows environment variables, such as `%LOCALAPPDATA%\tkmr-akhs\xls-web-tools`, in places like settings sheets for `BrowserProfilePath` or `WebDriverPath`. The previous `FileSystemService.GetAbsolutePath` did not expand environment variables, so values that users expected as OS path notation could differ from the actual values used for existence checks through common services or for paths passed directly to WebDriver without common services.

At the same time, Windows file and folder names can contain `%`, so if common modules strictly parsed `%NAME%` on their own and treated undefined variables as errors, valid paths that contain literal `%` could be broken. Users are also likely to expect `%VAR%` expansion rules to match the OS behavior.

## Decision

`FileSystemService.GetAbsolutePath` expands environment variables using OS behavior before absolute-path judgment, but only for local paths that are not URLs. The implementation uses `WScript.Shell.ExpandEnvironmentStrings`, and handling of undefined variables and literal `%` follows the OS result.

Environment variable expansion is not applied to URL strings. A string such as `https://example.com/%LOCALAPPDATA%/x` is treated as a URL and only URL path normalization is performed, as before.

If the expanded string is a relative path, it is converted to an absolute path relative to `WorkbookService.GetThisWorkbookDirectoryPath`, as before.

## Consequences

- `FsSrv.GetAbsolutePath("%LOCALAPPDATA%\foo")` returns the local absolute path after OS environment variable expansion.
- `FsSrv.GetAbsolutePath("https://example.com/%LOCALAPPDATA%/foo")` does not expand `%LOCALAPPDATA%`.
- Handling of local file names that contain `%` and undefined variables follows the result of the OS `ExpandEnvironmentStrings`, rather than a CommonModules-specific strict error.
- In tools such as `xls-web-tools`, when local paths from a settings sheet are passed to consumers, returning values after `FsSrv.GetAbsolutePath` aligns interpretation across existence checks, process startup, and paths passed to WebDriver capabilities.
