# Do Not Limit ClearButton Deletion Targets to Created Buttons

`ClearButton` is treated as a common API that deletes every button on a worksheet.
It is not a feature for deleting only buttons created by `AddButton` or UnitTest rerun buttons.

## Why this is out of scope

In this repository, a button is treated as "a shape on a worksheet that runs a macro through user interaction."
Therefore, limiting `ClearButton` deletion targets only to artifacts created by `AddButton` would conflict with the API name and responsibility.

The deletion-target rule should remain the same as the current implementation: treat any `Shape` with `OnAction` set as a button.
Even if it is an image or an arbitrary shape, treat it as a button when `OnAction` is set so it can run a macro.

If a use case that needs to delete only buttons created by `AddButton` becomes necessary in the future, handle it as a separate API or as caller-side management responsibility instead of narrowing the meaning of `ClearButton`.

## Prior requests

- #1 - Limit AddButton / ClearButton deletion targets to created buttons
