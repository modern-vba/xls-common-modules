# ADR 0017: Standard For Each Is Exposed as Snapshot Enumeration

- Status: Accepted
- Date: 2026-05-30

## Context

`ObjectList`, `ObjectSet`, and `WorksheetRangeBounds` have `GetEnumerator` methods that return `IEnumerator`, but VBA's standard `For Each` cannot directly specify custom method names or enumerators with arguments. Consuming code ends up with many `Do While enum.MoveNext()` loops, which is verbose even for simple read-only enumeration.

Previously, `WorksheetRangeBounds` let each call specify row enumeration, column enumeration, horizontal cell enumeration, vertical cell enumeration, and descending order through the `EnumerateType`, `ColumnDirection`, and `Descending` arguments of `GetEnumerator`. However, standard `For Each` cannot pass arguments, so if the argument-based API and standard enumeration coexist, it becomes unclear which settings are the defaults.

## Decision

`ObjectList`, `ObjectSet`, and `WorksheetRangeBounds` support standard `For Each` through a hidden `NewEnum` property. `NewEnum` returns a snapshot by repacking the elements that exist when enumeration starts into a `Collection`.

`For Each` on `ObjectList` and `ObjectSet` enumerates elements in the same order as the existing `Item(index)`. When callers need explicit update, delete, descending order, or read-only control, or when they mutate the original collection during enumeration, they continue to use `GetEnumerator`.

`WorksheetRangeBounds` holds the defaults for standard enumeration in `EnumerationMode` and `EnumerationDescending`. `EnumerationMode` is one of `G_RANGE_ENUM_MODE_ROWS`, `G_RANGE_ENUM_MODE_COLUMNS`, `G_RANGE_ENUM_MODE_CELLS_HORIZONTAL`, or `G_RANGE_ENUM_MODE_CELLS_VERTICAL`, with row enumeration as the default. The default for `EnumerationDescending` is `False`.

`WorksheetRangeBounds.GetEnumerator` takes no arguments and follows `EnumerationMode` and `EnumerationDescending`. `WorksheetRangeBoundsEnumerator.Initialize` initializes from the enumeration mode and descending flag passed by `WorksheetRangeBounds`. The previous API that passed `EnumerateType`, `ColumnDirection`, and `Descending` to `WorksheetRangeBounds.GetEnumerator` is removed.

`EnumerationMode` and `EnumerationDescending` do not affect the identity of `WorksheetRangeBounds`. Ranges derived through `CopyObject`, `Transform`, `Shift`, `Intersect`, `GetRow`, `GetColumn`, `GetCell`, and similar operations inherit the original enumeration settings.

## Consequences

- Simple read-only enumeration can use `For Each item In collection`.
- `For Each` enumerates a snapshot from the start of enumeration. However, if the original `ObjectList`, `ObjectSet`, or `WorksheetRangeBounds` state or enumeration settings are changed during a standard `For Each`, continuation of the running loop is not guaranteed.
- Existing `WorksheetRangeBounds.GetEnumerator(EnumerateType:=..., ColumnDirection:=..., Descending:=...)` calls migrate by creating a copy of the target range, setting `EnumerationMode` and `EnumerationDescending`, and then calling `GetEnumerator()`.
- String representation, equality, and duplicate checking for `WorksheetRangeBounds` are not affected by enumeration settings.
