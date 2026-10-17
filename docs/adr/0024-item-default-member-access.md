# ADR 0024: Item Is Exposed as Default Member Access

- Status: Accepted
- Date: 2026-06-13

## Context

`ObjectList`, `ObjectSet`, `ObjectDictionary`, `ArrayObject`, `WorksheetRangeBounds`, and `WorksheetVirtualTable` expose public APIs that reference elements, rows, or cell ranges through `Item(...)`. In VBA, making `Item` the default member allows the same target to be referenced as `obj(index)` in addition to `obj.Item(index)`, but it also increases implicit calls, so there is a trade-off between readability and convenience.

## Decision

Common collections and range views that have `Item` expose `Item` as the default member. Existing `.Item(...)` calls remain valid; default member access is treated as an additional call surface.

`ObjectList` and `ArrayObject` provide `Let` and `Set` assignment to `Item(index)` as updates to an existing index. An out-of-range index is not treated as an add; it is handled as the same update error as the existing `Update`.

As described in ADR 0020, `ObjectDictionary` treats assignment to `Item(key)` as having the same meaning as `AddOrUpdate(key, value)`. `ObjectSet`, `WorksheetRangeBounds`, and `WorksheetVirtualTable` expose only read-only `Item` as the default member and do not provide assignment to `Item`.

## Consequences

- Callers can reference existing elements with either `list.Item(0)` or `list(0)`.
- This decision does not include replacing large numbers of existing calls with default member access.
- Update syntax is limited to ordered, updatable `ObjectList` and `ArrayObject`, and to `ObjectDictionary`, which supports keyed add/update.
