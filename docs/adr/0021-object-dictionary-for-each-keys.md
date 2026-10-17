# ADR 0021: Standard For Each on Keyed Typed Element Collections Enumerates Keys

- Status: Accepted
- Date: 2026-06-11

## Context

ADR 0017 decided that standard `For Each` on `ObjectList` and `ObjectSet` enumerates elements as a snapshot. A keyed typed element collection could follow the same policy and enumerate values as elements.

However, keyed typed element collections have a call surface close to `Scripting.Dictionary`. For users reading them as dictionaries, `For Each item In dictionary` is strongly expected to enumerate keys. When callers want to enumerate values, they can explicitly use `Items`, which returns the values as a typed element collection.

Because standard `For Each` does not let callers specify the enumeration target, choosing keys or values as the default is a public API decision that is hard to change later.

## Decision

Standard `For Each` on keyed typed element collections enumerates keys as a snapshot.

To enumerate values, callers enumerate the `ObjectList` returned by `Items`. To retrieve keys as an array, callers use `Keys`; to retrieve values as an array, callers use `ConvertToArray`.

The initial version does not expose `GetEnumerator`. The existing `IEnumerator` assumes index-based `Item`, `Remove`, and `Update`, but `Item` on keyed typed element collections is key-based, so it is not forced into the same enumerator contract.

The enumeration order for standard `For Each` is the same as the key order returned by `Keys`. Key order is insertion order. Updating an existing key preserves order; deletion compacts the remaining keys and values so deleted slots disappear from enumeration order and array conversion results.

## Consequences

- `For Each key In dictionary` enumerates keys in insertion order.
- `For Each value In dictionary.Items` enumerates values.
- Standard `For Each` on `ObjectList` and `ObjectSet` remains element enumeration, but keyed typed element collections prioritize dictionary usability.
- Code that implicitly expects value enumeration must explicitly use `Items`.
- If a dedicated enumerator that supports update or delete becomes necessary, design `GetKeyEnumerator`, `GetItemEnumerator`, or a dedicated enumerator separately.
