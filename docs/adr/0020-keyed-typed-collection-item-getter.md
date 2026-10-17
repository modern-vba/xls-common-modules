# ADR 0020: The Item Getter on Keyed Typed Element Collections Does Not Add Missing Keys

- Status: Accepted
- Date: 2026-06-11

## Context

Keyed typed element collections are expected to expose a call surface close to `Scripting.Dictionary`. `Scripting.Dictionary` is widely used as a standard associative array that maps keys to items, and keys can be any value except arrays.

However, the `Item(Key)` getter on `Scripting.Dictionary` adds a new key with an empty item merely by reading a missing key. This behavior does not fit the `CommonModules` collection contract because a read operation mutates state and can violate the element type contract held by a typed element collection.

`ObjectList` and `ObjectSet` explicitly manage the element type contract and required capabilities. Keyed typed element collections must prioritize predictable read operations and preservation of the element type contract over the convenience of key access.

## Decision

The `Item(Key)` getter on a keyed typed element collection does not add a new element when a missing key is referenced. It raises an error.

Assignment to `Item(Key)` adds the value when the key does not exist and updates the value when the key already exists. In both cases, the value is subject to the same element type contract and required capabilities as a typed element collection. Assignment to `Item(Key)` has the same meaning as `AddOrUpdate(Key, Value)`.

Expose explicit methods: `Add(Key, Value)`, `Update(Key, Value)`, and `AddOrUpdate(Key, Value)`. `Add` errors on an existing key, `Update` errors on a missing key, and `AddOrUpdate` adds or updates depending on whether the key exists.

This behavior difference is treated as an intentional public API difference from full `Scripting.Dictionary` compatibility.

The explicit element type contract for values is set with the same `Initialize(ElementTypeName, RequireComparable, ObjectKeyMode)` used by `ObjectList` and `ObjectSet`. Key comparison mode is a separate contract from the value element type contract and is set through the `CompareMode` property. `CompareMode` cannot be changed after the first add, and it is preserved together with the value element type contract after `RemoveAll`.

The initial implementation internally has a `Scripting.Dictionary` from key to index, a `Collection` that preserves key order, and an `ObjectList` that stores values. The key side follows the key contract of `Scripting.Dictionary`, so it is not managed by the typed element collection.

## Consequences

- Read-only references do not change the count or element type state of a keyed typed element collection.
- `RemoveAll` removes values and keys, but preserves the value element type contract and `CompareMode`.
- Callers check `Exists(Key)` before reading a key that might not exist.
- Assignment to `Item(Key)` and `AddOrUpdate(Key, Value)` can be used for add or update, but they raise an error if the value violates the type contract.
- Code ported from `Scripting.Dictionary` that depends on implicit add must be replaced with explicit `Add` or assignment to `Item(Key)`.
