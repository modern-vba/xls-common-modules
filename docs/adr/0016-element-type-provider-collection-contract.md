# ADR 0016: Typed Element Collections Use Element Self-Declaration for Caller-Side Interface Contracts

- Status: Accepted
- Date: 2026-05-30

## Context

`ObjectList` and `ObjectSet` infer their element type contract from the first added element. Previously, object elements used the concrete type name as the contract, so elements of different concrete classes could not be stored in the same collection even when they implemented the same caller-side interface.

`CommonModules` cannot directly reference interfaces that are specific to consuming projects. At runtime, it may be possible to inspect `Implements` lines through VBIDE, but that would require ordinary users to allow access to the VBA project object model. For the same reason as ADR 0013, "Common Service Initialization Does Not Use VBIDE Optional-Service Detection," this would make the startup requirements for business macros unnecessarily heavy.

On the other hand, simply accepting all object types would allow objects unrelated to the caller's intended interface contract to enter the collection. A typed element collection needs a contract representation that does not depend solely on concrete type-name equality and does not depend on VBIDE.

## Decision

`ObjectList` and `ObjectSet` handle caller-side interface contracts through element type self-declaration with `IElementTypeProvider`.

`IElementTypeProvider` returns `ElementTypeKey`. `ElementTypeKey` consists of alphanumeric characters and `_` that are valid for class module names, starts with a letter, and is no longer than 31 characters. Comparisons are case-insensitive. If `ElementTypeKey` is empty or has an invalid form, operations that evaluate the element raise an error.

For object elements, the element type contract name is `ElementTypeKey` when the element implements `IElementTypeProvider`, and `TypeName` otherwise. `Nothing` is treated as a special null reference accepted by any object type contract. `ObjectList` and `ObjectSet` expose the raw contract name that can be passed to explicit initialization as `ElementTypeName`, and expose the string corresponding to the previous type display as `ItemTypeName`.

Explicit element type contracts are set through `ObjectList.Initialize` and `ObjectSet.Initialize`. Explicit initialization is only for object type contracts, and reinitialization after initialization or after adding elements is an error. When using only `New ObjectList` or `New ObjectSet`, the contract is inferred from the first element as before.

Required capabilities are managed as ordering through `IComparable` plus one mutually exclusive identity/duplicate-check mode. The identity/duplicate-check mode is one of `G_OBJECT_KEY_MODE_DUPLICATE_CHECKABLE`, `G_OBJECT_KEY_MODE_I_EQUATABLE`, or `G_OBJECT_KEY_MODE_REFERENCE`; explicit initialization prioritizes the specified value. Legacy mode infers from the first element in the order `IDuplicateCheckable`, `IEquatable`, then reference equality. `IStringable` is a display capability and is not included in required capabilities.

Add a `UseElementTypeKey` option to `Lib_Common.GetTypeString`, `GetTypedValueKey`, `GetValueKey`, and the multi-key APIs. The default is `False` to preserve compatibility with existing key strings. Internal key generation in `ObjectList` and `ObjectSet`, and display for non-`IStringable` objects, use `UseElementTypeKey:=True` only for top-level object elements. Even in reference-equality mode, the key string may include the element type contract name, but because element type contracts match within the same collection, practical duplicate checking is determined by object reference.

Array elements and primitive elements are outside element type self-declaration and keep their previous type-contract, comparison, and sorting behavior. The public scope and naming of `ArrayObject`, and the mutual import method names of `ObjectList` and `ObjectSet`, remain unchanged.

## Consequences

- Consuming projects that want to treat multiple concrete classes as elements under the same interface contract implement `IElementTypeProvider` on those classes and return the same `ElementTypeKey`.
- The implementation class is responsible for ensuring that the object is actually assignable to the type indicated by `ElementTypeKey`. Without VBIDE, `CommonModules` does not prove assignability to consuming-project-specific interfaces.
- `ElementTypeName` on `ObjectList` and `ObjectSet` returns the self-declared `ElementTypeKey` or `TypeName` for objects, and is used as the raw contract name that can be passed to `Initialize`. `ItemTypeName` returns a string corresponding to the previous type display, and for objects returns the `Object@ElementTypeName` form that reflects `UseElementTypeKey:=True`.
- `CopyList`, `CopySet`, `RemoveAll`, `ObjectSet.Sort`, and mutual collection import preserve or inherit element type contracts and required capabilities.
- Because the default behavior of `GetTypedValueKey` and related APIs is preserved, `TestDoubleVariantKeyBuilder` and `TestDoubleBehaviorStore` are not affected by element type self-declaration by default.
