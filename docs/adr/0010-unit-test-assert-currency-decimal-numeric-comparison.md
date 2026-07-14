# ADR 0010: UnitTestAssert.EqualsNumeric Does Not Include Currency and Decimal in Normal Numeric Comparison

- Status: Accepted
- Date: 2026-05-23

## Context

`UnitTestAssert.EqualsNumeric` treats types such as `Byte`, `Integer`, `Long`, `LongLong`, `Single`, and `Double` as numeric, but it does not include `Currency` or `Decimal`.

`Currency` is fixed-point and has rounding and scale semantics. `Decimal` is handled as a `Variant` subtype and is hard to handle strictly with normal type checks.

## Decision

`Currency` and `Decimal` are not included in the normal numeric comparison targets for `EqualsNumeric`.

If dedicated comparisons for money or high-precision decimals become necessary, consider them as separate APIs with explicit rounding, scale, and tolerance.

## Consequences

- Comparisons such as `EqualsNumeric CCur(1), 1` do not succeed as normal numeric comparisons.
- For tests involving money or high-precision decimals, callers currently need to express the expected comparison contract explicitly.
- Future APIs should not mix `Currency` / `Decimal` precision contracts into `EqualsNumeric`.
