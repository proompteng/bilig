# WASM Runtime Contract

## Scope

`@bilig/wasm-kernel` is the production formula executor for closed families.

## Required runtime capabilities

- typed scalar values:
  - empty
  - number
  - boolean
  - string
  - error
- array/spill values
- range/reference iteration
- builtin dispatch by stable builtin id
- structured error output
- recalc epoch input for volatile behavior

## Promotion rule

A formula family can route to WASM in production only if:

- parser/binder support is complete
- JS oracle is fixture-green
- WASM matches JS in differential tests
- kernel transport supports all value shapes needed by that family

## Explicit non-goal

JS and WASM must not each own independent semantics. JS defines semantics; WASM mirrors them exactly until the family is closed.
