# VBA Style Guide

This repository uses a lightweight, Access-friendly VBA naming convention.

The goal is consistency, readability, and easier maintenance in exported source.

## Core Rules

1. Use `PascalCase` for procedures, functions, properties, classes, and public API names.
2. Use `mod`, `cls`, and `frm` prefixes for Access object modules.
3. Prefix shared public routines by domain when they are part of a reusable module API.
4. Use lightweight type/context prefixes for local variables.
5. Use boolean names that read like predicates or decisions.
6. Use `p...` for private backing fields in classes.
7. Use one constant style consistently within a module.
8. Prefer descriptive names over short abbreviations unless the abbreviation is standard in Access/DAO code.

## Module Names

- Standard modules: `modOracleDataAccess`, `modOracleAdmin`, `modIniConfig`
- Class modules: `clsOracleFormField`
- Form modules: `frmLogin`

## Procedures And Functions

Use `PascalCase`.

- Public shared API: `OracleAdmin_SwitchEnvironment`
- Public shared API: `OracleLink_SetLinkedTableConnection`
- Public shared API: `OracleSession_IsConnected`
- Private helper: `CreatePassthroughQueryDef`
- Private helper: `ValidateModuleConfiguration`

For boolean-returning functions, use names that read like questions or states.

- `OracleSession_IsConnected`
- `OracleAdmin_tblConn_Exists`
- `OracleLink_IsLinkedTable`

## Variables

Use lightweight prefixes when the type or role improves clarity.

- Strings: `sDSN`, `sSQL`, `sSchema`
- Booleans: `bSucceeded`, `bUseSchemaSwap`, `bShouldProcess`
- Long integers: `lStartPos`, `lEndPos`, `lDotPos`, `lRowCount`
- Variants: `vValue`, `vResult`
- Collections: `colFields`, `colRows`
- Dictionaries / generic objects: `dictRow`, `objItem`
- DAO objects: `db`, `rs`, `qdf`, `tdf`, `fld`, `prp`
- Forms / controls: `frm`, `ctl`

Counter variables should match the actual type. Prefer `lIndex` over `i` when the variable is declared `As Long`.

## Constants

Use one of these patterns consistently within a module:

- Private module constant: `cModuleName`
- Public shared constant: `ORACLE_CONNECTION_STRING_BASE`

Avoid mixing several unrelated constant styles in the same module unless there is a clear reason.

## Class Properties And Fields

Use `PascalCase` for public properties.

- `DefaultValue`
- `ControlKind`
- `OracleDataType`

Use `p...` for private backing fields.

- `pDefaultValue`
- `pControlKind`

## Documentation

Each exported VBA module should keep the current README-style header blocks used in this repository.

Headers should stay aligned with the actual code:

- list real dependencies
- list real public helpers
- avoid references to retired or renamed modules

## Error Handling

Shared modules should prefer explicit `Err.Raise` behavior over silent fallback when the failure represents a real configuration or runtime problem.

UI modules may present `MsgBox` messages, but should still use a consistent `On Error GoTo ...` pattern for operations that can fail unexpectedly.
