# MSAccessOracle
Modules and Tools for MS Access and Oracle

Build using msaccess-vcs-addin:
https://github.com/joyfullservice/msaccess-vcs-addin

## Setup

This repository seeds the sample configuration with the DSN name `MY_DATA_SOURCE`.

After building the database from source, update the environment configuration to match
your Oracle database / ODBC DSN by running `OracleAdmin_SwitchEnvironment(...)` from
the Immediate Window or your own setup routine.

Example:

```vb
Call OracleAdmin_SwitchEnvironment( _
    envName:="PROD", _
    dsnName:="YOUR_DATABASE_DSN", _
    schemaName:="YOUR_SCHEMA")
```

This updates:

- `tblConn.ENV`
- `tblConn.DSN`
- `tblConn.SCHEMA`

It does not test the DSN connection or update saved passthrough queries / linked
tables unless you ask for those updates.

If you also use Oracle ODBC linked tables and want those retargeted too, run:

```vb
Call OracleAdmin_SwitchEnvironment( _
    envName:="PROD", _
    dsnName:="YOUR_DATABASE_DSN", _
    schemaName:="YOUR_SCHEMA", _
    updatePassthroughQueries:=True, _
    updateLinkedTables:=True, _
    linkedTableFromSchema:="OLD_SCHEMA", _
    linkedTableToSchema:="YOUR_SCHEMA")
```

For stateless / passthrough-only use, the first example is usually all you need.
If you still maintain saved passthrough query objects for admin/testing work, pass
`updatePassthroughQueries:=True`.

## Documentation

- [Using `clsOracleFormField` and `modOracleFormEngine`](./FORMFIELD_GUIDE.md)
- [Base form pattern](./BASE_FORM_PATTERN.md)
- [Proposed form test plan](./FORM_TEST_PLAN.md)

After logging in, you can run a quick runtime smoke test from the Immediate Window:

```vb
OracleSmoke_RunRuntime
```
