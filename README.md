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
- passthrough query DSNs

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
