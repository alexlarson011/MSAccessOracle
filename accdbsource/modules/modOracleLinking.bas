Attribute VB_Name = "modOracleLinking"
'====================================================================================
' modOracleLinking
'====================================================================================
'
' README
' ------
' Purpose
' -------
' Provides Oracle linked-table inspection, creation, relinking, and validation
' helpers for this Access application.
'
' Although the application is moving toward a passthrough-first architecture,
' linked tables may still exist for legacy forms, setup workflows, compatibility,
' diagnostics, or transitional support. This module isolates all linked-table-
' specific logic in one place.
'
'
' Responsibilities
' ----------------
' This module is responsible for:
'
'     - determining whether a table exists
'     - determining whether a table is an Oracle-manageable ODBC linked table
'     - reading DSN and schema information from linked tables
'     - creating Oracle linked tables
'     - deleting linked tables
'     - refreshing linked tables
'     - changing linked table DSN and/or schema targets
'     - bulk retargeting linked tables
'     - testing linked tables
'     - detecting DSN mismatches across linked tables
'
'
' Key public helpers
' ------------------
' inspection:
'     OracleLink_IsLinkedTable
'     OracleLink_IsOracleODBCLinkedTable
'     OracleLink_TableExists
'     OracleLink_GetLinkedTableDSN
'     OracleLink_GetLinkedTableSchema
'     OracleLink_GetLinkedTableSourceName
'     OracleLink_GetLinkedTableNames
'     OracleLink_GetLinkedTableDetails
'
' creation / deletion / relinking:
'     OracleLink_CreateLinkedTable
'     OracleLink_DeleteLinkedTable
'     OracleLink_RefreshLinkedTable
'     OracleLink_SetLinkedTableConnection
'     OracleLink_SetAllLinkedTableConnections
'
' validation / diagnostics:
'     OracleLink_TestLinkedTable
'     OracleLink_TestAllLinkedTables
'     OracleLink_CheckLinkedTableDSNMismatch
'     OracleLink_ValidateLinkedTable
'
'
' Typical usage
' -------------
' Create a linked table:
'
'     OracleLink_CreateLinkedTable "my_local_name", "MY_DATA_SOURCE", "MY_SCHEMA", "MY_TABLE"
'
' Refresh a linked table:
'
'     OracleLink_RefreshLinkedTable "MY_LINKED_TABLE"
'
' Retarget one linked table:
'
'     Call OracleLink_SetLinkedTableConnection("MY_LINKED_TABLE", "MY_DATA_SOURCE_TEST", "MY_SCHEMA_TEST")
'
' Retarget all linked tables:
'
'     Set results = OracleLink_SetAllLinkedTableConnections("MY_DATA_SOURCE_TEST", "MY_SCHEMA", "MY_SCHEMA_TEST")
'
' Test one linked table:
'
'     If OracleLink_TestLinkedTable("MY_LINKED_TABLE") Then ...
'
'
' Dependencies
' ------------
' Depends on:
'
'     - modOracleDataAccess
'
'
' Common callers
' --------------
' Common callers include:
'
'     - frmLogin (legacy / compatibility flows)
'     - deployment utilities
'     - environment switching tools
'     - admin diagnostics
'
'
' External link handling
' ----------------------
' This module intentionally targets only Oracle-manageable ODBC linked tables.
'
' By default, bulk inspection, testing, and retargeting routines ignore other
' external linked tables such as SharePoint, Excel, text, or non-ODBC links.
'
'
' Design notes
' ------------
' This module intentionally focuses only on linked-table behavior.
'
' It should not contain:
'
'     - general passthrough query logic
'     - UI logic
'     - Oracle form-engine logic
'
' If the application eventually becomes fully unbound and linked tables are no longer
' needed, this module may become mostly legacy / admin-only infrastructure.
'
'
' Version
' -------
' v1
'
'====================================================================================

Option Compare Database
Option Explicit

Private Const cModuleName As String = "modOracleLinking"

'------------------------------------------------------------------------------------
' Linked-table inspection helpers
'------------------------------------------------------------------------------------

Private Function OracleLink_GetTableDef(ByVal sTableName As String) As DAO.TableDef

    Dim db As DAO.Database
    Dim tdf As DAO.TableDef

    Set db = CurrentDb

    On Error Resume Next
    Set tdf = db.TableDefs(sTableName)
    On Error GoTo 0

    Set OracleLink_GetTableDef = tdf

    Set tdf = Nothing
    Set db = Nothing

End Function

Private Function OracleLink_IsOracleODBCLinkedTableDef(ByRef tdf As DAO.TableDef) As Boolean

    Dim sConnect As String

    If tdf Is Nothing Then Exit Function

    sConnect = Nz(tdf.Connect, vbNullString)

    If Len(sConnect) = 0 Then Exit Function
    If (tdf.Attributes And dbAttachedODBC) = 0 Then Exit Function
    If InStr(1, sConnect, "ODBC;", vbTextCompare) = 0 Then Exit Function
    If InStr(1, sConnect, "DSN=", vbTextCompare) = 0 Then Exit Function
    If Len(Nz(tdf.SourceTableName, vbNullString)) = 0 Then Exit Function

    OracleLink_IsOracleODBCLinkedTableDef = True

End Function

Public Function OracleLink_IsLinkedTable(ByVal sTableName As String) As Boolean

    Dim tdf As DAO.TableDef

    Set tdf = OracleLink_GetTableDef(sTableName)

    If tdf Is Nothing Then GoTo Cleanup

    OracleLink_IsLinkedTable = OracleLink_IsOracleODBCLinkedTableDef(tdf)

Cleanup:
    Set tdf = Nothing

End Function

Public Function OracleLink_IsOracleODBCLinkedTable(ByVal sTableName As String) As Boolean

    Dim tdf As DAO.TableDef

    Set tdf = OracleLink_GetTableDef(sTableName)
    OracleLink_IsOracleODBCLinkedTable = OracleLink_IsOracleODBCLinkedTableDef(tdf)

Cleanup:
    Set tdf = Nothing

End Function

Public Function OracleLink_TableExists(ByVal sTableName As String) As Boolean

    Dim db As DAO.Database
    Dim tdf As DAO.TableDef

    Set db = CurrentDb

    On Error Resume Next
    Set tdf = db.TableDefs(sTableName)
    On Error GoTo 0

    OracleLink_TableExists = Not tdf Is Nothing

Cleanup:
    Set tdf = Nothing
    Set db = Nothing

End Function

Public Function OracleLink_GetLinkedTableDSN(ByVal sTableName As String) As String

    Dim tdf As DAO.TableDef
    Dim sConnect As String
    Dim lStartPos As Long
    Dim lEndPos As Long

    Set tdf = OracleLink_GetTableDef(sTableName)

    If tdf Is Nothing Then Exit Function
    If Not OracleLink_IsOracleODBCLinkedTableDef(tdf) Then Exit Function

    sConnect = Nz(tdf.Connect, vbNullString)

    If Len(sConnect) = 0 Then Exit Function

    lStartPos = InStr(1, sConnect, "DSN=", vbTextCompare)

    If lStartPos > 0 Then
        lEndPos = InStr(lStartPos, sConnect & ";", ";")
        If lEndPos > 0 Then
            OracleLink_GetLinkedTableDSN = UCase$(Mid$(sConnect, lStartPos + 4, lEndPos - lStartPos - 4))
        End If
    End If

Cleanup:
    Set tdf = Nothing

End Function

Public Function OracleLink_GetLinkedTableSchema(ByVal sTableName As String) As String

    Dim tdf As DAO.TableDef
    Dim lDotPos As Long
    Dim sSourceTableName As String

    Set tdf = OracleLink_GetTableDef(sTableName)

    If tdf Is Nothing Then Exit Function
    If Not OracleLink_IsOracleODBCLinkedTableDef(tdf) Then Exit Function

    sSourceTableName = Nz(tdf.SourceTableName, vbNullString)

    If Len(sSourceTableName) = 0 Then Exit Function

    lDotPos = InStr(1, sSourceTableName, ".", vbTextCompare)

    If lDotPos > 1 Then
        OracleLink_GetLinkedTableSchema = UCase$(Left$(sSourceTableName, lDotPos - 1))
    End If

Cleanup:
    Set tdf = Nothing

End Function

Public Function OracleLink_GetLinkedTableSourceName(ByVal sTableName As String) As String

    Dim tdf As DAO.TableDef
    Dim lDotPos As Long
    Dim sSourceTableName As String

    Set tdf = OracleLink_GetTableDef(sTableName)

    If tdf Is Nothing Then Exit Function
    If Not OracleLink_IsOracleODBCLinkedTableDef(tdf) Then Exit Function

    sSourceTableName = Nz(tdf.SourceTableName, vbNullString)

    If Len(sSourceTableName) = 0 Then Exit Function

    lDotPos = InStr(1, sSourceTableName, ".", vbTextCompare)

    If lDotPos > 0 Then
        OracleLink_GetLinkedTableSourceName = Mid$(sSourceTableName, lDotPos + 1)
    Else
        OracleLink_GetLinkedTableSourceName = sSourceTableName
    End If

Cleanup:
    Set tdf = Nothing

End Function

Public Function OracleLink_GetLinkedTableNames() As Collection

    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim result As Collection

    Set result = New Collection
    Set db = CurrentDb

    For Each tdf In db.TableDefs
        If OracleLink_IsOracleODBCLinkedTableDef(tdf) Then
            result.Add tdf.Name
        End If
    Next tdf

    Set OracleLink_GetLinkedTableNames = result

Cleanup:
    Set tdf = Nothing
    Set db = Nothing

End Function

Public Function OracleLink_GetLinkedTableDetails() As Collection

    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim result As Collection
    Dim item As Object

    Set result = New Collection
    Set db = CurrentDb

    For Each tdf In db.TableDefs
        If OracleLink_IsOracleODBCLinkedTableDef(tdf) Then
            Set item = CreateObject("Scripting.Dictionary")
            item.Add "TableName", tdf.Name
            item.Add "DSN", OracleLink_GetLinkedTableDSN(tdf.Name)
            item.Add "Schema", OracleLink_GetLinkedTableSchema(tdf.Name)
            item.Add "SourceTableName", Nz(tdf.SourceTableName, vbNullString)
            item.Add "Connect", Nz(tdf.Connect, vbNullString)
            result.Add item
        End If
    Next tdf

    Set OracleLink_GetLinkedTableDetails = result

Cleanup:
    Set item = Nothing
    Set tdf = Nothing
    Set db = Nothing

End Function

'------------------------------------------------------------------------------------
' Linked-table creation / deletion / relinking
'------------------------------------------------------------------------------------

Public Function OracleLink_CreateLinkedTable( _
    ByVal sTableName As String, _
    ByVal sDSN As String, _
    ByVal sSourceSchema As String, _
    ByVal sSourceODBCTableName As String, _
    Optional ByVal failIfExists As Boolean = True _
) As Boolean

    Dim db As DAO.Database
    Dim tdf As DAO.TableDef

    sTableName = Trim$(sTableName)
    sDSN = UCase$(Trim$(sDSN))
    sSourceSchema = UCase$(Trim$(sSourceSchema))
    sSourceODBCTableName = Trim$(sSourceODBCTableName)

    If Len(sTableName) = 0 Then
        Err.Raise vbObjectError + 3000, cModuleName & ".OracleLink_CreateLinkedTable", "Table name cannot be blank."
    End If

    If Len(sDSN) = 0 Then
        Err.Raise vbObjectError + 3001, cModuleName & ".OracleLink_CreateLinkedTable", "DSN cannot be blank."
    End If

    If Len(sSourceSchema) = 0 Then
        Err.Raise vbObjectError + 3002, cModuleName & ".OracleLink_CreateLinkedTable", "Source schema cannot be blank."
    End If

    If Len(sSourceODBCTableName) = 0 Then
        Err.Raise vbObjectError + 3003, cModuleName & ".OracleLink_CreateLinkedTable", "Source ODBC table name cannot be blank."
    End If

    If Not Test_ODBC_Conn(sDSN) Then
        Err.Raise vbObjectError + 3004, cModuleName & ".OracleLink_CreateLinkedTable", "Unable to connect to DSN: " & sDSN
    End If

    Set db = CurrentDb

    If OracleLink_TableExists(sTableName) Then

        If failIfExists Then
            Err.Raise vbObjectError + 3005, cModuleName & ".OracleLink_CreateLinkedTable", _
                      "Table already exists: " & sTableName
        Else
            OracleLink_CreateLinkedTable = False
            GoTo Cleanup
        End If

    End If

    Set tdf = db.CreateTableDef(sTableName)

    With tdf
        .Connect = Get_ODBC_Conn_Str(sDSN)
        .SourceTableName = sSourceSchema & "." & sSourceODBCTableName
    End With

    db.TableDefs.Append tdf
    db.TableDefs.Refresh

    OracleLink_CreateLinkedTable = True

Cleanup:
    Set tdf = Nothing
    Set db = Nothing

End Function

Public Sub OracleLink_DeleteLinkedTable(ByVal sTableName As String)

    Dim db As DAO.Database

    sTableName = Trim$(sTableName)

    If Len(sTableName) = 0 Then
        Err.Raise vbObjectError + 3010, cModuleName & ".OracleLink_DeleteLinkedTable", "Table name cannot be blank."
    End If

    If Not OracleLink_TableExists(sTableName) Then
        Err.Raise vbObjectError + 3011, cModuleName & ".OracleLink_DeleteLinkedTable", "Table does not exist: " & sTableName
    End If

    If Not OracleLink_IsLinkedTable(sTableName) Then
        Err.Raise vbObjectError + 3012, cModuleName & ".OracleLink_DeleteLinkedTable", "Table is not an Oracle ODBC linked table: " & sTableName
    End If

    Set db = CurrentDb
    db.TableDefs.Delete sTableName
    db.TableDefs.Refresh

Cleanup:
    Set db = Nothing

End Sub

Public Sub OracleLink_RefreshLinkedTable(ByVal sTableName As String)

    Dim db As DAO.Database
    Dim tdf As DAO.TableDef

    sTableName = Trim$(sTableName)

    If Len(sTableName) = 0 Then
        Err.Raise vbObjectError + 3013, cModuleName & ".OracleLink_RefreshLinkedTable", "Table name cannot be blank."
    End If

    Set db = CurrentDb

    On Error Resume Next
    Set tdf = db.TableDefs(sTableName)
    On Error GoTo HandleErr

    If tdf Is Nothing Then
        Err.Raise vbObjectError + 3014, cModuleName & ".OracleLink_RefreshLinkedTable", "Table not found: " & sTableName
    End If

    If Not OracleLink_IsOracleODBCLinkedTableDef(tdf) Then
        Err.Raise vbObjectError + 3015, cModuleName & ".OracleLink_RefreshLinkedTable", "Table is not an Oracle ODBC linked table: " & sTableName
    End If

    tdf.RefreshLink

Cleanup:
    Set tdf = Nothing
    Set db = Nothing
    Exit Sub

HandleErr:
    Err.Raise vbObjectError + 3016, cModuleName & ".OracleLink_RefreshLinkedTable", _
              "RefreshLink failed for " & sTableName & ". Details: " & Err.Description

End Sub

Public Function OracleLink_SetLinkedTableConnection( _
    ByVal sTableName As String, _
    Optional ByVal sToDSN As String = "", _
    Optional ByVal sToSchema As String = "" _
) As Boolean

    Dim db As DAO.Database
    Dim tdf As DAO.TableDef

    Dim sStartSchema As String
    Dim sStartDSN As String
    Dim sSourceODBCTableName As String

    Dim bNewDSN As Boolean
    Dim bNewSchema As Boolean

    sTableName = Trim$(sTableName)
    sToDSN = UCase$(Trim$(sToDSN))
    sToSchema = UCase$(Trim$(sToSchema))

    If Len(sTableName) = 0 Then
        Err.Raise vbObjectError + 3020, cModuleName & ".OracleLink_SetLinkedTableConnection", "Table name cannot be blank."
    End If

    Set db = CurrentDb

    On Error Resume Next
    Set tdf = db.TableDefs(sTableName)
    On Error GoTo HandleErr

    If tdf Is Nothing Then
        Err.Raise vbObjectError + 3021, cModuleName & ".OracleLink_SetLinkedTableConnection", "Table not found: " & sTableName
    End If

    If Not OracleLink_IsOracleODBCLinkedTableDef(tdf) Then
        Err.Raise vbObjectError + 3022, cModuleName & ".OracleLink_SetLinkedTableConnection", "Table is not an Oracle ODBC linked table: " & sTableName
    End If

    sStartDSN = OracleLink_GetLinkedTableDSN(sTableName)
    sStartSchema = OracleLink_GetLinkedTableSchema(sTableName)
    sSourceODBCTableName = OracleLink_GetLinkedTableSourceName(sTableName)

    If Len(sToDSN) = 0 Then sToDSN = sStartDSN
    If Len(sToSchema) = 0 Then sToSchema = sStartSchema

    If Len(sToDSN) = 0 Then
        Err.Raise vbObjectError + 3023, cModuleName & ".OracleLink_SetLinkedTableConnection", "Unable to resolve target DSN."
    End If

    If Len(sToSchema) = 0 Then
        Err.Raise vbObjectError + 3024, cModuleName & ".OracleLink_SetLinkedTableConnection", "Unable to resolve target schema."
    End If

    If Not Test_ODBC_Conn(sToDSN) Then
        Err.Raise vbObjectError + 3025, cModuleName & ".OracleLink_SetLinkedTableConnection", "Unable to connect to target DSN: " & sToDSN
    End If

    bNewDSN = (sToDSN <> sStartDSN)
    bNewSchema = (sToSchema <> sStartSchema)

    If Not bNewDSN And Not bNewSchema Then
        OracleLink_SetLinkedTableConnection = True
        GoTo Cleanup
    End If

    If bNewDSN And Not bNewSchema Then

        tdf.Connect = Get_ODBC_Conn_Str(sToDSN)
        tdf.RefreshLink

        OracleLink_SetLinkedTableConnection = True
        GoTo FinalValidation

    End If

    If bNewSchema Then

        db.TableDefs.Delete sTableName
        db.TableDefs.Refresh

        Call OracleLink_CreateLinkedTable( _
            sTableName:=sTableName, _
            sDSN:=sToDSN, _
            sSourceSchema:=sToSchema, _
            sSourceODBCTableName:=sSourceODBCTableName, _
            failIfExists:=True)

        OracleLink_SetLinkedTableConnection = True

    End If

FinalValidation:

    If OracleLink_GetLinkedTableDSN(sTableName) <> sToDSN Then
        OracleLink_SetLinkedTableConnection = False
    End If

    If OracleLink_GetLinkedTableSchema(sTableName) <> sToSchema Then
        OracleLink_SetLinkedTableConnection = False
    End If

Cleanup:
    Set tdf = Nothing
    Set db = Nothing
    Exit Function

HandleErr:
    Err.Raise vbObjectError + 3026, cModuleName & ".OracleLink_SetLinkedTableConnection", _
              "Set linked table connection failed for " & sTableName & ". Details: " & Err.Description

End Function

Public Function OracleLink_SetAllLinkedTableConnections( _
    ByVal sToDSN As String, _
    Optional ByVal sFromSchema As String = "", _
    Optional ByVal sToSchema As String = "" _
) As Collection

    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim results As Collection
    Dim item As Object

    Dim bUseSchemaSwap As Boolean
    Dim bShouldProcess As Boolean
    Dim bSucceeded As Boolean

    sToDSN = UCase$(Trim$(sToDSN))
    sFromSchema = UCase$(Trim$(sFromSchema))
    sToSchema = UCase$(Trim$(sToSchema))

    If Len(sToDSN) = 0 Then
        Err.Raise vbObjectError + 3030, cModuleName & ".OracleLink_SetAllLinkedTableConnections", "Target DSN cannot be blank."
    End If

    If (Len(sFromSchema) = 0 Xor Len(sToSchema) = 0) Then
        Err.Raise vbObjectError + 3031, cModuleName & ".OracleLink_SetAllLinkedTableConnections", _
                  "Source and target schema must both be supplied or both omitted."
    End If

    bUseSchemaSwap = (Len(sFromSchema) > 0)

    If Not Test_ODBC_Conn(sToDSN) Then
        Err.Raise vbObjectError + 3032, cModuleName & ".OracleLink_SetAllLinkedTableConnections", "Unable to connect to target DSN: " & sToDSN
    End If

    Set results = New Collection
    Set db = CurrentDb

    For Each tdf In db.TableDefs

        bShouldProcess = OracleLink_IsOracleODBCLinkedTableDef(tdf)

        If bShouldProcess Then

            Set item = CreateObject("Scripting.Dictionary")
            item.Add "TableName", tdf.Name
            item.Add "StartDSN", OracleLink_GetLinkedTableDSN(tdf.Name)
            item.Add "StartSchema", OracleLink_GetLinkedTableSchema(tdf.Name)

            On Error Resume Next

            If bUseSchemaSwap Then
                If OracleLink_GetLinkedTableSchema(tdf.Name) = sFromSchema Then
                    bSucceeded = OracleLink_SetLinkedTableConnection(tdf.Name, sToDSN, sToSchema)
                Else
                    bSucceeded = OracleLink_SetLinkedTableConnection(tdf.Name, sToDSN)
                End If
            Else
                bSucceeded = OracleLink_SetLinkedTableConnection(tdf.Name, sToDSN)
            End If

            If Err.Number = 0 Then
                item.Add "Succeeded", bSucceeded
                item.Add "ErrorText", vbNullString
            Else
                item.Add "Succeeded", False
                item.Add "ErrorText", Err.Description
                Err.Clear
            End If

            On Error GoTo 0

            item.Add "EndDSN", OracleLink_GetLinkedTableDSN(tdf.Name)
            item.Add "EndSchema", OracleLink_GetLinkedTableSchema(tdf.Name)

            results.Add item

        End If

    Next tdf

    Set OracleLink_SetAllLinkedTableConnections = results

Cleanup:
    Set item = Nothing
    Set tdf = Nothing
    Set db = Nothing

End Function

'------------------------------------------------------------------------------------
' Linked-table test / validation helpers
'------------------------------------------------------------------------------------

Public Function OracleLink_TestLinkedTable(ByVal sTableName As String) As Boolean

    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim rs As DAO.Recordset

    sTableName = Trim$(sTableName)

    If Len(sTableName) = 0 Then
        Err.Raise vbObjectError + 3040, cModuleName & ".OracleLink_TestLinkedTable", "Table name cannot be blank."
    End If

    Set db = CurrentDb

    On Error Resume Next
    Set tdf = db.TableDefs(sTableName)
    On Error GoTo HandleErr

    If tdf Is Nothing Then
        Err.Raise vbObjectError + 3041, cModuleName & ".OracleLink_TestLinkedTable", "Table not found: " & sTableName
    End If

    If Not OracleLink_IsOracleODBCLinkedTableDef(tdf) Then
        Err.Raise vbObjectError + 3042, cModuleName & ".OracleLink_TestLinkedTable", "Table is not an Oracle ODBC linked table: " & sTableName
    End If

    Set rs = tdf.OpenRecordset(dbOpenSnapshot)
    OracleLink_TestLinkedTable = True

Cleanup:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set tdf = Nothing
    Set db = Nothing
    Exit Function

HandleErr:
    Err.Raise vbObjectError + 3043, cModuleName & ".OracleLink_TestLinkedTable", _
              "Linked table test failed for " & sTableName & ". Details: " & Err.Description

End Function

Public Function OracleLink_TestAllLinkedTables() As Collection

    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim results As Collection
    Dim item As Object
    Dim shouldTest As Boolean

    Set results = New Collection
    Set db = CurrentDb

    For Each tdf In db.TableDefs

        shouldTest = OracleLink_IsOracleODBCLinkedTableDef(tdf)

        If shouldTest Then
            Set item = CreateObject("Scripting.Dictionary")
            item.Add "TableName", tdf.Name
            item.Add "DSN", OracleLink_GetLinkedTableDSN(tdf.Name)
            item.Add "Schema", OracleLink_GetLinkedTableSchema(tdf.Name)

            On Error Resume Next
            Call OracleLink_TestLinkedTable(tdf.Name)
            If Err.Number = 0 Then
                item.Add "Succeeded", True
                item.Add "ErrorText", vbNullString
            Else
                item.Add "Succeeded", False
                item.Add "ErrorText", Err.Description
                Err.Clear
            End If
            On Error GoTo 0

            results.Add item
        End If

    Next tdf

    Set OracleLink_TestAllLinkedTables = results

Cleanup:
    Set item = Nothing
    Set tdf = Nothing
    Set db = Nothing

End Function

Public Function OracleLink_CheckLinkedTableDSNMismatch( _
    Optional ByVal sDSNCheckValue As String = "" _
) As Boolean

    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim sCheck As String

    If Len(Trim$(sDSNCheckValue)) = 0 Then
        sCheck = Get_DB_DSN()
    Else
        sCheck = UCase$(Trim$(sDSNCheckValue))
    End If

    Set db = CurrentDb

    For Each tdf In db.TableDefs
        If OracleLink_IsOracleODBCLinkedTableDef(tdf) Then
            If OracleLink_GetLinkedTableDSN(tdf.Name) <> sCheck Then
                OracleLink_CheckLinkedTableDSNMismatch = True
                Exit Function
            End If
        End If
    Next tdf

Cleanup:
    Set tdf = Nothing
    Set db = Nothing

End Function

Public Sub OracleLink_ValidateLinkedTable(ByVal sTableName As String)

    If Not OracleLink_TestLinkedTable(sTableName) Then
        Err.Raise vbObjectError + 3050, cModuleName & ".OracleLink_ValidateLinkedTable", _
                  "Linked table validation failed: " & sTableName
    End If

End Sub
