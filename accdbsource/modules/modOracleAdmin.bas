Attribute VB_Name = "modOracleAdmin"
'====================================================================================
' modOracleAdmin
'====================================================================================
'
' README
' ------
' Purpose
' -------
' Provides Oracle / ODBC administration helpers for this Access application.
'
' This module handles application-level configuration checks and passthrough-query
' administration tasks that are related to setup, validation, inspection, and bulk
' maintenance, but are not part of the low-level query execution layer.
'
'
' Responsibilities
' ----------------
' This module is responsible for:
'
'     - ensuring tblConn exists
'     - validating tblConn configuration
'     - validating DSN presence and DSN connectivity
'     - exposing current Oracle user / role checks
'     - changing the current Oracle user's password
'     - switching application environment / DSN / schema targets
'     - inspecting passthrough queries
'     - reading passthrough query DSNs
'     - bulk-updating passthrough query DSNs
'     - bulk-updating passthrough schema references in SQL text
'     - testing passthrough queries
'     - summarizing current Oracle application configuration
'
'
' Key public helpers
' ------------------
' tblConn setup / validation:
'     OracleAdmin_Ensure_tblConn
'     OracleAdmin_tblConn_Exists
'     OracleAdmin_tblConn_HasRow
'     OracleAdmin_Validate_tblConn
'
' DSN validation:
'     OracleAdmin_Get_ODBC_DSN_Exists
'     OracleAdmin_Validate_CurrentDSNExists
'     OracleAdmin_Validate_DSNConnection
'
' Oracle user / role helpers:
'     OracleAdmin_Get_ODBC_User
'     OracleAdmin_Check_Oracle_User_Role
'     OracleAdmin_Validate_Oracle_User_Role
'
' password helpers:
'     OracleAdmin_ChangeCurrentUserPassword
'
' environment switching:
'     OracleAdmin_SwitchEnvironment
'
' passthrough query inspection:
'     OracleAdmin_IsPassthroughQuery
'     OracleAdmin_Get_PTQ_Count
'     OracleAdmin_Get_PTQ_DSN
'     OracleAdmin_Get_PTQ_Names
'     OracleAdmin_Get_PTQ_Details
'
' passthrough query maintenance:
'     OracleAdmin_Set_PTQ_DSNS
'     OracleAdmin_Set_PTQ_SchemaRefs
'
' passthrough query diagnostics:
'     OracleAdmin_Check_PTQ_DSN_Mismatch
'     OracleAdmin_Test_PTQ
'     OracleAdmin_Test_PTQ_All
'
' config summaries:
'     OracleAdmin_Get_CurrentConfig
'     OracleAdmin_Debug_CurrentConfig
'
' helpers:
'     AddFieldDescription
'
' Typical usage
' -------------
' Ensure tblConn exists:
'
'     OracleAdmin_Ensure_tblConn
'
' Validate application configuration:
'
'     OracleAdmin_Validate_tblConn
'
' Validate a DSN can be reached:
'
'     OracleAdmin_Validate_DSNConnection Get_DB_DSN()

' Switch application environment:
'
'     OracleAdmin_SwitchEnvironment "TEST", "MY_DATA_SOURCE_TEST", "MY_SCHEMA_TEST"
'
' Bulk retarget passthrough queries:
'
'     OracleAdmin_Set_PTQ_DSNS "MY_DATA_SOURCE_TEST"
'
' Update schema references in PTQ SQL:
'
'     OracleAdmin_Set_PTQ_SchemaRefs "OLD_SCHEMA", "NEW_SCHEMA"
'
' Detect DSN mismatches:
'
'     If OracleAdmin_Check_PTQ_DSN_Mismatch(Get_DB_DSN()) Then ...
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
'     - frmLogin
'     - environment switching routines
'     - deployment / admin utilities
'
'
' Return style
' ------------
' This module generally uses:
'
'     - raised errors for validation failures
'     - Collections of dictionaries for inspection / test summaries
'
' It intentionally avoids UI behavior so callers can decide how to present errors
' or results.
'
'
' Design notes
' ------------
' This module should not contain:
'
'     - low-level SQL execution internals
'     - form business logic
'     - linked-table creation / relinking code
'     - user-facing message boxes
'
' Linked-table responsibilities belong in modOracleLinking.
'
'
' Version
' -------
' v1
'
'====================================================================================

Option Compare Database
Option Explicit

Private Const cModuleName As String = "modOracleAdmin"

'------------------------------------------------------------------------------------
' tblConn setup / validation
'------------------------------------------------------------------------------------

Public Sub OracleAdmin_Ensure_tblConn()

    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim rs As DAO.Recordset

    Set db = CurrentDb

    On Error Resume Next
    Set tdf = db.TableDefs("tblConn")
    On Error GoTo 0

    If tdf Is Nothing Then

        Set tdf = db.CreateTableDef("tblConn")

        With tdf.fields
            .Append tdf.CreateField("ENV", dbText, 15)
            .Append tdf.CreateField("DSN", dbText, 15)
            .Append tdf.CreateField("SCHEMA", dbText, 30)
            .Append tdf.CreateField("DSN_INIT", dbText, 15)
        End With

        db.TableDefs.Append tdf
        db.TableDefs.Refresh

        ' Re-open the persisted tabledef, then set field attributes/properties
        Set tdf = db.TableDefs("tblConn")

        tdf.fields("ENV").Required = True
        tdf.fields("DSN").Required = True

        tdf.fields("SCHEMA").Required = False
        tdf.fields("SCHEMA").AllowZeroLength = True

        tdf.fields("DSN_INIT").Required = False
        tdf.fields("DSN_INIT").AllowZeroLength = True

        AddFieldDescription tdf.fields("ENV"), "Environment Name (PROD is default)"
        AddFieldDescription tdf.fields("DSN"), "Default DSN name"
        AddFieldDescription tdf.fields("SCHEMA"), "Optional default schema name"
        AddFieldDescription tdf.fields("DSN_INIT"), "Original DSN value to support custom DSNs on login form if revert is needed"

    End If

    Set rs = db.OpenRecordset("SELECT COUNT(*) AS row_cnt FROM tblConn", dbOpenSnapshot)

    If CLng(rs!row_cnt) = 0 Then
        db.Execute _
            "INSERT INTO tblConn (ENV, DSN, SCHEMA, DSN_INIT) " & _
            "VALUES ('PROD', 'MY_DATA_SOURCE', '', 'MY_DATA_SOURCE')", _
            dbFailOnError
    End If

Cleanup:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set tdf = Nothing
    Set db = Nothing

End Sub

Public Function OracleAdmin_tblConn_Exists() As Boolean

    Dim db As DAO.Database
    Dim tdf As DAO.TableDef

    Set db = CurrentDb

    On Error Resume Next
    Set tdf = db.TableDefs("tblConn")
    On Error GoTo 0

    OracleAdmin_tblConn_Exists = Not tdf Is Nothing

Cleanup:
    Set tdf = Nothing
    Set db = Nothing

End Function

Public Function OracleAdmin_tblConn_HasRow() As Boolean

    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    If Not OracleAdmin_tblConn_Exists() Then Exit Function

    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT COUNT(*) AS row_cnt FROM tblConn", dbOpenSnapshot)

    OracleAdmin_tblConn_HasRow = (CLng(rs!row_cnt) > 0)

Cleanup:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing

End Function

Public Sub OracleAdmin_Validate_tblConn()

    If Not OracleAdmin_tblConn_Exists() Then
        Err.Raise vbObjectError + 2000, cModuleName & ".OracleAdmin_Validate_tblConn", "tblConn does not exist."
    End If

    If Not OracleAdmin_tblConn_HasRow() Then
        Err.Raise vbObjectError + 2001, cModuleName & ".OracleAdmin_Validate_tblConn", "tblConn does not contain a configuration row."
    End If

    If Len(Get_DB_DSN()) = 0 Then
        Err.Raise vbObjectError + 2002, cModuleName & ".OracleAdmin_Validate_tblConn", "tblConn.DSN is blank."
    End If

End Sub

'------------------------------------------------------------------------------------
' DSN helpers
'------------------------------------------------------------------------------------

Public Function OracleAdmin_Get_ODBC_DSN_Exists(ByVal sDSNName As String) As Boolean

    Dim reg As Object
    Dim dsnPath As String
    Dim hKey As Variant

    Set reg = CreateObject("WScript.Shell")

    For Each hKey In Array( _
        "HKEY_CURRENT_USER\Software\ODBC\ODBC.INI\", _
        "HKEY_LOCAL_MACHINE\Software\ODBC\ODBC.INI\", _
        "HKEY_LOCAL_MACHINE\Software\WOW6432Node\ODBC\ODBC.INI\" _
    )
        On Error Resume Next
        dsnPath = CStr(hKey) & sDSNName & "\"
        reg.RegRead dsnPath
        If Err.Number = 0 Then
            OracleAdmin_Get_ODBC_DSN_Exists = True
            Exit Function
        End If
        Err.Clear
        On Error GoTo 0
    Next hKey

End Function

Public Sub OracleAdmin_Validate_CurrentDSNExists()

    Dim sDSN As String

    sDSN = Get_DB_DSN()

    If Len(sDSN) = 0 Then
        Err.Raise vbObjectError + 2010, cModuleName & ".OracleAdmin_Validate_CurrentDSNExists", "Current DSN is blank."
    End If

    If Not OracleAdmin_Get_ODBC_DSN_Exists(sDSN) Then
        Err.Raise vbObjectError + 2011, cModuleName & ".OracleAdmin_Validate_CurrentDSNExists", "Current DSN does not exist in the registry: " & sDSN
    End If

End Sub

Public Sub OracleAdmin_Validate_DSNConnection(ByVal sDSN As String)

    If Len(Trim$(sDSN)) = 0 Then
        Err.Raise vbObjectError + 2012, cModuleName & ".OracleAdmin_Validate_DSNConnection", "DSN cannot be blank."
    End If

    If Not Test_ODBC_Conn(sDSN) Then
        Err.Raise vbObjectError + 2013, cModuleName & ".OracleAdmin_Validate_DSNConnection", "Unable to connect to DSN: " & sDSN
    End If

End Sub

'------------------------------------------------------------------------------------
' Oracle user / role helpers
'------------------------------------------------------------------------------------

Private Function OracleAdmin_QuotedIdentifier(ByVal sIdentifier As String) As String
    OracleAdmin_QuotedIdentifier = """" & Replace$(sIdentifier, """", """""") & """"
End Function

Private Function OracleAdmin_QuotedPasswordToken(ByVal sPassword As String) As String
    OracleAdmin_QuotedPasswordToken = """" & Replace$(sPassword, """", """""") & """"
End Function

Public Function OracleAdmin_Get_ODBC_User() As String
    OracleAdmin_Get_ODBC_User = Get_ODBC_User()
End Function

Public Function OracleAdmin_Check_Oracle_User_Role(ByVal sOracleRoleName As String) As Boolean
    OracleAdmin_Check_Oracle_User_Role = Check_Oracle_User_Role(sOracleRoleName)
End Function

Public Sub OracleAdmin_Validate_Oracle_User_Role(ByVal sOracleRoleName As String)

    If Not Check_Oracle_User_Role(sOracleRoleName) Then
        Err.Raise vbObjectError + 2020, cModuleName & ".OracleAdmin_Validate_Oracle_User_Role", _
                  "Current Oracle user does not have required role: " & UCase$(Trim$(sOracleRoleName))
    End If

End Sub

Public Function OracleAdmin_ChangeCurrentUserPassword( _
    ByVal sDSN As String, _
    ByVal sUserName As String, _
    ByVal sOldPassword As String, _
    ByVal sNewPassword As String _
) As String

    Dim conn As Object
    Dim rs As Object
    Dim sAdoConn As String
    Dim sValidatedOracleUser As String
    Dim sAlterUserSql As String

    On Error GoTo ErrHandler

    sDSN = UCase$(Trim$(sDSN))
    sUserName = UCase$(Trim$(sUserName))

    If Len(sDSN) = 0 Then
        Err.Raise vbObjectError + 2021, cModuleName & ".OracleAdmin_ChangeCurrentUserPassword", "DSN cannot be blank."
    End If

    If Len(sUserName) = 0 Then
        Err.Raise vbObjectError + 2022, cModuleName & ".OracleAdmin_ChangeCurrentUserPassword", "User name cannot be blank."
    End If

    If Len(sOldPassword) = 0 Then
        Err.Raise vbObjectError + 2023, cModuleName & ".OracleAdmin_ChangeCurrentUserPassword", "Current password cannot be blank."
    End If

    If Len(sNewPassword) = 0 Then
        Err.Raise vbObjectError + 2024, cModuleName & ".OracleAdmin_ChangeCurrentUserPassword", "New password cannot be blank."
    End If

    sAdoConn = Get_ADO_Login_Conn_Str(sDSN, sUserName, sOldPassword)

    Set conn = CreateObject("ADODB.Connection")
    conn.Open sAdoConn

    Set rs = conn.Execute("SELECT USER FROM DUAL")

    If rs.EOF Then
        Err.Raise vbObjectError + 2025, cModuleName & ".OracleAdmin_ChangeCurrentUserPassword", _
                  "Unable to validate the Oracle user for the current password."
    End If

    sValidatedOracleUser = UCase$(Nz(rs.fields(0).Value, vbNullString))

    If Len(sValidatedOracleUser) = 0 Then
        Err.Raise vbObjectError + 2026, cModuleName & ".OracleAdmin_ChangeCurrentUserPassword", _
                  "Oracle returned a blank user name during password change."
    End If

    sAlterUserSql = _
        "ALTER USER " & OracleAdmin_QuotedIdentifier(sValidatedOracleUser) & _
        " IDENTIFIED BY " & OracleAdmin_QuotedPasswordToken(sNewPassword) & _
        " REPLACE " & OracleAdmin_QuotedPasswordToken(sOldPassword)

    conn.Execute sAlterUserSql

    On Error Resume Next
    rs.Close
    conn.Close
    On Error GoTo ErrHandler

    Set conn = CreateObject("ADODB.Connection")
    conn.Open Get_ADO_Login_Conn_Str(sDSN, sValidatedOracleUser, sNewPassword)

    OracleAdmin_ChangeCurrentUserPassword = Get_ODBC_Conn_Str(sDSN, sValidatedOracleUser, sNewPassword)

Cleanup:
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State <> 0 Then rs.Close
    End If
    Set rs = Nothing

    If Not conn Is Nothing Then
        If conn.State <> 0 Then conn.Close
    End If
    Set conn = Nothing
    Exit Function

ErrHandler:
    Err.Raise Err.Number, cModuleName & ".OracleAdmin_ChangeCurrentUserPassword", Err.Description

End Function

'------------------------------------------------------------------------------------
' Switch environments
'------------------------------------------------------------------------------------

Public Sub OracleAdmin_SwitchEnvironment( _
    ByVal envName As String, _
    ByVal dsnName As String, _
    Optional ByVal schemaName As String = "", _
    Optional ByVal updatePassthroughQueries As Boolean = True, _
    Optional ByVal updateLinkedTables As Boolean = False, _
    Optional ByVal linkedTableFromSchema As String = "", _
    Optional ByVal linkedTableToSchema As String = "" _
)

    On Error GoTo ErrHandler

    Debug.Print String$(80, "-")
    Debug.Print "OracleAdmin_SwitchEnvironment starting"
    Debug.Print "  Requested ENV: " & envName
    Debug.Print "  Requested DSN: " & dsnName
    Debug.Print "  Requested SCHEMA: " & IIf(Len(Trim$(schemaName)) > 0, schemaName, "<unchanged>")
    Debug.Print "  Update PTQs: " & CStr(updatePassthroughQueries)
    Debug.Print "  Update linked tables: " & CStr(updateLinkedTables)

    envName = UCase$(Trim$(envName))
    dsnName = UCase$(Trim$(dsnName))
    schemaName = UCase$(Trim$(schemaName))
    linkedTableFromSchema = UCase$(Trim$(linkedTableFromSchema))
    linkedTableToSchema = UCase$(Trim$(linkedTableToSchema))

    If Len(envName) = 0 Then
        Err.Raise vbObjectError + 2100, cModuleName & ".OracleAdmin_SwitchEnvironment", _
                  "Environment name cannot be blank."
    End If

    If Len(dsnName) = 0 Then
        Err.Raise vbObjectError + 2101, cModuleName & ".OracleAdmin_SwitchEnvironment", _
                  "DSN name cannot be blank."
    End If

    If Not OracleAdmin_Get_ODBC_DSN_Exists(dsnName) Then
        Err.Raise vbObjectError + 2102, cModuleName & ".OracleAdmin_SwitchEnvironment", _
                  "DSN does not exist on this machine: " & dsnName
    End If

    If Not Test_ODBC_Conn(dsnName) Then
        Err.Raise vbObjectError + 2103, cModuleName & ".OracleAdmin_SwitchEnvironment", _
                  "Unable to connect to DSN: " & dsnName
    End If

    If (Len(linkedTableFromSchema) = 0 Xor Len(linkedTableToSchema) = 0) Then
        Err.Raise vbObjectError + 2104, cModuleName & ".OracleAdmin_SwitchEnvironment", _
                  "linkedTableFromSchema and linkedTableToSchema must both be supplied or both omitted."
    End If

    Set_DB_Env envName
    Debug.Print "  Updated tblConn.ENV -> " & envName

    Set_DB_DSN dsnName
    Debug.Print "  Updated tblConn.DSN -> " & dsnName

    If Len(schemaName) > 0 Then
        Set_DB_Schema schemaName
        Debug.Print "  Updated tblConn.SCHEMA -> " & schemaName
    Else
        Debug.Print "  Left tblConn.SCHEMA unchanged."
    End If

    If updatePassthroughQueries Then
        OracleAdmin_Set_PTQ_DSNS dsnName, False
        Debug.Print "  Updated passthrough query DSNs."
    Else
        Debug.Print "  Skipped passthrough query DSN updates."
    End If

    If updateLinkedTables Then
        Call OracleLink_SetAllLinkedTableConnections( _
            sToDSN:=dsnName, _
            sFromSchema:=linkedTableFromSchema, _
            sToSchema:=linkedTableToSchema)
        Debug.Print "  Updated Oracle ODBC linked tables."
        If Len(linkedTableFromSchema) > 0 Or Len(linkedTableToSchema) > 0 Then
            Debug.Print "  Linked-table schema remap: " & linkedTableFromSchema & " -> " & linkedTableToSchema
        End If
    Else
        Debug.Print "  Skipped linked-table updates."
    End If

    OracleSession_Clear
    Debug.Print "  Cleared runtime Oracle session."
    Debug.Print "OracleAdmin_SwitchEnvironment completed successfully."
    Debug.Print String$(80, "-")

    Exit Sub

ErrHandler:
    Debug.Print "OracleAdmin_SwitchEnvironment failed: " & Err.Number & " - " & Err.Description
    Debug.Print String$(80, "-")
    Err.Raise Err.Number, cModuleName & ".OracleAdmin_SwitchEnvironment", Err.Description

End Sub

'------------------------------------------------------------------------------------
' Passthrough query inspection helpers
'------------------------------------------------------------------------------------

Public Function OracleAdmin_IsPassthroughQuery(ByVal sQueryName As String) As Boolean

    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef

    Set db = CurrentDb

    On Error Resume Next
    Set qdf = db.QueryDefs(sQueryName)
    On Error GoTo 0

    If qdf Is Nothing Then Exit Function

    OracleAdmin_IsPassthroughQuery = (Len(Nz(qdf.Connect, vbNullString)) > 0)

Cleanup:
    Set qdf = Nothing
    Set db = Nothing

End Function

Public Function OracleAdmin_Get_PTQ_Count() As Long

    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef

    Set db = CurrentDb

    For Each qdf In db.QueryDefs
        If Len(Nz(qdf.Connect, vbNullString)) > 0 Then
            OracleAdmin_Get_PTQ_Count = OracleAdmin_Get_PTQ_Count + 1
        End If
    Next qdf

Cleanup:
    Set qdf = Nothing
    Set db = Nothing

End Function

Public Function OracleAdmin_Get_PTQ_DSN(ByVal sQueryName As String) As String

    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim lStartPos As Long
    Dim lEndPos As Long
    Dim sConnect As String

    Set db = CurrentDb

    On Error Resume Next
    Set qdf = db.QueryDefs(sQueryName)
    On Error GoTo 0

    If qdf Is Nothing Then Exit Function

    sConnect = Nz(qdf.Connect, vbNullString)

    If Len(sConnect) = 0 Then Exit Function

    lStartPos = InStr(1, sConnect, "DSN=", vbTextCompare)

    If lStartPos > 0 Then
        lEndPos = InStr(lStartPos, sConnect & ";", ";")
        If lEndPos > 0 Then
            OracleAdmin_Get_PTQ_DSN = UCase$(Mid$(sConnect, lStartPos + 4, lEndPos - lStartPos - 4))
        End If
    End If

Cleanup:
    Set qdf = Nothing
    Set db = Nothing

End Function

Public Function OracleAdmin_Get_PTQ_Names() As Collection

    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim result As Collection

    Set result = New Collection
    Set db = CurrentDb

    For Each qdf In db.QueryDefs
        If Len(Nz(qdf.Connect, vbNullString)) > 0 Then
            result.Add qdf.Name
        End If
    Next qdf

    Set OracleAdmin_Get_PTQ_Names = result

Cleanup:
    Set qdf = Nothing
    Set db = Nothing

End Function

Public Function OracleAdmin_Get_PTQ_Details() As Collection

    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim result As Collection
    Dim item As Object

    Set result = New Collection
    Set db = CurrentDb

    For Each qdf In db.QueryDefs
        If Len(Nz(qdf.Connect, vbNullString)) > 0 Then
            Set item = CreateObject("Scripting.Dictionary")
            item.Add "QueryName", qdf.Name
            item.Add "DSN", OracleAdmin_Get_PTQ_DSN(qdf.Name)
            item.Add "ReturnsRecords", qdf.returnsRecords
            item.Add "ODBCTimeout", qdf.ODBCTimeout
            item.Add "SQL", qdf.sql
            result.Add item
        End If
    Next qdf

    Set OracleAdmin_Get_PTQ_Details = result

Cleanup:
    Set item = Nothing
    Set qdf = Nothing
    Set db = Nothing

End Function

Public Sub OracleAdmin_Set_PTQ_DSNS( _
    ByVal sToDSN As String, _
    Optional ByVal includeSystemQueries As Boolean = False _
)

    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim sNewConnect As String

    sToDSN = UCase$(Trim$(sToDSN))

    If Len(sToDSN) = 0 Then
        Err.Raise vbObjectError + 2030, cModuleName & ".OracleAdmin_Set_PTQ_DSNS", "Target DSN cannot be blank."
    End If

    Set db = CurrentDb
    sNewConnect = "ODBC;DSN=" & sToDSN

    For Each qdf In db.QueryDefs

        If Len(Nz(qdf.Connect, vbNullString)) > 0 Then
            If includeSystemQueries Or Left$(qdf.Name, 1) <> "~" Then
                qdf.Connect = sNewConnect
            End If
        End If

    Next qdf

Cleanup:
    Set qdf = Nothing
    Set db = Nothing

End Sub

Public Sub OracleAdmin_Set_PTQ_SchemaRefs( _
    ByVal sFromSchema As String, _
    ByVal sToSchema As String, _
    Optional ByVal includeSystemQueries As Boolean = False _
)

    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim sStartSQL As String
    Dim sUpdatedSQL As String
    Dim sFromToken As String
    Dim sToToken As String

    sFromSchema = UCase$(Trim$(sFromSchema))
    sToSchema = UCase$(Trim$(sToSchema))

    If Len(sFromSchema) = 0 Or Len(sToSchema) = 0 Then
        Err.Raise vbObjectError + 2031, cModuleName & ".OracleAdmin_Set_PTQ_SchemaRefs", "Both source and target schema names are required."
    End If

    If sFromSchema = sToSchema Then Exit Sub

    sFromToken = sFromSchema & "."
    sToToken = sToSchema & "."

    Set db = CurrentDb

    For Each qdf In db.QueryDefs

        If Len(Nz(qdf.Connect, vbNullString)) > 0 Then
            If includeSystemQueries Or Left$(qdf.Name, 1) <> "~" Then
                sStartSQL = qdf.sql
                sUpdatedSQL = Replace$(sStartSQL, sFromToken, sToToken, , , vbTextCompare)

                If sUpdatedSQL <> sStartSQL Then
                    qdf.sql = sUpdatedSQL
                End If
            End If
        End If

    Next qdf

Cleanup:
    Set qdf = Nothing
    Set db = Nothing

End Sub

Public Function OracleAdmin_Check_PTQ_DSN_Mismatch( _
    Optional ByVal sDSNCheckValue As String = "" _
) As Boolean

    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim sCheck As String

    If Len(Trim$(sDSNCheckValue)) = 0 Then
        sCheck = Get_DB_DSN()
    Else
        sCheck = UCase$(Trim$(sDSNCheckValue))
    End If

    Set db = CurrentDb

    For Each qdf In db.QueryDefs
        If Len(Nz(qdf.Connect, vbNullString)) > 0 Then
            If OracleAdmin_Get_PTQ_DSN(qdf.Name) <> sCheck Then
                OracleAdmin_Check_PTQ_DSN_Mismatch = True
                Exit Function
            End If
        End If
    Next qdf

Cleanup:
    Set qdf = Nothing
    Set db = Nothing

End Function

Public Function OracleAdmin_Test_PTQ(ByVal sQueryName As String) As Boolean

    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim rs As DAO.Recordset

    Set db = CurrentDb

    On Error Resume Next
    Set qdf = db.QueryDefs(sQueryName)
    On Error GoTo HandleErr

    If qdf Is Nothing Then
        Err.Raise vbObjectError + 2040, cModuleName & ".OracleAdmin_Test_PTQ", "Query not found: " & sQueryName
    End If

    If Len(Nz(qdf.Connect, vbNullString)) = 0 Then
        Err.Raise vbObjectError + 2041, cModuleName & ".OracleAdmin_Test_PTQ", "Query is not a passthrough query: " & sQueryName
    End If

    If Not qdf.returnsRecords Then
        OracleAdmin_Test_PTQ = True
        GoTo Cleanup
    End If

    Set rs = qdf.OpenRecordset(dbOpenSnapshot)
    OracleAdmin_Test_PTQ = True

Cleanup:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set qdf = Nothing
    Set db = Nothing
    Exit Function

HandleErr:
    Err.Raise vbObjectError + 2042, cModuleName & ".OracleAdmin_Test_PTQ", _
              "Passthrough query test failed for " & sQueryName & ". Details: " & Err.Description

End Function

Public Function OracleAdmin_Test_PTQ_All( _
    Optional ByVal includeTempQueries As Boolean = False _
) As Collection

    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim results As Collection
    Dim item As Object
    Dim shouldTest As Boolean

    Set results = New Collection
    Set db = CurrentDb

    For Each qdf In db.QueryDefs

        shouldTest = (Len(Nz(qdf.Connect, vbNullString)) > 0)

        If shouldTest Then
            If Not includeTempQueries Then
                If Left$(qdf.Name, 4) = "~TMP" Then shouldTest = False
            End If
        End If

        If shouldTest Then
            Set item = CreateObject("Scripting.Dictionary")
            item.Add "QueryName", qdf.Name

            On Error Resume Next
            Call OracleAdmin_Test_PTQ(qdf.Name)
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

    Next qdf

    Set OracleAdmin_Test_PTQ_All = results

Cleanup:
    Set item = Nothing
    Set qdf = Nothing
    Set db = Nothing

End Function

'------------------------------------------------------------------------------------
' Utility summaries
'------------------------------------------------------------------------------------

Public Function OracleAdmin_Get_CurrentConfig() As Object

    Dim d As Object

    Set d = CreateObject("Scripting.Dictionary")

    d.Add "ENV", Get_DB_Env()
    d.Add "DSN", Get_DB_DSN()
    d.Add "SCHEMA", Get_DB_Schema()
    d.Add "DSN_INIT", Get_DB_DSN_Init()
    d.Add "ODBC_USER", Get_ODBC_User()

    Set OracleAdmin_Get_CurrentConfig = d

End Function

Public Function OracleAdmin_Debug_CurrentConfig() As String

    OracleAdmin_Debug_CurrentConfig = _
        "ENV=" & Get_DB_Env() & _
        "; DSN=" & Get_DB_DSN() & _
        "; SCHEMA=" & Get_DB_Schema() & _
        "; DSN_INIT=" & Get_DB_DSN_Init() & _
        "; ODBC_USER=" & Get_ODBC_User()

End Function

'------------------------------------------------------------------------------------
' Helpers
'------------------------------------------------------------------------------------

Private Sub AddFieldDescription(ByRef fld As DAO.Field, ByVal descriptionText As String)

    Dim prop As DAO.Property

    On Error Resume Next
    fld.Properties.Delete "Description"
    On Error GoTo 0

    Set prop = fld.CreateProperty("Description", dbText, descriptionText)
    fld.Properties.Append prop

    Set prop = Nothing

End Sub
