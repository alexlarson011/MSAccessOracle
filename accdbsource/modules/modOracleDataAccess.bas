Attribute VB_Name = "modOracleDataAccess"
'====================================================================================
' modOracleDataAccess
'====================================================================================
'
' README
' ------
' Purpose
' -------
' Provides the core Oracle data-access layer for this Access application.
'
' This module is the primary dependency for Oracle query execution and is designed
' for a passthrough-first architecture using ADO / DAO and ODBC connection strings,
' without requiring bound forms or persistent Oracle linked tables.
'
'
' Responsibilities
' ----------------
' This module is responsible for:
'
'     - reading and updating tblConn configuration values
'     - building DSN-based Oracle ODBC connection strings
'     - exposing a fully parameterized Oracle ODBC connection-string builder
'     - preferring the runtime Oracle session connection string when available
'     - executing passthrough scalar queries
'     - executing passthrough action SQL
'     - materializing query results into dictionaries / collections
'     - materializing query results into case-insensitive dictionaries / collections
'     - exposing typed scalar helpers
'     - exposing Oracle helpers such as sequence NEXTVAL access
'     - exposing SQL literal helper functions
'
'
' Key public helpers
' ------------------
' tblConn getters / setters:
'     Get_DB_Env
'     Get_DB_DSN
'     Get_DB_DSN_Init
'     Get_DB_Schema
'     Set_DB_Env
'     Set_DB_DSN
'     Set_DB_DSN_Init
'     Set_DB_Schema
'
' runtime session helpers:
'     Get_Runtime_ODBC_Conn_Str
'     RequireOracleSession
'
' connection helpers:
'     Build_Oracle_ODBC_Conn_Str
'     Get_ODBC_Conn_Str
'     Get_ADO_Login_Conn_Str
'     Test_ODBC_Conn
'
' passthrough scalar helpers:
'     PTQ_Select
'     PTQ_SelectString
'     PTQ_SelectLong
'     PTQ_SelectDouble
'     PTQ_SelectDate
'     PTQ_Exists
'
' passthrough execution helpers:
'     PTQ_Execute
'     PTQ_Run_Proc
'
' row materialization helpers:
'     PTQ_GetRow
'     PTQ_GetRows
'     PTQ_Rs (retired; raises an error if called)
'
' Oracle helpers:
'     Get_ODBC_User
'     Check_Oracle_User_Role
'     Oracle_GetNextSequenceValue
'
' SQL literal helpers:
'     SqlTextLiteral
'     SqlStringOrNull
'     SqlNumberOrNull
'     SqlDateOrNull
'     SqlTimestampOrNull
'     SqlBooleanNumber
'     SqlYNFlag
'
'
' Typical usage
' -------------
' Scalar query:
'
'     v = PTQ_Select("SELECT COUNT(*) FROM my_table")
'
' Typed scalar query:
'
'     nextId = PTQ_SelectLong("SELECT my_seq.NEXTVAL FROM dual")
'
' Action SQL:
'
'     PTQ_Execute "UPDATE my_table SET col = 'X' WHERE id = 1"
'
' Full connection-string build:
'
'     sConn = Build_Oracle_ODBC_Conn_Str( _
'         sDSN:="MY_DATA_SOURCE", _
'         sUserName:="scott", _
'         sPassword:="tiger", _
'         lFBS:=128000, _
'         lTSZ:=16384)
'
' Single-row retrieval:
'
'     Set rowData = PTQ_GetRow("SELECT col1, col2 FROM my_table WHERE id = 1")
'
' Returned row dictionaries use case-insensitive key lookup, which makes them
' friendlier for aliased SQL and form-engine read models.
'
' Joined SQL should alias duplicate column names. PTQ_GetRows / PTQ_GetRow raise a
' clear error if Oracle returns the same column name more than once.
'
' Sequence helper:
'
'     nextId = Oracle_GetNextSequenceValue(Get_DB_Schema(), "my_seq")
'
' Most query helpers place SQL or the primary inputs first and accept DSN as an
' optional final argument. If omitted, DSN defaults to Get_DB_DSN().
'
'
' Runtime connection behavior
' ---------------------------
' This module supports two connection modes:
'
' 1. DSN mode
'    Used before login or when no runtime Oracle session exists.
'
' 2. Runtime session mode
'    Used after login when modOracleSession contains a credentialed connection
'    string for the current user.
'
' When a runtime session exists, query helpers prefer that credentialed connection
' string over a DSN-only connection.
'
' Runtime session execution uses direct ADO connections for scalar and action helpers
' so re-login flows do not accidentally reuse stale Oracle user state inside the
' Access/DAO ODBC layer.
'
'
' Important architecture note
' ---------------------------
' This module does NOT keep a persistent Oracle connection open.
'
' Pre-login / DSN-only connectivity helpers still use temporary DAO passthrough
' QueryDefs.
'
' Logged-in runtime scalar/action helpers open short-lived ADO connections directly
' from the stored runtime connection string.
'
' Row materialization helpers use fresh isolated DAO passthrough QueryDefs so combo /
' list style result loading stays compatible with Oracle ODBC behavior in Access.
'
' DAO passthrough work uses a fresh isolated workspace per call to reduce stale
' ODBC-session reuse inside Access.
'
'
' Dependencies
' ------------
' Depends on:
'
'     - modOracleSession
'     - tblConn existing in the local Access database
'
'
' Common callers
' --------------
' Common callers include:
'
'     - frmLogin
'     - modOracleAdmin
'     - modOracleLinking
'     - modOracleFormEngine
'
'
' Design notes
' ------------
' This module should remain focused on runtime data access.
'
' It should not contain:
'
'     - message boxes
'     - form UI logic
'     - linked-table relinking logic
'     - business validation rules
'
' Row materialization is intentionally generic so callers can use aliased SELECT
' lists without depending on exact Oracle column-name casing.
'
' Get_ODBC_Conn_Str remains the compatibility wrapper that uses the builder defaults.
'
' PTQ_Rs is intentionally retired because returning a DAO.Recordset from a temporary
' passthrough object is not a stable contract for this architecture.
'
'
' Version
' -------
' v1
'
'====================================================================================

Option Compare Database
Option Explicit

Private Const cModuleName As String = "modOracleDataAccess"
Private Const cAdoConnectionTimeoutSeconds As Long = 15
Private Const cAdoQueryTimeoutSeconds As Long = 60

Private Function OracleBoolSetting(ByVal bValue As Boolean) As String
    If bValue Then
        OracleBoolSetting = "T"
    Else
        OracleBoolSetting = "F"
    End If
End Function

Private Sub AppendConnPart(ByRef sConn As String, ByVal sKey As String, ByVal sValue As String)
    sConn = sConn & sKey & "=" & sValue & ";"
End Sub

Private Function OpenIsolatedCurrentDb(ByRef ws As DAO.Workspace) As DAO.Database

    Dim dbCurrent As DAO.Database
    Dim sWorkspaceName As String

    Set dbCurrent = CurrentDb
    sWorkspaceName = "wsPTQ_" & Replace$(Format$(Timer, "0.000"), ".", "_")

    Set ws = DBEngine.CreateWorkspace(sWorkspaceName, "admin", vbNullString, dbUseJet)
    ws.IsolateODBCTrans = True
    Set OpenIsolatedCurrentDb = ws.OpenDatabase(dbCurrent.Name)

Cleanup:
    Set dbCurrent = Nothing

End Function

Private Function Get_Runtime_ADO_Conn_Str(Optional ByVal sDSN As String = "") As String

    Dim sConn As String

    sConn = Get_Runtime_ODBC_Conn_Str(sDSN)

    If Left$(sConn, 5) = "ODBC;" Then
        sConn = Mid$(sConn, 6)
    End If

    Get_Runtime_ADO_Conn_Str = sConn

End Function

Private Function ResolveDefaultDSN(Optional ByVal sDSN As String = "") As String

    sDSN = Trim$(sDSN)

    If Len(sDSN) = 0 Then
        sDSN = Get_DB_DSN()
    End If

    ResolveDefaultDSN = sDSN

End Function

Private Function CreateAdoConnection(ByVal sDSN As String) As Object

    Dim conn As Object

    Set conn = CreateObject("ADODB.Connection")
    conn.ConnectionTimeout = cAdoConnectionTimeoutSeconds
    conn.CommandTimeout = cAdoQueryTimeoutSeconds
    conn.Open Get_Runtime_ADO_Conn_Str(sDSN)

    Set CreateAdoConnection = conn

End Function

Private Function OpenAdoRecordset( _
    ByVal conn As Object, _
    ByVal sSQL As String, _
    Optional ByVal timeoutSeconds As Long = cAdoQueryTimeoutSeconds _
) As Object

    Dim cmd As Object
    Dim rs As Object

    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = conn
    cmd.CommandType = 1
    cmd.CommandText = sSQL
    cmd.CommandTimeout = timeoutSeconds

    Set rs = CreateObject("ADODB.Recordset")
    rs.Open cmd, , 0, 1

    Set OpenAdoRecordset = rs

End Function

Private Function PTQ_SelectAdo(ByVal sDSN As String, ByVal sSQL As String) As Variant

    Dim conn As Object
    Dim rs As Object

    On Error GoTo HandleErr

    Set conn = CreateAdoConnection(sDSN)
    Set rs = OpenAdoRecordset(conn, sSQL)

    If rs.EOF Then
        PTQ_SelectAdo = Null
    Else
        PTQ_SelectAdo = rs.Fields(0).Value
    End If

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

HandleErr:
    Err.Raise _
        vbObjectError + 1040, _
        cModuleName & ".PTQ_SelectAdo", _
        "ADO scalar query failed." & vbCrLf & _
        "DSN: " & sDSN & vbCrLf & _
        "SQL: " & sSQL & vbCrLf & _
        "Details: " & Err.Description

End Function

Private Sub PTQ_ExecuteAdo( _
    ByVal sDSN As String, _
    ByVal sSQL As String, _
    Optional ByVal timeoutSeconds As Long = 60 _
)

    Dim conn As Object

    On Error GoTo HandleErr

    Set conn = CreateAdoConnection(sDSN)
    conn.CommandTimeout = timeoutSeconds
    conn.Execute sSQL

Cleanup:
    On Error Resume Next
    If Not conn Is Nothing Then
        If conn.State <> 0 Then conn.Close
    End If
    Set conn = Nothing
    Exit Sub

HandleErr:
    Err.Raise _
        vbObjectError + 1041, _
        cModuleName & ".PTQ_ExecuteAdo", _
        "ADO execute failed." & vbCrLf & _
        "DSN: " & sDSN & vbCrLf & _
        "SQL: " & sSQL & vbCrLf & _
        "Details: " & Err.Description

End Sub

Private Function PTQ_GetRowsAdo( _
    ByVal sDSN As String, _
    ByVal sSQL As String _
) As Collection

    Dim conn As Object
    Dim rs As Object
    Dim rows As Collection
    Dim rowDict As Object
    Dim vData As Variant
    Dim lFieldIndex As Long
    Dim lRowIndex As Long

    On Error GoTo HandleErr

    Set rows = New Collection
    Set conn = CreateAdoConnection(sDSN)
    Set rs = OpenAdoRecordset(conn, sSQL)

    If Not rs.EOF Then
        vData = rs.GetRows

        For lRowIndex = LBound(vData, 2) To UBound(vData, 2)
            Set rowDict = CreateObject("Scripting.Dictionary")
            rowDict.CompareMode = vbTextCompare

            For lFieldIndex = LBound(vData, 1) To UBound(vData, 1)
                AddRowDictValue rowDict, rs.Fields(lFieldIndex).Name, vData(lFieldIndex, lRowIndex), cModuleName & ".PTQ_GetRowsAdo"
            Next lFieldIndex

            rows.Add rowDict
        Next lRowIndex
    End If

    Set PTQ_GetRowsAdo = rows

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

HandleErr:
    Err.Raise _
        vbObjectError + 1042, _
        cModuleName & ".PTQ_GetRowsAdo", _
        "ADO row retrieval failed." & vbCrLf & _
        "DSN: " & sDSN & vbCrLf & _
        "SQL: " & sSQL & vbCrLf & _
        "Details: " & Err.Description

End Function

Private Sub AddRowDictValue( _
    ByRef rowDict As Object, _
    ByVal fieldName As String, _
    ByVal fieldValue As Variant, _
    ByVal sourceProcName As String _
)

    If rowDict.Exists(fieldName) Then
        Err.Raise vbObjectError + 1043, sourceProcName, _
                  "Query returned duplicate column name '" & fieldName & "'. " & _
                  "Alias duplicate columns so every returned field name is unique."
    End If

    rowDict.Add fieldName, fieldValue

End Sub

'------------------------------------------------------------------------------------
' tblConn getters / setters
'------------------------------------------------------------------------------------

Public Function Get_DB_Env() As String
    Get_DB_Env = Get_tblConn_Value("ENV", "")
End Function

Public Function Get_DB_DSN() As String
    Get_DB_DSN = Get_tblConn_Value("DSN", "")
End Function

Public Function Get_DB_DSN_Init() As String
    Get_DB_DSN_Init = Get_tblConn_Value("DSN_INIT", "")
End Function

Public Function Get_DB_Schema() As String
    Get_DB_Schema = Get_tblConn_Value("SCHEMA", "")
End Function

Public Sub Set_DB_Env(ByVal sToEnv As String)

    sToEnv = UCase$(Trim$(sToEnv))

    If Len(sToEnv) = 0 Then
        Err.Raise vbObjectError + 1000, cModuleName & ".Set_DB_Env", "Environment cannot be blank."
    End If

    CurrentDb.Execute _
        "UPDATE tblConn SET ENV = '" & SqlTextLiteral(sToEnv) & "'", _
        dbFailOnError

End Sub

Public Sub Set_DB_DSN(ByVal sToDSN As String)

    sToDSN = UCase$(Trim$(sToDSN))

    If Len(sToDSN) = 0 Then
        Err.Raise vbObjectError + 1001, cModuleName & ".Set_DB_DSN", "DSN cannot be blank."
    End If

    CurrentDb.Execute _
        "UPDATE tblConn SET DSN = '" & SqlTextLiteral(sToDSN) & "'", _
        dbFailOnError

End Sub

Public Sub Set_DB_DSN_Init(ByVal sToDSN As String)

    sToDSN = UCase$(Trim$(sToDSN))

    If Len(sToDSN) = 0 Then
        Err.Raise vbObjectError + 1002, cModuleName & ".Set_DB_DSN_Init", "Initial DSN cannot be blank."
    End If

    CurrentDb.Execute _
        "UPDATE tblConn SET DSN_INIT = '" & SqlTextLiteral(sToDSN) & "'", _
        dbFailOnError

End Sub

Public Sub Set_DB_Schema(ByVal sToSchema As String)

    sToSchema = UCase$(Trim$(sToSchema))

    If Len(sToSchema) = 0 Then
        CurrentDb.Execute "UPDATE tblConn SET SCHEMA = Null", dbFailOnError
    Else
        CurrentDb.Execute _
            "UPDATE tblConn SET SCHEMA = '" & SqlTextLiteral(sToSchema) & "'", _
            dbFailOnError
    End If

End Sub

Private Function Get_tblConn_Value(ByVal fieldName As String, ByVal DefaultValue As String) As String

    Dim v As Variant

    On Error GoTo HandleErr

    v = DLookup(fieldName, "tblConn")

    If IsNull(v) Then
        Get_tblConn_Value = DefaultValue
    Else
        Get_tblConn_Value = CStr(v)
    End If

    Exit Function

HandleErr:
    Get_tblConn_Value = DefaultValue

End Function

'------------------------------------------------------------------------------------
' Runtime session helpers
'------------------------------------------------------------------------------------

Public Function Get_Runtime_ODBC_Conn_Str(Optional ByVal sDSN As String = "") As String

    If OracleSession_IsConnected() Then
        Get_Runtime_ODBC_Conn_Str = g_OracleConnectionString
    Else
        Get_Runtime_ODBC_Conn_Str = Get_ODBC_Conn_Str(sDSN)
    End If

End Function

Public Sub RequireOracleSession()

    If Not OracleSession_IsConnected() Then
        Err.Raise vbObjectError + 1035, cModuleName & ".RequireOracleSession", _
                  "No Oracle runtime session is active. Please log in."
    End If

End Sub

'------------------------------------------------------------------------------------
' Connection string helpers
'------------------------------------------------------------------------------------

Public Function Build_Oracle_ODBC_Conn_Str( _
    Optional ByVal sDSN As String = "", Optional ByVal sUserName As String = "", _
    Optional ByVal sPassword As String = "", Optional ByVal sDBA As String = "W", _
    Optional ByVal bAPA As Boolean = True, Optional ByVal bEXC As Boolean = False, _
    Optional ByVal bFEN As Boolean = True, Optional ByVal bQTO As Boolean = True, _
    Optional ByVal lFRC As Long = 10, Optional ByVal lFDL As Long = 10, _
    Optional ByVal bLOB As Boolean = True, Optional ByVal bRST As Boolean = True, _
    Optional ByVal bBTD As Boolean = False, Optional ByVal bBNF As Boolean = False, _
    Optional ByVal sBAM As String = "IfAllSuccessful", Optional ByVal sNUM As String = "NLS", _
    Optional ByVal bDPM As Boolean = False, Optional ByVal bMTS As Boolean = True, _
    Optional ByVal bMDI As Boolean = False, Optional ByVal bCSR As Boolean = False, _
    Optional ByVal bFWC As Boolean = False, Optional ByVal lFBS As Long = 64000, _
    Optional ByVal sTLO As String = "O", Optional ByVal lMLD As Long = 0, _
    Optional ByVal bODA As Boolean = False, Optional ByVal bSTE As Boolean = False, _
    Optional ByVal lTSZ As Long = 8192, Optional ByVal sAST As String = "FLOAT", _
    Optional ByVal lLPS As Long = 8192 _
) As String

    Dim sConn As String

    If Len(sDSN) = 0 Then sDSN = Get_DB_DSN()

    If Len(sDSN) = 0 Then
        Err.Raise vbObjectError + 1010, cModuleName & ".Build_Oracle_ODBC_Conn_Str", _
                  "No DSN was supplied and tblConn.DSN is blank."
    End If

    If (Len(sUserName) = 0 Xor Len(sPassword) = 0) Then
        Err.Raise vbObjectError + 1011, cModuleName & ".Build_Oracle_ODBC_Conn_Str", _
                  "User name and password must both be supplied or both omitted."
    End If

    sConn = "ODBC;"
    sConn = sConn & "DSN=" & sDSN & ";"

    If Len(sUserName) > 0 Then
        Call AppendConnPart(sConn, "Uid", sUserName)
        Call AppendConnPart(sConn, "Pwd", sPassword)
    End If

    Call AppendConnPart(sConn, "DBQ", sDSN)
    Call AppendConnPart(sConn, "DBA", sDBA)
    Call AppendConnPart(sConn, "APA", OracleBoolSetting(bAPA))
    Call AppendConnPart(sConn, "EXC", OracleBoolSetting(bEXC))
    Call AppendConnPart(sConn, "FEN", OracleBoolSetting(bFEN))
    Call AppendConnPart(sConn, "QTO", OracleBoolSetting(bQTO))
    Call AppendConnPart(sConn, "FRC", CStr(lFRC))
    Call AppendConnPart(sConn, "FDL", CStr(lFDL))
    Call AppendConnPart(sConn, "LOB", OracleBoolSetting(bLOB))
    Call AppendConnPart(sConn, "RST", OracleBoolSetting(bRST))
    Call AppendConnPart(sConn, "BTD", OracleBoolSetting(bBTD))
    Call AppendConnPart(sConn, "BNF", OracleBoolSetting(bBNF))
    Call AppendConnPart(sConn, "BAM", sBAM)
    Call AppendConnPart(sConn, "NUM", sNUM)
    Call AppendConnPart(sConn, "DPM", OracleBoolSetting(bDPM))
    Call AppendConnPart(sConn, "MTS", OracleBoolSetting(bMTS))
    Call AppendConnPart(sConn, "MDI", OracleBoolSetting(bMDI))
    Call AppendConnPart(sConn, "CSR", OracleBoolSetting(bCSR))
    Call AppendConnPart(sConn, "FWC", OracleBoolSetting(bFWC))
    Call AppendConnPart(sConn, "FBS", CStr(lFBS))
    Call AppendConnPart(sConn, "TLO", sTLO)
    Call AppendConnPart(sConn, "MLD", CStr(lMLD))
    Call AppendConnPart(sConn, "ODA", OracleBoolSetting(bODA))
    Call AppendConnPart(sConn, "STE", OracleBoolSetting(bSTE))
    Call AppendConnPart(sConn, "TSZ", CStr(lTSZ))
    Call AppendConnPart(sConn, "AST", sAST)
    Call AppendConnPart(sConn, "LPS", CStr(lLPS))

    Build_Oracle_ODBC_Conn_Str = sConn

End Function

Public Function Get_ODBC_Conn_Str( _
    Optional ByVal sDSN As String = "", _
    Optional ByVal sUserName As String = "", _
    Optional ByVal sPassword As String = "" _
) As String

    Get_ODBC_Conn_Str = Build_Oracle_ODBC_Conn_Str(sDSN, sUserName, sPassword)

End Function

Public Function Get_ADO_Login_Conn_Str( _
    Optional ByVal sDSN As String = "", _
    Optional ByVal sUserName As String = "", _
    Optional ByVal sPassword As String = "" _
) As String

    Dim sConn As String

    sConn = Get_ODBC_Conn_Str(sDSN, sUserName, sPassword)

    If Left$(sConn, 5) = "ODBC;" Then
        sConn = Mid$(sConn, 6)
    End If

    Get_ADO_Login_Conn_Str = sConn

End Function

Public Function Test_ODBC_Conn(Optional ByVal sDSN As String = "") As Boolean

    Dim ws As DAO.Workspace
    Dim db As DAO.Database
    Dim qdfTemp As DAO.QueryDef
    Dim rs As DAO.Recordset

    On Error GoTo HandleErr

    sDSN = ResolveDefaultDSN(sDSN)

    Set db = OpenIsolatedCurrentDb(ws)
    Set qdfTemp = db.CreateQueryDef(vbNullString)

    With qdfTemp
        .Connect = "ODBC;DSN=" & sDSN
        .returnsRecords = True
        .ODBCTimeout = 5
        .sql = "SELECT 1 FROM DUAL"
    End With

    Set rs = qdfTemp.OpenRecordset(dbOpenSnapshot)
    Test_ODBC_Conn = Not (rs.BOF And rs.EOF)

Cleanup:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set qdfTemp = Nothing
    If Not db Is Nothing Then db.Close
    Set db = Nothing
    If Not ws Is Nothing Then ws.Close
    Set ws = Nothing
    Exit Function

HandleErr:
    Test_ODBC_Conn = False
    Resume Cleanup

End Function

'------------------------------------------------------------------------------------
' Internal passthrough querydef helper
'------------------------------------------------------------------------------------

Private Function CreatePassthroughQueryDef( _
    ByVal db As DAO.Database, _
    ByVal sDSN As String, _
    ByVal sSQL As String, _
    ByVal returnsRecords As Boolean, _
    Optional ByVal timeoutSeconds As Long = 60 _
) As DAO.QueryDef

    Dim qdfTemp As DAO.QueryDef
    Dim sConnect As String

    If Len(Trim$(sSQL)) = 0 Then
        Err.Raise vbObjectError + 1021, cModuleName & ".CreatePassthroughQueryDef", "SQL cannot be blank."
    End If

    sConnect = Get_Runtime_ODBC_Conn_Str(sDSN)

    If Len(Trim$(sConnect)) = 0 Then
        Err.Raise vbObjectError + 1020, cModuleName & ".CreatePassthroughQueryDef", _
                  "No runtime connection string or DSN-based connection string could be resolved."
    End If

    Set qdfTemp = db.CreateQueryDef(vbNullString)

    With qdfTemp
        .Connect = sConnect
        .returnsRecords = returnsRecords
        .ODBCTimeout = timeoutSeconds
        .sql = sSQL
    End With

    Set CreatePassthroughQueryDef = qdfTemp

End Function

'------------------------------------------------------------------------------------
' Passthrough scalar helpers
'------------------------------------------------------------------------------------

Public Function PTQ_Select(ByVal sSQL As String, Optional ByVal sDSN As String = "") As Variant

    Dim ws As DAO.Workspace
    Dim db As DAO.Database
    Dim qdfTemp As DAO.QueryDef
    Dim rs As DAO.Recordset

    On Error GoTo HandleErr

    sDSN = ResolveDefaultDSN(sDSN)

    If OracleSession_IsConnected() Then
        PTQ_Select = PTQ_SelectAdo(sDSN, sSQL)
        Exit Function
    End If

    Set db = OpenIsolatedCurrentDb(ws)
    Set qdfTemp = CreatePassthroughQueryDef(db, sDSN, sSQL, True)
    Set rs = qdfTemp.OpenRecordset(dbOpenSnapshot)

    If rs.BOF And rs.EOF Then
        PTQ_Select = Null
    Else
        PTQ_Select = rs.fields(0).Value
    End If

Cleanup:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set qdfTemp = Nothing
    If Not db Is Nothing Then db.Close
    Set db = Nothing
    If Not ws Is Nothing Then ws.Close
    Set ws = Nothing
    Exit Function

HandleErr:
    Err.Raise _
        vbObjectError + 1022, _
        cModuleName & ".PTQ_Select", _
        "Passthrough scalar query failed." & vbCrLf & _
        "DSN: " & sDSN & vbCrLf & _
        "SQL: " & sSQL & vbCrLf & _
        "Details: " & Err.Description

End Function

Public Function PTQ_SelectString(ByVal sSQL As String, Optional ByVal sDSN As String = "") As String

    Dim v As Variant

    v = PTQ_Select(sSQL, sDSN)

    If IsNull(v) Then
        PTQ_SelectString = vbNullString
    Else
        PTQ_SelectString = CStr(v)
    End If

End Function

Public Function PTQ_SelectLong(ByVal sSQL As String, Optional ByVal sDSN As String = "") As Long

    Dim v As Variant

    v = PTQ_Select(sSQL, sDSN)

    If IsNull(v) Then
        Err.Raise vbObjectError + 1023, cModuleName & ".PTQ_SelectLong", "Scalar query returned Null."
    End If

    PTQ_SelectLong = CLng(v)

End Function

Public Function PTQ_SelectDouble(ByVal sSQL As String, Optional ByVal sDSN As String = "") As Double

    Dim v As Variant

    v = PTQ_Select(sSQL, sDSN)

    If IsNull(v) Then
        Err.Raise vbObjectError + 1024, cModuleName & ".PTQ_SelectDouble", "Scalar query returned Null."
    End If

    PTQ_SelectDouble = CDbl(v)

End Function

Public Function PTQ_SelectDate(ByVal sSQL As String, Optional ByVal sDSN As String = "") As Date

    Dim v As Variant

    v = PTQ_Select(sSQL, sDSN)

    If IsNull(v) Then
        Err.Raise vbObjectError + 1025, cModuleName & ".PTQ_SelectDate", "Scalar query returned Null."
    End If

    PTQ_SelectDate = CDate(v)

End Function

Public Function PTQ_Exists(ByVal sSQL As String, Optional ByVal sDSN As String = "") As Boolean
    PTQ_Exists = Not IsNull(PTQ_Select(sSQL, sDSN))
End Function

'------------------------------------------------------------------------------------
' Passthrough action / procedure helpers
'------------------------------------------------------------------------------------

Public Sub PTQ_Execute( _
    ByVal sSQL As String, _
    Optional ByVal timeoutSeconds As Long = 60, _
    Optional ByVal sDSN As String = "" _
)

    Dim ws As DAO.Workspace
    Dim db As DAO.Database
    Dim qdfTemp As DAO.QueryDef

    On Error GoTo HandleErr

    sDSN = ResolveDefaultDSN(sDSN)

    If OracleSession_IsConnected() Then
        PTQ_ExecuteAdo sDSN, sSQL, timeoutSeconds
        Exit Sub
    End If

    Set db = OpenIsolatedCurrentDb(ws)
    Set qdfTemp = CreatePassthroughQueryDef(db, sDSN, sSQL, False, timeoutSeconds)

    qdfTemp.Execute

Cleanup:
    On Error Resume Next
    Set qdfTemp = Nothing
    If Not db Is Nothing Then db.Close
    Set db = Nothing
    If Not ws Is Nothing Then ws.Close
    Set ws = Nothing
    Exit Sub

HandleErr:
    Err.Raise _
        vbObjectError + 1026, _
        cModuleName & ".PTQ_Execute", _
        "Passthrough execute failed." & vbCrLf & _
        "DSN: " & sDSN & vbCrLf & _
        "SQL: " & sSQL & vbCrLf & _
        "Details: " & Err.Description

End Sub

Public Sub PTQ_Run_Proc( _
    ByVal sOracleProcName As String, _
    Optional ByVal commitAfter As Boolean = True, _
    Optional ByVal sDSN As String = "" _
)

    Dim sSQL As String

    sSQL = "BEGIN" & vbCrLf & _
           "    " & sOracleProcName & ";" & vbCrLf

    If commitAfter Then
        sSQL = sSQL & "    COMMIT;" & vbCrLf
    End If

    sSQL = sSQL & "END;"

    PTQ_Execute sSQL, , sDSN

End Sub

'------------------------------------------------------------------------------------
' Passthrough row materialization helpers
'------------------------------------------------------------------------------------

Public Function PTQ_GetRows( _
    ByVal sSQL As String, _
    Optional ByVal sDSN As String = "" _
) As Collection

    Dim ws As DAO.Workspace
    Dim db As DAO.Database
    Dim qdfTemp As DAO.QueryDef
    Dim rs As DAO.Recordset
    Dim rows As Collection
    Dim rowDict As Object
    Dim fld As DAO.Field

    On Error GoTo HandleErr

    sDSN = ResolveDefaultDSN(sDSN)

    Set rows = New Collection
    Set db = OpenIsolatedCurrentDb(ws)
    Set qdfTemp = CreatePassthroughQueryDef(db, sDSN, sSQL, True)
    Set rs = qdfTemp.OpenRecordset(dbOpenSnapshot)

    Do While Not rs.EOF
        Set rowDict = CreateObject("Scripting.Dictionary")
        rowDict.CompareMode = vbTextCompare

        For Each fld In rs.fields
            AddRowDictValue rowDict, fld.Name, fld.Value, cModuleName & ".PTQ_GetRows"
        Next fld

        rows.Add rowDict
        rs.MoveNext
    Loop

    Set PTQ_GetRows = rows

Cleanup:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set qdfTemp = Nothing
    If Not db Is Nothing Then db.Close
    Set db = Nothing
    If Not ws Is Nothing Then ws.Close
    Set ws = Nothing
    Exit Function

HandleErr:
    Err.Raise _
        vbObjectError + 1027, _
        cModuleName & ".PTQ_GetRows", _
        "Passthrough row retrieval failed." & vbCrLf & _
        "DSN: " & sDSN & vbCrLf & _
        "SQL: " & sSQL & vbCrLf & _
        "Details: " & Err.Description

End Function

Public Function PTQ_GetRow( _
    ByVal sSQL As String, _
    Optional ByVal sDSN As String = "" _
) As Object

    Dim rows As Collection

    Set rows = PTQ_GetRows(sSQL, sDSN)

    If rows.Count = 0 Then
        Set PTQ_GetRow = Nothing
    Else
        Set PTQ_GetRow = rows(1)
    End If

End Function

Public Function PTQ_Rs(ByVal sSQL As String, Optional ByVal sDSN As String = "") As DAO.Recordset
    Err.Raise _
        vbObjectError + 1028, _
        cModuleName & ".PTQ_Rs", _
        "PTQ_Rs has been retired because returning a DAO.Recordset from a temporary passthrough QueryDef is not a stable contract. Use PTQ_GetRow / PTQ_GetRows or an immediate-consumption helper."
End Function

'------------------------------------------------------------------------------------
' Oracle helpers
'------------------------------------------------------------------------------------

Public Function Get_ODBC_User() As String

    If Len(g_OracleSessionUser) > 0 Then
        Get_ODBC_User = g_OracleSessionUser
    Else
        Get_ODBC_User = UCase$(PTQ_SelectString("SELECT USER FROM DUAL"))
    End If

End Function

Public Function Check_Oracle_User_Role(ByVal sOracleRoleName As String) As Boolean

    Dim sSQL As String

    sSQL = "SELECT COUNT(*) " & _
           "FROM user_role_privs " & _
           "WHERE granted_role = '" & SqlTextLiteral(UCase$(Trim$(sOracleRoleName))) & "'"

    Check_Oracle_User_Role = (PTQ_SelectLong(sSQL) > 0)

End Function

Public Function Oracle_GetNextSequenceValue( _
    ByVal sSchema As String, _
    ByVal sSequenceName As String, _
    Optional ByVal sDSN As String = "" _
) As Long

    Dim sObjectName As String
    Dim sSQL As String

    sSequenceName = Trim$(sSequenceName)
    sSchema = Trim$(sSchema)
    sDSN = ResolveDefaultDSN(sDSN)

    If Len(sSequenceName) = 0 Then
        Err.Raise vbObjectError + 1030, cModuleName & ".Oracle_GetNextSequenceValue", "Sequence name cannot be blank."
    End If

    If Len(sSchema) > 0 Then
        sObjectName = sSchema & "." & sSequenceName
    Else
        sObjectName = sSequenceName
    End If

    sSQL = "SELECT " & sObjectName & ".NEXTVAL FROM DUAL"

    Oracle_GetNextSequenceValue = PTQ_SelectLong(sSQL, sDSN)

End Function

'------------------------------------------------------------------------------------
' SQL literal / string helpers
'------------------------------------------------------------------------------------

Public Function SqlTextLiteral(ByVal sValue As String) As String
    SqlTextLiteral = Replace$(sValue, "'", "''")
End Function

Public Function SqlStringOrNull(ByVal v As Variant) As String
    If IsNull(v) Then
        SqlStringOrNull = "NULL"
    Else
        SqlStringOrNull = "'" & SqlTextLiteral(CStr(v)) & "'"
    End If
End Function

Public Function SqlNumberOrNull(ByVal v As Variant) As String
    If IsNull(v) Then
        SqlNumberOrNull = "NULL"
    ElseIf Len(Trim$(CStr(v))) = 0 Then
        SqlNumberOrNull = "NULL"
    Else
        SqlNumberOrNull = Replace$(Trim$(CStr(v)), ",", ".")
    End If
End Function

Public Function SqlDateOrNull(ByVal v As Variant) As String
    If IsNull(v) Then
        SqlDateOrNull = "NULL"
    Else
        SqlDateOrNull = _
            "TO_DATE('" & Format$(CDate(v), "yyyy-mm-dd hh:nn:ss") & "', 'YYYY-MM-DD HH24:MI:SS')"
    End If
End Function

Public Function SqlTimestampOrNull(ByVal v As Variant) As String
    If IsNull(v) Then
        SqlTimestampOrNull = "NULL"
    Else
        SqlTimestampOrNull = _
            "TO_TIMESTAMP('" & Format$(CDate(v), "yyyy-mm-dd hh:nn:ss") & "', 'YYYY-MM-DD HH24:MI:SS')"
    End If
End Function

Public Function SqlBooleanNumber(ByVal v As Variant) As String
    If IsNull(v) Then
        SqlBooleanNumber = "NULL"
    ElseIf CBool(v) Then
        SqlBooleanNumber = "1"
    Else
        SqlBooleanNumber = "0"
    End If
End Function

Public Function SqlYNFlag(ByVal v As Variant) As String
    If IsNull(v) Then
        SqlYNFlag = "NULL"
    ElseIf CBool(v) Then
        SqlYNFlag = "'Y'"
    Else
        SqlYNFlag = "'N'"
    End If
End Function
