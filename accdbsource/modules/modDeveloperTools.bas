Attribute VB_Name = "modDeveloperTools"
'====================================================================================
' modDeveloperTools
'====================================================================================
'
' README
' ------
' Purpose
' -------
' Contains developer-only helper routines for working with exported Access/VBA source.
'
' These helpers are intended for use from the Immediate Window while building forms,
' SQL strings, and scaffolding. They should not be called by runtime application code.
'
'
' Responsibilities
' ----------------
' This module is responsible for:
'
'     - reading Oracle table metadata for developer scaffolding
'     - printing starter clsOracleFormField configuration routines
'     - printing starter form modules for unbound Oracle maintenance forms
'     - formatting pasted Oracle SQL as copy/paste-ready Access VBA string literals
'     - printing long generated text safely to the Immediate Window
'
'
' Key public helpers
' ------------------
' scaffolding:
'     Ofm_Debug_PrintFieldScaffold
'     Ofm_Debug_PrintFormModuleScaffold
'     Ofm_Debug_PrintFieldScaffoldFromSql
'
' SQL formatting:
'     Dev_FormatAccessSqlLiteral
'     Dev_PrintAccessSqlLiteral
'
'
' Typical usage
' -------------
' Print a starter ConfigureFields routine:
'
'     Ofm_Debug_PrintFieldScaffold "APP_RECORD"
'
' Print a starter ConfigureFields routine from a joined/read-model SQL statement:
'
'     Ofm_Debug_PrintFieldScaffoldFromSql sSQL, "APP_RECORD"
'
' Print a starter form module:
'
'     Ofm_Debug_PrintFormModuleScaffold "APP_RECORD"
'
' From the Immediate Window:
'
'     Dev_PrintAccessSqlLiteral _
'         "SELECT RECORD_ID, RECORD_NAME" & vbCrLf & _
'         "FROM APP_RECORD" & vbCrLf & _
'         "WHERE STATUS_CD = 'A'"
'
' This prints:
'
'     sSQL = _
'         "SELECT RECORD_ID, RECORD_NAME" & vbCrLf & _
'         "FROM APP_RECORD" & vbCrLf & _
'         "WHERE STATUS_CD = 'A'"
'
' For very long SQL, use incremental assignment output:
'
'     Dev_PrintAccessSqlLiteral sSqlText, "sSQL", True, True
'
'
' Design notes
' ------------
' This module is developer tooling. Runtime forms should depend on
' modOracleFormEngine and clsOracleFormField, not on this module.
'
'
' Version
' -------
' v1
'
'====================================================================================

Option Compare Database
Option Explicit

Private Const cModuleName As String = "modDeveloperTools"
Private Const cDebugPrintChunkLength As Long = 900

'------------------------------------------------------------------------------------
' Metadata / scaffolding helpers
'------------------------------------------------------------------------------------

Private Function Ofm_ResolveScaffoldSchemaName(ByVal schemaName As String) As String

    schemaName = UCase$(Trim$(schemaName))

    If Len(schemaName) = 0 Then
        schemaName = UCase$(Trim$(Get_DB_Schema()))
    End If

    If Len(schemaName) = 0 Then
        schemaName = UCase$(Trim$(Get_ODBC_User()))
    End If

    Ofm_ResolveScaffoldSchemaName = schemaName

End Function

Private Function Ofm_VbaStringLiteral(ByVal sValue As String) As String
    Ofm_VbaStringLiteral = Replace$(sValue, """", """""")
End Function

Private Function Ofm_IsStringDataType(ByVal dataType As String) As Boolean

    dataType = UCase$(Trim$(dataType))

    Select Case dataType
        Case "CHAR", "NCHAR", "VARCHAR2", "NVARCHAR2", "VARCHAR", "CLOB", "NCLOB", _
             "LONG", "ROWID", "UROWID"
            Ofm_IsStringDataType = True
    End Select

End Function

Private Function Ofm_IsYesNoFlagColumn(ByVal columnName As String, ByVal dataType As String) As Boolean

    columnName = UCase$(Trim$(columnName))
    dataType = UCase$(Trim$(dataType))

    If Not Ofm_IsStringDataType(dataType) Then Exit Function

    If Right$(columnName, 3) = "_YN" _
       Or Right$(columnName, 5) = "_FLAG" _
       Or Right$(columnName, 5) = "_INDC" Then
        Ofm_IsYesNoFlagColumn = True
    End If

End Function

Private Function Ofm_GetPrimaryKeyColumns( _
    ByVal schemaName As String, _
    ByVal tableName As String, _
    Optional ByVal dsn As String = "" _
) As Collection

    Dim sSQL As String
    Dim currentUser As String

    schemaName = Ofm_ResolveScaffoldSchemaName(schemaName)
    tableName = UCase$(Trim$(tableName))
    currentUser = UCase$(Trim$(Get_ODBC_User()))

    If Len(tableName) = 0 Then
        Err.Raise vbObjectError + 5060, cModuleName & ".Ofm_GetPrimaryKeyColumns", "Table name cannot be blank."
    End If

    If Len(schemaName) = 0 Or StrComp(schemaName, currentUser, vbTextCompare) = 0 Then
        sSQL = _
            "SELECT ucc.column_name " & _
            "FROM user_constraints uc " & _
            "INNER JOIN user_cons_columns ucc " & _
            "    ON ucc.constraint_name = uc.constraint_name " & _
            "WHERE uc.constraint_type = 'P' " & _
            "  AND uc.table_name = '" & SqlTextLiteral(tableName) & "' " & _
            "ORDER BY ucc.position"
    Else
        sSQL = _
            "SELECT acc.column_name " & _
            "FROM all_constraints ac " & _
            "INNER JOIN all_cons_columns acc " & _
            "    ON acc.owner = ac.owner " & _
            "   AND acc.constraint_name = ac.constraint_name " & _
            "WHERE ac.constraint_type = 'P' " & _
            "  AND ac.owner = '" & SqlTextLiteral(schemaName) & "' " & _
            "  AND ac.table_name = '" & SqlTextLiteral(tableName) & "' " & _
            "ORDER BY acc.position"
    End If

    Set Ofm_GetPrimaryKeyColumns = PTQ_GetRows(sSQL, dsn)

End Function

Private Function Ofm_GetSinglePrimaryKeyFieldName( _
    ByVal schemaName As String, _
    ByVal tableName As String, _
    Optional ByVal dsn As String = "" _
) As String

    Dim pkRows As Collection

    Set pkRows = Ofm_GetPrimaryKeyColumns(schemaName, tableName, dsn)

    If pkRows.Count = 1 Then
        Ofm_GetSinglePrimaryKeyFieldName = CStr(pkRows(1)("COLUMN_NAME"))
    End If

End Function

Private Function Ofm_GetTableColumnMetadata( _
    ByVal schemaName As String, _
    ByVal tableName As String, _
    Optional ByVal dsn As String = "" _
) As Collection

    Dim sSQL As String
    Dim currentUser As String
    Dim rows As Collection

    schemaName = Ofm_ResolveScaffoldSchemaName(schemaName)
    tableName = UCase$(Trim$(tableName))
    currentUser = UCase$(Trim$(Get_ODBC_User()))

    If Len(tableName) = 0 Then
        Err.Raise vbObjectError + 5061, cModuleName & ".Ofm_GetTableColumnMetadata", "Table name cannot be blank."
    End If

    If Len(schemaName) = 0 Or StrComp(schemaName, currentUser, vbTextCompare) = 0 Then
        sSQL = _
            "SELECT utc.column_name, " & _
            "       utc.data_type, " & _
            "       utc.nullable, " & _
            "       NVL(utc.data_default, '') AS data_default, " & _
            "       utc.column_id, " & _
            "       NVL(ucc.comments, '') AS comments " & _
            "FROM user_tab_columns utc " & _
            "LEFT JOIN user_col_comments ucc " & _
            "    ON ucc.table_name = utc.table_name " & _
            "   AND ucc.column_name = utc.column_name " & _
            "WHERE utc.table_name = '" & SqlTextLiteral(tableName) & "' " & _
            "ORDER BY utc.column_id"
    Else
        sSQL = _
            "SELECT atc.column_name, " & _
            "       atc.data_type, " & _
            "       atc.nullable, " & _
            "       NVL(atc.data_default, '') AS data_default, " & _
            "       atc.column_id, " & _
            "       NVL(acc.comments, '') AS comments " & _
            "FROM all_tab_columns atc " & _
            "LEFT JOIN all_col_comments acc " & _
            "    ON acc.owner = atc.owner " & _
            "   AND acc.table_name = atc.table_name " & _
            "   AND acc.column_name = atc.column_name " & _
            "WHERE atc.owner = '" & SqlTextLiteral(schemaName) & "' " & _
            "  AND atc.table_name = '" & SqlTextLiteral(tableName) & "' " & _
            "ORDER BY atc.column_id"
    End If

    Set rows = PTQ_GetRows(sSQL, dsn)

    If rows.Count = 0 Then
        Err.Raise vbObjectError + 5062, cModuleName & ".Ofm_GetTableColumnMetadata", _
                  "No Oracle columns were found for " & Ofm_GetQualifiedObjectName(schemaName, tableName) & "."
    End If

    Set Ofm_GetTableColumnMetadata = rows

End Function

Private Function Ofm_NormalizeScaffoldSql(ByVal sSQL As String) As String

    sSQL = Trim$(sSQL)

    Do While Len(sSQL) > 0 And (Right$(sSQL, 1) = ";" Or Right$(sSQL, 1) = "/")
        sSQL = Trim$(Left$(sSQL, Len(sSQL) - 1))
    Loop

    If Len(sSQL) = 0 Then
        Err.Raise vbObjectError + 5065, cModuleName & ".Ofm_NormalizeScaffoldSql", "SQL cannot be blank."
    End If

    Ofm_NormalizeScaffoldSql = sSQL

End Function

Private Function Ofm_WrapSqlForMetadata(ByVal sSQL As String) As String

    Ofm_WrapSqlForMetadata = _
        "SELECT * " & _
        "FROM (" & vbCrLf & _
        Ofm_NormalizeScaffoldSql(sSQL) & vbCrLf & _
        ") OFM_SQL_SRC " & _
        "WHERE 1 = 0"

End Function

Private Function Ofm_DaoFieldTypeToOracleType(ByVal fieldType As Long) As String

    Select Case fieldType
        Case dbBoolean
            Ofm_DaoFieldTypeToOracleType = "NUMBER"
        Case dbByte, dbInteger, dbLong, dbBigInt, dbSingle, dbDouble, dbCurrency, dbNumeric, dbDecimal, dbFloat
            Ofm_DaoFieldTypeToOracleType = "NUMBER"
        Case dbDate, dbTime, dbTimeStamp
            Ofm_DaoFieldTypeToOracleType = "DATE"
        Case dbText, dbChar
            Ofm_DaoFieldTypeToOracleType = "VARCHAR2"
        Case dbMemo
            Ofm_DaoFieldTypeToOracleType = "CLOB"
        Case dbBinary, dbLongBinary, dbVarBinary
            Ofm_DaoFieldTypeToOracleType = "RAW"
        Case Else
            Ofm_DaoFieldTypeToOracleType = "VARCHAR2"
    End Select

End Function

Private Function Ofm_GetSqlColumnMetadata( _
    ByVal sSQL As String, _
    Optional ByVal dsn As String = "" _
) As Collection

    Dim db As DAO.Database
    Dim qdfTemp As DAO.QueryDef
    Dim rs As DAO.Recordset
    Dim fld As DAO.Field
    Dim rows As Collection
    Dim rowData As Object
    Dim columnIndex As Long
    Dim sWrappedSql As String
    Dim sConnect As String

    On Error GoTo HandleErr

    sWrappedSql = Ofm_WrapSqlForMetadata(sSQL)
    sConnect = Get_Runtime_ODBC_Conn_Str(dsn)

    If Len(Trim$(sConnect)) = 0 Then
        Err.Raise vbObjectError + 5066, cModuleName & ".Ofm_GetSqlColumnMetadata", _
                  "No runtime connection string or DSN-based connection string could be resolved."
    End If

    Set rows = New Collection
    Set db = CurrentDb
    Set qdfTemp = db.CreateQueryDef(vbNullString)

    With qdfTemp
        .Connect = sConnect
        .returnsRecords = True
        .ODBCTimeout = 30
        .sql = sWrappedSql
    End With

    Set rs = qdfTemp.OpenRecordset(dbOpenSnapshot)

    columnIndex = 0

    For Each fld In rs.fields
        columnIndex = columnIndex + 1

        Set rowData = CreateObject("Scripting.Dictionary")
        rowData.CompareMode = vbTextCompare
        rowData("COLUMN_NAME") = UCase$(Trim$(fld.Name))
        rowData("DATA_TYPE") = Ofm_DaoFieldTypeToOracleType(fld.Type)
        rowData("NULLABLE") = "Y"
        rowData("DATA_DEFAULT") = vbNullString
        rowData("COLUMN_ID") = columnIndex
        rowData("COMMENTS") = vbNullString

        rows.Add rowData
    Next fld

    If rows.Count = 0 Then
        Err.Raise vbObjectError + 5067, cModuleName & ".Ofm_GetSqlColumnMetadata", _
                  "No columns were returned by the SQL metadata probe."
    End If

    Set Ofm_GetSqlColumnMetadata = rows

Cleanup:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set qdfTemp = Nothing
    Set db = Nothing
    Exit Function

HandleErr:
    Err.Raise _
        vbObjectError + 5068, _
        cModuleName & ".Ofm_GetSqlColumnMetadata", _
        "SQL scaffold metadata probe failed." & vbCrLf & _
        "SQL: " & sWrappedSql & vbCrLf & _
        "Details: " & Err.Description

End Function

Private Function Ofm_GetColumnMetadataDictionary(ByRef columnRows As Collection) As Object

    Dim result As Object
    Dim rowData As Object
    Dim columnName As String

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare

    For Each rowData In columnRows
        columnName = UCase$(Trim$(Nz(rowData("COLUMN_NAME"), vbNullString)))

        If Len(columnName) > 0 Then
            If result.Exists(columnName) Then
                Set result.Item(columnName) = rowData
            Else
                result.Add columnName, rowData
            End If
        End If
    Next rowData

    Set Ofm_GetColumnMetadataDictionary = result

End Function

Public Sub Ofm_Debug_PrintFieldScaffold( _
    ByVal tableName As String, _
    Optional ByVal schemaName As String = "", _
    Optional ByVal keyFieldName As String = "", _
    Optional ByVal sequenceName As String = "", _
    Optional ByVal dsn As String = "" _
)

    Dim schemaNameResolved As String
    Dim pkRows As Collection
    Dim columnRows As Collection
    Dim rowData As Object
    Dim columnName As String
    Dim dataType As String
    Dim nullableFlag As String
    Dim commentText As String
    Dim isKey As Boolean
    Dim isRequired As Boolean
    Dim isStringField As Boolean
    Dim lineText As String

    tableName = UCase$(Trim$(tableName))
    schemaNameResolved = Ofm_ResolveScaffoldSchemaName(schemaName)
    keyFieldName = UCase$(Trim$(keyFieldName))
    sequenceName = UCase$(Trim$(sequenceName))

    If Len(tableName) = 0 Then
        Err.Raise vbObjectError + 5063, cModuleName & ".Ofm_Debug_PrintFieldScaffold", "Table name cannot be blank."
    End If

    Set pkRows = Ofm_GetPrimaryKeyColumns(schemaNameResolved, tableName, dsn)

    If Len(keyFieldName) = 0 Then
        keyFieldName = Ofm_GetSinglePrimaryKeyFieldName(schemaNameResolved, tableName, dsn)
    End If

    Set columnRows = Ofm_GetTableColumnMetadata(schemaNameResolved, tableName, dsn)

    Debug.Print String$(90, "=")
    Debug.Print "' clsOracleFormField scaffold for " & Ofm_GetQualifiedObjectName(schemaNameResolved, tableName)
    Debug.Print "' Generated by " & cModuleName & ".Ofm_Debug_PrintFieldScaffold on " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Debug.Print String$(90, "=")
    Debug.Print "Private Const cTableName As String = """ & tableName & """"

    If Len(keyFieldName) > 0 Then
        Debug.Print "Private Const cKeyField As String = """ & keyFieldName & """"
    Else
        Debug.Print "' No single-column primary key was inferred. Set cKeyField manually if needed."
    End If

    If Len(sequenceName) > 0 Then
        Debug.Print "Private Const cSequenceName As String = """ & sequenceName & """"
    Else
        Debug.Print "' Optional: Private Const cSequenceName As String = """ & tableName & "_SEQ"""
    End If

    Debug.Print ""
    Debug.Print "Private Sub ConfigureFields()"
    Debug.Print ""
    Debug.Print "    Dim f As clsOracleFormField"
    Debug.Print ""

    For Each rowData In columnRows

        columnName = UCase$(Trim$(Nz(rowData("COLUMN_NAME"), vbNullString)))
        dataType = UCase$(Trim$(Nz(rowData("DATA_TYPE"), vbNullString)))
        nullableFlag = UCase$(Trim$(Nz(rowData("NULLABLE"), vbNullString)))
        commentText = Trim$(Nz(rowData("COMMENTS"), vbNullString))
        isKey = (Len(keyFieldName) > 0 And StrComp(columnName, keyFieldName, vbTextCompare) = 0)
        isRequired = (nullableFlag = "N")
        isStringField = Ofm_IsStringDataType(dataType)

        If isKey Then
            lineText = "    Set f = Ofm_AddField(mFields, """ & columnName & """, """ & columnName & """, True, True, False)"
        Else
            lineText = "    Set f = Ofm_AddField(mFields, """ & columnName & """, """ & columnName & """)"
        End If

        Debug.Print lineText
        Debug.Print "    f.OracleDataType = """ & Ofm_VbaStringLiteral(dataType) & """"

        If Len(commentText) > 0 Then
            Debug.Print "    f.Description = """ & Ofm_VbaStringLiteral(commentText) & """"
        End If

        If isKey Then
            Debug.Print "    f.IsKey = True"
            Debug.Print "    f.IsUpdatable = False"
            If Len(sequenceName) > 0 Then
                Debug.Print "    f.IsDbGenerated = True"
            End If
        End If

        If isRequired Then
            Debug.Print "    f.IsRequired = True"
        End If

        If isStringField Then
            Debug.Print "    f.TrimOnSave = True"
        End If

        If Not isRequired Then
            Debug.Print "    f.NullIfBlank = True"
        End If

        If Ofm_IsYesNoFlagColumn(columnName, dataType) Then
            Debug.Print "    ' Review: this looks like a Y/N-style flag field."
            Debug.Print "    f.ControlKind = ""CHECKBOX"""
            Debug.Print "    f.UseCustomBooleanMapping = True"
            Debug.Print "    f.CheckedValue = ""Y"""
            Debug.Print "    f.UncheckedValue = ""N"""
        End If

        Debug.Print ""
    Next rowData

    Debug.Print "End Sub"
    Debug.Print ""

    If pkRows.Count = 0 Then
        Debug.Print "' Note: no primary-key metadata was found."
    ElseIf pkRows.Count > 1 Then
        Debug.Print "' Note: a composite primary key was found. The current form engine expects one key field."
        Debug.Print "' Primary-key columns:"
        For Each rowData In pkRows
            Debug.Print "'     " & CStr(rowData("COLUMN_NAME"))
        Next rowData
        Debug.Print "' Choose a strategy manually before using Ofm_SaveRecord / Ofm_LoadForm."
    End If

    Debug.Print "' Review the generated defaults before using them in a form."
    Debug.Print String$(90, "=")

End Sub

Public Sub Ofm_Debug_PrintFieldScaffoldFromSql( _
    ByVal sSQL As String, _
    Optional ByVal baseTableName As String = "", _
    Optional ByVal schemaName As String = "", _
    Optional ByVal keyFieldName As String = "", _
    Optional ByVal sequenceName As String = "", _
    Optional ByVal dsn As String = "" _
)

    Dim schemaNameResolved As String
    Dim sqlColumnRows As Collection
    Dim baseColumnRows As Collection
    Dim baseColumns As Object
    Dim pkRows As Collection
    Dim rowData As Object
    Dim metadataRow As Object
    Dim columnName As String
    Dim dataType As String
    Dim nullableFlag As String
    Dim commentText As String
    Dim controlName As String
    Dim isBaseField As Boolean
    Dim isKey As Boolean
    Dim isRequired As Boolean
    Dim isStringField As Boolean

    baseTableName = UCase$(Trim$(baseTableName))
    schemaNameResolved = Ofm_ResolveScaffoldSchemaName(schemaName)
    keyFieldName = UCase$(Trim$(keyFieldName))
    sequenceName = UCase$(Trim$(sequenceName))

    Set sqlColumnRows = Ofm_GetSqlColumnMetadata(sSQL, dsn)

    If Len(baseTableName) > 0 Then
        Set pkRows = Ofm_GetPrimaryKeyColumns(schemaNameResolved, baseTableName, dsn)

        If Len(keyFieldName) = 0 Then
            keyFieldName = Ofm_GetSinglePrimaryKeyFieldName(schemaNameResolved, baseTableName, dsn)
        End If

        Set baseColumnRows = Ofm_GetTableColumnMetadata(schemaNameResolved, baseTableName, dsn)
        Set baseColumns = Ofm_GetColumnMetadataDictionary(baseColumnRows)
    Else
        Set pkRows = New Collection
    End If

    Debug.Print String$(90, "=")
    Debug.Print "' clsOracleFormField scaffold from SQL read model"
    If Len(baseTableName) > 0 Then
        Debug.Print "' Base table: " & Ofm_GetQualifiedObjectName(schemaNameResolved, baseTableName)
    Else
        Debug.Print "' Base table: not supplied. SQL columns are scaffolded as read-only by default."
    End If
    Debug.Print "' Generated by " & cModuleName & ".Ofm_Debug_PrintFieldScaffoldFromSql on " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Debug.Print String$(90, "=")

    If Len(baseTableName) > 0 Then
        Debug.Print "Private Const cTableName As String = """ & baseTableName & """"
    Else
        Debug.Print "' Set cTableName manually if this form will save to a base table."
    End If

    If Len(keyFieldName) > 0 Then
        Debug.Print "Private Const cKeyField As String = """ & keyFieldName & """"
    Else
        Debug.Print "' Set cKeyField manually if this form will save records."
    End If

    If Len(sequenceName) > 0 Then
        Debug.Print "Private Const cSequenceName As String = """ & sequenceName & """"
    ElseIf Len(baseTableName) > 0 Then
        Debug.Print "' Optional: Private Const cSequenceName As String = """ & baseTableName & "_SEQ"""
    End If

    Debug.Print ""
    Debug.Print "' SQL used for the read model:"
    Debug.Print "'     Review aliases carefully. Joined/display-only columns should remain non-updatable."
    Debug.Print ""
    Debug.Print "Private Sub ConfigureFields()"
    Debug.Print ""
    Debug.Print "    Dim f As clsOracleFormField"
    Debug.Print ""

    For Each rowData In sqlColumnRows

        columnName = UCase$(Trim$(Nz(rowData("COLUMN_NAME"), vbNullString)))
        Set metadataRow = rowData
        isBaseField = False

        If Not baseColumns Is Nothing Then
            If baseColumns.Exists(columnName) Then
                Set metadataRow = baseColumns(columnName)
                isBaseField = True
            End If
        End If

        dataType = UCase$(Trim$(Nz(metadataRow("DATA_TYPE"), vbNullString)))
        nullableFlag = UCase$(Trim$(Nz(metadataRow("NULLABLE"), vbNullString)))
        commentText = Trim$(Nz(metadataRow("COMMENTS"), vbNullString))
        isKey = (Len(keyFieldName) > 0 And StrComp(columnName, keyFieldName, vbTextCompare) = 0)
        isRequired = (nullableFlag = "N")
        isStringField = Ofm_IsStringDataType(dataType)
        controlName = Ofm_GetSuggestedControlName(columnName, dataType, isKey)

        If isKey Then
            Debug.Print "    Set f = Ofm_AddField(mFields, """ & columnName & """, """ & controlName & """, True, True, False)"
        ElseIf isBaseField Then
            Debug.Print "    Set f = Ofm_AddField(mFields, """ & columnName & """, """ & controlName & """)"
        Else
            Debug.Print "    Set f = Ofm_AddField(mFields, """", """ & controlName & """, False, False, False)"
            Debug.Print "    f.LoadFieldName = """ & columnName & """"
            Debug.Print "    ' Display-only SQL column. Set DbFieldName / IsInsertable / IsUpdatable manually if this should write to the base table."
        End If

        Debug.Print "    f.OracleDataType = """ & Ofm_VbaStringLiteral(dataType) & """"

        If Len(commentText) > 0 Then
            Debug.Print "    f.Description = """ & Ofm_VbaStringLiteral(commentText) & """"
        End If

        If isKey Then
            Debug.Print "    f.IsKey = True"
            Debug.Print "    f.IsUpdatable = False"
            If Len(sequenceName) > 0 Then
                Debug.Print "    f.IsDbGenerated = True"
            End If
        End If

        If isRequired And (isBaseField Or isKey) Then
            Debug.Print "    f.IsRequired = True"
        End If

        If isStringField Then
            Debug.Print "    f.TrimOnSave = True"
        End If

        If Not isRequired Then
            Debug.Print "    f.NullIfBlank = True"
        End If

        If Ofm_IsYesNoFlagColumn(columnName, dataType) Then
            Debug.Print "    ' Review: this looks like a Y/N-style flag field."
            Debug.Print "    f.ControlKind = ""CHECKBOX"""
            Debug.Print "    f.UseCustomBooleanMapping = True"
            Debug.Print "    f.CheckedValue = ""Y"""
            Debug.Print "    f.UncheckedValue = ""N"""
        End If

        Debug.Print ""
    Next rowData

    Debug.Print "End Sub"
    Debug.Print ""

    If Len(baseTableName) = 0 Then
        Debug.Print "' Note: no base table was supplied, so generated fields are read-only unless reviewed manually."
    ElseIf pkRows.Count = 0 Then
        Debug.Print "' Note: no primary-key metadata was found for the base table."
    ElseIf pkRows.Count > 1 Then
        Debug.Print "' Note: a composite primary key was found. The current form engine expects one key field."
        Debug.Print "' Primary-key columns:"
        For Each rowData In pkRows
            Debug.Print "'     " & CStr(rowData("COLUMN_NAME"))
        Next rowData
        Debug.Print "' Choose a strategy manually before using Ofm_SaveRecord / Ofm_LoadFormBySql."
    End If

    Debug.Print "' Review aliases, display-only fields, lookup controls, and DbFieldName mappings before using this scaffold."
    Debug.Print String$(90, "=")

End Sub

Private Function Ofm_ToPascalCase(ByVal sValue As String) As String

    Dim parts() As String
    Dim i As Long
    Dim partText As String
    Dim result As String

    sValue = LCase$(Trim$(sValue))

    If Len(sValue) = 0 Then Exit Function

    parts = Split(sValue, "_")

    For i = LBound(parts) To UBound(parts)
        partText = Trim$(parts(i))

        If Len(partText) > 0 Then
            result = result & UCase$(Left$(partText, 1)) & Mid$(partText, 2)
        End If
    Next i

    Ofm_ToPascalCase = result

End Function

Private Function Ofm_GetSuggestedControlName( _
    ByVal columnName As String, _
    ByVal dataType As String, _
    ByVal isKey As Boolean _
) As String

    Dim suffixText As String

    columnName = UCase$(Trim$(columnName))
    suffixText = Ofm_ToPascalCase(columnName)

    If isKey Then
        Ofm_GetSuggestedControlName = columnName
        Exit Function
    End If

    If Ofm_IsYesNoFlagColumn(columnName, dataType) Then
        Ofm_GetSuggestedControlName = "chk" & suffixText
    Else
        Ofm_GetSuggestedControlName = "txt" & suffixText
    End If

End Function

Public Sub Ofm_Debug_PrintFormModuleScaffold( _
    ByVal tableName As String, _
    Optional ByVal formName As String = "", _
    Optional ByVal schemaName As String = "", _
    Optional ByVal keyFieldName As String = "", _
    Optional ByVal sequenceName As String = "", _
    Optional ByVal dsn As String = "" _
)

    Dim schemaNameResolved As String
    Dim columnRows As Collection
    Dim pkRows As Collection
    Dim rowData As Object
    Dim columnName As String
    Dim dataType As String
    Dim nullableFlag As String
    Dim commentText As String
    Dim controlName As String
    Dim isKey As Boolean
    Dim isRequired As Boolean
    Dim isStringField As Boolean

    tableName = UCase$(Trim$(tableName))
    formName = Trim$(formName)
    schemaNameResolved = Ofm_ResolveScaffoldSchemaName(schemaName)
    keyFieldName = UCase$(Trim$(keyFieldName))
    sequenceName = UCase$(Trim$(sequenceName))

    If Len(tableName) = 0 Then
        Err.Raise vbObjectError + 5064, cModuleName & ".Ofm_Debug_PrintFormModuleScaffold", "Table name cannot be blank."
    End If

    If Len(formName) = 0 Then
        formName = "frm" & Ofm_ToPascalCase(tableName)
    End If

    Set pkRows = Ofm_GetPrimaryKeyColumns(schemaNameResolved, tableName, dsn)

    If Len(keyFieldName) = 0 Then
        keyFieldName = Ofm_GetSinglePrimaryKeyFieldName(schemaNameResolved, tableName, dsn)
    End If

    Set columnRows = Ofm_GetTableColumnMetadata(schemaNameResolved, tableName, dsn)

    Debug.Print String$(90, "=")
    Debug.Print "' Form-module scaffold for " & formName
    Debug.Print "' Table: " & Ofm_GetQualifiedObjectName(schemaNameResolved, tableName)
    Debug.Print "' Generated by " & cModuleName & ".Ofm_Debug_PrintFormModuleScaffold on " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Debug.Print String$(90, "=")
    Debug.Print "Option Compare Database"
    Debug.Print "Option Explicit"
    Debug.Print ""
    Debug.Print "Private Const cTableName As String = """ & tableName & """"

    If Len(keyFieldName) > 0 Then
        Debug.Print "Private Const cKeyField As String = """ & keyFieldName & """"
    Else
        Debug.Print "' Set cKeyField manually. A single-column primary key was not inferred."
    End If

    If Len(sequenceName) > 0 Then
        Debug.Print "Private Const cSequenceName As String = """ & sequenceName & """"
    Else
        Debug.Print "' Optional: Private Const cSequenceName As String = """ & tableName & "_SEQ"""
    End If

    Debug.Print ""
    Debug.Print "Private mFields As Collection"
    Debug.Print "Private mOriginalValues As Object"
    Debug.Print "Private mIsNewRecord As Boolean"
    Debug.Print "Private mAllowClose As Boolean"
    Debug.Print ""
    Debug.Print "Private Sub Form_Open(Cancel As Integer)"
    Debug.Print ""
    Debug.Print "    Set mFields = New Collection"
    Debug.Print "    Set mOriginalValues = CreateObject(""Scripting.Dictionary"")"
    Debug.Print ""
    Debug.Print "    ConfigureFields"
    Debug.Print ""
    Debug.Print "End Sub"
    Debug.Print ""
    Debug.Print "Private Sub Form_Load()"
    Debug.Print ""

    If Len(keyFieldName) > 0 Then
        Debug.Print "    Dim keyValue As Variant"
        Debug.Print ""
        Debug.Print "    keyValue = Nz(Me.OpenArgs, vbNullString)"
        Debug.Print ""
        Debug.Print "    If Len(CStr(keyValue)) = 0 Then"
        Debug.Print "        mIsNewRecord = True"
        Debug.Print "        Ofm_InitNewForm Me, mFields, mOriginalValues"
        Debug.Print "    Else"
        Debug.Print "        mIsNewRecord = False"
        Debug.Print "        Ofm_LoadForm Me, Get_DB_Schema(), cTableName, cKeyField, keyValue, mFields, mOriginalValues"
        Debug.Print "    End If"
    Else
        Debug.Print "    mIsNewRecord = True"
        Debug.Print "    Ofm_InitNewForm Me, mFields, mOriginalValues"
        Debug.Print "    ' Review: no single-column primary key was inferred, so load logic must be added manually."
    End If

    Debug.Print ""
    Debug.Print "End Sub"
    Debug.Print ""
    Debug.Print "Private Sub ConfigureFields()"
    Debug.Print ""
    Debug.Print "    Dim f As clsOracleFormField"
    Debug.Print ""

    For Each rowData In columnRows

        columnName = UCase$(Trim$(Nz(rowData("COLUMN_NAME"), vbNullString)))
        dataType = UCase$(Trim$(Nz(rowData("DATA_TYPE"), vbNullString)))
        nullableFlag = UCase$(Trim$(Nz(rowData("NULLABLE"), vbNullString)))
        commentText = Trim$(Nz(rowData("COMMENTS"), vbNullString))
        isKey = (Len(keyFieldName) > 0 And StrComp(columnName, keyFieldName, vbTextCompare) = 0)
        isRequired = (nullableFlag = "N")
        isStringField = Ofm_IsStringDataType(dataType)
        controlName = Ofm_GetSuggestedControlName(columnName, dataType, isKey)

        If isKey Then
            Debug.Print "    Set f = Ofm_AddField(mFields, """ & columnName & """, """ & controlName & """, True, True, False)"
        Else
            Debug.Print "    Set f = Ofm_AddField(mFields, """ & columnName & """, """ & controlName & """)"
        End If

        Debug.Print "    f.OracleDataType = """ & Ofm_VbaStringLiteral(dataType) & """"

        If Len(commentText) > 0 Then
            Debug.Print "    f.Description = """ & Ofm_VbaStringLiteral(commentText) & """"
        End If

        If isKey Then
            Debug.Print "    f.IsKey = True"
            Debug.Print "    f.IsUpdatable = False"
            If Len(sequenceName) > 0 Then
                Debug.Print "    f.IsDbGenerated = True"
            End If
        End If

        If isRequired Then
            Debug.Print "    f.IsRequired = True"
        End If

        If isStringField Then
            Debug.Print "    f.TrimOnSave = True"
        End If

        If Not isRequired Then
            Debug.Print "    f.NullIfBlank = True"
        End If

        If Ofm_IsYesNoFlagColumn(columnName, dataType) Then
            Debug.Print "    f.ControlKind = ""CHECKBOX"""
            Debug.Print "    f.UseCustomBooleanMapping = True"
            Debug.Print "    f.CheckedValue = ""Y"""
            Debug.Print "    f.UncheckedValue = ""N"""
        End If

        Debug.Print ""
    Next rowData

    Debug.Print "End Sub"
    Debug.Print ""
    Debug.Print "Private Sub ValidateForm()"
    Debug.Print ""
    Debug.Print "    Ofm_ValidateRequiredFields Me, mFields"
    Debug.Print "    ' Add form-specific validation rules here."
    Debug.Print ""
    Debug.Print "End Sub"
    Debug.Print ""
    Debug.Print "Private Function SaveRecord() As Boolean"
    Debug.Print ""
    Debug.Print "    On Error GoTo ErrHandler"
    Debug.Print ""

    If Len(keyFieldName) > 0 Then
        Debug.Print "    Dim savedKey As Variant"
        Debug.Print ""
        Debug.Print "    ValidateForm"
        Debug.Print ""
        Debug.Print "    savedKey = Ofm_SaveRecord( _"
        Debug.Print "        Me, _"
        Debug.Print "        Get_DB_Schema(), _"
        Debug.Print "        cTableName, _"
        Debug.Print "        cKeyField, _"
        Debug.Print "        mFields, _"
        Debug.Print "        mOriginalValues, _"
        Debug.Print "        mIsNewRecord, _"
        If Len(sequenceName) > 0 Then
            Debug.Print "        cSequenceName, _"
        Else
            Debug.Print "        vbNullString, _"
        End If
        Debug.Print "        True)"
        Debug.Print ""
        Debug.Print "    mIsNewRecord = False"
        Debug.Print "    Me." & keyFieldName & " = savedKey"
        Debug.Print "    SaveRecord = True"
    Else
        Debug.Print "    ' Review: add custom save logic here if the form does not use a single-column primary key."
    End If

    Debug.Print ""
    Debug.Print "    Exit Function"
    Debug.Print ""
    Debug.Print "ErrHandler:"
    Debug.Print "    MsgBox Err.Description, vbExclamation"
    Debug.Print ""
    Debug.Print "End Function"
    Debug.Print ""
    Debug.Print "Private Sub btnSave_Click()"
    Debug.Print ""
    Debug.Print "    If SaveRecord() Then"
    Debug.Print "        MsgBox ""Record saved."", vbInformation"
    Debug.Print "    End If"
    Debug.Print ""
    Debug.Print "End Sub"
    Debug.Print ""
    Debug.Print "Private Sub btnDelete_Click()"
    Debug.Print ""
    Debug.Print "    If mIsNewRecord Then Exit Sub"
    Debug.Print ""
    Debug.Print "    If MsgBox(""Delete this record?"", vbQuestion + vbYesNo) = vbNo Then Exit Sub"
    Debug.Print ""

    If Len(keyFieldName) > 0 Then
        Debug.Print "    Ofm_Delete Get_DB_Schema(), cTableName, cKeyField, Me." & keyFieldName
        Debug.Print ""
        Debug.Print "    mAllowClose = True"
        Debug.Print "    DoCmd.Close acForm, Me.Name"
    Else
        Debug.Print "    ' Review: add custom delete logic here if the form does not use a single-column primary key."
    End If

    Debug.Print ""
    Debug.Print "End Sub"
    Debug.Print ""
    Debug.Print "Private Function PromptToSaveIfDirty() As Boolean"
    Debug.Print ""
    Debug.Print "    Dim response As VbMsgBoxResult"
    Debug.Print ""
    Debug.Print "    PromptToSaveIfDirty = False"
    Debug.Print ""
    Debug.Print "    If Not Ofm_IsDirty(Me, mFields, mOriginalValues) Then"
    Debug.Print "        PromptToSaveIfDirty = True"
    Debug.Print "        Exit Function"
    Debug.Print "    End If"
    Debug.Print ""
    Debug.Print "    response = MsgBox(""Save changes before closing?"", vbYesNoCancel + vbQuestion, ""Unsaved Changes"")"
    Debug.Print ""
    Debug.Print "    Select Case response"
    Debug.Print "        Case vbYes"
    Debug.Print "            PromptToSaveIfDirty = SaveRecord()"
    Debug.Print "        Case vbNo"
    Debug.Print "            PromptToSaveIfDirty = True"
    Debug.Print "        Case vbCancel"
    Debug.Print "            PromptToSaveIfDirty = False"
    Debug.Print "    End Select"
    Debug.Print ""
    Debug.Print "End Function"
    Debug.Print ""
    Debug.Print "Private Sub btnClose_Click()"
    Debug.Print ""
    Debug.Print "    If Not PromptToSaveIfDirty() Then Exit Sub"
    Debug.Print ""
    Debug.Print "    mAllowClose = True"
    Debug.Print "    DoCmd.Close acForm, Me.Name"
    Debug.Print ""
    Debug.Print "End Sub"
    Debug.Print ""
    Debug.Print "Private Sub Form_Unload(Cancel As Integer)"
    Debug.Print ""
    Debug.Print "    If mAllowClose Then Exit Sub"
    Debug.Print ""
    Debug.Print "    If Not PromptToSaveIfDirty() Then"
    Debug.Print "        Cancel = True"
    Debug.Print "    End If"
    Debug.Print ""
    Debug.Print "End Sub"
    Debug.Print ""
    Debug.Print "' Suggested controls based on metadata:"

    For Each rowData In columnRows
        columnName = UCase$(Trim$(Nz(rowData("COLUMN_NAME"), vbNullString)))
        dataType = UCase$(Trim$(Nz(rowData("DATA_TYPE"), vbNullString)))
        isKey = (Len(keyFieldName) > 0 And StrComp(columnName, keyFieldName, vbTextCompare) = 0)
        controlName = Ofm_GetSuggestedControlName(columnName, dataType, isKey)
        Debug.Print "'     " & columnName & " -> " & controlName
    Next rowData

    If pkRows.Count = 0 Then
        Debug.Print "' Note: no primary-key metadata was found."
    ElseIf pkRows.Count > 1 Then
        Debug.Print "' Note: a composite primary key was found. The scaffold assumes a single key field and must be adjusted."
    End If

    Debug.Print "' Review lookup fields, combo-box choices, joined read models, and sequence usage before finalizing the form."
    Debug.Print String$(90, "=")

End Sub


'------------------------------------------------------------------------------------
' SQL literal formatting
'------------------------------------------------------------------------------------

Public Function Dev_FormatAccessSqlLiteral( _
    ByVal oracleSql As String, _
    Optional ByVal variableName As String = "sSQL", _
    Optional ByVal includeSqlLineBreaks As Boolean = True, _
    Optional ByVal useIncrementalAssignments As Boolean = False _
) As String

    Dim normalizedSql As String
    Dim lines() As String
    Dim i As Long
    Dim result As String
    Dim lineText As String
    Dim literalPart As String
    Dim lineBreakPart As String

    If Len(Trim$(variableName)) = 0 Then
        Err.Raise vbObjectError + 9000, cModuleName & ".Dev_FormatAccessSqlLiteral", _
                  "Variable name cannot be blank."
    End If

    normalizedSql = Replace$(oracleSql, vbCrLf, vbLf)
    normalizedSql = Replace$(normalizedSql, vbCr, vbLf)
    lines = Split(normalizedSql, vbLf)

    If useIncrementalAssignments Then
        result = variableName & " = vbNullString" & vbCrLf
    Else
        result = variableName & " = _" & vbCrLf
    End If

    For i = LBound(lines) To UBound(lines)

        lineText = Replace$(lines(i), """", """""")
        literalPart = """" & lineText & """"

        If includeSqlLineBreaks And i < UBound(lines) Then
            lineBreakPart = " & vbCrLf"
        Else
            lineBreakPart = vbNullString
        End If

        If useIncrementalAssignments Then
            result = result & variableName & " = " & variableName & " & " & literalPart & lineBreakPart
        Else
            result = result & "    " & literalPart & lineBreakPart

            If i < UBound(lines) Then
                result = result & " & _"
            End If
        End If

        If i < UBound(lines) Then
            result = result & vbCrLf
        End If

    Next i

    Dev_FormatAccessSqlLiteral = result

End Function

Public Sub Dev_PrintAccessSqlLiteral( _
    ByVal oracleSql As String, _
    Optional ByVal variableName As String = "sSQL", _
    Optional ByVal includeSqlLineBreaks As Boolean = True, _
    Optional ByVal useIncrementalAssignments As Boolean = False _
)

    Dim formattedText As String

    formattedText = Dev_FormatAccessSqlLiteral( _
        oracleSql, _
        variableName, _
        includeSqlLineBreaks, _
        useIncrementalAssignments)

    Dev_PrintLongText formattedText

End Sub

'------------------------------------------------------------------------------------
' Immediate Window helpers
'------------------------------------------------------------------------------------

Public Sub Dev_PrintLongText(ByVal textToPrint As String)

    Dim lines() As String
    Dim i As Long

    textToPrint = Replace$(textToPrint, vbCrLf, vbLf)
    textToPrint = Replace$(textToPrint, vbCr, vbLf)
    lines = Split(textToPrint, vbLf)

    For i = LBound(lines) To UBound(lines)
        Dev_PrintDebugLine lines(i)
    Next i

End Sub

Private Sub Dev_PrintDebugLine(ByVal textLine As String)

    Dim lStartPos As Long

    If Len(textLine) = 0 Then
        Debug.Print vbNullString
        Exit Sub
    End If

    lStartPos = 1

    Do While lStartPos <= Len(textLine)
        Debug.Print Mid$(textLine, lStartPos, cDebugPrintChunkLength)
        lStartPos = lStartPos + cDebugPrintChunkLength
    Loop

End Sub
