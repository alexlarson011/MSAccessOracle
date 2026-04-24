Attribute VB_Name = "modOracleFormEngine"
'====================================================================================
' modOracleFormEngine
'====================================================================================
'
' README
' ------
' Purpose
' -------
' Provides a reusable CRUD engine for unbound Access forms that read and write
' Oracle data without requiring bound linked tables.
'
' This module is designed for forms that:
'
'     - load data manually from Oracle
'     - push data into controls manually
'     - snapshot original values manually
'     - detect changes manually
'     - generate INSERT / UPDATE / DELETE SQL explicitly
'
'
' Responsibilities
' ----------------
' This module is responsible for:
'
'     - managing field definitions through clsOracleFormField
'     - loading an existing Oracle row into a form
'     - initializing a new blank / default row in a form
'     - converting Oracle values to UI control values
'     - converting UI control values back to Oracle values
'     - validating required fields
'     - snapshotting original values
'     - detecting dirty state
'     - building changed-fields-only UPDATE SQL
'     - building INSERT SQL
'     - building DELETE SQL
'     - executing insert / update / delete operations
'     - supporting sequence-first inserts
'
'
' Key public helpers
' ------------------
' field definition helpers:
'     Ofm_AddField
'     Ofm_GetKeyField
'     Ofm_GetFieldByControlName
'     Ofm_GetFieldByDbFieldName
'     Ofm_GetSelectList
'
' value translation helpers:
'     Ofm_ValuesEqual
'     Ofm_DbToControlValue
'     Ofm_ControlToDbValue
'     Ofm_GetControlValue
'     Ofm_SetControlValue
'
' list / combo helpers:
'     Ofm_LoadListControlBySql
'     Ofm_LoadLookupControl
'     Ofm_LoadLookupControls
'
' snapshot / dirty helpers:
'     Ofm_SnapshotValues
'     Ofm_IsDirty
'     Ofm_GetChangedFields
'
' validation helpers:
'     Ofm_ValidateRequiredFields
'
' load / new-record helpers:
'     Ofm_LoadForm
'     Ofm_LoadFormBySql
'     Ofm_InitNewForm
'
' SQL builders:
'     Ofm_GetQualifiedObjectName
'     Ofm_SqlValue
'     Ofm_BuildUpdateSql
'     Ofm_BuildInsertSql
'     Ofm_BuildDeleteSql
'
' CRUD execution helpers:
'     Ofm_Insert
'     Ofm_Update
'     Ofm_Delete
'     Ofm_SaveRecord
'
' debugging:
'     Ofm_Debug_PrintFieldScaffold
'     Ofm_Debug_PrintFormModuleScaffold
'     Ofm_DebugFieldSummary
'
'
' Typical usage
' -------------
' In a form module:
'
'     Private mFields As Collection
'     Private mOriginalValues As Object
'     Private mIsNewRecord As Boolean
'
' Build field config:
'
'     Set mFields = New Collection
'     Set mOriginalValues = CreateObject("Scripting.Dictionary")
'     Set f = Ofm_AddField(mFields, "PROJ_OPTN_ID", "PROJ_OPTN_ID", True, True, False)
'
' Joined read-model field example:
'
'     Set f = Ofm_AddField(mFields, "STATUS_CD", "txtStatusText", False, False, False)
'     f.LoadFieldName = "STATUS_TEXT"
'
' Combo-box lookup example:
'
'     Set f = Ofm_AddField(mFields, "STATUS_CD", "cboStatus")
'     f.ControlKind = "COMBOBOX"
'     f.LookupSql = "SELECT STATUS_CD, STATUS_TEXT FROM APP_STATUS_LU ORDER BY STATUS_TEXT"
'     f.LookupBoundColumn = 1
'     f.LookupDisplayColumn = 2
'     f.LookupShowColumnHeads = True
'     f.LookupIncludeBlankRow = True
'
' Load an existing record:
'
'     Ofm_LoadForm Me, Get_DB_Schema(), cTableName, cKeyField, keyValue, mFields, mOriginalValues
'
' Load an existing record from arbitrary SQL:
'
'     Ofm_LoadFormBySql Me, _
'         "SELECT p.PROJ_ID, s.STATUS_TEXT " & _
'         "FROM PROJECT p " & _
'         "LEFT JOIN STATUS_LU s ON s.STATUS_CD = p.STATUS_CD " & _
'         "WHERE p.PROJ_ID = " & Ofm_SqlValue(keyValue), _
'         mFields, mOriginalValues
'
' Initialize a new record:
'
'     Ofm_InitNewForm Me, mFields, mOriginalValues
'
' Load configured combo/list lookups:
'
'     Ofm_LoadLookupControls Me, mFields
'
' Save:
'
'     savedKey = Ofm_SaveRecord(Me, Get_DB_Schema(), cTableName, cKeyField, mFields, mOriginalValues, mIsNewRecord, cSequenceName, True)
'
' Save and reload from a read-model SQL statement:
'
'     savedKey = Ofm_SaveRecord(Me, Get_DB_Schema(), cTableName, cKeyField, mFields, mOriginalValues, mIsNewRecord, cSequenceName, True, _
'         "SELECT p.PROJ_ID, s.STATUS_TEXT " & _
'         "FROM PROJECT p " & _
'         "LEFT JOIN STATUS_LU s ON s.STATUS_CD = p.STATUS_CD " & _
'         "WHERE p.PROJ_ID = [OFM_KEY_VALUE]")
'
' Delete:
'
'     Ofm_Delete Get_DB_Schema(), cTableName, cKeyField, Me!PROJ_OPTN_ID
'
' Print a starter ConfigureFields routine to the Immediate Window:
'
'     Ofm_Debug_PrintFieldScaffold "APP_RECORD"
'
' Print a starter form-module scaffold to the Immediate Window:
'
'     Ofm_Debug_PrintFormModuleScaffold "APP_RECORD"
'
' The CRUD/load helpers place the main form and business inputs first and accept
' DSN as an optional final argument. If omitted, DSN defaults to Get_DB_DSN().
'
'
' Sequence-first insert model
' ---------------------------
' This module supports Oracle inserts where the primary key is obtained first using:
'
'     sequence_name.NEXTVAL
'
' The generated value is written into the key control and then included directly in
' the INSERT statement.
'
' This avoids the fragile Oracle pattern of trying to fetch a generated key after
' insert via CURRVAL or by relying on a trigger-generated value across connections.
'
'
' Checkbox mapping support
' ------------------------
' This module supports checkbox control translation through clsOracleFormField using:
'
'     - ControlKind
'     - UseCustomBooleanMapping
'     - CheckedValue
'     - UncheckedValue
'
' This allows Access checkbox values to map cleanly to Oracle storage formats such
' as:
'
'     - Y / N
'     - 1 / 0
'     - Y / NULL
'
'
' Read/write separation support
' -----------------------------
' This module can load a form from richer Oracle read models while still writing only
' to the base table passed into the CRUD helpers.
'
' This is useful for highly normalized schemas where forms need lookup labels,
' joined display values, or view-based reads, but inserts and updates should still
' target one base table.
'
' Use:
'
'     - clsOracleFormField.LoadFieldName to map a returned column/alias to a control
'     - Ofm_LoadFormBySql for arbitrary joined or aliased read SQL
'     - Ofm_SaveRecord / Ofm_Insert / Ofm_Update reloadSql to refresh from a richer
'       read model after save
'
' If reloadSql is supplied, [OFM_KEY_VALUE] is replaced with the saved key value.
'
'
' Dependencies
' ------------
' Depends on:
'
'     - modOracleDataAccess
'     - clsOracleFormField
'
'
' Common callers
' --------------
' Intended callers are:
'
'     - unbound Access forms
'
'
' Design notes
' ------------
' This module is intentionally generic.
'
' It should not contain:
'
'     - form-specific business rules
'     - form-specific UI messaging
'     - application-specific authorization logic
'
' Those belong in the form module or in business-specific modules.
'
'
' Version
' -------
' v1
'
'====================================================================================

Option Compare Database
Option Explicit

Private Const cModuleName As String = "modOracleFormEngine"
Private Const cReloadKeyToken As String = "[OFM_KEY_VALUE]"

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
        Case "CHAR", "NCHAR", "VARCHAR2", "NVARCHAR2", "VARCHAR", "CLOB", "NCLOB"
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
            If Not isRequired Then
                Debug.Print "    f.NullIfBlank = True"
            End If
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
            If Not isRequired Then
                Debug.Print "    f.NullIfBlank = True"
            End If
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
    Debug.Print "Private Sub btnSave_Click()"
    Debug.Print ""

    If Len(keyFieldName) > 0 Then
        Debug.Print "    Dim savedKey As Variant"
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
    Else
        Debug.Print "    ' Review: add custom save logic here if the form does not use a single-column primary key."
    End If

    Debug.Print ""
    Debug.Print "    MsgBox ""Record saved."", vbInformation"
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
        Debug.Print "    DoCmd.Close acForm, Me.Name"
    Else
        Debug.Print "    ' Review: add custom delete logic here if the form does not use a single-column primary key."
    End If

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
' Field definition helpers
'------------------------------------------------------------------------------------

Public Function Ofm_AddField( _
    ByRef fields As Collection, _
    ByVal DbFieldName As String, _
    ByVal ControlName As String, _
    Optional ByVal IsKey As Boolean = False, _
    Optional ByVal IsInsertable As Boolean = True, _
    Optional ByVal IsUpdatable As Boolean = True _
) As clsOracleFormField

    Dim f As clsOracleFormField

    Set f = New clsOracleFormField
    f.DbFieldName = DbFieldName
    f.ControlName = ControlName
    f.IsKey = IsKey
    f.IsInsertable = IsInsertable
    f.IsUpdatable = IsUpdatable

    fields.Add f
    Set Ofm_AddField = f

End Function

Public Function Ofm_GetKeyField(ByRef fields As Collection) As clsOracleFormField

    Dim f As clsOracleFormField

    For Each f In fields
        If f.IsKey Then
            Set Ofm_GetKeyField = f
            Exit Function
        End If
    Next f

    Err.Raise vbObjectError + 5000, cModuleName & ".Ofm_GetKeyField", "No key field is defined."

End Function

Public Function Ofm_GetFieldByControlName( _
    ByRef fields As Collection, _
    ByVal ControlName As String _
) As clsOracleFormField

    Dim f As clsOracleFormField

    For Each f In fields
        If StrComp(f.ControlName, ControlName, vbTextCompare) = 0 Then
            Set Ofm_GetFieldByControlName = f
            Exit Function
        End If
    Next f

    Err.Raise vbObjectError + 5001, cModuleName & ".Ofm_GetFieldByControlName", _
              "Field not found for control name: " & ControlName

End Function

Public Function Ofm_GetFieldByDbFieldName( _
    ByRef fields As Collection, _
    ByVal DbFieldName As String _
) As clsOracleFormField

    Dim f As clsOracleFormField

    For Each f In fields
        If StrComp(f.DbFieldName, DbFieldName, vbTextCompare) = 0 Then
            Set Ofm_GetFieldByDbFieldName = f
            Exit Function
        End If
    Next f

    Err.Raise vbObjectError + 5002, cModuleName & ".Ofm_GetFieldByDbFieldName", _
              "Field not found for database field name: " & DbFieldName

End Function

Public Function Ofm_GetSelectList(ByRef fields As Collection) As String

    Dim f As clsOracleFormField
    Dim s As String

    For Each f In fields
        If Len(s) > 0 Then s = s & ", "
        s = s & f.LoadFieldName
    Next f

    Ofm_GetSelectList = s

End Function

'------------------------------------------------------------------------------------
' Value translation helpers
'------------------------------------------------------------------------------------

Public Function Ofm_ValuesEqual(ByVal v1 As Variant, ByVal v2 As Variant) As Boolean

    If IsNull(v1) And IsNull(v2) Then
        Ofm_ValuesEqual = True
    ElseIf IsNull(v1) Xor IsNull(v2) Then
        Ofm_ValuesEqual = False
    Else
        Ofm_ValuesEqual = (v1 = v2)
    End If

End Function

Public Function Ofm_DbToControlValue( _
    ByRef fieldDef As clsOracleFormField, _
    ByVal dbValue As Variant _
) As Variant

    Select Case UCase$(fieldDef.ControlKind)

        Case "CHECKBOX"
            If fieldDef.UseCustomBooleanMapping Then
                If IsNull(dbValue) Then
                    Ofm_DbToControlValue = Null
                ElseIf Ofm_ValuesEqual(dbValue, fieldDef.CheckedValue) Then
                    Ofm_DbToControlValue = True
                Else
                    Ofm_DbToControlValue = False
                End If
            Else
                Ofm_DbToControlValue = dbValue
            End If

        Case Else
            Ofm_DbToControlValue = dbValue

    End Select

End Function

Public Function Ofm_ControlToDbValue( _
    ByRef fieldDef As clsOracleFormField, _
    ByVal controlValue As Variant _
) As Variant

    Select Case UCase$(fieldDef.ControlKind)

        Case "CHECKBOX"
            If fieldDef.UseCustomBooleanMapping Then
                If IsNull(controlValue) Then
                    Ofm_ControlToDbValue = Null
                ElseIf CBool(controlValue) Then
                    Ofm_ControlToDbValue = fieldDef.CheckedValue
                Else
                    Ofm_ControlToDbValue = fieldDef.UncheckedValue
                End If
            Else
                Ofm_ControlToDbValue = controlValue
            End If

        Case Else
            Ofm_ControlToDbValue = controlValue

    End Select

End Function

Public Function Ofm_GetControlValue( _
    ByRef frm As Access.Form, _
    ByRef fieldDef As clsOracleFormField _
) As Variant

    Dim v As Variant

    v = frm.Controls(fieldDef.ControlName).Value
    v = Ofm_ControlToDbValue(fieldDef, v)
    Ofm_GetControlValue = fieldDef.GetNormalizedValue(v)

End Function

Public Sub Ofm_SetControlValue( _
    ByRef frm As Access.Form, _
    ByRef fieldDef As clsOracleFormField, _
    ByVal dbValue As Variant _
)

    frm.Controls(fieldDef.ControlName).Value = Ofm_DbToControlValue(fieldDef, dbValue)

End Sub

'------------------------------------------------------------------------------------
' List / combo helpers
'------------------------------------------------------------------------------------

Private Function Ofm_ValueListCell(ByVal v As Variant) As String

    Dim s As String

    If IsNull(v) Then Exit Function

    s = CStr(v)
    s = Replace$(s, ";", ",")
    s = Replace$(s, vbCr, " ")
    s = Replace$(s, vbLf, " ")

    Ofm_ValueListCell = s

End Function

Private Function Ofm_BuildBlankValueListRow( _
    ByVal columnCount As Long, _
    Optional ByVal displayColumn As Long = 2, _
    Optional ByVal blankCaption As String = "" _
) As String

    Dim i As Long
    Dim s As String

    If columnCount <= 0 Then columnCount = 1
    If displayColumn <= 0 Then displayColumn = 1

    For i = 1 To columnCount
        If i > 1 Then s = s & ";"

        If i = displayColumn Then
            s = s & Ofm_ValueListCell(blankCaption)
        End If
    Next i

    Ofm_BuildBlankValueListRow = s

End Function

Private Function Ofm_BuildColumnHeadRow( _
    ByVal rowData As Object, _
    ByVal columnCount As Long, _
    Optional ByVal columnHeadCaptions As String = "" _
) As String

    Dim captions() As String
    Dim keys As Variant
    Dim i As Long
    Dim s As String
    Dim captionValue As String

    If columnCount <= 0 Then columnCount = 1

    If Len(Trim$(columnHeadCaptions)) > 0 Then
        captions = Split(columnHeadCaptions, ";")

        For i = 1 To columnCount
            If i > 1 Then s = s & ";"

            If (i - 1) <= UBound(captions) Then
                captionValue = Trim$(captions(i - 1))
            Else
                captionValue = vbNullString
            End If

            s = s & Ofm_ValueListCell(captionValue)
        Next i

        Ofm_BuildColumnHeadRow = s
        Exit Function
    End If

    If rowData Is Nothing Then
        For i = 1 To columnCount
            If i > 1 Then s = s & ";"
            s = s & Ofm_ValueListCell("Column " & CStr(i))
        Next i

        Ofm_BuildColumnHeadRow = s
        Exit Function
    End If

    keys = rowData.Keys

    For i = LBound(keys) To UBound(keys)
        If i > LBound(keys) Then s = s & ";"
        s = s & Ofm_ValueListCell(keys(i))
    Next i

    Ofm_BuildColumnHeadRow = s

End Function

Private Function Ofm_BuildValueListRow(ByVal rowData As Object) As String

    Dim keys As Variant
    Dim i As Long
    Dim s As String

    keys = rowData.Keys

    For i = LBound(keys) To UBound(keys)
        If i > LBound(keys) Then s = s & ";"
        s = s & Ofm_ValueListCell(rowData(keys(i)))
    Next i

    Ofm_BuildValueListRow = s

End Function

Private Function Ofm_WrapLookupSqlWithRowLimit(ByVal sSQL As String, ByVal maxRows As Long) As String

    If maxRows <= 0 Then
        Ofm_WrapLookupSqlWithRowLimit = sSQL
    Else
        Ofm_WrapLookupSqlWithRowLimit = _
            "SELECT * FROM (" & vbCrLf & _
            sSQL & vbCrLf & _
            ") lookup_src " & _
            "WHERE ROWNUM <= " & CStr(maxRows + 1)
    End If

End Function

Private Sub Ofm_PrepareListControl( _
    ByVal ctl As Object, _
    ByVal columnCount As Long, _
    ByVal boundColumn As Long, _
    Optional ByVal showColumnHeads As Boolean = False, _
    Optional ByVal columnWidths As String = "" _
)

    If columnCount <= 0 Then columnCount = 1
    If boundColumn <= 0 Then boundColumn = 1

    ctl.RowSourceType = "Value List"
    ctl.RowSource = vbNullString
    ctl.ColumnCount = columnCount
    ctl.BoundColumn = boundColumn

    On Error Resume Next
    ctl.ControlSource = vbNullString
    ctl.MultiSelect = 0
    ctl.ColumnHeads = showColumnHeads
    ctl.ListStyle = 0
    ctl.Enabled = True
    ctl.Locked = False
    ctl.TabStop = True
    ctl.ListIndex = -1
    On Error GoTo 0

    If Len(columnWidths) > 0 Then
        ctl.ColumnWidths = columnWidths
    End If

End Sub

Public Sub Ofm_LoadListControlBySql( _
    ByRef frm As Access.Form, _
    ByVal controlName As String, _
    ByVal sSQL As String, _
    Optional ByVal boundColumn As Long = 1, _
    Optional ByVal displayColumn As Long = 2, _
    Optional ByVal includeBlankRow As Boolean = False, _
    Optional ByVal blankCaption As String = "", _
    Optional ByVal columnWidths As String = "", _
    Optional ByVal showColumnHeads As Boolean = False, _
    Optional ByVal columnHeadCaptions As String = "", _
    Optional ByVal maxRows As Long = 2500, _
    Optional ByVal dsn As String = "" _
)

    Dim rows As Collection
    Dim rowData As Object
    Dim ctl As Object
    Dim columnCount As Long

    If Len(Trim$(controlName)) = 0 Then
        Err.Raise vbObjectError + 5070, cModuleName & ".Ofm_LoadListControlBySql", "Control name cannot be blank."
    End If

    If Len(Trim$(sSQL)) = 0 Then
        Err.Raise vbObjectError + 5071, cModuleName & ".Ofm_LoadListControlBySql", "Lookup SQL cannot be blank."
    End If

    If maxRows <= 0 Then maxRows = 2500

    Set rows = PTQ_GetRows(Ofm_WrapLookupSqlWithRowLimit(sSQL, maxRows), dsn)

    If rows.Count > maxRows Then
        Err.Raise vbObjectError + 5074, cModuleName & ".Ofm_LoadListControlBySql", _
                  "Lookup for control " & controlName & " returned " & CStr(rows.Count) & _
                  " rows, which exceeds the configured limit of " & CStr(maxRows) & "."
    End If

    Set ctl = frm.Controls(controlName)

    If rows.Count > 0 Then
        columnCount = rows(1).Count
    Else
        columnCount = IIf(displayColumn > boundColumn, displayColumn, boundColumn)
    End If

    Call Ofm_PrepareListControl(ctl, columnCount, boundColumn, showColumnHeads, columnWidths)

    If showColumnHeads Then
        If rows.Count > 0 Then
            ctl.AddItem Ofm_BuildColumnHeadRow(rows(1), columnCount, columnHeadCaptions)
        Else
            ctl.AddItem Ofm_BuildColumnHeadRow(Nothing, columnCount, columnHeadCaptions)
        End If
    End If

    If includeBlankRow Then
        ctl.AddItem Ofm_BuildBlankValueListRow(columnCount, displayColumn, blankCaption)
    End If

    For Each rowData In rows
        ctl.AddItem Ofm_BuildValueListRow(rowData)
    Next rowData

End Sub

Public Sub Ofm_LoadLookupControl( _
    ByRef frm As Access.Form, _
    ByRef fieldDef As clsOracleFormField, _
    Optional ByVal dsn As String = "" _
)

    If Not fieldDef.HasLookupSql Then
        Err.Raise vbObjectError + 5072, cModuleName & ".Ofm_LoadLookupControl", _
                  "Field " & fieldDef.ControlName & " does not define LookupSql."
    End If

    If Not (fieldDef.IsComboBox Or fieldDef.IsListBox) Then
        Err.Raise vbObjectError + 5073, cModuleName & ".Ofm_LoadLookupControl", _
                  "Field " & fieldDef.ControlName & " is not marked as COMBOBOX or LISTBOX."
    End If

    Ofm_LoadListControlBySql _
        frm:=frm, _
        controlName:=fieldDef.ControlName, _
        sSQL:=fieldDef.LookupSql, _
        boundColumn:=fieldDef.LookupBoundColumn, _
        displayColumn:=fieldDef.LookupDisplayColumn, _
        includeBlankRow:=fieldDef.LookupIncludeBlankRow, _
        blankCaption:=fieldDef.LookupBlankCaption, _
        columnWidths:=fieldDef.LookupColumnWidths, _
        showColumnHeads:=fieldDef.LookupShowColumnHeads, _
        columnHeadCaptions:=fieldDef.LookupColumnHeadCaptions, _
        maxRows:=fieldDef.LookupMaxRows, _
        dsn:=dsn

End Sub

Public Sub Ofm_LoadLookupControls( _
    ByRef frm As Access.Form, _
    ByRef fields As Collection, _
    Optional ByVal dsn As String = "" _
)

    Dim f As clsOracleFormField

    For Each f In fields
        If f.HasLookupSql Then
            Ofm_LoadLookupControl frm, f, dsn
        End If
    Next f

End Sub

'------------------------------------------------------------------------------------
' Snapshot / dirty helpers
'------------------------------------------------------------------------------------

Public Sub Ofm_SnapshotValues( _
    ByRef frm As Access.Form, _
    ByRef fields As Collection, _
    ByRef originalValues As Object _
)

    Dim f As clsOracleFormField

    If originalValues Is Nothing Then
        Set originalValues = CreateObject("Scripting.Dictionary")
    Else
        originalValues.RemoveAll
    End If

    For Each f In fields
        originalValues(f.ControlName) = Ofm_GetControlValue(frm, f)
    Next f

End Sub

Public Function Ofm_IsDirty( _
    ByRef frm As Access.Form, _
    ByRef fields As Collection, _
    ByRef originalValues As Object _
) As Boolean

    Dim f As clsOracleFormField
    Dim currentValue As Variant
    Dim originalValue As Variant

    For Each f In fields
        currentValue = Ofm_GetControlValue(frm, f)
        originalValue = originalValues(f.ControlName)

        If Not Ofm_ValuesEqual(currentValue, originalValue) Then
            Ofm_IsDirty = True
            Exit Function
        End If
    Next f

End Function

Public Function Ofm_GetChangedFields( _
    ByRef frm As Access.Form, _
    ByRef fields As Collection, _
    ByRef originalValues As Object _
) As Collection

    Dim f As clsOracleFormField
    Dim result As Collection
    Dim currentValue As Variant
    Dim originalValue As Variant

    Set result = New Collection

    For Each f In fields
        currentValue = Ofm_GetControlValue(frm, f)
        originalValue = originalValues(f.ControlName)

        If Not Ofm_ValuesEqual(currentValue, originalValue) Then
            result.Add f
        End If
    Next f

    Set Ofm_GetChangedFields = result

End Function

'------------------------------------------------------------------------------------
' Validation helpers
'------------------------------------------------------------------------------------

Public Function Ofm_ValidateRequiredFields( _
    ByRef frm As Access.Form, _
    ByRef fields As Collection, _
    Optional ByVal setFocusToFirstInvalid As Boolean = True _
) As Boolean

    Dim f As clsOracleFormField
    Dim v As Variant

    For Each f In fields
        v = Ofm_GetControlValue(frm, f)

        If f.IsMissingRequiredValue(v) Then
            If setFocusToFirstInvalid Then
                frm.Controls(f.ControlName).SetFocus
            End If

            Err.Raise vbObjectError + 5010, cModuleName & ".Ofm_ValidateRequiredFields", _
                      "Required field is missing: " & f.ControlName
        End If
    Next f

    Ofm_ValidateRequiredFields = True

End Function

'------------------------------------------------------------------------------------
' Load / new-record helpers
'------------------------------------------------------------------------------------

Private Function Ofm_BuildLoadSql( _
    ByVal schemaName As String, _
    ByVal objectName As String, _
    ByVal keyField As String, _
    ByVal keyValue As Variant, _
    ByRef fields As Collection _
) As String

    Ofm_BuildLoadSql = _
        "SELECT " & Ofm_GetSelectList(fields) & " " & _
        "FROM " & Ofm_GetQualifiedObjectName(schemaName, objectName) & " " & _
        "WHERE " & keyField & " = " & Ofm_SqlValue(keyValue)

End Function

Private Function Ofm_ResolveReloadSql( _
    ByVal schemaName As String, _
    ByVal objectName As String, _
    ByVal keyField As String, _
    ByVal keyValue As Variant, _
    ByRef fields As Collection, _
    Optional ByVal reloadSql As String = "" _
) As String

    If Len(Trim$(reloadSql)) = 0 Then
        Ofm_ResolveReloadSql = Ofm_BuildLoadSql(schemaName, objectName, keyField, keyValue, fields)
    Else
        Ofm_ResolveReloadSql = Replace$(reloadSql, cReloadKeyToken, Ofm_SqlValue(keyValue), , , vbTextCompare)
    End If

End Function

Private Sub Ofm_LoadFormFromRow( _
    ByRef frm As Access.Form, _
    ByVal rowData As Object, _
    ByRef fields As Collection, _
    ByRef originalValues As Object, _
    ByVal sourceProcName As String _
)

    Dim f As clsOracleFormField
    Dim sLoadFieldName As String

    If rowData Is Nothing Then
        Err.Raise vbObjectError + 5020, sourceProcName, "No row found."
    End If

    For Each f In fields
        sLoadFieldName = f.LoadFieldName

        If rowData.Exists(sLoadFieldName) Then
            Ofm_SetControlValue frm, f, rowData(sLoadFieldName)
        Else
            Err.Raise vbObjectError + 5021, sourceProcName, _
                      "Returned row does not contain expected load field: " & sLoadFieldName & _
                      " (control: " & f.ControlName & ")."
        End If
    Next f

    Ofm_SnapshotValues frm, fields, originalValues

End Sub

Public Sub Ofm_LoadForm( _
    ByRef frm As Access.Form, _
    ByVal schemaName As String, _
    ByVal tableName As String, _
    ByVal keyField As String, _
    ByVal keyValue As Variant, _
    ByRef fields As Collection, _
    ByRef originalValues As Object, _
    Optional ByVal dsn As String = "" _
)

    Dim sSQL As String

    sSQL = Ofm_BuildLoadSql(schemaName, tableName, keyField, keyValue, fields)

    Ofm_LoadFormBySql frm, sSQL, fields, originalValues, cModuleName & ".Ofm_LoadForm", dsn

End Sub

Public Sub Ofm_LoadFormBySql( _
    ByRef frm As Access.Form, _
    ByVal sSQL As String, _
    ByRef fields As Collection, _
    ByRef originalValues As Object, _
    Optional ByVal sourceProcName As String = "", _
    Optional ByVal dsn As String = "" _
)

    Dim rowData As Object

    If Len(Trim$(sSQL)) = 0 Then
        Err.Raise vbObjectError + 5022, cModuleName & ".Ofm_LoadFormBySql", "Load SQL cannot be blank."
    End If

    If Len(sourceProcName) = 0 Then
        sourceProcName = cModuleName & ".Ofm_LoadFormBySql"
    End If

    Set rowData = PTQ_GetRow(sSQL, dsn)

    Ofm_LoadFormFromRow frm, rowData, fields, originalValues, sourceProcName

End Sub

Public Sub Ofm_InitNewForm( _
    ByRef frm As Access.Form, _
    ByRef fields As Collection, _
    ByRef originalValues As Object _
)

    Dim f As clsOracleFormField
    Dim v As Variant

    For Each f In fields

        v = Null

        If Not IsNull(f.DefaultValue) Then
            v = f.DefaultValue
        ElseIf UCase$(f.ControlKind) = "CHECKBOX" Then
            If f.UseCustomBooleanMapping Then
                v = False
            End If
        End If

        frm.Controls(f.ControlName).Value = v
    Next f

    Ofm_SnapshotValues frm, fields, originalValues

End Sub

'------------------------------------------------------------------------------------
' SQL helpers
'------------------------------------------------------------------------------------

Public Function Ofm_GetQualifiedObjectName( _
    ByVal schemaName As String, _
    ByVal objectName As String _
) As String

    schemaName = Trim$(schemaName)
    objectName = Trim$(objectName)

    If Len(objectName) = 0 Then
        Err.Raise vbObjectError + 5030, cModuleName & ".Ofm_GetQualifiedObjectName", "Object name cannot be blank."
    End If

    If Len(schemaName) = 0 Then
        Ofm_GetQualifiedObjectName = objectName
    Else
        Ofm_GetQualifiedObjectName = schemaName & "." & objectName
    End If

End Function

Public Function Ofm_SqlValue(ByVal v As Variant) As String

    If IsNull(v) Then
        Ofm_SqlValue = "NULL"
    ElseIf IsDate(v) Then
        Ofm_SqlValue = SqlDateOrNull(v)
    ElseIf VarType(v) = vbString Then
        Ofm_SqlValue = SqlStringOrNull(v)
    ElseIf VarType(v) = vbBoolean Then
        Ofm_SqlValue = SqlBooleanNumber(v)
    Else
        Ofm_SqlValue = SqlNumberOrNull(v)
    End If

End Function

Public Function Ofm_BuildUpdateSql( _
    ByRef frm As Access.Form, _
    ByVal schemaName As String, _
    ByVal tableName As String, _
    ByVal keyField As String, _
    ByVal keyValue As Variant, _
    ByRef fields As Collection, _
    ByRef originalValues As Object _
) As String

    Dim f As clsOracleFormField
    Dim setClause As String
    Dim currentValue As Variant
    Dim originalValue As Variant

    For Each f In fields

        If f.IsUpdatable And Not f.IsKey Then

            currentValue = Ofm_GetControlValue(frm, f)
            originalValue = originalValues(f.ControlName)

            If Not Ofm_ValuesEqual(currentValue, originalValue) Then
                If Len(setClause) > 0 Then setClause = setClause & ", "
                setClause = setClause & f.DbFieldName & " = " & Ofm_SqlValue(currentValue)
            End If

        End If

    Next f

    If Len(setClause) = 0 Then Exit Function

    Ofm_BuildUpdateSql = _
        "UPDATE " & Ofm_GetQualifiedObjectName(schemaName, tableName) & " " & _
        "SET " & setClause & " " & _
        "WHERE " & keyField & " = " & Ofm_SqlValue(keyValue)

End Function

Public Function Ofm_BuildInsertSql( _
    ByRef frm As Access.Form, _
    ByVal schemaName As String, _
    ByVal tableName As String, _
    ByRef fields As Collection _
) As String

    Dim f As clsOracleFormField
    Dim fieldList As String
    Dim valueList As String
    Dim currentValue As Variant

    For Each f In fields

        If f.IsInsertable Then
            currentValue = Ofm_GetControlValue(frm, f)

            If Len(fieldList) > 0 Then
                fieldList = fieldList & ", "
                valueList = valueList & ", "
            End If

            fieldList = fieldList & f.DbFieldName
            valueList = valueList & Ofm_SqlValue(currentValue)
        End If

    Next f

    If Len(fieldList) = 0 Then
        Err.Raise vbObjectError + 5031, cModuleName & ".Ofm_BuildInsertSql", _
                  "No insertable fields are defined."
    End If

    Ofm_BuildInsertSql = _
        "INSERT INTO " & Ofm_GetQualifiedObjectName(schemaName, tableName) & _
        " (" & fieldList & ") " & _
        "VALUES (" & valueList & ")"

End Function

Public Function Ofm_BuildDeleteSql( _
    ByVal schemaName As String, _
    ByVal tableName As String, _
    ByVal keyField As String, _
    ByVal keyValue As Variant _
) As String

    If IsNull(keyValue) Then
        Err.Raise vbObjectError + 5032, cModuleName & ".Ofm_BuildDeleteSql", _
                  "Delete requires a non-null key value."
    End If

    Ofm_BuildDeleteSql = _
        "DELETE FROM " & Ofm_GetQualifiedObjectName(schemaName, tableName) & " " & _
        "WHERE " & keyField & " = " & Ofm_SqlValue(keyValue)

End Function

'------------------------------------------------------------------------------------
' CRUD execution helpers
'------------------------------------------------------------------------------------

Public Function Ofm_Insert( _
    ByRef frm As Access.Form, _
    ByVal schemaName As String, _
    ByVal tableName As String, _
    ByVal keyField As String, _
    ByRef fields As Collection, _
    ByRef originalValues As Object, _
    Optional ByVal sequenceName As String = "", _
    Optional ByVal reloadAfterInsert As Boolean = True, _
    Optional ByVal reloadSql As String = "", _
    Optional ByVal dsn As String = "" _
) As Variant

    Dim sSQL As String
    Dim sReloadSql As String
    Dim keyDef As clsOracleFormField
    Dim newKeyValue As Variant

    Call Ofm_ValidateRequiredFields(frm, fields, True)

    Set keyDef = Ofm_GetKeyField(fields)

    If Len(Trim$(sequenceName)) > 0 Then
        newKeyValue = Oracle_GetNextSequenceValue(schemaName, sequenceName, dsn)
        frm.Controls(keyDef.ControlName).Value = newKeyValue
    Else
        newKeyValue = Ofm_GetControlValue(frm, keyDef)
        If IsNull(newKeyValue) Then
            Err.Raise vbObjectError + 5040, cModuleName & ".Ofm_Insert", _
                      "Insert requires a key value or a sequence name."
        End If
    End If

    sSQL = Ofm_BuildInsertSql(frm, schemaName, tableName, fields)
    PTQ_Execute sSQL, , dsn

    If reloadAfterInsert Then
        sReloadSql = Ofm_ResolveReloadSql(schemaName, tableName, keyField, newKeyValue, fields, reloadSql)
        Ofm_LoadFormBySql frm, sReloadSql, fields, originalValues, cModuleName & ".Ofm_Insert", dsn
    Else
        Ofm_SnapshotValues frm, fields, originalValues
    End If

    Ofm_Insert = newKeyValue

End Function

Public Sub Ofm_Update( _
    ByRef frm As Access.Form, _
    ByVal schemaName As String, _
    ByVal tableName As String, _
    ByVal keyField As String, _
    ByVal keyValue As Variant, _
    ByRef fields As Collection, _
    ByRef originalValues As Object, _
    Optional ByVal reloadAfterUpdate As Boolean = True, _
    Optional ByVal reloadSql As String = "", _
    Optional ByVal dsn As String = "" _
)

    Dim sSQL As String
    Dim sReloadSql As String

    Call Ofm_ValidateRequiredFields(frm, fields, True)

    sSQL = Ofm_BuildUpdateSql(frm, schemaName, tableName, keyField, keyValue, fields, originalValues)

    If Len(sSQL) = 0 Then Exit Sub

    PTQ_Execute sSQL, , dsn

    If reloadAfterUpdate Then
        sReloadSql = Ofm_ResolveReloadSql(schemaName, tableName, keyField, keyValue, fields, reloadSql)
        Ofm_LoadFormBySql frm, sReloadSql, fields, originalValues, cModuleName & ".Ofm_Update", dsn
    Else
        Ofm_SnapshotValues frm, fields, originalValues
    End If

End Sub

Public Sub Ofm_Delete( _
    ByVal schemaName As String, _
    ByVal tableName As String, _
    ByVal keyField As String, _
    ByVal keyValue As Variant, _
    Optional ByVal dsn As String = "" _
)

    Dim sSQL As String

    sSQL = Ofm_BuildDeleteSql(schemaName, tableName, keyField, keyValue)
    PTQ_Execute sSQL, , dsn

End Sub

Public Function Ofm_SaveRecord( _
    ByRef frm As Access.Form, _
    ByVal schemaName As String, _
    ByVal tableName As String, _
    ByVal keyField As String, _
    ByRef fields As Collection, _
    ByRef originalValues As Object, _
    ByVal isNewRecord As Boolean, _
    Optional ByVal sequenceName As String = "", _
    Optional ByVal reloadAfterSave As Boolean = True, _
    Optional ByVal reloadSql As String = "", _
    Optional ByVal dsn As String = "" _
) As Variant

    Dim keyDef As clsOracleFormField
    Dim keyValue As Variant

    Set keyDef = Ofm_GetKeyField(fields)

    If isNewRecord Then
        Ofm_SaveRecord = Ofm_Insert( _
            frm:=frm, _
            schemaName:=schemaName, _
            tableName:=tableName, _
            keyField:=keyField, _
            fields:=fields, _
            originalValues:=originalValues, _
            sequenceName:=sequenceName, _
            reloadAfterInsert:=reloadAfterSave, _
            reloadSql:=reloadSql, _
            dsn:=dsn)
    Else
        keyValue = Ofm_GetControlValue(frm, keyDef)

        If IsNull(keyValue) Then
            Err.Raise vbObjectError + 5050, cModuleName & ".Ofm_SaveRecord", _
                      "Existing record save requires a non-null key value."
        End If

        Ofm_Update _
            frm:=frm, _
            schemaName:=schemaName, _
            tableName:=tableName, _
            keyField:=keyField, _
            keyValue:=keyValue, _
            fields:=fields, _
            originalValues:=originalValues, _
            reloadAfterUpdate:=reloadAfterSave, _
            reloadSql:=reloadSql, _
            dsn:=dsn

        Ofm_SaveRecord = keyValue
    End If

End Function

'------------------------------------------------------------------------------------
' Debug helpers
'------------------------------------------------------------------------------------

Public Function Ofm_DebugFieldSummary(ByRef fields As Collection) As String

    Dim f As clsOracleFormField
    Dim s As String

    For Each f In fields
        If Len(s) > 0 Then s = s & vbCrLf
        s = s & f.DebugSummary
    Next f

    Ofm_DebugFieldSummary = s

End Function
