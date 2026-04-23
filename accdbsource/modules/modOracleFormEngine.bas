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
' Load an existing record:
'
'     Ofm_LoadForm Me, Get_DB_DSN(), Get_DB_Schema(), cTableName, cKeyField, keyValue, mFields, mOriginalValues
'
' Load an existing record from arbitrary SQL:
'
'     Ofm_LoadFormBySql Me, Get_DB_DSN(), _
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
' Save:
'
'     savedKey = Ofm_SaveRecord(Me, Get_DB_DSN(), Get_DB_Schema(), cTableName, cKeyField, mFields, mOriginalValues, mIsNewRecord, cSequenceName, True)
'
' Save and reload from a read-model SQL statement:
'
'     savedKey = Ofm_SaveRecord(Me, Get_DB_DSN(), Get_DB_Schema(), cTableName, cKeyField, mFields, mOriginalValues, mIsNewRecord, cSequenceName, True, _
'         "SELECT p.PROJ_ID, s.STATUS_TEXT " & _
'         "FROM PROJECT p " & _
'         "LEFT JOIN STATUS_LU s ON s.STATUS_CD = p.STATUS_CD " & _
'         "WHERE p.PROJ_ID = [OFM_KEY_VALUE]")
'
' Delete:
'
'     Ofm_Delete Get_DB_DSN(), Get_DB_Schema(), cTableName, cKeyField, Me!PROJ_OPTN_ID
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
    ByVal dsn As String, _
    ByVal schemaName As String, _
    ByVal tableName As String, _
    ByVal keyField As String, _
    ByVal keyValue As Variant, _
    ByRef fields As Collection, _
    ByRef originalValues As Object _
)

    Dim sSQL As String

    sSQL = Ofm_BuildLoadSql(schemaName, tableName, keyField, keyValue, fields)

    Ofm_LoadFormBySql frm, dsn, sSQL, fields, originalValues, cModuleName & ".Ofm_LoadForm"

End Sub

Public Sub Ofm_LoadFormBySql( _
    ByRef frm As Access.Form, _
    ByVal dsn As String, _
    ByVal sSQL As String, _
    ByRef fields As Collection, _
    ByRef originalValues As Object, _
    Optional ByVal sourceProcName As String = "" _
)

    Dim rowData As Object

    If Len(Trim$(sSQL)) = 0 Then
        Err.Raise vbObjectError + 5022, cModuleName & ".Ofm_LoadFormBySql", "Load SQL cannot be blank."
    End If

    If Len(sourceProcName) = 0 Then
        sourceProcName = cModuleName & ".Ofm_LoadFormBySql"
    End If

    Set rowData = PTQ_GetRow(dsn, sSQL)

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
    ByVal dsn As String, _
    ByVal schemaName As String, _
    ByVal tableName As String, _
    ByVal keyField As String, _
    ByRef fields As Collection, _
    ByRef originalValues As Object, _
    Optional ByVal sequenceName As String = "", _
    Optional ByVal reloadAfterInsert As Boolean = True, _
    Optional ByVal reloadSql As String = "" _
) As Variant

    Dim sSQL As String
    Dim sReloadSql As String
    Dim keyDef As clsOracleFormField
    Dim newKeyValue As Variant

    Call Ofm_ValidateRequiredFields(frm, fields, True)

    Set keyDef = Ofm_GetKeyField(fields)

    If Len(Trim$(sequenceName)) > 0 Then
        newKeyValue = Oracle_GetNextSequenceValue(dsn, schemaName, sequenceName)
        frm.Controls(keyDef.ControlName).Value = newKeyValue
    Else
        newKeyValue = Ofm_GetControlValue(frm, keyDef)
        If IsNull(newKeyValue) Then
            Err.Raise vbObjectError + 5040, cModuleName & ".Ofm_Insert", _
                      "Insert requires a key value or a sequence name."
        End If
    End If

    sSQL = Ofm_BuildInsertSql(frm, schemaName, tableName, fields)
    PTQ_Execute dsn, sSQL

    If reloadAfterInsert Then
        sReloadSql = Ofm_ResolveReloadSql(schemaName, tableName, keyField, newKeyValue, fields, reloadSql)
        Ofm_LoadFormBySql frm, dsn, sReloadSql, fields, originalValues, cModuleName & ".Ofm_Insert"
    Else
        Ofm_SnapshotValues frm, fields, originalValues
    End If

    Ofm_Insert = newKeyValue

End Function

Public Sub Ofm_Update( _
    ByRef frm As Access.Form, _
    ByVal dsn As String, _
    ByVal schemaName As String, _
    ByVal tableName As String, _
    ByVal keyField As String, _
    ByVal keyValue As Variant, _
    ByRef fields As Collection, _
    ByRef originalValues As Object, _
    Optional ByVal reloadAfterUpdate As Boolean = True, _
    Optional ByVal reloadSql As String = "" _
)

    Dim sSQL As String
    Dim sReloadSql As String

    Call Ofm_ValidateRequiredFields(frm, fields, True)

    sSQL = Ofm_BuildUpdateSql(frm, schemaName, tableName, keyField, keyValue, fields, originalValues)

    If Len(sSQL) = 0 Then Exit Sub

    PTQ_Execute dsn, sSQL

    If reloadAfterUpdate Then
        sReloadSql = Ofm_ResolveReloadSql(schemaName, tableName, keyField, keyValue, fields, reloadSql)
        Ofm_LoadFormBySql frm, dsn, sReloadSql, fields, originalValues, cModuleName & ".Ofm_Update"
    Else
        Ofm_SnapshotValues frm, fields, originalValues
    End If

End Sub

Public Sub Ofm_Delete( _
    ByVal dsn As String, _
    ByVal schemaName As String, _
    ByVal tableName As String, _
    ByVal keyField As String, _
    ByVal keyValue As Variant _
)

    Dim sSQL As String

    sSQL = Ofm_BuildDeleteSql(schemaName, tableName, keyField, keyValue)
    PTQ_Execute dsn, sSQL

End Sub

Public Function Ofm_SaveRecord( _
    ByRef frm As Access.Form, _
    ByVal dsn As String, _
    ByVal schemaName As String, _
    ByVal tableName As String, _
    ByVal keyField As String, _
    ByRef fields As Collection, _
    ByRef originalValues As Object, _
    ByVal isNewRecord As Boolean, _
    Optional ByVal sequenceName As String = "", _
    Optional ByVal reloadAfterSave As Boolean = True, _
    Optional ByVal reloadSql As String = "" _
) As Variant

    Dim keyDef As clsOracleFormField
    Dim keyValue As Variant

    Set keyDef = Ofm_GetKeyField(fields)

    If isNewRecord Then
        Ofm_SaveRecord = Ofm_Insert( _
            frm:=frm, _
            dsn:=dsn, _
            schemaName:=schemaName, _
            tableName:=tableName, _
            keyField:=keyField, _
            fields:=fields, _
            originalValues:=originalValues, _
            sequenceName:=sequenceName, _
            reloadAfterInsert:=reloadAfterSave, _
            reloadSql:=reloadSql)
    Else
        keyValue = Ofm_GetControlValue(frm, keyDef)

        If IsNull(keyValue) Then
            Err.Raise vbObjectError + 5050, cModuleName & ".Ofm_SaveRecord", _
                      "Existing record save requires a non-null key value."
        End If

        Ofm_Update _
            frm:=frm, _
            dsn:=dsn, _
            schemaName:=schemaName, _
            tableName:=tableName, _
            keyField:=keyField, _
            keyValue:=keyValue, _
            fields:=fields, _
            originalValues:=originalValues, _
            reloadAfterUpdate:=reloadAfterSave, _
            reloadSql:=reloadSql

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
