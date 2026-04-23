# Base Form Pattern

This document describes the recommended "golden path" for building unbound Oracle
forms in this repository.

The goal is consistency.

Every new form should follow roughly the same structure so that:

- field configuration is easy to read
- load/save/delete behavior is predictable
- lookups are loaded the same way
- validation is easy to find
- debugging is easier across forms

Examples use generic names:

- table: `APP_RECORD`
- key field: `RECORD_ID`
- sequence: `APP_RECORD_SEQ`


## 1. Recommended responsibilities

For this architecture, a form should own:

- control layout
- field configuration
- form-specific validation
- form-specific lookup loading
- user-facing messages

The shared engine should own:

- loading rows from Oracle
- value normalization
- required-field validation
- dirty tracking
- insert/update/delete SQL generation
- execution of those statements

Keep the split clean:

- generic behavior in `modOracleFormEngine`
- field metadata in `clsOracleFormField`
- Oracle query/session behavior in `modOracleDataAccess`
- form-specific business rules in the form module


## 2. Recommended form lifecycle

The recommended lifecycle is:

1. `Form_Open`
   - initialize `mFields`
   - initialize `mOriginalValues`
   - call `ConfigureFields`

2. `Form_Load`
   - determine whether this is a new or existing record
   - call `Ofm_InitNewForm` or `Ofm_LoadForm` / `Ofm_LoadFormBySql`
   - call `LoadLookups`

3. Save button
   - call form-specific validation
   - call `Ofm_SaveRecord`
   - update `mIsNewRecord`
   - show success message if appropriate

4. Delete button
   - confirm delete
   - call `Ofm_Delete`
   - close the form

5. Close / cancel behavior
   - optionally check `Ofm_IsDirty`
   - prompt before discarding changes


## 3. Recommended module-level members

Every standard maintenance form should start with:

```vb
Option Compare Database
Option Explicit

Private Const cTableName As String = "APP_RECORD"
Private Const cKeyField As String = "RECORD_ID"
Private Const cSequenceName As String = "APP_RECORD_SEQ"

Private mFields As Collection
Private mOriginalValues As Object
Private mIsNewRecord As Boolean
```

If the form does not use a sequence for inserts, omit `cSequenceName` or leave it
as `vbNullString` when saving.


## 4. Recommended procedures

The standard form should usually have these procedures:

- `Form_Open`
- `Form_Load`
- `ConfigureFields`
- `LoadLookups`
- `ValidateForm`
- `btnSave_Click`
- `btnDelete_Click`
- optionally `btnClose_Click`

That gives each form one obvious place for each kind of logic.


## 5. Base template

```vb
Option Compare Database
Option Explicit

Private Const cTableName As String = "APP_RECORD"
Private Const cKeyField As String = "RECORD_ID"
Private Const cSequenceName As String = "APP_RECORD_SEQ"

Private mFields As Collection
Private mOriginalValues As Object
Private mIsNewRecord As Boolean

Private Sub Form_Open(Cancel As Integer)

    Set mFields = New Collection
    Set mOriginalValues = CreateObject("Scripting.Dictionary")

    ConfigureFields

End Sub

Private Sub Form_Load()

    Dim keyValue As Variant

    keyValue = Nz(Me.OpenArgs, vbNullString)

    If Len(CStr(keyValue)) = 0 Then
        mIsNewRecord = True
        Ofm_InitNewForm Me, mFields, mOriginalValues
    Else
        mIsNewRecord = False
        Ofm_LoadForm Me, Get_DB_Schema(), cTableName, cKeyField, keyValue, mFields, mOriginalValues
    End If

    LoadLookups

End Sub

Private Sub ConfigureFields()

    Dim f As clsOracleFormField

    Set f = Ofm_AddField(mFields, "RECORD_ID", "RECORD_ID", True, True, False)
    f.IsKey = True
    f.IsUpdatable = False
    f.IsDbGenerated = True

    Set f = Ofm_AddField(mFields, "RECORD_NAME", "txtRecordName")
    f.IsRequired = True
    f.TrimOnSave = True
    f.NullIfBlank = True

    Set f = Ofm_AddField(mFields, "STATUS_CD", "cboStatus")
    f.ControlKind = "COMBOBOX"
    f.IsRequired = True
    f.TrimOnSave = True
    f.LookupSql = _
        "SELECT STATUS_CD, STATUS_TEXT " & _
        "FROM APP_STATUS_LU " & _
        "ORDER BY STATUS_TEXT"
    f.LookupBoundColumn = 1
    f.LookupDisplayColumn = 2
    f.LookupIncludeBlankRow = True
    f.LookupColumnWidths = "0;1.5"""

    Set f = Ofm_AddField(mFields, "IS_ACTIVE_YN", "chkIsActive")
    f.ControlKind = "CHECKBOX"
    f.UseCustomBooleanMapping = True
    f.CheckedValue = "Y"
    f.UncheckedValue = "N"
    f.DefaultValue = "Y"

End Sub

Private Sub LoadLookups()
    Ofm_LoadLookupControls Me, mFields
End Sub

Private Sub ValidateForm()

    If Not IsNull(Me.txtRecordName.Value) Then
        If Len(Trim$(Me.txtRecordName.Value)) < 3 Then
            Me.txtRecordName.SetFocus
            Err.Raise vbObjectError + 9000, Me.Name & ".ValidateForm", _
                      "Record Name must be at least 3 characters."
        End If
    End If

End Sub

Private Sub btnSave_Click()

    Dim savedKey As Variant

    On Error GoTo ErrHandler

    ValidateForm

    savedKey = Ofm_SaveRecord( _
        Me, _
        Get_DB_Schema(), _
        cTableName, _
        cKeyField, _
        mFields, _
        mOriginalValues, _
        mIsNewRecord, _
        cSequenceName, _
        True)

    Me.RECORD_ID = savedKey
    mIsNewRecord = False

    MsgBox "Record saved.", vbInformation
    Exit Sub

ErrHandler:
    MsgBox Err.Description, vbExclamation

End Sub

Private Sub btnDelete_Click()

    On Error GoTo ErrHandler

    If mIsNewRecord Then Exit Sub

    If MsgBox("Delete this record?", vbQuestion + vbYesNo) = vbNo Then Exit Sub

    Ofm_Delete Get_DB_Schema(), cTableName, cKeyField, Me.RECORD_ID

    DoCmd.Close acForm, Me.Name
    Exit Sub

ErrHandler:
    MsgBox Err.Description, vbExclamation

End Sub
```


## 6. Recommended load patterns

There are two standard load patterns.

### Pattern A: Base-table load

Use:

```vb
Ofm_LoadForm Me, Get_DB_Schema(), cTableName, cKeyField, keyValue, mFields, mOriginalValues
```

Use this when:

- the form edits one base table
- the form does not need joined display fields

This should be the default starting pattern.


### Pattern B: Joined-read load

Use:

```vb
Ofm_LoadFormBySql _
    Me, _
    "SELECT r.RECORD_ID, " & _
    "       r.RECORD_NAME, " & _
    "       r.STATUS_CD, " & _
    "       s.STATUS_TEXT " & _
    "FROM APP_RECORD r " & _
    "LEFT JOIN APP_STATUS_LU s ON s.STATUS_CD = r.STATUS_CD " & _
    "WHERE r.RECORD_ID = " & Ofm_SqlValue(keyValue), _
    mFields, _
    mOriginalValues
```

Use this when:

- the form needs lookup labels
- the form needs derived display fields
- the read model is richer than the write model

If you use this pattern, prefer pairing it with `reloadSql` during save so the form
comes back with the same joined display values after insert/update.


## 7. Recommended lookup pattern

For combo boxes and list boxes:

1. define lookup metadata in `ConfigureFields`
2. call `Ofm_LoadLookupControls Me, mFields` in `LoadLookups`

Example:

```vb
Set f = Ofm_AddField(mFields, "STATUS_CD", "cboStatus")
f.ControlKind = "COMBOBOX"
f.LookupSql = _
    "SELECT STATUS_CD, STATUS_TEXT " & _
    "FROM APP_STATUS_LU " & _
    "ORDER BY STATUS_TEXT"
f.LookupBoundColumn = 1
f.LookupDisplayColumn = 2
f.LookupIncludeBlankRow = True
f.LookupColumnWidths = "0;1.5"""
```

Then:

```vb
Private Sub LoadLookups()
    Ofm_LoadLookupControls Me, mFields
End Sub
```

Do not use saved Access queries as runtime lookup row sources if you want to stay
inside the stateless Oracle session model.


## 8. Recommended validation pattern

Use two layers of validation:

1. generic validation from the engine
   - `IsRequired`
   - normalization rules

2. form-specific validation in `ValidateForm`
   - length rules
   - cross-field rules
   - business-specific checks

Example:

```vb
Private Sub ValidateForm()

    If Nz(Me.cboStatus.Value, vbNullString) = "CLOSED" Then
        If IsNull(Me.txtClosedDate.Value) Then
            Me.txtClosedDate.SetFocus
            Err.Raise vbObjectError + 9001, Me.Name & ".ValidateForm", _
                      "Closed Date is required when Status is CLOSED."
        End If
    End If

End Sub
```

Keep `ValidateForm` short and readable.

If it grows large, move business rules to a dedicated domain module.


## 9. Recommended save pattern

The save button should usually:

1. run `ValidateForm`
2. call `Ofm_SaveRecord`
3. update `mIsNewRecord`
4. update the key control if needed
5. show a user-facing success message if appropriate

Recommended shape:

```vb
Private Sub btnSave_Click()

    Dim savedKey As Variant

    On Error GoTo ErrHandler

    ValidateForm

    savedKey = Ofm_SaveRecord( _
        Me, _
        Get_DB_Schema(), _
        cTableName, _
        cKeyField, _
        mFields, _
        mOriginalValues, _
        mIsNewRecord, _
        cSequenceName, _
        True)

    Me.RECORD_ID = savedKey
    mIsNewRecord = False
    Exit Sub

ErrHandler:
    MsgBox Err.Description, vbExclamation

End Sub
```

For joined-read forms, include `reloadSql`.


## 10. Recommended delete pattern

Delete behavior should stay simple:

1. skip delete for new records
2. confirm with the user
3. call `Ofm_Delete`
4. close the form

Recommended shape:

```vb
Private Sub btnDelete_Click()

    If mIsNewRecord Then Exit Sub

    If MsgBox("Delete this record?", vbQuestion + vbYesNo) = vbNo Then Exit Sub

    Ofm_Delete Get_DB_Schema(), cTableName, cKeyField, Me.RECORD_ID

    DoCmd.Close acForm, Me.Name

End Sub
```


## 11. Recommended close/cancel pattern

If the form has a close button, prefer this pattern:

```vb
Private Sub btnClose_Click()

    If Ofm_IsDirty(Me, mFields, mOriginalValues) Then
        If MsgBox("Discard unsaved changes?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If

    DoCmd.Close acForm, Me.Name

End Sub
```

That gives a consistent unsaved-changes experience.


## 12. Recommended naming pattern

Try to keep form code predictable:

- `cTableName`
- `cKeyField`
- `cSequenceName`
- `mFields`
- `mOriginalValues`
- `mIsNewRecord`
- `ConfigureFields`
- `LoadLookups`
- `ValidateForm`

For controls:

- text box: `txtRecordName`
- combo box: `cboStatus`
- list box: `lstAvailableItems`
- checkbox: `chkIsActive`
- key control: `RECORD_ID` or `txtRecordId`

Consistency matters more than perfection.


## 13. What not to put in the base form

Avoid putting these directly in every form unless necessary:

- raw SQL-building logic scattered across click events
- repeated combo-loading loops
- Oracle connection handling
- global environment switching
- linked-table refresh logic
- saved passthrough query management

Those either already belong elsewhere, or should become shared helpers.


## 14. When to deviate from the pattern

Deviate only when the form genuinely needs it.

Good reasons to deviate:

- composite keys
- multi-table transactional workflows
- master/detail save behavior
- grid/list editing patterns
- highly specialized validation

Even then, try to preserve the same overall structure:

- configure
- load
- lookups
- validate
- save
- delete


## 15. Recommended first build order

For a new form, build in this order:

1. create the form and controls
2. write the module-level constants and variables
3. write `ConfigureFields`
4. make `Ofm_InitNewForm` work
5. make `Ofm_LoadForm` work
6. add `LoadLookups`
7. add `btnSave_Click`
8. add `btnDelete_Click`
9. add `ValidateForm`
10. add dirty-check / close behavior

That gives you a working form quickly and keeps debugging incremental.


## 16. Golden-path checklist

Before calling a form "done", check:

- `mFields` is initialized in `Form_Open`
- every managed control has a field definition
- exactly one key field is defined for standard forms
- required fields are marked
- checkbox mappings are explicit
- lookups are loaded through `Ofm_LoadLookupControls` or a deliberate equivalent
- new record path works
- existing record load works
- save works
- delete works
- dirty-check works if the form allows cancel/close
- business validation is in `ValidateForm`


## 17. Recommended next step

Once you have one form built successfully with this pattern, use it as the template
for the next forms.

Do not try to invent a new pattern per form.

This architecture will stay maintainable if most forms look nearly identical at the
module level and differ mainly in:

- the field definitions
- lookup SQL
- joined load SQL
- business rules
