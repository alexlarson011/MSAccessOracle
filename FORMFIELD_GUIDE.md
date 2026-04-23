# Using `clsOracleFormField` and `modOracleFormEngine`

This guide explains how to use the unbound Oracle form engine in this repository to
load, edit, insert, update, and delete Oracle-backed records from Access forms
without binding the form to linked tables.

The examples below use generic names on purpose:

- table: `APP_RECORD`
- key field: `RECORD_ID`
- editable fields: `RECORD_NAME`, `STATUS_CD`, `IS_ACTIVE_YN`
- lookup/display field: `STATUS_TEXT`


## 1. What the form-field class does

`clsOracleFormField` describes one mapping between:

- an Oracle field used for writes: `DbFieldName`
- an Oracle field or alias used for reads: `LoadFieldName`
- an Access control: `ControlName`

Each field definition also describes how that control should behave:

- whether it is the key
- whether it participates in `INSERT`
- whether it participates in `UPDATE`
- whether it is required
- whether blank strings should become `Null`
- whether strings should be trimmed or uppercased
- whether a checkbox should map to custom Oracle values like `Y` / `N`
- what default value should be used for new records

Think of `clsOracleFormField` as metadata, not behavior.

The class does not execute SQL, open recordsets, or know anything about a specific
form. It only describes a field. `modOracleFormEngine` reads those descriptions and
does the actual work.


## 2. When to use this pattern

Use this pattern when:

- your Access form is unbound
- you want explicit SQL instead of Access bound-form behavior
- you want to write back to a base Oracle table
- you may want to load from a view or joined SQL
- you want the form to be stateless and login-session driven

Do not use this pattern when:

- the form is already built around Access bound forms and linked tables
- the record structure is too dynamic to describe with a stable field collection


## 3. The moving parts

Typical form-engine usage has three pieces:

1. A form module that owns:
   - `mFields`
   - `mOriginalValues`
   - `mIsNewRecord`

2. A field-definition routine that builds a `Collection` of `clsOracleFormField`
   objects.

3. Calls into `modOracleFormEngine` to:
   - initialize a new form
   - load an existing row
   - validate and save
   - delete


## 4. Minimal form module skeleton

Here is the basic shape of an Access form module:

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

    Dim recordId As Variant

    recordId = Nz(Me.OpenArgs, vbNullString)

    If Len(CStr(recordId)) = 0 Then
        mIsNewRecord = True
        Ofm_InitNewForm Me, mFields, mOriginalValues
    Else
        mIsNewRecord = False
        Ofm_LoadForm Me, Get_DB_Schema(), cTableName, cKeyField, CLng(recordId), mFields, mOriginalValues
    End If

End Sub

Private Sub btnSave_Click()

    Dim savedKey As Variant

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

    mIsNewRecord = False
    Me.RECORD_ID = savedKey

    MsgBox "Record saved.", vbInformation

End Sub

Private Sub btnDelete_Click()

    If mIsNewRecord Then Exit Sub

    Ofm_Delete Get_DB_Schema(), cTableName, cKeyField, Me.RECORD_ID

    MsgBox "Record deleted.", vbInformation
    DoCmd.Close acForm, Me.Name

End Sub
```


## 5. Building the field collection

This is the core of the pattern. Each control you want the engine to manage should
have a corresponding `clsOracleFormField`.

Example:

```vb
Private Sub ConfigureFields()

    Dim f As clsOracleFormField

    Set f = Ofm_AddField(mFields, "RECORD_ID", "RECORD_ID", True, True, False)
    f.IsRequired = True

    Set f = Ofm_AddField(mFields, "RECORD_NAME", "txtRecordName")
    f.IsRequired = True
    f.TrimOnSave = True
    f.NullIfBlank = True

    Set f = Ofm_AddField(mFields, "STATUS_CD", "cboStatus")
    f.IsRequired = True

    Set f = Ofm_AddField(mFields, "IS_ACTIVE_YN", "chkIsActive")
    f.ControlKind = "CHECKBOX"
    f.UseCustomBooleanMapping = True
    f.CheckedValue = "Y"
    f.UncheckedValue = "N"

End Sub
```

That one routine tells the engine:

- which Oracle columns matter
- which Access controls map to those columns
- which values are required
- which fields can be updated
- how booleans should translate


## 6. Understanding the important properties

### 6.1 `DbFieldName`

This is the Oracle column used in generated `INSERT` and `UPDATE` SQL.

Example:

```vb
Set f = Ofm_AddField(mFields, "RECORD_NAME", "txtRecordName")
```

When the engine builds SQL, it will write:

```sql
RECORD_NAME = ...
```


### 6.2 `LoadFieldName`

This is the field name or alias expected in the row returned during load.

If you do not set it, `LoadFieldName` defaults to `DbFieldName`.

Example:

```vb
Set f = Ofm_AddField(mFields, "STATUS_CD", "txtStatusText", False, False, False)
f.LoadFieldName = "STATUS_TEXT"
```

This means:

- read `STATUS_TEXT` from the query result
- place it into `txtStatusText`
- do not write it back on insert or update

This property is what makes joined read models practical.


### 6.3 `ControlName`

This must match the Access control name on the form.

Examples:

- `"txtRecordName"`
- `"cboStatus"`
- `"chkIsActive"`
- `"RECORD_ID"`

If the control name does not exist on the form, the engine will fail when it tries
to read or write that control.


### 6.4 `IsKey`

Marks the primary key field definition.

The engine uses this for:

- identifying the current key during save
- validating that updates and deletes have a key
- writing a sequence-generated key back into the right control

Every form definition should have exactly one key field.


### 6.5 `IsInsertable`

Controls whether the field appears in generated `INSERT` SQL.

Set this to `False` for:

- read-only display fields
- joined lookup text fields
- fields populated entirely by the database that you do not want to send

Example:

```vb
Set f = Ofm_AddField(mFields, "STATUS_CD", "txtStatusText", False, False, False)
f.LoadFieldName = "STATUS_TEXT"
```


### 6.6 `IsUpdatable`

Controls whether the field can appear in generated `UPDATE` SQL.

Set this to `False` for:

- keys that should never change
- read-only display fields
- fields you want to load but never update from this form


### 6.7 `IsDbGenerated`

This property exists for metadata and clarity, but the current engine does not make
special SQL decisions from it by itself.

In practice, you still control write behavior through:

- `IsInsertable`
- `IsUpdatable`
- `sequenceName` passed to `Ofm_SaveRecord` / `Ofm_Insert`

It is still useful to set `IsDbGenerated = True` for readability.


### 6.8 `IsRequired`

If `True`, the engine checks the field during `Ofm_ValidateRequiredFields`.

A required field is considered missing when the normalized value is:

- `Null`
- an empty string

Because validation uses normalized values, properties like `TrimOnSave` and
`NullIfBlank` affect required-field behavior.

Example:

```vb
Set f = Ofm_AddField(mFields, "RECORD_NAME", "txtRecordName")
f.IsRequired = True
f.TrimOnSave = True
f.NullIfBlank = True
```

With that setup, a value like `"   "` becomes `Null`, and the field fails required
validation.


### 6.9 `TrimOnSave`

If `True`, string values are trimmed before insert, update, validation comparison,
and snapshot comparison.

Use this for fields where leading/trailing spaces should never matter.


### 6.10 `UppercaseOnSave`

If `True`, string values are converted to uppercase before save and comparison.

This is useful for:

- codes
- abbreviations
- user IDs
- flag fields stored as uppercase text

Example:

```vb
Set f = Ofm_AddField(mFields, "STATUS_CD", "cboStatus")
f.UppercaseOnSave = True
```


### 6.11 `NullIfBlank`

If `True`, a blank string becomes `Null`.

This is useful for optional text values where Oracle should store `NULL`, not an
empty string.

Example:

```vb
Set f = Ofm_AddField(mFields, "NOTES_TXT", "txtNotes")
f.TrimOnSave = True
f.NullIfBlank = True
```


### 6.12 `DefaultValue`

`DefaultValue` is used in two places in the current implementation:

- during `Ofm_InitNewForm`
- as a normalization fallback when the incoming value is `Null`

That second behavior is important.

Example:

```vb
Set f = Ofm_AddField(mFields, "IS_ACTIVE_YN", "chkIsActive")
f.ControlKind = "CHECKBOX"
f.UseCustomBooleanMapping = True
f.CheckedValue = "Y"
f.UncheckedValue = "N"
f.DefaultValue = "Y"
```

That means:

- a new record starts checked
- a `Null` value can normalize to `"Y"` if passed through `GetNormalizedValue`

Use `DefaultValue` deliberately, especially on editable fields.


### 6.13 `ControlKind`

The current engine cares most about:

- `"TEXT"`
- `"CHECKBOX"`
- `"OPTIONGROUP"`

In practice, checkbox behavior is the one with special translation logic.

If you do nothing, the class defaults to `"TEXT"`.


### 6.14 `UseCustomBooleanMapping`, `CheckedValue`, `UncheckedValue`

These control how checkbox values are translated between Access and Oracle.

Example for `Y` / `N` storage:

```vb
Set f = Ofm_AddField(mFields, "IS_ACTIVE_YN", "chkIsActive")
f.ControlKind = "CHECKBOX"
f.UseCustomBooleanMapping = True
f.CheckedValue = "Y"
f.UncheckedValue = "N"
```

Behavior:

- Oracle `"Y"` loads into the checkbox as `True`
- Oracle `"N"` loads into the checkbox as `False`
- checked in Access saves as `"Y"`
- unchecked in Access saves as `"N"`

Example for `1` / `0`:

```vb
f.CheckedValue = 1
f.UncheckedValue = 0
```

Example for `Y` / `Null`:

```vb
f.CheckedValue = "Y"
f.UncheckedValue = Null
```


## 7. Loading an existing base-table record

The simplest load path is `Ofm_LoadForm`.

Example:

```vb
Ofm_LoadForm _
    Me, _
    Get_DB_Schema(), _
    "APP_RECORD", _
    "RECORD_ID", _
    1001, _
    mFields, _
    mOriginalValues
```

What happens:

1. The engine builds a `SELECT` list from `LoadFieldName`.
2. It queries:

```sql
SELECT RECORD_ID, RECORD_NAME, STATUS_CD, IS_ACTIVE_YN
FROM YOUR_SCHEMA.APP_RECORD
WHERE RECORD_ID = 1001
```

3. It writes each returned value into the matching Access control.
4. It snapshots the normalized values into `mOriginalValues`.

That snapshot is later used for dirty checking and changed-fields-only updates.


## 8. Loading from joined or custom SQL

Use `Ofm_LoadFormBySql` when the read model is richer than the write model.

Example field configuration:

```vb
Private Sub ConfigureFields()

    Dim f As clsOracleFormField

    Set f = Ofm_AddField(mFields, "RECORD_ID", "RECORD_ID", True, True, False)

    Set f = Ofm_AddField(mFields, "RECORD_NAME", "txtRecordName")
    f.IsRequired = True
    f.TrimOnSave = True

    Set f = Ofm_AddField(mFields, "STATUS_CD", "cboStatus")
    f.IsRequired = True

    Set f = Ofm_AddField(mFields, "STATUS_CD", "txtStatusText", False, False, False)
    f.LoadFieldName = "STATUS_TEXT"

End Sub
```

Example load:

```vb
Ofm_LoadFormBySql _
    Me, _
    "SELECT r.RECORD_ID, " & _
    "       r.RECORD_NAME, " & _
    "       r.STATUS_CD, " & _
    "       s.STATUS_TEXT " & _
    "FROM APP_RECORD r " & _
    "LEFT JOIN APP_STATUS_LU s ON s.STATUS_CD = r.STATUS_CD " & _
    "WHERE r.RECORD_ID = 1001", _
    mFields, _
    mOriginalValues
```

Why this works:

- editable controls still map through `DbFieldName`
- display-only text can come from `LoadFieldName`
- save operations still target the base table and base columns

This is one of the strongest uses of the pattern.


## 9. New-record initialization

Use `Ofm_InitNewForm` when you want to start from defaults instead of loading from
Oracle.

Example:

```vb
mIsNewRecord = True
Ofm_InitNewForm Me, mFields, mOriginalValues
```

Current behavior:

- if a field has `DefaultValue`, that value is assigned to the control
- otherwise, a checkbox with custom boolean mapping starts as `False`
- otherwise, the control starts as `Null`

Then the engine snapshots those starting values into `mOriginalValues`.


## 10. Saving a record

The usual entry point is `Ofm_SaveRecord`.

Example:

```vb
savedKey = Ofm_SaveRecord( _
    Me, _
    Get_DB_Schema(), _
    "APP_RECORD", _
    "RECORD_ID", _
    mFields, _
    mOriginalValues, _
    mIsNewRecord, _
    "APP_RECORD_SEQ", _
    True)
```

How it behaves:

- if `mIsNewRecord = True`, it runs `Ofm_Insert`
- otherwise, it runs `Ofm_Update`

Before save, it:

- normalizes values
- validates required fields

After save, it:

- reloads the record by default
- refreshes `mOriginalValues`


## 11. What insert actually does

During `Ofm_Insert`, the engine:

1. validates required fields
2. finds the key field
3. if `sequenceName` is supplied, gets `NEXTVAL`
4. writes that generated key into the key control
5. builds `INSERT` SQL from all `IsInsertable = True` fields
6. executes the SQL
7. reloads the record or snapshots current values

Example generated SQL shape:

```sql
INSERT INTO YOUR_SCHEMA.APP_RECORD
    (RECORD_ID, RECORD_NAME, STATUS_CD, IS_ACTIVE_YN)
VALUES
    (1005, 'Example Name', 'OPEN', 'Y')
```


## 12. What update actually does

During `Ofm_Update`, the engine:

1. validates required fields
2. compares current normalized values to `mOriginalValues`
3. includes only changed fields in the `SET` clause
4. excludes the key field from the update
5. executes the SQL
6. reloads the record or snapshots current values

Example generated SQL shape:

```sql
UPDATE YOUR_SCHEMA.APP_RECORD
SET RECORD_NAME = 'Updated Name',
    STATUS_CD = 'CLOSED'
WHERE RECORD_ID = 1005
```

If no updatable field changed, the engine exits without running an update.


## 13. Deleting a record

Use `Ofm_Delete`:

```vb
Ofm_Delete Get_DB_Schema(), "APP_RECORD", "RECORD_ID", Me.RECORD_ID
```

Generated SQL shape:

```sql
DELETE FROM YOUR_SCHEMA.APP_RECORD
WHERE RECORD_ID = 1005
```


## 14. Dirty checking and changed-field inspection

The engine gives you helpers for UI logic.

Check whether anything changed:

```vb
If Ofm_IsDirty(Me, mFields, mOriginalValues) Then
    MsgBox "You have unsaved changes."
End If
```

Get the changed field definitions:

```vb
Dim changedFields As Collection
Dim f As clsOracleFormField

Set changedFields = Ofm_GetChangedFields(Me, mFields, mOriginalValues)

For Each f In changedFields
    Debug.Print f.ControlName
Next f
```

This is useful for:

- save prompts
- debug tracing
- custom auditing


## 15. Generic end-to-end example

Below is a more complete pattern for a maintainable form.

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

    Dim recordId As Variant

    recordId = Nz(Me.OpenArgs, vbNullString)

    If Len(CStr(recordId)) = 0 Then
        mIsNewRecord = True
        Ofm_InitNewForm Me, mFields, mOriginalValues
        LoadStatusChoices
    Else
        mIsNewRecord = False
        Ofm_LoadFormBySql _
            Me, _
            "SELECT r.RECORD_ID, " & _
            "       r.RECORD_NAME, " & _
            "       r.STATUS_CD, " & _
            "       r.IS_ACTIVE_YN, " & _
            "       s.STATUS_TEXT " & _
            "FROM APP_RECORD r " & _
            "LEFT JOIN APP_STATUS_LU s ON s.STATUS_CD = r.STATUS_CD " & _
            "WHERE r.RECORD_ID = " & CLng(recordId), _
            mFields, _
            mOriginalValues
        LoadStatusChoices
    End If

End Sub

Private Sub ConfigureFields()

    Dim f As clsOracleFormField

    Set f = Ofm_AddField(mFields, "RECORD_ID", "RECORD_ID", True, True, False)
    f.IsRequired = True
    f.IsDbGenerated = True

    Set f = Ofm_AddField(mFields, "RECORD_NAME", "txtRecordName")
    f.IsRequired = True
    f.TrimOnSave = True
    f.NullIfBlank = True

    Set f = Ofm_AddField(mFields, "STATUS_CD", "cboStatus")
    f.IsRequired = True
    f.TrimOnSave = True
    f.UppercaseOnSave = True

    Set f = Ofm_AddField(mFields, "IS_ACTIVE_YN", "chkIsActive")
    f.ControlKind = "CHECKBOX"
    f.UseCustomBooleanMapping = True
    f.CheckedValue = "Y"
    f.UncheckedValue = "N"
    f.DefaultValue = "Y"

    Set f = Ofm_AddField(mFields, "STATUS_CD", "txtStatusText", False, False, False)
    f.LoadFieldName = "STATUS_TEXT"

End Sub

Private Sub btnSave_Click()

    Dim savedKey As Variant

    savedKey = Ofm_SaveRecord( _
        Me, _
        Get_DB_Schema(), _
        cTableName, _
        cKeyField, _
        mFields, _
        mOriginalValues, _
        mIsNewRecord, _
        cSequenceName, _
        True, _
        "SELECT r.RECORD_ID, " & _
        "       r.RECORD_NAME, " & _
        "       r.STATUS_CD, " & _
        "       r.IS_ACTIVE_YN, " & _
        "       s.STATUS_TEXT " & _
        "FROM APP_RECORD r " & _
        "LEFT JOIN APP_STATUS_LU s ON s.STATUS_CD = r.STATUS_CD " & _
        "WHERE r.RECORD_ID = [OFM_KEY_VALUE]")

    Me.RECORD_ID = savedKey
    mIsNewRecord = False

    MsgBox "Record saved.", vbInformation

End Sub

Private Sub btnDelete_Click()

    If mIsNewRecord Then Exit Sub

    If MsgBox("Delete this record?", vbQuestion + vbYesNo) = vbNo Then Exit Sub

    Ofm_Delete Get_DB_Schema(), cTableName, cKeyField, Me.RECORD_ID

    DoCmd.Close acForm, Me.Name

End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    Cancel = True
End Sub

Private Sub LoadStatusChoices()
    ' Load combo values using your preferred lookup-loading helper.
End Sub
```


## 16. Joined read model, base-table write model

This is a very important pattern in this repository.

You can:

- load from joined SQL
- show lookup text on the form
- still write only to the base table

The rules are:

1. For writable fields, keep `DbFieldName` pointed at the real base-table column.
2. For read-only joined/display fields, use `LoadFieldName` for the returned alias.
3. Mark display-only fields as:
   - `IsInsertable = False`
   - `IsUpdatable = False`
4. If you want the richer joined data to come back after save, pass `reloadSql` to
   `Ofm_SaveRecord`.

Example display-only field:

```vb
Set f = Ofm_AddField(mFields, "STATUS_CD", "txtStatusText", False, False, False)
f.LoadFieldName = "STATUS_TEXT"
```


## 17. Optional DSN override

The current engine places `dsn` last and makes it optional.

Most forms should just omit it:

```vb
Ofm_LoadForm Me, Get_DB_Schema(), cTableName, cKeyField, keyValue, mFields, mOriginalValues
```

That defaults to `Get_DB_DSN()`.

If you want to force a specific DSN:

```vb
Ofm_LoadForm Me, Get_DB_Schema(), cTableName, cKeyField, keyValue, mFields, mOriginalValues, "MY_ALT_DSN"
```

That same pattern applies to:

- `Ofm_LoadForm`
- `Ofm_LoadFormBySql`
- `Ofm_Insert`
- `Ofm_Update`
- `Ofm_Delete`
- `Ofm_SaveRecord`


## 18. Common design patterns

### Pattern A: Simple maintenance form

Use when the form edits one Oracle table directly.

Typical setup:

- `Ofm_LoadForm`
- `Ofm_SaveRecord`
- no custom `LoadFieldName`
- no custom `reloadSql`


### Pattern B: Joined read, base-table write

Use when lookup text or descriptive labels are needed.

Typical setup:

- `Ofm_LoadFormBySql`
- one or more display-only fields using `LoadFieldName`
- `Ofm_SaveRecord` with `reloadSql`


### Pattern C: Sequence-generated insert

Use when Oracle key values come from a sequence.

Typical setup:

- mark the key field with `IsKey = True`
- pass `sequenceName` to `Ofm_SaveRecord`
- keep the key field not updatable


### Pattern D: Checkbox-to-Oracle flag mapping

Use when the database stores flags as `Y/N`, `1/0`, or similar.

Typical setup:

- `ControlKind = "CHECKBOX"`
- `UseCustomBooleanMapping = True`
- set `CheckedValue` and `UncheckedValue`


## 19. Common mistakes to avoid

### Mistake 1: forgetting to configure fields before load/save

If `mFields` is empty or not configured before load or save, the engine cannot work.


### Mistake 2: missing key field

Every save/update/delete form should define one field with `IsKey = True`.


### Mistake 3: using a load alias without `LoadFieldName`

If your SQL returns:

```sql
SELECT STATUS_TEXT
```

but your field definition expects `STATUS_CD`, the engine will raise an error
because the row does not contain the expected load field.


### Mistake 4: making display-only fields insertable or updatable

Joined display fields should usually be:

- `IsInsertable = False`
- `IsUpdatable = False`


### Mistake 5: forgetting that normalization affects dirty checking

The engine compares normalized values, not raw control text.

That means:

- trimming can make `"ABC "` equal `"ABC"`
- uppercasing can make `"abc"` equal `"ABC"`
- `NullIfBlank` can make `""` equal `Null`

This is usually good, but you should know it is happening.


### Mistake 6: relying on `IsDbGenerated` alone

`IsDbGenerated` is descriptive. It does not, by itself, remove a field from insert
or update SQL.

Actual write participation still comes from:

- `IsInsertable`
- `IsUpdatable`
- the presence of a `sequenceName`


## 20. Debugging tips

### Inspect one field definition

```vb
Dim f As clsOracleFormField

Set f = Ofm_GetFieldByControlName(mFields, "txtRecordName")
Debug.Print f.DebugSummary
```


### Inspect all configured fields

```vb
Debug.Print Ofm_DebugFieldSummary(mFields)
```


### See which fields changed

```vb
Dim changedFields As Collection
Dim f As clsOracleFormField

Set changedFields = Ofm_GetChangedFields(Me, mFields, mOriginalValues)

For Each f In changedFields
    Debug.Print f.ControlName & " -> " & f.DbFieldName
Next f
```


### Confirm field validity while building config

```vb
Dim f As clsOracleFormField

For Each f In mFields
    Debug.Print f.ControlName, f.IsValid
Next f
```


## 21. Quick reference

Use this when you need a memory jog.

### Create one field

```vb
Set f = Ofm_AddField(mFields, "COLUMN_NAME", "txtControl")
```

### Mark it required

```vb
f.IsRequired = True
```

### Trim and null blank text

```vb
f.TrimOnSave = True
f.NullIfBlank = True
```

### Map a checkbox to `Y` / `N`

```vb
f.ControlKind = "CHECKBOX"
f.UseCustomBooleanMapping = True
f.CheckedValue = "Y"
f.UncheckedValue = "N"
```

### Load from a base table

```vb
Ofm_LoadForm Me, Get_DB_Schema(), "APP_RECORD", "RECORD_ID", 1001, mFields, mOriginalValues
```

### Load from custom SQL

```vb
Ofm_LoadFormBySql Me, "SELECT ...", mFields, mOriginalValues
```

### Save

```vb
savedKey = Ofm_SaveRecord(Me, Get_DB_Schema(), "APP_RECORD", "RECORD_ID", mFields, mOriginalValues, mIsNewRecord, "APP_RECORD_SEQ", True)
```

### Delete

```vb
Ofm_Delete Get_DB_Schema(), "APP_RECORD", "RECORD_ID", Me.RECORD_ID
```


## 22. Recommended starting pattern

If you are building a new form, start simple:

1. Create the form unbound.
2. Add controls with stable names.
3. Build `mFields` in one `ConfigureFields` routine.
4. Use `Ofm_InitNewForm` for new rows.
5. Use `Ofm_LoadForm` for simple forms.
6. Move to `Ofm_LoadFormBySql` only when you need joined reads.
7. Use `Ofm_SaveRecord` as the default save entry point.
8. Add `reloadSql` only when you want richer display values after save.

That path keeps the form easy to reason about and makes it clear where the Oracle
read model and write model differ.
