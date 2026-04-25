# Form Test Plan

This document proposes a practical first test plan for forms built on:

- `clsOracleFormField`
- `modOracleFormEngine`
- `modOracleDataAccess`
- runtime Oracle session login

The goal is to validate the architecture before many forms are built.

Use generic object names in the examples:

- table: `APP_RECORD`
- key field: `RECORD_ID`
- sequence: `APP_RECORD_SEQ`


## 1. Test strategy

Start with one "golden path" sample form and one safe test table.

Do not try to test every future form at first.

Instead, prove that the pattern works for:

- new record
- existing record
- lookups
- validation
- save
- delete
- dirty tracking

If those work cleanly once, later forms mostly become configuration and business-rule
work.


## 2. Recommended test layers

Use three layers:

1. environment/session tests
2. engine/helper tests
3. form behavior tests


## 3. Environment/session tests

These confirm the runtime Oracle session model is healthy before you blame the form.

After logging in through `frmLogin`, run the built-in smoke helper from the Immediate
Window:

```vb
OracleSmoke_RunRuntime
```

That checks:

- runtime Oracle session exists
- `SELECT USER FROM DUAL` matches the logged-in user
- scalar query execution works
- row materialization works

You can also run this diagnostic when testing joined SQL aliases:

```vb
OracleSmoke_TestDuplicateColumnDiagnostic
```

That confirms duplicate returned column names produce a clear aliasing error.

### Test E1: login succeeds

Verify:

- login form accepts valid credentials
- `SELECT USER FROM DUAL` returns the expected user

### Test E2: re-login as another user

Verify:

- log in as user A
- reopen login
- log in as user B
- runtime queries now return user B

### Test E3: lookup SQL uses current runtime session

Verify:

- after login, `Ofm_LoadLookupControls` loads values without prompting
- values reflect the current Oracle session/user permissions


## 4. Engine/helper tests

These can be done from the Immediate Window or a small test harness form.

### Test G1: field normalization

For a field with:

- `TrimOnSave = True`
- `NullIfBlank = True`

Verify:

- `"  ABC  "` normalizes to `"ABC"`
- `"   "` normalizes to `Null`

### Test G2: checkbox mapping

For a checkbox field with:

- `CheckedValue = "Y"`
- `UncheckedValue = "N"`

Verify:

- `True` saves as `"Y"`
- `False` saves as `"N"`
- `"Y"` loads as checked
- `"N"` loads as unchecked

### Test G3: lookup control loading

Verify:

- `Ofm_LoadLookupControls` populates a combo box
- bound values save correctly
- visible text matches the intended display column

### Test G4: dirty tracking

Verify:

- immediately after load, `Ofm_IsDirty = False`
- change one control, `Ofm_IsDirty = True`
- save record, `Ofm_IsDirty = False`

### Test G5: changed-fields-only update

Verify:

- update one field only
- save succeeds
- no unrelated values are changed in Oracle


## 5. Base form behavior tests

These should be the first real form tests.

### Test F1: new record path

Steps:

1. open the form without `OpenArgs`
2. verify defaults populate correctly
3. verify lookups load

Expected:

- `mIsNewRecord = True`
- key is blank until save, unless you intentionally prefill it
- defaults and blank lookups behave as expected

### Test F2: existing record load path

Steps:

1. open the form with a valid key in `OpenArgs`
2. verify row data loads
3. verify `mOriginalValues` is populated

Expected:

- correct record is displayed
- combo/list values are loaded
- joined display fields load if used

### Test F3: required-field validation

Steps:

1. leave a required field blank
2. click Save

Expected:

- save is blocked
- focus goes to the missing field or a clear error is shown

### Test F4: custom business validation

Steps:

1. enter values that violate `ValidateForm`
2. click Save

Expected:

- save is blocked
- error message clearly explains why

### Test F5: insert

Steps:

1. open new form
2. enter valid data
3. click Save

Expected:

- record is inserted
- key is populated after save
- `mIsNewRecord` becomes `False`
- reload happens correctly if enabled

### Test F6: update

Steps:

1. open an existing record
2. change one or two fields
3. click Save

Expected:

- only intended values change
- record reloads cleanly
- dirty state resets

### Test F7: delete

Steps:

1. open an existing record
2. click Delete
3. confirm delete

Expected:

- row is deleted
- form closes or resets as designed

### Test F8: cancel/close with dirty form

Steps:

1. open an existing record
2. change a value
3. close or cancel

Expected:

- user is prompted before changes are discarded


## 6. Joined-read/base-write tests

If the sample form uses joined display fields, add these tests.

### Test J1: joined display field loads

Verify:

- lookup text from joined SQL appears correctly
- display-only fields are populated

### Test J2: display-only field is not written

Verify:

- save does not attempt to update read-only joined display columns

### Test J3: post-save reload restores joined text

Verify:

- if `reloadSql` is used, joined display text comes back after save


## 7. Lookup-specific tests

### Test L1: blank-row handling

Verify:

- a combo configured with `LookupIncludeBlankRow = True` gets a blank choice
- a required field still fails validation if left blank

### Test L2: code/value pair behavior

Verify:

- the stored combo value is the code column
- the visible text is the description column

### Test L3: zero-row lookup

Verify:

- the control still loads without crashing when the lookup query returns no rows


## 8. Suggested first test table

Use a very simple test table first.

Recommended characteristics:

- single numeric primary key
- Oracle sequence
- one required text field
- one optional text field
- one lookup/code field
- one `Y/N` checkbox-style field

That lets you test nearly the whole form pattern without unnecessary complexity.


## 9. Suggested first sample form

The first sample form should include:

- one key field
- one required text box
- one optional text box
- one combo box backed by lookup SQL
- one checkbox with custom mapping

That gives you:

- normalization coverage
- lookup coverage
- checkbox coverage
- save/update/delete coverage


## 10. Immediate Window smoke tests

Before opening the form, these checks are useful:

```vb
? PTQ_SelectString("SELECT USER FROM DUAL")
? Oracle_GetNextSequenceValue(Get_DB_Schema(), "APP_RECORD_SEQ")
```

If the form uses a lookup:

```vb
Set rows = PTQ_GetRows("SELECT STATUS_CD, STATUS_TEXT FROM APP_STATUS_LU ORDER BY STATUS_TEXT")
? rows.Count
```

If these fail, fix the session/query problem before debugging the form.


## 11. Recommended order for testing

Run tests in this order:

1. login/session tests
2. lookup SQL smoke tests
3. new-form load
4. existing-form load
5. required-field validation
6. insert
7. update
8. delete
9. dirty-check
10. joined-read tests if applicable

That order isolates problems quickly.


## 12. Recommended first pass of manual tests

For the very first sample form, I would manually verify:

- login works
- sample form opens new
- sample form opens existing
- combo loads
- checkbox maps correctly
- insert works
- update works
- delete works
- close-with-unsaved-changes works

Only after those pass would I start cloning the pattern into additional forms.


## 13. Next step

Once you pick the first sample table/form, convert this plan into:

- one manual checklist for you
- optionally one small VBA smoke-test module for repeated sanity checks
