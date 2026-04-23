VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOracleFormField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'====================================================================================
' clsOracleFormField
'====================================================================================
'
' README
' ------
' Purpose
' -------
' Represents one field mapping between an Oracle database column and an Access form
' control for use with the unbound Oracle form engine.
'
' Each instance of this class describes how one field should behave during:
'
'     - form load
'     - new-record initialization
'     - validation
'     - normalization
'     - insert
'     - update
'     - checkbox / control value translation
'
'
' Responsibilities
' ----------------
' Each clsOracleFormField instance can describe:
'
'     - database field name
'     - form control name
'     - whether the field is the key
'     - whether the field is insertable
'     - whether the field is updatable
'     - whether the field is database-generated
'     - whether the field is required
'     - whether trimming / uppercasing / null-if-blank normalization should occur
'     - default value behavior
'     - control type
'     - checkbox mapping values
'
'
' Key properties
' --------------
' mapping:
'     DbFieldName
'     ControlName
'
' CRUD behavior:
'     IsKey
'     IsInsertable
'     IsUpdatable
'     IsDbGenerated
'
' validation / normalization:
'     IsRequired
'     TrimOnSave
'     UppercaseOnSave
'     NullIfBlank
'     DefaultValue
'
' metadata:
'     OracleDataType
'     Description
'
' UI / control behavior:
'     ControlKind
'     CheckedValue
'     UncheckedValue
'     UseCustomBooleanMapping
'
'
' Helper methods
' --------------
' validation / structure:
'     HasDbFieldName
'     HasControlName
'     IsValid
'
' control type checks:
'     IsCheckbox
'     IsOptionGroup
'     IsTextControl
'
' normalization:
'     GetNormalizedValue
'     IsMissingRequiredValue
'
' debugging:
'     DebugSummary
'
'
' Typical usage
' -------------
' In a form config routine:
'
'     Dim f As clsOracleFormField
'
'     Set f = Ofm_AddField(mFields, "PROJECT_ID", "PROJECT_ID")
'     f.IsRequired = True
'     f.TrimOnSave = True
'
' Checkbox example:
'
'     Set f = Ofm_AddField(mFields, "DO_NOT_USE_INDC", "DO_NOT_USE_INDC")
'     f.ControlKind = "CHECKBOX"
'     f.UseCustomBooleanMapping = True
'     f.CheckedValue = "Y"
'     f.UncheckedValue = "N"
'
' Key field example:
'
'     Set f = Ofm_AddField(mFields, "PROJ_OPTN_ID", "PROJ_OPTN_ID", True, True, False)
'
'
' DefaultValue note
' -----------------
' DefaultValue is currently available to support both:
'
'     - new-record initialization
'     - normalization fallback
'
' If future behavior needs to be stricter, DefaultValue handling can be narrowed so
' it only applies during new-record initialization.
'
'
' Dependencies
' ------------
' No module dependencies.
'
'
' Common callers
' --------------
' Intended callers are:
'
'     - modOracleFormEngine
'     - unbound form modules that define field collections
'
'
' Design notes
' ------------
' This class should remain lightweight and declarative.
'
' It should not:
'
'     - execute SQL
'     - read recordsets
'     - contain form-specific business rules
'
' Its job is only to describe one field and how that field should behave.
'
'
' Version
' -------
' v1
'
'====================================================================================

'------------------------------------------------------------------------------------
' Private backing fields
'------------------------------------------------------------------------------------

Private pDbFieldName As String
Private pControlName As String

Private pIsKey As Boolean
Private pIsInsertable As Boolean
Private pIsUpdatable As Boolean
Private pIsDbGenerated As Boolean

Private pIsRequired As Boolean
Private pTrimOnSave As Boolean
Private pUppercaseOnSave As Boolean
Private pNullIfBlank As Boolean

Private pDefaultValue As Variant
Private pOracleDataType As String
Private pDescription As String

Private pControlKind As String
Private pCheckedValue As Variant
Private pUncheckedValue As Variant
Private pUseCustomBooleanMapping As Boolean

'------------------------------------------------------------------------------------
' Initialization
'------------------------------------------------------------------------------------

Private Sub Class_Initialize()

    pDbFieldName = vbNullString
    pControlName = vbNullString

    pIsKey = False
    pIsInsertable = True
    pIsUpdatable = True
    pIsDbGenerated = False

    pIsRequired = False
    pTrimOnSave = False
    pUppercaseOnSave = False
    pNullIfBlank = False

    pDefaultValue = Null
    pOracleDataType = vbNullString
    pDescription = vbNullString

    pControlKind = "TEXT"
    pCheckedValue = 1
    pUncheckedValue = 0
    pUseCustomBooleanMapping = False

End Sub

'------------------------------------------------------------------------------------
' Core mapping properties
'------------------------------------------------------------------------------------

Public Property Get DbFieldName() As String
    DbFieldName = pDbFieldName
End Property

Public Property Let DbFieldName(ByVal Value As String)
    pDbFieldName = Trim$(Nz(Value, vbNullString))
End Property


Public Property Get ControlName() As String
    ControlName = pControlName
End Property

Public Property Let ControlName(ByVal Value As String)
    pControlName = Trim$(Nz(Value, vbNullString))
End Property

'------------------------------------------------------------------------------------
' CRUD behavior properties
'------------------------------------------------------------------------------------

Public Property Get IsKey() As Boolean
    IsKey = pIsKey
End Property

Public Property Let IsKey(ByVal Value As Boolean)
    pIsKey = Value
End Property


Public Property Get IsInsertable() As Boolean
    IsInsertable = pIsInsertable
End Property

Public Property Let IsInsertable(ByVal Value As Boolean)
    pIsInsertable = Value
End Property


Public Property Get IsUpdatable() As Boolean
    IsUpdatable = pIsUpdatable
End Property

Public Property Let IsUpdatable(ByVal Value As Boolean)
    pIsUpdatable = Value
End Property


Public Property Get IsDbGenerated() As Boolean
    IsDbGenerated = pIsDbGenerated
End Property

Public Property Let IsDbGenerated(ByVal Value As Boolean)
    pIsDbGenerated = Value
End Property

'------------------------------------------------------------------------------------
' Validation / normalization properties
'------------------------------------------------------------------------------------

Public Property Get IsRequired() As Boolean
    IsRequired = pIsRequired
End Property

Public Property Let IsRequired(ByVal Value As Boolean)
    pIsRequired = Value
End Property


Public Property Get TrimOnSave() As Boolean
    TrimOnSave = pTrimOnSave
End Property

Public Property Let TrimOnSave(ByVal Value As Boolean)
    pTrimOnSave = Value
End Property


Public Property Get UppercaseOnSave() As Boolean
    UppercaseOnSave = pUppercaseOnSave
End Property

Public Property Let UppercaseOnSave(ByVal Value As Boolean)
    pUppercaseOnSave = Value
End Property


Public Property Get NullIfBlank() As Boolean
    NullIfBlank = pNullIfBlank
End Property

Public Property Let NullIfBlank(ByVal Value As Boolean)
    pNullIfBlank = Value
End Property

'------------------------------------------------------------------------------------
' Metadata properties
'------------------------------------------------------------------------------------

Public Property Get DefaultValue() As Variant
    DefaultValue = pDefaultValue
End Property

Public Property Let DefaultValue(ByVal Value As Variant)
    pDefaultValue = Value
End Property

Public Property Set DefaultValue(ByVal Value As Variant)
    pDefaultValue = Value
End Property


Public Property Get OracleDataType() As String
    OracleDataType = pOracleDataType
End Property

Public Property Let OracleDataType(ByVal Value As String)
    pOracleDataType = Trim$(Nz(Value, vbNullString))
End Property


Public Property Get Description() As String
    Description = pDescription
End Property

Public Property Let Description(ByVal Value As String)
    pDescription = Trim$(Nz(Value, vbNullString))
End Property

'------------------------------------------------------------------------------------
' Control mapping / UI behavior properties
'------------------------------------------------------------------------------------

Public Property Get ControlKind() As String
    ControlKind = pControlKind
End Property

Public Property Let ControlKind(ByVal Value As String)
    pControlKind = UCase$(Trim$(Nz(Value, "TEXT")))
End Property


Public Property Get CheckedValue() As Variant
    CheckedValue = pCheckedValue
End Property

Public Property Let CheckedValue(ByVal Value As Variant)
    pCheckedValue = Value
End Property

Public Property Set CheckedValue(ByVal Value As Variant)
    pCheckedValue = Value
End Property


Public Property Get UncheckedValue() As Variant
    UncheckedValue = pUncheckedValue
End Property

Public Property Let UncheckedValue(ByVal Value As Variant)
    pUncheckedValue = Value
End Property

Public Property Set UncheckedValue(ByVal Value As Variant)
    pUncheckedValue = Value
End Property


Public Property Get UseCustomBooleanMapping() As Boolean
    UseCustomBooleanMapping = pUseCustomBooleanMapping
End Property

Public Property Let UseCustomBooleanMapping(ByVal Value As Boolean)
    pUseCustomBooleanMapping = Value
End Property

'------------------------------------------------------------------------------------
' Simple validation helpers
'------------------------------------------------------------------------------------

Public Function HasDbFieldName() As Boolean
    HasDbFieldName = (Len(pDbFieldName) > 0)
End Function


Public Function HasControlName() As Boolean
    HasControlName = (Len(pControlName) > 0)
End Function


Public Function IsValid() As Boolean

    IsValid = False

    If Len(pDbFieldName) = 0 Then Exit Function
    If Len(pControlName) = 0 Then Exit Function

    IsValid = True

End Function

'------------------------------------------------------------------------------------
' Control type helpers
'------------------------------------------------------------------------------------

Public Function IsCheckbox() As Boolean
    IsCheckbox = (UCase$(pControlKind) = "CHECKBOX")
End Function


Public Function IsOptionGroup() As Boolean
    IsOptionGroup = (UCase$(pControlKind) = "OPTIONGROUP")
End Function


Public Function IsTextControl() As Boolean
    IsTextControl = (UCase$(pControlKind) = "TEXT" Or Len(pControlKind) = 0)
End Function

'------------------------------------------------------------------------------------
' Value normalization helpers
'------------------------------------------------------------------------------------

Public Function GetNormalizedValue(ByVal InputValue As Variant) As Variant

    Dim v As Variant

    v = InputValue

    If IsNull(v) Then
        If Not IsNull(pDefaultValue) Then
            v = pDefaultValue
        Else
            GetNormalizedValue = Null
            Exit Function
        End If
    End If

    If VarType(v) = vbString Then

        If pTrimOnSave Then
            v = Trim$(CStr(v))
        End If

        If pUppercaseOnSave Then
            v = UCase$(CStr(v))
        End If

        If pNullIfBlank Then
            If Len(CStr(v)) = 0 Then
                GetNormalizedValue = Null
                Exit Function
            End If
        End If

    End If

    GetNormalizedValue = v

End Function


Public Function IsMissingRequiredValue(ByVal InputValue As Variant) As Boolean

    Dim v As Variant

    If Not pIsRequired Then Exit Function

    v = GetNormalizedValue(InputValue)

    If IsNull(v) Then
        IsMissingRequiredValue = True
        Exit Function
    End If

    If VarType(v) = vbString Then
        If Len(CStr(v)) = 0 Then
            IsMissingRequiredValue = True
            Exit Function
        End If
    End If

End Function

'------------------------------------------------------------------------------------
' Debug helpers
'------------------------------------------------------------------------------------

Public Function DebugSummary() As String

    DebugSummary = _
        "DbFieldName=" & pDbFieldName & _
        "; ControlName=" & pControlName & _
        "; IsKey=" & CStr(pIsKey) & _
        "; IsInsertable=" & CStr(pIsInsertable) & _
        "; IsUpdatable=" & CStr(pIsUpdatable) & _
        "; IsDbGenerated=" & CStr(pIsDbGenerated) & _
        "; IsRequired=" & CStr(pIsRequired) & _
        "; TrimOnSave=" & CStr(pTrimOnSave) & _
        "; UppercaseOnSave=" & CStr(pUppercaseOnSave) & _
        "; NullIfBlank=" & CStr(pNullIfBlank) & _
        "; ControlKind=" & pControlKind & _
        "; UseCustomBooleanMapping=" & CStr(pUseCustomBooleanMapping) & _
        "; OracleDataType=" & pOracleDataType

End Function
