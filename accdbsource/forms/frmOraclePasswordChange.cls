Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' See "frmOraclePasswordChange.cls"

'====================================================================================
' frmOraclePasswordChange
'====================================================================================
'
' README
' ------
' Purpose
' -------
' Provides a small Oracle password-change form for the Access application.
'
' This form supports a user-driven password update workflow using the current Oracle
' credentials and a new target password, then optionally pushes the new password back
' into frmLogin if that form is open.
'
'
' Responsibilities
' ----------------
' This form is responsible for:
'
'     - collecting the Oracle user name
'     - collecting the current Oracle password
'     - collecting and confirming a new Oracle password
'     - validating the entered values
'     - calling OracleAdmin_ChangeCurrentUserPassword
'     - updating frmLogin controls when launched from the login flow
'     - optionally refreshing the runtime Oracle session credentials
'
'
' Dependencies
' ------------
' Depends on:
'
'     - modOracleAdmin
'     - modOracleDataAccess
'     - modOracleSession
'
'
' Version
' -------
' v1
'
'====================================================================================

Option Compare Database
Option Explicit

Private Const cFormTitle As String = "Change Oracle Password"

Private Sub Form_Open(Cancel As Integer)

    Dim sUserName As String

    On Error GoTo ErrHandler

    Me.txtAppName.ControlSource = "=""" & cFormTitle & """"
    Me.txtHeader.ControlSource = "=""" & _
        "Enter your current Oracle password and choose a new password." & """"
    Me.lblDSNName.Caption = "DSN: " & Get_DB_DSN()

    sUserName = GetOpenArgValue("USER")

    If Len(sUserName) = 0 Then
        If CurrentProject.AllForms("frmLogin").IsLoaded Then
            sUserName = Trim$(Nz(Forms("frmLogin").Controls("txtUser").Value, vbNullString))
        End If
    End If

    If Len(sUserName) > 0 Then
        Me.txtUserName.Value = UCase$(sUserName)
        Me.txtOldPassword.SetFocus
    Else
        Me.txtUserName.SetFocus
    End If

    Exit Sub

ErrHandler:
    MsgBox "Unable to open the Oracle password change form." & vbCrLf & vbCrLf & _
           Err.Description, vbCritical
    Cancel = True

End Sub

Private Sub btnUpdatePassword_Click()

    Dim sUserName As String
    Dim sOldPassword As String
    Dim sNewPassword As String
    Dim sConfirmPassword As String
    Dim sRuntimeConn As String

    On Error GoTo ErrHandler

    sUserName = UCase$(Trim$(Nz(Me.txtUserName.Value, vbNullString)))
    sOldPassword = Nz(Me.txtOldPassword.Value, vbNullString)
    sNewPassword = Nz(Me.txtNewPassword.Value, vbNullString)
    sConfirmPassword = Nz(Me.txtConfirmPassword.Value, vbNullString)

    If Len(sUserName) = 0 Then
        MsgBox "Enter your Oracle user name.", vbExclamation
        Me.txtUserName.SetFocus
        Exit Sub
    End If

    If Len(sOldPassword) = 0 Then
        MsgBox "Enter your current Oracle password.", vbExclamation
        Me.txtOldPassword.SetFocus
        Exit Sub
    End If

    If Len(sNewPassword) = 0 Then
        MsgBox "Enter your new Oracle password.", vbExclamation
        Me.txtNewPassword.SetFocus
        Exit Sub
    End If

    If Len(sConfirmPassword) = 0 Then
        MsgBox "Confirm your new Oracle password.", vbExclamation
        Me.txtConfirmPassword.SetFocus
        Exit Sub
    End If

    If StrComp(sNewPassword, sConfirmPassword, vbBinaryCompare) <> 0 Then
        MsgBox "The new password and confirmation do not match.", vbExclamation
        Me.txtConfirmPassword.SetFocus
        Exit Sub
    End If

    sRuntimeConn = OracleAdmin_ChangeCurrentUserPassword(Get_DB_DSN(), sUserName, sOldPassword, sNewPassword)

    If CurrentProject.AllForms("frmLogin").IsLoaded Then
        Forms("frmLogin").Controls("txtUser").Value = sUserName
        Forms("frmLogin").Controls("txtPass").Value = sNewPassword
        MsgBox "Oracle password updated. The login form has been updated with your new password. Click Login to continue.", vbInformation
    Else
        g_OracleConnectionString = sRuntimeConn
        g_OracleSessionUser = sUserName
        MsgBox "Oracle password updated.", vbInformation
    End If

    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub

ErrHandler:
    If InStr(1, Err.Description, "ORA-28001", vbTextCompare) > 0 Then
        MsgBox "Oracle reported an expired password and did not allow the custom password change flow to complete." & vbCrLf & vbCrLf & _
               "The Oracle ODBC driver may still require its own password-change prompt for expired passwords.", vbExclamation
    Else
        MsgBox "Unable to update the Oracle password." & vbCrLf & vbCrLf & Err.Description, vbCritical
    End If

End Sub

Private Sub btnCancel_Click()
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Function GetOpenArgValue(ByVal keyName As String) As String

    Dim parts As Variant
    Dim part As Variant
    Dim eqPos As Long
    Dim sKey As String
    Dim sValue As String

    If Len(Nz(Me.OpenArgs, vbNullString)) = 0 Then Exit Function

    parts = Split(Me.OpenArgs, ";")

    For Each part In parts
        eqPos = InStr(1, CStr(part), "=")
        If eqPos > 0 Then
            sKey = UCase$(Trim$(Left$(CStr(part), eqPos - 1)))
            sValue = Mid$(CStr(part), eqPos + 1)

            If sKey = UCase$(Trim$(keyName)) Then
                GetOpenArgValue = sValue
                Exit Function
            End If
        End If
    Next part

End Function
