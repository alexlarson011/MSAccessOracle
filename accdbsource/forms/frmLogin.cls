Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'====================================================================================
' frmLogin
'====================================================================================
'
' README
' ------
' Purpose
' -------
' Provides the Oracle login form for the Access application.
'
' This form is responsible for collecting user credentials, validating them against
' the configured Oracle ODBC DSN, establishing a runtime Oracle session for the
' current Access application session, and then opening the main application form.
'
'
' Responsibilities
' ----------------
' This form is responsible for:
'
'     - displaying application name / version / environment information
'     - prompting the user for Oracle credentials
'     - validating credentials using ADODB
'     - classifying common Oracle / ODBC login failures for the user
'     - displaying login status and failure reasons in the form UI
'     - storing the validated runtime Oracle connection string in modOracleSession
'     - validating that Oracle passthrough access works after login
'     - opening the Oracle password change form when requested
'     - optionally opening the application menu or startup form
'     - optionally allowing a machine-specific custom DSN override
'     - presenting login-help messaging to users
'
'
' Architecture notes
' ------------------
' This form now supports a passthrough-first runtime login model.
'
' The login process works like this:
'
'     1. User enters Oracle credentials
'     2. frmLogin creates a temporary ADODB connection to validate them
'     3. If login succeeds, the credentialed connection string is stored in
'        modOracleSession
'     4. frmLogin verifies the authenticated Oracle user directly from that ADO
'        connection
'     5. modOracleDataAccess query helpers then reuse the runtime connection
'        string for the life of the Access session
'
' Important:
'
'     - the ADODB connection itself is NOT kept open
'     - the application keeps only the credentialed connection string in memory
'     - each passthrough call still creates its own short-lived Oracle connection
'
'
' Why this form exists
' --------------------
' Access applications that use Oracle often need a login flow that is more explicit
' and user-friendly than relying on linked-table prompts or hidden ODBC behavior.
'
' This form provides:
'
'     - predictable login UX
'     - clearer error handling
'     - inline login status messaging inside the form
'     - runtime credential injection for passthrough SQL
'     - a driver-prompt fallback for expired Oracle passwords
'     - optional DSN override behavior for per-machine custom configuration
'
'
' Dependencies
' ------------
' Depends on:
'
'     - modOracleSession
'     - modOracleDataAccess
'     - modOracleAdmin
'
' Optional dependencies:
'
'     - modIniConfig (or equivalent config module) if ALLOW_CUSTOM_DSN = True
'     - Outlook automation if the Help button sends email
'     - frmOraclePasswordChange for the custom password-update workflow
'
'
' Common callers
' --------------
' Common callers include:
'
'     - startup macros
'     - application launcher forms / startup logic
'
'
' Key form events / procedures
' ----------------------------
' form lifecycle:
'     Form_Open
'
' login flow:
'     btnLogin_Click
'     btnChangePassword_Click
'
' exit / help:
'     btnExit_Click
'     btnHelp_Click
'     imgHelp_Click
'
' DSN override:
'     lblDSNName_DblClick
'     UserManageCustomDSN
'
' validation helpers:
'     ValidateModuleConfiguration
'
' password helpers:
'     OpenPasswordChangeDialog
'     PromptExpiredPasswordChangeViaDriver
'
' login status helpers:
'     ClearLoginStatus
'     SetLoginHeaderText
'     ShowLoginStatus
'     ShowLoginFailure
'     GetLoginFailureSummary
'     ShortLoginErrorDetails
'
'
' Configuration constants
' -----------------------
' The form is configured through module-level constants for behavior such as:
'
'     - application display name
'     - whether to hide the Access ribbon
'     - whether to open a menu form after login
'     - whether to populate the Windows username automatically
'     - whether to show version / environment / DSN information
'     - whether to allow custom DSN override
'     - whether to show help UI and help-email prompts
'
' Review the constants at the top of the form module before reuse in another
' application.
'
'
' Runtime session behavior
' ------------------------
' On successful login, frmLogin stores:
'
'     g_OracleConnectionString
'     g_OracleSessionUser
'
' in modOracleSession.
'
' This allows the rest of the application to execute passthrough Oracle SQL without
' rebinding linked tables or keeping a persistent ADODB connection open.
'
'
' Custom DSN support
' ------------------
' If ALLOW_CUSTOM_DSN is enabled, this form can read and store a user- or machine-
' specific DSN override through the external config module.
'
' This is useful when:
'
'     - users need different DSNs across environments
'     - a local machine needs a temporary alternate DSN
'     - testing / troubleshooting requires per-user DSN redirection
'
'
' Design notes
' ------------
' This form intentionally owns UI messaging and user interaction.
'
' Low-level Oracle modules should raise errors and avoid message boxes.
' frmLogin is the correct place to:
'
'     - present credential prompts
'     - display login failures
'     - classify common Oracle / ODBC login problems into user-friendly messages
'     - explain help options to users
'
'
' This form should not:
'
'     - contain generic Oracle query logic
'     - contain linked-table maintenance logic unrelated to login
'     - duplicate business logic that belongs in shared modules
'
'
' Version
' -------
' v1
'
'====================================================================================

Option Compare Database
Option Explicit

'--------------Options-------------
Private Const APP_DISPLAY_NAME As String = "[APP TITLE]"
Private Const LOGIN_FORM_CAPTION = "Login"
Private Const HIDE_ACCESS_RIBBON As Boolean = False
Private Const OPEN_FORM_ON_LOGIN As Boolean = False
Private Const OPEN_FORM_NM As String = "frmMain"
Private Const CHANGE_PASSWORD_FORM_NM As String = "frmOraclePasswordChange"
Private Const POP_WINDOWS_USERNM_ON_FORM As Boolean = False
Private Const SHOW_ENV_NM As Boolean = False
Private Const ALLOW_CUSTOM_DSN As Boolean = False
Private Const SHOW_LOGIN_MSGBOX As Boolean = False
Private Const SHOW_HELP_BTN As Boolean = True
Private Const HELP_EMAIL_ADDR As String = "[HELPEMAIL@HELP.DOMAIN]"
Private Const PROMPT_SENDEMAIL As Boolean = True
'----------------------------------

Private Const cHeaderTextColor As Long = 4210752
Private Const cHeaderErrorColor As Long = 192

Private mDefaultHeaderText As String

Private Sub Form_Open(Cancel As Integer)

    Dim appNameText As String
    Dim customDSN As String
    Dim textValue As String
    Dim versionValue As String: versionValue = "" 'Add version value or constant here to display on form

    On Error GoTo ErrHandler

    If Not ValidateModuleConfiguration() Then
        Cancel = True
        Exit Sub
    End If

    Call OracleSession_Clear

    If HIDE_ACCESS_RIBBON Then
        DoCmd.ShowToolbar "Ribbon", acToolbarNo
    End If

    If ALLOW_CUSTOM_DSN Then
        customDSN = Trim$(Nz(GetConfig("customDSN"), vbNullString))
        If Len(customDSN) > 0 Then
            customDSN = UCase$(customDSN)
            If customDSN <> Get_DB_DSN() Then
                Set_DB_DSN customDSN
            End If
        End If
    End If

    Me.Caption = LOGIN_FORM_CAPTION

    appNameText = APP_DISPLAY_NAME
    If SHOW_ENV_NM Then
        If UCase$(Get_DB_Env()) <> "PROD" Then
            appNameText = appNameText & " (" & Get_DB_Env() & ")"
        End If
    End If

    Me.txtAppName.ControlSource = "=""" & appNameText & """"

    textValue = "Enter your " & Get_DB_DSN() & " credentials to sign in:"
    mDefaultHeaderText = textValue
    SetLoginHeaderText mDefaultHeaderText

    If Len(Nz(versionValue, vbNullString)) > 0 Then
        textValue = "Version: " & versionValue
        Me.txtVersion.ControlSource = "=""" & textValue & """"
        Me.txtVersion.Visible = True
    Else
        Me.txtVersion.Visible = False
    End If

    If ALLOW_CUSTOM_DSN Then
        Me.lblDSNName.Caption = Get_DB_DSN()
        Me.lblDSNName.Visible = True
        If Not Me.txtVersion.Visible Then
            Me.lblDSNName.Top = Me.txtVersion.Top
        End If
    Else
        Me.lblDSNName.Visible = False
    End If

    Me.imgHelp.ControlTipText = "Enter your " & Get_DB_DSN() & " password here."
    Me.btnHelp.Visible = SHOW_HELP_BTN

    If POP_WINDOWS_USERNM_ON_FORM Then
        Me.txtUser = UCase$(Environ$("USERNAME"))
        Me.txtPass.SetFocus
    Else
        Me.txtUser.SetFocus
    End If

    Exit Sub

ErrHandler:
    MsgBox "Unable to open login form." & vbCrLf & vbCrLf & Err.Description, vbCritical

End Sub

Private Sub btnLogin_Click()

    Dim conn As Object
    Dim rs As Object
    Dim adoLoginConnectionString As String
    Dim runtimeConnectionString As String
    Dim validatedOracleUser As String
    Dim ptqValidatedUser As String
    Dim sLoginErrorText As String

    On Error GoTo ErrHandler

    ClearLoginStatus

    If Len(Trim$(Nz(Me.txtUser, vbNullString))) = 0 Then
        ShowLoginStatus "Enter a username to continue."
        Me.txtUser.SetFocus
        Exit Sub
    End If

    If Len(Trim$(Nz(Me.txtPass, vbNullString))) = 0 Then
        ShowLoginStatus "Enter a password to continue."
        Me.txtPass.SetFocus
        Exit Sub
    End If

    If Not OracleAdmin_Get_ODBC_DSN_Exists(Get_DB_DSN()) Then
        ShowLoginStatus _
            "The Oracle ODBC DSN '" & Get_DB_DSN() & "' is not set up on this machine.", _
            "Create the DSN in the Windows ODBC Data Source Administrator, or switch the application to a valid DSN."
        Me.txtUser.SetFocus
        Exit Sub
    End If

    adoLoginConnectionString = Get_ADO_Login_Conn_Str(Get_DB_DSN(), Me.txtUser, Me.txtPass)
    runtimeConnectionString = Get_ODBC_Conn_Str(Get_DB_DSN(), Me.txtUser, Me.txtPass)

    Set conn = CreateObject("ADODB.Connection")

    On Error Resume Next
    conn.Open adoLoginConnectionString

    If Err.Number <> 0 Then
        sLoginErrorText = Err.Description

        If InStr(1, sLoginErrorText, "ORA-28001", vbTextCompare) > 0 Then
            Err.Clear
            On Error GoTo ErrHandler
            ShowLoginStatus _
                "Your Oracle password is expired. The Oracle driver password-change prompt will open if the driver supports it.", _
                ShortLoginErrorDetails(sLoginErrorText)
            If PromptExpiredPasswordChangeViaDriver(Trim$(Nz(Me.txtUser, vbNullString))) Then
                ShowLoginStatus "The Oracle driver password-change prompt completed. Enter your new password and click Login again."
            End If
            Me.txtPass = Null
            Me.txtPass.SetFocus
            GoTo Cleanup
        Else
            ShowLoginFailure sLoginErrorText
            Me.txtPass.SetFocus
            GoTo Cleanup
        End If

    End If
    On Error GoTo ErrHandler

    Set rs = conn.Execute("SELECT USER FROM DUAL")

    If rs.EOF Then
        Err.Raise vbObjectError + 7100, "frmLogin.btnLogin_Click", _
                  "Login succeeded, but Oracle did not return the authenticated user."
    End If

    validatedOracleUser = UCase$(Nz(rs.Fields(0).Value, vbNullString))

    If Len(validatedOracleUser) = 0 Then
        Err.Raise vbObjectError + 7101, "frmLogin.btnLogin_Click", _
                  "Login succeeded, but the runtime Oracle session could not be validated."
    End If

    Call OracleSession_Clear
    g_OracleConnectionString = runtimeConnectionString
    g_OracleSessionUser = validatedOracleUser

    ptqValidatedUser = UCase$(PTQ_SelectString(Get_DB_DSN(), "SELECT USER FROM DUAL"))

    If Len(ptqValidatedUser) = 0 Then
        Err.Raise vbObjectError + 7102, "frmLogin.btnLogin_Click", _
                  "Passthrough validation returned a blank Oracle user."
    End If

    If StrComp(ptqValidatedUser, validatedOracleUser, vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + 7103, "frmLogin.btnLogin_Click", _
                  "Passthrough validation returned Oracle user '" & ptqValidatedUser & _
                  "' after authenticating as '" & validatedOracleUser & "'."
    End If

    DoCmd.Close acForm, Me.Name, acSaveNo

    If SHOW_LOGIN_MSGBOX Then
        MsgBox "Login to " & Get_DB_DSN() & " (" & Get_DB_DSN() & ") successful!", vbInformation
    End If

    If OPEN_FORM_ON_LOGIN Then
        DoCmd.OpenForm OPEN_FORM_NM
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
    Exit Sub

ErrHandler:
    Call OracleSession_Clear
    ShowLoginFailure Err.Description
    Resume Cleanup

End Sub

Private Sub btnExit_Click()

    Dim response As VbMsgBoxResult

    response = MsgBox("Are you sure you want to exit " & APP_DISPLAY_NAME & "?", vbQuestion + vbYesNo)

    If response = vbYes Then
        Application.Quit
    End If

End Sub

Private Sub btnHelp_Click()

    Dim response As VbMsgBoxResult
    Dim outlookApp As Object
    Dim emailItem As Object

    On Error GoTo ErrHandler

    MsgBox "For assistance with logging in to " & Get_DB_DSN() & " please contact " & HELP_EMAIL_ADDR & ".", vbInformation

    If Not PROMPT_SENDEMAIL Then Exit Sub

    response = MsgBox("Would you like to send an email now?", vbYesNo + vbQuestion)

    If response = vbYes Then
        Set outlookApp = CreateObject("Outlook.Application")
        Set emailItem = outlookApp.CreateItem(0)

        emailItem.To = HELP_EMAIL_ADDR
        emailItem.Subject = Get_DB_DSN() & " login assistance"
        emailItem.Body = "Please enter the details of your issue here."
        emailItem.Display
    End If

Cleanup:
    Set emailItem = Nothing
    Set outlookApp = Nothing
    Exit Sub

ErrHandler:
    MsgBox "Unable to open the help email workflow." & vbCrLf & vbCrLf & Err.Description, vbExclamation
    Resume Cleanup

End Sub

Private Sub btnChangePassword_Click()
    OpenPasswordChangeDialog False
End Sub

Private Sub imgHelp_Click()

    MsgBox "Enter your " & Get_DB_DSN() & " password here.", vbInformation

End Sub

Private Sub lblDSNName_DblClick(Cancel As Integer)

    If ALLOW_CUSTOM_DSN Then
        UserManageCustomDSN
    End If

End Sub

Private Function ValidateModuleConfiguration() As Boolean

    On Error GoTo HandleErr

    Call OracleAdmin_Ensure_tblConn
    Call OracleAdmin_Validate_tblConn

    If Len(Get_DB_DSN_Init()) = 0 Then
        MsgBox "tblConn.DSN_INIT is blank. Review application configuration before using frmLogin.", vbCritical
        Exit Function
    End If

    ValidateModuleConfiguration = True
    Exit Function

HandleErr:
    MsgBox "frmLogin configuration is invalid." & vbCrLf & vbCrLf & Err.Description, vbCritical

End Function

Private Sub OpenPasswordChangeDialog(Optional ByVal fromExpiredPassword As Boolean = False)

    Dim sUserName As String
    Dim sMode As String

    sUserName = Trim$(Nz(Me.txtUser, vbNullString))

    If fromExpiredPassword Then
        sMode = "EXPIRED"
    Else
        sMode = "STANDARD"
    End If

    DoCmd.OpenForm CHANGE_PASSWORD_FORM_NM, , , , , acDialog, "USER=" & sUserName & ";MODE=" & sMode

End Sub

Private Function PromptExpiredPasswordChangeViaDriver(ByVal sUserName As String) As Boolean

    Dim ws As DAO.Workspace
    Dim dbCurrent As DAO.Database
    Dim db As DAO.Database
    Dim qdfPrompt As DAO.QueryDef
    Dim rs As DAO.Recordset
    Dim sConnect As String

    On Error GoTo ErrHandler

    Call OracleSession_Clear

    Set dbCurrent = CurrentDb
    Set ws = DBEngine.CreateWorkspace( _
        "wsPwdPrompt_" & Replace$(Format$(Timer, "0.000"), ".", "_"), _
        "admin", _
        vbNullString, _
        dbUseJet)
    ws.IsolateODBCTrans = True
    Set db = ws.OpenDatabase(dbCurrent.Name)
    Set qdfPrompt = db.CreateQueryDef(vbNullString)

    sConnect = "ODBC;DSN=" & Get_DB_DSN() & ";"

    If Len(Trim$(sUserName)) > 0 Then
        sConnect = sConnect & "UID=" & UCase$(Trim$(sUserName)) & ";"
    End If

    With qdfPrompt
        .Connect = sConnect
        .returnsRecords = True
        .ODBCTimeout = 30
        .sql = "SELECT USER FROM DUAL"
    End With

    Set rs = qdfPrompt.OpenRecordset(dbOpenSnapshot)
    PromptExpiredPasswordChangeViaDriver = Not (rs.BOF And rs.EOF)

Cleanup:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set qdfPrompt = Nothing
    If Not db Is Nothing Then db.Close
    Set db = Nothing
    If Not ws Is Nothing Then ws.Close
    Set ws = Nothing
    Set dbCurrent = Nothing
    Exit Function

ErrHandler:
    ShowLoginStatus _
        "The Oracle driver password-change prompt did not complete.", _
        ShortLoginErrorDetails(Err.Description)
    Resume Cleanup

End Function

Private Sub UserManageCustomDSN()

    Dim response As Variant
    Dim question As String
    Dim hasCustomDSN As Boolean
    Dim isValidDSN As Boolean
    Dim currentFormName As String
    Dim customDSN As String

    On Error GoTo ErrHandler

    customDSN = Trim$(Nz(GetConfig("customDSN"), vbNullString))
    hasCustomDSN = (Len(customDSN) > 0)

    If hasCustomDSN Then

        question = "Remove existing custom DSN configuration?" & vbCrLf & vbCrLf & _
                   "Current custom DSN: " & customDSN & vbCrLf & _
                   "Reset to original DSN: " & Get_DB_DSN_Init()

        response = MsgBox(question, vbYesNo + vbQuestion)

        If response = vbYes Then
            DeleteConfig "customDSN"
            DeleteConfig "disableDeleteOnClose"
            Set_DB_DSN Get_DB_DSN_Init()

            currentFormName = Me.Name
            DoCmd.Close acForm, currentFormName, acSaveNo
            MsgBox "Custom DSN configuration removed.", vbInformation
            DoCmd.OpenForm currentFormName
        End If

        Exit Sub
    End If

    response = MsgBox("Do you wish to specify a custom DSN value for this application?", vbYesNo + vbQuestion)
    If response <> vbYes Then Exit Sub

    Do Until isValidDSN

        response = InputBox("Please enter the custom DSN value to use with this application:")

        If Len(Trim$(Nz(response, vbNullString))) = 0 Then Exit Sub

        isValidDSN = OracleAdmin_Get_ODBC_DSN_Exists(CStr(response))

        If Not isValidDSN Then
            question = UCase$(CStr(response)) & " is not a valid DSN name on this machine. Try again?"
            response = MsgBox(question, vbRetryCancel + vbQuestion)
            If response <> vbRetry Then Exit Sub
        End If

    Loop

    SetConfig "customDSN", UCase$(CStr(response))
    SetConfig "disableDeleteOnClose", 1

    currentFormName = Me.Name
    DoCmd.Close acForm, currentFormName, acSaveNo

    MsgBox "Custom DSN stored for this machine. (New DSN value = " & GetConfig("customDSN") & ")" & _
           vbCrLf & "The application will now exit. Relaunch to complete setup.", vbInformation

    Application.Quit
    Exit Sub

ErrHandler:
    MsgBox "Unable to update custom DSN." & vbCrLf & vbCrLf & Err.Description, vbCritical

End Sub

Private Sub ClearLoginStatus()
    SetLoginHeaderText mDefaultHeaderText
End Sub

Private Sub SetLoginHeaderText(ByVal sText As String, Optional ByVal lForeColor As Long = cHeaderTextColor)

    Me.txtHeader.ControlSource = vbNullString
    Me.txtHeader.Value = sText
    Me.txtHeader.ForeColor = lForeColor

End Sub

Private Sub ShowLoginStatus(ByVal sSummary As String, Optional ByVal sDetails As String = "")

    Dim sText As String

    sText = sSummary

    If Len(Trim$(sDetails)) > 0 Then
        sText = sText & vbCrLf & sDetails
    End If

    SetLoginHeaderText sText, cHeaderErrorColor

End Sub

Private Sub ShowLoginFailure(ByVal sErrorText As String)
    ShowLoginStatus GetLoginFailureSummary(sErrorText), ShortLoginErrorDetails(sErrorText)
End Sub

Private Function GetLoginFailureSummary(ByVal sErrorText As String) As String

    Dim sUpper As String

    sUpper = UCase$(Nz(sErrorText, vbNullString))

    If Len(sUpper) = 0 Then
        GetLoginFailureSummary = "Oracle login failed."
    ElseIf InStr(1, sUpper, "ORA-01017", vbTextCompare) > 0 Then
        GetLoginFailureSummary = "The Oracle user name or password is incorrect."
    ElseIf InStr(1, sUpper, "ORA-28000", vbTextCompare) > 0 Then
        GetLoginFailureSummary = "The Oracle user account is locked."
    ElseIf InStr(1, sUpper, "ORA-28001", vbTextCompare) > 0 Then
        GetLoginFailureSummary = "The Oracle password is expired."
    ElseIf InStr(1, sUpper, "IM002", vbTextCompare) > 0 _
        Or InStr(1, sUpper, "DATA SOURCE NAME NOT FOUND", vbTextCompare) > 0 _
        Or InStr(1, sUpper, "NO DEFAULT DRIVER SPECIFIED", vbTextCompare) > 0 Then
        GetLoginFailureSummary = "The Oracle ODBC DSN '" & Get_DB_DSN() & "' is not set up on this machine."
    ElseIf InStr(1, sUpper, "IM003", vbTextCompare) > 0 _
        Or InStr(1, sUpper, "SPECIFIED DRIVER COULD NOT BE LOADED", vbTextCompare) > 0 _
        Or InStr(1, sUpper, "DRIVER COULD NOT BE LOADED", vbTextCompare) > 0 Then
        GetLoginFailureSummary = "The Oracle ODBC client driver is not installed or could not be loaded."
    ElseIf InStr(1, sUpper, "ORA-12170", vbTextCompare) > 0 _
        Or InStr(1, sUpper, "ORA-12541", vbTextCompare) > 0 _
        Or InStr(1, sUpper, "ORA-12543", vbTextCompare) > 0 _
        Or InStr(1, sUpper, "ORA-12545", vbTextCompare) > 0 _
        Or InStr(1, sUpper, "ORA-12537", vbTextCompare) > 0 _
        Or InStr(1, sUpper, "ORA-12547", vbTextCompare) > 0 _
        Or InStr(1, sUpper, "ORA-12560", vbTextCompare) > 0 Then
        GetLoginFailureSummary = "The Oracle database could not be reached. Check VPN/network connectivity and database availability."
    ElseIf InStr(1, sUpper, "ORA-12154", vbTextCompare) > 0 _
        Or InStr(1, sUpper, "ORA-12514", vbTextCompare) > 0 _
        Or InStr(1, sUpper, "ORA-12505", vbTextCompare) > 0 Then
        GetLoginFailureSummary = "The Oracle DSN or service configuration is not resolving correctly."
    ElseIf InStr(1, sUpper, "PASSTHROUGH VALIDATION RETURNED ORACLE USER", vbTextCompare) > 0 Then
        GetLoginFailureSummary = "Login succeeded, but runtime Oracle validation returned a different user than expected."
    Else
        GetLoginFailureSummary = "Oracle login failed."
    End If

End Function

Private Function ShortLoginErrorDetails(ByVal sErrorText As String) As String

    Dim sDetails As String

    sDetails = Trim$(Nz(sErrorText, vbNullString))

    If Len(sDetails) = 0 Then Exit Function

    sDetails = Replace$(sDetails, vbCrLf, " ")
    sDetails = Replace$(sDetails, vbCr, " ")
    sDetails = Replace$(sDetails, vbLf, " ")

    Do While InStr(1, sDetails, "  ", vbBinaryCompare) > 0
        sDetails = Replace$(sDetails, "  ", " ")
    Loop

    If Len(sDetails) > 180 Then
        sDetails = Left$(sDetails, 177) & "..."
    End If

    ShortLoginErrorDetails = "Details: " & sDetails

End Function
