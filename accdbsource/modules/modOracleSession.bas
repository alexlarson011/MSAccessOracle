Attribute VB_Name = "modOracleSession"
'====================================================================================
' modOracleSession
'====================================================================================
'
' README
' ------
' Purpose
' -------
' Stores runtime Oracle session state for the current Access application session.
'
' This module exists to hold the credentialed Oracle connection information that is
' established at login time and then reused by the passthrough data-access layer for
' the life of the Access session.
'
'
' Responsibilities
' ----------------
' This module is responsible for:
'
'     - storing the runtime Oracle ODBC connection string
'     - storing the validated Oracle username
'     - clearing runtime Oracle session state
'     - reporting whether a runtime Oracle session is currently active
'
'
' Key public members
' ------------------
' runtime state:
'     g_OracleConnectionString
'     g_OracleSessionUser
'
' helpers:
'     OracleSession_Clear
'     OracleSession_IsConnected
'
'
' Typical usage
' -------------
' After successful login:
'
'     g_OracleConnectionString = loginConnectionString
'     g_OracleSessionUser = validatedOracleUser
'
' To clear the runtime session:
'
'     OracleSession_Clear
'
' To check whether a runtime session exists:
'
'     If OracleSession_IsConnected() Then ...
'
'
' Runtime behavior
' ----------------
' This module does NOT keep a live Oracle connection open.
'
' It stores a connection string only.
'
' The data-access layer creates short-lived passthrough connections as needed using
' the stored runtime connection string.
'
' This means:
'
'     - the user stays logically "logged in" for the Access session
'     - Oracle work remains stateless across individual passthrough calls
'     - code should not rely on Oracle session-specific state persisting across calls
'
'
' Dependencies
' ------------
' No module dependencies.
'
'
' Common callers
' --------------
' Common callers include:
'
'     - frmLogin
'     - modOracleDataAccess
'
'
' Design notes
' ------------
' This module should remain small and focused.
'
' It should not contain:
'
'     - SQL execution logic
'     - UI logic
'     - linked-table logic
'     - business rules
'
'
' Version
' -------
' v1
'
'====================================================================================

Option Compare Database
Option Explicit

Public g_OracleConnectionString As String
Public g_OracleSessionUser As String

Public Sub OracleSession_Clear()

    g_OracleConnectionString = vbNullString
    g_OracleSessionUser = vbNullString

End Sub

Public Function OracleSession_IsConnected() As Boolean

    OracleSession_IsConnected = (Len(g_OracleConnectionString) > 0)

End Function
