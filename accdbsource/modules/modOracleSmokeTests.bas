Attribute VB_Name = "modOracleSmokeTests"
'====================================================================================
' modOracleSmokeTests
'====================================================================================
'
' README
' ------
' Purpose
' -------
' Provides lightweight developer smoke tests for the stateless Oracle runtime layer.
'
' These routines are intended to be run manually from the Immediate Window after
' building the Access database and logging in through frmLogin.
'
'
' Responsibilities
' ----------------
' This module is responsible for:
'
'     - checking that a runtime Oracle session is active
'     - checking that the runtime passthrough path authenticates as the expected user
'     - checking scalar query execution
'     - checking row materialization into dictionaries
'     - checking duplicate-column diagnostics when developers forget SQL aliases
'
'
' Key public helpers
' ------------------
' smoke tests:
'     OracleSmoke_RunRuntime
'     OracleSmoke_TestDuplicateColumnDiagnostic
'
'
' Typical usage
' -------------
' After logging in:
'
'     OracleSmoke_RunRuntime
'
' Optional diagnostic:
'
'     OracleSmoke_TestDuplicateColumnDiagnostic
'
'
' Design notes
' ------------
' These tests intentionally avoid application tables so they can run against any
' reachable Oracle database. Write tests should live next to a known disposable test
' table or fixture.
'
'
' Version
' -------
' v1
'
'====================================================================================

Option Compare Database
Option Explicit

Private Const cModuleName As String = "modOracleSmokeTests"

Public Sub OracleSmoke_RunRuntime(Optional ByVal dsn As String = "")

    Dim expectedUser As String
    Dim actualUser As String
    Dim scalarValue As Long
    Dim rows As Collection
    Dim rowData As Object

    On Error GoTo ErrHandler

    Debug.Print String$(80, "-")
    Debug.Print "OracleSmoke_RunRuntime starting"

    RequireOracleSession

    expectedUser = UCase$(Trim$(g_OracleSessionUser))
    actualUser = UCase$(Trim$(PTQ_SelectString("SELECT USER FROM DUAL", dsn)))

    If Len(expectedUser) > 0 Then
        If StrComp(actualUser, expectedUser, vbTextCompare) <> 0 Then
            Err.Raise vbObjectError + 9201, cModuleName & ".OracleSmoke_RunRuntime", _
                      "Expected Oracle user " & expectedUser & " but query returned " & actualUser & "."
        End If
    End If

    Debug.Print "  Runtime user: " & actualUser

    scalarValue = PTQ_SelectLong("SELECT 1 FROM DUAL", dsn)
    OracleSmoke_Assert scalarValue = 1, "Scalar query did not return 1."
    Debug.Print "  Scalar query: OK"

    Set rows = PTQ_GetRows("SELECT 1 AS TEST_VALUE, 'OK' AS TEST_TEXT FROM DUAL", dsn)
    OracleSmoke_Assert rows.Count = 1, "Expected one row from row materialization smoke test."

    Set rowData = rows(1)
    OracleSmoke_Assert CLng(rowData("TEST_VALUE")) = 1, "TEST_VALUE was not materialized correctly."
    OracleSmoke_Assert CStr(rowData("TEST_TEXT")) = "OK", "TEST_TEXT was not materialized correctly."
    Debug.Print "  Row materialization: OK"

    Debug.Print "OracleSmoke_RunRuntime completed successfully."
    Debug.Print String$(80, "-")
    Exit Sub

ErrHandler:
    Debug.Print "OracleSmoke_RunRuntime failed: " & Err.Number & " - " & Err.Description
    Debug.Print String$(80, "-")
    Err.Raise Err.Number, cModuleName & ".OracleSmoke_RunRuntime", Err.Description

End Sub

Public Sub OracleSmoke_TestDuplicateColumnDiagnostic(Optional ByVal dsn As String = "")

    Dim rows As Collection

    On Error GoTo ExpectedErr

    Debug.Print String$(80, "-")
    Debug.Print "OracleSmoke_TestDuplicateColumnDiagnostic starting"

    Set rows = PTQ_GetRows("SELECT 1 AS DUP_VALUE, 2 AS DUP_VALUE FROM DUAL", dsn)

    Err.Raise vbObjectError + 9202, cModuleName & ".OracleSmoke_TestDuplicateColumnDiagnostic", _
              "Duplicate-column diagnostic did not fire. Returned rows: " & CStr(rows.Count)

ExpectedErr:
    If InStr(1, Err.Description, "duplicate column name", vbTextCompare) = 0 Then
        Debug.Print "OracleSmoke_TestDuplicateColumnDiagnostic failed: " & Err.Number & " - " & Err.Description
        Debug.Print String$(80, "-")
        Err.Raise Err.Number, cModuleName & ".OracleSmoke_TestDuplicateColumnDiagnostic", Err.Description
    End If

    Debug.Print "  Duplicate-column diagnostic: OK"
    Debug.Print String$(80, "-")

End Sub

Private Sub OracleSmoke_Assert(ByVal condition As Boolean, ByVal message As String)

    If Not condition Then
        Err.Raise vbObjectError + 9200, cModuleName & ".OracleSmoke_Assert", message
    End If

End Sub
