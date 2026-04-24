Attribute VB_Name = "modDeveloperTools"
'====================================================================================
' modDeveloperTools
'====================================================================================
'
' README
' ------
' Purpose
' -------
' Contains developer-only helper routines for working with exported Access/VBA source.
'
' These helpers are intended for use from the Immediate Window while building forms,
' SQL strings, and scaffolding. They should not be called by runtime application code.
'
'
' Responsibilities
' ----------------
' This module is responsible for:
'
'     - formatting pasted Oracle SQL as copy/paste-ready Access VBA string literals
'     - printing long generated text safely to the Immediate Window
'
'
' Key public helpers
' ------------------
' SQL formatting:
'     Dev_FormatAccessSqlLiteral
'     Dev_PrintAccessSqlLiteral
'
'
' Typical usage
' -------------
' From the Immediate Window:
'
'     Dev_PrintAccessSqlLiteral _
'         "SELECT RECORD_ID, RECORD_NAME" & vbCrLf & _
'         "FROM APP_RECORD" & vbCrLf & _
'         "WHERE STATUS_CD = 'A'"
'
' This prints:
'
'     sSQL = _
'         "SELECT RECORD_ID, RECORD_NAME" & vbCrLf & _
'         "FROM APP_RECORD" & vbCrLf & _
'         "WHERE STATUS_CD = 'A'"
'
' For very long SQL, use incremental assignment output:
'
'     Dev_PrintAccessSqlLiteral sSqlText, "sSQL", True, True
'
'
' Design notes
' ------------
' This module intentionally avoids dependencies on Oracle, DAO, ADO, Access forms, or
' the runtime session layer.
'
'
' Version
' -------
' v1
'
'====================================================================================

Option Compare Database
Option Explicit

Private Const cModuleName As String = "modDeveloperTools"
Private Const cDebugPrintChunkLength As Long = 900

'------------------------------------------------------------------------------------
' SQL literal formatting
'------------------------------------------------------------------------------------

Public Function Dev_FormatAccessSqlLiteral( _
    ByVal oracleSql As String, _
    Optional ByVal variableName As String = "sSQL", _
    Optional ByVal includeSqlLineBreaks As Boolean = True, _
    Optional ByVal useIncrementalAssignments As Boolean = False _
) As String

    Dim normalizedSql As String
    Dim lines() As String
    Dim i As Long
    Dim result As String
    Dim lineText As String
    Dim literalPart As String
    Dim lineBreakPart As String

    If Len(Trim$(variableName)) = 0 Then
        Err.Raise vbObjectError + 9000, cModuleName & ".Dev_FormatAccessSqlLiteral", _
                  "Variable name cannot be blank."
    End If

    normalizedSql = Replace$(oracleSql, vbCrLf, vbLf)
    normalizedSql = Replace$(normalizedSql, vbCr, vbLf)
    lines = Split(normalizedSql, vbLf)

    If useIncrementalAssignments Then
        result = variableName & " = vbNullString" & vbCrLf
    Else
        result = variableName & " = _" & vbCrLf
    End If

    For i = LBound(lines) To UBound(lines)

        lineText = Replace$(lines(i), """", """""")
        literalPart = """" & lineText & """"

        If includeSqlLineBreaks And i < UBound(lines) Then
            lineBreakPart = " & vbCrLf"
        Else
            lineBreakPart = vbNullString
        End If

        If useIncrementalAssignments Then
            result = result & variableName & " = " & variableName & " & " & literalPart & lineBreakPart
        Else
            result = result & "    " & literalPart & lineBreakPart

            If i < UBound(lines) Then
                result = result & " & _"
            End If
        End If

        If i < UBound(lines) Then
            result = result & vbCrLf
        End If

    Next i

    Dev_FormatAccessSqlLiteral = result

End Function

Public Sub Dev_PrintAccessSqlLiteral( _
    ByVal oracleSql As String, _
    Optional ByVal variableName As String = "sSQL", _
    Optional ByVal includeSqlLineBreaks As Boolean = True, _
    Optional ByVal useIncrementalAssignments As Boolean = False _
)

    Dim formattedText As String

    formattedText = Dev_FormatAccessSqlLiteral( _
        oracleSql, _
        variableName, _
        includeSqlLineBreaks, _
        useIncrementalAssignments)

    Dev_PrintLongText formattedText

End Sub

'------------------------------------------------------------------------------------
' Immediate Window helpers
'------------------------------------------------------------------------------------

Public Sub Dev_PrintLongText(ByVal textToPrint As String)

    Dim lines() As String
    Dim i As Long

    textToPrint = Replace$(textToPrint, vbCrLf, vbLf)
    textToPrint = Replace$(textToPrint, vbCr, vbLf)
    lines = Split(textToPrint, vbLf)

    For i = LBound(lines) To UBound(lines)
        Dev_PrintDebugLine lines(i)
    Next i

End Sub

Private Sub Dev_PrintDebugLine(ByVal textLine As String)

    Dim lStartPos As Long

    If Len(textLine) = 0 Then
        Debug.Print vbNullString
        Exit Sub
    End If

    lStartPos = 1

    Do While lStartPos <= Len(textLine)
        Debug.Print Mid$(textLine, lStartPos, cDebugPrintChunkLength)
        lStartPos = lStartPos + cDebugPrintChunkLength
    Loop

End Sub
