Attribute VB_Name = "modIniConfig"
'====================================================================================
' modIniConfig
'====================================================================================
'
' README
' ------
' Purpose
' -------
' Provides helpers for reading, writing, cleaning, importing, and exporting an
' INI-style configuration file for the current Access application.
'
' This module supports:
'
'     - optional section headers
'     - key / value pairs
'     - typed getters
'     - atomic file writes
'     - import to a local Access table
'     - export from a local Access table
'
'
' Responsibilities
' ----------------
' This module is responsible for:
'
'     - resolving the active INI file path
'     - reading config values from an INI file
'     - writing or updating config values in an INI file
'     - deleting config values from an INI file
'     - checking whether config values exist
'     - cleaning and sorting INI files
'     - printing config contents to the Immediate window
'     - loading config values into a local Access table
'     - exporting config values from a local Access table back to an INI file
'
'
' Key public helpers
' ------------------
' path helpers:
'     ConfigFilePath (Property Get / Let)
'
' config read / write helpers:
'     GetConfigValue
'     SetConfigValue
'     DeleteConfigValue
'     ConfigValueExists
'
' typed getters:
'     GetConfigLong
'     GetConfigBoolean

' compatibility wrappers:
'     GetConfig
'     SetConfig
'     DeleteConfig
'     ConfigKeyExists
'     GetConfigLng
'     GetConfigBool
'
' file helpers:
'     CleanIniFile
'     PrintConfig
'
' table helpers:
'     LoadConfigToTable
'     ExportConfigFromTable
'     SaveConfigCopyAs
'     DeleteLocalConfigTable
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
'     - frmLogin custom DSN logic
'     - setup / deployment routines
'     - local admin / support forms
'
'
' Design notes
' ------------
' This module raises errors instead of displaying UI.
'
' Callers are responsible for presenting any errors to the user.
'
' Reads and writes are section-aware. If sectionName is omitted, only top-level
' (non-section) keys are considered.
'
'
' Version
' -------
' v1
'
'====================================================================================

Option Compare Database
Option Explicit

Private Const cModuleName As String = "modIniConfig"

Private Const CONFIG_FILENAME As String = "My_App_Config.ini"
Private Const CONFIG_TABLENAME As String = "tblLocalConfig"
Private Const DEBUG_PRINT As Boolean = False

Private mCustomConfigPath As String

'------------------------------------------------------------------------------------
' Path helpers
'------------------------------------------------------------------------------------

Public Property Get ConfigFilePath() As String

    If Len(mCustomConfigPath) > 0 Then
        ConfigFilePath = mCustomConfigPath
    Else
        ConfigFilePath = CurrentProject.path & "\" & CONFIG_FILENAME
    End If

End Property

Public Property Let ConfigFilePath(ByVal filePath As String)
    mCustomConfigPath = Trim$(filePath)
End Property

'------------------------------------------------------------------------------------
' Read / write helpers
'------------------------------------------------------------------------------------

Public Sub SetConfigValue( _
    ByVal paramName As String, _
    ByVal paramValue As Variant, _
    Optional ByVal sectionName As String = "" _
)

    Dim iniPath As String
    Dim tempPath As String
    Dim fileNum As Integer
    Dim lines As Collection
    Dim lineText As Variant
    Dim currentSection As String
    Dim normalizedSection As String
    Dim targetLine As String
    Dim foundParam As Boolean
    Dim inserted As Boolean
    Dim sectionFound As Boolean
    Dim i As Long
    Dim tempLines As Collection

    If Len(Trim$(paramName)) = 0 Then
        Err.Raise vbObjectError + 6000, cModuleName & ".SetConfigValue", "Config param name cannot be blank."
    End If

    If IsNull(paramValue) Then paramValue = vbNullString

    On Error GoTo ErrHandler

    iniPath = ConfigFilePath
    tempPath = iniPath & ".tmp"
    normalizedSection = Trim$(sectionName)
    targetLine = paramName & "=" & CStr(paramValue)

    Set lines = New Collection
    currentSection = vbNullString
    foundParam = False

    If Dir(iniPath) <> vbNullString Then
        fileNum = FreeFile
        Open iniPath For Input Lock Read Write As #fileNum

        Do While Not EOF(fileNum)
            Line Input #fileNum, lineText
            lineText = Trim$(CStr(lineText))

            If IsSectionHeader(CStr(lineText)) Then
                currentSection = ExtractSectionName(CStr(lineText))
                lines.Add CStr(lineText)
            ElseIf IsMatchingKeyLine(CStr(lineText), paramName) And StrComp(currentSection, normalizedSection, vbTextCompare) = 0 Then
                lines.Add targetLine
                foundParam = True
            Else
                lines.Add CStr(lineText)
            End If
        Loop

        Close #fileNum
    End If

    If Not foundParam Then

        If Len(normalizedSection) = 0 Then

            Set tempLines = New Collection
            inserted = False

            For i = 1 To lines.Count
                If IsSectionHeader(CStr(lines(i))) And Not inserted Then
                    tempLines.Add targetLine
                    inserted = True
                End If
                tempLines.Add lines(i)
            Next i

            If Not inserted Then tempLines.Add targetLine
            Set lines = tempLines

        Else

            Set tempLines = New Collection
            inserted = False
            sectionFound = False
            currentSection = vbNullString

            For i = 1 To lines.Count

                If IsSectionHeader(CStr(lines(i))) Then

                    ' If we were already inside the target section and hit the next section,
                    ' insert the new key at the bottom of the target section before this header.
                    If sectionFound And Not inserted Then
                        tempLines.Add targetLine
                        inserted = True
                    End If

                    currentSection = ExtractSectionName(CStr(lines(i)))

                    If StrComp(currentSection, normalizedSection, vbTextCompare) = 0 Then
                        sectionFound = True
                    Else
                        sectionFound = False
                    End If

                End If

                tempLines.Add lines(i)

            Next i

            ' If target section was found and was the last section in the file,
            ' append the new key at the end of that section.
            If sectionFound And Not inserted Then
                tempLines.Add targetLine
                inserted = True
            End If

            ' If target section was never found, add a new section at the end.
            If Not inserted Then
                If tempLines.Count > 0 Then tempLines.Add vbNullString
                tempLines.Add "[" & normalizedSection & "]"
                tempLines.Add targetLine
            End If

            Set lines = tempLines

        End If

    End If

    WriteLinesAtomically lines, iniPath, tempPath

    If DEBUG_PRINT Then Debug.Print "[SetConfigValue] " & paramName & "=" & CStr(paramValue)
    Exit Sub

ErrHandler:
    CleanupFailedTempWrite fileNum, tempPath
    Err.Raise Err.Number, cModuleName & ".SetConfigValue", Err.Description

End Sub

Public Function GetConfigValue( _
    ByVal paramName As String, _
    Optional ByVal sectionName As String = "", _
    Optional ByVal nullOnNoValue As Boolean = False _
) As Variant

    Dim iniPath As String
    Dim fileNum As Integer
    Dim lineText As String
    Dim currentSection As String
    Dim eqPos As Long
    Dim foundCount As Long
    Dim foundValue As String
    Dim candidateKey As String

    On Error GoTo ErrHandler

    If Len(Trim$(paramName)) = 0 Then
        Err.Raise vbObjectError + 6001, cModuleName & ".GetConfigValue", "Config param name cannot be blank."
    End If

    iniPath = ConfigFilePath
    GetConfigValue = vbNullString

    If Dir(iniPath) = vbNullString Then Exit Function

    fileNum = FreeFile
    Open iniPath For Input As #fileNum

    currentSection = vbNullString

    Do While Not EOF(fileNum)
        Line Input #fileNum, lineText
        lineText = Trim$(lineText)

        If IsSectionHeader(lineText) Then
            currentSection = ExtractSectionName(lineText)
        Else
            eqPos = InStr(1, lineText, "=")
            If eqPos > 0 Then
                candidateKey = Trim$(Left$(lineText, eqPos - 1))

                If StrComp(candidateKey, paramName, vbTextCompare) = 0 _
                    And StrComp(currentSection, Trim$(sectionName), vbTextCompare) = 0 Then

                    foundCount = foundCount + 1
                    foundValue = Mid$(lineText, eqPos + 1)

                    If foundCount > 1 Then
                        Close #fileNum
                        Err.Raise vbObjectError + 6002, cModuleName & ".GetConfigValue", _
                                  "Multiple occurrences of '" & paramName & "' were found in section '" & sectionName & "'."
                    End If
                End If
            End If
        End If
    Loop

    Close #fileNum

    If foundCount = 1 Then
        GetConfigValue = foundValue
    ElseIf nullOnNoValue Then
        GetConfigValue = Null
    Else
        GetConfigValue = vbNullString
    End If

    Exit Function

ErrHandler:
    On Error Resume Next
    Close #fileNum
    Err.Raise Err.Number, cModuleName & ".GetConfigValue", Err.Description

End Function

Public Function GetConfigLong( _
    ByVal paramName As String, _
    Optional ByVal sectionName As String = "", _
    Optional ByVal errorIfNonNumerical As Boolean = True _
) As Variant

    Dim sValue As String

    On Error GoTo ErrHandler

    sValue = Trim$(Nz(GetConfigValue(paramName, sectionName), vbNullString))

    If Len(sValue) = 0 Then
        GetConfigLong = Null
    ElseIf Not IsNumeric(sValue) Then
        If errorIfNonNumerical Then
            Err.Raise vbObjectError + 6003, cModuleName & ".GetConfigLong", _
                      "Value for '" & paramName & "' is not a valid number: '" & sValue & "'"
        Else
            GetConfigLong = Null
        End If
    Else
        GetConfigLong = CLng(sValue)
    End If

    Exit Function

ErrHandler:
    Err.Raise Err.Number, cModuleName & ".GetConfigLong", Err.Description

End Function

Public Function GetConfigBoolean( _
    ByVal paramName As String, _
    Optional ByVal sectionName As String = "", _
    Optional ByVal errorIfNonBoolean As Boolean = True _
) As Variant

    Dim sValue As String

    On Error GoTo ErrHandler

    sValue = Trim$(Nz(GetConfigValue(paramName, sectionName), vbNullString))
    sValue = Replace$(Replace$(sValue, """", vbNullString), "'", vbNullString)

    If Len(sValue) = 0 Then
        GetConfigBoolean = Null
    Else
        Select Case LCase$(sValue)
            Case "true", "1", "yes", "y", "on"
                GetConfigBoolean = True
            Case "false", "0", "no", "n", "off"
                GetConfigBoolean = False
            Case Else
                If errorIfNonBoolean Then
                    Err.Raise vbObjectError + 6004, cModuleName & ".GetConfigBoolean", _
                              "Value for '" & paramName & "' is not a valid boolean: '" & sValue & "'"
                Else
                    GetConfigBoolean = Null
                End If
        End Select
    End If

    Exit Function

ErrHandler:
    Err.Raise Err.Number, cModuleName & ".GetConfigBoolean", Err.Description

End Function

Public Sub DeleteConfigValue( _
    ByVal paramName As String, _
    Optional ByVal sectionName As String = "" _
)

    Dim iniPath As String
    Dim tempPath As String
    Dim fileNum As Integer
    Dim lines As Collection
    Dim lineText As Variant
    Dim currentSection As String

    On Error GoTo ErrHandler

    If Len(Trim$(paramName)) = 0 Then
        Err.Raise vbObjectError + 6005, cModuleName & ".DeleteConfigValue", "Config param name cannot be blank."
    End If

    iniPath = ConfigFilePath
    tempPath = iniPath & ".tmp"

    If Dir(iniPath) = vbNullString Then Exit Sub

    Set lines = New Collection
    fileNum = FreeFile
    Open iniPath For Input Lock Read Write As #fileNum

    currentSection = vbNullString

    Do While Not EOF(fileNum)
        Line Input #fileNum, lineText
        lineText = Trim$(CStr(lineText))

        If IsSectionHeader(CStr(lineText)) Then
            currentSection = ExtractSectionName(CStr(lineText))
            lines.Add CStr(lineText)
        ElseIf IsMatchingKeyLine(CStr(lineText), paramName) And StrComp(currentSection, Trim$(sectionName), vbTextCompare) = 0 Then
            ' skip
        Else
            lines.Add CStr(lineText)
        End If
    Loop

    Close #fileNum

    WriteLinesAtomically lines, iniPath, tempPath

    If DEBUG_PRINT Then Debug.Print "[DeleteConfigValue] " & paramName
    Exit Sub

ErrHandler:
    CleanupFailedTempWrite fileNum, tempPath
    Err.Raise Err.Number, cModuleName & ".DeleteConfigValue", Err.Description

End Sub

Public Function ConfigValueExists( _
    ByVal paramName As String, _
    Optional ByVal sectionName As String = "" _
) As Boolean

    On Error GoTo ErrHandler

    ConfigValueExists = Not IsNull(GetConfigValue(paramName, sectionName, True))
    Exit Function

ErrHandler:
    ConfigValueExists = False

End Function

'------------------------------------------------------------------------------------
' File helpers
'------------------------------------------------------------------------------------

Public Sub CleanIniFile(Optional ByVal filePath As String = "")

    Dim iniPath As String
    Dim tempPath As String
    Dim fileNum As Integer
    Dim tempFileNum As Integer
    Dim sectionDict As Object
    Dim sectionOrder As Collection
    Dim currentSection As String
    Dim sectionLines As Collection
    Dim lineText As Variant
    Dim i As Long
    Dim j As Long
    Dim sortedParams As Variant
    Dim sectionHeader As String
    Dim key As Variant
    Dim paramLines As Collection
    Dim hasKeys As Boolean
    Dim idx As Long

    On Error GoTo ErrHandler

    iniPath = IIf(Len(Trim$(filePath)) > 0, filePath, ConfigFilePath)
    tempPath = iniPath & ".tmp"

    If Dir(iniPath) = vbNullString Then
        Err.Raise vbObjectError + 6006, cModuleName & ".CleanIniFile", _
                  "Config file not found: " & iniPath
    End If

    Set sectionDict = CreateObject("Scripting.Dictionary")
    Set sectionOrder = New Collection
    currentSection = vbNullString
    Set sectionLines = New Collection

    fileNum = FreeFile
    Open iniPath For Input As #fileNum

    Do While Not EOF(fileNum)
        Line Input #fileNum, lineText
        lineText = Trim$(CStr(lineText))

        If IsSectionHeader(CStr(lineText)) Then

            hasKeys = SectionCollectionHasKeys(sectionLines)

            If currentSection = vbNullString Or hasKeys Then
                If Not sectionDict.Exists(currentSection) Then
                    sectionDict.Add currentSection, sectionLines
                    sectionOrder.Add currentSection
                Else
                    Set sectionDict(currentSection) = sectionLines
                End If
            End If

            currentSection = ExtractSectionName(CStr(lineText))
            Set sectionLines = New Collection
            sectionLines.Add CStr(lineText)

        ElseIf Len(CStr(lineText)) > 0 Then
            sectionLines.Add CStr(lineText)
        End If
    Loop

    Close #fileNum

    hasKeys = SectionCollectionHasKeys(sectionLines)
    If currentSection = vbNullString Or hasKeys Then
        If Not sectionDict.Exists(currentSection) Then
            sectionDict.Add currentSection, sectionLines
            sectionOrder.Add currentSection
        Else
            Set sectionDict(currentSection) = sectionLines
        End If
    End If

    tempFileNum = FreeFile
    Open tempPath For Output As #tempFileNum

    For i = 1 To sectionOrder.Count

        key = sectionOrder(i)
        Set paramLines = sectionDict(key)
        sectionHeader = vbNullString

        If CStr(key) <> vbNullString And i > 1 Then
            Print #tempFileNum, vbNullString
        End If

        If CStr(key) <> vbNullString Then
            sectionHeader = CStr(paramLines(1))
            Print #tempFileNum, sectionHeader
        End If

        sortedParams = SortParamLines(GetOnlyParamLines(paramLines, CStr(key) <> vbNullString))

        If IsArrayAllocated(sortedParams) Then
            For idx = LBound(sortedParams) To UBound(sortedParams)
                Print #tempFileNum, sortedParams(idx)
            Next idx
        End If

    Next i

    Close #tempFileNum

    If Dir(iniPath) <> vbNullString Then Kill iniPath
    Name tempPath As iniPath

    If DEBUG_PRINT Then Debug.Print "[CleanIniFile] " & iniPath
    Exit Sub

ErrHandler:
    On Error Resume Next
    If fileNum <> 0 Then Close #fileNum
    If tempFileNum <> 0 Then Close #tempFileNum
    If Dir(tempPath) <> vbNullString Then Kill tempPath
    Err.Raise Err.Number, cModuleName & ".CleanIniFile", Err.Description

End Sub

Public Sub PrintConfig()

    Dim iniPath As String
    Dim fileNum As Integer
    Dim lineText As String

    On Error GoTo ErrHandler

    iniPath = ConfigFilePath

    If Dir(iniPath) = vbNullString Then
        Err.Raise vbObjectError + 6007, cModuleName & ".PrintConfig", _
                  "Config file not found: " & iniPath
    End If

    Debug.Print "[Printing configuration values stored in " & iniPath & "]"
    Debug.Print

    fileNum = FreeFile
    Open iniPath For Input Lock Read Write As #fileNum

    Do While Not EOF(fileNum)
        Line Input #fileNum, lineText
        If Len(Trim$(lineText)) > 0 Then Debug.Print lineText
    Loop

    Close #fileNum
    Exit Sub

ErrHandler:
    On Error Resume Next
    Close #fileNum
    Err.Raise Err.Number, cModuleName & ".PrintConfig", Err.Description

End Sub

'------------------------------------------------------------------------------------
' Table helpers
'------------------------------------------------------------------------------------

Public Sub LoadConfigToTable(Optional ByVal remoteIniPath As String = "")

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim tdf As DAO.TableDef
    Dim iniPath As String
    Dim fileNum As Integer
    Dim lineText As String
    Dim currentSection As String
    Dim parts As Variant
    Dim anyParamAdded As Boolean

    On Error GoTo ErrHandler

    iniPath = IIf(Len(Trim$(remoteIniPath)) > 0, remoteIniPath, ConfigFilePath)

    If Dir(iniPath) = vbNullString Then
        Err.Raise vbObjectError + 6008, cModuleName & ".LoadConfigToTable", _
                  "Config file not found: " & iniPath
    End If

    Set db = CurrentDb

    On Error Resume Next
    Set tdf = db.TableDefs(CONFIG_TABLENAME)
    On Error GoTo ErrHandler

    If tdf Is Nothing Then
        Set tdf = db.CreateTableDef(CONFIG_TABLENAME)
        With tdf.fields
            .Append tdf.CreateField("Section", dbText, 255)
            .Append tdf.CreateField("Param", dbText, 255)
            .Append tdf.CreateField("Value", dbText, 255)
        End With
        db.TableDefs.Append tdf
        db.TableDefs(CONFIG_TABLENAME).fields("Section").AllowZeroLength = True
        db.TableDefs(CONFIG_TABLENAME).fields("Value").AllowZeroLength = True
    End If

    db.Execute "DELETE * FROM " & CONFIG_TABLENAME, dbFailOnError
    Set rs = db.OpenRecordset(CONFIG_TABLENAME, dbOpenDynaset)

    fileNum = FreeFile
    Open iniPath For Input Lock Read Write As #fileNum

    currentSection = vbNullString

    Do While Not EOF(fileNum)
        Line Input #fileNum, lineText
        lineText = Trim$(lineText)

        If IsSectionHeader(lineText) Then
            currentSection = ExtractSectionName(lineText)
        ElseIf Len(lineText) > 0 And Left$(lineText, 1) <> ";" And InStr(1, lineText, "=") > 0 Then
            parts = Split(lineText, "=", 2)
            rs.AddNew
            rs!Section = currentSection
            rs!param = parts(0)
            rs!Value = parts(1)
            rs.Update
            anyParamAdded = True
        End If
    Loop

    Close #fileNum

    If Not anyParamAdded Then
        Err.Raise vbObjectError + 6009, cModuleName & ".LoadConfigToTable", _
                  "No configuration entries were found in: " & iniPath
    End If

    If DEBUG_PRINT Then Debug.Print "[LoadConfigToTable] " & iniPath
    GoTo Cleanup

ErrHandler:
    On Error Resume Next
    Close #fileNum
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, cModuleName & ".LoadConfigToTable", Err.Description

Cleanup:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set tdf = Nothing
    Set db = Nothing

End Sub

Public Sub ExportConfigFromTable(Optional ByVal filePath As String = "")

    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim rs As DAO.Recordset
    Dim iniPath As String
    Dim tempPath As String
    Dim fileNum As Integer
    Dim currentSection As String

    On Error GoTo ErrHandler

    iniPath = IIf(Len(Trim$(filePath)) > 0, filePath, ConfigFilePath)
    tempPath = iniPath & ".tmp"

    Set db = CurrentDb

    On Error Resume Next
    Set tdf = db.TableDefs(CONFIG_TABLENAME)
    On Error GoTo ErrHandler

    If tdf Is Nothing Then
        Err.Raise vbObjectError + 6010, cModuleName & ".ExportConfigFromTable", _
                  CONFIG_TABLENAME & " not found."
    End If

    If DCount("*", CONFIG_TABLENAME) = 0 Then
        Err.Raise vbObjectError + 6011, cModuleName & ".ExportConfigFromTable", _
                  CONFIG_TABLENAME & " is empty."
    End If

    Set rs = db.OpenRecordset( _
        "SELECT Section, Param, Value " & _
        "FROM " & CONFIG_TABLENAME & " " & _
        "ORDER BY Section, Param", dbOpenSnapshot)

    fileNum = FreeFile
    Open tempPath For Output As #fileNum

    currentSection = vbNullString

    Do While Not rs.EOF

        If Nz(rs!Section, vbNullString) <> currentSection Then
            If Len(currentSection) > 0 Then Print #fileNum, vbNullString
            currentSection = Nz(rs!Section, vbNullString)
            If Len(currentSection) > 0 Then
                Print #fileNum, "[" & currentSection & "]"
            End If
        End If

        Print #fileNum, rs!param & "=" & Nz(rs!Value, vbNullString)
        rs.MoveNext
    Loop

    Close #fileNum

    If Dir(iniPath) <> vbNullString Then Kill iniPath
    Name tempPath As iniPath

    CleanIniFile iniPath

    If DEBUG_PRINT Then Debug.Print "[ExportConfigFromTable] " & iniPath
    GoTo Cleanup

ErrHandler:
    On Error Resume Next
    Close #fileNum
    If Dir(tempPath) <> vbNullString Then Kill tempPath
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, cModuleName & ".ExportConfigFromTable", Err.Description

Cleanup:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set tdf = Nothing
    Set db = Nothing

End Sub

Public Sub SaveConfigCopyAs(ByVal filePath As String)

    If Len(Trim$(filePath)) = 0 Then
        Err.Raise vbObjectError + 6012, cModuleName & ".SaveConfigCopyAs", _
                  "Target file path cannot be blank."
    End If

    ExportConfigFromTable filePath

End Sub

Public Sub DeleteLocalConfigTable()

    Dim db As DAO.Database
    Dim tdf As DAO.TableDef

    On Error GoTo ErrHandler

    Set db = CurrentDb

    On Error Resume Next
    Set tdf = db.TableDefs(CONFIG_TABLENAME)
    On Error GoTo ErrHandler

    If Not tdf Is Nothing Then
        db.TableDefs.Delete CONFIG_TABLENAME
    End If

    If DEBUG_PRINT Then Debug.Print "[DeleteLocalConfigTable] " & CONFIG_TABLENAME
    GoTo Cleanup

ErrHandler:
    Err.Raise Err.Number, cModuleName & ".DeleteLocalConfigTable", Err.Description

Cleanup:
    Set tdf = Nothing
    Set db = Nothing

End Sub

'------------------------------------------------------------------------------------
' Internal helpers
'------------------------------------------------------------------------------------

Private Function IsSectionHeader(ByVal lineText As String) As Boolean
    IsSectionHeader = (Left$(Trim$(lineText), 1) = "[" And Right$(Trim$(lineText), 1) = "]")
End Function

Private Function ExtractSectionName(ByVal lineText As String) As String
    If IsSectionHeader(lineText) Then
        ExtractSectionName = Mid$(Trim$(lineText), 2, Len(Trim$(lineText)) - 2)
    End If
End Function

Private Function IsMatchingKeyLine(ByVal lineText As String, ByVal paramName As String) As Boolean

    Dim eqPos As Long
    Dim keyName As String

    eqPos = InStr(1, lineText, "=")
    If eqPos < 1 Then Exit Function

    keyName = Trim$(Left$(lineText, eqPos - 1))
    IsMatchingKeyLine = (StrComp(keyName, paramName, vbTextCompare) = 0)

End Function

Private Sub WriteLinesAtomically( _
    ByRef lines As Collection, _
    ByVal iniPath As String, _
    ByVal tempPath As String _
)

    Dim fileNum As Integer
    Dim lineText As Variant

    fileNum = FreeFile
    Open tempPath For Output Lock Read Write As #fileNum

    For Each lineText In lines
        Print #fileNum, CStr(lineText)
    Next lineText

    Close #fileNum

    If Dir(iniPath) <> vbNullString Then Kill iniPath
    Name tempPath As iniPath

End Sub

Private Sub CleanupFailedTempWrite(ByVal fileNum As Integer, ByVal tempPath As String)

    On Error Resume Next
    Close #fileNum
    If Len(tempPath) > 0 Then
        If Dir(tempPath) <> vbNullString Then Kill tempPath
    End If

End Sub

Private Function SectionCollectionHasKeys(ByRef sectionLines As Collection) As Boolean

    Dim i As Long
    Dim lineText As String

    For i = 1 To sectionLines.Count
        lineText = CStr(sectionLines(i))
        If InStr(1, lineText, "=") > 0 And Left$(Trim$(lineText), 1) <> ";" Then
            SectionCollectionHasKeys = True
            Exit Function
        End If
    Next i

End Function

Private Function GetOnlyParamLines( _
    ByRef sectionLines As Collection, _
    ByVal hasSectionHeader As Boolean _
) As Collection

    Dim result As Collection
    Dim i As Long
    Dim startPos As Long
    Dim lineText As String

    Set result = New Collection

    If hasSectionHeader Then
        startPos = 2
    Else
        startPos = 1
    End If

    For i = startPos To sectionLines.Count
        lineText = CStr(sectionLines(i))
        If InStr(1, lineText, "=") > 0 And Left$(Trim$(lineText), 1) <> ";" Then
            result.Add lineText
        End If
    Next i

    Set GetOnlyParamLines = result

End Function

Private Function SortParamLines(ByRef paramPairs As Collection) As Variant

    Dim arr() As String
    Dim i As Long
    Dim j As Long
    Dim swapped As Boolean
    Dim tempValue As String

    If paramPairs.Count = 0 Then Exit Function

    ReDim arr(1 To paramPairs.Count)

    For i = 1 To paramPairs.Count
        arr(i) = paramPairs(i)
    Next i

    Do
        swapped = False
        For i = LBound(arr) To UBound(arr) - 1
            If CompareParams(arr(i), arr(i + 1)) > 0 Then
                tempValue = arr(i)
                arr(i) = arr(i + 1)
                arr(i + 1) = tempValue
                swapped = True
            End If
        Next i
    Loop While swapped

    SortParamLines = arr

End Function

Private Function CompareParams(ByVal a As String, ByVal b As String) As Integer

    Dim aName As String
    Dim bName As String
    Dim aIsNum As Boolean
    Dim bIsNum As Boolean

    aName = LCase$(Trim$(Split(a, "=")(0)))
    bName = LCase$(Trim$(Split(b, "=")(0)))

    aIsNum = IsNumeric(Left$(aName, 1))
    bIsNum = IsNumeric(Left$(bName, 1))

    If aIsNum And Not bIsNum Then
        CompareParams = -1
    ElseIf Not aIsNum And bIsNum Then
        CompareParams = 1
    Else
        CompareParams = StrComp(aName, bName, vbTextCompare)
    End If

End Function

Private Function IsArrayAllocated(ByVal arr As Variant) As Boolean
    On Error GoTo HandleErr
    If IsArray(arr) Then
        Dim l As Long
        Dim u As Long
        l = LBound(arr)
        u = UBound(arr)
        IsArrayAllocated = (u >= l)
    End If
    Exit Function
HandleErr:
    IsArrayAllocated = False
End Function

'------------------------------------------------------------------------------------
' Backward-compatible wrappers
'------------------------------------------------------------------------------------

Public Function GetConfig( _
    ByVal paramName As String, _
    Optional ByVal nullOnNoVal As Boolean = False _
) As Variant
    GetConfig = GetConfigValue(paramName, vbNullString, nullOnNoVal)
End Function

Public Sub SetConfig( _
    ByVal paramName As String, _
    ByVal paramValue As Variant, _
    Optional ByVal sectionName As String = "" _
)
    SetConfigValue paramName, paramValue, sectionName
End Sub

Public Sub DeleteConfig( _
    ByVal paramName As String, _
    Optional ByVal sectionName As String = "" _
)
    DeleteConfigValue paramName, sectionName
End Sub

Public Function ConfigKeyExists(ByVal paramName As String) As Boolean
    ConfigKeyExists = ConfigValueExists(paramName)
End Function

Public Function GetConfigLng( _
    ByVal paramName As String, _
    Optional ByVal errorIfNonNumerical As Boolean = True _
) As Variant
    GetConfigLng = GetConfigLong(paramName, vbNullString, errorIfNonNumerical)
End Function

Public Function GetConfigBool( _
    ByVal paramName As String, _
    Optional ByVal errorIfNonBool As Boolean = True _
) As Variant
    GetConfigBool = GetConfigBoolean(paramName, vbNullString, errorIfNonBool)
End Function
