Attribute VB_Name = "ConfigManager"

' =============================================================================
' ConfigManager Module - Settings Management
' =============================================================================

Public Function IsFirstRun() As Boolean
    IsFirstRun = (GetSetting(FIRST_RUN_KEY) <> "False")
End Function

Public Function IsTutorialShown() As Boolean
    IsTutorialShown = (GetSetting("TutorialShown") = "True")
End Function

Public Sub SetTutorialShown()
    SetSetting "TutorialShown", "True"
End Sub

Public Sub ClearTutorialFlag()
    RemoveSetting "TutorialShown"
End Sub

Public Function GetUserTotalChars() As Long
    Dim userTotal As String
    userTotal = GetSetting(USER_TOTAL_CHARS_KEY)
    
    If IsNumeric(userTotal) And Val(userTotal) > 0 Then
        GetUserTotalChars = CLng(Val(userTotal))
    Else
        GetUserTotalChars = Len(ActiveDocument.Range.text)
    End If
End Function
Public Sub SetUserTotalChars(newTotal As Long)
    SetSetting USER_TOTAL_CHARS_KEY, CStr(newTotal)
    ' Cache invalidation
    HeadingProcessor.InvalidateCache
End Sub

Public Sub SaveSettings(settings As Object)
    SetSetting FIRST_RUN_KEY, "False"
    SetSetting USER_TOTAL_CHARS_KEY, CStr(settings("TotalChars"))
    SetSetting AUTO_SAVE_KEY, IIf(settings("AutoSave"), "True", "False")
    
    ' Speech Time disabled by default
    If GetSetting("DisplaySpeechTime") = "" Then
        SetSetting "DisplaySpeechTime", "False"
    End If
    
    ' Default value for Speech Tempo (if enabled later)
    If GetSetting("SpeechTempo") = "" Then
        SetSetting "SpeechTempo", CStr(DEFAULT_SPEECH_TEMPO)
    End If
End Sub

Public Sub ClearAllSettings()
    RemoveSetting FIRST_RUN_KEY
    RemoveSetting USER_TOTAL_CHARS_KEY
End Sub

' General setting handler functions
Public Function GetSetting(propName As String) As String
    On Error Resume Next
    GetSetting = ActiveDocument.CustomDocumentProperties(propName).value
    If Err.Number <> 0 Then GetSetting = ""
    Err.Clear
End Function

Public Sub SetSetting(propName As String, propValue As String)
    On Error Resume Next
    ActiveDocument.CustomDocumentProperties(propName).value = propValue
    If Err.Number <> 0 Then
        ActiveDocument.CustomDocumentProperties.Add _
            Name:=propName, _
            LinkToContent:=False, _
            Type:=msoPropertyTypeString, _
            value:=propValue
    End If
    Err.Clear
End Sub

Public Sub RemoveSetting(propName As String)
    On Error Resume Next
    ActiveDocument.CustomDocumentProperties(propName).Delete
    Err.Clear
End Sub
