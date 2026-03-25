Attribute VB_Name = "Main"

' =============================================================================
' RIBBON Buttons
' =============================================================================

Public Sub InitializeModules()
    On Error Resume Next
    ' Initialize modules
    Call InitializeErrorManager
    
    ' Check default settings
    If ConfigManager.GetSetting("AutoSave") = "" Then
        ConfigManager.SetSetting "AutoSave", "False"
    End If
    
    ' NEW: Speech Time disabled by default
    If ConfigManager.GetSetting("DisplaySpeechTime") = "" Then
        ConfigManager.SetSetting "DisplaySpeechTime", "False"
    End If
    
    ' NEW: Default value for Speech Tempo
    if ConfigManager.GetSetting("SpeechTempo") = "" Then
        ConfigManager.SetSetting "SpeechTempo", CStr(DEFAULT_SPEECH_TEMPO)
    End If
    
    If Err.Number <> 0 Then
        Debug.Print "InitializeModules error: " & Err.description
        Err.Clear
    End If
    On Error GoTo 0
End Sub

Public Sub OpenFeedbackPage(control As IRibbonControl)
    ActiveDocument.FollowHyperlink "https://github.com/textbalance/macro/issues"
End Sub

Public Sub OpenBuyMeACoffee(control As IRibbonControl)
    ActiveDocument.FollowHyperlink "https://buymeacoffee.com/textbalance"
End Sub

Sub TextBalance(Optional control As Object = Nothing)
    Call InitializeErrorManager
    Dim ur As UndoRecord
    Set ur = Application.UndoRecord
    ur.StartCustomRecord "Run TextBalance"
    ActiveDocument.ActiveWindow.View.ShowHiddenText = True
    On Error GoTo ErrorHandler
    Call InitializeErrorManager
    
    Application.ScreenUpdating = False
    If ConfigManager.IsFirstRun() Then
        TextBalanceInstall
    Else
        TextBalanceUpdate
    End If

    Application.ScreenUpdating = True
    
    ur.EndCustomRecord
    Exit Sub
    
ErrorHandler:
    ur.EndCustomRecord
    ErrorManager.HandleError "TextBalance", Err.description, esCritical, ecGeneral
End Sub

Public Sub SetSpeechTempo(Optional control As Object = Nothing)
    Dim userInput As String
    Dim tempo As Integer
    
    If ConfigManager.GetSetting("DisplaySpeechTime") = "True" Then
        Dim previous As Integer
        previous = Val(ConfigManager.GetSetting("SpeechTempo"))
    Else
        previous = 180
    End If
    
    userInput = InputBox("Enter speech tempo (50-1000 characters/minute):", "Speech Tempo", previous)
    
    If userInput = "" Then Exit Sub
    
    If IsNumeric(userInput) Then
        tempo = CInt(userInput)
        If tempo >= 50 And tempo <= 1000 Then
            ConfigManager.SetSetting "SpeechTempo", CStr(tempo)
            TextBalance
        Else
            MsgBox "Please enter a number between 50 and 1000", vbExclamation
        End If
    Else
        MsgBox "Please enter a valid number", vbExclamation
    End If
End Sub

Public Sub ToggleDisplayMode(control As IRibbonControl, pressed As Boolean)
    ConfigManager.SetSetting "DisplaySpeechTime", IIf(pressed, "True", "False")
    TextBalance
End Sub
Public Sub ToggleAutoSave(control As IRibbonControl, pressed As Boolean)
    ConfigManager.SetSetting "AutoSave", IIf(pressed, "True", "False")
End Sub

Public Sub GetDisplayModeState(control As IRibbonControl, ByRef returnedVal)
    ' Ribbon callback - DisplayMode button state
    returnedVal = (ConfigManager.GetSetting("DisplaySpeechTime") = "True")
End Sub

Public Sub GetAutoSaveState(control As IRibbonControl, ByRef returnedVal)
    ' Ribbon callback - AutoSave button state
    returnedVal = (ConfigManager.GetSetting("AutoSave") = "True")
End Sub

Sub TextBalanceRemove(Optional control As Object = Nothing)
    If Not UIHelpers.ConfirmRemoval() Then Exit Sub
    
    Dim ur As UndoRecord
    Set ur = Application.UndoRecord
    ur.StartCustomRecord "Remove TextBalance"
    ActiveDocument.ActiveWindow.View.ShowHiddenText = True
    Application.ScreenUpdating = False

    AnnotationManager.RemoveAnnotationCharacters
    Dim tbl As table
    Set tbl = FindExistingTable()
    
    If Not tbl Is Nothing Then
        tbl.Delete
    End If
    ConfigManager.ClearAllSettings
    
    Application.ScreenUpdating = True
    UIHelpers.ShowRemovalSuccess
    ur.EndCustomRecord
    Exit Sub
End Sub

Sub AnnotationRemove(Optional control As Object = Nothing)
    Dim ur As UndoRecord
    Set ur = Application.UndoRecord
    ur.StartCustomRecord "Remove Annotations"
    ActiveDocument.ActiveWindow.View.ShowHiddenText = True
    Application.ScreenUpdating = False
    
    ActiveDocument.Content.Characters(1).Font.Superscript = True 'dummy to prevent crash on empty running .ur
    AnnotationManager.RemoveAnnotationCharacters
    ActiveDocument.Content.Characters(1).Font.Superscript = False
    
    Application.ScreenUpdating = True
    UIHelpers.ShowAnnotationRemovalSuccess
    ur.EndCustomRecord
    Exit Sub
End Sub

Public Sub TableRemove(Optional control As Object = Nothing)
    Dim ur As UndoRecord
    Set ur = Application.UndoRecord
    ur.StartCustomRecord "Remove Table"
    ActiveDocument.ActiveWindow.View.ShowHiddenText = True
    Application.ScreenUpdating = False
    
    Dim tbl As table
    Set tbl = FindExistingTable()
    
    If Not tbl Is Nothing Then
        tbl.Delete
    End If
    
    Application.ScreenUpdating = True
    UIHelpers.ShowTableRemovalSuccess
    ur.EndCustomRecord
    Exit Sub
End Sub

Public Sub HelpButton(Optional control As Object = Nothing)
    ShowHTMLHelp
End Sub


'==================================
'Main functions
'==================================

Sub TextBalanceInstall()
    If Not UIHelpers.ShowInstallWelcome() Then Exit Sub
    
    Dim settings As Object
    Set settings = UIHelpers.GetInstallationSettings()
    If settings Is Nothing Then Exit Sub
    
    ' Saving settings
    ConfigManager.SaveSettings settings
    
    ' Fix ligatures in heading styles to prevent ZWNJ display
    FixHeadingLigatures
    
    If Not PreFlightCheck() Then Exit Sub
    ' Performing installation
    If PerformInstallation() Then
        UIHelpers.ShowInstallSuccess
    Else
        ErrorManager.ValidateDocument "Validation"
    End If
End Sub

Private Sub FixHeadingLigatures()
    On Error Resume Next
    ActiveDocument.Styles(wdStyleHeading1).Font.Ligatures = wdLigaturesStandardContextual
    ActiveDocument.Styles(wdStyleHeading2).Font.Ligatures = wdLigaturesStandardContextual
    Err.Clear
End Sub

Public Sub CheckForUpdatesButton(control As IRibbonControl)
    UpdateChecker.CheckForUpdates showUpToDate:=True
End Sub

Sub TextBalanceUpdate()
    Dim startTime As Double
    startTime = Timer
    Dim currentTableTotal As Long
    currentTableTotal = TableManager.GetCurrentTableTotalChars()
    If currentTableTotal <> ConfigManager.GetUserTotalChars() Then
         ConfigManager.SetUserTotalChars currentTableTotal
    End If
    If PreFlightCheck() Then
        If PerformUpdate() Then
            Dim runtime As Double
            runtime = Timer - startTime
            Debug.Print "runtime: " & runtime
            If ConfigManager.GetSetting("AutoSave") = "True" Then
                If Not ActiveDocument.Saved Then
                    ActiveDocument.Save
                End If
            End If
        End If
    Else
        Exit Sub
    End If
End Sub


Private Function PreFlightCheck() As Boolean
    PreFlightCheck = False
    If ErrorManager.ValidateDocument("validation") Then
        PreFlightCheck = True
    Else
        Exit Function
    End If
End Function

Private Function PerformInstallation() As Boolean
    On Error GoTo ErrorHandler
    Dim tbl As table
    Set tbl = TableManager.GetOrCreateTable()
    If tbl Is Nothing Then GoTo ErrorHandler
    Dim headings As Collection
    Set headings = HeadingProcessor.ProcessAllHeadingsWithTable(tbl)
    If headings Is Nothing Then GoTo ErrorHandler
    AnnotationManager.AddAnnotations headings
    TableManager.UpdateTable tbl, headings

    PerformInstallation = True
    Exit Function
    
ErrorHandler:
    ErrorManager.HandleError "installation", Err.description, esError, ecGeneral
    PerformInstallation = False
End Function

Private Function PerformUpdate() As Boolean
    On Error GoTo ErrorHandler
    Dim tbl As table
    Set tbl = TableManager.GetOrCreateTable()
    If tbl Is Nothing Then GoTo ErrorHandler
 
    Dim headings As Collection
    Set headings = HeadingProcessor.ProcessAllHeadingsWithTable(tbl)
   
    AnnotationManager.RemoveAnnotationCharacters
    AnnotationManager.AddAnnotations headings
    TableManager.UpdateTable tbl, headings
   
    PerformUpdate = True
    Exit Function

ErrorHandler:
    ErrorManager.HandleError "Update", Err.description, esError, ecGeneral
    PerformUpdate = False
End Function
