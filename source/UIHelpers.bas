Attribute VB_Name = "UIHelpers"
' =============================================================================
' UIHelpers Module - Unified User Interface
' =============================================================================

Public Sub HandleError(functionName As String, errorMsg As String)
    ErrorManager.HandleError functionName, errorMsg, esError, ecGeneral
End Sub

Public Sub ShowError(message As String)
    MsgBox message, vbCritical, "TextBalance Error"
End Sub

Public Function ShowInstallWelcome() As Boolean
    Dim msg As String
    msg = ChrW(182) & " Welcome to TextBalance!" & vbCrLf & vbCrLf & _
          "This add-in will help you:" & vbCrLf & _
          "  - Track document structure and balance" & vbCrLf & _
          "  - Set target document and heading lengths" & vbCrLf & _
          "  - Monitor speech time for presentations" & vbCrLf & _
          "  - Get visual feedback on document progress" & vbCrLf & vbCrLf & _
          "Do you want to add it to this document?"
          
    ShowInstallWelcome = (MsgBox(msg, vbYesNo + vbQuestion, "Adding TextBalance to document") = vbYes)
End Function

Public Function GetInstallationSettings() As Object
    Dim settings As Object
    Set settings = CreateObject("Scripting.Dictionary")
    
    ' Document length target
    Dim currentChars As Long
    currentChars = Len(ActiveDocument.Range.text)
    
    Dim userInput As String
    userInput = InputBox("The TextBalance add-in will:" & vbCrLf & _
                        "  - Create a summary table from headings" & vbCrLf & _
                        "  - Add progress indicators to headings" & vbCrLf & vbCrLf & _
                        "Default ideal:" & vbCrLf & _
                        "    available ideal% / headings without ideal%" & vbCrLf & _
                        "Default tolerance: 5%" & vbCrLf & vbCrLf & _
                        "Set your target document length:" & vbCrLf & _
                        "(Current document: " & Format(currentChars, "# ##0") & " characters)" & vbCrLf & vbCrLf & _
                        "The actual percentage of headings and progress indicators will be based on Target Document Length, which can be set below:", _
                        "Target Document Length", CStr(currentChars))
    If userInput = "" Then
        Set GetInstallationSettings = Nothing
        Exit Function
    End If
    
    If Not IsNumeric(userInput) Or Val(userInput) <= 0 Then
        userInput = CStr(currentChars)
    End If
    
    settings("TotalChars") = CLng(Val(userInput))
    ' AutoSave disabled by default
    settings("AutoSave") = False
    
    Set GetInstallationSettings = settings
End Function

Public Sub ShowInstallSuccess()
    MsgBox " TextBalance added successfully!" & vbCrLf & vbCrLf & _
           "  - Summary table created" & vbCrLf & _
           "  - Heading indicators added" & vbCrLf & _
           "  - Settings saved to document" & vbCrLf & vbCrLf & _
           "Use the TextBalance tab (or Alt+,+, shortcut) to refresh data or adjust settings." & vbCrLf & vbCrLf & _
           "IMPORTANT: If you want faster performance, remove Table of Contents from the document.", _
           vbInformation, ChrW(182) & " Complete"
           
    ' Show tutorial
    ShowInstallTutorial
End Sub

Public Sub ShowInstallError()
    MsgBox "Installation encountered an error" & vbCrLf & vbCrLf & _
           "Please check your document structure and try again. " & _
           "Ensure your document has properly formatted headings (Heading 1, Heading 2).", _
           vbCritical, "Installation Error"
End Sub

Public Sub ShowUpdateSuccess(runtime As Double)
    Dim timeStr As String
    If runtime < 1 Then
        timeStr = Format(runtime * 1000, "0") & " ms"
    Else
        timeStr = Format(runtime, "0.0") & " seconds"
    End If
    
    MsgBox ChrW(182) & " TextBalance updated successfully!" & vbCrLf & vbCrLf & _
           " Analysis complete in " & timeStr & vbCrLf & _
           " Progress indicators refreshed" & vbCrLf & _
           " Summary table updated", _
           vbInformation, "Update Complete"
End Sub

Public Function ConfirmRemoval() As Boolean
    Dim msg As String
    msg = "  Remove TextBalance from this document?" & vbCrLf & vbCrLf & _
          "This will permanently delete:" & vbCrLf & _
          "   - Character count summary table" & vbCrLf & _
          "   - All heading annotations and progress bars" & vbCrLf & _
          "   - Stored settings and preferences" & vbCrLf & _
          "   - Hidden text from headings" & vbCrLf & vbCrLf & _
          "This action cannot be undone!" & vbCrLf & vbCrLf & _
          "The TextBalance macro will remain available for other documents."
          
    ConfirmRemoval = (MsgBox(msg, vbYesNo + vbExclamation, "Confirm Complete Removal") = vbYes)
End Function

Public Sub ShowRemovalSuccess()
    MsgBox ChrW(182) & "  TextBalance removed successfully" & vbCrLf & vbCrLf & _
           "    All data and formatting cleared" & vbCrLf & _
           "    Document restored to original state" & vbCrLf & _
           "    TextBalance add-in remains available" & vbCrLf & vbCrLf & _
           "Use Refresh to reload add-in into document.", _
           vbInformation, "Removal Complete"
End Sub

Public Sub ShowAnnotationRemovalSuccess()
    MsgBox ChrW(182) & " Annotations removed successfully!" & vbCrLf & vbCrLf & _
           "  " & ChrW(149) & "  All heading indicators cleared" & vbCrLf & _
           "  " & ChrW(149) & "  Progress bars removed" & vbCrLf & _
           "  " & ChrW(149) & "  Summary table preserved" & vbCrLf & vbCrLf & _
           "Use Refresh to restore annotations if needed.", _
           vbInformation, "Annotations Removed"
End Sub

Public Sub ShowTableRemovalSuccess()
    MsgBox ChrW(182) & " Summary table removed successfully!" & vbCrLf & vbCrLf & _
           "  •  Character count table deleted" & vbCrLf & _
           "  •  Heading annotations preserved" & vbCrLf & vbCrLf & _
           "Use Refresh to recreate the table if needed.", _
           vbInformation, "Table Removed"
End Sub

' =============================================================================
' Tutorial Functions
' =============================================================================

Public Function ShowInstallTutorial() As Boolean
    ' Called during installation - can be skipped
    Dim result As VbMsgBoxResult
    result = MsgBox(ChrW(182) & " Would you like a quick tutorial on how to use TextBalance?" & vbCrLf & vbCrLf & _
                   "This will show you:" & vbCrLf & _
                   "  - How headings work with TextBalance" & vbCrLf & _
                   "  - What the colors and numbers mean" & vbCrLf & _
                   "  - Tips for getting the best results" & vbCrLf & vbCrLf & _
                   "You can always access help later from the Help group in the ribbon.", _
                   vbYesNo + vbQuestion, "TextBalance Tutorial")
    
    If result = vbYes Then
        ShowHTMLHelp
        ConfigManager.SetTutorialShown
        ShowInstallTutorial = True
    Else
        ConfigManager.SetTutorialShown
        ShowInstallTutorial = False
    End If
End Function
Public Sub ShowHTMLHelp()
    Dim url As String
    url = "https://textbalance.github.io/tutorial"

    ' Cross-platform browser opening
    On Error GoTo ErrorHandler
    Dim shellCmd As String
    
    ' Platform detection and appropriate command
    If InStr(1, Application.System.OperatingSystem, "Windows", vbTextCompare) > 0 Then
        ' Windows
        shellCmd = "cmd /c start """" """ & url & """"
        Shell shellCmd, vbHide
    ElseIf InStr(1, Application.System.OperatingSystem, "Mac", vbTextCompare) > 0 Then
        ' macOS
        shellCmd = "open """ & url & """"
        Shell shellCmd, vbHide
    Else
        ' Linux/Unix (assumption)
        shellCmd = "xdg-open """ & url & """"
        Shell shellCmd, vbHide
    End If
    
    Exit Sub

ErrorHandler:
    MsgBox "Failed to open browser. You can visit the tutorial at: " & url, vbExclamation
End Sub
