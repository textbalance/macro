Attribute VB_Name = "ErrorManager"

' =============================================================================
' ErrorManager Module - Unified Error Handling
' =============================================================================

Public Enum ErrorSeverity
    esInfo = 1
    esWarning = 2
    esError = 3
    esCritical = 4
End Enum

Public Enum ErrorCategory
    ecGeneral = 1
    ecDocument = 2
    ecRange = 3
    ecTable = 4
    ecHeading = 5
    ecAnnotation = 6
    ecConfiguration = 7
End Enum

Private m_errorLog As Collection
Private m_debugMode As Boolean

Private Function CreateErrorInfo(source As String, description As String, severity As ErrorSeverity, category As ErrorCategory, context As String) As Object
    Dim errorInfo As Object
    Set errorInfo = CreateObject("Scripting.Dictionary")
    
    errorInfo("Number") = Err.Number
    errorInfo("Description") = description
    errorInfo("Source") = source
    errorInfo("Severity") = severity
    errorInfo("Category") = category
    errorInfo("Timestamp") = Now
    errorInfo("Context") = context
    
    Set CreateErrorInfo = errorInfo
End Function

Function InitializeErrorManager()
    Set m_errorLog = New Collection
    m_debugMode = True ' True during development
End Function

' === MAIN PUBLIC FUNCTIONS ===

Public Function SafeExecute(functionName As String, Optional context As String = "") As Boolean
    ' Safe function call wrapper
    On Error GoTo ErrorHandler
    
    SafeExecute = True
    Exit Function
    
ErrorHandler:
    Call HandleError(functionName, Err.description, esError, ecGeneral, context)
    SafeExecute = False
End Function

Public Sub HandleError(source As String, description As String, _
                      Optional severity As ErrorSeverity = esError, _
                      Optional category As ErrorCategory = ecGeneral, _
                      Optional context As String = "")
    
    Dim errorInfo As Object
    Set errorInfo = CreateErrorInfo(source, description, severity, category, context)
    
    ' Log errors
    LogError errorInfo
    
    ' User notification
    If severity = esCritical Then
        ShowErrorToUser errorInfo
    End If
    
    
    ' Debug information
    If m_debugMode Then
        Debug.Print FormatErrorForDebug(errorInfo)
    End If
End Sub

Public Function ValidateRange(rng As Range, functionName As String) As Boolean
    ' Range validation
    ValidateRange = False
    
    If rng Is Nothing Then
        Call HandleError(functionName, "Range object is null", esError, ecRange)
        Exit Function
    End If
    
    If rng.Start < 0 Or rng.End > ActiveDocument.Range.End Then
        Call HandleError(functionName, "Range boundaries are invalid", esError, ecRange)
        Exit Function
    End If
    
    If rng.Start >= rng.End Then
        Call HandleError(functionName, "Range start >= end", esWarning, ecRange)
        ' Warning, but not critical
    End If
    
    ValidateRange = True
End Function

Public Function ValidateDocument(functionName As String) As Boolean
    ' Document validation
    ValidateDocument = False
    
    If ActiveDocument Is Nothing Then
        Call HandleError(functionName, "No active document", esCritical, ecDocument)
        Exit Function
    End If
    
    If ActiveDocument.ProtectionType <> wdNoProtection Then
        Call HandleError(functionName, "Document is protected", esCritical, ecDocument)
        Exit Function
    End If
    
    If Not HasHeadingsInDocument() Then
        Call HandleError(functionName, "No headings found in document." & vbCrLf & vbCrLf & "Please add Heading 1 or Heading 2 styles, or create headings based on these styles.", esCritical, ecDocument)
        Exit Function
    End If
    
    ValidateDocument = True
End Function

Private Function HasHeadingsInDocument() As Boolean
    Dim rng As Range
    Set rng = ActiveDocument.Range
    
    ' Level 1 Search
    With rng.Find
        .ClearFormatting
        .ParagraphFormat.OutlineLevel = wdOutlineLevel1
        .text = ""
        .Forward = True
        If .Execute Then
            HasHeadingsInDocument = True
            Exit Function
        End If
    End With
    
    ' Level 2 Search
    Set rng = ActiveDocument.Range
    With rng.Find
        .ClearFormatting
        .ParagraphFormat.OutlineLevel = wdOutlineLevel2
        .text = ""
        .Forward = True
        If .Execute Then
            HasHeadingsInDocument = True
            Exit Function
        End If
    End With
    
    HasHeadingsInDocument = False
End Function

Public Function ValidateTable(tbl As table, functionName As String) As Boolean
    ' Table validation
    ValidateTable = False
    
    If tbl Is Nothing Then
        Call HandleError(functionName, "Table object is null", esError, ecTable)
        Exit Function
    End If
    
    If tbl.Rows.count < 3 Or tbl.Columns.count < 4 Then
        Call HandleError(functionName, "Table structure is invalid", esError, ecTable, _
                        "Rows: " & tbl.Rows.count & ", Columns: " & tbl.Columns.count)
        Exit Function
    End If
    
    ValidateTable = True
End Function

Public Sub ClearErrorLog()
    Set m_errorLog = New Collection
End Sub

Public Function GetErrorCount(Optional severity As ErrorSeverity = 0) As Long
    If severity = 0 Then
        GetErrorCount = m_errorLog.count
    Else
        Dim count As Long
        Dim i As Long
        For i = 1 To m_errorLog.count
            Dim errorInfo As Variant
            errorInfo = m_errorLog(i)
            If errorInfo("Severity") = severity Then count = count + 1
        Next i
        GetErrorCount = count
    End If
End Function

' === PRIVATE HELPER FUNCTIONS ===

Private Sub LogError(errorInfo As Object)
    ' Log error
    If m_errorLog Is Nothing Then Set m_errorLog = New Collection
    
    If m_errorLog.count > 100 Then
        ' Remove old errors (memory saving)
        m_errorLog.Remove 1
    End If
    
    'm_errorLog.Add errorInfo
End Sub

Private Sub ShowErrorToUser(errorInfo As Object)
    Dim msg As String
    Dim icon As VbMsgBoxStyle
    
    Select Case errorInfo("Severity")
        Case esError
            msg = "Error: " & errorInfo("Description")
            icon = vbCritical
        Case esCritical
            msg = "Error: " & errorInfo("Description")
            icon = vbCritical
    End Select
    
    If Len(errorInfo("Context")) > 0 Then
        msg = msg & vbCrLf & "Context: " & errorInfo("Context")
    End If
    
    MsgBox msg, icon, "TextBalance - " & errorInfo("Source")
End Sub

Private Function FormatErrorForDebug(errorInfo As Object) As String
    FormatErrorForDebug = "[" & Format(errorInfo("Timestamp"), "hh:mm:ss") & "] " & _
                         GetSeverityText(errorInfo("Severity")) & " " & _
                         errorInfo("Source") & ": " & errorInfo("Description")
    
    If Len(errorInfo("Context")) > 0 Then
        FormatErrorForDebug = FormatErrorForDebug & " [" & errorInfo("Context") & "]"
    End If
End Function

Private Function GetSeverityText(severity As ErrorSeverity) As String
    Select Case severity
        Case esInfo: GetSeverityText = "INFO"
        Case esWarning: GetSeverityText = "WARN"
        Case esError: GetSeverityText = "ERROR"
        Case esCritical: GetSeverityText = "CRITICAL"
        Case Else: GetSeverityText = "UNKNOWN"
    End Select
End Function

Public Property Get DebugMode() As Boolean
    DebugMode = m_debugMode
End Property

Public Property Let DebugMode(value As Boolean)
    m_debugMode = value
End Property
