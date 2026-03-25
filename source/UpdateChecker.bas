
' =============================================================================
' UpdateChecker Module - Version checking and update notification
' =============================================================================

Public Sub CheckForUpdates(Optional showUpToDate As Boolean = False)
    On Error GoTo ErrorHandler
    
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    http.Open "GET", VERSION_CHECK_URL, False
    http.send
    
    If http.Status <> 200 Then
        If showUpToDate Then
            MsgBox "Could not check for updates. HTTP status: " & http.Status & vbCrLf & _
                   "Please check your internet connection.", _
                   vbExclamation, "Update Check"
        End If
        Exit Sub
    End If

    ' Parse response
    Dim responseText As String
    responseText = http.responseText
    
    Dim latestVersion As String
    latestVersion = ExtractJsonValue(responseText, "version")
    
    Dim downloadUrl As String
    downloadUrl = ExtractJsonValue(responseText, "downloadUrl")
    
    Dim releaseNotes As String
    releaseNotes = ExtractJsonValue(responseText, "releaseNotes")
    
    ' Compare versions
    If IsNewerVersion(latestVersion, APP_VERSION) Then
        Dim msg As String
        msg = "A new version of TextBalance is available!" & vbCrLf & vbCrLf & _
              "Current version: " & APP_VERSION & vbCrLf & _
              "Latest version: " & latestVersion & vbCrLf
        
        If Len(releaseNotes) > 0 Then
            msg = msg & vbCrLf & "What's new: " & releaseNotes & vbCrLf
        End If
        
        msg = msg & vbCrLf & "Would you like to open the download page?"
        
        If MsgBox(msg, vbYesNo + vbInformation, "TextBalance Update") = vbYes Then
            If Len(downloadUrl) > 0 Then
                ActiveDocument.FollowHyperlink downloadUrl
            Else
                ActiveDocument.FollowHyperlink "https://github.com/textbalance/macro/releases/latest"
            End If
        End If
    ElseIf showUpToDate Then
        MsgBox "TextBalance is up to date!" & vbCrLf & vbCrLf & _
               "Current version: " & APP_VERSION, _
               vbInformation, "Update Check"
    End If
    
    Exit Sub

ErrorHandler:
    If showUpToDate Then
        MsgBox "Could not check for updates." & vbCrLf & _
               "Error: " & Err.description, _
               vbExclamation, "Update Check"
    End If
End Sub

Private Function IsNewerVersion(remote As String, current As String) As Boolean

    Dim remoteParts() As String
    Dim localParts() As String
    
    remoteParts = Split(remote, ".")
    localParts = Split(current, ".")
    
    Dim maxParts As Long
    If UBound(remoteParts) > UBound(localParts) Then
        maxParts = UBound(remoteParts)
    Else
        maxParts = UBound(localParts)
    End If
    
    Dim i As Long
    For i = 0 To maxParts
        Dim r As Long, l As Long
        If i <= UBound(remoteParts) Then r = Val(remoteParts(i)) Else r = 0
        If i <= UBound(localParts) Then l = Val(localParts(i)) Else l = 0
        
        If r > l Then
            IsNewerVersion = True
            Exit Function
        ElseIf r < l Then
            IsNewerVersion = False
            Exit Function
        End If
    Next i
    
    IsNewerVersion = False
End Function

Private Function ExtractJsonValue(json As String, key As String) As String
    ' Simple JSON value extractor (no external dependency)
    Dim searchKey As String
    searchKey = """" & key & """"
    
    Dim keyPos As Long
    keyPos = InStr(json, searchKey)
    If keyPos = 0 Then
        ExtractJsonValue = ""
        Exit Function
    End If
    
    ' Find colon after key
    Dim colonPos As Long
    colonPos = InStr(keyPos + Len(searchKey), json, ":")
    If colonPos = 0 Then
        ExtractJsonValue = ""
        Exit Function
    End If
    
    ' Find opening quote of value
    Dim valueStart As Long
    valueStart = InStr(colonPos, json, """")
    If valueStart = 0 Then
        ExtractJsonValue = ""
        Exit Function
    End If
    
    ' Find closing quote
    Dim valueEnd As Long
    valueEnd = InStr(valueStart + 1, json, """")
    If valueEnd = 0 Then
        ExtractJsonValue = ""
        Exit Function
    End If
    
    ExtractJsonValue = Mid(json, valueStart + 1, valueEnd - valueStart - 1)
End Function
