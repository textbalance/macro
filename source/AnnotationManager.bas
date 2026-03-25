Attribute VB_Name = "AnnotationManager"
' =============================================================================
' AnnotationManager Module
' ============================================================================


' ==================================
' Annotation Removal
' ==================================

Public Sub RemoveAnnotationCharacters()
    Dim rng As Range
    Set rng = ActiveDocument.Range
    
    With rng.Find
        .text = ChrW(ANNOTATION_START)
        .Forward = True
        .MatchCase = True
        
        Do While .Execute
            ' CHECK: Are we in a table?
            Dim inTable As Boolean
            inTable = (rng.Information(wdWithInTable) = True)
            
            If inTable Then
                ' Identify table cell
                Dim cell As cell
                Set cell = rng.Cells(1)
            End If
            
            Dim inTOC As Boolean
            inTOC = IsInsideTOCField(rng)
            
            ' Search for paragraph
            Dim Para As Paragraph
            Set Para = rng.Paragraphs(1)
            
            ' Deletion range
            Dim deleteRange As Range
            Set deleteRange = ActiveDocument.Range(rng.Start - 1, Para.Range.End - 1)
               
            ' ONLY DELETE IF NOT IN A TABLE
            If Not inTable And Not inTOC And rng.Start < Para.Range.End - 1 Then
                deleteRange.Delete
            End If
        Loop
    End With
End Sub

Private Function IsInsideTOCField(rng As Range) As Boolean
    On Error Resume Next
    Dim fld As Field
    For Each fld In rng.Paragraphs(1).Range.Fields
        If fld.Type = wdFieldTOC Or fld.Type = wdFieldTOCEntry Then
            IsInsideTOCField = True
            Exit Function
        End If
    Next fld
    
    ' Fallback: check if any parent field is TOC
    Dim paraRange As Range
    Set paraRange = rng.Paragraphs(1).Range
    For Each fld In ActiveDocument.Fields
        If fld.Type = wdFieldTOC Then
            If paraRange.Start >= fld.result.Start And paraRange.End <= fld.result.End Then
                IsInsideTOCField = True
                Exit Function
            End If
        End If
    Next fld
    
    IsInsideTOCField = False
End Function

' =============================================================================
' Add annotations - main
' ============================================================================

Public Sub AddAnnotations(headings As Collection)
    Dim i As Long
    Dim heading As clsHeadingInfo
    
    For i = 1 To headings.count
        Set heading = headings(i)
        
        On Error Resume Next
        
        If heading.level = 1 Then
            AddLevel1Annotation heading, headings
        ElseIf heading.level = 2 Then
            ' Here we calculate sibling information
            Dim siblingInfo As clsSiblingInfo
            Set siblingInfo = HeadingProcessor.CalculateSiblingInfo(headings, i)
            AddLevel2Annotation heading, siblingInfo
        End If
        
        If Err.Number <> 0 Then
            ErrorManager.HandleError "Annotations", Err.description, esError, ecAnnotation
            Err.Clear
        End If
        
        On Error GoTo 0
    Next i
    
End Sub

' =============================================================================
' Level 1 Annotation - create and format
' ============================================================================

Private Sub AddLevel1Annotation(heading As clsHeadingInfo, headings As Collection)
    Dim limits As Object
    Set limits = GetTableLimits(heading, headings)
    Dim headingId As String
    headingId = heading.headingId

    Dim annotationText As String
    Dim displayText As String
    If ConfigManager.GetSetting("DisplaySpeechTime") = "True" Then
        Dim speechMinutes As Double
        Dim tempo As Integer
        tempo = Val(ConfigManager.GetSetting("SpeechTempo"))
        If tempo = 0 Then tempo = DEFAULT_SPEECH_TEMPO
        speechMinutes = heading.charCount / tempo
        displayText = Format(speechMinutes, "0.0") & " min"
    Else
        displayText = heading.charCount & " char"
    End If

annotationText = ChrW(ANNOTATION_START) & headingId & " (" & displayText & ", " & _
                Format(heading.percentage, "0.0") & "%)" & ChrW(ANNOTATION_END) & ChrW(ANNOTATION_END)
    
    Dim color As Long
    color = DetermineColor(heading.percentage, limits("Ideal"), limits("Tolerance"))
    
    AddFormattedAnnotation heading.Range, annotationText, color, headingId
End Sub

Public Function GetTableLimits(heading As clsHeadingInfo, headings As Collection) As Object
    Dim limits As Object
    Set limits = CreateObject("Scripting.Dictionary")

    ' Defaults
    limits("Ideal") = HeadingProcessor.GetDefaultIdealPercent()
    limits("Tolerance") = DEFAULT_TOLERANCE

    If heading.isExcluded Then
        limits("Ideal") = -1
    ElseIf heading.idealPercent > 0 Then
        limits("Ideal") = heading.idealPercent
    End If
    
    If Not heading.isExcluded And heading.limitPercent > 0 Then
        limits("Tolerance") = heading.limitPercent
    End If

    Set GetTableLimits = limits
End Function

Private Function DetermineColor(actualPercent As Double, idealPercent As Double, tolerance As Double) As Long
    ' Excluded headings get gray color
    If idealPercent = -1 Then
        DetermineColor = COLOR_DARK_GRAY
        Exit Function
    End If
    
    Dim difference As Double
    difference = Abs(actualPercent - idealPercent)
    
    If difference <= tolerance Then
        DetermineColor = COLOR_GREEN
    ElseIf difference <= (2 * tolerance) Then
        DetermineColor = COLOR_ORANGE
    Else
        DetermineColor = COLOR_RED
    End If
End Function

Public Sub AddFormattedAnnotation(headingRange As Range, annotationText As String, color As Long, headingId As String)
    Dim Para As Paragraph
    Set Para = headingRange.Paragraphs(1)
    
    Dim insertPos As Long
    insertPos = Para.Range.End - 1
    
    Dim insertRange As Range
    Set insertRange = ActiveDocument.Range(insertPos, insertPos)
    insertRange.text = " " & annotationText

    Dim IdRange As Range
    Set IdRange = ActiveDocument.Range(insertPos, insertPos + Len(headingId) + 2)

    Dim annotationRange As Range
    Set annotationRange = ActiveDocument.Range(insertPos + Len(headingId) + 2, insertPos + Len(annotationText))
    
    With IdRange
        With .Font
            .size = 1
            .color = RGB(255, 255, 255)
            .Hidden = True
        End With
    End With

    With annotationRange
        With .Font
            .size = DEFAULT_FONT_SIZE
            .color = color
            .Hidden = True
        End With
    End With

    End Sub

Public Sub AddLevel2Annotation(heading As clsHeadingInfo, siblingInfo As clsSiblingInfo)
    ' Simplified: just space + start character, then insert segments
    Dim Para As Paragraph
    Set Para = heading.Range.Paragraphs(1)
    
    Dim insertPos As Long
    insertPos = Para.Range.End - 1
    
    ' Insert start and end characters (to get the heading formatting)
    Dim insertRange As Range
    Set insertRange = ActiveDocument.Range(insertPos, insertPos)
    insertRange.text = " " & ChrW(ANNOTATION_START) & ChrW(ANNOTATION_END) & ChrW(ANNOTATION_END)
    
    Dim startRange As Range
    Set startRange = ActiveDocument.Range(insertPos, insertPos + 4) ' space + start character
    
    Dim startPos As Long
    startPos = insertPos + 2  ' after space + start character
          
    ' Insert segments and format immediately
    InsertAndFormatSegments startPos, siblingInfo
      
End Sub

Private Sub InsertAndFormatSegments(startPos As Long, siblingInfo As clsSiblingInfo)
    Dim currentPos As Long
    currentPos = startPos
    
    ' If no segments, insert empty bar
    If siblingInfo.GetBarSegmentCount() = 0 Then
        InsertEmptyBar currentPos, siblingInfo.barWidth
        Exit Sub
    End If
    
    ' Create consolidated segments by type
    Dim consolidatedSegments As Collection
    Set consolidatedSegments = ConsolidateAdjacentSegments(siblingInfo)
    
    Dim totalInserted As Integer
    totalInserted = 0
    
    ' Insert consolidated segments
    Dim i As Long
    For i = 1 To consolidatedSegments.count
        Dim segment As Variant
        segment = consolidatedSegments(i)
        
        Dim segmentType As String: segmentType = segment(0)
        Dim segmentWidth As Integer: segmentWidth = segment(1)
        
        If segmentWidth > 0 Then
            InsertAndFormatSegment currentPos, segmentWidth, segmentType
            currentPos = currentPos + segmentWidth
            totalInserted = totalInserted + segmentWidth
        End If
    Next i
    
    ' If fewer segments inserted than bar width, fill with empty segments
    If totalInserted < siblingInfo.barWidth Then
        InsertEmptyBar currentPos, siblingInfo.barWidth - totalInserted
    End If

End Sub

Private Function ConsolidateAdjacentSegments(siblingInfo As clsSiblingInfo) As Collection
    Dim consolidated As New Collection
    
    If siblingInfo.GetBarSegmentCount() = 0 Then
        Set ConsolidateAdjacentSegments = consolidated
        Exit Function
    End If
    
    Dim currentType As String
    Dim currentWidth As Integer
    Dim i As Long
    
    ' Initialize first segment
    Dim firstSegment As Variant
    firstSegment = siblingInfo.barSegments(1)
    Dim firstSegmentType As String
    firstSegmentType = CStr(firstSegment(3))
    currentType = DetermineConsolidationType(firstSegmentType)
    currentWidth = firstSegment(2)
    
    ' Process remaining segments
    For i = 2 To siblingInfo.barSegments.count
        Dim segment As Variant
        segment = siblingInfo.barSegments(i)
        
        Dim segmentTypeStr As String
        segmentTypeStr = CStr(segment(3))
        Dim segmentType As String
        segmentType = DetermineConsolidationType(segmentTypeStr)
        
        If segmentType = currentType Then
            ' Same type, add to width
            currentWidth = currentWidth + segment(2)
        Else
            ' New type, save current and start new
            consolidated.Add Array(currentType, currentWidth)
            currentType = segmentType
            currentWidth = segment(2)
        End If
    Next i
    
    ' Add last segment
    If currentWidth > 0 Then
        consolidated.Add Array(currentType, currentWidth)
    End If
    
    Set ConsolidateAdjacentSegments = consolidated
End Function

Private Function DetermineConsolidationType(originalType As String) As String
    Select Case originalType
        Case "current"
            DetermineConsolidationType = "current"
        Case "orphan", "sibling"
            DetermineConsolidationType = "sibling"
        Case Else
            DetermineConsolidationType = "sibling"
    End Select
End Function

Private Sub InsertAndFormatSegment(position As Long, width As Integer, segmentType As String)
    ' Select appropriate character
    Dim charCode As String
    Dim currentBarSet As String
    currentBarSet = ConfigManager.GetSetting("BarCharacterSet")
    If currentBarSet = "" Then currentBarSet = BAR_SET_STANDARD
    
    ' Select character based on set
    Select Case currentBarSet
        Case BAR_SET_STANDARD
            Select Case segmentType
                Case "current": charCode = STANDARD_CHAR_FILLED
                Case "sibling": charCode = STANDARD_CHAR_SIBLING
                Case Else: charCode = STANDARD_CHAR_EMPTY
            End Select
            
        Case BAR_SET_ASCII
            Select Case segmentType
                Case "current": charCode = ASCII_CHAR_FILLED
                Case "sibling": charCode = ASCII_CHAR_SIBLING
                Case Else: charCode = ASCII_CHAR_EMPTY
            End Select
            
        Case BAR_SET_BRAILLE
            Select Case segmentType
                Case "current": charCode = BRAILLE_CHAR_FILLED
                Case "sibling": charCode = BRAILLE_CHAR_SIBLING
                Case Else: charCode = BRAILLE_CHAR_EMPTY
            End Select
            
        Case BAR_SET_GEOMETRIC
            Select Case segmentType
                Case "current": charCode = GEOMETRIC_CHAR_FILLED
                Case "sibling": charCode = GEOMETRIC_CHAR_SIBLING
                Case Else: charCode = GEOMETRIC_CHAR_EMPTY
            End Select
            
        Case Else ' Fallback
            Select Case segmentType
                Case "current": charCode = CHAR_FILLED
                Case "sibling": charCode = CHAR_SIBLING
                Case Else: charCode = CHAR_EMPTY
            End Select
    End Select
    
    ' Assemble text
    Dim segmentText As String
    Dim i As Integer
    For i = 1 To width
        segmentText = segmentText & ChrW(charCode)
    Next i
    
    ' Insertion
    Dim insertRange As Range
    Set insertRange = ActiveDocument.Range(position, position)
    insertRange.text = segmentText
    
    ' Immediate formatting
    Dim formatRange As Range
    Set formatRange = ActiveDocument.Range(insertRange.Start, insertRange.End)
    
    Select Case segmentType
        Case "current"
            FormatRangeAsCurrent formatRange
        Case "sibling"
            FormatRangeAsSibling formatRange
        Case Else
            FormatRangeAsSibling formatRange
    End Select
End Sub

Private Sub InsertEmptyBar(position As Long, width As Integer)
    Dim emptyText As String
    Dim i As Integer
    For i = 1 To width
        emptyText = emptyText & ChrW(CHAR_EMPTY)
    Next i
    
    Dim insertRange As Range
    Set insertRange = ActiveDocument.Range(position, position)
    insertRange.text = emptyText
    
    Dim formatRange As Range
   
    Set formatRange = ActiveDocument.Range(position, position + width)
    FormatRangeAsEmpty formatRange
End Sub

Private Sub FormatRangeAsCurrent(rng As Range)
    With rng.Font
        .color = COLOR_BLACK
        .Bold = True
        .Scaling = DEFAULT_SCALING
        .size = DEFAULT_FONT_SIZE
        .Hidden = True
    End With
    
End Sub

Private Sub FormatRangeAsSibling(rng As Range)
    With rng.Font
        .color = COLOR_DARK_GRAY
        .Bold = True
        .Scaling = DEFAULT_SCALING
        .size = DEFAULT_FONT_SIZE
        .Hidden = True
    End With
End Sub

Private Sub FormatRangeAsEmpty(rng As Range)
    With rng.Font
        .color = COLOR_LIGHT_GRAY
        .Bold = False
        .Scaling = DEFAULT_SCALING
        .size = DEFAULT_FONT_SIZE
        .Hidden = True
    End With

End Sub
