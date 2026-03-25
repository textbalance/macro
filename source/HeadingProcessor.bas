Attribute VB_Name = "HeadingProcessor"

' =============================================================================
' HeadingProcessor Module - Processing headings
' =============================================================================

Private Type ProcessingCache
    totalChars As Long
    level1Count As Long
    defaultIdealPercent As Double
    headingCollection As Collection
    isValid As Boolean
End Type

Private m_cache As ProcessingCache

' === IMPORTANT PUBLIC FUNCTIONS ===

Public Function ProcessAllHeadingsWithTable(Optional tbl As table = Nothing) As Collection
    InitializeCache
    
    Dim headings As Collection
    Set headings = CollectAllHeadings()
    
    EstablishHierarchy headings
      
    If Not tbl Is Nothing Then
        LoadTableDataIntoHeadings headings, tbl
    End If
    
    Set m_cache.headingCollection = headings
    
    Set ProcessAllHeadingsWithTable = headings
End Function

Public Function GetCachedHeadings() As Collection
    ' Quick access to cached data
    If m_cache.isValid And Not m_cache.headingCollection Is Nothing Then
        Set GetCachedHeadings = m_cache.headingCollection
    Else
        Set GetCachedHeadings = ProcessAllHeadingsWithTable()
    End If
End Function

Public Function GetScaledIdealPercent(originalIdeal As Double, totalSum As Double) As Double
    If totalSum > 100 Then
        GetScaledIdealPercent = originalIdeal * (100 / totalSum)
    Else
        GetScaledIdealPercent = originalIdeal
    End If
End Function

Public Sub CalculateBarSegments(siblingInfo As clsSiblingInfo, Optional barWidth As Integer = 40)
    
    ' Set bar width
    siblingInfo.barWidth = barWidth
    
    ' If no siblings, exit
    If siblingInfo.GetSiblingCount() = 0 Then
        siblingInfo.totalFilledBarWidth = 0
        Exit Sub
    End If
    
    ' NEW LOGIC: Determine target character count
    Dim targetCharCount As Long
    ' Use ParentIdealChars as set in CalculateSiblingInfo
    targetCharCount = siblingInfo.ParentIdealChars
    
    ' Calculate total fill width
    Dim totalFilledWidth As Integer
    If targetCharCount > 0 Then
        totalFilledWidth = CInt((siblingInfo.totalSiblingChars / targetCharCount) * barWidth)
        If totalFilledWidth > barWidth Then totalFilledWidth = barWidth
        If totalFilledWidth < 1 And siblingInfo.totalSiblingChars > 0 Then totalFilledWidth = 1
    Else
        totalFilledWidth = barWidth ' Fallback
    End If
    
    siblingInfo.totalFilledBarWidth = totalFilledWidth
    
    ' Calculate segment widths
    Dim i As Long
    Dim currentBarPos As Integer: currentBarPos = 1
    Dim currentSiblingFound As Boolean: currentSiblingFound = False
    
    For i = 1 To siblingInfo.siblings.count
        Dim siblingData As Variant
        siblingData = siblingInfo.siblings(i)
        
        ' Calculate segment width proportionally
        Dim segmentWidth As Integer
        If siblingInfo.totalSiblingChars > 0 Then
            segmentWidth = CInt((siblingData(2) / siblingInfo.totalSiblingChars) * totalFilledWidth)
            If segmentWidth < 1 And siblingData(2) > 0 Then segmentWidth = 1
        Else
            segmentWidth = 0
        End If
        
        ' Determine segment type
        Dim segmentType As String
        If siblingData(0) = siblingInfo.currentSiblingIndex Then
            segmentType = "current"
            siblingInfo.currentSiblingBarStart = currentBarPos
            siblingInfo.currentSiblingBarWidth = segmentWidth
            currentSiblingFound = True
        ElseIf siblingData(0) = -1 Then
            segmentType = "orphan"
        Else
            segmentType = "sibling"
        End If
        
        ' Add segment
        Dim siblingId As Long
        siblingId = CLng(siblingData(0))
        siblingInfo.AddBarSegment siblingId, currentBarPos, segmentWidth, segmentType
        
        currentBarPos = currentBarPos + segmentWidth
    Next i
    
    ' Check
    If Not currentSiblingFound Then
        siblingInfo.currentSiblingBarStart = 1
        siblingInfo.currentSiblingBarWidth = 1
    End If
End Sub

' === COLLECT HEADINGS AND COUNT CHARACTERS ===

Private Function CollectAllHeadings() As Collection
    
    Dim headings As New Collection
    Dim rng As Range
    Dim maxId As Long: maxId = 0
    Dim i As Long
    Dim headingInfo As clsHeadingInfo
    
    ' Search for HEADING 1 (wdOutlineLevel1)
    Set rng = ActiveDocument.Range
    With rng.Find
        .ClearFormatting
        .ParagraphFormat.OutlineLevel = wdOutlineLevel1
        .text = ""
        .Forward = True
        
        Do While .Execute
            i = i + 1
            Set headingInfo = New clsHeadingInfo
            With headingInfo
                Set .Range = rng.Duplicate  ' Important: copy!
                .level = 1
                .charCount = 0
                .percentage = 0
                .parentIndex = -1
                
                ' ID handling
                Dim existingId As String
                existingId = ExtractIdFromParagraph(rng.Paragraphs(1))
                If Len(existingId) > 0 Then
                    .headingId = existingId
                    Dim currentId As Long
                    currentId = Val(Right(existingId, 3))
                    If currentId > maxId Then maxId = currentId
                Else
                    maxId = maxId + 1
                    .headingId = "ID" & Format(maxId, "000")
                End If
                
                .cleanText = CleanHeadingText(.Range)
            End With
            
            headings.Add headingInfo
            Dim nextPos As Long
            nextPos = rng.Paragraphs(1).Range.End
            If nextPos >= ActiveDocument.Range.End Then Exit Do
            
            rng.SetRange nextPos, ActiveDocument.Range.End
        Loop
    End With
    
    ' Search for HEADING 2 (wdOutlineLevel2)
    Set rng = ActiveDocument.Range
    With rng.Find
        .ClearFormatting
        .ParagraphFormat.OutlineLevel = wdOutlineLevel2
        .text = ""
        .Forward = True
        
        Do While .Execute
            i = i + 1
            Set headingInfo = New clsHeadingInfo
            With headingInfo
                Set .Range = rng.Duplicate
                .level = 2
                .charCount = 0
                .percentage = 0
                .parentIndex = -1
                .cleanText = CleanHeadingText(.Range)
            End With
            
            headings.Add headingInfo
            nextPos = rng.Paragraphs(1).Range.End
            If nextPos >= ActiveDocument.Range.End Then Exit Do
            
            rng.SetRange nextPos, ActiveDocument.Range.End
        Loop
    End With
    
    SortHeadingsByPosition headings
    ' Character count calculation
    CalculateCharacterCountsBatch headings
    Set CollectAllHeadings = headings
End Function

Private Sub SortHeadingsByPosition(headings As Collection)
    ' Create temporary list for sorting
    Dim sortedList As Collection
    Set sortedList = New Collection
    
    ' Insert every item to the appropriate position
    Dim i As Long, j As Long
    For i = 1 To headings.count
        Dim currentHeading As clsHeadingInfo
        Set currentHeading = headings(i)
        
        ' Where to insert?
        Dim insertPos As Long
        insertPos = sortedList.count + 1  ' Default: at the end
        
        For j = 1 To sortedList.count
            Dim existingHeading As clsHeadingInfo
            Set existingHeading = sortedList(j)
            
            If currentHeading.Range.Start < existingHeading.Range.Start Then
                insertPos = j
                Exit For
            End If
        Next j
        
        ' Insertion to the appropriate position
        If insertPos > sortedList.count Then
            sortedList.Add currentHeading
        Else
            sortedList.Add currentHeading, , insertPos
        End If
    Next i
    
    ' Empty and reload original collection
    Set headings = sortedList
End Sub

Private Sub CalculateCharacterCountsBatch(headings As Collection)
    Dim i As Long
    Dim heading As clsHeadingInfo
    Dim startPos As Long, endPos As Long
    Dim totalCalculatedChars As Long: totalCalculatedChars = 0
    
    For i = 1 To headings.count
        Set heading = headings(i)
        startPos = heading.Range.End
                
        ' End position determination
        If heading.level = 1 Then
            endPos = FindNextLevel1PositionFast(headings, i)
            If endPos = -1 Then
                endPos = ActiveDocument.Range.End
            End If
        Else
            endPos = FindNextAnyHeadingPositionSafe(headings, i)
        End If
        
        ' Range check
        If endPos <= startPos Then
            heading.charCount = 0
        ElseIf startPos >= ActiveDocument.Range.End Then
            heading.charCount = 0
        Else
            ' Safe end position
            If endPos > ActiveDocument.Range.End Then endPos = ActiveDocument.Range.End
            
            ' Raw text
            Dim contentRange As Range
            Set contentRange = ActiveDocument.Range(startPos, endPos)
            Dim rawText As String
            rawText = contentRange.text
            
            ' Number of visible characters
            If i = headings.count Then
                Dim cleanText As String
                cleanText = Replace(rawText, vbCr, "")
                heading.charCount = Len(cleanText)
            Else
                heading.charCount = CountVisibleCharacters(contentRange)
            End If
        End If
        
        ' Percentage calculation
        If m_cache.totalChars > 0 Then
            heading.percentage = (heading.charCount / m_cache.totalChars) * 100#
        Else
            heading.percentage = 0
        End If
        
        totalCalculatedChars = totalCalculatedChars + heading.charCount
    Next i
    
End Sub
Private Function CountVisibleCharacters(rng As Range) As Long
    On Error Resume Next
    
    Dim cleanText As String
    cleanText = Replace(rng.text, vbCr, "")
    
    ' Quick path: no annotation
    If InStr(cleanText, ChrW(ANNOTATION_START)) = 0 Then
        CountVisibleCharacters = Len(cleanText)
        Exit Function
    End If

    ' === STABLE SOLUTION: String operation ===
    Dim visibleText As String
    visibleText = RemoveAnnotationsFromText(cleanText)

    CountVisibleCharacters = Len(visibleText)

    If Err.Number <> 0 Then
        ' Fallback: original length
        CountVisibleCharacters = Len(cleanText)
        Err.Clear
    End If
End Function

Private Function RemoveAnnotationsFromText(text As String) As String
    Dim result As String
    Dim marker As String
    Dim startPos As Long
    Dim endPos As Long
    Dim nextPos As Long
    
    marker = ChrW(ANNOTATION_START) ' START and END are same
    result = text
    
    Dim i As Long
    For i = 1 To 50
        startPos = InStr(result, marker)
        if startPos = 0 Then Exit For
        
        ' The first marker is opener - finding NEXT marker (closer)
        endPos = InStr(startPos + 1, result, marker)
        if endPos = 0 Then
            result = Left(result, startPos - 1)
            Exit For
        End If
        
        ' In case of double closing marker (Level 2 annotation), skip ahead
        Do While endPos + 1 <= Len(result) And Mid(result, endPos + 1, 1) = marker
            endPos = endPos + 1
        Loop
        
        result = Left(result, startPos - 1) & Mid(result, endPos + 1)
    Next i
    
    RemoveAnnotationsFromText = result
End Function

Private Function FindNextLevel1PositionFast(headings As Collection, currentIndex As Long) As Long
    ' Try from collection first (faster)
    Dim i As Long
    Dim heading As clsHeadingInfo
    
    For i = currentIndex + 1 To headings.count
        Set heading = headings(i)
        If heading.level = 1 Then
            FindNextLevel1PositionFast = heading.Range.Start
            Exit Function
        End If
    Next i
    
    ' If not in collection, end of document (safe)
    FindNextLevel1PositionFast = ActiveDocument.Range.End
End Function
Private Function FindNextAnyHeadingPositionSafe(headings As Collection, currentIndex As Long) As Long
    If currentIndex < headings.count Then
        Dim nextHeading As clsHeadingInfo
        Set nextHeading = headings(currentIndex + 1)
        FindNextAnyHeadingPositionSafe = nextHeading.Range.Start
    Else
        ' Safe end of document
        FindNextAnyHeadingPositionSafe = ActiveDocument.Range.End
    End If
End Function

Private Sub CalculateOrphanTextBatch(headings As Collection)
    Dim i As Long
    Dim heading As clsHeadingInfo
    
    For i = 1 To headings.count
        Set heading = headings(i)
        
        If heading.level = 1 And heading.hasChildren Then
            ' Find first child position
            Dim firstChildPos As Long
            firstChildPos = FindFirstChildPosition(headings, i)
            
            If firstChildPos > 0 Then
                Dim orphanRange As Range
                Set orphanRange = ActiveDocument.Range(heading.Range.End, firstChildPos)
                heading.orphanTextSize = Len(orphanRange.text)
            Else
                heading.orphanTextSize = 0
            End If
        End If
    Next i
End Sub

Private Function FindFirstChildPosition(headings As Collection, parentIndex As Long) As Long
    Dim parentHeading As clsHeadingInfo
    Set parentHeading = headings(parentIndex)
    
    If parentHeading.childrenCount > 0 Then
        Dim firstChildIndex As Long
        firstChildIndex = parentHeading.childrenIndices(1)
        Dim firstChild As clsHeadingInfo
        Set firstChild = headings(firstChildIndex)
        FindFirstChildPosition = firstChild.Range.Start
    Else
        FindFirstChildPosition = -1
    End If
End Function

' === HIERARCHY ===

Private Sub EstablishHierarchy(headings As Collection)
    Dim i As Long, j As Long
    Dim currentHeading As clsHeadingInfo
    Dim parentHeading As clsHeadingInfo

    For i = 1 To headings.count
        Set currentHeading = headings(i)
        
        If currentHeading.level = 2 Then
            ' finding parent backwards
            For j = i - 1 To 1 Step -1
                Set parentHeading = headings(j)
                If parentHeading.level = 1 Then
                    currentHeading.parentIndex = j
                    ' add child index to heading of parent
                    parentHeading.AddChildIndex i
                    Exit For
                End If
            Next j
        End If
    Next i

    ' Calculate orphan text for level 1 headings
    CalculateOrphanTextBatch headings
End Sub

' === SAVING EXISTING TABLE DATA INTO HEADING OBJECT ===

Private Sub LoadTableDataIntoHeadings(headings As Collection, tbl As table)
    ' Loading table data into heading objects with SCALING
    Dim i As Long, j As Long
    Dim heading As clsHeadingInfo
    Dim cellText As String
    Dim headingId As String

    ' FIRST PASS: Loading original values
    For i = 3 To tbl.Rows.count - 1 ' Skip last row (summary)
        headingId = TableManager.ExtractHeadingId(CleanCellText(tbl.cell(i, 1).Range.text))
        If Len(headingId) > 0 Then
            ' Searching for appropriate heading
            For j = 1 To headings.count
                Set heading = headings(j)
                If heading.level = 1 And heading.headingId = headingId Then
                    ' Loading Ideal %
                    cellText = CleanCellText(tbl.cell(i, 3).Range.text)
                    If cellText = "-" Or cellText = "EXCL" Or cellText = "N/A" Then
                        heading.idealPercent = -1
                    ElseIf IsNumeric(Replace(cellText, "%", "")) Then
                        heading.idealPercent = Val(Replace(cellText, "%", ""))
                    End If
                    
                    ' Loading Limit %
                    cellText = CleanCellText(tbl.cell(i, 4).Range.text)
                    If cellText = "-" Or cellText = "EXCL" Or cellText = "N/A" Then
                        heading.limitPercent = -1
                    ElseIf IsNumeric(Replace(cellText, "%", "")) Then
                        heading.limitPercent = Val(Replace(cellText, "%", ""))
                    End If
                    
                    Exit For
                End If
            Next j
        End If
    Next i
    
    ' SECOND PASS: Now we can calculate the actual totalIdeal
    Dim totalIdeal As Double
    totalIdeal = CalculateTableIdealSum(headings)
    
    ' THIRD PASS: Applying scaling if necessary
    If totalIdeal > 100 Then
        For i = 1 To headings.count
            Set heading = headings(i)
            If heading.level = 1 And heading.hasIdealPercent And Not heading.isExcluded Then
                Dim originalIdeal As Double
                originalIdeal = heading.idealPercent
                heading.idealPercent = GetScaledIdealPercent(heading.idealPercent, totalIdeal)
            End If
        Next i
    End If
End Sub

'==== IMPORTANT PERCENTAGE CALCULATIONS ====

Public Function GetDefaultIdealPercent() As Double
    ' Collecting headings
    Dim headings As Collection
    Set headings = GetCachedHeadings()
    
    ' Table ideal sum
    Dim tableIdealSum As Double
    tableIdealSum = CalculateTableIdealSum(headings)
    
    ' Proportional scaling if necessary
    If tableIdealSum > 100 Then
        tableIdealSum = 100 ' Max 100 after scaling
    End If
    
    ' Remainder calculation
    Dim remainingPercent As Double
    remainingPercent = 100 - tableIdealSum
    
    ' How many Level1 headings have no ideal value?
    Dim headingsWithoutIdeal As Long
    headingsWithoutIdeal = CountLevel1HeadingsWithoutIdeal(headings)
    
    If headingsWithoutIdeal > 0 And remainingPercent > 0 Then
        GetDefaultIdealPercent = remainingPercent / headingsWithoutIdeal
    ElseIf headingsWithoutIdeal > 0 Then
        GetDefaultIdealPercent = 0 ' No remainder
    Else
        GetDefaultIdealPercent = 0 ' All headings have ideal or are excluded
    End If
End Function

Public Function CalculateTableIdealSum(headings As Collection) As Double
    Dim totalIdeal As Double
    totalIdeal = 0
    
    Dim i As Long
    Dim heading As clsHeadingInfo
    
    For i = 1 To headings.count
        Set heading = headings(i)
        If heading.level = 1 And heading.hasIdealPercent And Not heading.isExcluded Then
            totalIdeal = totalIdeal + heading.idealPercent
        End If
    Next i
    
    CalculateTableIdealSum = totalIdeal
End Function

' === SIBLING INFO ===

Public Function CalculateSiblingInfo(headings As Collection, currentIndex As Long) As clsSiblingInfo
    Dim siblingInfo As New clsSiblingInfo
    Dim currentHeading As clsHeadingInfo
    Set currentHeading = headings(currentIndex)
    
    If currentHeading.level = 2 And currentHeading.parentIndex = -1 Then
        siblingInfo.currentSiblingIndex = currentIndex
        siblingInfo.parentIndex = -1
        
        ' Search for all Level 2 siblings without a parent
        Dim totalChars As Long: totalChars = 0
        Dim i As Long
        For i = 1 To headings.count
            Dim heading As clsHeadingInfo
            Set heading = headings(i)
            
            If heading.level = 2 And heading.parentIndex = -1 Then
                siblingInfo.AddSibling i, heading.Range.Start, heading.charCount
                totalChars = totalChars + heading.charCount
            End If
        Next i
        
        siblingInfo.totalSiblingChars = totalChars
        siblingInfo.ParentIdealChars = totalChars ' Showing internal proportions
        
        CalculateBarSegments siblingInfo
        Set CalculateSiblingInfo = siblingInfo
        Exit Function
    End If
    
    If currentHeading.level <> 2 Or currentHeading.parentIndex = -1 Then
        Set CalculateSiblingInfo = siblingInfo
        Exit Function
    End If
    
    Dim parentHeading As clsHeadingInfo
    Set parentHeading = headings(currentHeading.parentIndex)
    
    siblingInfo.currentSiblingIndex = currentIndex
    siblingInfo.parentIndex = currentHeading.parentIndex
    
    ' Adding orphan text
    If parentHeading.hasOrphanText Then
        siblingInfo.AddSibling -1, parentHeading.Range.End, parentHeading.orphanTextSize
    End If
    
    ' Adding siblings
    Dim j As Long, childIndex As Long
    Dim childHeading As clsHeadingInfo
    
    For j = 1 To parentHeading.childrenCount
        childIndex = parentHeading.childrenIndices(j)
        Set childHeading = headings(childIndex)
        siblingInfo.AddSibling childIndex, childHeading.Range.Start, childHeading.charCount
        totalChars = totalChars + childHeading.charCount
    Next j
    
    siblingInfo.totalSiblingChars = totalChars + parentHeading.orphanTextSize
    
    ' NEW LOGIC: Determine parent ideal character count
    If parentHeading.hasIdealPercent Then
        ' There is an ideal value - use it (scaled if necessary)
        Dim scaledIdeal As Double
        Dim totalSum As Double
        totalSum = CalculateTableIdealSum(headings)
        scaledIdeal = GetScaledIdealPercent(parentHeading.idealPercent, totalSum)
        
        siblingInfo.ParentIdealChars = CLng((scaledIdeal / 100#) * m_cache.totalChars)
    Else
        ' NO ideal value - use actual sibling proportions
        siblingInfo.ParentIdealChars = siblingInfo.totalSiblingChars
    End If
    
    CalculateBarSegments siblingInfo
    Set CalculateSiblingInfo = siblingInfo
End Function

' === CACHE ===

Private Sub InitializeCache()
    ' Always update the cache, not just when invalid
    m_cache.totalChars = ConfigManager.GetUserTotalChars()
    m_cache.level1Count = CountLevel1Headings()
    
    If m_cache.level1Count > 0 Then
        m_cache.defaultIdealPercent = 100# / m_cache.level1Count
    Else
        m_cache.defaultIdealPercent = 100#
    End If
    
    m_cache.isValid = True
End Sub

Public Sub InvalidateCache()
    m_cache.isValid = False
    Set m_cache.headingCollection = Nothing
End Sub

' === CLEAN TEXT FUNCTIONS ===
Private Function CleanHeadingText(rng As Range) As String
    Dim text As String
    text = rng.text
    
    ' Searching for annotation start character
    Dim annotationPos As Long
    annotationPos = InStr(text, ChrW(ANNOTATION_START))
    
    If annotationPos > 0 Then
        ' There is an annotation - simply cut off
        text = Trim(Left(text, annotationPos - 1))
    Else
        ' No annotation - only removing the CR character
        text = Trim(text)
        If Right(text, 1) = vbCr Then text = Left(text, Len(text) - 1)
    End If
    
    CleanHeadingText = text
End Function

Private Function CleanCellText(cellText As String) As String
    Dim i As Integer
    Dim cleanText As String
    
    For i = 1 To Len(cellText)
        Dim c As String
        c = Mid(cellText, i, 1)
        If Asc(c) >= 32 And Asc(c) <> 127 Then
            cleanText = cleanText & c
        End If
    Next i
    
    CleanCellText = Trim(cleanText)
End Function

'=== COUNT HEADINGS ===

Public Function CountHeadings() As Long
    CountHeadings = CountLevel1Headings() + CountLevel2Headings()
End Function

Public Function CountLevel1Headings() As Long
    If m_cache.isValid Then
        CountLevel1Headings = m_cache.level1Count
    Else
        Dim Para As Paragraph
        Dim count As Long
        For Each Para In ActiveDocument.Paragraphs
            If Para.OutlineLevel = wdOutlineLevel1 Then count = count + 1
        Next Para
        CountLevel1Headings = count
    End If
End Function

Private Function CountLevel1HeadingsWithoutIdeal(headings As Collection) As Long
    Dim count As Long
    Dim i As Long
    Dim heading As clsHeadingInfo
    
    For i = 1 To headings.count
        Set heading = headings(i)
        If heading.level = 1 And Not heading.hasIdealPercent And Not heading.isExcluded Then
            count = count + 1
        End If
    Next i
    
    CountLevel1HeadingsWithoutIdeal = count
End Function

Public Function CountLevel2Headings() As Long
    Dim Para As Paragraph
    Dim count As Long
    For Each Para In ActiveDocument.Paragraphs
        If Para.OutlineLevel = wdOutlineLevel2 Then count = count + 1
    Next Para
    CountLevel2Headings = count
End Function

' === ID-S ===

Private Function GetHighestExistingIdFromDocument() As Long
    Dim Para As Paragraph
    Dim existingId As String
    Dim maxId As Long
    
    For Each Para In ActiveDocument.Paragraphs
        If Para.OutlineLevel = wdOutlineLevel1 Then
            existingId = ExtractIdFromParagraph(Para)
                        
            If Len(existingId) > 0 Then
                Dim idNumber As Long
                idNumber = Val(Right(existingId, 3))
                If idNumber > maxId Then maxId = idNumber
            End If
        End If
    Next Para

    GetHighestExistingIdFromDocument = maxId
End Function

Private Function ExtractIdFromParagraph(Para As Paragraph) As String
    Dim text As String
    text = Para.Range.text
    
    ' Search: ANNOTATION_START + ID pattern
    Dim startPos As Long
    startPos = InStr(text, ChrW(ANNOTATION_START))
    If startPos = 0 Then
        ' No annotation start character
        ExtractIdFromParagraph = ""
        Exit Function
    End If
    
    ' Examining 5 characters after ANNOTATION_START (ID + 3 digits)
    If Len(text) < startPos + 5 Then
        ' Not enough characters
        ExtractIdFromParagraph = ""
        Exit Function
    End If
    
    ' Extracting ID
    Dim potentialId As String
    potentialId = Mid(text, startPos + 1, 5) ' 5 characters after START character
    
    ' Validation: "ID" + 3 digits
    If Left(potentialId, 2) = "ID" And IsNumeric(Right(potentialId, 3)) Then
        ExtractIdFromParagraph = potentialId
    Else
        ' Invalid ID format
        ExtractIdFromParagraph = ""
    End If
End Function
