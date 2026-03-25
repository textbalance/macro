Attribute VB_Name = "TableManager"

' =============================================================================
' TableManager Module - Table Management
' =============================================================================
Private Const TABLE_ID As String = "TEXT_BALANCE_TABLE_ID"

'==============================================================================
' Other important Public Subs and Functions
'==============================================================================

Public Function CleanCellText(cellText As String) As String
    Dim i As Integer
    Dim cleanText As String
    cleanText = ""
    
    ' Iterate through every character in cellText
    For i = 1 To Len(cellText)
        Dim c As String
        c = Mid(cellText, i, 1)
        
        ' If character is printable, add to cleanText
        If Asc(c) >= 32 And Asc(c) <> 127 Then
            cleanText = cleanText & c
        End If
    Next i
    
    ' Return cleaned text with leading/trailing spaces removed
    CleanCellText = Trim(cleanText)
End Function

Public Function GetCurrentTableTotalChars() As Long
    Dim tbl As table
    Set tbl = FindExistingTable()
    
    If tbl Is Nothing Then
        GetCurrentTableTotalChars = 0
        Exit Function
    End If
    
    ' TotalChar value is in the second cell of the first row
    Dim cellText As String
    cellText = CleanCellText(tbl.cell(1, 2).Range.text)
    
    If IsNumeric(cellText) Then
        GetCurrentTableTotalChars = CLng(cellText)
    Else
        GetCurrentTableTotalChars = 0
    End If
End Function

'==============================================================================
' Table Existence Check
'==============================================================================

Public Function GetOrCreateTable() As table
    Set GetOrCreateTable = FindExistingTable()
    
    If GetOrCreateTable Is Nothing Then
        Set GetOrCreateTable = CreateTable()
    Else
        ' If it exists but in the wrong place, move it
        If GetOrCreateTable.Range.Start > 0 Then
            Dim newTable As table
            Set newTable = CreateTable()
            
            GetOrCreateTable.Delete
            Set GetOrCreateTable = newTable
        Else
            ValidateAndRepairTable GetOrCreateTable
        End If
    End If
End Function

Private Sub ValidateAndRepairTable(tbl As table)
    ' Simple repairs if necessary
    If tbl.Rows.count < 3 Then
        ' Replace missing rows
        Do While tbl.Rows.count < 3
            tbl.Rows.Add
        Loop
        PopulateTableHeaders tbl
    End If
End Sub

Public Function FindExistingTable() As table
    Dim tbl As table

    For Each tbl In ActiveDocument.tables
        If IsCharacterCountTable(tbl) Then
            Set FindExistingTable = tbl
            Exit Function
        End If
    Next tbl
    
    Set FindExistingTable = Nothing
End Function


Private Function IsCharacterCountTable(tbl As table) As Boolean
    If tbl.Rows.count < 3 Or tbl.Columns.count < 4 Then
        IsCharacterCountTable = False
        Exit Function
    End If
        
    IsCharacterCountTable = (InStr(tbl.cell(2, 1).Range.text, TABLE_ID) > 0 Or InStr(tbl.cell(2, 1).Range.text, "CHAR_COUNT_TABLE_ID") > 0)
End Function

'==============================================================================
' Table Creation and Formatting
'==============================================================================

Public Function CreateTable() As table
    Dim tbl As table
    Dim rng As Range
    Dim doc As Document
    Set doc = ActiveDocument
        
    Dim insertPos As Long
    insertPos = 0
    
    Set rng = doc.Range(insertPos, insertPos)
    Set tbl = doc.tables.Add(rng, 3, 4)
    
    ' Style repair if necessary
    Dim Para As Paragraph
    Set Para = doc.Range(insertPos, insertPos).Paragraphs(1)
    If Para.Style <> ActiveDocument.Styles(wdStyleNormal) Then
        tbl.Range.Style = wdStyleNormal
    End If
    
    FormatTable tbl
    
    PopulateTableHeaders tbl
    
    Set CreateTable = tbl
End Function


Private Function IsHeadingStyle(Para As Paragraph) As Boolean
    IsHeadingStyle = (Para.Style <> ActiveDocument.Styles(wdStyleNormal))
End Function

Private Sub FormatTable(tbl As table)
    With tbl
        .Borders.Enable = True
        .Range.Font.size = 8
        .Range.Font.color = vbBlack
        .Range.Font.Hidden = True

        With .Range.ParagraphFormat
            .LeftIndent = 0
            .RightIndent = 0
            .FirstLineIndent = 0
            .SpaceBefore = 3
            .SpaceAfter = 0
            .LineSpacing = LinesToPoints(1.5)
        End With

        .Columns(1).width = Application.CentimetersToPoints(4)
        .Columns(2).width = Application.CentimetersToPoints(1.5)
        .Columns(3).width = Application.CentimetersToPoints(1.5)
        .Columns(4).width = Application.CentimetersToPoints(1.5)

        With .Rows
            .WrapAroundText = True
            .HorizontalPosition = Application.CentimetersToPoints(12)
            .VerticalPosition = Application.CentimetersToPoints(12)
            .RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
            .RelativeVerticalPosition = wdRelativeVerticalPositionPage
            .Alignment = wdAlignRowLeft
        End With
    End With
End Sub

Private Sub PopulateTableHeaders(tbl As table)
    With tbl
        .cell(1, 1).Range.text = "TotalChar"
        .cell(1, 2).Range.text = CStr(ConfigManager.GetUserTotalChars())
        .cell(1, 2).Merge MergeTo:=.cell(1, 4)

        .cell(2, 1).Range.text = TABLE_ID
        .cell(2, 2).Range.text = "Actual%"
        .cell(2, 3).Range.text = "Ideal%"
        .cell(2, 4).Range.text = "Limit%"
    End With
End Sub

'==============================================================================
' Updating Table - Main Logic
'==============================================================================

Public Sub UpdateTable(tbl As table, headings As Collection)
    SyncTableWithHeadings tbl, headings
    
    ' Update summary row
    UpdateSummaryRow tbl, headings
End Sub

Private Sub SyncTableWithHeadings(tbl As table, headings As Collection)
    
    Dim documentHeadingIds As Object
    Set documentHeadingIds = BuildHeadingIdIndex(headings)
    ClearDataRows tbl
    RebuildRows tbl, headings

End Sub
Private Sub UpdateSummaryRow(tbl As table, headings As Collection)
    Dim summaryRowIndex As Long
    summaryRowIndex = tbl.Rows.count
    
    ' Total actual percentage
    Dim totalActual As Double
    Dim i As Long
    Dim heading As clsHeadingInfo
    
    For i = 1 To headings.count
        Set heading = headings(i)
        If heading.level = 1 And Not heading.isExcluded Then
            totalActual = totalActual + heading.percentage
        End If
    Next i
    
    ' Total ideal percentage (excluded headings are skipped by CalculateTableIdealSum)
    Dim totalIdeal As Double
    totalIdeal = HeadingProcessor.CalculateTableIdealSum(headings)
    
    ' Update table cells
    tbl.cell(summaryRowIndex, 2).Range.text = Format(totalActual, "0.0") & "%"
    tbl.cell(summaryRowIndex, 3).Range.text = Format(totalIdeal, "0.0") & "%"
    
    ' Coloring - Ideal column
    If totalIdeal > 100 Then
        tbl.cell(summaryRowIndex, 3).Range.Font.color = COLOR_RED
    Else
        tbl.cell(summaryRowIndex, 3).Range.Font.color = COLOR_BLACK
    End If
End Sub

Private Function GetCurrentTableIdealSum(tbl As table) As Double
    Dim totalSum As Double
    totalSum = 0
    
    For i = 3 To tbl.Rows.count - 1 ' Skip last row
        Dim cellText As String
        cellText = CleanCellText(tbl.cell(i, 3).Range.text)
        If Len(cellText) > 0 And IsNumeric(Replace(cellText, "%", "")) Then
            totalSum = totalSum + Val(Replace(cellText, "%", ""))
        End If
    Next i
    
    GetCurrentTableIdealSum = totalSum
End Function


Private Function BuildHeadingIdIndex(headings As Collection) As Object
    Dim index As Object
    Set index = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    Dim heading As clsHeadingInfo
    
    For i = 1 To headings.count
        Set heading = headings(i)
        If heading.level = 1 Then
            index(heading.headingId) = i
        End If
    Next i
    
    Set BuildHeadingIdIndex = index
End Function

Private Function ExtractHeadingFromCell(cellText As String) As String
    ' Extract heading text without headingId
    Dim bracketPos As Long
    bracketPos = InStr(cellText, " [ID_")
    
    If bracketPos > 0 Then
        ExtractHeadingFromCell = Trim(Left(cellText, bracketPos - 1))
    Else
        ExtractHeadingFromCell = CleanCellText(cellText)
    End If
End Function

'==============================================================================
' Update Helpers
'==============================================================================

Private Sub ClearDataRows(tbl As table)
    Dim i As Long
    For i = tbl.Rows.count To 3 Step -1
        tbl.Rows(i).Delete
    Next i
End Sub

Public Function ExtractHeadingId(cellText As String) As String
    Dim startPos As Long
    Dim endPos As Long
    
    startPos = InStr(cellText, "[ID")
    If startPos > 0 Then
        endPos = InStr(startPos, cellText, "]")
        If endPos > startPos Then
            ExtractHeadingId = Mid(cellText, startPos + 1, endPos - startPos - 1)
        End If
    End If
End Function

'==============================================================================
' Table Rebuild
'==============================================================================

Private Sub RebuildRows(tbl As table, headings As Collection)
    ' Simplified - works from heading objects
    Dim i As Long
    Dim heading As clsHeadingInfo
    Dim newRow As row
    Dim headingId As String

    For i = 1 To headings.count
        Set heading = headings(i)
        
        If heading.level = 1 Then
            Set newRow = tbl.Rows.Add
            headingId = heading.headingId
            
            ' Heading text + hidden headingID
            newRow.Cells(1).Range.text = heading.cleanText & " [" & headingId & "]"
            
            ' Actual
            If heading.isExcluded Then
                newRow.Cells(2).Range.text = "-"
            Else
                newRow.Cells(2).Range.text = Format(heading.percentage, "0.0") & "%"
            End If
            
            ' Writing back data from heading object
            If heading.hasIdealPercent Then
                If heading.isExcluded Then
                    newRow.Cells(3).Range.text = "-"
                Else
                    newRow.Cells(3).Range.text = Format(heading.idealPercent, "0.0") & "%"
                End If
            End If
            
            If heading.hasLimitPercent Then
                If heading.isExcluded Then
                    newRow.Cells(4).Range.text = "-"
                Else
                    newRow.Cells(4).Range.text = Format(heading.limitPercent, "0.0") & "%"
                End If
            End If
            
            ' Hide Heading ID
            HideHeadingIdInCell newRow.Cells(1), headingId
        End If
    Next i
    AddSummaryRow tbl
End Sub

Private Sub AddSummaryRow(tbl As table)
    Dim summaryRow As row
    Set summaryRow = tbl.Rows.Add
    
    With summaryRow
        .Cells(1).Range.text = "SUM:"
        .Cells(2).Range.text = "" ' Actual% - leave empty
        .Cells(3).Range.text = "0.0%" ' Ideal% - we will calculate this
        .Cells(4).Range.text = "" ' Limit% - leave empty
    End With
End Sub

'==============================================================================
' Common Helpers
'==============================================================================

Private Sub HideHeadingIdInCell(cell As cell, headingId As String)
    Dim hdRange As Range
    Set hdRange = cell.Range
    Dim startPos As Long
    
    startPos = InStr(hdRange.text, "[" & headingId)
    If startPos > 0 Then
        hdRange.SetRange hdRange.Start + startPos - 1, hdRange.End - 1
        hdRange.Font.size = 1
        hdRange.Font.color = wdColorWhite
    End If
End Sub
