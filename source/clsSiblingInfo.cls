VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSiblingInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' =============================================================================
' clsSiblingInfo Class - Sibling Headings Information
' =============================================================================

Option Explicit

Private m_totalSiblingChars As Long
Private m_parentIdealChars As Long
Private m_currentSiblingIndex As Long
Private m_parentIndex As Long
Private m_siblings As Collection

' NEW: Bar positional data
Private m_barWidth As Integer
Private m_barSegments As Collection
Private m_currentSiblingBarStart As Integer
Private m_currentSiblingBarWidth As Integer
Private m_totalFilledBarWidth As Integer

' Constructor
Private Sub Class_Initialize()
    Set m_siblings = New Collection
    Set m_barSegments = New Collection
    m_totalSiblingChars = 0
    m_parentIdealChars = 0
    m_currentSiblingIndex = -1
    m_parentIndex = -1
    m_barWidth = 40 ' Default
End Sub

' Original Properties
Public Property Get totalSiblingChars() As Long
    totalSiblingChars = m_totalSiblingChars
End Property

Public Property Let totalSiblingChars(value As Long)
    m_totalSiblingChars = value
End Property

Public Property Get ParentIdealChars() As Long
    ParentIdealChars = m_parentIdealChars
End Property

Public Property Let ParentIdealChars(value As Long)
    m_parentIdealChars = value
End Property

Public Property Get currentSiblingIndex() As Long
    currentSiblingIndex = m_currentSiblingIndex
End Property

Public Property Let currentSiblingIndex(value As Long)
    m_currentSiblingIndex = value
End Property

Public Property Get parentIndex() As Long
    parentIndex = m_parentIndex
End Property

Public Property Let parentIndex(value As Long)
    m_parentIndex = value
End Property

Public Property Get siblings() As Collection
    Set siblings = m_siblings
End Property

Public Property Set siblings(value As Collection)
    Set m_siblings = value
End Property

' Bar positional properties
Public Property Get barWidth() As Integer
    barWidth = m_barWidth
End Property

Public Property Let barWidth(value As Integer)
    m_barWidth = value
End Property

Public Property Get barSegments() As Collection
    Set barSegments = m_barSegments
End Property

Public Property Get currentSiblingBarStart() As Integer
    currentSiblingBarStart = m_currentSiblingBarStart
End Property

Public Property Let currentSiblingBarStart(value As Integer)
    m_currentSiblingBarStart = value
End Property

Public Property Get currentSiblingBarWidth() As Integer
    currentSiblingBarWidth = m_currentSiblingBarWidth
End Property

Public Property Let currentSiblingBarWidth(value As Integer)
    m_currentSiblingBarWidth = value
End Property

Public Property Get totalFilledBarWidth() As Integer
    totalFilledBarWidth = m_totalFilledBarWidth
End Property

Public Property Let totalFilledBarWidth(value As Integer)
    m_totalFilledBarWidth = value
End Property

' Original helper methods
Public Sub AddSibling(index As Long, position As Long, charCount As Long)
    m_siblings.Add Array(index, position, charCount)
End Sub

Public Function GetSiblingCount() As Long
    GetSiblingCount = m_siblings.count
End Function

Public Sub ClearSiblings()
    Set m_siblings = New Collection
    Set m_barSegments = New Collection
End Sub

' Add bar segment
Public Sub AddBarSegment(siblingIndex As Long, startPos As Integer, width As Integer, segmentType As String)
    ' segmentType: "current", "sibling", "orphan"
    m_barSegments.Add Array(siblingIndex, startPos, width, segmentType)
End Sub

Public Function GetBarSegmentCount() As Long
    GetBarSegmentCount = m_barSegments.count
End Function
