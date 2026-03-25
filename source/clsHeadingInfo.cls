VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsHeadingInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' =============================================================================
' clsHeadingInfo - Optimized Class Module
' =============================================================================

Option Explicit

' Basic Data
Public Range As Range
Public level As Integer
Public cleanText As String
Public charCount As Long
Public parentIndex As Long
Public percentage As Double

' Heading information
Private m_headingId As String

' Table-related data
Private m_idealPercent As Double
Private m_limitPercent As Double
Private m_idealPercentCalculated As Boolean

' Hierarchy data
Private m_parentIndex As Long
Private m_childrenIndices As Collection
Private m_orphanTextSize As Long
Private m_orphanTextCalculated As Boolean

' Positional data (for faster searching)
Private m_startPosition As Long
Private m_endPosition As Long
Private m_positionsCalculated As Boolean

' Constructor
Private Sub Class_Initialize()
    Set m_childrenIndices = New Collection
    m_parentIndex = -1
    m_idealPercent = 0  ' 0 = not set, -1 = excluded, >0 = has value
    m_limitPercent = 0
    m_orphanTextSize = -1
    m_idealPercentCalculated = False
    m_orphanTextCalculated = False
    m_positionsCalculated = False
End Sub

' === HEADING HANDLING ===
Public Property Get headingId() As String
    headingId = m_headingId
End Property

Public Property Let headingId(value As String)
    m_headingId = value
End Property

' === TABLE DATA ===
Public Property Get idealPercent() As Double
    idealPercent = m_idealPercent
End Property

Public Property Let idealPercent(value As Double)
    m_idealPercent = value
    m_idealPercentCalculated = True
End Property

Public Property Get limitPercent() As Double
    limitPercent = m_limitPercent
End Property

Public Property Let limitPercent(value As Double)
    m_limitPercent = value
End Property

Public Property Get hasIdealPercent() As Boolean
    hasIdealPercent = (m_idealPercent <> 0)  ' 0 = not set, -1 = excluded, >0 = has value
End Property

Public Property Get hasLimitPercent() As Boolean
    hasLimitPercent = (m_limitPercent <> 0)
End Property

Public Property Get isExcluded() As Boolean
    isExcluded = (m_idealPercent = -1)
End Property

' === HIERARCHY HANDLING ===
Public Property Get childrenIndices() As Collection
    Set childrenIndices = m_childrenIndices
End Property

Public Sub AddChildIndex(childIndex As Long)
    m_childrenIndices.Add childIndex
End Sub

Public Property Get hasChildren() As Boolean
    hasChildren = (m_childrenIndices.count > 0)
End Property

Public Property Get childrenCount() As Long
    childrenCount = m_childrenIndices.count
End Property

' === ORPHAN TEXT HANDLING ===
Public Property Get orphanTextSize() As Long
    orphanTextSize = m_orphanTextSize
End Property

Public Property Let orphanTextSize(value As Long)
    m_orphanTextSize = value
    m_orphanTextCalculated = True
End Property

Public Property Get hasOrphanText() As Boolean
    hasOrphanText = (m_orphanTextSize > 0)
End Property

Public Property Get orphanTextCalculated() As Boolean
    orphanTextCalculated = m_orphanTextCalculated
End Property

' === POSITIONAL DATA ===
Public Property Get startPosition() As Long
    If Not m_positionsCalculated Then
        CalculatePositions
    End If
    startPosition = m_startPosition
End Property

Public Property Get endPosition() As Long
    If Not m_positionsCalculated Then
        CalculatePositions
    End If
    endPosition = m_endPosition
End Property

Private Sub CalculatePositions()
    m_startPosition = Range.Start
    m_endPosition = Range.End
    m_positionsCalculated = True
End Sub

' === HELPER FUNCTIONS ===
Public Function GetIdealCharCount() As Long
    If hasIdealPercent And Not isExcluded Then
        Dim totalChars As Long
        totalChars = ConfigManager.GetUserTotalChars()
        GetIdealCharCount = CLng((m_idealPercent / 100#) * totalChars)
    Else
        GetIdealCharCount = 0
    End If
End Function

Public Function GetLimitCharCount() As Long
    If hasLimitPercent And Not isExcluded Then
        Dim totalChars As Long
        totalChars = ConfigManager.GetUserTotalChars()
        GetLimitCharCount = CLng((m_limitPercent / 100#) * totalChars)
    Else
        GetLimitCharCount = 0
    End If
End Function

Public Function IsOverLimit() As Boolean
    If hasLimitPercent And Not isExcluded Then
        IsOverLimit = (percentage > m_limitPercent)
    Else
        IsOverLimit = False
    End If
End Function

Public Function GetPercentageDeviation() As Double
    If hasIdealPercent And Not isExcluded Then
        GetPercentageDeviation = percentage - m_idealPercent
    Else
        GetPercentageDeviation = 0
    End If
End Function

' Invalidate cache on modification
Public Sub InvalidateCache()
    m_headingCalculated = False 'Gemini deleted this line
    m_orphanTextCalculated = False
    m_positionsCalculated = False
End Sub
