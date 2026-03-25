Attribute VB_Name = "Constants"

' =============================================================================
' Constants Module
' =============================================================================

Public Const ANNOTATION_START As Long = &H200C
Public Const ANNOTATION_END As Long = &H200C
Public Const FIRST_RUN_KEY As String = "TextBalanceFirstRun"
Public Const AUTO_SAVE_KEY As String = ""
Public Const USER_TOTAL_CHARS_KEY As String = "TextBalanceUserTotal"
Public Const DEFAULT_TOLERANCE As Double = 5#
Public Const DEFAULT_FONT_SIZE As Integer = 13

' Color constants
Public Const COLOR_GREEN As Long = 32768        ' RGB(0, 128, 0)
Public Const COLOR_ORANGE As Long = 26316       ' RGB(204, 102, 0)
Public Const COLOR_RED As Long = 204            ' RGB(204, 0, 0)
Public Const COLOR_BLACK As Long = 0            ' RGB(0, 0, 0)
Public Const COLOR_DARK_GRAY As Long = 6710886  ' RGB(102, 102, 102)
Public Const COLOR_LIGHT_GRAY As Long = 13421772 ' RGB(204, 204, 204)

Public Const DEFAULT_SCALING As Long = 33

' Unicode characters for the visual bar
Public Const CHAR_FILLED As Long = 9608
Public Const CHAR_SIBLING As Long = 9619
Public Const CHAR_EMPTY As Long = 9617


Public Const DEFAULT_BAR_WIDTH As Integer = 40


' Segment types
Public Const SEGMENT_TYPE_CURRENT As String = "current"
Public Const SEGMENT_TYPE_SIBLING As String = "sibling"
Public Const SEGMENT_TYPE_ORPHAN As String = "orphan"

' Separate default values for H1 and H2
Public Const DEFAULT_H1_FONT_SIZE As Integer = 13
Public Const DEFAULT_H2_FONT_SIZE As Integer = 11

' Speech tempo
Public Const DEFAULT_SPEECH_TEMPO As Integer = 180 ' karakter/perc

' Version checking
Public Const APP_VERSION As String = "1.0.0"
Public Const VERSION_CHECK_URL As String = "https://raw.githubusercontent.com/textbalance/macro/main/version.json"
' Colorings
' PASTEL colorings
'Public Const COLORING_PASTEL_OPTIMAL As Long = 10079487    ' Soft green RGB(160, 230, 153)
'Public Const COLORING_PASTEL_CLOSE As Long = 8421631       ' Soft orange RGB(255, 204, 128)
'Public Const COLORING_PASTEL_OUT As Long = 9695743         ' Soft red RGB(255, 153, 153)

' VIBRANT colorings
'Public Const COLORING_VIBRANT_OPTIMAL As Long = 65280      ' Bright green RGB(0, 255, 0)
'Public Const COLORING_VIBRANT_CLOSE As Long = 36863        ' Bright orange RGB(255, 140, 0)
'Public Const COLORING_VIBRANT_OUT As Long = 255            ' Bright red RGB(255, 0, 0)

' NATURAL colorings
'Public Const COLORING_NATURAL_OPTIMAL As Long = 6723891    ' Forest green RGB(51, 102, 102)
'Public Const COLORING_NATURAL_CLOSE As Long = 2263842      ' Earth brown RGB(162, 82, 34)
'Public Const COLORING_NATURAL_OUT As Long = 3355443        ' Deep red RGB(51, 51, 51)

' MONOCHROME colorings
'Public Const COLORING_MONO_OPTIMAL As Long = 8421504       ' Dark gray RGB(128, 128, 128)
'Public Const COLORING_MONO_CLOSE As Long = 6710886         ' Medium gray RGB(166, 166, 102)
'Public Const COLORING_MONO_OUT As Long = 3355443           ' Light gray RGB(51, 51, 51)

' CONTRAST colorings
'Public Const COLORING_CONTRAST_OPTIMAL As Long = 32768     ' Strong green RGB(0, 128, 0)
'Public Const COLORING_CONTRAST_CLOSE As Long = 255         ' Strong orange RGB(255, 128, 0)
'Public Const COLORING_CONTRAST_OUT As Long = 128

' Table méretek
'Public Const TABLE_SIZE_SMALL_MULTIPLIER As Double = 0.8
'Public Const TABLE_SIZE_NORMAL_MULTIPLIER As Double = 1#
'Public Const TABLE_SIZE_LARGE_MULTIPLIER As Double = 1.2

' ============================================================================
' BAR KARAKTEREK DEFINÍCIÓK
' ============================================================================
'Public Const BAR_SET_STANDARD As String = "standard"
'Public Const BAR_SET_ASCII As String = "ascii"
'Public Const BAR_SET_BRAILLE As String = "braille"
'Public Const BAR_SET_GEOMETRIC As String = "geometric"

' Standard set (current)
Public Const STANDARD_CHAR_FILLED As Long = 9608       ' █
Public Const STANDARD_CHAR_SIBLING As Long = 9619      ' ▓
Public Const STANDARD_CHAR_EMPTY As Long = 9617        ' ░

' ASCII set
'Public Const ASCII_CHAR_FILLED As String = "#"
'Public Const ASCII_CHAR_SIBLING As String = "+"
'Public Const ASCII_CHAR_EMPTY As String = "-"

' Braille set
'Public Const BRAILLE_CHAR_FILLED As Long = 10239       ' ⣿
'Public Const BRAILLE_CHAR_SIBLING As Long = 10212      ' ⣤
'Public Const BRAILLE_CHAR_EMPTY As Long = 10176        ' ⣀

' Geometric set
'Public Const GEOMETRIC_CHAR_FILLED As Long = 9632      ' ■
'Public Const GEOMETRIC_CHAR_SIBLING As Long = 9670     ' ◆
'Public Const GEOMETRIC_CHAR_EMPTY As Long = 9675       ' ○
