Attribute VB_Name = "Declaration_Module"

'Declare the picture pasting function
Public Declare Function StretchBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal ClipX As Long, ByVal ClipY As Long, ByVal dwRop As Long) As Long
'The array of values to keep track of what squares are which type
Global maparray() As Integer
'tree map array
'The currently clicked position
Global CurX, CurY As Integer
'The view position
Global ViewX, ViewY As Integer
'The size of the map
Global SizeX, SizeY As Integer
'Whether the mouse is up or down
Global MouseState As Byte
'The currently selected picture
Global CurPicIndex As Integer
'The zoom value
Global Zoom As Integer
