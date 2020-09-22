Attribute VB_Name = "MouseCoords"
Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)


Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long


Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
    Public Const MOUSEEVENTF_LEFTDOWN = &H2
    Public Const MOUSEEVENTF_LEFTUP = &H4
    Public Const MOUSEEVENTF_MIDDLEDOWN = &H20
    Public Const MOUSEEVENTF_MIDDLEUP = &H40
    Public Const MOUSEEVENTF_RIGHTDOWN = &H8
    Public Const MOUSEEVENTF_RIGHTUP = &H10
    Public Const MOUSEEVENTF_MOVE = &H1
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Const MERGEPAINT = &HBB0226
Private Const SRCAND = &H8800C6
Private Const SRCCOPY = &HCC0020
Public Type POINTAPI
    mousex As Long
    mousey As Long
    End Type
'**************************************
' Name: A mouse module, FINALLY!!! Move,
'     click, +more
' Description:This module has the follow
'     ing functions (pretty self explanitory):
'     GetX, GetY, LeftClick, LeftDown, LeftUp,
'     RightClick, RightUp, RightDown, MiddleCl
'     ick, MiddleDown, MiddleUp, MoveMouse, Se
'     tMousePos
' By: Arthur Chaparyan
'
'
' Inputs:None
'
' Returns:None
'
'Assumes:You should know how to create a
'     nd use a module. If you have any questio
'     ns, please submit a comment, thanX
'
'Side Effects:None
'
'Warranty:
'Code provided by Planet Source Code(tm)
'     (http://www.Planet-Source-Code.com) 'as
'     is', without warranties as to performanc
'     e, fitness, merchantability,and any othe
'     r warranty (whether expressed or implied
'     ).
'Terms of Agreement:
'By using this source code, you agree to
'     the following terms...
' 1) You may use this source code in per
'     sonal projects and may compile it into a
'     n .exe/.dll/.ocx and distribute it in bi
'     nary format freely and with no charge.
' 2) You MAY NOT redistribute this sourc
'     e code (for example to a web site) witho
'     ut written permission from the original
'     author.Failure to do so is a violation o
'     f copyright laws.
' 3) You may link to this code from anot
'     her website, provided it is not wrapped
'     in a frame.
' 4) The author of this code may have re
'     tained certain additional copyright righ
'     ts.If so, this is indicated in the autho
'     r's description.
'**************************************




Public Function GetX() As Long
    Dim n As POINTAPI
    GetCursorPos n
    GetX = n.mousex
End Function


Public Function GetY() As Long
    Dim n As POINTAPI
    GetCursorPos n
    GetY = n.mousey
End Function


Public Sub LeftClick()
    LeftDown
    LeftUp
End Sub


Public Sub LeftDown()
    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
End Sub


Public Sub LeftUp()
    mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
End Sub


Public Sub MiddleClick()
    MiddleDown
    MiddleUp
End Sub


Public Sub MiddleDown()
    mouse_event MOUSEEVENTF_MIDDLEDOWN, 0, 0, 0, 0
End Sub


Public Sub MiddleUp()
    mouse_event MOUSEEVENTF_MIDDLEUP, 0, 0, 0, 0
End Sub


Public Sub MoveMouse(xMove As Long, yMove As Long)
    mouse_event MOUSEEVENTF_MOVE, xMove, yMove, 0, 0
End Sub


Public Sub RightClick()
    RightDown
    RightUp
End Sub


Public Sub RightDown()
    mouse_event MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0
End Sub


Public Sub RightUp()
    mouse_event MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0
End Sub


Public Sub SetMousePos(XPos As Long, YPos As Long)
    SetCursorPos XPos, YPos
End Sub

