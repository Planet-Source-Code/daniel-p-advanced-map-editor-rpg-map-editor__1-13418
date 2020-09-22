VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000006&
   BorderStyle     =   0  'None
   Caption         =   "Map Editor"
   ClientHeight    =   6180
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   10620
   DrawMode        =   1  'Blackness
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   11.25
      Charset         =   204
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000006&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   412
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   708
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   1440
      ScaleHeight     =   195
      ScaleWidth      =   435
      TabIndex        =   0
      Top             =   240
      Width           =   495
   End
   Begin MSComDlg.CommonDialog cmndlg1 
      Left            =   720
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FileName        =   "default.mpd"
      Filter          =   "*.mpd"
   End
   Begin VB.Menu TS 
      Caption         =   "Select TileSet"
      Begin VB.Menu EG 
         Caption         =   "Enable Grid"
      End
      Begin VB.Menu TGL1 
         Caption         =   " Tileset Ground"
      End
      Begin VB.Menu TGL2 
         Caption         =   "TileSet Walls"
      End
      Begin VB.Menu TT 
         Caption         =   "TileSet Things"
         Begin VB.Menu DT 
            Caption         =   "Delete Thing"
         End
         Begin VB.Menu AST1 
            Caption         =   "A Statue (Type 1)"
         End
         Begin VB.Menu AST1L 
            Caption         =   "A Statue (Type 1 Left)"
         End
         Begin VB.Menu AST1R 
            Caption         =   "A Statue (Type 1 Right)"
         End
         Begin VB.Menu AST2L 
            Caption         =   "A Statue (Type 2 Left)"
         End
         Begin VB.Menu AST2R 
            Caption         =   "A Statue (Type 2 Right)"
         End
         Begin VB.Menu AOCT1 
            Caption         =   "An Open Chest (Type 1)"
         End
         Begin VB.Menu AOCT2 
            Caption         =   "An Open Chest (Type 2)"
         End
         Begin VB.Menu ACCT1 
            Caption         =   "A Closed Chest (Type 1)"
         End
         Begin VB.Menu ACCT2 
            Caption         =   "A Closed Chest (Type 2)"
         End
         Begin VB.Menu ART 
            Caption         =   "A Red Torch"
         End
         Begin VB.Menu ABT 
            Caption         =   "A Blue Torch"
         End
         Begin VB.Menu CS 
            Caption         =   "Crossed Swords "
         End
         Begin VB.Menu AW 
            Caption         =   "A Well"
         End
         Begin VB.Menu RIPS 
            Caption         =   "Rest in Peace Solid"
         End
         Begin VB.Menu RIPC 
            Caption         =   "Rest in Peace Cross"
         End
         Begin VB.Menu AWind 
            Caption         =   "A Window"
         End
         Begin VB.Menu ARS 
            Caption         =   "A Red Shield"
         End
         Begin VB.Menu ARC 
            Caption         =   "A Red Cirtains"
         End
         Begin VB.Menu AA 
            Caption         =   "An Angel"
         End
      End
      Begin VB.Menu LM 
         Caption         =   "Load Map"
      End
      Begin VB.Menu SM 
         Caption         =   "Save Map"
      End
      Begin VB.Menu cg 
         Caption         =   "Change Ground Filling Tile"
         Begin VB.Menu TNO 
            Caption         =   "Tile NO.0"
         End
         Begin VB.Menu l1 
            Caption         =   "Tile N0.1"
         End
         Begin VB.Menu l2 
            Caption         =   "Tile N0.2"
         End
         Begin VB.Menu l3 
            Caption         =   "Tile N0.3"
         End
         Begin VB.Menu l4 
            Caption         =   "Tile N0.4"
         End
         Begin VB.Menu l5 
            Caption         =   "Tile N0.5"
         End
         Begin VB.Menu l6 
            Caption         =   "Tile N0.6"
         End
      End
      Begin VB.Menu fr 
         Caption         =   "Frame Rate"
         Begin VB.Menu fr10 
            Caption         =   "10 Frames"
         End
         Begin VB.Menu fr15 
            Caption         =   "15 Frames"
         End
         Begin VB.Menu fr20 
            Caption         =   "20 Frames"
         End
         Begin VB.Menu fr25 
            Caption         =   "25 Frames"
         End
         Begin VB.Menu fr30 
            Caption         =   "30 Frames"
         End
         Begin VB.Menu SCFrameRate 
            Caption         =   "Set Custom Frame Rate"
         End
      End
      Begin VB.Menu E 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim mdx As New DirectX7
Dim mdd As DirectDraw7
Dim Remo, Num As Integer
Dim Wthing() As Things
Dim GroundSurf(0 To 15) As DirectDrawSurface7
Dim GroundDesc(0 To 15) As DDSURFACEDESC2


Dim WallsSurf(0 To 30) As DirectDrawSurface7
Dim WallsDesc(0 To 30) As DDSURFACEDESC2


Dim ThingsSurf(0 To 20) As DirectDrawSurface7
Dim ThingsDesc(0 To 20) As DDSURFACEDESC2
Dim MyThingRECT As RECT

Dim GroundA(0 To 500, 0 To 500) As Integer
Dim WallsA(0 To 500, 0 To 500) As Integer
Dim UPDATEDMAP As Boolean




Dim GroundL1RECT As RECT

Dim WallsRECT As RECT
Dim Frame As Integer
Dim ThingsRECT As RECT

Dim UpLeftKey As Boolean
Dim UpRightKey As Boolean
Dim DownRightKey As Boolean
Dim DownLeftKey As Boolean

Dim LastTimeChecked As Long
Dim DelayTime As Integer
Dim CustomFrameRate As Integer
Dim MyFont As StdFont
Dim msurfFront As DirectDrawSurface7 'Front Surface
Dim msurfBack As DirectDrawSurface7 'Back Surface
Dim FinalPos As Integer
Dim MouseSurf As DirectDrawSurface7 'Mouse Surface

'strings..


Dim ddsd5 As DDSURFACEDESC2 'Desc. for Mouse Surface
Dim ddsdMain As DDSURFACEDESC2 'Main Surface Desc.
Dim ddsdFlip As DDSURFACEDESC2 'Flip(back) Surf. Desc.

 Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Dim FrameIndex As Integer

Const SCREEN_WIDTH = 800
Const SCREEN_HEIGHT = 600
Const SCREEN_BITDEPTH = 16
Const TILE_WIDTH = 48
Const TILE_HEIGHT = 48

Dim ssTexts As Integer
Dim mrectScreen As RECT                     'Rectangle the size of the screen
Public DoUnitMove As Boolean
Dim UnitRECT As RECT

Dim MyUnit As DirectDrawSurface7
Dim MyUnitDesc As DDSURFACEDESC2
Dim SubIndex As Integer
Dim ViewX As Long                       '"Player" X coordinate
Dim ViewY As Long                        '"Player" Y coordinate
Dim i, j As Integer
Dim mousex As Integer
Dim mousey As Integer
Dim CurX As Integer
Dim CurY As Integer
Dim MenuRECT As RECT

Public Marked As Integer
Dim UnitXPos, UnitYPos As Integer


Dim Finded As Boolean
Dim Thing As Integer


Dim mblnRunning As Boolean                  'Boolean Flags for run - time
Dim LeftKey As Boolean
Dim RightKey As Boolean
Dim UpKey As Boolean
Dim DownKey As Boolean
Public CurrentPos, NextPos As Integer
Dim TileMenuX  As Integer
Dim TileMenuY As Integer
Dim MoveMenu1, MoveMenu2 As Boolean
Dim DrawX, DrawY As Integer
Dim DrawSub As Boolean
Dim STX, STY As Integer
Dim CurrentTile As DirectDrawSurface7
Dim CTDesc As DDSURFACEDESC2
Dim CTRECT As RECT
Dim SelectedRECT As RECT
Dim CtrlX, CtrlY, CtrlXTile, CtrlYTile As Integer
Dim ff, gg As Integer
Dim MousePressedOnClick As Boolean
Dim FullRECT As RECT
Dim key As DDCOLORKEY
Dim GroundFillTile, WallsFillTile As Integer
Public StartInput As Boolean

Dim WallsT, ind, sl As Integer
Dim MyConsole(1 To 5) As String
Public ConsoleInput As String
Public myindexes As Integer
Public MyIndex As Integer
Public CSearch As Integer
Public CodeFlag As Integer
Public SumOfAllTiles, NumberOfGTiles, NumberOfWTiles As Integer
Dim ThingsT As Integer
Dim ScrollRate As Integer
Dim DrawSTileX, DrawSTileY As Integer
Dim fx, fy, myg As Integer
Dim MaxT, tr, myrt, drf As Integer
Dim WTX, WTY, WTAS As Long
Dim Finished As Boolean




'Main Start - Load of Form
Public Sub Form1_Load()
EG.Checked = False
ScrollRate = 24
'calculating length of borders to draw...
Thing = 0

'some parameters..
NumberOfGTiles = 6
NumberOfWTiles = 18
MyConsole(1) = "quit"
MyConsole(2) = "go"
MyConsole(3) = "thru"
MyConsole(4) = "hehe"
MyConsole(5) = "hoho"



ConsoleInput = ""
'Default Filling tiles..
GroundFillTile = 2
WallsFillTile = 0

'Setting frame rate(default).....
CustomFrameRate = 40

'Not needed right now..
'UnitXPos = 320
'UnitYPos = 320




'Current Tileset X and Y's.. '''''
TileMenuX = 50
TileMenuY = 150
Marked = 1 'Default Loaded Tileset..
STX = 0
''''''''''''''''''''''''
'Default Loading Procedure.. ''''
For i = 0 To 500
For j = 0 To 500
GroundA(i, j) = GroundFillTile
WallsA(i, j) = WallsFillTile

Next
Next
''''''''''''''''''''''''''''

    'DoUnitMove = False
   'Show the main form
    Me.Show
    Me.WindowState = 2
    
    'Initialize DirectDraw
    Set mdd = mdx.DirectDrawCreate("")
  
    'Set the cooperative level (Fullscreen exclusive)
     mdd.SetCooperativeLevel frmMain.hWnd, DDSCL_FULLSCREEN Or DDSCL_EXCLUSIVE
    
    'Set the resolution
     mdd.SetDisplayMode 800, 600, 16, 0, DDSDM_DEFAULT
    
    'Describe the flipping chain architecture
    ddsdMain.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
    ddsdMain.lBackBufferCount = 1
    ddsdMain.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_FLIP Or DDSCAPS_COMPLEX
    
   
    'Create the primary surface
    Set msurfFront = mdd.CreateSurface(ddsdMain)
    
    'Create the backbuffer
    ddsdFlip.ddsCaps.lCaps = DDSCAPS_BACKBUFFER Or DDSCAPS_VIDEOMEMORY
    Set msurfBack = msurfFront.GetAttachedSurface(ddsdFlip.ddsCaps)
    
    'Setting Font for all text written using  msurfback.Drawtext.
     msurfBack.SetFont frmMain.Font
    'The font is now the same as in our main form (frmMain)...
   

   
    'Create our screen-sized rectangle
    mrectScreen.Bottom = SCREEN_HEIGHT
    mrectScreen.Right = SCREEN_WIDTH
    
    'Loading our surfaces.....
    LoadSurfaces
    CalcTLen
    'Proceeding to the main loop.....
    LoadFile "default.txt"
              MainLoop
End Sub

Private Sub MainLoop()

    'Starting the loop..
    mblnRunning = True
Do While mblnRunning
 
  'Frame rate calculations and settings...
  DelayTime = 1000 / CustomFrameRate
     If mdx.TickCount - LastTimeChecked >= DelayTime Then
        'Restore all surfaces if we loose 'em occasionally..
        If LostSurfaces Then LoadSurfaces
        msurfBack.BltColorFill mrectScreen, 0   'Clear the backbuffer
        MoveScreen                              'Move the screen (Scrolling)
        scrolling
                 UpdateMap
                 DrawTiles                               'Main drawing procedure..
           'Optional Unit Moving Procedure..
       'UnitFrames
      'If DoUnitMove = True Then
       'MoveUnit
       'End If
          
       'Flipping our main surface to the screen..
       msurfFront.Flip Nothing, 0
        DoEvents                                'Let other events occur
       LastTimeChecked = mdx.TickCount
       End If
       'Checking our display mode..
       If cmndlg1.CancelError = True Then
       Else
       End If
          ExclusiveMode
       Loop
        
        'Unload everything if someone has exited the main loop..
    Terminate

End Sub

Public Sub DrawTiles()



'Drawing Filling tiles first...
For j = Int(ViewY / TILE_HEIGHT) To Int(ViewY / TILE_HEIGHT) + 14
For i = Int(ViewX / TILE_WIDTH) To Int(ViewX / TILE_WIDTH) + 18
  
With GroundL1RECT
 .Bottom = TILE_WIDTH
 .Left = 0
 .Right = .Left + TILE_WIDTH
 .Top = 0
End With

'Calc X,Y coords for this tile's placement
CurX = i * TILE_WIDTH - ViewX
CurY = j * TILE_HEIGHT - ViewY
If GroundA(i, j) = 0 Then
msurfBack.BltFast CurX - (Int(GroundDesc(GroundA(i, j)).lWidth - TILE_WIDTH) / 2), CurY - (Int(GroundDesc(GroundA(i, j)).lHeight - TILE_WIDTH) / 2), GroundSurf(GroundFillTile), GroundL1RECT, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
Else
End If
 Next
Next

'First Ground Layer.'''''''''''''''''''''''''''''''''''''''''''''''
    'Draw the tiles according to the GroundA(x,y) array
For j = Int(ViewY / TILE_HEIGHT) To Int(ViewY / TILE_HEIGHT) + 14
For i = Int(ViewX / TILE_WIDTH) To Int(ViewX / TILE_WIDTH) + 18

With GroundL1RECT
 .Bottom = TILE_WIDTH
 .Left = 0
 .Right = .Left + TILE_WIDTH
 .Top = 0
End With

'Calc X,Y coords for this tile's placement
CurX = i * TILE_WIDTH - ViewX
CurY = j * TILE_HEIGHT - ViewY

msurfBack.BltFast CurX - (Int(GroundDesc(GroundA(i, j)).lWidth - TILE_WIDTH) / 2), CurY - (Int(GroundDesc(GroundA(i, j)).lHeight - TILE_WIDTH) / 2), GroundSurf(GroundA(i, j)), GroundL1RECT, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
 Next
Next
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''




'Walls Layer  ''''''''''''''''''''''''''''''''''''''''''''
'Draw the tiles according to the WallsA(x,y) array
For j = Int(ViewY / TILE_HEIGHT) To Int(ViewY / TILE_HEIGHT) + 14
 For i = Int(ViewX / TILE_WIDTH) To Int(ViewX / TILE_WIDTH) + 18
 
 With WallsRECT
 .Top = 0
 .Bottom = WallsDesc(WallsA(i, j)).lHeight
 .Left = 0
 .Right = WallsDesc(WallsA(i, j)).lWidth
 End With

CurX = i * TILE_WIDTH - ViewX
CurY = j * TILE_HEIGHT - ViewY
If WallsA(i, j) <> 0 Then
 msurfBack.BltFast CurX - (Int(WallsDesc(WallsA(i, j)).lWidth - TILE_WIDTH) / 2), CurY - (Int(WallsDesc(WallsA(i, j)).lHeight - TILE_WIDTH) / 2), WallsSurf(WallsA(i, j)), WallsRECT, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
  End If

'If UnitXPos > CurX And UnitXPos < CurX + 32 And UnitYPos > CurY And UnitYPos < CurY + 32 Then
'msurfBack.BltFast UnitXPos, UnitYPos, MyUnit, UnitRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
 ' End If

   Next
Next
If Num > 0 Then
For myg = 0 To Num
If Wthing(myg).WorldY > ViewY And Wthing(myg).WorldY < ViewY + 600 And Wthing(myg).WorldX > ViewX And Wthing(myg).WorldX < ViewX + 800 Then
With ThingsRECT
.Right = ThingsDesc(Wthing(myg).ThingAs).lWidth
.Bottom = ThingsDesc(Wthing(myg).ThingAs).lHeight
End With
msurfBack.BltFast Wthing(myg).WorldX - ViewX, Wthing(myg).WorldY - ViewY, ThingsSurf(Wthing(myg).ThingAs), ThingsRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
End If
Next
End If


'Unit Drawing............
' With UnitRECT
'.Left = Frame * 64
'.Right = .Left + 64
'.Top = CurrentPos * 64
'.Bottom = .Top + 64
'End With

'FrameIndex = FrameIndex + 1
'If FrameIndex = 9 Then
'FrameIndex = 0
'End If

'Menu Drawing...


Select Case Marked
'Floor tiles have been selected, so drawing them in menu box..
Case Is = 1
For SubIndex = 0 To NumberOfGTiles
With MenuRECT
.Left = 0
.Top = 0
.Right = GroundDesc(SubIndex).lWidth
.Bottom = GroundDesc(SubIndex).lHeight
End With
If SubIndex = STX Then
DrawSTileX = Int((GetX + ViewX) / TILE_WIDTH) * TILE_WIDTH - ViewX
DrawSTileY = Int((GetY + ViewY) / TILE_HEIGHT) * TILE_HEIGHT - ViewY
msurfBack.BltFast DrawSTileX, DrawSTileY, GroundSurf(SubIndex), SelectedRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
End If
msurfBack.BltFast TileMenuX + SubIndex * TILE_WIDTH, TileMenuY, GroundSurf(SubIndex), MenuRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
Next

'Walls tiles have been selected.. Drawing them? Sure..
Case Is = 2
DrawSTileX = Int((GetX + ViewX) / TILE_WIDTH) * TILE_WIDTH - ViewX
DrawSTileY = Int((GetY + ViewY) / TILE_HEIGHT) * TILE_HEIGHT - ViewY
For SubIndex = 0 To NumberOfWTiles
With WallsRECT
.Left = 0
.Top = 0
.Right = WallsDesc(SubIndex).lWidth
.Bottom = WallsDesc(SubIndex).lHeight
End With
If SubIndex = STX Then
msurfBack.BltFast DrawSTileX, DrawSTileY, WallsSurf(SubIndex), WallsRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
End If
msurfBack.BltFast TileMenuX + SubIndex * TILE_WIDTH, TileMenuY, WallsSurf(SubIndex), MenuRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY

Next
End Select


'Drawing all Things to a map..


'Drawing selected thing to map..
MyThingRECT.Right = ThingsDesc(Thing).lWidth
MyThingRECT.Bottom = ThingsDesc(Thing).lHeight
If Marked = 3 Then
msurfBack.BltFast GetX, GetY, ThingsSurf(Thing), MyThingRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
End If
'drawing MyConsole command...

msurfBack.DrawText 100, 100, ViewX, False
msurfBack.DrawText 200, 100, ViewY, False


'Borders...........................
msurfBack.DrawText TileMenuX, TileMenuY - 20, "Your current Tileset. Double click to select tile.", False

msurfBack.DrawBox TileMenuX - 2, TileMenuY - 2, 2 + TileMenuX + SumOfAllTiles, TileMenuY + TILE_HEIGHT + 2


'Calling mouse procedure for ... Hmm.. Check it to see urself..
Call DrawGrid
Call Do_Mouse

End Sub

Private Sub MoveScreen()

    'Move screen
    'Ensure we don't go off the edge, that'd cause an error!
    If ViewX < 0 Then ViewX = 0
    If ViewX > (UBound(GroundA, 1) - 18) * TILE_WIDTH Then ViewX = (UBound(GroundA, 1) - 18) * TILE_WIDTH
    If ViewY < 0 Then ViewY = 0
    If ViewY > (UBound(GroundA, 2) - 16) * TILE_HEIGHT Then ViewY = (UBound(GroundA, 2) - 16) * TILE_HEIGHT

End Sub

Private Sub AA_Click()
Marked = 3
Thing = 15
End Sub

Private Sub ABT_Click()
Marked = 3
Thing = 14
End Sub

Private Sub ACCT1_Click()
Marked = 3
Thing = 13
End Sub

Private Sub ACCT2_Click()
Marked = 3
Thing = 12
End Sub

Private Sub AOCT1_Click()
Marked = 3
Thing = 11

End Sub

Private Sub AOCT2_Click()
Marked = 3
Thing = 10
End Sub

Private Sub ARC_Click()
Thing = 19
Marked = 3
End Sub

Private Sub ARS_Click()
Marked = 3
Thing = 18
End Sub

Private Sub ART_Click()
Marked = 3
Thing = 9
End Sub

Private Sub AST1_Click()
Marked = 3
Thing = 1
End Sub

Private Sub AST1L_Click()
Marked = 3
Thing = 2
End Sub

Private Sub AST1R_Click()
Marked = 3
Thing = 3
End Sub

Private Sub AST2L_Click()
Marked = 3
Thing = 4
End Sub

Private Sub AST2R_Click()
Marked = 3
Thing = 5
End Sub

Private Sub AW_Click()
Marked = 3
Thing = 6
End Sub

Private Sub AWind_Click()
Marked = 3
Thing = 7
End Sub

Private Sub CS_Click()
Marked = 3
Thing = 8
End Sub

Private Sub DT_Click()
Marked = 3
Thing = 0
End Sub

Private Sub E_Click()
Terminate
End
End Sub

Private Sub EG_Click()
If EG.Checked = False Then
EG.Checked = True
Else
If EG.Checked = True Then
EG.Checked = False
End If
End If
End Sub

Private Sub Form_DblClick()
DrawX = GetX
DrawY = GetY
'Define whether you clicked on main field or tileset..
If DrawX > TileMenuX And DrawX < TileMenuX + 20 * TILE_WIDTH And DrawY > TileMenuY And DrawY < TileMenuY + TILE_WIDTH Then
STX = Int((DrawX - TileMenuX) / TILE_WIDTH)
STY = Int((DrawY - TileMenuY) / TILE_WIDTH)
DrawSub = True
Else
DrawSub = False
End If
End Sub

Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   'Exit program on escape key
'If KeyCode = vbKeyUp Then
 '    UpKey = True
  '  NextPos = 1
   ' End If
'If KeyCode = vbKeyPageUp Then
 '    UpRightKey = True
  '  NextPos = 2
   ' End If
'If KeyCode = vbKeyRight Then
 '    RightKey = True
  '  NextPos = 3
   ' End If
'If KeyCode = vbKeyPageDown Then
 '    DownRightKey = True
  '  NextPos = 4
   ' End If
'If KeyCode = vbKeyDown Then
 '    DownKey = True
  '   NextPos = 5
   ' End If
'If KeyCode = vbKeyEnd Then
 '    DownLeftKey = True
  '  NextPos = 6
   ' End If
'If KeyCode = vbKeyLeft Then
 '    LeftKey = True
  '  NextPos = 7
  '  End If
  'If KeyCode = vbKeyHome Then
   'UpLeftKey = True
  'NextPos = 0
  'End If
  
  'DoUnitMove = True
   'FinalPos = NextPos
    If KeyCode = vbKeyEscape Then mblnRunning = False

  ' If Frame = 2 Then Frame = 0

End Sub

Public Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Terminate
'MyConsole commands..MyConsole(0) = "exit"
If KeyAscii = 92 Then
ConsoleInput = InputBox("Enter your command", "Console Prompt", "", 100, 100)
End If
'If KeyAscii = 115 Then
'msurfBack.ReleaseDC Picture1.hDC
'SavePicture Picture1.Picture, "c:/games/Mypic.bmp"
'End If
For MyIndex = 1 To 5
If MyConsole(MyIndex) = ConsoleInput Then
Finded = True
Else
Finded = False
End If
Next

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
DoUnitMove = False
    'Stop moving screen
    If KeyCode = vbKeyUp Then UpKey = False
    If KeyCode = vbKeyDown Then DownKey = False
    If KeyCode = vbKeyLeft Then LeftKey = False
    If KeyCode = vbKeyRight Then RightKey = False
    If KeyCode = vbKeyHome Then UpLeftKey = False
    If KeyCode = vbKeyPageUp Then UpRightKey = False
    If KeyCode = vbKeyPageDown Then DownRightKey = False
    If KeyCode = vbKeyEnd Then DownLeftKey = False
'Frame = 2
End Sub
Private Sub LoadSurfaces()
'Loading our resources file..
Open App.Path & "\resources.txt" For Input As #3
Input #3, CustomFrameRate
Close #3

'Setting colour key for our surfaces.. They all have a black colour background..
key.low = 0
key.high = 0
msurfBack.SetColorKey DDCKEY_SRCBLT, key
''''''''''''''''''''''''''''''''''''''

'mouse description and..
ddsd5.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
ddsd5.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_VIDEOMEMORY
ddsd5.lWidth = 32
ddsd5.lHeight = 32
'mouse surface....
Set MouseSurf = mdd.CreateSurfaceFromFile(App.Path & "\bmps\mouse.bmp", ddsd5) 'Create "MouseSurf" from the mouse.bmp file.
MouseSurf.SetColorKey DDCKEY_SRCBLT, key 'Applies the colour key to the mouse surface.
''''''''''''''''''''''''''''''''''''''


'Ground properties..''''''''''''''''''''''''''''''''''''''''''''''''
'First Ground Layer... '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
GroundDesc(0).ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
GroundDesc(0).lFlags = DDSD_CKSRCBLT Or DDSD_HEIGHT Or DDSD_WIDTH
GroundDesc(0).lWidth = TILE_WIDTH
GroundDesc(0).lHeight = TILE_WIDTH
Set GroundSurf(0) = mdd.CreateSurfaceFromFile(App.Path & "\bmps\default.bmp", GroundDesc(0))
GroundSurf(0).SetColorKey DDCKEY_SRCBLT, key

GroundDesc(1).ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
GroundDesc(1).lFlags = DDSD_CKSRCBLT Or DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
GroundDesc(1).lWidth = TILE_WIDTH
GroundDesc(1).lHeight = TILE_WIDTH
Set GroundSurf(1) = mdd.CreateSurfaceFromFile(App.Path & "\bmps\ground1.bmp", GroundDesc(1))
GroundSurf(1).SetColorKey DDCKEY_SRCBLT, key

GroundDesc(2).ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
GroundDesc(2).lFlags = DDSD_CKSRCBLT Or DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
GroundDesc(2).lWidth = TILE_WIDTH
GroundDesc(2).lHeight = TILE_WIDTH
Set GroundSurf(2) = mdd.CreateSurfaceFromFile(App.Path & "\bmps\ground2.bmp", GroundDesc(2))
GroundSurf(2).SetColorKey DDCKEY_SRCBLT, key

GroundDesc(3).ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
GroundDesc(3).lFlags = DDSD_CKSRCBLT Or DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
GroundDesc(3).lWidth = TILE_WIDTH
GroundDesc(3).lHeight = TILE_WIDTH
Set GroundSurf(3) = mdd.CreateSurfaceFromFile(App.Path & "\bmps\ground3.bmp", GroundDesc(3))
GroundSurf(3).SetColorKey DDCKEY_SRCBLT, key

GroundDesc(4).ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
GroundDesc(4).lFlags = DDSD_CKSRCBLT Or DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
GroundDesc(4).lWidth = TILE_WIDTH
GroundDesc(4).lHeight = TILE_WIDTH
Set GroundSurf(4) = mdd.CreateSurfaceFromFile(App.Path & "\bmps\ground4.bmp", GroundDesc(4))
GroundSurf(4).SetColorKey DDCKEY_SRCBLT, key

GroundDesc(5).ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
GroundDesc(5).lFlags = DDSD_CKSRCBLT Or DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
GroundDesc(5).lWidth = TILE_WIDTH
GroundDesc(5).lHeight = TILE_WIDTH
Set GroundSurf(5) = mdd.CreateSurfaceFromFile(App.Path & "\bmps\Ground5.bmp", GroundDesc(5))
GroundSurf(5).SetColorKey DDCKEY_SRCBLT, key

GroundDesc(6).ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
GroundDesc(6).lFlags = DDSD_CKSRCBLT Or DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
GroundDesc(6).lWidth = TILE_WIDTH
GroundDesc(6).lHeight = TILE_WIDTH
Set GroundSurf(6) = mdd.CreateSurfaceFromFile(App.Path & "\bmps\ground6.bmp", GroundDesc(6))
GroundSurf(6).SetColorKey DDCKEY_SRCBLT, key
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 

'Walls Properties Layer 1..'''''''''''''''''''''''''''''''''''''''''''''''''''
WallsDesc(0).ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
WallsDesc(0).lFlags = DDSD_CAPS Or DDSD_CKSRCBLT Or DDSD_HEIGHT Or DDSD_WIDTH
WallsDesc(0).lHeight = TILE_WIDTH
WallsDesc(0).lWidth = TILE_WIDTH
Set WallsSurf(0) = mdd.CreateSurfaceFromFile(App.Path & "\bmps\default.bmp", WallsDesc(0))
WallsSurf(0).SetColorKey DDCKEY_SRCBLT, key

WallsDesc(1).ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
WallsDesc(1).lFlags = DDSD_CAPS Or DDSD_CKSRCBLT Or DDSD_HEIGHT Or DDSD_WIDTH
WallsDesc(1).lHeight = TILE_WIDTH
WallsDesc(1).lWidth = TILE_WIDTH
Set WallsSurf(1) = mdd.CreateSurfaceFromFile(App.Path & "\bmps\wall1(1).bmp", WallsDesc(1))
WallsSurf(1).SetColorKey DDCKEY_SRCBLT, key

WallsDesc(2).ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
WallsDesc(2).lFlags = DDSD_CAPS Or DDSD_CKSRCBLT Or DDSD_HEIGHT Or DDSD_WIDTH
WallsDesc(2).lHeight = TILE_WIDTH
WallsDesc(2).lWidth = TILE_WIDTH
Set WallsSurf(2) = mdd.CreateSurfaceFromFile(App.Path & "\bmps\wall2(1).bmp", WallsDesc(2))
WallsSurf(2).SetColorKey DDCKEY_SRCBLT, key

WallsDesc(3).ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
WallsDesc(3).lFlags = DDSD_CAPS Or DDSD_CKSRCBLT Or DDSD_HEIGHT Or DDSD_WIDTH
WallsDesc(3).lHeight = TILE_WIDTH
WallsDesc(3).lWidth = TILE_WIDTH
Set WallsSurf(3) = mdd.CreateSurfaceFromFile(App.Path & "\bmps\wall3(1).bmp", WallsDesc(3))
WallsSurf(3).SetColorKey DDCKEY_SRCBLT, key

'walls1-3(1)

WallsDesc(4).ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
WallsDesc(4).lFlags = DDSD_CAPS Or DDSD_CKSRCBLT Or DDSD_HEIGHT Or DDSD_WIDTH
WallsDesc(4).lHeight = TILE_WIDTH
WallsDesc(4).lWidth = TILE_WIDTH
Set WallsSurf(4) = mdd.CreateSurfaceFromFile(App.Path & "\bmps\wall1(2).bmp", WallsDesc(4))
WallsSurf(4).SetColorKey DDCKEY_SRCBLT, key

WallsDesc(5).ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
WallsDesc(5).lFlags = DDSD_CAPS Or DDSD_CKSRCBLT Or DDSD_HEIGHT Or DDSD_WIDTH
WallsDesc(5).lHeight = TILE_WIDTH
WallsDesc(5).lWidth = TILE_WIDTH
Set WallsSurf(5) = mdd.CreateSurfaceFromFile(App.Path & "\bmps\wall2(2).bmp", WallsDesc(5))
WallsSurf(5).SetColorKey DDCKEY_SRCBLT, key

WallsDesc(6).ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
WallsDesc(6).lFlags = DDSD_CAPS Or DDSD_CKSRCBLT Or DDSD_HEIGHT Or DDSD_WIDTH
WallsDesc(6).lHeight = TILE_WIDTH
WallsDesc(6).lWidth = TILE_WIDTH
Set WallsSurf(6) = mdd.CreateSurfaceFromFile(App.Path & "\bmps\wall3(2).bmp", WallsDesc(6))
WallsSurf(6).SetColorKey DDCKEY_SRCBLT, key

'walls1-3(2)


WallsDesc(7).ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
WallsDesc(7).lFlags = DDSD_CAPS Or DDSD_CKSRCBLT Or DDSD_HEIGHT Or DDSD_WIDTH
WallsDesc(7).lHeight = TILE_WIDTH
WallsDesc(7).lWidth = TILE_WIDTH
Set WallsSurf(7) = mdd.CreateSurfaceFromFile(App.Path & "\bmps\wall1(3).bmp", WallsDesc(7))
WallsSurf(7).SetColorKey DDCKEY_SRCBLT, key

WallsDesc(8).ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
WallsDesc(8).lFlags = DDSD_CAPS Or DDSD_CKSRCBLT Or DDSD_HEIGHT Or DDSD_WIDTH
WallsDesc(8).lHeight = TILE_WIDTH
WallsDesc(8).lWidth = TILE_WIDTH
Set WallsSurf(8) = mdd.CreateSurfaceFromFile(App.Path & "\bmps\wall2(3).bmp", WallsDesc(8))
WallsSurf(8).SetColorKey DDCKEY_SRCBLT, key

WallsDesc(9).ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
WallsDesc(9).lFlags = DDSD_CAPS Or DDSD_CKSRCBLT Or DDSD_HEIGHT Or DDSD_WIDTH
WallsDesc(9).lHeight = TILE_WIDTH
WallsDesc(9).lWidth = TILE_WIDTH
Set WallsSurf(9) = mdd.CreateSurfaceFromFile(App.Path & "\bmps\wall3(3).bmp", WallsDesc(9))
WallsSurf(9).SetColorKey DDCKEY_SRCBLT, key

'walls 1-3(3)

WallsDesc(10).ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
WallsDesc(10).lFlags = DDSD_CAPS Or DDSD_CKSRCBLT Or DDSD_HEIGHT Or DDSD_WIDTH
WallsDesc(10).lHeight = TILE_WIDTH
WallsDesc(10).lWidth = TILE_WIDTH
Set WallsSurf(10) = mdd.CreateSurfaceFromFile(App.Path & "\bmps\wall1(4).bmp", WallsDesc(10))
WallsSurf(10).SetColorKey DDCKEY_SRCBLT, key

WallsDesc(11).ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
WallsDesc(11).lFlags = DDSD_CAPS Or DDSD_CKSRCBLT Or DDSD_HEIGHT Or DDSD_WIDTH
WallsDesc(11).lHeight = TILE_WIDTH
WallsDesc(11).lWidth = TILE_WIDTH
Set WallsSurf(11) = mdd.CreateSurfaceFromFile(App.Path & "\bmps\wall2(4).bmp", WallsDesc(11))
WallsSurf(11).SetColorKey DDCKEY_SRCBLT, key

WallsDesc(12).ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
WallsDesc(12).lFlags = DDSD_CAPS Or DDSD_CKSRCBLT Or DDSD_HEIGHT Or DDSD_WIDTH
WallsDesc(12).lHeight = TILE_WIDTH
WallsDesc(12).lWidth = TILE_WIDTH
Set WallsSurf(12) = mdd.CreateSurfaceFromFile(App.Path & "\bmps\wall3(4).bmp", WallsDesc(12))
WallsSurf(12).SetColorKey DDCKEY_SRCBLT, key

'walls 1-3(4)

WallsDesc(13).ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
WallsDesc(13).lFlags = DDSD_CAPS Or DDSD_CKSRCBLT Or DDSD_HEIGHT Or DDSD_WIDTH
WallsDesc(13).lHeight = TILE_WIDTH
WallsDesc(13).lWidth = TILE_WIDTH
Set WallsSurf(13) = mdd.CreateSurfaceFromFile(App.Path & "\bmps\wall1(5).bmp", WallsDesc(13))
WallsSurf(13).SetColorKey DDCKEY_SRCBLT, key

WallsDesc(14).ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
WallsDesc(14).lFlags = DDSD_CAPS Or DDSD_CKSRCBLT Or DDSD_HEIGHT Or DDSD_WIDTH
WallsDesc(14).lHeight = TILE_WIDTH
WallsDesc(14).lWidth = TILE_WIDTH
Set WallsSurf(14) = mdd.CreateSurfaceFromFile(App.Path & "\bmps\wall2(5).bmp", WallsDesc(14))
WallsSurf(14).SetColorKey DDCKEY_SRCBLT, key

WallsDesc(15).ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
WallsDesc(15).lFlags = DDSD_CAPS Or DDSD_CKSRCBLT Or DDSD_HEIGHT Or DDSD_WIDTH
WallsDesc(15).lHeight = TILE_WIDTH
WallsDesc(15).lWidth = TILE_WIDTH
Set WallsSurf(15) = mdd.CreateSurfaceFromFile(App.Path & "\bmps\wall3(5).bmp", WallsDesc(15))
WallsSurf(15).SetColorKey DDCKEY_SRCBLT, key

WallsDesc(16).ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
WallsDesc(16).lFlags = DDSD_CAPS Or DDSD_CKSRCBLT Or DDSD_HEIGHT Or DDSD_WIDTH
WallsDesc(16).lHeight = TILE_WIDTH
WallsDesc(16).lWidth = 64
Set WallsSurf(16) = mdd.CreateSurfaceFromFile(App.Path & "\bmps\wall1(6).bmp", WallsDesc(16))
WallsSurf(16).SetColorKey DDCKEY_SRCBLT, key

WallsDesc(17).ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
WallsDesc(17).lFlags = DDSD_CAPS Or DDSD_CKSRCBLT Or DDSD_HEIGHT Or DDSD_WIDTH
WallsDesc(17).lHeight = TILE_WIDTH
WallsDesc(17).lWidth = 64
Set WallsSurf(17) = mdd.CreateSurfaceFromFile(App.Path & "\bmps\wall2(6).bmp", WallsDesc(17))
WallsSurf(17).SetColorKey DDCKEY_SRCBLT, key

WallsDesc(18).ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
WallsDesc(18).lFlags = DDSD_CAPS Or DDSD_CKSRCBLT Or DDSD_HEIGHT Or DDSD_WIDTH
WallsDesc(18).lHeight = TILE_WIDTH
WallsDesc(18).lWidth = 64
Set WallsSurf(18) = mdd.CreateSurfaceFromFile(App.Path & "\bmps\wall3(6).bmp", WallsDesc(18))
WallsSurf(18).SetColorKey DDCKEY_SRCBLT, key

ThingsDesc(0).ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
ThingsDesc(0).lFlags = DDSD_CAPS Or DDSD_CKSRCBLT Or DDSD_HEIGHT Or DDSD_WIDTH
ThingsDesc(0).lHeight = 48
ThingsDesc(0).lWidth = 48
Set ThingsSurf(0) = mdd.CreateSurfaceFromFile(App.Path & "\bmps\default.bmp", ThingsDesc(0))
ThingsSurf(0).SetColorKey DDCKEY_SRCBLT, key

ThingsDesc(1).ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
ThingsDesc(1).lFlags = DDSD_CAPS Or DDSD_CKSRCBLT Or DDSD_HEIGHT Or DDSD_WIDTH
ThingsDesc(1).lHeight = 64
ThingsDesc(1).lWidth = 32
Set ThingsSurf(1) = mdd.CreateSurfaceFromFile(App.Path & "\bmps\statuetype0.bmp", ThingsDesc(1))
ThingsSurf(1).SetColorKey DDCKEY_SRCBLT, key

ThingsDesc(2).ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
ThingsDesc(2).lFlags = DDSD_CAPS Or DDSD_CKSRCBLT Or DDSD_HEIGHT Or DDSD_WIDTH
ThingsDesc(2).lHeight = 64
ThingsDesc(2).lWidth = 32
Set ThingsSurf(2) = mdd.CreateSurfaceFromFile(App.Path & "\bmps\statuetype1left.bmp", ThingsDesc(2))
ThingsSurf(2).SetColorKey DDCKEY_SRCBLT, key

ThingsDesc(3).ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
ThingsDesc(3).lFlags = DDSD_CAPS Or DDSD_CKSRCBLT Or DDSD_HEIGHT Or DDSD_WIDTH
ThingsDesc(3).lHeight = 64
ThingsDesc(3).lWidth = 32
Set ThingsSurf(3) = mdd.CreateSurfaceFromFile(App.Path & "\bmps\statuetype1right.bmp", ThingsDesc(3))
ThingsSurf(3).SetColorKey DDCKEY_SRCBLT, key

ThingsDesc(4).ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
ThingsDesc(4).lFlags = DDSD_CAPS Or DDSD_CKSRCBLT Or DDSD_HEIGHT Or DDSD_WIDTH
ThingsDesc(4).lHeight = 64
ThingsDesc(4).lWidth = 32
Set ThingsSurf(4) = mdd.CreateSurfaceFromFile(App.Path & "\bmps\statuetype2left.bmp", ThingsDesc(4))
ThingsSurf(4).SetColorKey DDCKEY_SRCBLT, key

ThingsDesc(5).ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
ThingsDesc(5).lFlags = DDSD_CAPS Or DDSD_CKSRCBLT Or DDSD_HEIGHT Or DDSD_WIDTH
ThingsDesc(5).lHeight = 64
ThingsDesc(5).lWidth = 32
Set ThingsSurf(5) = mdd.CreateSurfaceFromFile(App.Path & "\bmps\statuetype2right.bmp", ThingsDesc(5))
ThingsSurf(5).SetColorKey DDCKEY_SRCBLT, key

ThingsDesc(6).ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
ThingsDesc(6).lFlags = DDSD_CAPS Or DDSD_CKSRCBLT Or DDSD_HEIGHT Or DDSD_WIDTH
ThingsDesc(6).lHeight = 32
ThingsDesc(6).lWidth = 32
Set ThingsSurf(6) = mdd.CreateSurfaceFromFile(App.Path & "\bmps\well.bmp", ThingsDesc(6))
ThingsSurf(6).SetColorKey DDCKEY_SRCBLT, key

ThingsDesc(7).ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
ThingsDesc(7).lFlags = DDSD_CAPS Or DDSD_CKSRCBLT Or DDSD_HEIGHT Or DDSD_WIDTH
ThingsDesc(7).lHeight = 32
ThingsDesc(7).lWidth = 32
Set ThingsSurf(7) = mdd.CreateSurfaceFromFile(App.Path & "\bmps\window.bmp", ThingsDesc(7))
ThingsSurf(7).SetColorKey DDCKEY_SRCBLT, key

ThingsDesc(8).ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
ThingsDesc(8).lFlags = DDSD_CAPS Or DDSD_CKSRCBLT Or DDSD_HEIGHT Or DDSD_WIDTH
ThingsDesc(8).lHeight = 32
ThingsDesc(8).lWidth = 32
Set ThingsSurf(8) = mdd.CreateSurfaceFromFile(App.Path & "\bmps\crossedswords.bmp", ThingsDesc(8))
ThingsSurf(8).SetColorKey DDCKEY_SRCBLT, key

ThingsDesc(9).ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
ThingsDesc(9).lFlags = DDSD_CAPS Or DDSD_CKSRCBLT Or DDSD_HEIGHT Or DDSD_WIDTH
ThingsDesc(9).lHeight = 32
ThingsDesc(9).lWidth = 32
Set ThingsSurf(9) = mdd.CreateSurfaceFromFile(App.Path & "\bmps\redtorch.bmp", ThingsDesc(9))
ThingsSurf(9).SetColorKey DDCKEY_SRCBLT, key

ThingsDesc(10).ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
ThingsDesc(10).lFlags = DDSD_CAPS Or DDSD_CKSRCBLT Or DDSD_HEIGHT Or DDSD_WIDTH
ThingsDesc(10).lHeight = 32
ThingsDesc(10).lWidth = 32
Set ThingsSurf(10) = mdd.CreateSurfaceFromFile(App.Path & "\bmps\chestopened2.bmp", ThingsDesc(10))
ThingsSurf(10).SetColorKey DDCKEY_SRCBLT, key

ThingsDesc(11).ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
ThingsDesc(11).lFlags = DDSD_CAPS Or DDSD_CKSRCBLT Or DDSD_HEIGHT Or DDSD_WIDTH
ThingsDesc(11).lHeight = 32
ThingsDesc(11).lWidth = 32
Set ThingsSurf(11) = mdd.CreateSurfaceFromFile(App.Path & "\bmps\chestopened1.bmp", ThingsDesc(11))
ThingsSurf(11).SetColorKey DDCKEY_SRCBLT, key

ThingsDesc(12).ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
ThingsDesc(12).lFlags = DDSD_CAPS Or DDSD_CKSRCBLT Or DDSD_HEIGHT Or DDSD_WIDTH
ThingsDesc(12).lHeight = 32
ThingsDesc(12).lWidth = 32
Set ThingsSurf(12) = mdd.CreateSurfaceFromFile(App.Path & "\bmps\chestclosed2.bmp", ThingsDesc(12))
ThingsSurf(12).SetColorKey DDCKEY_SRCBLT, key

ThingsDesc(13).ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
ThingsDesc(13).lFlags = DDSD_CAPS Or DDSD_CKSRCBLT Or DDSD_HEIGHT Or DDSD_WIDTH
ThingsDesc(13).lHeight = 32
ThingsDesc(13).lWidth = 32
Set ThingsSurf(13) = mdd.CreateSurfaceFromFile(App.Path & "\bmps\chestclosed1.bmp", ThingsDesc(13))
ThingsSurf(13).SetColorKey DDCKEY_SRCBLT, key

ThingsDesc(14).ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
ThingsDesc(14).lFlags = DDSD_CAPS Or DDSD_CKSRCBLT Or DDSD_HEIGHT Or DDSD_WIDTH
ThingsDesc(14).lHeight = 32
ThingsDesc(14).lWidth = 32
Set ThingsSurf(14) = mdd.CreateSurfaceFromFile(App.Path & "\bmps\bluetorch.bmp", ThingsDesc(14))
ThingsSurf(14).SetColorKey DDCKEY_SRCBLT, key

ThingsDesc(15).ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
ThingsDesc(15).lFlags = DDSD_CAPS Or DDSD_CKSRCBLT Or DDSD_HEIGHT Or DDSD_WIDTH
ThingsDesc(15).lHeight = 32
ThingsDesc(15).lWidth = 32
Set ThingsSurf(15) = mdd.CreateSurfaceFromFile(App.Path & "\bmps\an angel.bmp", ThingsDesc(15))
ThingsSurf(15).SetColorKey DDCKEY_SRCBLT, key

ThingsDesc(16).ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
ThingsDesc(16).lFlags = DDSD_CAPS Or DDSD_CKSRCBLT Or DDSD_HEIGHT Or DDSD_WIDTH
ThingsDesc(16).lHeight = 32
ThingsDesc(16).lWidth = 32
Set ThingsSurf(16) = mdd.CreateSurfaceFromFile(App.Path & "\bmps\RIPS.bmp", ThingsDesc(16))
ThingsSurf(16).SetColorKey DDCKEY_SRCBLT, key

ThingsDesc(17).ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
ThingsDesc(17).lFlags = DDSD_CAPS Or DDSD_CKSRCBLT Or DDSD_HEIGHT Or DDSD_WIDTH
ThingsDesc(17).lHeight = 32
ThingsDesc(17).lWidth = 32
Set ThingsSurf(17) = mdd.CreateSurfaceFromFile(App.Path & "\bmps\RIPC.bmp", ThingsDesc(17))
ThingsSurf(17).SetColorKey DDCKEY_SRCBLT, key

ThingsDesc(18).ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
ThingsDesc(18).lFlags = DDSD_CAPS Or DDSD_CKSRCBLT Or DDSD_HEIGHT Or DDSD_WIDTH
ThingsDesc(18).lHeight = 32
ThingsDesc(18).lWidth = 32
Set ThingsSurf(18) = mdd.CreateSurfaceFromFile(App.Path & "\bmps\redshield.bmp", ThingsDesc(18))
ThingsSurf(18).SetColorKey DDCKEY_SRCBLT, key

ThingsDesc(19).ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
ThingsDesc(19).lFlags = DDSD_CAPS Or DDSD_CKSRCBLT Or DDSD_HEIGHT Or DDSD_WIDTH
ThingsDesc(19).lHeight = 32
ThingsDesc(19).lWidth = 32
Set ThingsSurf(19) = mdd.CreateSurfaceFromFile(App.Path & "\bmps\red cirtains.bmp", ThingsDesc(19))
ThingsSurf(19).SetColorKey DDCKEY_SRCBLT, key

'unit properties.. '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Set MyUnit = mdd.CreateSurfaceFromFile(App.Path & "\bowmansetwalking.bmp", MyUnitDesc)
 '  MyUnitDesc.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
  ' MyUnitDesc.lFlags = DDSD_CAPS Or DDSD_CKSRCBLT
   'MyUnitDesc.lWidth = 192
   'MyUnitDesc.lHeight = 512
   'MyUnit.SetColorKey DDCKEY_SRCBLT, key
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Current Selected Tile.. Default - 1..''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
CTDesc.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_VIDEOMEMORY
CTDesc.lHeight = TILE_WIDTH
CTDesc.lWidth = TILE_WIDTH
CTDesc.lFlags = DDSD_CAPS Or DDSD_CKSRCBLT Or DDSD_WIDTH Or DDSD_HEIGHT
End Sub

Private Function ExclusiveMode() As Boolean

Dim lngTestExMode As Long

'Testing the Exclusive mode..
lngTestExMode = mdd.TestCooperativeLevel
If (lngTestExMode = DD_OK) Then
  ExclusiveMode = True
    Else
        ExclusiveMode = False
End If
End Function

Public Function LostSurfaces() As Boolean
LostSurfaces = False
Do Until ExclusiveMode
    DoEvents
    LostSurfaces = True
Loop
    
 DoEvents

If LostSurfaces Then
 mdd.RestoreAllSurfaces
End If
End Function

Public Sub Terminate()
frmMain.TS = True
    'Terminate the render loop

    
    'Restore resolution
  Call mdd.RestoreDisplayMode 'sets the screen resolution back to what it was before the program was started.
     Call mdd.SetCooperativeLevel(Me.hWnd, DDSCL_NORMAL) 'tells DirectX that the application is no longer
Set msurfBack = Nothing
Set msurfFront = Nothing
Set MouseSurf = Nothing
    'Kill directdraw
    Set mdd = Nothing
For STX = 0 To 6
  Set GroundSurf(STX) = Nothing
 Next
 For SubIndex = 0 To 18
 Set WallsSurf(SubIndex) = Nothing
Next
    
    'Unload the form
 Unload Me
End
End Sub

Public Sub scrolling()
'right scrolling..
If GetX = 799 Then
  If ViewX < TILE_WIDTH * i Then
    ViewX = ViewX + ScrollRate
            If ViewX >= TILE_WIDTH * i Then ViewX = TILE_WIDTH * i

  End If
End If

'left scrolling..
If GetX = 0 Then
  If ViewX > 0 Then
    ViewX = ViewX - ScrollRate
      End If
   If ViewX < 0 Then ViewX = 0
End If

'top scrolling..
If GetY = 0 Then
   If ViewY > 0 Then
     ViewY = ViewY - ScrollRate
    End If
     If ViewY <= 0 Then ViewY = 0
End If

'bottom scrolling..
If GetY = SCREEN_HEIGHT - 1 Then
     If ViewY < TILE_HEIGHT * j Then
       ViewY = ViewY + ScrollRate
         End If
       If ViewY >= TILE_HEIGHT * j Then ViewY = TILE_HEIGHT * j
End If
End Sub

Private Sub Form_Load()
ReDim Wthing(0) As Things
Set Wthing(0) = New Things
Wthing(0).ThingAs = 0
Num = 0
ViewX = 48 * 500 / 2
ViewY = 48 * 500 / 2
Form1_Load
End Sub



Public Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
If Marked = 3 And Thing <> 0 Then
Num = Num + 1
ReDim Preserve Wthing(0 To Num) As Things
Set Wthing(Num) = New Things
Wthing(Num).WorldX = GetX + ViewX
Wthing(Num).WorldY = GetY + ViewY
Wthing(Num).ThingAs = Thing
Else
End If
End If
 If GetX > TileMenuX And GetX < TileMenuX + 48 * 7 And GetY > TileMenuY And GetY < TileMenuY + TILE_HEIGHT Then
 CtrlX = x - TileMenuX
 CtrlY = y - TileMenuY
 MoveMenu1 = True
Else
MousePressedOnClick = True
End If

End Sub

Sub Do_Mouse()
 'mouse image(cursor)..
 Dim Rval As Long    'to store the return value
 Dim rMouse As RECT
 rMouse.Top = 0
 rMouse.Left = 0
 rMouse.Right = 32
 rMouse.Bottom = 32
' Rval = msurfBack.BltFast(GetX - 5, GetY - 5, MouseSurf, rMouse, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT) 'IMPORTANT: ALL THIS NEEDS TO BE ON THE SAME LINE. When you paste it into VB,
 'make sure it all stays on the same line.
     End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 And Thing = 0 Then
RemoveTile
End If
If Button = 1 And MoveMenu1 = True Then
TileMenuX = x - CtrlX
TileMenuY = y - CtrlY
Else
If MousePressedOnClick = True And Button = 1 Then
ff = Int((GetX + ViewX) / TILE_WIDTH)
gg = Int((GetY + ViewY) / TILE_WIDTH)
Select Case Marked
Case Is = 1
GroundA(ff, gg) = STX
Case Is = 2
WallsA(ff, gg) = STX
End Select
End If
End If
End Sub
Sub UnitFrames()
   If CurrentPos = NextPos Then
CurrentPos = NextPos
Else
If CurrentPos - NextPos = 4 Then
End If
If CurrentPos > NextPos Then
If CurrentPos - NextPos > 4 Then
CurrentPos = CurrentPos + 1
ssTexts = 12
Else
CurrentPos = CurrentPos - 1
ssTexts = 13
End If
End If
If CurrentPos < NextPos Then
If NextPos - CurrentPos > 4 Then
CurrentPos = CurrentPos - 1
ssTexts = 13
Else
CurrentPos = CurrentPos + 1
ssTexts = 12
End If
End If
End If
If CurrentPos > 7 Then CurrentPos = 0
If CurrentPos < 0 Then CurrentPos = 7
End Sub
Sub MoveUnit()
'If FinalPos = CurrentPos Then
'If DownKey = True And UpKey = False And LeftKey = False And RightKey = False And UpLeftKey = False And UpRightKey = False And DownRightKey = False And DownLeftKey = False Then
'UnitYPos = UnitYPos + 8
'End If
'If RightKey = True And DownKey = False And UpKey = False And LeftKey = False And UpLeftKey = False And UpRightKey = False And DownRightKey = False And DownLeftKey = False Then
'UnitXPos = UnitXPos + 8
'End If
'If UpKey = True And RightKey = False And DownKey = False And LeftKey = False And UpLeftKey = False And UpRightKey = False And DownRightKey = False And DownLeftKey = False Then
'UnitYPos = UnitYPos - 8
'End If
'If LeftKey = True And RightKey = False And DownKey = False And UpKey = False And UpLeftKey = False And UpRightKey = False And DownRightKey = False And DownLeftKey = False Then
'UnitXPos = UnitXPos - 8
'End If
'If UpRightKey = True And LeftKey = False And RightKey = False And DownKey = False And UpKey = False And UpLeftKey = False And DownRightKey = False And DownLeftKey = False Then
'UnitXPos = UnitXPos + 8
'UnitYPos = UnitYPos - 8
'End If
'If DownRightKey = True And UpRightKey = False And LeftKey = False And RightKey = False And DownKey = False And UpKey = False And UpLeftKey = False And DownLeftKey = False Then
'UnitXPos = UnitXPos + 8
'UnitYPos = UnitYPos + 8
'End If
'If DownLeftKey = True And UpRightKey = False And LeftKey = False And RightKey = False And DownKey = False And UpKey = False And UpLeftKey = False And DownRightKey = False Then
'UnitXPos = UnitXPos - 8
'UnitYPos = UnitYPos + 8
'End If
'If UpLeftKey = True And UpRightKey = False And LeftKey = False And RightKey = False And DownKey = False And UpKey = False And DownRightKey = False And DownLeftKey = False Then
'UnitXPos = UnitXPos - 8
'UnitYPos = UnitYPos - 8
'End If
'End If
'Frame = Frame + 1
End Sub


Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
MoveMenu1 = False
MoveMenu2 = False
If Button = 2 Then
frmMain.PopupMenu TS, True
End If
MousePressedOnClick = False
End Sub

Private Sub TAL1_Click()
Call ClearAll

End Sub

Private Sub TAL2_Click()
Call ClearAll
End Sub

Private Sub fr10_Click()
CustomFrameRate = 10
End Sub

Private Sub fr15_Click()
CustomFrameRate = 15
End Sub

Private Sub fr20_Click()
CustomFrameRate = 20
End Sub

Private Sub fr25_Click()
CustomFrameRate = 25
End Sub

Private Sub fr30_Click()
CustomFrameRate = 30
End Sub

Private Sub l1_Click()
GroundFillTile = 1
Restart
End Sub

Private Sub l2_Click()
GroundFillTile = 2
Restart
End Sub

Private Sub l3_Click()
GroundFillTile = 3
Restart
End Sub

Private Sub l4_Click()
GroundFillTile = 4
Restart
End Sub

Private Sub l5_Click()
GroundFillTile = 5
Restart
End Sub

Private Sub l6_Click()
GroundFillTile = 6
Restart
End Sub

Public Sub LM_Click()
LoadFile "default"
End Sub



Private Sub RIPC_Click()
Marked = 3
Thing = 17
End Sub

Private Sub RIPS_Click()
Marked = 3
Thing = 16
End Sub

Private Sub SCFrameRate_Click()
CustomFrameRate = Val(InputBox("Please enter your custom frame rate(less than 50)", "Frame Rate Settings", "20"))
End Sub

Public Sub SM_Click()
Dim ax, ay, bx, by, tfr As Integer
cmndlg1.ShowSave
If Len(cmndlg1.FileName) <> 0 Then
Finished = False
Do While Finished = False
On Error Resume Next
Open cmndlg1.FileName For Output As #2
Write #2, GroundFillTile, WallsFillTile
For ax = 0 To 500
For ay = 0 To 500
If GroundA(ax, ay) <> GroundFillTile And GroundA(ax, ay) <> 0 Then
Write #2, ax, ay, GroundA(ax, ay)
End If
Next
Next
Write #2, 0, 0, 0

For bx = 0 To 500
For by = 0 To 500
If WallsA(bx, by) <> 0 Then
Write #2, bx, by, WallsA(bx, by)
End If
Next
Next
Write #2, 0, 0, 0


If UBound(Wthing) > 0 Then
For tfr = LBound(Wthing) To UBound(Wthing)
Write #2, Wthing(tfr).ThingAs, Wthing(tfr).WorldX, Wthing(tfr).WorldY
Next
End If
Close #2
Finished = True
Loop

End If
End Sub


Private Sub TGL1_Click()
Call ClearAll
Marked = 1
CalcTLen
End Sub

Private Sub TGL2_Click()
Call ClearAll
Marked = 2
CalcTLen
End Sub

Private Sub TPL1_Click()
Call ClearAll
End Sub

Private Sub TPL2_Click()
Call ClearAll
End Sub
Public Sub ClearAll()
frmMain.PopupMenu TS, False
End Sub
Public Sub Restart()
'Refilling ground filling (*smiles*) array with new values..
For fx = 0 To 500
For fy = 0 To 500
GroundA(fx, fy) = GroundFillTile
Next
Next
End Sub

Public Sub DrawGrid()
 For j = Int(ViewY / TILE_HEIGHT) To Int(ViewY / TILE_HEIGHT) + 14
For i = Int(ViewX / TILE_WIDTH) To Int(ViewX / TILE_WIDTH) + 18
CurX = i * TILE_WIDTH - ViewX
CurY = j * TILE_HEIGHT - ViewY
 If EG.Checked = True Then
  msurfBack.DrawLine CurX, CurY, CurX, CurY + SCREEN_HEIGHT
  msurfBack.DrawLine CurX, CurY, CurX + SCREEN_WIDTH, CurY
  Else
  End If
Next
Next
End Sub
Private Sub TNO_Click()
GroundFillTile = 0
Restart
End Sub

Public Sub CalcTLen()
SumOfAllTiles = 0
'ground length..
If Marked = 1 Then
For SubIndex = 0 To NumberOfGTiles
SumOfAllTiles = SumOfAllTiles + GroundDesc(SubIndex).lWidth
Next
Else
If Marked = 2 Then
'walls length..
For STX = 0 To NumberOfWTiles
SumOfAllTiles = SumOfAllTiles + WallsDesc(STX).lWidth
Next
End If
End If
End Sub

Sub UpdateMap()
For j = Int(ViewY / TILE_HEIGHT) To Int(ViewY / TILE_HEIGHT) + 14
For i = Int(ViewX / TILE_WIDTH) To Int(ViewX / TILE_WIDTH) + 18
If WallsA(i, j) = 3 And WallsA(i + 1, j - 1) = 6 And WallsA(i + 1, j + 1) <> 6 Then WallsA(i + 1, j) = 1
If WallsA(i, j) = 6 And WallsA(i, j - 1) = 3 Then WallsA(i, j - 1) = 12
If WallsA(i, j) = 6 And WallsA(i + 1, j - 1) = 3 Then WallsA(i, j - 1) = 12
If WallsA(i, j) = 6 And WallsA(i + 1, j + 1) = 9 Then WallsA(i, j + 1) = 15
If WallsA(i, j) = 1 And WallsA(i + 1, j) = 6 Then WallsA(i, j) = 4
If WallsA(i, j) = 9 And WallsA(i + 1, j + 1) = 6 And WallsA(i + 1, j) = 0 And WallsA(i + 1, j) = 0 Then WallsA(i + 1, j) = 7
Next
Next
End Sub
Sub LoadFile(myFileName As String)
cmndlg1.InitDir = App.Path & "\saves\"
Finished = False
'Init of Load Procedure
Dim Cancel As Boolean
Dim ax, ay, bx, by, sx, sy, GT, WT As Integer
myrt = 0
'''''''''''''''''''''''
If myFileName <> "default" Then
cmndlg1.FileName = myFileName
Else
cmndlg1.ShowOpen
End If
Open cmndlg1.InitDir & cmndlg1.FileName For Input As #1
Do While Finished = False
DoEvents

Input #1, GroundFillTile, WallsFillTile
'Default Filling...................
On Error Resume Next
For sx = 0 To 500
For sy = 0 To 500
GroundA(sx, sy) = GroundFillTile
WallsA(sx, sy) = WallsFillTile
Next
Next

Do
Input #1, ax, ay, GT
GroundA(ax, ay) = GT
Loop Until ax = 0 And ay = 0 And GT = 0

Do
Input #1, bx, by, WT
WallsA(bx, by) = WT
Loop Until bx = 0 And by = 0 And WT = 0

Do While Not EOF(1)
Input #1, WTAS, WTX, WTY
ReDim Preserve Wthing(0 To myrt) As Things
Set Wthing(myrt) = New Things
 Let Wthing(myrt).WorldX = WTX
 Let Wthing(myrt).WorldY = WTY
 Let Wthing(myrt).ThingAs = WTAS
myrt = myrt + 1
Loop
Close #1
Num = myrt - 1

Finished = True
Loop
Close #1
End Sub
Sub RemoveTile()
Dim tfr As Integer
For tfr = 0 To Num - 1
If GetX + ViewX > Wthing(tfr).WorldX And GetX + ViewX < Wthing(tfr).WorldX + ThingsDesc(Wthing(tfr).ThingAs).lWidth And GetY + ViewY > Wthing(tfr).WorldY And GetY + ViewY < Wthing(tfr).WorldY + ThingsDesc(Wthing(tfr).ThingAs).lHeight Then
Set Wthing(tfr) = Wthing(Num)
Num = Num - 1
ReDim Preserve Wthing(0 To Num) As Things
drf = tfr
End If
Next
End Sub
