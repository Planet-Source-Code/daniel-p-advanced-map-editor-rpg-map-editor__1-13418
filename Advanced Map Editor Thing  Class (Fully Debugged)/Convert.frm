VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   90
   ClientTop       =   375
   ClientWidth     =   4620
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4620
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Tell me the number !"
      Height          =   495
      Left            =   1080
      TabIndex        =   2
      Top             =   1440
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Text            =   "dfsdfsdf"
      Top             =   2160
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   $"Convert.frx":0000
      Height          =   615
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FirstDigits(0 To 9) As String
Dim SecondDigits(0 To 9) As String
Dim Tens(0 To 5) As String
Dim TtoN(0 To 9) As String
Dim ConsistsOf(1 To 4) As String
Dim OriginNumber As String

Private Sub Command1_Click()
Text1.Text = ""
Call Finalise
Call Finalise
End Sub

Private Sub Form_Load()
Text2.Text = ""
FirstDigits(1) = "One"
FirstDigits(2) = "Two"
FirstDigits(3) = "Three"
FirstDigits(4) = "Four"
FirstDigits(5) = "Five"
FirstDigits(6) = "Six"
FirstDigits(7) = "Seven"
FirstDigits(8) = "Eigth"
FirstDigits(9) = "Nine"
FirstDigits(0) = "Zero"
SecondDigits(0) = "Zero"
SecondDigits(2) = "Twenty"
SecondDigits(3) = "Thirty"
SecondDigits(4) = "Fourty"
SecondDigits(5) = "Fifty"
SecondDigits(6) = "Sixty"
SecondDigits(7) = "Seventy"
SecondDigits(8) = "Eighty"
SecondDigits(9) = "Ninety"
TtoN(0) = "Ten"
TtoN(1) = "Eleven"
TtoN(2) = "Twelve"
TtoN(3) = "Thirteen"
TtoN(4) = "Fourteen"
TtoN(5) = "Fifteen"
TtoN(6) = "Sixteen"
TtoN(7) = "Seventeen"
TtoN(8) = "Eighteen"
TtoN(9) = "Nineteen"
End Sub
Public Sub Finalise()
If Val(Text2.Text) < 0 Or Len(Text2.Text) > 4 Then MsgBox "This is not a valid number. Please read carefully instructions on top of the page.", , "Error message!"
'1912
'Reserving a number for further uses..
OriginNumber = Text2.Text
'Adding if need, zeros to 4 digits number ( xxxx )
If Len(Text2.Text) = 2 Then
Text2.Text = "00" & Text2.Text
Else
If Len(Text2.Text) = 3 Then
Text2.Text = "0" & Text2.Text
Else
If Len(Text2.Text) = 1 Then
Text2.Text = "000" & Text2.Text
End If
End If
End If
'Checking values.. if between 0 and 9 then..
If Val(Text2.Text) < 10 Then
ConsistsOf(1) = FirstDigits(Val(Right(Text2.Text, 1)))
Else
'Between 10 and 19....
If Val(Right(Text2.Text, 2)) >= 10 And Val(Right(Text2.Text, 2)) < 20 Then
ConsistsOf(2) = TtoN(Val(Right(Text2.Text, 1)))
ConsistsOf(1) = ""
Else
'Else ( >19 and <100)
ConsistsOf(2) = SecondDigits(Val(Left(Right(Text2.Text, 2), 1)))
ConsistsOf(1) = FirstDigits(Val(Right(Text2.Text, 1)))
End If
End If
If Len(OriginNumber) = 1 Then
ConsistsOf(2) = " "
ConsistsOf(3) = " "
ConsistsOf(4) = " "
End If
If Len(OriginNumber) = 2 Then
ConsistsOf(3) = " "
ConsistsOf(4) = " "
End If
'A hundreds check...
If Len(OriginNumber) > 2 Then
ConsistsOf(3) = FirstDigits(Val(Left(Right(Text2.Text, 3), 1))) & " Hundred"
End If
'Thousands check..
If Len(OriginNumber) > 3 Then
ConsistsOf(4) = FirstDigits(Val(Left(Right(Text2.Text, 4), 1))) & " Thousand"
End If
Text1.Text = ConsistsOf(4) & " " & ConsistsOf(3) & " " & ConsistsOf(2) & " " & ConsistsOf(1)
End Sub





