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
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   480
      Top             =   1920
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   2280
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   840
      TabIndex        =   3
      Text            =   "Text3"
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   720
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   240
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   855
      Left            =   1080
      TabIndex        =   0
      Top             =   960
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim MyUnit As BasicUnit
Public Sub Command1_Click()
x = x + 1
Set MyUnit = New BasicUnit
MyUnit.UnitX = 156 + x
MyUnit.UnitY = 145 + x
End Sub

Public Sub Command2_Click()
Open "c:/games/xx.txt" For Output As #1
For x = 0 To x
Write #1, MyUnit.UnitX
Write #1, MyUnit.UnitY
Next
Close #1
End Sub

