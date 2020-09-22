VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BasicUnit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarUnitX As Integer 'local copy
Private mvarUnitY As Integer 'local copy
Public Property Let UnitY(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UnitY = 5
    mvarUnitY = vData
End Property


Public Property Get UnitY() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UnitY
    UnitY = mvarUnitY
End Property



Public Property Let UnitX(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UnitX = 5
    mvarUnitX = vData
End Property


Public Property Get UnitX() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UnitX
    UnitX = mvarUnitX
End Property



