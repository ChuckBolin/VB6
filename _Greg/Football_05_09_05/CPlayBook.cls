VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPlayBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'*****************************************************
' CPLAYBOOK.CLS Written 7/27/2004
'
'*****************************************************
Option Explicit

Private nPlayNumber As Integer
Private sPosition As String
Private nX As Single
Private nY As Single
Private sAction As String
Private nWX1 As Single
Private nWX2 As Single
Private nWY1 As Single
Private nWY2 As Single

Private Sub Class_Initialize()
  nPlayNumber = 0
  sPosition = " "
  nX = 0
  nY = 0
  sAction = " "
   nWX1 = 0
  nWX2 = 0
  nWY1 = 0
  nWY2 = 0
End Sub

Public Property Get PlayNumber() As Integer
  PlayNumber = nPlayNumber
End Property

Public Property Let PlayNumber(ByVal nNewValue As Integer)
  nPlayNumber = nNewValue
End Property

Public Property Get Position() As String
  Position = sPosition
End Property

Public Property Let Position(ByVal sNewValue As String)
  sPosition = sNewValue
End Property

Public Property Get X() As Single
  X = nX
End Property

Public Property Let X(ByVal nNewValue As Single)
  nX = nNewValue
End Property

Public Property Get Y() As Single
  Y = nY
End Property

Public Property Let Y(ByVal nNewValue As Single)
  nY = nNewValue
End Property

Public Property Get Action() As String
  Action = sAction
End Property

Public Property Let Action(ByVal sNewValue As String)
  sAction = sNewValue
End Property

Public Property Get WX1() As Single
  WX1 = nWX1
End Property

Public Property Let WX1(ByVal nNewValue As Single)
  nWX1 = nNewValue
End Property

Public Property Get WX2() As Single
  WX2 = nWX2
End Property

Public Property Let WX2(ByVal nNewValue As Single)
  nWX2 = nNewValue
End Property

Public Property Get WY1() As Single
  WY1 = nWY1
End Property

Public Property Let WY1(ByVal nNewValue As Single)
  nWY1 = nNewValue
End Property

Public Property Get WY2() As Single
  WY2 = nWY2
End Property

Public Property Let WY2(ByVal nNewValue As Single)
  nWY2 = nNewValue
End Property


