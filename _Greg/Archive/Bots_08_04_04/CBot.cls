VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************************
' CBOT.CLS Written 8/4/2004
'
'*****************************************************
Option Explicit

Private nX As Single
Private nY As Single
Private nTX As Single
Private nTY As Single

Private nCX As Single
Private nCY As Single
Private nWX As Single
Private nWY As Single
Private nWX1 As Single
Private nWY1 As Single
Private nWX2 As Single
Private nWY2 As Single
Private nWX3 As Single
Private nWY3 As Single
Private nIndex As Integer
Private nWPCount As Integer
Private nActionMode As Integer
Private nBotType As Integer
Private nDirection As Single
Private nSpeed As Single
Private nDiameter As Single
Private bFound As Boolean
Private nDX As Single
Private nDY As Single

Private Sub Class_Initialize()
  nX = 0
  nY = 0
  nTX = 0
  nTY = 0
  nCX = 0
  nCY = 0
  nWX = 0
  nWY = 0
  nWX1 = 0
  nWY1 = 0
  nWX2 = 0
  nWY2 = 0
  nWX3 = 0
  nWY3 = 0
  nIndex = 0
  nWPCount = 0
  nActionMode = 0
  nBotType = 0
  nDirection = 0
  nSpeed = 0
  nDiameter = 1.5
  bFound = False
  nDX = 0
  nDY = 0
End Sub

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

Public Property Get TX() As Single
  TX = nTX
End Property

Public Property Let TX(ByVal nNewValue As Single)
  nTX = nNewValue
End Property

Public Property Get TY() As Single
  TY = nTY
End Property

Public Property Let TY(ByVal nNewValue As Single)
  nTY = nNewValue
End Property

Public Property Get CX() As Single
  CX = nCX
End Property

Public Property Let CX(ByVal nNewValue As Single)
  nCX = nNewValue
End Property

Public Property Get CY() As Single
  CY = nCY
End Property

Public Property Let CY(ByVal nNewValue As Single)
  nCY = nNewValue
End Property

Public Property Get WX() As Single
  WX = nWX
End Property

Public Property Let WX(ByVal nNewValue As Single)
  nWX = nNewValue
End Property

Public Property Get WY() As Single
  WY = nWY
End Property

Public Property Let WY(ByVal nNewValue As Single)
  nWY = nNewValue
End Property

Public Property Get WX1() As Single
  WX1 = nWX1
End Property

Public Property Let WX1(ByVal nNewValue As Single)
  nWX1 = nNewValue
End Property

Public Property Get WY1() As Single
  WY1 = nWY1
End Property

Public Property Let WY1(ByVal nNewValue As Single)
  nWY1 = nNewValue
End Property

Public Property Get WX2() As Single
  WX2 = nWX2
End Property

Public Property Let WX2(ByVal nNewValue As Single)
  nWX2 = nNewValue
End Property

Public Property Get WY2() As Single
  WY2 = nWY2
End Property

Public Property Let WY2(ByVal nNewValue As Single)
  nWY2 = nNewValue
End Property

Public Property Get WX3() As Single
  WX3 = nWX3
End Property

Public Property Let WX3(ByVal nNewValue As Single)
  nWX3 = nNewValue
End Property

Public Property Get WY3() As Single
  WY3 = nWY3
End Property

Public Property Let WY3(ByVal nNewValue As Single)
  nWY3 = nNewValue
End Property

Public Property Get Index() As Integer
  Index = nIndex
End Property

Public Property Let Index(ByVal nNewValue As Integer)
  If nNewValue > 3 Then Exit Property
  If nNewValue < 1 Then Exit Property
  nIndex = nNewValue
End Property

Public Property Get WPCount() As Integer
  WPCount = nWPCount
End Property

Public Property Let WPCount(ByVal nNewValue As Integer)
  If nNewValue > 3 Then Exit Property
  nWPCount = nNewValue
End Property

Public Property Get ActionMode() As Integer
  ActionMode = nActionMode
End Property

Public Property Let ActionMode(ByVal nNewValue As Integer)
  nActionMode = nNewValue
End Property

Public Property Get BotType() As Integer
  BotType = nBotType
End Property

Public Property Let BotType(ByVal nNewValue As Integer)
  nBotType = nNewValue
End Property

Public Property Get Direction() As Single
  Direction = nDirection
End Property

Public Property Let Direction(ByVal nNewValue As Single)
  If nNewValue > 6.284 Then nNewValue = 6.284
  If nNewValue < 0 Then nNewValue = 0
  nDirection = nNewValue
End Property

Public Property Get Speed() As Single
  Speed = nSpeed
End Property

Public Property Let Speed(ByVal nNewValue As Single)
  nSpeed = nNewValue
End Property

Public Property Get Diameter() As Single
  Diameter = nDiameter
End Property

Public Property Let Diameter(ByVal nNewValue As Single)
  nDiameter = nNewValue
End Property

Public Property Get Found() As Boolean
  Found = bFound
End Property

Public Property Let Found(ByVal bNewValue As Boolean)
  bFound = bNewValue
End Property

Public Property Get DX() As Single
  DX = nDX
End Property

Public Property Let DX(ByVal nNewValue As Single)
  nDX = nNewValue
End Property

Public Property Get DY() As Single
  DY = nDY
End Property

Public Property Let DY(ByVal nNewValue As Single)
  nDY = nNewValue
End Property






