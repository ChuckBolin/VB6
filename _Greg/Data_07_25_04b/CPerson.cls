VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPerson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************************************
'  CPERSON.CLS - July 25, 2004
'  All information and procedures for people
'***************************************************************************************************
Option Explicit

'private constants

'personal information about players
Private sFName As String 'first name
Private sLName As String 'last name
Private nJerseyNumber As Integer '
Private dDOB As Date  'date of birth
Private nAge As Integer 'age
Private nHeight As Single  'height and weight
Private nWeight As Single
Private nPriPos As Integer 'primary position played
Private nSecPos As Integer 'secondary position played

'real time data
Private nX As Single 'x pos
Private nY As Single 'y pos
Private nVel As Single 'velocity
Private nMaxVel As Single 'maximum velocity
Private nDir As Single 'direction of movement
Private nDX As Single 'change in X direction
Private nDY As Single 'change in Y direction

Public Property Get Vel() As Single
  Vel = nVel
End Property

Public Property Let Vel(ByVal vNewValue As Single)
  nVel = vNewValue
End Property

Public Property Get MaxVel() As Single
  MaxVel = nMaxVel
End Property

Public Property Let MaxVel(ByVal vNewValue As Single)
  nMaxVel = vNewValue
End Property

Public Property Get Dir() As Single
  Dir = nDir
End Property

Public Property Let Dir(ByVal vNewValue As Single)
  nDir = vNewValue
End Property

Public Property Get X() As Single
  X = nX
End Property

Public Property Let X(ByVal vNewValue As Single)
  nX = vNewValue
End Property

Public Property Get Y() As Single
  Y = nY
End Property

Public Property Let Y(ByVal vNewValue As Single)
  nY = vNewValue
End Property

Public Property Get DX() As Single
  DX = nDX
End Property

Public Property Let DX(ByVal vNewValue As Single)
  nDX = vNewValue
End Property

Public Property Get DY() As Single
  DY = nDY
End Property

Public Property Let DY(ByVal vNewValue As Single)
  nDY = vNewValue
End Property



Public Property Get GetFirstName() As String
  GetFirstName = sFName
End Property

Public Property Let GetFirstName(ByVal vNewValue As String)
  sFName = vNewValue
End Property

Public Property Get GetLastName() As String
  GetLastName = sLName
End Property

Public Property Let GetLastName(ByVal vNewValue As String)
  sLName = vNewValue
End Property

Public Property Get Age() As Variant
  Age = nAge
End Property

Public Property Get Height() As Single
  Height = nHeight
End Property

Public Property Let Height(ByVal vNewValue As Single)
  nHeight = vNewValue
End Property

Public Property Get Weight() As Single
  Weight = nWeight
End Property

Public Property Let Weight(ByVal vNewValue As Single)
   nWeight = vNewValue
End Property

Public Property Get PriPosition() As Integer
  PriPosition = nPriPos
End Property

Public Property Let PriPosition(ByVal vNewValue As Integer)
  nPriPos = vNewValue
End Property

Public Property Get SecPosition() As Integer
  SecPosition = nSecPos
End Property

Public Property Let SecPosition(ByVal vNewValue As Integer)
  nSecPos = vNewValue
End Property


