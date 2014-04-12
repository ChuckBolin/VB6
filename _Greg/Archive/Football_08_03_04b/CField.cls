VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'CField Class
Option Explicit

'constants
Private Const HOME = 0
Private Const VISITOR = 1
Private Const EAST = 0
Private Const WEST = 1
Private Const BALLSTEP = 10  'increase to make ball velocity decrease

'private variables
'form and field specific stuff
Private nFormWidth As Integer  'in twips, physical form
Private nFormHeight As Integer
Private nScaleTop As Integer    'for scale mode in yards
Private nScaleLeft As Integer
Private nScaleHeight As Integer
Private nScaleWidth As Integer
Private nOffLine As Single      'position of offensive line
Private nDefLine As Single      'position of defensive line
Private nBallX As Single         'ball X,Y,Z position
Private nBallY As Single
Private nBallZ As Single
Private nBallDir As Single 'direction ball is thrown
Private nBallVel As Single 'velocity of ball
Private nBallDX As Single 'change in ball x position
Private nBallDY As Single 'change in ball y position
Private bBallInAir As Boolean 'true if ball is being thrown
Private bOffLineOn As Boolean 'displays line if true
Private bDefLineOn As Boolean
Private nFirstDownLine As Single
Private bFirstDownLineOn As Boolean
Private nDown As Integer '1 through 4

'game specific stuff
Private nYFD As Single 'yards to go to first down
Private nYTD As Single ' yards to touchdown
Private nHomeScore As Integer
Private nVisitorScore As Integer
Private nPossession As Integer '0=Home, 1=Visitor
Private nDriveDir As Integer '0 = East, 1=West
Private nQBx As Single 'quarterback x,y position
Private nQBy As Single
Private nRecx As Single 'receiver x,y position
Private nRecy As Single

'initialize field
Public Sub Initialize()
  
  'form and stuff on form...field is 2.16 x longer than wide, or wider than high
  'nFormWidth = 8656
  'nFormHeight = 4000
  nFormWidth = 10800
  nFormHeight = nFormWidth / 1.89 ' 5000
  
  'includes field, endzones and 6' border around field
  nScaleTop = 35 '28.67
  nScaleHeight = -70 '-57.3
  nScaleLeft = -66
  nScaleWidth = 132
  nOffLine = -30
  nDefLine = 15
  nBallX = 15
  nBallY = 0
  nBallZ = 0
  nQBx = 15
  nQBy = 0
  nRecx = 45
  nRecy = 0
  bOffLineOn = True
  bDefLineOn = False
  nFirstDownLine = -20
  bFirstDownLineOn = True
  bBallInAir = False
End Sub

'properties
'goto tools, add procedure.select public property
Public Property Get Width() As Integer
  Width = nFormWidth
End Property

Public Property Let Width(ByVal vNewValue As Integer)
  nFormWidth = vNewValue
End Property

Public Property Get Height() As Integer
  Height = nFormHeight
End Property

Public Property Let Height(ByVal vNewValue As Integer)
  nFormHeight = vNewValue
End Property

Public Property Get ScaleTop() As Integer
  ScaleTop = nScaleTop
End Property

Public Property Let ScaleTop(ByVal vNewValue As Integer)
  nScaleTop = vNewValue
End Property

Public Property Get ScaleLeft() As Integer
  ScaleLeft = nScaleLeft
End Property

Public Property Let ScaleLeft(ByVal vNewValue As Integer)
  nScaleLeft = vNewValue
End Property

Public Property Get ScaleHeight() As Integer
  ScaleHeight = nScaleHeight
End Property

Public Property Let ScaleHeight(ByVal vNewValue As Integer)
  nScaleHeight = vNewValue
End Property

Public Property Get ScaleWidth() As Integer
  ScaleWidth = nScaleWidth
End Property

Public Property Let ScaleWidth(ByVal vNewValue As Integer)
  nScaleWidth = vNewValue
End Property

Public Property Get DefLine() As Single
  DefLine = nDefLine
End Property

Public Property Let DefLine(ByVal vNewValue As Single)
  nDefLine = vNewValue
End Property

Public Property Get OffLine() As Single
  OffLine = nOffLine
End Property

Public Property Let OffLine(ByVal vNewValue As Single)
  nOffLine = vNewValue
End Property

Public Property Get BallX() As Single
  BallX = nBallX
End Property

Public Property Let BallX(ByVal vNewValue As Single)
  nBallX = vNewValue
End Property

Public Property Get BallY() As Single
  BallY = nBallY
End Property

Public Property Let BallY(ByVal vNewValue As Single)
  nBallY = vNewValue
End Property

Public Property Get BallZ() As Single
  BallZ = nBallZ
End Property

Public Property Let BallZ(ByVal vNewValue As Single)
  nBallZ = vNewValue
End Property

Public Property Get OffLineOn() As Boolean
  OffLineOn = bOffLineOn
End Property

Public Property Let OffLineOn(ByVal vNewValue As Boolean)
  bOffLineOn = vNewValue
End Property

Public Property Get DefLineOn() As Boolean
  DefLineOn = bDefLineOn
End Property

Public Property Let DefLineOn(ByVal vNewValue As Boolean)
  bDefLineOn = vNewValue
End Property


Public Property Get FirstDownLine() As Single
  FirstDownLine = nFirstDownLine
End Property

Public Property Let FirstDownLine(ByVal vNewValue As Single)
  nFirstDownLine = vNewValue
End Property

Public Property Get FirstDownLineOn() As Boolean
  FirstDownLineOn = bFirstDownLineOn
End Property

Public Property Let FirstDownLineOn(ByVal vNewValue As Boolean)
  bFirstDownLineOn = vNewValue
End Property

Public Property Get QBX() As Single
  QBX = nQBx
End Property

Public Property Let QBX(ByVal vNewValue As Single)
  nQBx = vNewValue
End Property

Public Property Get QBY() As Single
  QBY = nQBy
End Property

Public Property Let QBY(ByVal vNewValue As Single)
  nQBy = vNewValue
End Property

Public Property Get RecX() As Single
  RecX = nRecx
End Property

Public Property Let RecX(ByVal vNewValue As Single)
  nRecx = vNewValue
End Property

Public Property Get RecY() As Single
  RecY = nRecy
End Property

Public Property Let RecY(ByVal vNewValue As Single)
  nRecy = vNewValue
End Property

'initialize vel and dir of ball being thrown
Public Sub MoveBall()
  Dim nDistance As Single
  Dim nHyp As Single
  
  nDistance = GetTargetDistance(nQBx, nQBy, nRecx, nRecy)
  If nDistance = 0 Then Exit Sub
  nBallVel = BALLSTEP
  nHyp = nDistance / nBallVel
  nBallDir = GetTargetDirection(nQBx, nQBy, nRecx, nRecy)
  nBallDY = nHyp * Cos(nBallDir)
  nBallDX = nHyp * sIn(nBallDir)
  bBallInAir = True
End Sub

'responsible for all animation
Public Sub Update()
  
  'update position of ball if enabled
  If bBallInAir = True Then
    nBallX = nBallX + nBallDX
    nBallY = nBallY + nBallDY
    If GetTargetDistance(nBallX, nBallY, nRecx, nRecy) < 1 Then bBallInAir = False
  End If
End Sub

