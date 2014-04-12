VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Manipulates ball."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarX As Single 'local copy
Private mvarY As Single 'local copy
Private mvarBallWidth As Single 'local copy
Private mvarBallLength As Single 'local copy
Private mvarVelocity As Single 'local copy
Private mvarDirection As Single 'local copy
Private mvarTargetX As Single 'local copy
Private mvarTargetY As Single 'local copy

'updates position of ball x,y if vel <> 0
Public Sub Update()
  If mvarVelocity = 0 Then Exit Sub
  
  Dim opp As Single, adj As Single
  Dim dx As Single, dy As Single
  Dim dir As Single
  
  opp = mvarTargetY - mvarY
  adj = mvarTargetX - mvarX
  If adj = 0 Then adj = 0.001
    
  'calculate dir based upon x,y,targetx, targety
  If mvarTargetX > mvarX And mvarTargetY > mvarY Then ' quadrant I
    dir = Atn(opp / adj)
    dx = mvarVelocity * Cos(dir)
    dy = mvarVelocity * Sin(dir)
  ElseIf mvarTargetX < mvarX And mvarTargetY > mvarY Then 'quadrant II
    dir = Atn(opp / adj) - 3.14
    dx = mvarVelocity * Cos(dir)
    dy = mvarVelocity * Sin(dir)
  ElseIf mvarTargetX < mvarX And mvarTargetY < mvarY Then ' quadrant III
    dir = Atn(opp / adj) + 3.14
    dx = mvarVelocity * Cos(dir)
    dy = mvarVelocity * Sin(dir)
  ElseIf mvarTargetX > mvarX And mvarTargetY < mvarY Then ' quadrant IV
    dir = -Atn(-opp / adj)
    dx = mvarVelocity * Cos(dir)
    dy = mvarVelocity * Sin(dir)
  ElseIf mvarTargetX = mvarX And mvarTargetY > mvarY Then ' 1.57 rad
    dir = 1.57
    dx = mvarVelocity * Cos(dir)
    dy = mvarVelocity * Sin(dir)
  ElseIf mvarTargetX < mvarX And mvarTargetY = mvarY Then ' 3.14 rad
    dir = 3.14
    dx = mvarVelocity * Cos(dir)
    dy = mvarVelocity * Sin(dir)
  ElseIf mvarTargetX = mvarX And mvarTargetY < mvarY Then ' 4.71 rad
    dir = 4.71
    dx = mvarVelocity * Cos(dir)
    dy = mvarVelocity * Sin(dir)
  ElseIf mvarTargetX > mvarX And mvarTargetY = mvarY Then ' 0 rad
    dir = 0
    dx = mvarVelocity * Cos(dir)
    dy = mvarVelocity * Sin(dir)
    
  End If
  
  'update position
  mvarX = mvarX + dx
  mvarY = mvarY + dy
  
  If Abs(mvarX - mvarTargetX) < 1 And Abs(mvarY - mvarTargetY) < 1 Then mvarVelocity = 0
  mvarDirection = dir
  
End Sub

Public Sub AssignTarget(x As Single, y As Single)
  mvarTargetX = x
  mvarTargetY = y
'  mvarVelocity = 1
End Sub

Public Property Let TargetY(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TargetY = 5
    mvarTargetY = vData
End Property


Public Property Get TargetY() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TargetY
    TargetY = mvarTargetY
End Property



Public Property Let TargetX(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TargetX = 5
    mvarTargetX = vData
End Property


Public Property Get TargetX() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TargetX
    TargetX = mvarTargetX
End Property



Public Property Let Direction(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Direction = 5
    mvarDirection = vData
End Property


Public Property Get Direction() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Direction
    Direction = mvarDirection
End Property



Public Property Let Velocity(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Velocity = 5
    mvarVelocity = vData
End Property


Public Property Get Velocity() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Velocity
    Velocity = mvarVelocity
End Property



Public Property Let BallLength(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BallLength = 5
    mvarBallLength = vData
End Property


Public Property Get BallLength() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BallLength
    BallLength = mvarBallLength
End Property



Public Property Let BallWidth(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BallWidth = 5
    mvarBallWidth = vData
End Property


Public Property Get BallWidth() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BallWidth
    BallWidth = mvarBallWidth
End Property



Public Property Let y(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Y = 5
    mvarY = vData
End Property


Public Property Get y() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Y
    y = mvarY
End Property



Public Property Let x(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.X = 5
    mvarX = vData
End Property


Public Property Get x() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.X
    x = mvarX
End Property



Private Sub Class_Initialize()
  mvarX = 50
  mvarY = 50
  mvarBallWidth = 1
  mvarBallLength = 1
End Sub
