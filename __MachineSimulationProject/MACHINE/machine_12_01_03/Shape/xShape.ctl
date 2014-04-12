VERSION 5.00
Begin VB.UserControl Shape 
   BackStyle       =   0  'Transparent
   ClientHeight    =   2190
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2145
   ScaleHeight     =   2190
   ScaleWidth      =   2145
   ToolboxBitmap   =   "xShape.ctx":0000
   Begin VB.Shape pt 
      BackColor       =   &H00808000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   105
      Index           =   7
      Left            =   660
      Top             =   1740
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Shape pt 
      BackColor       =   &H00808000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   105
      Index           =   6
      Left            =   660
      Top             =   1620
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Shape pt 
      BackColor       =   &H00808000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   105
      Index           =   5
      Left            =   660
      Top             =   1500
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Shape pt 
      BackColor       =   &H00808000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   105
      Index           =   4
      Left            =   660
      Top             =   1380
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Shape pt 
      BackColor       =   &H00808000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   105
      Index           =   3
      Left            =   420
      Top             =   1740
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Shape pt 
      BackColor       =   &H00808000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   105
      Index           =   2
      Left            =   420
      Top             =   1620
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Shape pt 
      BackColor       =   &H00808000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   105
      Index           =   1
      Left            =   420
      Top             =   1500
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Shape pt 
      BackColor       =   &H00808000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   105
      Index           =   0
      Left            =   420
      Top             =   1380
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Shape Shape 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      Height          =   555
      Left            =   0
      Top             =   0
      Width           =   675
   End
End
Attribute VB_Name = "Shape"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event Resize()
Attribute Resize.VB_Description = "Occurs when a form is first displayed or the size of an object changes."
'Default Property Values:
Const m_def_RunTime = False

'Property Variables:
Dim m_RunTime As Boolean
Dim m_Height As Integer
Dim m_Width As Integer
Dim mx, my As Integer
Dim mblnResize As Boolean
Dim m_MoveShape As Boolean


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Shape,Shape,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
  BackColor = Shape.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
  Shape.BackColor() = New_BackColor
  PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
  Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
  UserControl.Enabled() = New_Enabled
  PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Shape,Shape,-1,FillStyle
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Returns/sets the fill style of a shape."
  BackStyle = Shape.FillStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
  Shape.FillStyle() = New_BackStyle
  PropertyChanged "BackStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Shape,Shape,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
  BorderStyle = Shape.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
  Shape.BorderStyle() = New_BorderStyle
  PropertyChanged "BorderStyle"
End Property

Private Sub UserControl_Click()
  RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
  RaiseEvent DblClick
End Sub

Private Sub UserControl_Initialize()
  Dim x As Integer
  
  If RunTime = False Then
    'For x = 0 To 7
      pt(4).Visible = True
    'Next x
  End If
  m_Height = 1000
  m_Width = 2000
  'pt(0).Left = 0:  pt(0).Top = 0
  'pt(1).Left = 100 + m_Width / 2: pt(1).Top = 0
  'pt(2).Left = 100 + m_Width: pt(2).Top = 0
  'pt(3).Left = 100 + m_Width: pt(3).Top = 100 + m_Height / 2
  pt(4).Left = m_Width: pt(4).Top = m_Height
  
  Shape.Left = 0
  Shape.Top = 0
  Shape.Height = m_Height
  Shape.Width = m_Width
  
  UserControl.Width = Shape.Width + 100
  UserControl.Height = Shape.Height + 100

End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
  If Button = 1 And m_RunTime = False Then
    If x > 0 And x < UserControl.Width - 100 And Y > 0 And Y < UserControl.Height - 100 Then
      UserControl.MousePointer = 5
    End If
  
    If x > pt(4).Left And x < pt(4).Left + pt(4).Width And Y > pt(4).Top And Y < pt(4).Top + pt(4).Height Then
      mblnResize = True
      'UserControl.MousePointer = 5
      mx = x: my = Y
      m_MoveShape = False
    End If
  End If
  RaiseEvent MouseDown(Button, Shift, x, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  RaiseEvent MouseMove(Button, Shift, x, Y)
  If x > 0 And x < UserControl.Width - 100 And Y > 0 And Y < UserControl.Height - 100 Then
    UserControl.MousePointer = 5
  ElseIf x > pt(4).Left And x < pt(4).Left + pt(4).Width And Y > pt(4).Top And Y < pt(4).Top + pt(4).Height Then
    UserControl.MousePointer = 8
  Else
    UserControl.MousePointer = 0
  End If
  DoEvents
  
    
  If Button = 1 And m_RunTime = False And mblnResize = True And m_MoveShape = False Then
      If (Shape.Height + Y - my) > 0 Then
        Shape.Height = Shape.Height + (Y - my)
        Shape.Width = Shape.Width + (x - mx)
        pt(4).Top = Shape.Height
        pt(4).Left = Shape.Width
        UserControl.Height = Shape.Height + 100
        UserControl.Width = Shape.Width + 100
        mx = x: my = Y
      End If
  End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
  RaiseEvent MouseUp(Button, Shift, x, Y)
  
  If Button = 1 And m_RunTime = False And mblnResize = True Then
    If (Shape.Height + Y - my) > 0 Then
      Shape.Height = Shape.Height + (Y - my)
      Shape.Width = Shape.Width + (x - mx)
      pt(4).Top = Shape.Height
      pt(4).Left = Shape.Width
      UserControl.Height = Shape.Height + 100
      UserControl.Width = Shape.Width + 100
      mblnResize = False
      UserControl.MousePointer = 0
      m_MoveShape = True
    End If
  End If

End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Shape,Shape,-1,BorderColor
Public Property Get BorderColor() As Long
Attribute BorderColor.VB_Description = "Returns/sets the color of an object's border."
  BorderColor = Shape.BorderColor
End Property

Public Property Let BorderColor(ByVal New_BorderColor As Long)
  Shape.BorderColor() = New_BorderColor
  PropertyChanged "BorderColor"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  Shape.BackColor = PropBag.ReadProperty("BackColor", &H80C0FF)
  UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
  Shape.FillStyle = PropBag.ReadProperty("BackStyle", 1)
  Shape.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
  Shape.BorderColor = PropBag.ReadProperty("BorderColor", -2147483640)
  m_RunTime = PropBag.ReadProperty("RunTime", m_def_RunTime)
  m_MoveShape = PropBag.ReadProperty("MoveShape", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Call PropBag.WriteProperty("BackColor", Shape.BackColor, &H80C0FF)
  Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
  Call PropBag.WriteProperty("BackStyle", Shape.FillStyle, 1)
  Call PropBag.WriteProperty("BorderStyle", Shape.BorderStyle, 1)
  Call PropBag.WriteProperty("BorderColor", Shape.BorderColor, -2147483640)
  Call PropBag.WriteProperty("RunTime", m_RunTime, m_def_RunTime)
  Call PropBag.WriteProperty("MoveShape", m_MoveShape, True)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get RunTime() As Boolean
  RunTime = m_RunTime
End Property

Public Property Let RunTime(ByVal New_RunTime As Boolean)
  m_RunTime = New_RunTime
  PropertyChanged "RunTime"
  Dim x As Integer
  If m_RunTime = True Then
    'For x = 0 To 7
      pt(4).Visible = False
    'Next x
  Else
    'For x = 0 To 7
      pt(4).Visible = True
    'Next x
  End If
  
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
  m_RunTime = m_def_RunTime
  m_MoveShape = False
End Sub


Public Property Get MoveShape() As Boolean
  MoveShape = m_MoveShape
End Property

Public Property Let MoveShape(ByVal vNewValue As Boolean)
  m_MoveShape = vNewValue
  PropertyChanged "MoveShape"
End Property
