VERSION 5.00
Begin VB.UserControl Cylinder 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BackStyle       =   0  'Transparent
   ClientHeight    =   2190
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4275
   ScaleHeight     =   2190
   ScaleWidth      =   4275
   ToolboxBitmap   =   "Cylinder.ctx":0000
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2040
      Top             =   300
   End
   Begin VB.Label lblDesignation 
      BackColor       =   &H000000C0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   2100
      TabIndex        =   0
      Top             =   960
      Width           =   495
   End
   Begin VB.Shape shpExtendSw 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      Height          =   195
      Left            =   660
      Top             =   1080
      Width           =   315
   End
   Begin VB.Shape shpRetractSw 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   240
      Top             =   60
      Width           =   435
   End
   Begin VB.Shape shpPusher 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   1380
      Top             =   420
      Width           =   135
   End
   Begin VB.Shape shpShaft 
      BackStyle       =   1  'Opaque
      Height          =   195
      Left            =   960
      Top             =   180
      Width           =   315
   End
   Begin VB.Shape shpBody 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   1020
      Top             =   1680
      Width           =   795
   End
   Begin VB.Menu mnuCylinder 
      Caption         =   "Cylinder"
      Begin VB.Menu mnuCylOrient 
         Caption         =   "Orientation"
         Begin VB.Menu mnuCylOrientUp 
            Caption         =   "Up"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuCylOrientRight 
            Caption         =   "Right"
         End
         Begin VB.Menu mnuCylOrientDown 
            Caption         =   "Down"
         End
         Begin VB.Menu mnuCylOrientLeft 
            Caption         =   "Left"
         End
      End
      Begin VB.Menu mnuCylExtend 
         Caption         =   "Extend"
      End
   End
End
Attribute VB_Name = "Cylinder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Public Enum COMP_ORIENTATION
'  msNorth = 0
'  msEast = 1
'  msSouth = 2
'  msWest = 3
'End Enum

'used for coordinates of a point
Public Type COORDINATE_PAIR
  X As Single
  Y As Single
End Type

'used for coordinates of a quad
'names are direction of a compass..ie. NE = NorthEast
Public Type QUAD_COORDINATES
  NE As COORDINATE_PAIR
  SE As COORDINATE_PAIR
  SW As COORDINATE_PAIR
  NW As COORDINATE_PAIR
End Type

'Default Property Values:
Const m_def_MousePointer = 0
Const m_def_Stroke = 0
Const m_def_Orientation = 0
Const m_def_Value = 0
Const m_def_CylinderWidth = 400
Const m_def_CylinderLength = 1500
Const m_def_RetractSwitch = 1
Const m_def_ExtendSwitch = 0
Const m_def_Speed = 100
Const m_def_Force = 10
Const m_def_Extend = 0
Const m_def_ForeColor = 0
Const m_def_Enabled = 0
'Const m_def_BackStyle = 0
Const m_def_BorderStyle = 0

'Property Variables:
Dim m_MousePointer As Integer
Dim m_Stroke As Integer
Dim m_Orientation As xOrientation
Dim m_Value As Integer
Dim m_CylinderWidth As Single
Dim m_CylinderLength As Single
Dim m_RetractSwitch As Boolean
Dim m_ExtendSwitch As Boolean
Dim m_Speed As Integer
Dim m_Force As Single
Dim m_Extend As Boolean
Dim m_ForeColor As Long
Dim m_Enabled As Boolean
Dim m_Font As Font
'Dim m_BackStyle As Integer
Dim m_BorderStyle As Integer

'Event Declarations:
Event DblClick()
Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
'Event DblClick()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."

Private Sub lblDesignation_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub lblDesignation_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub mnuCylExtend_Click()
  If m_Extend = True Then
    mnuCylExtend.Caption = "Extend"
    m_Extend = False
    PropertyChanged "Extend"
    Timer1.Enabled = True
  Else
    mnuCylExtend.Caption = "Retract"
    m_Extend = True
    PropertyChanged "Extend"
    Timer1.Enabled = True
  End If
End Sub

Private Sub mnuCylOrientDown_Click()
  mnuCylOrientUp.Checked = False
  mnuCylOrientRight.Checked = False
  mnuCylOrientDown.Checked = True
  mnuCylOrientLeft.Checked = False
  m_Orientation = South
  SetSize
  PropertyChanged "Orientation"
End Sub

Private Sub mnuCylOrientLeft_Click()
  mnuCylOrientUp.Checked = False
  mnuCylOrientRight.Checked = False
  mnuCylOrientDown.Checked = False
  mnuCylOrientLeft.Checked = True
  m_Orientation = West
  SetSize
  PropertyChanged "Orientation"
End Sub

Private Sub mnuCylOrientRight_Click()
  mnuCylOrientUp.Checked = False
  mnuCylOrientRight.Checked = True
  mnuCylOrientDown.Checked = False
  mnuCylOrientLeft.Checked = False
  m_Orientation = East
  SetSize
  PropertyChanged "Orientation"
End Sub

Private Sub mnuCylOrientUp_Click()
  mnuCylOrientUp.Checked = True
  mnuCylOrientRight.Checked = False
  mnuCylOrientDown.Checked = False
  mnuCylOrientLeft.Checked = False
  m_Orientation = North
  SetSize
  PropertyChanged "Orientation"
End Sub

Private Sub mnuExtend_Click()
  If m_Extend = True Then 'need to retract
    
  
  Else 'need to extend
  
  
  End If
 m_Extend = New_Extend
  PropertyChanged "Extend"
  Timer1.Enabled = True
End Sub

'*********************************************************************
' T I M E R  -  Used to animate extending/retracting
'*********************************************************************
Private Sub Timer1_Timer()
  DoEvents
  If m_Extend = True Then  'means extend the cylinder
    If m_Value < m_Stroke Then
      shpRetractSw.BackColor = RGB(0, 100, 0)
      m_Value = m_Value + m_Speed
      If m_Value >= m_Stroke Then
        m_Value = m_Stroke
        shpExtendSw.BackColor = RGB(0, 255, 0)
        Timer1.Enabled = False
      End If
      DrawCylinder
      PropertyChanged "Value"
    End If
  Else 'retracts the cylinder
    If m_Value > 0 Then
      shpExtendSw.BackColor = RGB(0, 100, 0)
      m_Value = m_Value - m_Speed
      If m_Value <= 0 Then
        m_Value = 0
        shpRetractSw.BackColor = RGB(0, 255, 0)
        Timer1.Enabled = False
      End If
      DrawCylinder
      PropertyChanged "Value"
    End If
  End If
End Sub

Private Sub UserControl_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
 UserControl.MousePointer = 5
End Sub

Private Sub UserControl_GotFocus()
  UserControl.DrawStyle = 3
  UserControl.Line (5, 5)-(UserControl.Width - 10, UserControl.Height - 10), vbWhite, B
End Sub

Private Sub UserControl_LostFocus()
  UserControl.DrawStyle = 3
  UserControl.Line (5, 5)-(UserControl.Width - 10, UserControl.Height - 10), vbBlack, B
End Sub

'this stuff happens when this control is added
'to the project form
Private Sub UserControl_Initialize()
  m_Orientation = 1
End Sub

'*********************************************************************
' S E T  S I Z E
'*********************************************************************
Public Sub SetSize()
  Select Case m_Orientation
    Case 0: 'north pointing
      UserControl.Width = m_CylinderWidth * 1.35
      UserControl.Height = m_CylinderLength * 2
      DrawCylinder
    Case 1:
      UserControl.Width = m_CylinderLength * 2
      UserControl.Height = m_CylinderWidth * 1.35
      DrawCylinder
    Case 2:
      UserControl.Width = m_CylinderWidth * 1.35
      UserControl.Height = m_CylinderLength * 2
      DrawCylinder
    Case 3:
      UserControl.Width = m_CylinderLength * 2
      UserControl.Height = m_CylinderWidth * 1.35
      DrawCylinder
  End Select
End Sub

'*********************************************************************
' D R A W  C Y L I N D E R
'*********************************************************************
Private Sub DrawCylinder()
  Dim cl, cw As Integer
  
  'this makes the expressions below much shorter and easier to read
  'lblDesignation.Left = 0: lblDesignation.Top = 0
  'lblDesignation.Width = UserControl.Width
  'lblDesignation.Height = UserControl.Height
  cl = m_CylinderLength
  cw = m_CylinderWidth
  m_Stroke = cl * 0.8
  PropertyChanged "stroke"
  shpPusher.ZOrder 1
  shpShaft.ZOrder 1
  shpBody.ZOrder 0
 ' lblDesignation.ZOrder 0
        
  'draws cylinder based upon orientation and dimensions
  Select Case m_Orientation
    Case 0: 'north pointing
      shpBody.Left = cw * 0.2: shpBody.Width = cw: shpBody.Top = cl: shpBody.Height = cl
      shpRetractSw.Left = 0: shpRetractSw.Top = 2 * cl - cw * 0.4: shpRetractSw.Width = cw * 0.2: shpRetractSw.Height = cw * 0.4
      shpExtendSw.Left = 0: shpExtendSw.Top = cl: shpExtendSw.Width = cw * 0.2: shpExtendSw.Height = cw * 0.4
      shpShaft.Left = cw * 0.6: shpShaft.Width = cw * 0.2
      shpShaft.Top = (cl * 0.9) - m_Value: shpShaft.Height = (cl * 0.1) + m_Value 'cw * 0.2
      shpPusher.Left = shpBody.Left:  shpPusher.Width = shpBody.Width
      shpPusher.Top = (cl * 0.9) - (cw * 0.2) - m_Value: shpPusher.Height = cw * 0.2
     ' lblDesignation.Left = cw * 0.1
     ' lblDesignation.Top = shpBody.Top + cl * 0.6
    Case 1: 'east pointing
      shpBody.Left = 0: shpBody.Width = cl: shpBody.Top = cw * 0.2: shpBody.Height = cw
      shpRetractSw.Left = 0: shpRetractSw.Top = 0: shpRetractSw.Width = cw * 0.4: shpRetractSw.Height = cw * 0.2
      shpExtendSw.Left = cl - cw * 0.4: shpExtendSw.Top = 0: shpExtendSw.Width = cw * 0.4: shpExtendSw.Height = cw * 0.2
      shpShaft.Left = shpBody.Left + cl - 12: shpShaft.Width = (cl * 0.1) + m_Value
      shpShaft.Top = cw * 0.6: shpShaft.Height = cw * 0.2
      shpPusher.Left = shpShaft.Left + shpShaft.Width:  shpPusher.Width = cw * 0.2
      shpPusher.Top = shpBody.Top: shpPusher.Height = cw
     ' lblDesignation.Left = shpBody.Left + cl * 0.1
     ' lblDesignation.Top = shpBody.Top '+ cw * 0.2
    Case 2: 'south pointing
      shpBody.Left = 0: shpBody.Width = cw: shpBody.Top = 0: shpBody.Height = cl
      shpRetractSw.Left = cw: shpRetractSw.Top = 0: shpRetractSw.Width = cw * 0.2: shpRetractSw.Height = cw * 0.4
      shpExtendSw.Left = cw: shpExtendSw.Top = cl - cw * 0.4: shpExtendSw.Width = cw * 0.2: shpExtendSw.Height = cw * 0.4
      shpShaft.Left = cw * 0.35: shpShaft.Width = cw * 0.2
      shpShaft.Top = cl: shpShaft.Height = cw * 0.3 + m_Value
      shpPusher.Left = 0:  shpPusher.Width = cw
      shpPusher.Top = shpBody.Height + shpShaft.Height - 12: shpPusher.Height = cw * 0.2
     ' lblDesignation.Left = shpBody.Left
     ' lblDesignation.Top = cl * 0.4
    Case 3: 'west pointing
      shpBody.Left = cl: shpBody.Width = cl: shpBody.Top = cw * 0.2: shpBody.Height = cw
      shpRetractSw.Left = 2 * cl - cw * 0.4: shpRetractSw.Top = 0: shpRetractSw.Width = cw * 0.4: shpRetractSw.Height = cw * 0.2
      shpExtendSw.Left = cl: shpExtendSw.Top = 0: shpExtendSw.Width = cw * 0.4: shpExtendSw.Height = cw * 0.2
      shpShaft.Left = cl * 0.9 - m_Value: shpShaft.Width = 1.1 * cl - (cl - m_Value)
      shpShaft.Top = cw * 0.6: shpShaft.Height = cw * 0.2
      shpPusher.Left = shpShaft.Left - cw * 0.2: shpPusher.Width = cw * 0.2
      shpPusher.Top = cw * 0.2: shpPusher.Height = cw
     ' lblDesignation.Left = shpBody.Left + cl * 0.5
     ' lblDesignation.Top = shpBody.Top '+ cw * 0.2
  End Select
  
End Sub

'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=UserControl,UserControl,-1,BackColor
'Public Property Get BackColor() As OLE_COLOR
'    BackColor = UserControl.BackColor
'End Property
'
'Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
'    UserControl.BackColor() = New_BackColor
'    PropertyChanged "BackColor"
'End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get ForeColor() As Long
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As Long)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=6,0,0,0
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = m_Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set m_Font = New_Font
    PropertyChanged "Font"
End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=7,0,0,0
'Public Property Get BackStyle() As Integer
'    BackStyle = m_BackStyle
'End Property
'
'Public Property Let BackStyle(ByVal New_BackStyle As Integer)
'    m_BackStyle = New_BackStyle
'    PropertyChanged "BackStyle"
'End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    m_BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
     
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_ForeColor = m_def_ForeColor
    m_Enabled = m_def_Enabled
    Set m_Font = Ambient.Font
'    m_BackStyle = m_def_BackStyle
    m_BorderStyle = m_def_BorderStyle
  m_Orientation = m_def_Orientation
  m_Value = m_def_Value
  m_CylinderWidth = m_def_CylinderWidth
  m_CylinderLength = m_def_CylinderLength
  m_RetractSwitch = m_def_RetractSwitch
  m_ExtendSwitch = m_def_ExtendSwitch
  m_Speed = m_def_Speed
  m_Force = m_def_Force
  m_Extend = m_def_Extend
  m_Stroke = m_def_Stroke
  m_MousePointer = m_def_MousePointer
End Sub


Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent MouseDown(Button, Shift, X, Y)
  
  If Button = 1 And X > 0 And X < UserControl.Width - 100 And Y > 0 And Y < UserControl.Height - 100 Then
    UserControl.MousePointer = 5
  End If
  If Button = vbRightButton Then
    
    PopupMenu mnuCylinder
  End If

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If X > 0 And X < UserControl.Width - 100 And Y > 0 And Y < UserControl.Height - 100 Then
    UserControl.MousePointer = 5
  End If
  
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  
'    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
'    m_BackStyle = PropBag.ReadProperty("BackStyle", m_def_BackStyle)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
  shpBody.BackStyle = PropBag.ReadProperty("BackStyle", 1)
  m_Orientation = PropBag.ReadProperty("Orientation", m_def_Orientation)
  m_Value = PropBag.ReadProperty("Value", m_def_Value)
  m_CylinderWidth = PropBag.ReadProperty("CylinderWidth", m_def_CylinderWidth)
  m_CylinderLength = PropBag.ReadProperty("CylinderLength", m_def_CylinderLength)
  m_RetractSwitch = PropBag.ReadProperty("RetractSwitch", m_def_RetractSwitch)
  m_ExtendSwitch = PropBag.ReadProperty("ExtendSwitch", m_def_ExtendSwitch)
  m_Speed = PropBag.ReadProperty("Speed", m_def_Speed)
  m_Force = PropBag.ReadProperty("Force", m_def_Force)
  m_Extend = PropBag.ReadProperty("Extend", m_def_Extend)
  m_Stroke = PropBag.ReadProperty("Stroke", m_def_Stroke)
  lblDesignation.Caption = PropBag.ReadProperty("Designation", "")
  m_MousePointer = PropBag.ReadProperty("MousePointer", m_def_MousePointer)
End Sub

Private Sub UserControl_Resize()
  If UserControl.Height > 0 Then
  'UserControl.ScaleTop = UserControl.Height
  'UserControl.ScaleHeight = -UserControl.Height
  End If
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

'    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
'    Call PropBag.WriteProperty("BackStyle", m_BackStyle, m_def_BackStyle)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
  Call PropBag.WriteProperty("BackStyle", shpBody.BackStyle, 1)
  Call PropBag.WriteProperty("Orientation", m_Orientation, m_def_Orientation)
  Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
  Call PropBag.WriteProperty("CylinderWidth", m_CylinderWidth, m_def_CylinderWidth)
  Call PropBag.WriteProperty("CylinderLength", m_CylinderLength, m_def_CylinderLength)
  Call PropBag.WriteProperty("RetractSwitch", m_RetractSwitch, m_def_RetractSwitch)
  Call PropBag.WriteProperty("ExtendSwitch", m_ExtendSwitch, m_def_ExtendSwitch)
  Call PropBag.WriteProperty("Speed", m_Speed, m_def_Speed)
  Call PropBag.WriteProperty("Force", m_Force, m_def_Force)
  Call PropBag.WriteProperty("Extend", m_Extend, m_def_Extend)
  Call PropBag.WriteProperty("Stroke", m_Stroke, m_def_Stroke)
  Call PropBag.WriteProperty("Designation", lblDesignation.Caption, "")
  Call PropBag.WriteProperty("MousePointer", m_MousePointer, m_def_MousePointer)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=shpBody,shpBody,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
  BackColor = shpBody.BackColor
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=shpBody,shpBody,-1,BackStyle
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
  BackStyle = shpBody.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
  shpBody.BackStyle() = New_BackStyle
  PropertyChanged "BackStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,1
Public Property Get Orientation() As xOrientation
  Orientation = m_Orientation
End Property

Public Property Let Orientation(ByVal New_Orientation As xOrientation)
  m_Orientation = New_Orientation
  SetSize
  PropertyChanged "Orientation"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,1,0,0
Public Property Get value() As Integer
  value = m_Value
End Property

Public Property Let value(ByVal New_Value As Integer)
  If Ambient.UserMode Then Err.Raise 382
  m_Value = New_Value
  PropertyChanged "Value"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12,0,0,0
Public Property Get CylinderWidth() As Single
  CylinderWidth = m_CylinderWidth
End Property

Public Property Let CylinderWidth(ByVal New_CylinderWidth As Single)
  m_CylinderWidth = New_CylinderWidth
  PropertyChanged "CylinderWidth"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12,0,0,0
Public Property Get CylinderLength() As Single
  CylinderLength = m_CylinderLength
End Property

Public Property Let CylinderLength(ByVal New_CylinderLength As Single)
  m_CylinderLength = New_CylinderLength
  PropertyChanged "CylinderLength"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,1,0,1
Public Property Get RetractSwitch() As Boolean
  RetractSwitch = m_RetractSwitch
End Property

Public Property Let RetractSwitch(ByVal New_RetractSwitch As Boolean)
  If Ambient.UserMode Then Err.Raise 382
  m_RetractSwitch = New_RetractSwitch
  PropertyChanged "RetractSwitch"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,1,0,1
Public Property Get ExtendSwitch() As Boolean
  ExtendSwitch = m_ExtendSwitch
End Property

Public Property Let ExtendSwitch(ByVal New_ExtendSwitch As Boolean)
  If Ambient.UserMode Then Err.Raise 382
  m_ExtendSwitch = New_ExtendSwitch
  PropertyChanged "ExtendSwitch"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,100
Public Property Get Speed() As Integer
  Speed = m_Speed
End Property

Public Property Let Speed(ByVal New_Speed As Integer)
  m_Speed = New_Speed
  PropertyChanged "Speed"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12,0,0,10
Public Property Get Force() As Single
  Force = m_Force
End Property

Public Property Let Force(ByVal New_Force As Single)
  m_Force = New_Force
  PropertyChanged "Force"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get Extend() As Boolean
  Extend = m_Extend
End Property

Public Property Let Extend(ByVal New_Extend As Boolean)
  m_Extend = New_Extend
  PropertyChanged "Extend"
  Timer1.Enabled = True
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,1,0,0
Public Property Get Stroke() As Integer
  Stroke = m_Stroke
End Property

Public Property Let Stroke(ByVal New_Stroke As Integer)
  If Ambient.UserMode Then Err.Raise 382
  m_Stroke = New_Stroke
  PropertyChanged "Stroke"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblDesignation,lblDesignation,-1,Caption
Public Property Get Designation() As String
Attribute Designation.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
  'Designation = lblDesignation.Caption
End Property

Public Property Let Designation(ByVal New_Designation As String)
 ' lblDesignation.Caption() = New_Designation
 ' PropertyChanged "Designation"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get MousePointer() As Integer
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
  MousePointer = m_MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
  m_MousePointer = New_MousePointer
  PropertyChanged "MousePointer"
End Property

