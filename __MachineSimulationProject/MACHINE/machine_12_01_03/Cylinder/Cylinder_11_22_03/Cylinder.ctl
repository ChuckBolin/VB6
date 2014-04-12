VERSION 5.00
Begin VB.UserControl Cylinder 
   BackColor       =   &H00C0C0FF&
   BackStyle       =   0  'Transparent
   ClientHeight    =   660
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2070
   ScaleHeight     =   660
   ScaleWidth      =   2070
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1800
      Top             =   240
   End
   Begin VB.Label lblDesignation 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   300
      Width           =   915
   End
   Begin VB.Shape shpExtendSw 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   195
      Left            =   900
      Top             =   240
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
      Left            =   1560
      Top             =   240
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
      Left            =   900
      Top             =   120
      Width           =   795
   End
End
Attribute VB_Name = "Cylinder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Enum COMP_ORIENTATION
  msNorth = 0
  msEast = 1
  msSouth = 2
  msWest = 3
End Enum

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
Const m_def_Stroke = 0
Const m_def_Orientation = 1
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
Dim m_Stroke As Integer
Dim m_Orientation As Variant
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

Private Sub Timer1_Timer()
  If m_Extend = True Then  'means extend the cylinder
    If m_Value < m_Stroke Then
      shpRetractSw.BackColor = vbBlack
      m_Value = m_Value + m_Speed
      If m_Value >= m_Stroke Then
        m_Value = m_Stroke
        shpExtendSw.BackColor = vbGreen
        Timer1.Enabled = False
      End If
      DrawCylinder
      PropertyChanged "Value"
    End If
  Else 'retracts the cylinder
    If m_Value > 0 Then
      shpExtendSw.BackColor = vbBlack
      m_Value = m_Value - m_Speed
      If m_Value <= 0 Then
        m_Value = 0
        shpRetractSw.BackColor = vbGreen
        Timer1.Enabled = False
      End If
      DrawCylinder
      PropertyChanged "Value"
    End If
  End If
End Sub

'this stuff happens when this control is added
'to the project form
Private Sub UserControl_Initialize()
  m_Orientation = 1
End Sub

Public Sub SetSize()
  Select Case m_Orientation
    Case 0: 'north pointing
    
    Case 1:
      UserControl.Width = m_CylinderLength * 2
      UserControl.Height = m_CylinderWidth * 1.35
      DrawCylinder
    Case 2:
    
    
    Case 3:
  End Select
End Sub

'*********************************************************************
' D R A W  C Y L I N D E R
'*********************************************************************
Private Sub DrawCylinder()
  Dim cl, cw As Integer
  
  'this makes the expressions below much shorter and easier to read
  cl = m_CylinderLength
  cw = m_CylinderWidth
  m_Stroke = cl * 0.8
  PropertyChanged "stroke"
  
  'draws cylinder based upon orientation and dimensions
  Select Case m_Orientation
    Case 0: 'north pointing
    
    Case 1: 'east pointing
      shpBody.Left = 0: shpBody.Width = cl: shpBody.Top = cw * 0.2: shpBody.Height = cw
      shpRetractSw.Left = 0: shpRetractSw.Top = 0: shpRetractSw.Width = cw * 0.4: shpRetractSw.Height = cw * 0.3
      shpExtendSw.Left = cl - cw * 0.4: shpExtendSw.Top = 0: shpExtendSw.Width = cw * 0.4: shpExtendSw.Height = cw * 0.3
      shpShaft.Left = shpBody.Left + cl - 12: shpShaft.Width = (cl * 0.1) + m_Value
      shpShaft.Top = cw * 0.6: shpShaft.Height = cw * 0.2
      shpPusher.Left = shpShaft.Left + shpShaft.Width:  shpPusher.Width = cw * 0.2
      shpPusher.Top = shpBody.Top: shpPusher.Height = cw
      lblDesignation.Left = shpBody.Left + cl * 0.1
      lblDesignation.Top = shpBody.Top + cw * 0.2
      
    Case 2: 'south pointing
    
    Case 3: 'west pointing
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
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent MouseDown(Button, Shift, X, Y)

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
Public Property Get Orientation() As Variant
  Orientation = m_Orientation
End Property

Public Property Let Orientation(ByVal New_Orientation As Variant)
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
  Designation = lblDesignation.Caption
End Property

Public Property Let Designation(ByVal New_Designation As String)
  lblDesignation.Caption() = New_Designation
  PropertyChanged "Designation"
End Property

