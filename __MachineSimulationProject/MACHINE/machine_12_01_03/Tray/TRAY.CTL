VERSION 5.00
Begin VB.UserControl Tray 
   BackStyle       =   0  'Transparent
   ClientHeight    =   615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   600
   ScaleHeight     =   615
   ScaleWidth      =   600
   ToolboxBitmap   =   "tray.ctx":0000
   Begin VB.Shape shape2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   120
      Shape           =   3  'Circle
      Top             =   180
      Width           =   315
   End
   Begin VB.Shape shape1 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      Height          =   555
      Left            =   0
      Top             =   0
      Width           =   555
   End
End
Attribute VB_Name = "Tray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Default Property Values:
Const m_def_ForeColor = 0
Const m_def_Enabled = 0
Const m_def_BackStyle = vbRed
Const m_def_BorderStyle = 0
'Property Variables:
Dim m_ForeColor As Long
Dim m_Enabled As Boolean
Dim m_Font As Font
Dim m_BackStyle As Integer
Dim m_BorderStyle As Integer
'Event Declarations:
Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick()
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
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

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Shape1,Shape1,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
  BackColor = shape1.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
  shape1.BackColor() = New_BackColor
  PropertyChanged "BackColor"
End Property

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

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
  BackStyle = m_BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
  m_BackStyle = New_BackStyle
  PropertyChanged "BackStyle"
End Property

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
  m_BackStyle = m_def_BackStyle
  m_BorderStyle = m_def_BorderStyle
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent MouseDown(Button, Shift, X, Y)
  UserControl.MousePointer = 5
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  UserControl.MousePointer = 5
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  shape1.BackColor = PropBag.ReadProperty("BackColor", &H404040)
  m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
  m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
  Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
  m_BackStyle = PropBag.ReadProperty("BackStyle", m_def_BackStyle)
  m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Call PropBag.WriteProperty("BackColor", shape1.BackColor, &H404040)
  Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
  Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
  Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
  Call PropBag.WriteProperty("BackStyle", m_BackStyle, m_def_BackStyle)
  Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
End Sub

