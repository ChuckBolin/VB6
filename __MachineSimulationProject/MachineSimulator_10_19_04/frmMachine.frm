VERSION 5.00
Begin VB.Form frmMachine 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Machine (Top View)"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   9510
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFFF&
      Height          =   4215
      Left            =   30
      ScaleHeight     =   4155
      ScaleWidth      =   9255
      TabIndex        =   0
      Top             =   0
      Width           =   9315
      Begin VB.Timer tmrUpdate 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   8250
         Top             =   2100
      End
      Begin VB.Frame frame1 
         Caption         =   "Manual Operation"
         Height          =   1185
         Left            =   60
         TabIndex        =   1
         Top             =   2910
         Width           =   8985
         Begin VB.OptionButton optRetract 
            Caption         =   "Retract"
            Height          =   315
            Left            =   1710
            TabIndex        =   4
            Top             =   690
            Value           =   -1  'True
            Width           =   915
         End
         Begin VB.OptionButton optExtend 
            Caption         =   "Extend"
            Height          =   255
            Left            =   1710
            TabIndex        =   3
            Top             =   360
            Width           =   915
         End
         Begin VB.HScrollBar HScroll1 
            Height          =   285
            LargeChange     =   10
            Left            =   240
            Max             =   425
            TabIndex        =   2
            Top             =   360
            Width           =   1125
         End
      End
      Begin VB.Shape shpS21 
         BackColor       =   &H00000080&
         BackStyle       =   1  'Opaque
         Height          =   100
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Width           =   100
      End
      Begin VB.Shape shpS20 
         BackColor       =   &H00000080&
         BackStyle       =   1  'Opaque
         Height          =   100
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Width           =   100
      End
      Begin VB.Shape shpS15 
         BackColor       =   &H00000080&
         BackStyle       =   1  'Opaque
         Height          =   100
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Width           =   100
      End
      Begin VB.Shape shpS14 
         BackColor       =   &H00000080&
         BackStyle       =   1  'Opaque
         Height          =   100
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Width           =   100
      End
      Begin VB.Shape shpS12 
         BackColor       =   &H00000080&
         BackStyle       =   1  'Opaque
         Height          =   100
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Width           =   100
      End
      Begin VB.Shape shpS11 
         BackColor       =   &H00000080&
         BackStyle       =   1  'Opaque
         Height          =   100
         Left            =   3210
         Shape           =   3  'Circle
         Top             =   1590
         Width           =   100
      End
      Begin VB.Shape shpZ7 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         Height          =   105
         Left            =   0
         Top             =   0
         Width           =   105
      End
      Begin VB.Shape shpZ3 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         Height          =   105
         Left            =   0
         Top             =   0
         Width           =   105
      End
      Begin VB.Shape shpZ1 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         Height          =   105
         Left            =   3300
         Top             =   960
         Width           =   105
      End
      Begin VB.Shape shpState 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  'Opaque
         Height          =   75
         Index           =   0
         Left            =   1890
         Shape           =   1  'Square
         Top             =   1860
         Width           =   75
      End
      Begin VB.Shape shpPwr 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   75
         Index           =   0
         Left            =   1740
         Shape           =   3  'Circle
         Top             =   1230
         Width           =   75
      End
      Begin VB.Shape shpWT 
         BackColor       =   &H00808080&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         Height          =   380
         Left            =   780
         Top             =   2370
         Width           =   380
      End
   End
End
Attribute VB_Name = "frmMachine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'enumerations for machines
Private Enum MOTOR_POS
  NW = 0
  NE
  SE
  SW
End Enum

Private Enum ORIENTATION
  NORTH = 0
  EAST
  SOUTH
  WEST
End Enum

Private Sub Form_Load()
    'frmMachine.Left = frmCP.Width + frmCP.Left
    'frmMachine.Top = 0
    'frmMachine.Height = frmCP.Height
    pic.Top = 0: pic.Left = 0
    pic.Height = frmMachine.Height - 400
    pic.Width = frmMachine.Width - 100
    
    'initial layout
    DrawConveyor 100, 1000, 5000, 100, NW  'top conveyor
    
    DrawConveyor 4600, 1400, 4000, 100, NE 'bottom conveyor
    
    DrawCylinder 4700, 275, 600, 200, SOUTH 'top cylinder Z1
    shpZ1.Top = 875: shpZ1.Left = 4750: shpZ1.Width = 100: shpZ1.Height = 100
    shpS11.Left = 4750: shpS11.Top = 280
    shpS12.Left = 4750: shpS12.Top = 780
    
    DrawCylinder 3900, 1500, 600, 200, EAST 'left cylinder Z3
    shpZ3.Top = 1550: shpZ3.Left = 4500: shpZ3.Height = 100: shpZ3.Width = 100
    shpS14.Left = 3920: shpS14.Top = 1550
    shpS15.Left = 4390: shpS15.Top = 1550
    
    DrawCylinder 7000, 350, 600, 200, SOUTH 'right cylinder Z7
    shpZ7.Top = 950: shpZ7.Left = 7050: shpZ7.Height = 100: shpZ7.Width = 100
    shpS20.Left = 7050: shpS20.Top = 370
    shpS21.Left = 7050: shpS21.Top = 860
        
    DrawSeparator 3000, 1100, 300, 200, EAST 'left sep
    DrawSeparator 6600, 1500, 300, 200, EAST 'middle sep
    DrawSeparator 7200, 1500, 300, 200, EAST 'right sep
    
    DrawProx 3050, 650, SOUTH 'S10
    shpPwr(0).Left = 3050: shpPwr(0).Top = 650 ': shpPwr(0).Width = 75: shpPwr(0).Height = 75
    shpState(0).Left = 3050: shpState(0).Top = 875 ': shpState(0).Width = 75: shpState(0).Height = 75
    
    DrawProx 5100, 1150, WEST 'S13
    DrawProx 4900, 1850, NORTH 's16
    DrawProx 6650, 1850, NORTH 'S18
    DrawProx 7250, 1850, NORTH 's19
    DrawClamp 6950, 1850, NORTH 'Z5
    
    'position tray
    shpWT.Left = 100
    shpWT.Top = 1020
    
    tmrUpdate_Timer
End Sub

'****************************************************** DrawClamp
Private Sub DrawClamp(X As Single, Y As Single, o As ORIENTATION)
  Select Case o
    Case NORTH:
      pic.Line (X, Y)-(X + 300, Y + 200), , B 'main part of clamp
      pic.Line (X + 25, Y - 45)-(X + 275, Y), , BF 'clamping part
      pic.Line (X - 200, Y + 75)-(X, Y + 100), , B
    Case EAST:
    
    Case SOUTH:
    
    Case WEST:
  
  End Select
End Sub

'***************************************************** DrawProx
Private Sub DrawProx(X As Single, Y As Single, o As ORIENTATION)
  Select Case o
    Case NORTH:
      pic.Line (X, Y)-(X + 50, Y + 300), , B
      pic.Line (X, Y)-(X + 50, Y + 50), , BF
    Case EAST:
      pic.Line (X, Y)-(X + 300, Y + 50), , B
      pic.Line (X + 250, Y)-(X + 300, Y + 50), , BF
    Case SOUTH:
      pic.Line (X, Y)-(X + 50, Y + 300), , B
      pic.Line (X, Y + 250)-(X + 50, Y + 300), , BF
    Case WEST:
      pic.Line (X, Y)-(X + 300, Y + 50), , B
      pic.Line (X, Y)-(X + 50, Y + 50), , BF
  End Select
End Sub
'***************************************************** DrawSeparator
Private Sub DrawSeparator(X As Single, Y As Single, length As Single, w As Single, o As ORIENTATION)
  Select Case o
    Case NORTH Or SOUTH
    
    Case EAST: ' Or WEST
      pic.Line (X, Y)-(X + length, Y + w), , B 'draws body
      pic.Line (X + length / 4, Y + w / 4)-(X + length * 3 / 4, Y + w * 3 / 4), , BF
      
    
  End Select
End Sub

'**************************************************** DrawCylinder
Private Sub DrawCylinder(X As Single, Y As Single, length As Single, w As Single, o As ORIENTATION)
  Select Case o
    Case NORTH:
      
    Case EAST:
      pic.Line (X, Y)-(X + length, Y + w), , B 'draws body
      'pic.Line (X + length, Y + w / 4)-(X + length + w / 2, Y + w * 3 / 4), , B 'draws shaft
      pic.Line (X, Y)-(X + w / 2, Y + w), , BF 'draws retracted reed switch
      pic.Line (X + length - w / 2, Y)-(X + length, Y + w), , BF 'draws extended switch
    Case SOUTH:
      pic.Line (X, Y)-(X + w, Y + length), , B  'draws body
     ' pic.Line (X + w / 4, Y + length)-(X + w * 3 / 4, Y + length + w / 2), , B 'draws shaft retracted
      pic.Line (X, Y)-(X + w, Y + w / 2), , BF 'retracted switch
      pic.Line (X, Y + length - w / 2)-(X + w, Y + length), , BF 'extended switch
    Case WEST:
    
  End Select
End Sub

'**************************************************** DrawConveyor
'x,y = top-left corner of conveyor
'length = how long is conveyor, tw = track width
Private Sub DrawConveyor(X As Single, Y As Single, length As Single, tw As Single, m As MOTOR_POS)
  pic.Line (X, Y)-(X + length, Y + tw), , B
  pic.Line (X, Y + 3 * tw)-(X + length, Y + 4 * tw), , B
  
  Select Case m
    Case NW:
      pic.Circle (X + 2 * tw, Y - 2 * tw), 2 * tw
    Case NE:
      pic.Circle (X + length - 2 * tw, Y - 2 * tw), 2 * tw
    Case SE:
      pic.Circle (X + length - 2 * tw, Y + 6 * tw), 2 * tw
    Case SW:
      pic.Circle (X + 2 * tw, Y + 6 * tw), 2 * tw
  End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If frmMain.mnuViewMachine.Checked = True Then frmMain.mnuViewMachine.Checked = False
End Sub

Private Sub HScroll1_Change()
  shpZ1.Height = 100 + HScroll1.Value
  shpZ3.Width = 100 + HScroll1.Value
  shpZ7.Height = 100 + HScroll1.Value
End Sub

Private Sub HScroll1_Scroll()
  HScroll1_Change
End Sub

Private Sub optExtend_Click()
  tmrUpdate.Enabled = True
End Sub

Private Sub optRetract_Click()
  tmrUpdate.Enabled = True
End Sub

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'  frmMachine.Caption = X & " " & Y
  
  'displays tool text tip for each component
  If Hotspot(X, Y, 3075, 800) Then
    pic.ToolTipText = "S10 Prox Sw"
  ElseIf Hotspot(X, Y, 5250, 1175) Then
    pic.ToolTipText = "S13 Prox Sw"
  ElseIf Hotspot(X, Y, 4925, 2025) Then
    pic.ToolTipText = "S16 Prox Sw"
  ElseIf Hotspot(X, Y, 6660, 2085) Then
    pic.ToolTipText = "S17 Prox Sw"
  ElseIf Hotspot(X, Y, 6855, 1920) Then
    pic.ToolTipText = "S18 Prox Sw"
  ElseIf Hotspot(X, Y, 7275, 1995) Then
    pic.ToolTipText = "S19 Prox Sw"
  ElseIf Hotspot(X, Y, 7125, 1935) Then
    pic.ToolTipText = "Z5 Clamp"
  ElseIf Hotspot(X, Y, 6735, 1575) Then
    pic.ToolTipText = "Z4 Stop Gate"
  ElseIf Hotspot(X, Y, 7350, 1575) Then
    pic.ToolTipText = "Z6 Stop Gate"
  ElseIf Hotspot(X, Y, 315, 780) Then
    pic.ToolTipText = "MOT1 Conveyor Motor"
  ElseIf Hotspot(X, Y, 8400, 1185) Then
    pic.ToolTipText = "MOT2 Conveyor Motor"
  ElseIf Hotspot(X, Y, 3135, 1185) Then
    pic.ToolTipText = "Z2 Stop Gate"
  ElseIf Hotspot(X, Y, 4800, 285) Then
    pic.ToolTipText = "S11 Magnetic Sw"
  ElseIf Hotspot(X, Y, 4785, 555) Then
    pic.ToolTipText = "Z1 Cylinder"
  ElseIf Hotspot(X, Y, 4785, 795) Then
    pic.ToolTipText = "S12 Magnetic Sw"
  ElseIf Hotspot(X, Y, 3930, 1575) Then
    pic.ToolTipText = "S14 Magnetic Sw"
  ElseIf Hotspot(X, Y, 4215, 1590) Then
    pic.ToolTipText = "Z3 Cylinder"
  ElseIf Hotspot(X, Y, 4440, 1575) Then
    pic.ToolTipText = "S15 Magnetic Sw"
  ElseIf Hotspot(X, Y, 7095, 570) Then
    pic.ToolTipText = "S20 Magnetic Sw"
  ElseIf Hotspot(X, Y, 7095, 825) Then
    pic.ToolTipText = "Z7 Cylinder"
  ElseIf Hotspot(X, Y, 7095, 1080) Then
    pic.ToolTipText = "S21 Magnetic Sw"
  Else
    pic.ToolTipText = ""
  End If
End Sub

'returns true if mouseclick is inside circle of diameter 120
Private Function Hotspot(x1 As Single, y1 As Single, x2 As Single, y2 As Single) As Boolean
  Hotspot = False
  If Sqr((x1 - x2) ^ 2 + (y1 - y2) ^ 2) < 150 Then Hotspot = True
End Function

Private Sub tmrUpdate_Timer()
  If optExtend.Value = True And HScroll1.Value < 405 Then
    HScroll1.Value = HScroll1.Value + 20
    If HScroll1.Value > 405 Then
      HScroll1.Value = 425
      'tmrUpdate.Enabled = False
    End If
  ElseIf optRetract.Value = True And HScroll1.Value > 20 Then
    HScroll1.Value = HScroll1.Value - 20
    If HScroll1.Value < 20 Then
      HScroll1.Value = 0
      'tmrUpdate.Enabled = False
    End If
  End If
  
  
  If HScroll1.Value < 25 Then
    e(S11_Z1_RETRACT) = True
    shpS11.BackColor = vbRed
    e(S14_Z3_RETRACT) = True
    shpS14.BackColor = vbRed
    e(S20_Z7_RETRACT) = True
    shpS20.BackColor = vbRed
  Else
    e(S11_Z1_RETRACT) = False
    shpS11.BackColor = RGB(100, 0, 0)
    e(S14_Z3_RETRACT) = False
    shpS14.BackColor = RGB(100, 0, 0)
    e(S20_Z7_RETRACT) = False
    shpS20.BackColor = RGB(100, 0, 0)
  End If
  
  If HScroll1.Value > 405 Then
    e(S12_Z1_EXTEND) = True
    shpS12.BackColor = vbRed
    e(S15_Z3_EXTEND) = True
    shpS15.BackColor = vbRed
    e(S21_Z7_EXTEND) = True
    shpS21.BackColor = vbRed
  Else
    e(S12_Z1_EXTEND) = False
    shpS12.BackColor = RGB(100, 0, 0)
    e(S15_Z3_EXTEND) = False
    shpS15.BackColor = RGB(100, 0, 0)
    e(S21_Z7_EXTEND) = False
    shpS21.BackColor = RGB(100, 0, 0)
  End If
  
  If v(V_BE_24V) = True Then
    shpPwr(0).BackColor = vbRed
  Else
    shpPwr(0).BackColor = RGB(100, 0, 0)
  End If
  
  If e(S10_PROX) = True Then
    shpState(0).BackColor = vbGreen
  Else
    shpState(0).BackColor = RGB(0, 100, 0)
  End If
  'frmMachine.Caption = e(S11_Z1_RETRACT)
End Sub
