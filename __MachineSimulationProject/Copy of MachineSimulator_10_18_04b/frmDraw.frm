VERSION 5.00
Begin VB.Form frmDraw 
   Caption         =   "Electrical Schematic"
   ClientHeight    =   6390
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9195
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6390
   ScaleWidth      =   9195
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4215
      Left            =   0
      ScaleHeight     =   4215
      ScaleWidth      =   6675
      TabIndex        =   2
      Top             =   0
      Width           =   6675
   End
   Begin VB.VScrollBar vsb 
      Height          =   5175
      LargeChange     =   1000
      Left            =   8940
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
   Begin VB.HScrollBar hsb 
      Height          =   255
      LargeChange     =   1000
      Left            =   0
      TabIndex        =   0
      Top             =   5760
      Width           =   7035
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Zoom"
      Begin VB.Menu mnuZoom25 
         Caption         =   "25%"
      End
      Begin VB.Menu mnuZoom50 
         Caption         =   "50%"
      End
      Begin VB.Menu mnuZoom100 
         Caption         =   "100%"
      End
      Begin VB.Menu mnu200 
         Caption         =   "200%"
      End
      Begin VB.Menu mnu300 
         Caption         =   "300%"
      End
      Begin VB.Menu mnu400 
         Caption         =   "400%"
      End
   End
End
Attribute VB_Name = "frmDraw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Z As Single 'zoom scale multiplier

Private Sub Form_Load()
  Z = 2
  pic.FontSize = 26
  mnu300.Checked = True
End Sub

Private Sub Form_Resize()
 vsb.Left = frmDraw.Width - vsb.Width - 120
  If frmDraw.Height > 660 Then vsb.Height = frmDraw.Height - 660
  hsb.Top = frmDraw.Height - hsb.Height - 400
  hsb.Width = frmDraw.Width - 400
  pic.Width = vsb.Left
  If hsb.Top > 0 Then pic.Height = hsb.Top
  DrawSchematic
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If frmMain.mnuViewSchematic.Checked = True Then frmMain.mnuViewSchematic.Checked = False
End Sub

'********************************* DrawSchematic
Private Sub DrawSchematic()
  pic.Cls
  DrawDisconnect pic, 1000 - hsb.Value, 1000 - vsb.Value
  DrawThreePhaseFuse pic, 1900 - hsb.Value, 1000 - vsb.Value, "F1", "0"
  DrawLine pic, 2700 - hsb.Value, 1000 - vsb.Value, 4000 - hsb.Value, 1000 - vsb.Value
  DrawLine pic, 2700 - hsb.Value, 1300 - vsb.Value, 4000 - hsb.Value, 1300 - vsb.Value
  DrawLine pic, 2700 - hsb.Value, 1600 - vsb.Value, 4000 - hsb.Value, 1600 - vsb.Value
  DrawThreePhaseFuse pic, 4000 - hsb.Value, 1000 - vsb.Value, "F3", "1"
  DrawThreePhaseFuse pic, 4000 - hsb.Value, 2000 - vsb.Value, "F4", "1"
  DrawTerminal pic, 3800 - hsb.Value, 1000 - vsb.Value
  DrawTerminal pic, 3700 - hsb.Value, 1300 - vsb.Value
  DrawTerminal pic, 3600 - hsb.Value, 1600 - vsb.Value
  DrawLine pic, 3800 - hsb.Value, 1000 - vsb.Value, 3800 - hsb.Value, 2000 - vsb.Value
  DrawLine pic, 3700 - hsb.Value, 1300 - vsb.Value, 3700 - hsb.Value, 2300 - vsb.Value
  DrawLine pic, 3600 - hsb.Value, 1600 - vsb.Value, 3600 - hsb.Value, 2600 - vsb.Value
  
  
  
End Sub


'*********************************************************************
'*********************************
'********************************* D R A W I N G   R O U T I N E S
'*********************************
'*********************************************************************
Private Sub DrawDisconnect(p As PictureBox, x As Single, y As Single)
  Dim nFont As Integer
  nFont = p.FontSize
  
  p.FontSize = nFont - 8
  DrawText p, x - 700, y, "480VAC"
  DrawText p, x - 700, y + 200, "3 Phase"
  DrawText p, x - 700, y + 400, "60 Hz"
  DrawText p, x + 375, y - 300, "Q0"
  DrawText p, x, y - 150, "L1"
  DrawText p, x, y + 150, "L2"
  DrawText p, x, y + 450, "L3"
  
  p.FontSize = p.FontSize / 2
  DrawText p, x + 275, y - 120, "1"
  DrawText p, x + 600, y - 120, "2"
  DrawText p, x + 275, y + 180, "3"
  DrawText p, x + 600, y + 180, "4"
  DrawText p, x + 275, y + 480, "5"
  DrawText p, x + 600, y + 480, "6"

  DrawNOSwitch p, x, y
  DrawNOSwitch p, x, y + 300
  DrawNOSwitch p, x, y + 600
  p.DrawStyle = 1
  p.Line ((x + 450) * Z, y * Z)-((x + 450) * Z, (y + 550) * Z)
  p.DrawStyle = 0
  p.FontSize = nFont
End Sub

Private Sub DrawThreePhaseFuse(p As PictureBox, x As Single, y As Single, sName As String, sWire As String)
  Dim nFont As Integer
  nFont = p.FontSize
  
  p.FontSize = nFont - 8
  DrawText p, x + 350, y - 200, sName & "A"
  DrawText p, x + 350, y + 100, sName & "B"
  DrawText p, x + 350, y + 400, sName & "C"
  
  If Len(sWire) > 0 Then
    DrawText p, x, y - 150, sWire & "L1"
    DrawText p, x, y + 150, sWire & "L2"
    DrawText p, x, y + 450, sWire & "L3"
  End If
  
  p.FontSize = nFont / 2
  DrawText p, x + 250, y - 100, "1"
  DrawText p, x + 625, y - 100, "2"
  DrawText p, x + 250, y + 200, "1"
  DrawText p, x + 625, y + 200, "2"
  DrawText p, x + 250, y + 500, "1"
  DrawText p, x + 625, y + 500, "2"
    
  DrawFuse p, x, y
  DrawFuse p, x, y + 300
  DrawFuse p, x, y + 600
  p.FontSize = nFont
  
End Sub

'*********************************************************************
'*********************************  D R A W I N G  P R I M I T I V E S
'*********************************************************************
Private Sub DrawText(p As PictureBox, x As Single, y As Single, s As String)
  p.CurrentX = x * Z
  p.CurrentY = y * Z
  p.Print s
End Sub

Private Sub DrawTerminal(p As PictureBox, x As Single, y As Single)
  p.Circle (x * Z, y * Z), Z * 10
End Sub

Private Sub DrawLine(p As PictureBox, x1 As Single, y1 As Single, x2 As Single, y2 As Single)
  p.Line (x1 * Z, y1 * Z)-(x2 * Z, y2 * Z)
End Sub

Private Sub DrawNOSwitch(p As PictureBox, x As Single, y As Single)
  p.Line (x * Z, y * Z)-((x + 300) * Z, y * Z)
  p.Circle ((x + 300) * Z, y * Z), 25 * Z
  p.Line ((x + 300) * Z, y * Z)-((x + 550) * Z, (y - 100) * Z)
  p.Circle ((x + 600) * Z, y * Z), 25 * Z
  p.Line ((x + 600) * Z, y * Z)-((x + 900) * Z, y * Z)
End Sub

Private Sub DrawFuse(p As PictureBox, x As Single, y As Single)
  p.Line (x * Z, y * Z)-((x + 300) * Z, y * Z)
  p.Line ((x + 300) * Z, (y - 50) * Z)-((x + 600) * Z, (y + 50) * Z), , B
  p.Line ((x + 600) * Z, y * Z)-((x + 900) * Z, y * Z)
End Sub

Private Sub hsb_Change()
  DrawSchematic
End Sub

Private Sub hsb_Scroll()
  hsb_Change
End Sub

'********************************* ZOOM
Private Sub mnuZoom25_Click()
  Z = 0.25
  pic.FontSize = 3
  DrawSchematic
  mnuZoom25.Checked = True
  mnuZoom50.Checked = False
  mnuZoom100.Checked = False
  mnu200.Checked = False
  mnu300.Checked = False
  mnu400.Checked = False
End Sub

Private Sub mnuZoom50_Click()
  Z = 0.5
  pic.FontSize = 6
  DrawSchematic
  mnuZoom25.Checked = False
  mnuZoom50.Checked = True
  mnuZoom100.Checked = False
  mnu200.Checked = False
  mnu300.Checked = False
  mnu400.Checked = False
End Sub

Private Sub mnuZoom100_Click()
  Z = 1
  pic.FontSize = 10
  DrawSchematic
  mnuZoom25.Checked = False
  mnuZoom50.Checked = False
  mnuZoom100.Checked = True
  mnu200.Checked = False
  mnu300.Checked = False
  mnu400.Checked = False
End Sub

Private Sub mnu200_Click()
  Z = 2
  pic.FontSize = 18
  DrawSchematic
  mnuZoom25.Checked = False
  mnuZoom50.Checked = False
  mnuZoom100.Checked = False
  mnu200.Checked = True
  mnu300.Checked = False
  mnu400.Checked = False
End Sub

Private Sub mnu300_Click()
  Z = 3
  pic.FontSize = 26
  DrawSchematic
  mnuZoom25.Checked = False
  mnuZoom50.Checked = False
  mnuZoom100.Checked = False
  mnu200.Checked = False
  mnu300.Checked = True
  mnu400.Checked = False
End Sub

Private Sub mnu400_Click()
  Z = 4
  pic.FontSize = 36
  DrawSchematic
  mnuZoom25.Checked = False
  mnuZoom50.Checked = False
  mnuZoom100.Checked = False
  mnu200.Checked = False
  mnu300.Checked = False
  mnu400.Checked = True
End Sub

Private Sub pic_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = 2 Then PopupMenu mnuPopup
End Sub

Private Sub vsb_Change()
 DrawSchematic
End Sub

Private Sub vsb_Scroll()
  vsb_Change
End Sub
