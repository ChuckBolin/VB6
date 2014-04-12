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
      WindowList      =   -1  'True
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
  Z = 3
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
  DrawMainFuse pic, 1900 - hsb.Value, 1000 - vsb.Value
End Sub


'*********************************************************************
'*********************************
'********************************* D R A W I N G   R O U T I N E S
'*********************************
'*********************************************************************
Private Sub DrawDisconnect(p As PictureBox, X As Single, Y As Single)
  Dim nFont As Integer
  nFont = p.FontSize
  If p.FontSize > 8 Then p.FontSize = nFont - 8
  DrawText p, X - 700, Y, "480VAC"
  DrawText p, X - 700, Y + 200, "3 Phase"
  DrawText p, X - 700, Y + 400, "60 Hz"
  DrawText p, X + 375, Y - 300, "Q0"
  DrawText p, X, Y - 150, "L1"
  DrawText p, X, Y + 150, "L2"
  DrawText p, X, Y + 450, "L3"
  
  p.FontSize = p.FontSize / 2
  DrawText p, X + 275, Y - 120, "1"
  DrawText p, X + 600, Y - 120, "2"
  DrawText p, X + 275, Y + 180, "3"
  DrawText p, X + 600, Y + 180, "4"
  DrawText p, X + 275, Y + 480, "5"
  DrawText p, X + 600, Y + 480, "6"

  DrawNOSwitch p, X, Y
  DrawNOSwitch p, X, Y + 300
  DrawNOSwitch p, X, Y + 600
  p.DrawStyle = 1
  p.Line ((X + 450) * Z, Y * Z)-((X + 450) * Z, (Y + 550) * Z)
  p.DrawStyle = 0
  p.FontSize = nFont
End Sub

Private Sub DrawMainFuse(p As PictureBox, X As Single, Y As Single)
  Dim nFont As Integer
  nFont = p.FontSize
  
  If p.FontSize > 8 Then p.FontSize = nFont - 8
  DrawText p, X + 350, Y - 200, "F1A"
  DrawText p, X + 350, Y + 100, "F1B"
  DrawText p, X + 350, Y + 400, "F1C"
  
  DrawText p, X - 100, Y - 150, "0L1"
  DrawText p, X - 100, Y + 150, "0L2"
  DrawText p, X - 100, Y + 450, "0L3"
  
  p.FontSize = nFont / 2
  DrawText p, X + 250, Y - 100, "1"
  DrawText p, X + 625, Y - 100, "2"
  DrawText p, X + 250, Y + 200, "3"
  DrawText p, X + 625, Y + 200, "4"
  DrawText p, X + 250, Y + 500, "5"
  DrawText p, X + 625, Y + 500, "6"
    
  DrawFuse p, X, Y
  DrawFuse p, X, Y + 300
  DrawFuse p, X, Y + 600
  p.FontSize = nFont
  
End Sub

'*********************************************************************
'*********************************  D R A W I N G  P R I M I T I V E S
'*********************************************************************
Private Sub DrawText(p As PictureBox, X As Single, Y As Single, s As String)
  p.CurrentX = X * Z
  p.CurrentY = Y * Z
  p.Print s
End Sub

Private Sub DrawNOSwitch(p As PictureBox, X As Single, Y As Single)
  p.Line (X * Z, Y * Z)-((X + 300) * Z, Y * Z)
  p.Circle ((X + 300) * Z, Y * Z), 25 * Z
  p.Line ((X + 300) * Z, Y * Z)-((X + 550) * Z, (Y - 100) * Z)
  p.Circle ((X + 600) * Z, Y * Z), 25 * Z
  p.Line ((X + 600) * Z, Y * Z)-((X + 900) * Z, Y * Z)
End Sub

Private Sub DrawFuse(p As PictureBox, X As Single, Y As Single)
  p.Line (X * Z, Y * Z)-((X + 300) * Z, Y * Z)
  p.Line ((X + 300) * Z, (Y - 50) * Z)-((X + 600) * Z, (Y + 50) * Z), , B
  p.Line ((X + 600) * Z, Y * Z)-((X + 900) * Z, Y * Z)
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

Private Sub pic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then PopupMenu mnuPopup
End Sub

Private Sub vsb_Change()
 DrawSchematic
End Sub

Private Sub vsb_Scroll()
  vsb_Change
End Sub
