VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   7680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDir 
      Height          =   315
      Left            =   6120
      TabIndex        =   3
      Top             =   1500
      Width           =   375
   End
   Begin VB.VScrollBar vsbDir 
      Height          =   1695
      Left            =   6180
      Max             =   628
      TabIndex        =   2
      Top             =   1920
      Width           =   255
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "&Reset"
      Height          =   435
      Left            =   6120
      TabIndex        =   1
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Timer tmrUpdate 
      Interval        =   100
      Left            =   6720
      Top             =   4200
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   6000
      Left            =   60
      MousePointer    =   2  'Cross
      ScaleHeight     =   -100
      ScaleMode       =   0  'User
      ScaleTop        =   100
      ScaleWidth      =   100
      TabIndex        =   0
      Top             =   60
      Width           =   6000
      Begin VB.Line linRF 
         BorderColor     =   &H0000C000&
         X1              =   25.253
         X2              =   25.253
         Y1              =   36.364
         Y2              =   32.323
      End
      Begin VB.Line linLF 
         BorderColor     =   &H000000C0&
         X1              =   16.162
         X2              =   21.212
         Y1              =   37.374
         Y2              =   42.424
      End
      Begin VB.Line linRA 
         BorderColor     =   &H0000C000&
         BorderWidth     =   2
         X1              =   53.535
         X2              =   67.677
         Y1              =   25.253
         Y2              =   19.192
      End
      Begin VB.Line linLA 
         BorderColor     =   &H000000C0&
         BorderWidth     =   2
         X1              =   26.263
         X2              =   20.202
         Y1              =   76.768
         Y2              =   68.687
      End
      Begin VB.Line linLeft 
         X1              =   24.242
         X2              =   38.384
         Y1              =   14.141
         Y2              =   11.111
      End
      Begin VB.Line linRight 
         X1              =   43.434
         X2              =   65.657
         Y1              =   40.404
         Y2              =   39.394
      End
      Begin VB.Line linFront 
         X1              =   50.505
         X2              =   61.616
         Y1              =   59.596
         Y2              =   52.525
      End
      Begin VB.Shape shpBug 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         Height          =   356
         Left            =   2820
         Shape           =   3  'Circle
         Top             =   2820
         Width           =   356
      End
      Begin VB.Shape shpTarget 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  'Opaque
         Height          =   240
         Left            =   2160
         Shape           =   3  'Circle
         Top             =   2640
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdReset_Click()
  LoadBug
  vsbDir.Value = 0
End Sub

Private Sub Form_Load()
  LoadBug
  vsbDir_Change
End Sub

Private Sub pic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  target.X = X
  target.Y = Y
End Sub

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  frmMain.Caption = FormatNumber(X, 2) & " " & FormatNumber(Y, 2)
End Sub

Private Sub tmrUpdate_Timer()
  UpdateBug
  UpdateGraphics
End Sub

Private Sub UpdateGraphics()
  pic.Cls
  shpTarget.Left = target.X - (shpTarget.Width / 2)
  shpTarget.Top = target.Y + (shpTarget.Height / 2)
  
  shpBug.Width = 0.5 * bug.size
  shpBug.Height = 0.5 * bug.size
  
  shpBug.Left = bug.X - (shpBug.Width / 2)
  shpBug.Top = bug.Y + (shpBug.Height / 2)
  linFront.X1 = bug.a.X
  linFront.Y1 = bug.a.Y
  linFront.X2 = bug.b.X
  linFront.Y2 = bug.b.Y
  linRight.X1 = bug.b.X
  linRight.Y1 = bug.b.Y
  linRight.X2 = bug.c.X
  linRight.Y2 = bug.c.Y
  linLeft.X1 = bug.a.X
  linLeft.Y1 = bug.a.Y
  linLeft.X2 = bug.c.X
  linLeft.Y2 = bug.c.Y
  
  linLA.X1 = bug.elbow1.X
  linLA.Y1 = bug.elbow1.Y
  linLA.X2 = bug.a.X
  linLA.Y2 = bug.a.Y
  linRA.X1 = bug.elbow2.X
  linRA.Y1 = bug.elbow2.Y
  linRA.X2 = bug.b.X
  linRA.Y2 = bug.b.Y
  
  linLF.X1 = bug.wrist1.X
  linLF.Y1 = bug.wrist1.Y
  linLF.X2 = bug.elbow1.X
  linLF.Y2 = bug.elbow1.Y
  linRF.X1 = bug.wrist2.X
  linRF.Y1 = bug.wrist2.Y
  linRF.X2 = bug.elbow2.X
  linRF.Y2 = bug.elbow2.Y
  
  
End Sub

Private Sub vsbDir_Change()
  txtDir = vsbDir.Value * 0.01
  bug.direction = vsbDir.Value * 0.01
End Sub

Private Sub vsbDir_Scroll()
  vsbDir_Change
End Sub
