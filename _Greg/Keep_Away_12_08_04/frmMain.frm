VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   7590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9330
   LinkTopic       =   "Form1"
   ScaleHeight     =   7590
   ScaleWidth      =   9330
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrUpdate 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   7620
      Top             =   3480
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H00C0C0FF&
      Height          =   7000
      Left            =   0
      ScaleHeight     =   -100
      ScaleMode       =   0  'User
      ScaleTop        =   100
      ScaleWidth      =   100
      TabIndex        =   0
      Top             =   0
      Width           =   7000
      Begin VB.Shape shpBot 
         BackColor       =   &H000000C0&
         BackStyle       =   1  'Opaque
         Height          =   139
         Index           =   5
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Width           =   139
      End
      Begin VB.Shape shpBot 
         BackColor       =   &H000000C0&
         BackStyle       =   1  'Opaque
         Height          =   139
         Index           =   4
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Width           =   139
      End
      Begin VB.Shape shpBot 
         BackColor       =   &H000000C0&
         BackStyle       =   1  'Opaque
         Height          =   139
         Index           =   3
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Width           =   139
      End
      Begin VB.Shape shpBot 
         BackColor       =   &H000000C0&
         BackStyle       =   1  'Opaque
         Height          =   139
         Index           =   2
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Width           =   139
      End
      Begin VB.Shape shpBot 
         BackColor       =   &H000000C0&
         BackStyle       =   1  'Opaque
         Height          =   139
         Index           =   1
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Width           =   139
      End
      Begin VB.Shape shpBot 
         BackColor       =   &H000000C0&
         BackStyle       =   1  'Opaque
         Height          =   139
         Index           =   6
         Left            =   1500
         Shape           =   3  'Circle
         Top             =   2820
         Width           =   139
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Pic_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  frmMain.Caption = x & "  " & y
End Sub

Private Sub Form_Load()
  LoadVariables
  UpdateDisplay
  
End Sub

'updates all graphics
Private Sub UpdateDisplay()
  Dim i As Integer
  
  For i = 1 To 6
    shpBot(i).Left = P.GetX(i) - P.GetDiameter(i) / 2
    shpBot(i).Top = P.GetY(i) + P.GetDiameter(i) / 2
    shpBot(i).Height = P.GetDiameter(i)
    shpBot(i).Width = shpBot(i).Height
    shpBot(i).BackColor = P.GetColor(i)
  Next i
End Sub

'updates all bot data
Private Sub tmrUpdate_Timer()
  Dim i As Integer
  Dim bRet As Boolean
  
  'update bot data
  P.UpdateBots
  
  For i = 1 To P.GetMaxBots
    DoEvents
    If P.AtTarget(i) = True Then
      bRet = P.SetTargetX(i, GetRandomSingle(2, 98))
      bRet = P.SetTargetY(i, GetRandomSingle(2, 98))
      bRet = P.SetVelocity(i, 1)
    End If
  Next i

 
  UpdateDisplay
  
End Sub
