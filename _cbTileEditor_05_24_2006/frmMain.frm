VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "cbTileEditor v0.1 - Written by Chuck Bolin, May 2006"
   ClientHeight    =   9240
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10770
   LinkTopic       =   "Form1"
   ScaleHeight     =   9240
   ScaleWidth      =   10770
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picMap 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H00FFFFFF&
      Height          =   6270
      Left            =   3180
      ScaleHeight     =   6240
      ScaleWidth      =   6780
      TabIndex        =   4
      Top             =   240
      Width           =   6810
   End
   Begin VB.Timer tmrLoad 
      Interval        =   250
      Left            =   6120
      Top             =   7500
   End
   Begin VB.PictureBox picMaster 
      AutoSize        =   -1  'True
      Height          =   2460
      Left            =   120
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   2400
      ScaleWidth      =   4800
      TabIndex        =   1
      Top             =   8340
      Width           =   4860
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tiles"
      Height          =   5115
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   2835
      Begin VB.CommandButton cmdClear 
         Caption         =   "&Clear"
         Height          =   315
         Left            =   600
         TabIndex        =   5
         Top             =   4560
         Width           =   1335
      End
      Begin VB.PictureBox picTile 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   510
         Index           =   1
         Left            =   780
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   3
         Top             =   300
         Width           =   510
      End
      Begin VB.PictureBox picTile 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   510
         Index           =   0
         Left            =   180
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   2
         Top             =   300
         Width           =   510
      End
      Begin VB.Shape shpMarker 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   4
         Height          =   645
         Left            =   120
         Top             =   240
         Width           =   645
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private gTileX As Integer
Private gTileY As Integer


Private Sub cmdClear_Click()
  picMap.Cls
End Sub

Private Sub picMap_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim tileX As Integer
  Dim tileY As Integer
  Dim clearX As Integer
  Dim clearY As Integer
    
  tileX = (CLng(X) \ 480) * 32
  tileY = (CLng(Y) \ 480) * 32
  
  clearX = (CInt(X) \ 480) * 480
  clearY = (CInt(Y) \ 480) * 480
  
  If Button = 1 Then
    BitBlt picMap.hDC, tileX, tileY, 32, 32, picMaster.hDC, gTileX, gTileY, vbSrcCopy
  ElseIf Button = 2 Then
    picMap.Line (clearX, clearY)-(clearX + 480, clearY + 480), , B
  End If
  
End Sub

Private Sub picMap_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim tileX As Integer
  Dim tileY As Integer
  
  tileX = (CLng(X) \ picTile(0).Width) * 32
  tileY = (CLng(Y) \ picTile(0).Height) * 32
  frmMain.Caption = tileX & "     " & tileY
End Sub

Private Sub picTile_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  shpMarker.Left = picTile(Index).Left - 60
  shpMarker.Top = picTile(Index).Top - 60
  
  If Index = 0 Then
    gTileX = 0
    gTileY = 0
  Else
    gTileX = 32
    gTileY = 0
  End If
End Sub

Private Sub tmrLoad_Timer()
  'BitBlt picMap.hDC, 0, 0, 32, 32, picMaster.hDC, 0, 0, vbSrcCopy
  BitBlt picTile(0).hDC, 0, 0, 32, 32, picMaster.hDC, 0, 0, vbSrcCopy
  BitBlt picTile(1).hDC, 0, 0, 32, 32, picMaster.hDC, 32, 0, vbSrcCopy
  'BitBlt picMap.hDC, 0, 0, 32, 32, picMaster.hDC, 0, 0, vbSrcCopy
  tmrLoad.Enabled = False
End Sub
