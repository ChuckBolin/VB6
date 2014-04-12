VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "cbTileEditor v0.1 - Written by Chuck Bolin, May 2006"
   ClientHeight    =   7725
   ClientLeft      =   60
   ClientTop       =   540
   ClientWidth     =   12330
   LinkTopic       =   "Form1"
   ScaleHeight     =   7725
   ScaleWidth      =   12330
   Begin VB.PictureBox picMaster 
      AutoSize        =   -1  'True
      Height          =   2460
      Left            =   3060
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   2400
      ScaleWidth      =   4800
      TabIndex        =   1
      Top             =   240
      Width           =   4860
   End
   Begin VB.VScrollBar vsb 
      Height          =   7095
      Left            =   11940
      Max             =   15
      TabIndex        =   9
      Top             =   120
      Width           =   315
   End
   Begin VB.HScrollBar hsb 
      Height          =   315
      Left            =   2940
      Max             =   300
      TabIndex        =   8
      Top             =   7320
      Width           =   9015
   End
   Begin VB.PictureBox picMap 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H00FFFFFF&
      Height          =   7200
      Left            =   2940
      ScaleHeight     =   7170
      ScaleWidth      =   8970
      TabIndex        =   4
      Top             =   120
      Width           =   9000
   End
   Begin VB.Timer tmrLoad 
      Interval        =   50
      Left            =   2220
      Top             =   6960
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tiles"
      Height          =   7575
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   2835
      Begin VB.PictureBox picTile 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   510
         Index           =   3
         Left            =   1980
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   7
         Top             =   300
         Width           =   510
      End
      Begin VB.PictureBox picTile 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   510
         Index           =   2
         Left            =   1380
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   6
         Top             =   300
         Width           =   510
      End
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
         AutoRedraw      =   -1  'True
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
         AutoRedraw      =   -1  'True
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
'***********************************************************************
' cbMapEditor - written by Chuck Bolin, May 2006
' Allows user to add 50 different tiles to a side-scrolling map of
' 30 rows by 300 columns.  Data is saved into a comma-delimited array
' for use in a side-scrolling game.
'
'***********************************************************************
Option Explicit

'BitBlt is a Win32 API that allows copying tiles from a bitmap to picture box
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

'global variables
Private gTileX As Integer         'left coordinate of specific tile from bitamp
Private gTileY As Integer         'top coordinate of specific tile from bitmap
Private gMap(30, 318) As Integer  'stores data for map
Private gIndex As Integer         'tracks value to be stored in array based upon tile

'zeros entire array
Private Sub clearMap()
  Dim i As Integer
  Dim j As Integer
  
  For i = 0 To 30
    For j = 0 To 318
      gMap(i, j) = 0
    Next j
  Next i

End Sub

'clears picMap and calls sub to clear out the array
Private Sub cmdClear_Click()
  picMap.Cls
  clearMap
End Sub

'called anytime the horizontal or vertical scrollbars are operated
Private Sub hsb_Change()
  Dim i As Integer
  Dim j As Integer
  Dim value As Integer
  Dim tileX As Integer
  Dim tileY As Integer
    
  picMap.Cls
  
  'reads array and draws it on the map
  For i = vsb To vsb + 15
    For j = hsb To hsb + 18
      value = gMap(i, j) - 1
      If value = 0 Then
        tileX = 0
        tileY = 0
      ElseIf value = 1 Then
        tileX = 32
        tileY = 0
      ElseIf value = 2 Then
        tileX = 64
        tileY = 0
      ElseIf value = 3 Then
        tileX = 96
        tileY = 0
      End If

      If (value > -1) Then
        'BitBlt picMap.hDC, (j - hsb) * 32, (i - vsb) * 32, 32, 32, picMaster.hDC, tileX, tileY, vbSrcCopy
        BitBlt picMap.hDC, (j - hsb) * 32, (i - vsb) * 32, 32, 32, picTile(value).hDC, 0, 0, vbSrcCopy
      End If
    Next j
  Next i
End Sub

'moves map horizontally
Private Sub hsb_Scroll()
 hsb_Change
End Sub

'left button to draw tile, right button to erase tile
Private Sub picMap_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim tileX As Integer
  Dim tileY As Integer
  Dim clearX As Integer
  Dim clearY As Integer
  Dim row As Integer
  Dim col As Integer
  
  tileX = (CLng(X) \ 480) * 32
  tileY = (CLng(Y) \ 480) * 32
  clearX = (CInt(X) \ 480) * 480
  clearY = (CInt(Y) \ 480) * 480
  col = hsb + (CLng(X) \ 480)
  row = vsb + (CLng(Y) \ 480)
  
  If Button = 1 Then
    'BitBlt picMap.hDC, tileX, tileY, 32, 32, picMaster.hDC, gTileX, gTileY, vbSrcCopy
    BitBlt picMap.hDC, tileX, tileY, 32, 32, picTile(gIndex).hDC, 0, 0, vbSrcCopy
    gMap(row, col) = gIndex + 1
  ElseIf Button = 2 Then
    picMap.FillColor = picMap.BackColor
    picMap.ForeColor = picMap.BackColor
    picMap.Line (clearX, clearY)-(clearX + 480, clearY + 480), , B
    gMap(row, col) = 0
  End If
  
End Sub

'allows mouse button to be held down and many tiles to be drawn
Private Sub picMap_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim tileX As Integer
  Dim tileY As Integer
  Dim clearX As Integer
  Dim clearY As Integer
  Dim row As Integer
  Dim col As Integer
  
  tileX = (CLng(X) \ 480) * 32
  tileY = (CLng(Y) \ 480) * 32
  clearX = (CInt(X) \ 480) * 480
  clearY = (CInt(Y) \ 480) * 480
  col = hsb + (CLng(X) \ 480)
  row = vsb + (CLng(Y) \ 480)
  
  If Button = 1 Then
    'BitBlt picMap.hDC, tileX, tileY, 32, 32, picMaster.hDC, gTileX, gTileY, vbSrcCopy
    BitBlt picMap.hDC, tileX, tileY, 32, 32, picTile(gIndex).hDC, 0, 0, vbSrcCopy
    gMap(row, col) = gIndex + 1
  ElseIf Button = 2 Then
    picMap.Line (clearX, clearY)-(clearX + 480, clearY + 480), , B
    gMap(row, col) = 0
  End If

 ' frmMain.Caption = col & "     " & row
End Sub

'user selects tile
Private Sub picTile_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  'this is blue outline showing current selected tile
  shpMarker.Left = picTile(Index).Left - 60
  shpMarker.Top = picTile(Index).Top - 60
  
  gIndex = Index 'value stored in array
  
  'top left corner of tile location in bitmap
  If Index = 0 Then
    gTileX = 0
    gTileY = 0
  ElseIf Index = 1 Then
    gTileX = 32
    gTileY = 0
  ElseIf Index = 2 Then
    gTileX = 64
    gTileY = 0
  ElseIf Index = 3 Then
    gTileX = 96
    gTileY = 0
    
  End If
End Sub

'used only at first to display tiles on left side
Private Sub tmrLoad_Timer()
  BitBlt picTile(0).hDC, 0, 0, 32, 32, picMaster.hDC, 0, 0, vbSrcCopy
  BitBlt picTile(1).hDC, 0, 0, 32, 32, picMaster.hDC, 32, 0, vbSrcCopy
  BitBlt picTile(2).hDC, 0, 0, 32, 32, picMaster.hDC, 64, 0, vbSrcCopy
  BitBlt picTile(3).hDC, 0, 0, 32, 32, picMaster.hDC, 96, 0, vbSrcCopy
  
  'hide this bitmap after tiles on left are drawn
  picMaster.ZOrder (1)
  
  'set erase color to back color
  picMap.FillColor = picMap.BackColor
  picMap.ForeColor = picMap.BackColor
  
  'need to refresh these so they show
  Dim i As Integer
  For i = 0 To 3
    picTile(i).Refresh
  Next i
  
  'disable timer...no longer needed
  tmrLoad.Enabled = False
End Sub

'moves map vertically
Private Sub vsb_Change()
  hsb_Change
End Sub

'moves map vertically
Private Sub vsb_Scroll()
  vsb_Change
End Sub
