VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bitmap Generator (8x8 Fonts) - December 2005"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   9765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Generate"
      Height          =   315
      Left            =   6840
      TabIndex        =   3
      Top             =   6720
      Width           =   975
   End
   Begin VB.TextBox txtCode 
      Height          =   6615
      Left            =   5040
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   0
      Width           =   4395
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Draw Grid"
      Height          =   315
      Left            =   1920
      TabIndex        =   1
      Top             =   6720
      Width           =   1035
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Left            =   60
      ScaleHeight     =   6555
      ScaleWidth      =   4815
      TabIndex        =   0
      Top             =   0
      Width           =   4875
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private dat(47, 47) As Integer

Private Sub Command1_Click()
  Dim i, j As Integer
  
  pic.Cls
  pic.FillColor = vbWhite
  
  'small grid squares
  pic.DrawWidth = 1
  For i = 0 To 47
    For j = 0 To 47
      pic.Line (i * 100, j * 100)-((i + 1) * 100, (j + 1) * 100), , B
      dat(i, j) = 0
    Next j
  Next i
  
  'large grids
  pic.DrawWidth = 2
  pic.FillStyle = 1
  For i = 0 To 5
    For j = 0 To 5
      pic.Line (i * 800, j * 800)-((i + 1) * 800, (j + 1) * 800), , B
    Next j
  Next i
  
  pic.CurrentX = 10
  pic.CurrentY = 5000
  pic.Print "ABCDEFGHIJKLM"
  pic.CurrentX = 10
  pic.Print "NOPQRSTUVWXYZ"
  
  
End Sub

Private Sub Command2_Click()
  Dim i, j As Integer
  Dim m, n As Integer
  Dim count As Integer
  
  txtCode = ""
  For i = 0 To 5
    For j = 0 To 5
      count = count + 1
      If count < 27 Then
        txtCode = txtCode & "static GLubyte letter" & Chr(64 + count) & "[] =" & vbCrLf
        txtCode = txtCode & "{" & vbCrLf
        
        
        
        txtCode = txtCode & "};" & vbCrLf
      End If
    Next j
  Next i

End Sub

Private Sub Form_Load()
  Command1_Click
End Sub

Private Sub pic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim px, py As Single
  px = ((X \ 100) * 100)
  py = ((Y \ 100) * 100)
  
  pic.DrawWidth = 1
  pic.FillStyle = 0
  If Button = 1 Then
    pic.FillColor = vbRed
    pic.Line (px + 5, py + 5)-(px + 90, py + 90), , B
    dat(px / 100, py / 100) = 1
  ElseIf Button = 2 Then
    pic.FillColor = vbWhite
    pic.Line (px + 5, py + 5)-(px + 90, py + 90), , B
    dat(px / 100, py / 100) = 0
  End If
  
  Form1.Caption = px / 100 & "  " & py / 100
End Sub

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim px, py As Single
  px = ((X \ 100) * 100)
  py = ((Y \ 100) * 100)
  
  pic.DrawWidth = 1
  pic.FillStyle = 0
  If Button = 1 Then
    pic.FillColor = vbRed
    pic.Line (px + 5, py + 5)-(px + 90, py + 90), , B
    dat(px / 100, py / 100) = 1
  ElseIf Button = 2 Then
    pic.FillColor = vbWhite
    pic.Line (px + 5, py + 5)-(px + 90, py + 90), , B
    dat(px / 100, py / 100) = 0
  End If

End Sub
