VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Player Generator"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   5745
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   ">"
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   1680
      Width           =   735
   End
   Begin VB.ListBox lstPlayer 
      Height          =   3180
      ItemData        =   "Form1.frx":0000
      Left            =   1740
      List            =   "Form1.frx":0002
      TabIndex        =   0
      Top             =   60
      Width           =   3795
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private sPos(15) As String
Private sSec(15) As String

Private Sub Command1_Click()
  Dim i As Integer
  Dim nPos As Integer, nSec As Integer
  
  lstPlayer.Clear
  i = 0
  Do
    nPos = 1 + (Rnd * 15) Mod 15
    If Len(sPos(nPos)) > 0 Then  'primary position found
      Do
        nSec = 1 + (Rnd * 15) Mod 15
        If Len(sSec(nSec)) > 0 Then  'secondary position found
          lstPlayer.AddItem i + 1 & ": " & vbTab & sPos(nPos) & vbTab & sSec(nSec)
          sPos(nPos) = "" 'delete letter
          sSec(nSec) = ""
          i = i + 1
          Exit Do
        End If
      Loop
    End If
  Loop Until i = 15
  
  LoadPositions
End Sub

Private Sub LoadPositions()
  Dim i As Integer
  
  sPos(1) = "K"
  sPos(2) = "K"
  sPos(3) = "Q"
  sPos(4) = "Q"
  sPos(5) = "Q"
  sPos(6) = "T"
  sPos(7) = "T"
  sPos(8) = "T"
  sPos(9) = "T"
  sPos(10) = "C"
  sPos(11) = "C"
  sPos(12) = "G"
  sPos(13) = "G"
  sPos(14) = "TE"
  sPos(15) = "TE"

  For i = 1 To 15
    sSec(i) = sPos(i)
  Next i
End Sub

Private Sub Form_Load()
  LoadPositions

End Sub
