VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6750
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7650
   LinkTopic       =   "Form1"
   ScaleHeight     =   6750
   ScaleWidth      =   7650
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   720
      Top             =   4740
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Hit Right"
      Height          =   315
      Left            =   3540
      TabIndex        =   4
      Top             =   2940
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Hit Left"
      Height          =   315
      Left            =   1320
      TabIndex        =   3
      Top             =   2880
      Width           =   1095
   End
   Begin VB.ListBox lstRight 
      Height          =   1620
      ItemData        =   "Form1.frx":0D54
      Left            =   3300
      List            =   "Form1.frx":0D56
      TabIndex        =   2
      Top             =   1020
      Width           =   1515
   End
   Begin VB.ListBox lstLeft 
      Height          =   1620
      ItemData        =   "Form1.frx":0D58
      Left            =   1080
      List            =   "Form1.frx":0D5A
      TabIndex        =   1
      Top             =   1020
      Width           =   1515
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Deal"
      Height          =   555
      Left            =   4200
      TabIndex        =   0
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   675
      Left            =   240
      Top             =   60
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private nCard(52) As String

Private Sub Command1_Click()
  Form1.Caption = "Card Deal"
  lstLeft.Clear
  lstRight.Clear

  Shuffle
  Dim nCt As Integer
  Dim i As Integer
  
  Do
    nCt = 1 + (Rnd * 52) Mod 52
    If nCard(nCt) = 1 Then
      If nCt > 0 And nCt < 14 Then
        lstLeft.AddItem Convert(nCt) & " Heart"
      ElseIf nCt > 13 And nCt < 27 Then
        lstLeft.AddItem Convert(nCt - 13) & " Club"
      ElseIf nCt > 26 And nCt < 40 Then
        lstLeft.AddItem Convert(nCt - 26) & " Diamond"
      ElseIf nCt > 39 And nCt < 53 Then
        lstLeft.AddItem Convert(nCt - 39) & " Spade"
      End If
      nCard(nCt) = 0
      i = i + 1
    End If
  Loop Until i = 5
  Print
  
  i = 0
  Do
    nCt = 1 + (Rnd * 52) Mod 52
    If nCard(nCt) = 1 Then
      If nCt > 0 And nCt < 14 Then
        lstRight.AddItem Convert(nCt) & " Heart"
      ElseIf nCt > 13 And nCt < 27 Then
        lstRight.AddItem Convert(nCt - 13) & " Club"
      ElseIf nCt > 26 And nCt < 40 Then
        lstRight.AddItem Convert(nCt - 26) & " Diamond"
      ElseIf nCt > 39 And nCt < 53 Then
        lstRight.AddItem Convert(nCt - 39) & " Spade"
      End If
      nCard(nCt) = 0
      i = i + 1
    End If
  Loop Until i = 5
  
  
  
End Sub

Private Function Convert(num As Integer) As String
  If num < 11 Then
    Convert = CStr(num)
  ElseIf num = 11 Then
    Convert = "J"
  ElseIf num = 12 Then
    Convert = "Q"
  ElseIf num = 13 Then
    Convert = "K"
  End If
End Function


Private Sub Command2_Click()
  Dim nCt As Integer
  Dim i As Integer
  Dim nMax As Integer
  nMax = 5 - lstLeft.ListCount
  If nMax = 0 Then Exit Sub
  
  Do
    nCt = 1 + (Rnd * 52) Mod 52
    If nCard(nCt) = 1 Then
      If nCt > 0 And nCt < 14 Then
        lstLeft.AddItem Convert(nCt) & " Heart"
      ElseIf nCt > 13 And nCt < 27 Then
        lstLeft.AddItem Convert(nCt - 13) & " Club"
      ElseIf nCt > 26 And nCt < 40 Then
        lstLeft.AddItem Convert(nCt - 26) & " Diamond"
      ElseIf nCt > 39 And nCt < 53 Then
        lstLeft.AddItem Convert(nCt - 39) & " Spade"
      End If
      nCard(nCt) = 0
      i = i + 1
    End If
  Loop Until i = nMax
End Sub

Private Sub Command3_Click()
 Dim nCt As Integer
  Dim i As Integer
  Dim nMax As Integer
  nMax = 5 - lstRight.ListCount
  If nMax = 0 Then Exit Sub
  
  Do
    nCt = 1 + (Rnd * 52) Mod 52
    If nCard(nCt) = 1 Then
      If nCt > 0 And nCt < 14 Then
        lstRight.AddItem Convert(nCt) & " Heart"
      ElseIf nCt > 13 And nCt < 27 Then
        lstRight.AddItem Convert(nCt - 13) & " Club"
      ElseIf nCt > 26 And nCt < 40 Then
        lstRight.AddItem Convert(nCt - 26) & " Diamond"
      ElseIf nCt > 39 And nCt < 53 Then
        lstRight.AddItem Convert(nCt - 39) & " Spade"
      End If
      nCard(nCt) = 0
      i = i + 1
    End If
  Loop Until i = nMax
End Sub

Private Sub Form_Load()
  Shuffle
End Sub

Private Sub Shuffle()
  Dim i As Integer
  
  For i = 1 To 52
    nCard(i) = 1
  Next i
  
End Sub

Private Sub lstLeft_Click()
  Dim i As Integer
  For i = 0 To lstLeft.ListCount - 1
    If lstLeft.Selected(i) = True Then
      lstLeft.RemoveItem (i)
      Exit For
    End If
  Next i
End Sub

Private Sub lstRight_Click()
 Dim i As Integer
  For i = 0 To lstRight.ListCount - 1
    If lstRight.Selected(i) = True Then
      lstRight.RemoveItem (i)
      Exit For
    End If
  Next i
End Sub
