VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Play Reader"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   7230
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   5520
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   4215
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private tp As New CTextParser
Private pb() As New CPlayBook

Private Sub Command1_Click()
  Dim i As Integer
  Dim sIn As String
  Text1.Text = ""
  For i = 1 To UBound(pb)
    sIn = ""
    sIn = pb(i).PlayNumber & " - " & pb(i).Position & " - " & pb(i).X & " - " & pb(i).Y & " - " & pb(i).Action & _
        " - " & pb(i).WX1 & " - " & pb(i).WY1 & " - " & pb(i).WX2 & " - " & pb(i).WY2 & vbCrLf
    Text1.Text = Text1.Text & sIn
  Next i
End Sub

Private Sub Form_Load()
  Dim sIn As String 'line read from file
  Dim i As Integer
  Dim nFields As Integer 'total number of fields for line
  Dim nIndex As Integer
  Dim nPlayNum As Integer 'number of play in playbook
  
  tp.DelimitChar = ","
  
  Open App.Path & "\home.csv" For Input As #1
    Do
      Line Input #1, sIn
      nFields = tp.ProcessString(LCase(sIn))
      If nFields > 0 Then
        If tp.GetField(1) = "playnum" Then
          nPlayNum = CInt(tp.GetField(2))
        Else
          nIndex = nIndex + 1
          ReDim Preserve pb(nIndex)
          pb(nIndex).PlayNumber = nPlayNum
          pb(nIndex).Position = tp.GetField(1)
          pb(nIndex).X = tp.GetField(2)
          pb(nIndex).Y = tp.GetField(3)
          pb(nIndex).Action = tp.GetField(4)
          pb(nIndex).WX1 = CSng(Val(tp.GetField(5)))
          pb(nIndex).WY1 = CSng(Val(tp.GetField(6)))
          pb(nIndex).WX2 = CSng(Val(tp.GetField(7)))
          pb(nIndex).WY2 = CSng(Val(tp.GetField(8)))
        End If
      End If
        
    Loop Until EOF(1)
  Close #1
End Sub
