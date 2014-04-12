VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Team Generator"
   ClientHeight    =   4845
   ClientLeft      =   3165
   ClientTop       =   2160
   ClientWidth     =   8040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   8040
   Begin VB.ListBox lstName 
      Height          =   3570
      ItemData        =   "Form1.frx":0000
      Left            =   2820
      List            =   "Form1.frx":0002
      TabIndex        =   4
      Top             =   120
      Width           =   555
   End
   Begin VB.TextBox Text1 
      Height          =   3555
      Left            =   3540
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   120
      Width           =   4275
   End
   Begin VB.ListBox lstTeam 
      Height          =   3570
      ItemData        =   "Form1.frx":0004
      Left            =   2220
      List            =   "Form1.frx":0006
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   ">"
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   1440
      Width           =   675
   End
   Begin VB.ListBox lstLetters 
      Height          =   3570
      ItemData        =   "Form1.frx":0008
      Left            =   180
      List            =   "Form1.frx":000A
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private sLet() As String
Private sTeam() As String

Private Sub cmdSelect_Click()
  Dim i As Integer, j As Integer
  
  Dim nLet As Integer, sID As String
  
  lstTeam.Clear
  lstName.Clear
  Text1.Text = ""
  
  For i = 1 To 5
    
    
    
    'first name
    nLet = 1 + (Rnd * 26) Mod 26 'returns 1 to 26
    lstTeam.AddItem sLet(nLet)
    Text1.Text = Text1.Text & sLet(nLet) & ", " & vbTab
    
    'last name
    nLet = 1 + (Rnd * 26) Mod 26 'returns 1 to 26
    lstName.AddItem sLet(nLet)
    Text1.Text = Text1.Text & sLet(nLet) & ", " & vbTab
    
    'unique ID number
    sID = ""
    For j = 1 To 5
      sID = sID & sLet(1 + (Rnd * 26) Mod 26)   '1 to 26
      sID = sID & CStr(CInt((Rnd * 10) Mod 10)) '0 to 9
    Next j
    Text1.Text = Text1.Text & sID & vbCrLf
    
  Next i
  
  'Text1.Text = Text1.Text & nLet & vbCrLf
  
End Sub

Private Sub Form_Load()
  Dim i As Integer
  
  Randomize Timer
  
  ReDim sLet(26)
  For i = 1 To 26
    sLet(i) = Chr(64 + i)
    lstLetters.AddItem sLet(i)
  Next i
  
  
  
  
End Sub
