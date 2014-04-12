VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Flashcard v0.01 - Written by C. Bolin,  October 2004"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   10410
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkRandom 
      Caption         =   "Randomize Questions"
      Height          =   285
      Left            =   60
      TabIndex        =   13
      Top             =   5460
      Width           =   2145
   End
   Begin VB.ComboBox cboCat 
      Height          =   315
      Left            =   30
      TabIndex        =   12
      Text            =   "(Select Category)"
      Top             =   5040
      Width           =   2175
   End
   Begin VB.CommandButton cmdPrev 
      Caption         =   "<"
      Height          =   405
      Left            =   2370
      TabIndex        =   8
      Top             =   5040
      Width           =   495
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">"
      Height          =   405
      Left            =   3690
      TabIndex        =   7
      Top             =   5040
      Width           =   495
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "&Show"
      Height          =   405
      Left            =   2850
      TabIndex        =   6
      Top             =   5040
      Width           =   825
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   345
      Left            =   9150
      TabIndex        =   5
      Top             =   5850
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   5745
      Left            =   4230
      ScaleHeight     =   5685
      ScaleWidth      =   6075
      TabIndex        =   4
      Top             =   30
      Width           =   6135
   End
   Begin VB.TextBox Text1 
      Height          =   1605
      Left            =   30
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   3270
      Width           =   4155
   End
   Begin VB.TextBox txtQ 
      Height          =   2595
      Left            =   30
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   330
      Width           =   4155
   End
   Begin VB.Label lblTotal 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3660
      TabIndex        =   11
      Top             =   5490
      Width           =   525
   End
   Begin VB.Label lblCurrent 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2370
      TabIndex        =   10
      Top             =   5490
      Width           =   525
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "of"
      Height          =   255
      Left            =   3000
      TabIndex        =   9
      Top             =   5490
      Width           =   525
   End
   Begin VB.Label Label2 
      Caption         =   "Answer:"
      Height          =   285
      Left            =   30
      TabIndex        =   2
      Top             =   2970
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Question:"
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   765
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_sFilename As String

Private Sub Form_Load()
  m_sFilename = App.Path & "\eem117.fl"
  ReadCategories
End Sub

Private Sub ReadCategories()
  Dim sIn As String
  Dim sElement As String 'category, question, answer, picture
  Dim i As Integer
  Dim bFound As Boolean
  
  ReDim g_sCat(0) As String
  g_sCat(0) = "All"
   
  Open m_sFilename For Input As #1
    Do
      Line Input #1, sIn
      sIn = LTrim(RTrim(sIn))
      If Len(sIn) > 0 Then
        sElement = LCase(Left(sIn, 2))
        Select Case sElement
          Case "c:"
            bFound = False
            For i = 1 To UBound(g_sCat)
              If LCase(g_sCat(i)) = LCase(Mid(sIn, 3)) Then bFound = True
            Next i
            If bFound = False Then  'category not found...add to list
              ReDim Preserve g_sCat(UBound(g_sCat) + 1)
              g_sCat(UBound(g_sCat)) = Mid(sIn, 3)
            End If
          Case "q:"
        
          Case "a:"
          
          Case "p:"
          
        End Select
      
      
      End If
    Loop Until EOF(1)
  
  Close #1
  
  g_nTotalCategories = UBound(g_sCat)
  If g_nTotalCategories < 1 Then
    MsgBox "File " & m_sFilename & " not found!  Aborting program."
    End
  End If
  
  'add to list
  For i = 0 To UBound(g_sCat)
    cboCat.AddItem g_sCat(i)
  Next i
  
End Sub
