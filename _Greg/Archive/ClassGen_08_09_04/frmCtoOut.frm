VERSION 5.00
Begin VB.Form frmCtoOut 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Convert Code to Output"
   ClientHeight    =   9225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9225
   ScaleWidth      =   9180
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   7320
      TabIndex        =   5
      Top             =   8640
      Width           =   1695
   End
   Begin VB.TextBox txtOutput 
      Height          =   3615
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   4920
      Width           =   8895
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "Convert Code"
      Height          =   495
      Left            =   3240
      TabIndex        =   2
      Top             =   4200
      Width           =   1935
   End
   Begin VB.TextBox txtInput 
      Height          =   3615
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   480
      Width           =   8895
   End
   Begin VB.Label Label2 
      Caption         =   "Output Code:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Input VB Code:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmCtoOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
  Unload Me
  
End Sub

Private Sub cmdConvert_Click()
  If Len(txtInput.Text) < 1 Then Exit Sub
  
  Dim sIn As String
  Dim sOut As String
  Dim nCrLf As Integer
  Dim nBegin As Integer
  
  Dim nQuote As Integer
  Dim sQuote As String
  Dim i As Integer
  
  sQuote = Chr$(34)
  
  sIn = txtInput.Text
  txtOutput.Text = ""
  
  sOut = "Private Function GenerateCode (sOut As String) As String" & vbCrLf
  'i = 1
  nBegin = 1
  
  For i = 1 To Len(sIn)
    If Mid(sIn, i, 2) = vbCrLf Then
      txtOutput.Text = txtOutput.Text & Mid(sIn, nBegin, i - nBegin - 1) & vbCrLf
      nBegin = i + 1
      
    End If
  Next i
  
  sOut = sOut & "End Function" & vbCrLf & vbCrLf
  txtOutput.Text = sOut
End Sub
