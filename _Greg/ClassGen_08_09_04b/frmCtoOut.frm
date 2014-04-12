VERSION 5.00
Begin VB.Form frmCtoOut 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Convert Code to Output"
   ClientHeight    =   9225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9180
   Icon            =   "frmCtoOut.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9225
   ScaleWidth      =   9180
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdShow 
      Enabled         =   0   'False
      Height          =   375
      Left            =   4560
      TabIndex        =   6
      Top             =   4320
      Width           =   4395
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   7320
      TabIndex        =   5
      Top             =   8760
      Width           =   1695
   End
   Begin VB.TextBox txtOutput 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   4320
      Width           =   1935
   End
   Begin VB.TextBox txtInput 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
'********************************************************************************************
'frmCtoOut.frm - Utitility for converting VB code to a VB hardcode output format
'********************************************************************************************
Option Explicit

Private m_sProcedureCode As String  'stores content of created procedure
Private m_sProcedureName As String 'stores the name of created procedure

'close window
Private Sub cmdClose_Click()
  Unload Me
End Sub

'This allows you to copy code from VB into a text box.  This is then converted to
'a series of sOut = sOut & "....." strings in order to autogenerate this same function
'via hard code.  Sometimes you may want your program to automatically generate
'VB code in a text box...this will do that. Most cool!
Private Sub cmdConvert_Click()
  
  'creates variables
  Dim sOut As String
  Dim sIn As String
  Dim i As Integer
  Dim j As Integer
  Dim sCode() As String 'stores delimited strings
  Dim sQuote As String 'CHR(34)
  Dim sFrag As String  'holds temporary string so quotes are replaced by & sQuote &
  Dim sProcedure As String 'stores name of procedure
  
  'setup if text has content
  If Len(txtInput.Text) < 1 Then Exit Sub
  sQuote = Chr(34)
  sIn = txtInput.Text
  txtOutput.Text = ""
  sCode = Split(sIn, vbCrLf)
  
  sProcedure = InputBox("Enter name of Function: ", "Function Name")
  If Len(sProcedure) < 1 Then sProcedure = "GenerateCode"
  
  'constructs string from entire piece of code
  sOut = ""
  sOut = sOut & "'*********************************************" & vbCrLf
  sOut = sOut & "'* Code Generated by ClassGen v0.3    " & vbCrLf
  sOut = sOut & "'*********************************************" & vbCrLf
  sOut = sOut & "Public Function " & sProcedure & " ( ) As String" & vbCrLf
  sOut = sOut & "    Dim sOut As String" & vbCrLf & vbCrLf
  sOut = sOut & "    sOut = " & sQuote & sQuote & vbCrLf
  
  'go through each line...must convert quotation marks
  For i = 0 To UBound(sCode) - 1
    sFrag = ""
    For j = 1 To Len(sCode(i))
      If Mid(sCode(i), j, 1) = sQuote Then
        sFrag = sFrag & sQuote & " & " & sQuote
      Else
        sFrag = sFrag & Mid(sCode(i), j, 1)
      End If
    Next j
    
    sOut = sOut & "    sOut = sOut  & " & sQuote & sFrag & sQuote & " & vbCrLf" & vbCrLf
  Next i
  
  'end code creation
  sOut = sOut & "    " & sProcedure & " = sOut" & vbCrLf
  sOut = sOut & "End Function" & vbCrLf

  txtOutput.Text = sOut
  m_sProcedureCode = sOut
  m_sProcedureName = sProcedure
  cmdShow.Caption = "Show " & m_sProcedureName & "( ) "
  cmdShow.Enabled = True
  
End Sub

Private Sub cmdShow_Click()
  MsgBox m_sProcedureCode
End Sub
