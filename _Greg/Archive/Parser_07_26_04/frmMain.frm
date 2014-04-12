VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Text Parser Demo v0.1 Written July 2004"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   8925
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdParse 
      Caption         =   "&Parse"
      Height          =   375
      Left            =   4800
      TabIndex        =   13
      Top             =   3000
      Width           =   1095
   End
   Begin VB.TextBox txtField 
      Height          =   285
      Left            =   3600
      TabIndex        =   12
      Top             =   2520
      Width           =   5175
   End
   Begin VB.HScrollBar hsbNum 
      Height          =   255
      Left            =   1680
      TabIndex        =   11
      Top             =   2520
      Width           =   1695
   End
   Begin VB.TextBox txtFieldNum 
      Height          =   285
      Left            =   960
      TabIndex        =   10
      Top             =   2520
      Width           =   495
   End
   Begin VB.Frame Frame1 
      Caption         =   "Delimiting Character"
      Height          =   615
      Left            =   1680
      TabIndex        =   3
      Top             =   120
      Width           =   6975
      Begin VB.OptionButton optOther 
         Caption         =   "Other"
         Height          =   255
         Left            =   2160
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optTab 
         Caption         =   "TAB"
         Height          =   255
         Left            =   1200
         TabIndex        =   6
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton optComma 
         Caption         =   "Comma"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.TextBox txtDelimit 
         Height          =   285
         Left            =   3000
         TabIndex        =   4
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "(Length must be 1)"
         Height          =   255
         Left            =   3840
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.TextBox txtString 
      Height          =   1485
      Left            =   960
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   840
      Width           =   7815
   End
   Begin VB.Label Label4 
      Caption         =   "Field:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "String:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Delimit Character:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public tp As New CTextParser
Private nNumField As Integer

Private Sub cmdParse_Click()
  
  'initialize
  hsbNum.Max = 0
  hsbNum.Min = 0
  hsbNum.Value = 0
  txtField.Text = ""
  txtFieldNum.Text = hsbNum.Value
    
  'selects delimiter
  If optComma.Value = True Then
    tp.DelimitChar = ","
  ElseIf optTab.Value = True Then
    tp.DelimitChar = vbTab
  ElseIf optOther.Value = True Then
    If Len(txtDelimit.Text) <> 1 Then Exit Sub
    tp.DelimitChar = LTrim(RTrim(txtDelimit.Text))
  End If
  
  nNumField = tp.ProcessString(txtString.Text)
  If nNumField = 0 Then Exit Sub
  
  hsbNum.Max = nNumField
  hsbNum.Min = 1
  hsbNum.Value = 1
  txtField.Text = tp.GetField(hsbNum.Value)
  txtFieldNum.Text = hsbNum.Value
End Sub

Private Sub hsbNum_Change()
  txtField.Text = tp.GetField(hsbNum.Value)
  txtFieldNum.Text = hsbNum.Value
End Sub

Private Sub hsbNum_Scroll()
  hsbNum_Change
End Sub
