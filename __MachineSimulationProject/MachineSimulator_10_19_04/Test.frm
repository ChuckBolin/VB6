VERSION 5.00
Begin VB.Form Test 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test Form for CBooleanEvaluator"
   ClientHeight    =   4545
   ClientLeft      =   3120
   ClientTop       =   2115
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   6855
   Begin VB.TextBox txtProg 
      Height          =   1935
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   90
      Width           =   6765
   End
End
Attribute VB_Name = "Test"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim boo As New CBooleanEvaluator

Private Sub Form_Load()
  Dim sProg As String, nRet As Integer
  Dim sProgArray() As Variant
  
  sProg = "OUT3 = IN1 & IN2" & vbCrLf
  sProg = sProg & "OUT4 = IN3 & IN4"
  
  nRet = boo.LoadProgram(sProg)
  txtProg.Text = boo.GetProgramAsString
  sProgArray = boo.GetProgramAsArray
End Sub
