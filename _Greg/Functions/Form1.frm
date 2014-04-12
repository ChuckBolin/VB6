VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7020
   LinkTopic       =   "Form1"
   ScaleHeight     =   4680
   ScaleWidth      =   7020
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdParse 
      Caption         =   "Parse"
      Height          =   315
      Left            =   3600
      TabIndex        =   13
      Top             =   2340
      Width           =   1215
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "String is Empty"
      Height          =   315
      Left            =   3600
      TabIndex        =   12
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox txtIn3 
      Height          =   315
      Left            =   240
      TabIndex        =   11
      Top             =   1920
      Width           =   3195
   End
   Begin VB.CommandButton cmdCube2 
      Caption         =   "Cube Volume VAR"
      Height          =   375
      Left            =   3420
      TabIndex        =   10
      Top             =   240
      Width           =   1575
   End
   Begin VB.TextBox txtIn2b 
      Height          =   285
      Left            =   240
      TabIndex        =   9
      Text            =   "5"
      Top             =   1500
      Width           =   735
   End
   Begin VB.CommandButton cmdXL 
      Caption         =   "XL"
      Height          =   375
      Left            =   1080
      TabIndex        =   8
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox txtOut2 
      Height          =   285
      Left            =   2520
      TabIndex        =   7
      Top             =   1140
      Width           =   735
   End
   Begin VB.TextBox txtIn2a 
      Height          =   285
      Left            =   240
      TabIndex        =   6
      Text            =   "5"
      Top             =   1140
      Width           =   735
   End
   Begin VB.TextBox txtIn1 
      Height          =   285
      Left            =   240
      TabIndex        =   5
      Text            =   "5"
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox txtOut1 
      Height          =   285
      Left            =   2520
      TabIndex        =   4
      Top             =   720
      Width           =   735
   End
   Begin VB.CommandButton cmdArea 
      Caption         =   "Area of Circle"
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   660
      Width           =   1335
   End
   Begin VB.CommandButton cmdCube 
      Caption         =   "Cube Volume"
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox txtOut 
      Height          =   285
      Left            =   2520
      TabIndex        =   1
      Top             =   300
      Width           =   735
   End
   Begin VB.TextBox txtIn 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Text            =   "5"
      Top             =   300
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'constants
Private Const PI = 3.14159

'******************************* F U N C T I O N     C A L L S
Private Sub cmdArea_Click()
  txtOut1.Text = CalcAreaOfCircle(txtIn1.Text)
End Sub

Private Sub cmdCheck_Click()
  'MsgBox StringIsEmpty(txtIn3.Text)
  
  ' OR
  
  Dim bReturn As Boolean
  bReturn = StringIsEmpty(txtIn3.Text)
  MsgBox bReturn
  
End Sub

'calcs cube volume with one value
Private Sub cmdCube_Click()
  txtOut.Text = CalcCubeVolume(txtIn.Text)
End Sub

'demonstrates how the passing value BYREF can be changed...very bad!
'To test, remove BYVAL from function CalcCubeVolume and compare
'results to the function with BYVAL.
Private Sub cmdCube2_Click()
  Dim nLength As Single
  If Not IsNumeric(txtIn.Text) Then Exit Sub
  
  nLength = txtIn.Text
  txtOut.Text = CalcCubeVolume(nLength)
  txtIn.Text = nLength
End Sub

'parses a string delimited by spaces
'a string array is returned
Private Sub cmdParse_Click()
  Dim sData() As String
  Dim i As Integer
  Dim sOut As String
  
  sData = Parse(txtIn3.Text)
  For i = 0 To UBound(sData)
    sOut = sOut & sData(i) & vbCrLf
  Next i
  MsgBox sOut
End Sub

'passes two values to function
Private Sub cmdXL_Click()
  txtOut2.Text = CalcXL(txtIn2a.Text, txtIn2b.Text)
End Sub


'****************************** F U N C T I O N S
'NOTE: All parameters are passed to functions by default as
'BYREF. This is potentially dangerous. Inside the function,
'it is possible to change the value of the variable being passed.
'Try using BYVAL to prevent the function from modifying a passed
'variable.
'**********************************************************
'                 ---------------------
'    nSide ----> [   CalcCubeVolume    ] ----> Volume
'    (Single)     ---------------------        (Single)
'**********************************************************
Private Function CalcCubeVolume(ByVal nSide As Single) As Single
  CalcCubeVolume = nSide * nSide * nSide
  nSide = nSide - 1
End Function

'**********************************************************
'                   ---------------------
'    nRadius ----> [   CalcAreaOfCircle  ] ----> Area
'    (Single)       ---------------------      (Single)
'**********************************************************
Private Function CalcAreaOfCircle(nRadius As Single) As Single
  CalcAreaOfCircle = PI * nRadius * nRadius
End Function

'**********************************************************
'
'                 -------------
'    nFreq ----> |             |
'     (Single)   |   CalcXL    | ----> Volume
'    nInd  ----> |             |      (Single)
'    (Single)     -------------
'
'**********************************************************
Private Function CalcXL(nFrequency As Single, nInductance As Single) As Single
  CalcXL = 2 * PI * nFrequency * nInductance
End Function

'**********************************************************
'                   ------------------
'    sIn     ----> [   StringIsEmpty  ] ----> Status
'    (String)       ------------------       (Boolean)
'**********************************************************
'Returns TRUE if string is empty
Private Function StringIsEmpty(sIn As String) As Boolean
  StringIsEmpty = False
  If Len(sIn) < 1 Then StringIsEmpty = True
End Function

'**********************************************************
'                   ----------
'    sIn     ----> [   Parse  ] ----> Parsed Array
'    (String)       ----------       (Variant)
'**********************************************************
Private Function Parse(sIn As String) As String()
  Parse = Split(sIn, " ")
End Function

'***************************** M I S C E L L A N E O U S
'clears txtOut text as soon as you start typing an input value
Private Sub txtIn_Change()
  txtOut.Text = ""
End Sub

Private Sub txtIn1_Change()
  txtOut1.Text = ""
End Sub

Private Sub txtIn2a_Change()
  txtOut2.Text = ""
End Sub

Private Sub txtIn2b_Change()
  txtOut2.Text = ""
End Sub
