VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PLC Simulator v0.11"
   ClientHeight    =   9600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9600
   ScaleWidth      =   10950
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstCode 
      BackColor       =   &H00C0C0FF&
      Height          =   1425
      ItemData        =   "frmMain.frx":0000
      Left            =   1200
      List            =   "frmMain.frx":0002
      TabIndex        =   24
      Top             =   120
      Width           =   9615
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF0000&
      Height          =   7815
      Left            =   5880
      ScaleHeight     =   7755
      ScaleWidth      =   4875
      TabIndex        =   23
      Top             =   1680
      Width           =   4935
   End
   Begin VB.TextBox txtIL 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3435
      Left            =   1200
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   22
      Top             =   1680
      Width           =   4635
   End
   Begin VB.CommandButton cmdMode 
      Caption         =   "Run Mode"
      Height          =   375
      Left            =   1200
      TabIndex        =   20
      Top             =   5220
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   3435
      Left            =   1200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   19
      Top             =   1680
      Width           =   9615
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   4920
      TabIndex        =   18
      Top             =   9120
      Width           =   855
   End
   Begin VB.Timer tmrUpdate 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3240
      Top             =   6360
   End
   Begin VB.Frame Frame2 
      Caption         =   "Outputs"
      Height          =   2535
      Left            =   120
      TabIndex        =   9
      Top             =   2760
      Width           =   975
      Begin VB.Label lblOut 
         Caption         =   "OUT7"
         Height          =   255
         Index           =   7
         Left            =   360
         TabIndex        =   17
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label lblOut 
         Caption         =   "OUT6"
         Height          =   255
         Index           =   6
         Left            =   360
         TabIndex        =   16
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label lblOut 
         Caption         =   "OUT5"
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   15
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label lblOut 
         Caption         =   "OUT4"
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   14
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lblOut 
         Caption         =   "OUT3"
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   13
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label lblOut 
         Caption         =   "OUT2"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   12
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lblOut 
         Caption         =   "OUT1"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   11
         Top             =   600
         Width           =   495
      End
      Begin VB.Shape shpLED 
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   7
         Left            =   120
         Shape           =   3  'Circle
         Top             =   2040
         Width           =   135
      End
      Begin VB.Shape shpLED 
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   6
         Left            =   120
         Shape           =   3  'Circle
         Top             =   1800
         Width           =   135
      End
      Begin VB.Shape shpLED 
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   5
         Left            =   120
         Shape           =   3  'Circle
         Top             =   1560
         Width           =   135
      End
      Begin VB.Shape shpLED 
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   4
         Left            =   120
         Shape           =   3  'Circle
         Top             =   1320
         Width           =   135
      End
      Begin VB.Shape shpLED 
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   3
         Left            =   120
         Shape           =   3  'Circle
         Top             =   1080
         Width           =   135
      End
      Begin VB.Shape shpLED 
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   2
         Left            =   120
         Shape           =   3  'Circle
         Top             =   840
         Width           =   135
      End
      Begin VB.Shape shpLED 
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   1
         Left            =   120
         Shape           =   3  'Circle
         Top             =   600
         Width           =   135
      End
      Begin VB.Label lblOut 
         Caption         =   "OUT0"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   10
         Top             =   360
         Width           =   495
      End
      Begin VB.Shape shpLED 
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   0
         Left            =   120
         Shape           =   3  'Circle
         Top             =   360
         Width           =   135
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Inputs"
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
      Begin VB.CheckBox chkIn 
         Caption         =   "IN7"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   8
         Top             =   2040
         Width           =   735
      End
      Begin VB.CheckBox chkIn 
         Caption         =   "IN6"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   7
         Top             =   1800
         Width           =   735
      End
      Begin VB.CheckBox chkIn 
         Caption         =   "IN5"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   6
         Top             =   1560
         Width           =   735
      End
      Begin VB.CheckBox chkIn 
         Caption         =   "IN4"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   615
      End
      Begin VB.CheckBox chkIn 
         Caption         =   "IN3"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   735
      End
      Begin VB.CheckBox chkIn 
         Caption         =   "IN2"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Value           =   1  'Checked
         Width           =   615
      End
      Begin VB.CheckBox chkIn 
         Caption         =   "IN1"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CheckBox chkIn 
         Caption         =   "IN0"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Label lblMode 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Editor Mode"
      Height          =   255
      Left            =   2520
      TabIndex        =   21
      Top             =   5280
      Width           =   1515
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************
' PLC Simulator - Written by Chuck Bolin, September 23, 2004
'
' NOTES:
' Program PLC code is loaded in LoadProgram() procedure.
' Boolean expressions are loaded into array g_sCode().Code
' g_nLines states how many lines of code are in the array.
' g_uIn() holds state of PLC inputs.
' g_uOut() holds state of PLC outputs.
' tmrUpdate_Time controls the 10mSec update interval consisting of:
'   ReadInput
'   Substitute boolean expression inputs/bits with 1's or 0's.
'   Determine result of each boolean expression.
'   UpdateOutputDisplay based upon these boolean results.
'
' Version 0.11 - September 24, 2004, Added PROFI IL format
'*************************************************************
Option Explicit
Private de_bug As Boolean
Private m_nPLCMode As PLC_MODE
Private m_sTextCode As String

'************************************************************ LoadProgram
'loads program
Private Sub LoadProgram()
  Dim sCode As String
  Dim i As Integer
  
  g_nLines = 7
  g_sCode(0).Code = "Out0 = In0 + In1 + In2"
  g_sCode(1).Code = "Out1 = !In0"
  g_sCode(2).Code = "Out2 = In2 & In3"
  g_sCode(3).Code = "Out3 = In4 + In5"
  g_sCode(4).Code = "Out4 = In2 & (In3 + In4)"
  g_sCode(5).Code = "Out5 = !In4 + !In5"
  g_sCode(6).Code = "Out6 = In1 & (In2 + in3) & in4 + (in5 + in6) + !in7"
  g_sCode(7).Code = "Out7 = In1 & (In2 + in3)"
  
  'Cleans up code and places it into g_sCode().Code array.
  'Dim sTokens() As String, sList As String
  
  For i = 0 To g_nLines - 1
    g_sCode(i).CleanCode = CleanUpValidateCode(g_sCode(i).Code)
  Next i
  
End Sub

'*********************************************************** Displays Code (Boolean format)
'displays code in txtCode.text
Private Sub DisplayBooleanCode()
  Dim i As Integer
  
  'Places boolean code expressions into text box.
  'txtCode.Text = ""
  lstCode.Clear
  For i = 0 To g_nLines - 1
    'txtCode.Text = txtCode.Text & g_sCode(i).CleanCode & vbCrLf
    lstCode.AddItem g_sCode(i).CleanCode
  Next i

End Sub

'************************************************************ DisplayPROFILadderCode
'Displays code in PROFI Ladder format
Private Sub DisplayPROFILadderCode(nRung As Integer)
  Dim i As Integer, j As Integer
  Dim nWFactor As Integer, nHFactor As Integer
  Dim nEqual As Integer
  Dim sInput As String, sOutput As String, sCode As String, sIL As String
  Dim nFlag As Integer, nFlag2 As Integer
  Dim sTokens() As String
  Dim nRow As Integer, nCol As Integer
   
  nWFactor = pic.Width \ 8
  nHFactor = pic.Height \ 13
  
  'draw ladder diagram tic marks
  pic.Cls
  pic.ForeColor = vbWhite
  For i = 1 To 12
    For j = 1 To 7
      pic.Line (j * nWFactor - 50, i * nHFactor)-(j * nWFactor + 50, i * nHFactor) 'hor tic marks
      pic.Line (j * nWFactor, i * nHFactor - 100)-(j * nWFactor, i * nHFactor + 100) 'vert tic marks
    Next j
  Next

  nRow = 1: nCol = 1 'starting point
  For i = nRung To nRung '0 To g_nLines - 1
    sCode = g_sCode(i).CleanCode
    nEqual = InStr(1, sCode, "=")
    If nEqual > 0 Then 'process only if equal sign
      sOutput = Left(sCode, nEqual - 1) 'this is left of equal sign
      sInput = LTrim(RTrim(Mid(sCode, nEqual + 1)))   'this is right side of equal sign
      'sIL = sIL & ";Rung: " & CStr(i + 1) & vbCrLf
      sTokens = Split(sInput, " ")
      
      'now, evaluate all tokens to construct IL
      nFlag = 0
      For j = 0 To UBound(sTokens)
        
        'ends in right paren ))))))
        If Right(RTrim(sTokens(j)), 1) = ")" Then nFlag2 = 3
        
        'begins with exclamation point !!!!!!!
        If Mid(LTrim(sTokens(j)), 1, 1) = "!" Then
          If nFlag = 0 Then
            'sIL = sIL & "AN" & vbTab & "B" & vbTab & Mid(sTokens(j), 2) & vbCrLf
            'rights input symbol
            pic.ForeColor = vbYellow
            pic.CurrentX = (nCol - 1) * nWFactor + (nWFactor * 0.3)
            pic.CurrentY = nRow * nHFactor - (nHFactor * 0.6)
            pic.Print Mid(sTokens(j), 2)
            
            'draws NC switch
            pic.ForeColor = vbWhite
            pic.Line ((nCol - 1) * nWFactor, nRow * nHFactor)-((nCol - 1) * nWFactor + (nWFactor * 0.4), nRow * nHFactor)
            pic.Line ((nCol - 1) * nWFactor + (nWFactor * 0.4), nRow * nHFactor - 100)-((nCol - 1) * nWFactor + (nWFactor * 0.4), nRow * nHFactor + 100)
            pic.Line ((nCol - 1) * nWFactor + (nWFactor * 0.44), nRow * nHFactor + 80)-((nCol - 1) * nWFactor + (nWFactor * 0.56), nRow * nHFactor - 80)
            pic.Line ((nCol - 1) * nWFactor + (nWFactor * 0.6), nRow * nHFactor - 100)-((nCol - 1) * nWFactor + (nWFactor * 0.6), nRow * nHFactor + 100)
            pic.Line ((nCol - 1) * nWFactor + (nWFactor * 0.6), nRow * nHFactor)-(nCol * nWFactor, nRow * nHFactor)
            
            'nCol = nCol + 1
            nRow = nRow + 1

          ElseIf nFlag = 1 Then  'previous token OR
            'sIL = sIL & "ON" & vbTab & "B" & vbTab & Mid(sTokens(j), 2) & vbCrLf
            'writes input symbol
            pic.ForeColor = vbYellow
            pic.CurrentX = (nCol - 1) * nWFactor + (nWFactor * 0.3)
            pic.CurrentY = nRow * nHFactor - (nHFactor * 0.6)
            pic.Print Mid(sTokens(j), 2)
            
            'draws NC switch
            pic.ForeColor = vbWhite
            pic.Line ((nCol - 1) * nWFactor, nRow * nHFactor)-((nCol - 1) * nWFactor + (nWFactor * 0.4), nRow * nHFactor)
            pic.Line ((nCol - 1) * nWFactor + (nWFactor * 0.4), nRow * nHFactor - 100)-((nCol - 1) * nWFactor + (nWFactor * 0.4), nRow * nHFactor + 100)
            pic.Line ((nCol - 1) * nWFactor + (nWFactor * 0.44), nRow * nHFactor + 80)-((nCol - 1) * nWFactor + (nWFactor * 0.56), nRow * nHFactor - 80)
            pic.Line ((nCol - 1) * nWFactor + (nWFactor * 0.6), nRow * nHFactor - 100)-((nCol - 1) * nWFactor + (nWFactor * 0.6), nRow * nHFactor + 100)
            pic.Line ((nCol - 1) * nWFactor + (nWFactor * 0.6), nRow * nHFactor)-(nCol * nWFactor, nRow * nHFactor)
            pic.Line ((nCol - 1) * nWFactor, (nRow - 1) * nHFactor)-((nCol - 1) * nWFactor, (nRow) * nHFactor) 'drop down line
            
            nFlag = 0
          ElseIf nFlag = 2 Then 'previous token AND
            'writes input symbol
            pic.ForeColor = vbYellow
            pic.CurrentX = (nCol - 1) * nWFactor + (nWFactor * 0.3)
            pic.CurrentY = nRow * nHFactor - (nHFactor * 0.6)
            pic.Print Mid(sTokens(j), 2)
            
            'draws NC switch
            pic.ForeColor = vbWhite
            pic.Line ((nCol - 1) * nWFactor, nRow * nHFactor)-((nCol - 1) * nWFactor + (nWFactor * 0.4), nRow * nHFactor)
            pic.Line ((nCol - 1) * nWFactor + (nWFactor * 0.4), nRow * nHFactor - 100)-((nCol - 1) * nWFactor + (nWFactor * 0.4), nRow * nHFactor + 100)
            pic.Line ((nCol - 1) * nWFactor + (nWFactor * 0.44), nRow * nHFactor + 80)-((nCol - 1) * nWFactor + (nWFactor * 0.56), nRow * nHFactor - 80)
            pic.Line ((nCol - 1) * nWFactor + (nWFactor * 0.6), nRow * nHFactor - 100)-((nCol - 1) * nWFactor + (nWFactor * 0.6), nRow * nHFactor + 100)
            pic.Line ((nCol - 1) * nWFactor + (nWFactor * 0.6), nRow * nHFactor)-(nCol * nWFactor, nRow * nHFactor)
            
            nFlag = 0
            nCol = nCol + 1
            'sIL = sIL & "AN" & vbTab & "B" & vbTab & Mid(sTokens(j), 2) & vbCrLf
          End If
        
        'begins with open paren and exclamation point (!  (!  (!  (!  (!
        ElseIf Mid(LTrim(sTokens(j)), 1, 2) = "(!" Then
            'writes input symbol
            pic.ForeColor = vbYellow
            pic.CurrentX = (nCol - 1) * nWFactor + (nWFactor * 0.3)
            pic.CurrentY = nRow * nHFactor - (nHFactor * 0.6)
            pic.Print Mid(sTokens(j), 2)
            
            'draws NC switch
            pic.ForeColor = vbWhite
            pic.Line ((nCol - 1) * nWFactor, nRow * nHFactor)-((nCol - 1) * nWFactor + (nWFactor * 0.4), nRow * nHFactor)
            pic.Line ((nCol - 1) * nWFactor + (nWFactor * 0.4), nRow * nHFactor - 100)-((nCol - 1) * nWFactor + (nWFactor * 0.4), nRow * nHFactor + 100)
            pic.Line ((nCol - 1) * nWFactor + (nWFactor * 0.44), nRow * nHFactor + 80)-((nCol - 1) * nWFactor + (nWFactor * 0.56), nRow * nHFactor - 80)
            pic.Line ((nCol - 1) * nWFactor + (nWFactor * 0.6), nRow * nHFactor - 100)-((nCol - 1) * nWFactor + (nWFactor * 0.6), nRow * nHFactor + 100)
            pic.Line ((nCol - 1) * nWFactor + (nWFactor * 0.6), nRow * nHFactor)-(nCol * nWFactor, nRow * nHFactor)
            
            nFlag = 0

          'sIL = sIL & "(" & vbCrLf & "AN" & vbTab & "B" & vbTab & Mid(sTokens(j), 3) & vbCrLf
        ElseIf Mid(LTrim(sTokens(j)), 1, 1) = "(" Then
          If nFlag = 1 Then 'OR before (
            'sIL = sIL & "O(" & vbCrLf & "A" & vbTab & "B" & vbTab & Mid(sTokens(j), 2) & vbCrLf
            nFlag = 0
          ElseIf nFlag = 2 Then 'AND before (
            'sIL = sIL & "A(" & vbCrLf & "A" & vbTab & "B" & vbTab & Mid(sTokens(j), 2) & vbCrLf
            nFlag = 0
          Else
          
          End If
          
        'OR
        ElseIf Mid(LTrim(sTokens(j)), 1, 1) = "+" Then
          nFlag = 1 'OR
        
        'AND
        ElseIf Mid(LTrim(sTokens(j)), 1, 1) = "&" Then
          nFlag = 2 'AND
        
        'Everything else
        Else
          If nFlag = 1 And nFlag2 = 0 Then 'OR
            'sIL = sIL & "O" & vbTab & "B" & vbTab & sTokens(j) & vbCrLf
          ElseIf nFlag = 1 And nFlag2 = 3 Then 'OR with closed paren
            'sIL = sIL & "O" & vbTab & "B" & vbTab & Left(sTokens(j), Len(sTokens(j)) - 1) & vbCrLf & ")" & vbCrLf
            nFlag2 = 0
          ElseIf nFlag = 2 And nFlag2 = 0 Then 'AND
            'sIL = sIL & "A" & vbTab & "B" & vbTab & sTokens(j) & vbCrLf
          ElseIf nFlag = 2 And nFlag2 = 3 Then 'AND with closed paren
            'sIL = sIL & "A" & vbTab & "B" & vbTab & Left(sTokens(j), Len(sTokens(j)) - 1) & vbCrLf & ")" & vbCrLf
            nFlag = 0
          Else
            
            'prints input symbol
            pic.ForeColor = vbYellow
            pic.CurrentX = (nCol - 1) * nWFactor + (nWFactor * 0.3)
            pic.CurrentY = nRow * nHFactor - (nHFactor * 0.6)
            pic.Print sTokens(j)
            
            'draws NO switch
            pic.ForeColor = vbWhite
            pic.Line ((nCol - 1) * nWFactor, nRow * nHFactor)-((nCol - 1) * nWFactor + (nWFactor * 0.4), nRow * nHFactor)
            pic.Line ((nCol - 1) * nWFactor + (nWFactor * 0.4), nRow * nHFactor - 100)-((nCol - 1) * nWFactor + (nWFactor * 0.4), nRow * nHFactor + 100)
            pic.Line ((nCol - 1) * nWFactor + (nWFactor * 0.6), nRow * nHFactor - 100)-((nCol - 1) * nWFactor + (nWFactor * 0.6), nRow * nHFactor + 100)
            pic.Line ((nCol - 1) * nWFactor + (nWFactor * 0.6), nRow * nHFactor)-(nCol * nWFactor, nRow * nHFactor)
            nCol = nCol + 1
            'sIL = sIL & "A" & vbTab & "B" & vbTab & sTokens(j) & vbCrLf
          End If
        End If
      Next j
      
      'print output address
      pic.ForeColor = vbYellow
      pic.CurrentX = 7 * nWFactor + (nWFactor * 0.05)
      pic.CurrentY = 1 * nHFactor - (nHFactor * 0.6)
      pic.Print sOutput
      
      'draws open and closed paren
      pic.Line (7 * nWFactor, 1 * nHFactor)-(7 * nWFactor + (nWFactor * 0.3) - 60, 1 * nHFactor)
      pic.Circle (7 * nWFactor + (nWFactor * 0.3) + 60, 1 * nHFactor), 130, , 1.74, 4.34  'left paren
      pic.Circle (8 * nWFactor - (nWFactor * 0.2) - 190, 1 * nHFactor), 130, , 5.08, 1.47 'right paren
      pic.Line (8 * nWFactor - (nWFactor * 0.2) - 70, 1 * nHFactor)-(8 * nWFactor, 1 * nHFactor)
      
      'prints equal sign
      pic.ForeColor = vbRed
      pic.CurrentX = 7 * nWFactor + (nWFactor * 0.35)
      pic.CurrentY = 1 * nHFactor - (nHFactor * 0.2)
      pic.Print "="
      
      
      'sIL = sIL & "=" & vbTab & "B" & vbTab & sOutput & vbCrLf & vbCrLf
      'txtIL.Text = sIL & vbCrLf
    End If
  Next i
  
  
  'txtIL.Text = txtIL.Text & "EP" & vbCrLf
End Sub


'************************************************************ DisplayPROFIILCode
'Displays code in PROFI IL format
Private Sub DisplayPROFIILCode(nRung As Integer)
  Dim i As Integer, nEqual As Integer, j As Integer
  Dim sInput As String, sOutput As String, sCode As String, sIL As String
  Dim nFlag As Integer, nFlag2 As Integer
  Dim sTokens() As String
 
  For i = nRung To nRung '0 To g_nLines - 1
    sCode = g_sCode(i).CleanCode
    nEqual = InStr(1, sCode, "=")
    If nEqual > 0 Then 'process only if equal sign
      sOutput = Left(sCode, nEqual - 1) 'this is left of equal sign
      sInput = LTrim(RTrim(Mid(sCode, nEqual + 1)))   'this is right side of equal sign
      sIL = sIL & ";Rung: " & CStr(i + 1) & vbCrLf
      sTokens = Split(sInput, " ")
      
      'now, evaluate all tokens to construct IL
      nFlag = 0
      For j = 0 To UBound(sTokens)
        If Right(RTrim(sTokens(j)), 1) = ")" Then nFlag2 = 3
        If Mid(LTrim(sTokens(j)), 1, 1) = "!" Then
          If nFlag = 0 Then
            sIL = sIL & "AN" & vbTab & "B" & vbTab & Mid(sTokens(j), 2) & vbCrLf
          ElseIf nFlag = 1 Then  'previous token OR
            sIL = sIL & "ON" & vbTab & "B" & vbTab & Mid(sTokens(j), 2) & vbCrLf
            nFlag = 0
          ElseIf nFlag = 2 Then 'previous token AND
            sIL = sIL & "AN" & vbTab & "B" & vbTab & Mid(sTokens(j), 2) & vbCrLf
          End If
        ElseIf Mid(LTrim(sTokens(j)), 1, 2) = "(!" Then
          sIL = sIL & "(" & vbCrLf & "AN" & vbTab & "B" & vbTab & Mid(sTokens(j), 3) & vbCrLf
        ElseIf Mid(LTrim(sTokens(j)), 1, 1) = "(" Then
          If nFlag = 1 Then 'OR before (
            sIL = sIL & "O(" & vbCrLf & "A" & vbTab & "B" & vbTab & Mid(sTokens(j), 2) & vbCrLf
            nFlag = 0
          ElseIf nFlag = 2 Then 'AND before (
            sIL = sIL & "A(" & vbCrLf & "A" & vbTab & "B" & vbTab & Mid(sTokens(j), 2) & vbCrLf
            nFlag = 0
          Else
          
          End If
          
        ElseIf Mid(LTrim(sTokens(j)), 1, 1) = "+" Then
          nFlag = 1 'OR
        ElseIf Mid(LTrim(sTokens(j)), 1, 1) = "&" Then
          nFlag = 2 'AND
        
        Else
          If nFlag = 1 And nFlag2 = 0 Then 'OR
            sIL = sIL & "O" & vbTab & "B" & vbTab & sTokens(j) & vbCrLf
          ElseIf nFlag = 1 And nFlag2 = 3 Then 'OR with closed paren
            sIL = sIL & "O" & vbTab & "B" & vbTab & Left(sTokens(j), Len(sTokens(j)) - 1) & vbCrLf & ")" & vbCrLf
            nFlag2 = 0
          ElseIf nFlag = 2 And nFlag2 = 0 Then 'AND
            sIL = sIL & "A" & vbTab & "B" & vbTab & sTokens(j) & vbCrLf
          ElseIf nFlag = 2 And nFlag2 = 3 Then 'AND with closed paren
            sIL = sIL & "A" & vbTab & "B" & vbTab & Left(sTokens(j), Len(sTokens(j)) - 1) & vbCrLf & ")" & vbCrLf
            nFlag = 0
          Else
            sIL = sIL & "A" & vbTab & "B" & vbTab & sTokens(j) & vbCrLf
          End If
          
          
        End If
        
        
      Next j
      sIL = sIL & "=" & vbTab & "B" & vbTab & sOutput & vbCrLf & vbCrLf
      txtIL.Text = sIL & vbCrLf
    End If
  Next i
  txtIL.Text = txtIL.Text & "EP" & vbCrLf
End Sub


'********************************************************** FormLoad
'Program initialization
Private Sub Form_Load()
  LoadVariables
  LoadPatternsReplacements
  LoadProgram
  DisplayBooleanCode
  WriteListFromString (m_sTextCode)
  'm_sTextCode = WriteStringFromList 'txtCode.Text

  txtIL.Text = "Select Boolean expression above."
  DisplayPROFIILCode 6
  DisplayPROFILadderCode 6
End Sub

'writes list contents to a string
Private Function WriteStringFromList() As String
  Dim sOut As String, i As Integer
  For i = 0 To lstCode.ListCount - 1
    sOut = sOut & lstCode.List(i)
  Next i
  WriteStringFromList = sOut
End Function

'reads a string and loads into listbox
Private Sub WriteListFromString(sIn As String)
  Dim i As Integer, sCode() As String
  If Len(sIn) < 1 Then Exit Sub
  
  'grab code from string, vbCrLf delimited
  sCode = Split(sIn, vbCrLf)
  lstCode.Clear
  For i = 0 To UBound(sCode)
    lstCode.AddItem sCode(i)
  Next i
  
End Sub


'*************************************************************  Exit program
'quit program
Private Sub cmdExit_Click()
  End
End Sub

'*********************************************************** Change Modes
'Switches between Run and Stopped (Edit) mode
Private Sub cmdMode_Click()
  Dim i As Integer
  
  'starts run mode
  If m_nPLCMode = Edit Then
    m_nPLCMode = Run
    cmdMode.Caption = "Edit Mode"
    lblMode.Caption = "Running!"
    lblMode.BackColor = vbGreen
    tmrUpdate.Enabled = True
    'txtCode.Text = m_sTextCode
  'starts edit or stop mode
  Else
    m_nPLCMode = Edit
    cmdMode.Caption = "Run Mode"
    lblMode.Caption = "Stopped!"
    lblMode.BackColor = vbRed
    tmrUpdate.Enabled = False
    
    'turns outputs off
    For i = 0 To 7
      g_uOut(i).Value = 0
    Next i
    UpdateOutputDisplay
    
  End If
End Sub

Private Sub lstCode_Click()
  DisplayPROFIILCode lstCode.ListIndex
  DisplayPROFILadderCode lstCode.ListIndex
End Sub

'************************************************************* Restores code in text box
'after attempts to edit, at run time, the latest correct
'code is displayed
Private Sub txtCode_Change()
  If m_nPLCMode = Run Then txtCode.Text = m_sTextCode
End Sub


'********************************************************** LoadVariables
Private Sub LoadVariables()
  Dim i As Integer
  de_bug = False 'set true to debug one boolean equation
  If de_bug = True Then
    Text1.Visible = True
  Else
    Text1.Visible = False
  End If
  
  'mode setup
  m_nPLCMode = Edit
  cmdMode.Caption = "Run Mode"
  
  For i = 0 To MAX_INPUTS - 1
    g_uIn(i).Absolute = "IN" & CStr(i)
    g_uIn(i).Symbol = "S" & CStr(i)
    chkIn(i).ToolTipText = g_uIn(i).Symbol
  Next i
  For i = 0 To MAX_OUTPUTS - 1
    g_uOut(i).Absolute = "OUT" & CStr(i)
    g_uOut(i).Symbol = "S" & CStr(i)
  Next i
  
  'These are legal characters and operators
  g_sOperators = "()&+!="
  g_sLegalCharacters = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789-@"

  
End Sub

'*************************************************** LoadPatternsReplacements
'All input and output addresses in PLC code will be replaced by 1's and 0's.
'After this, the program must reduce the value to a single solution 1 or 0
'or -1(error)
Private Sub LoadPatternsReplacements()
  bp(0).Pattern = "!0"       'NOT
  bp(0).Replacement = "1"
  bp(1).Pattern = "!1"       'NOT
  bp(1).Replacement = "0"
  bp(2).Pattern = "(0)"      '(x)
  bp(2).Replacement = "0"
  bp(3).Pattern = "(1)"
  bp(3).Replacement = "1"
  bp(4).Pattern = "0&0"      'AND
  bp(4).Replacement = "0"
  bp(5).Pattern = "0&1"
  bp(5).Replacement = "0"
  bp(6).Pattern = "1&0"
  bp(6).Replacement = "0"
  bp(7).Pattern = "1&1"
  bp(7).Replacement = "1"
  bp(8).Pattern = "0+0"      'OR
  bp(8).Replacement = "0"
  bp(9).Pattern = "0+1"
  bp(9).Replacement = "1"
  bp(10).Pattern = "1+0"
  bp(10).Replacement = "1"
  bp(11).Pattern = "1+1"
  bp(11).Replacement = "1"
  bp(12).Pattern = "00"      'AND w/o symbol &
  bp(12).Replacement = "0"
  bp(13).Pattern = "01"
  bp(13).Replacement = "0"
  bp(14).Pattern = "10"
  bp(14).Replacement = "0"
  bp(15).Pattern = "11"
  bp(15).Replacement = "1"
  bp(16).Pattern = "0"
  bp(16).Replacement = "0"
  bp(17).Pattern = "1"
  bp(17).Replacement = "1"
  End Sub

'**********************************************************  Scan Cycle
'**********************************************************
'**********************************************************    10 mSec
'**********************************************************
'********************************************************** tmrUpdate_Timer
'This happens every 10mSec. Updates output based upon inputs and PLC code
Private Sub tmrUpdate_Timer()
  Dim nRet As Integer, i As Integer, nEqual As Integer, j As Integer
  
  'read current inputs into g_uIn()
  ReadInputs
  
  'substitute actual 1's and 0's into code
  For i = 0 To g_nLines
    g_sCode(i).Substitute = SubstituteNumbers(g_sCode(i).CleanCode)
    If de_bug = True Then frmMain.Caption = g_sCode(i).Substitute
  Next i
  
  'derive correct boolean answers
  For i = 0 To g_nLines - 1
    'nEqual = InStr(1, g_sCode(i).CleanCode, "=")
    'nRet = GetSolution(RemoveSpaces((Mid(g_sCode(i).Substitute, nEqual + 1))))
    nRet = GetSolution(g_sCode(i).Substitute)
    If nRet = -1 Then
      MsgBox "Syntax error: " & g_sCode(i).CleanCode
      End
    Else
      nEqual = InStr(1, g_sCode(i).CleanCode, "=")
      If nEqual > 0 Then
        g_sCode(i).ResultString = LTrim(Trim(Left(g_sCode(i).CleanCode, nEqual - 1)))
        g_sCode(i).Result = nRet
        For j = 0 To MAX_OUTPUTS - 1
          If g_sCode(i).ResultString = g_uOut(j).Absolute Then
            g_uOut(j).Value = nRet
          End If
        Next j
      End If
    End If
  Next i
  
  'update display to reflect logical answers to code
  UpdateOutputDisplay 'change g_byOut based upon code
End Sub


'***************************************************** ReadsInput
'reads inputs (check boxes) into g_byIn
' vbUnchecked = 0, vbChecked = 1
Private Sub ReadInputs()
  Dim i As Integer
  For i = 0 To MAX_INPUTS - 1
    g_uIn(i).Value = chkIn(i).Value
  Next i
End Sub

'********************************************************** SubstituteNumbers
'Takes a boolean string with abolute addresses (i.e. IN3,OUT4)
'and replaces the addresses to the right of the equal sign
'with a 1 or 0 read from ReadInput function
Private Function SubstituteNumbers(sIn As String) As String
  Dim sTokens() As String, sTemp As String, sOut As String
  Dim i As Integer, j As Integer
  Dim nPos As Integer
  Dim sSymbol As String
  If Len(sIn) < 1 Then Exit Function
  
  sTokens = Split(sIn, " ") 'loads all tokens into array sTokens
  
  
  'replace all symbols INx and OUTx with 1's or 0's.
  For i = 0 To UBound(sTokens)
    If Len(sTokens(i)) > 0 Then 'process only token with something there
      
      'look for INPUT absolute address
      If InStr(1, sTokens(i), "IN") > 0 Then
        nPos = InStr(1, sTokens(i), "IN") 'get starting positin of IN
        sSymbol = Mid(sTokens(i), nPos)   'grab token INx...maybe !INx
        
        If Left(sTokens(i), 1) = "!" Then 'must process not symbol before absolute
            sSymbol = Mid(sTokens(i), 2)
          For j = 0 To MAX_INPUTS - 1
          
            If sSymbol = g_uIn(j).Absolute Then
              If g_uIn(j).Value = 1 Then
                sTokens(i) = "0"
              Else
                sTokens(i) = "1"
              End If
            End If
          Next j
        
        ElseIf Left(sTokens(i), 1) = "(" Then 'must process not symbol before absolute
       
          sSymbol = Mid(sTokens(i), 2)
          For j = 0 To MAX_INPUTS - 1
            If sSymbol = g_uIn(j).Absolute Then
              If g_uIn(j).Value = 1 Then
                sTokens(i) = "(1"
              Else
                sTokens(i) = "(0"
              End If
            End If
          Next j
          
        ElseIf Right(sTokens(i), 1) = ")" Then 'must process not symbol before absolute
          sSymbol = Left(sTokens(i), Len(sTokens(i)) - 1) ' Mid(sTokens(i), 2)
          
          For j = 0 To MAX_INPUTS - 1
            If sSymbol = g_uIn(j).Absolute Then
              If g_uIn(j).Value = 1 Then
                sTokens(i) = "1)"
              Else
                sTokens(i) = "0)"
              End If
            End If
          Next j
          
          
        Else  'no extra !, ( or )
          For j = 0 To MAX_INPUTS - 1
            If sSymbol = g_uIn(j).Absolute Then
              sTokens(i) = CStr(g_uIn(j).Value)
            End If
          Next j
        End If
        
      ElseIf InStr(1, sTokens(i), "OUT") > 0 Then
        nPos = InStr(1, sTokens(i), "OUT")
        sSymbol = Mid(sTokens(i), nPos)
        For j = 0 To MAX_INPUTS - 1
          If sSymbol = g_uIn(j).Absolute Then
            sTokens(i) = CStr(g_uIn(j).Value)
          End If
        Next j
      Else
        'do nothing with this token
      End If
    
    End If
  Next i
  
  'reassemble string from tokens
  For i = 0 To UBound(sTokens)
    If Len(sTokens(i)) > 0 Then
      sOut = sOut & sTokens(i) & " "
    End If
  Next i

  SubstituteNumbers = sOut
  
End Function

'******************************************************* GetSolution
'Determines result based upon solution of boolean expression
Private Function GetSolution(sIn As String) As Integer
  GetSolution = -1
  If Len(sIn) < 1 Then Exit Function
        
  Dim i As Integer, j As Integer, nCt As Integer, nRet As Integer, nErr As Integer
  Dim sOut As String, sTemp As String, intEqual As Integer
  
  'removes spaces and verifies existence of equal sign
  sIn = RemoveSpaces(sIn)
  intEqual = InStr(1, sIn, "=")
  If intEqual > 0 Then
    sIn = Mid(sIn, intEqual + 1)
  Else
    MsgBox "Invalid expression: " & sIn & ". Aborting program"
    End
  End If
 
  If de_bug = True Then Text1.Text = Text1.Text & "START:   " & sIn & vbCrLf
  
  'look for patterns in code line from bp() array and replaces
  Do
    sOut = ""
    nErr = 0
    For j = 0 To MAX_PATTERNS
      nRet = InStr(1, sIn, bp(j).Pattern)
      If nRet > 0 Then
        sOut = ReplaceSubstring(sIn, bp(j).Pattern, bp(j).Replacement)
        sIn = sOut
        If de_bug = True Then Text1.Text = Text1.Text & j & " - " & " : " & bp(j).Pattern & " : " & bp(j).Replacement & "  -  " & sIn & vbCrLf
        nErr = 1
        Exit For
      End If
    Next j
    If nErr = 0 Then
      MsgBox "Stuck processing " & sIn & " . Exiting program!"
      End
    End If
  Loop Until Len(sIn) = 1
  If de_bug = True Then tmrUpdate.Enabled = False
  GetSolution = Val(sIn)
End Function


'***************************************************** UpdateOutputDisplay
'updates outputs (shapes) by g_byOut
Private Sub UpdateOutputDisplay()
  Dim i As Integer
  
  For i = 0 To MAX_OUTPUTS - 1
    If g_uOut(i).Value = 1 Then
      shpLED(i).FillColor = RGB(0, 255, 0)
    Else
      shpLED(i).FillColor = RGB(0, 155, 0)
    End If
  Next i

End Sub

Private Function ValidateCharactersInCode(sIn As String) As String
  ValidateCharactersInCode = ""
  If Len(sIn) < 1 Then Exit Function
  Dim i As Integer, sChar As String
  
  For i = 1 To Len(sIn)
    sChar = Mid(sIn, i, 1)
    If (InStr(1, g_sOperators, sChar) < 1) And (InStr(1, g_sLegalCharacters, sChar) < 1) Then
        ValidateCharactersInCode = sChar
    End If
  Next i
  
End Function
'********************************************************** CleanUpValidateCode
'cleans up code...spaces before and after certain operators
'this is done for proper display and writing to a file
'Operators:  = equals
'            & AND
'            | OR
Private Function CleanUpValidateCode(sIn As String) As String
  Dim sNew As String, sTemp As String, sRet As String
  
  Dim i As Integer
  Dim sTokens() As String
  
  'changes characters to upper case
  sIn = UCase(RemoveSpaces(sIn))
  sRet = ValidateCharactersInCode(sIn)
  If Len(sRet) > 0 Then
    MsgBox sRet & " is not a valid character or operator in code " & sIn
    End
  End If
  'validate tokens (non-operators)...must be INx or OUTx
  'replaces all operators with '*'. Everything else must be an absolute or
  'symbolic address
  For i = 1 To Len(sIn)
    If Mid(sIn, i, 1) = "=" Or Mid(sIn, i, 1) = "!" Or Mid(sIn, i, 1) = "&" Or Mid(sIn, i, 1) = "+" Or Mid(sIn, i, 1) = "(" Or Mid(sIn, i, 1) = ")" Then
      sTemp = sTemp & "*"
    Else
      sTemp = sTemp & Mid(sIn, i, 1)
    End If
  Next i
  sTokens = Split(sTemp, "*") 'loads all tokens into array sTokens
  
  'each token must be INx or OUTx
  For i = 0 To UBound(sTokens)
    If Len(sTokens(i)) > 0 Then
      If (Left(sTokens(i), 2) = "IN" And IsNumeric(Mid(sTokens(i), 3))) Or (Left(sTokens(i), 3) = "OUT" And IsNumeric(Mid(sTokens(i), 4))) Then
        'legitimate...do nothing
      Else 'unrecognized token
         MsgBox sTokens(i) & " is incorrect.Check spelling.", vbOKOnly, "Illegal Operand!"
         End
      End If
    End If
  Next i
  
  'add proper spacing about operators
  For i = 1 To Len(sIn)
    If Mid(sIn, i, 1) = "=" Then  'replace [=] with [space = space]
      sNew = sNew & " = "
    ElseIf Mid(sIn, i, 2) = "&!" Then  'replace [&!] with [space &! space]
      sNew = sNew & " & !"
      i = i + 1
    ElseIf Mid(sIn, i, 2) = "+!" Then  'replace [|!] with [space |! space]
      sNew = sNew & " + !"
      i = i + 1
    ElseIf Mid(sIn, i, 1) = "&" Then  'replace [&] with [space & space]
      sNew = sNew & " & "
    ElseIf Mid(sIn, i, 1) = "+" Then  'replace [&] with [space | space]
      sNew = sNew & " + "
    Else  'keeper
      sNew = sNew & (Mid(sIn, i, 1))
    End If
  Next i

  CleanUpValidateCode = sNew
End Function


'********************************************************** RemoveSpaces
'Removes all spaces from a string
Private Function RemoveSpaces(sIn) As String
  Dim i As Integer
  Dim sTemp As String
  
  For i = 1 To Len(sIn)
    If Mid(sIn, i, 1) <> " " Then sTemp = sTemp & Mid(sIn, i, 1)
  Next i
  RemoveSpaces = sTemp
End Function

'********************************************************** ReplaceSubstring
'Replaces a substring in a larger string with another substring.
Private Function ReplaceSubstring(sIn As String, sFind As String, sReplace As String) As String
  Dim sNew As String, sBefore As String, sAfter As String
  Dim nPos As Integer, nLen As Integer
  
  sNew = sIn
  nPos = InStr(1, sIn, sFind)
  sBefore = "": sAfter = ""
  sAfter = Right(sIn, 1)
  'MsgBox "Here: " & sFind & "  " & sIn & "   " & sReplace & "  " & nPos
  If nPos > 0 Then
    nLen = Len(sFind)
    sBefore = Left(sIn, nPos - 1)
    sAfter = Mid(sIn, nPos + nLen)
    sNew = sBefore & sReplace & sAfter
  End If
   
  If de_bug = True Then
    Text1.Text = Text1.Text & vbCrLf & "F: " & sFind & vbCrLf
    Text1.Text = Text1.Text & "B: " & sBefore & vbCrLf
    Text1.Text = Text1.Text & "M: " & sReplace & vbCrLf
    Text1.Text = Text1.Text & "A: " & sAfter & vbCrLf
  End If
  
  ReplaceSubstring = sNew
  
End Function

