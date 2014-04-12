VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PLC Simulator v0.1"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   8985
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   2955
      Left            =   1200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   20
      Top             =   2160
      Width           =   7695
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   5940
      TabIndex        =   19
      Top             =   5220
      Width           =   855
   End
   Begin VB.Timer tmrUpdate 
      Interval        =   10
      Left            =   5940
      Top             =   5640
   End
   Begin VB.Frame Frame2 
      Caption         =   "Outputs"
      Height          =   2535
      Left            =   120
      TabIndex        =   10
      Top             =   2760
      Width           =   975
      Begin VB.Label lblOut 
         Caption         =   "OUT7"
         Height          =   255
         Index           =   7
         Left            =   360
         TabIndex        =   18
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label lblOut 
         Caption         =   "OUT6"
         Height          =   255
         Index           =   6
         Left            =   360
         TabIndex        =   17
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label lblOut 
         Caption         =   "OUT5"
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   16
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label lblOut 
         Caption         =   "OUT4"
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   15
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lblOut 
         Caption         =   "OUT3"
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   14
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label lblOut 
         Caption         =   "OUT2"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   13
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lblOut 
         Caption         =   "OUT1"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   12
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
         TabIndex        =   11
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
      TabIndex        =   1
      Top             =   120
      Width           =   975
      Begin VB.CheckBox chkIn 
         Caption         =   "IN7"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   9
         Top             =   2040
         Width           =   735
      End
      Begin VB.CheckBox chkIn 
         Caption         =   "IN6"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   8
         Top             =   1800
         Width           =   735
      End
      Begin VB.CheckBox chkIn 
         Caption         =   "IN5"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   7
         Top             =   1560
         Width           =   735
      End
      Begin VB.CheckBox chkIn 
         Caption         =   "IN4"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   615
      End
      Begin VB.CheckBox chkIn 
         Caption         =   "IN3"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   735
      End
      Begin VB.CheckBox chkIn 
         Caption         =   "IN2"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   615
      End
      Begin VB.CheckBox chkIn 
         Caption         =   "IN1"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   735
      End
      Begin VB.CheckBox chkIn 
         Caption         =   "IN0"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.TextBox txtCode 
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
      Height          =   1875
      Left            =   1200
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Width           =   7695
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
'*************************************************************
Option Explicit
Private de_bug As Boolean


'************************************************************ LoadProgram
'loads program
Private Sub LoadProgram()
  Dim sCode As String
  Dim i As Integer
  g_nLines = 1
  'g_sCode(0).Code = "Out0 = In0"
  'g_sCode(1).Code = "Out1 = !In1"
  'g_sCode(2).Code = "Out2 = In2 & In3"
  'g_sCode(3).Code = "Out3 = In4 + In5"
  'g_sCode(4).Code = "Out4 = !In2 & !In3"
  'g_sCode(5).Code = "Out5 = !In4 + !In5"
  g_sCode(0).Code = "Out6 = In1 & In2 + in3 & in4 + in5 + in6 + !in7"
  
  txtCode.Text = ""
  For i = 0 To g_nLines - 1
    g_sCode(i).CleanCode = CleanUpValidateCode(g_sCode(i).Code)
    txtCode.Text = txtCode.Text & g_sCode(i).CleanCode & vbCrLf
  Next i
  
End Sub

'quit program
Private Sub cmdExit_Click()
  End
End Sub

'********************************************************** FormLoad
'Program initialization
Private Sub Form_Load()
  
  LoadVariables
  LoadPatternsReplacements
  LoadProgram
End Sub

'********************************************************** LoadVariables
Private Sub LoadVariables()
  Dim i As Integer
  de_bug = False
  For i = 0 To 7
    g_uIn(i).Absolute = "IN" & CStr(i)
    g_uOut(i).Absolute = "OUT" & CStr(i)
  Next i
End Sub

'*************************************************** LoadPatternsReplacements
'all input and output addresses in PLC code will be replaced by 1's and 0's
'after this, the program must reduce the value to a single solution 1 or 0 or -1(error)
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

End Sub
'**********************************************************  Scan Cycle
'**********************************************************
'**********************************************************    10 mSec
'**********************************************************
'********************************************************** tmrUpdate_Timer
'This happens every 10mSec. Updates output based upon inputs and PLC code
Private Sub tmrUpdate_Timer()
  Dim nRet As Integer, i As Integer, nEqual As Integer, j As Integer
  
  'read current inputs
  For j = 0 To 7
    g_uIn(j).Value = 0
  Next j
  ReadInputs  'read into g_uIn(7)
  
  'substitute actual 1's and 0's into code
  For i = 0 To g_nLines
    g_sCode(i).Substitute = SubstituteNumbers(g_sCode(i).CleanCode)
    'If i = 6 Then frmMain.Caption = g_sCode(i).Substitute
  Next i
  
  'derive correct boolean answers
  For i = 0 To g_nLines - 1
    nEqual = InStr(1, g_sCode(i).CleanCode, "=")
    nRet = GetSolution(RemoveSpaces((Mid(g_sCode(i).Substitute, nEqual + 1))))
    If nRet = -1 Then
      MsgBox "Syntax error: " & g_sCode(i).CleanCode
      End
    Else
      nEqual = InStr(1, g_sCode(i).CleanCode, "=")
      If nEqual > 0 Then
        g_sCode(i).ResultString = LTrim(Trim(Left(g_sCode(i).CleanCode, nEqual - 1)))
        g_sCode(i).Result = nRet
        For j = 0 To 7
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

'********************************************************** SubstituteNumbers
'Takes a boolean string with abolute addresses (i.e. IN3,OUT4)
'and replaces the addresses to the right of the equal sign
'with a 1 or 0 read from ReadInput function
Private Function SubstituteNumbers(sIn As String) As String
  Dim sTokens() As String, sTemp As String, sOut As String
  Dim i As Integer, j As Integer
  Dim nPos As Integer
  Dim sSymbol As String
        
  sTokens = Split(sIn, " ") 'loads all tokens into array sTokens
  
  'replace all symbols INx and OUTx with 1's or 0's.
  For i = 0 To UBound(sTokens)
    If Len(sTokens(i)) > 0 Then 'process only token with something there
      
      'look for INPUT absolute address
      If InStr(1, sTokens(i), "IN") > 0 Then
        nPos = InStr(1, sTokens(i), "IN") 'get starting positin of IN
        sSymbol = Mid(sTokens(i), nPos)   'grab token INx...maybe !INx
        
        If Left(sTokens(i), 1) = "!" Then 'must process not symbol before absolute
          For j = 0 To 7
            sSymbol = Mid(sTokens(i), 2)
            If sSymbol = g_uIn(j).Absolute Then
              If g_uIn(j).Value = 1 Then
                sTokens(i) = "0"
              Else
                sTokens(i) = "1"
              End If
            End If
          Next j
        Else  'no not
          For j = 0 To 7
            If sSymbol = g_uIn(j).Absolute Then
              sTokens(i) = CStr(g_uIn(j).Value)
            End If
          Next j
        End If
      ElseIf InStr(1, sTokens(i), "OUT") > 0 Then
        nPos = InStr(1, sTokens(i), "OUT")
        sSymbol = Mid(sTokens(i), nPos)
        For j = 0 To 7
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
'Replaces a substring with another string
Private Function ReplaceSubstring(sIn As String, sFind As String, sReplace As String) As String
  Dim sNew As String, sBefore As String, sAfter As String
  Dim nPos As Integer, nLen As Integer
  
  sNew = sIn
  nPos = InStr(1, sIn, sFind)
  sBefore = "": sAfter = ""
  If nPos > 0 Then
    nLen = Len(sFind)
    If nPos > 1 Then 'grab string before the part we want to replace
      sBefore = Left(sIn, nPos - 1)
    End If
    If nPos + nLen < Len(sIn) Then 'grab string after the part to replace
      sAfter = Mid(sIn, nPos + nLen)
    End If
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

'******************************************************* GetSolution
'Determines result based upon solution of boolean expression
Private Function GetSolution(sIn As String) As Integer
  GetSolution = -1
  If Len(sIn) < 1 Then Exit Function
  
  Dim i As Integer, j As Integer, nCt As Integer, nRet As Integer, nErr As Integer
  Dim sOut As String, sTemp As String
  
  sTemp = sIn
  nCt = Len(sIn)
  
  If de_bug = True Then Text1.Text = Text1.Text & "START:   " & sIn & vbCrLf
  'look for patterns in code line from bp() array and replace
  Do
    sOut = ""
    For j = 0 To MAX_PATTERNS
      nRet = InStr(1, sIn, bp(j).Pattern)
      If nRet > 0 Then
        sOut = ReplaceSubstring(sIn, bp(j).Pattern, bp(j).Replacement)
        sIn = sOut
        If de_bug = True Then Text1.Text = Text1.Text & j & " - " & " : " & bp(j).Pattern & " : " & bp(j).Replacement & "  -  " & sIn & vbCrLf
        Exit For
      End If
    Next j
  Loop Until Len(sIn) = 1
  If de_bug = True Then tmrUpdate.Enabled = False
  GetSolution = Val(sIn)
End Function

'***************************************************** ReadsInput
'reads inputs (check boxes) into g_byIn
' vbUnchecked = 0, vbChecked = 1
Private Sub ReadInputs()
  Dim i As Integer
  For i = 0 To 7
    If chkIn(i).Value = vbChecked Then
      g_uIn(i).Value = chkIn(i).Value
    End If
  Next i
End Sub

'***************************************************** UpdateOutputDisplay
'updates outputs (shapes) by g_byOut
Private Sub UpdateOutputDisplay()
  Dim i As Integer
  
  For i = 0 To 7
    If g_uOut(i).Value = 1 Then
      shpLED(i).FillColor = RGB(0, 255, 0)
    Else
      shpLED(i).FillColor = RGB(0, 155, 0)
    End If
  Next i

End Sub


'********************************************************** CleanUpValidateCode
'cleans up code...spaces before and after certain operators
'this is done for proper display and writing to a file
'Operators:  = equals
'            & AND
'            | OR
Private Function CleanUpValidateCode(sIn As String) As String
  Dim sNew As String, sTemp As String
  Dim i As Integer
  Dim sTokens() As String
  
  'keep all but spaces...they are added in later
  'changes characters to upper case
  'sNew = ""
  'For i = 1 To Len(sIn)
  '  If Mid(sIn, i, 1) <> " " Then sNew = sNew & UCase(Mid(sIn, i, 1))
  'Next i
  'sIn = sNew
  'sNew = ""
  sIn = UCase(RemoveSpaces(sIn))
  
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
         'End
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
