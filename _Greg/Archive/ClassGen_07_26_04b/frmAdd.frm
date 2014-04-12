VERSION 5.00
Begin VB.Form frmAdd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Property"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   3015
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtMin 
      Height          =   285
      Left            =   1080
      TabIndex        =   14
      Top             =   1680
      Width           =   1815
   End
   Begin VB.TextBox txtMax 
      Height          =   285
      Left            =   1080
      TabIndex        =   11
      Top             =   1320
      Width           =   1815
   End
   Begin VB.TextBox txtInitial 
      Height          =   285
      Left            =   1080
      TabIndex        =   10
      Top             =   960
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "Read / Write"
      Height          =   975
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   1695
      Begin VB.CheckBox chkWrite 
         Caption         =   "Write"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CheckBox chkRead 
         Caption         =   "Read"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Value           =   1  'Checked
         Width           =   1335
      End
   End
   Begin VB.ComboBox cboType 
      Height          =   315
      Left            =   840
      TabIndex        =   1
      Text            =   "(Select Type)"
      Top             =   480
      Width           =   2055
   End
   Begin VB.TextBox txtProp 
      Height          =   285
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton cmdAccept 
      Caption         =   "&Accept"
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Min Value:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Max Value:"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Initial Value:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Width           =   855
   End
   Begin VB.Label label2 
      Caption         =   "Type:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Property:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAccept_Click()
  Dim sProp As String         'the complete function
  Dim sPropName As String     'user selected name of property
  Dim sPropVar As String      'program created variable
  Dim sType As String         'variable type
  Dim sArg As String          'name of argument for LET property
  Dim sInit As String         'initial value
  Dim sMax As String          'max value
  Dim sMin As String          'min value
  Dim bMax As Boolean         'true means add this
  Dim bMin As Boolean         'true means add this
  
  'If cboType.Text = "(Select Type)" Then Exit Sub
  'If cboType = 0 Then Exit Sub
  
  sPropName = txtProp.Text
  sType = cboType.Text
  
  'construct variable name by prefixing 1 letter n, s, b, v
  If sType = "Integer" Then
    sPropVar = "n" & sPropName
    sArg = "nNewValue"
    If IsNumeric(txtInitial.Text) Then
      If Len(txtInitial.Text) > 0 Then sInit = txtInitial.Text
    Else
      sInit = "0"
    End If
    If IsNumeric(txtMax.Text) And Len(txtMax.Text) > 0 Then
      sMax = txtMax.Text
      bMax = True
    End If
    If IsNumeric(txtMin.Text) And Len(txtMin.Text) > 0 Then
      sMin = txtMin.Text
      bMin = True
    End If
  
  
  ElseIf sType = "Single" Then
    sPropVar = "n" & sPropName
    sArg = "nNewValue"
    If IsNumeric(txtInitial.Text) Then
      If Len(txtInitial.Text) > 0 Then sInit = txtInitial.Text
    Else
      sInit = "0"
    End If
    If IsNumeric(txtMax.Text) And Len(txtMax.Text) > 0 Then
      sMax = txtMax.Text
      bMax = True
    End If
    If IsNumeric(txtMin.Text) And Len(txtMin.Text) > 0 Then
      sMin = txtMin.Text
      bMin = True
    End If

  ElseIf sType = "Double" Then
    sPropVar = "n" & sPropName
    sArg = "nNewValue"
    If IsNumeric(txtInitial.Text) Then
      If Len(txtInitial.Text) > 0 Then sInit = txtInitial.Text
    Else
      sInit = "0"
    End If
    If IsNumeric(txtMax.Text) And Len(txtMax.Text) > 0 Then
      sMax = txtMax.Text
      bMax = True
    End If
    If IsNumeric(txtMin.Text) And Len(txtMin.Text) > 0 Then
      sMin = txtMin.Text
      bMin = True
    End If

  ElseIf sType = "String" Then
    sPropVar = "s" & sPropName
    sArg = "sNewValue"
    If Len(txtInitial.Text) > 0 Then
      sInit = txtInitial.Text
    Else
      sInit = g_sQuote & " " & g_sQuote
    End If
    
  ElseIf sType = "Boolean" Then
    sPropVar = "b" & sPropName
    sArg = "bNewValue"
    If LCase(txtInitial.Text) = "true" Or LCase(txtInitial.Text) = "false" Then
      If Len(txtInitial.Text) > 0 Then sInit = txtInitial.Text
    Else
      sInit = "False"
    End If
    
  Else
    sPropVar = "v" & sPropName
    sArg = "vNewValue"
    If Len(txtInitial.Text) > 0 Then
      sInit = txtInitial.Text
    Else
      sInit = g_sQuote & " " & g_sQuote
    End If
  End If
  
  'add GET property if allowed to read
  If chkRead.Value = vbChecked Then
    sProp = sProp & "Public Property Get " & sPropName & " ( ) As " & sType & vbCrLf
    sProp = sProp & "  " & sPropName & " = " & sPropVar & vbCrLf
    sProp = sProp & "End Property" & vbCrLf & vbCrLf
  End If
  
  'add LET property if allowed to write
  If chkWrite.Value = vbChecked Then
    sProp = sProp & "Public Property Let " & sPropName & " (ByVal " & sArg & " As " & sType & ")" & vbCrLf
    If bMax = True Then
      sProp = sProp & "  If " & sArg & " > " & sMax & " Then Exit Property" & vbCrLf
    End If
    If bMin = True Then
      sProp = sProp & "  If " & sArg & " < " & sMin & " Then Exit Property" & vbCrLf
    End If
    
    sProp = sProp & "  " & sPropVar & " = " & sArg & vbCrLf
    
    sProp = sProp & "End Property" & vbCrLf & vbCrLf
  End If
  
  'process variable
  g_sVar = g_sVar & "Private " & sPropVar & " As " & sType & vbCrLf
  
  'updates text box on frmMain
  sInit = "  " & sPropVar & " = " & sInit
  g_sInit = g_sInit & sInit & vbCrLf
  g_sProp = g_sProp & sProp
  frmMain.UpdateCode
  Unload Me
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  cboType.AddItem "Integer"
  cboType.AddItem "Single"
  cboType.AddItem "Double"
  cboType.AddItem "Boolean"
  cboType.AddItem "String"
End Sub
