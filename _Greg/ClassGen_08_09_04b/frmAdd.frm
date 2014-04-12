VERSION 5.00
Begin VB.Form frmAdd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Property"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3750
   Icon            =   "frmAdd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   3750
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtComment 
      Height          =   855
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   20
      ToolTipText     =   "Add comment here..."
      Top             =   1680
      Width           =   3495
   End
   Begin VB.Frame Frame2 
      Caption         =   "Scope:"
      Height          =   855
      Left            =   120
      TabIndex        =   15
      Top             =   480
      Width           =   2775
      Begin VB.CheckBox chkAbstract 
         Caption         =   "Abstract"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         ToolTipText     =   "Prevents Direct Access to Variables"
         Top             =   480
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.OptionButton optPublic 
         Caption         =   "Public"
         Height          =   255
         Left            =   1200
         TabIndex        =   17
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optPrivate 
         Caption         =   "Private"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.TextBox txtMin 
      Height          =   285
      Left            =   1080
      TabIndex        =   14
      Top             =   3840
      Width           =   1815
   End
   Begin VB.TextBox txtMax 
      Height          =   285
      Left            =   1080
      TabIndex        =   11
      Top             =   3480
      Width           =   1815
   End
   Begin VB.TextBox txtInitial 
      Height          =   285
      Left            =   1080
      TabIndex        =   10
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "Read / Write"
      Height          =   975
      Left            =   120
      TabIndex        =   8
      Top             =   4320
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
      ToolTipText     =   "Variable Type Here"
      Top             =   2640
      Width           =   2055
   End
   Begin VB.TextBox txtProp 
      Height          =   285
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
   Begin VB.CommandButton cmdAccept 
      Caption         =   "&Accept"
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   4920
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "Comment:"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Min Value:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Max Value:"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Initial Value:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label label2 
      Caption         =   "Type:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2640
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

Private m_nPropProc As Integer '0=property, 1=sub, 2=function
Private m_nScope As Integer    '0=Private, 1=Public
Private m_nDataType As Integer '0=integer, 1=single, 2=double,3=boolean,
                               '4=string, 5=byte
Private m_bReadValue As Boolean  'true to read
Private m_bWriteValue As Boolean 'true to write
Private m_bAbstract As Boolean
Private m_sInitialValue As String
Private m_sMinimumValue As String
Private m_sMaximumValue As String
Private m_sName As String
Private m_sComment As String

'selects property data type
Private Sub cboType_Click()
  m_nDataType = cboType.ListIndex
End Sub

'user accepts property information
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
  
  m_sName = txtProp.Text
  If Len(m_sName) < 1 Then Exit Sub  'exit sub if no property name
    
  'add GET property if allowed to read
  If chkRead.Value = vbChecked Then
    m_bReadValue = True 'read
  Else
    m_bReadValue = False
  End If
  
  'add LET property if allowed to write
  If chkWrite.Value = vbChecked Then
    m_bWriteValue = True 'write
  Else
    m_bWriteValue = False
  End If
  
  'scope
  If optPrivate.Value = True Then
    m_nScope = 0 'private
  Else
    m_nScope = 1 'public
  End If
  
  'if true, user can use .X or .Y to read/write.Must use GetX() or SetX()
  If chkAbstract.Value = vbChecked Then
    m_bAbstract = True
  Else
    m_bAbstract = False
  End If
  
  'values
  If Len(txtInitial.Text) > 0 Then m_sInitialValue = txtInitial.Text
  If Len(txtMin.Text) > 0 Then m_sMinimumValue = txtMin.Text
  If Len(txtMax.Text) > 0 Then m_sMaximumValue = txtMax.Text
  
  'comment
  If Len(txtComment.Text) > 0 Then m_sComment = txtComment.Text
      
  'adds to global array for class_content
  g_nIndex = g_nIndex + 1
  ReDim Preserve g_uClassContent(g_nIndex)
  g_uClassContent(g_nIndex).PropProc = m_nPropProc
  g_uClassContent(g_nIndex).Scope = m_nScope
  g_uClassContent(g_nIndex).DataType = m_nDataType
  g_uClassContent(g_nIndex).Comment = m_sComment
  g_uClassContent(g_nIndex).ReadValue = m_bReadValue
  g_uClassContent(g_nIndex).WriteValue = m_bWriteValue
  g_uClassContent(g_nIndex).Name = m_sName
  g_uClassContent(g_nIndex).InitialValue = m_sInitialValue
  g_uClassContent(g_nIndex).MinimumValue = m_sMinimumValue
  g_uClassContent(g_nIndex).MaximumValue = m_sMaximumValue
  g_uClassContent(g_nIndex).Abstract = m_bAbstract
  
  'done
  frmMain.lstProp.AddItem m_sName
  
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
  
  'initialize
  m_nPropProc = 0 'Property by default
  m_nScope = 0 'Private by default
  m_nDataType = 0 'Integer by default
  cboType.ListIndex = m_nDataType
  m_sInitialValue = ""
  m_sMinimumValue = ""
  m_sMaximumValue = ""
  m_sComment = ""
End Sub
