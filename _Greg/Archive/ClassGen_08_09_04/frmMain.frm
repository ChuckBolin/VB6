VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Class Generator v0.3 - Chuck Bolin, August 2004"
   ClientHeight    =   9480
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   9555
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9480
   ScaleWidth      =   9555
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstSub 
      Height          =   1230
      ItemData        =   "frmMain.frx":030A
      Left            =   1680
      List            =   "frmMain.frx":030C
      TabIndex        =   12
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton cmdDeleteSub 
      Caption         =   "Delete Sub/Function"
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton cmdAddSub 
      Caption         =   "Add Sub/Function"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton cmdDeleteProp 
      Caption         =   "Delete Property"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox txtSample 
      Height          =   3495
      Left            =   3480
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   8
      Top             =   120
      Width           =   6015
   End
   Begin VB.ListBox lstProp 
      Height          =   1425
      ItemData        =   "frmMain.frx":030E
      Left            =   1680
      List            =   "frmMain.frx":0310
      TabIndex        =   7
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "&Generate"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear Class Info"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   8280
      Width           =   1695
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy to Clipboard"
      Height          =   375
      Left            =   7800
      TabIndex        =   4
      Top             =   8280
      Width           =   1695
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   7800
      TabIndex        =   3
      Top             =   9000
      Width           =   1695
   End
   Begin VB.TextBox txtCode 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   3720
      Width           =   9375
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add Property"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "&Create Class"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Menu mnuUtil 
      Caption         =   "Utilities"
      Begin VB.Menu mnuUtilCtoOut 
         Caption         =   "Code To Output"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'pops up the Add frmAdd properties box
Private Sub cmdAdd_Click()
  frmAdd.Show
End Sub

Private Sub cmdClear_Click()
  txtCode.Text = ""
  g_sVar = ""
  g_sProp = ""
  g_sHeader = ""
  g_sInit = ""
  
End Sub

'copies contents of txtCode to clipboard
Private Sub cmdCopy_Click()
  Clipboard.Clear
  Clipboard.SetText txtCode.Text
End Sub

'user creates name of class
Private Sub cmdCreate_Click()
  Dim sIn As String
  
  'get name of class
  sIn = InputBox("Enter Class Name (no extensions) : i.e. 'CPerson' ", "Class Name")
  If Len(sIn) < 1 Then sIn = "CGeneralClass"
  If InStr(1, sIn, ".") > 0 Then Exit Sub
  g_uClassHeader.ClassName = sIn
  
  'get name of programmer (author)
  sIn = InputBox("Enter Programmer's  : i.e. 'John Doe' ", "Programmer's Name")
  If Len(sIn) < 1 Then sIn = "Chuck Bolin"
  g_uClassHeader.Author = sIn
  
  'get purpose of program
  sIn = InputBox("Enter Purpose of program: ", "Purpose of Program")
  If Len(sIn) < 1 Then sIn = "General Class...nothing more...nothing less!"
  g_uClassHeader.Purpose = sIn
  
  'get current date
  g_uClassHeader.Date = CStr(Date)
  
  txtSample.Text = ConstructHeader
  
End Sub

'deletes or undeletes items in list box. A delete does not eliminate
'all information, just doesn't put it in text box when Generate is clicked.
'Undelete allows it to be restored
Private Sub cmdDeleteProp_Click()
  Dim i As Integer
  
  For i = 0 To lstProp.ListCount - 1
    If lstProp.Selected(i) = True Then
      If cmdDeleteProp.Caption = "Delete Property" Then
        lstProp.List(i) = g_uClassContent(i + 1).Name & " [Deleted]"
        g_uClassContent(i + 1).Delete = True
        cmdDeleteProp.Caption = "Undelete Property"
      Else
        lstProp.List(i) = g_uClassContent(i + 1).Name
        g_uClassContent(i + 1).Delete = False
        cmdDeleteProp.Caption = "Delete Property"
      End If
    End If
  Next i
End Sub

'Quit program
Private Sub cmdExit_Click()
  Unload Me
End Sub

Private Sub cmdGenerate_Click()
  Dim i As Integer
  If i < 0 Then Exit Sub
  
  txtCode.Text = ""
  
  'get header
  txtCode.Text = ConstructHeader
  
  'get declaration data
  txtCode.Text = txtCode.Text & "'Variable Declaration" & vbCrLf
  For i = 1 To g_nIndex
    If g_uClassContent(i).Delete = False Then
      txtCode.Text = txtCode.Text & frmMain.ConstructDeclaration(i)
    End If
  Next i
  txtCode.Text = txtCode.Text & vbCrLf
  
  'get initialization data
  txtCode.Text = txtCode.Text & "'Initialize Variables" & vbCrLf
  txtCode.Text = txtCode.Text & "Private Sub Class_Initialize()" & vbCrLf
  For i = 1 To g_nIndex
    If g_uClassContent(i).Delete = False Then
      txtCode.Text = txtCode.Text & frmMain.ConstructInitialization(i)
    End If
  Next i
  txtCode.Text = txtCode.Text & "End Sub" & vbCrLf & vbCrLf
    
  'add properties
  For i = 1 To g_nIndex
    If g_uClassContent(i).Delete = False Then
      txtCode.Text = txtCode.Text & frmMain.ConstructProperty(i)
    End If
  Next i
End Sub

'program initialization
Private Sub Form_Load()
  g_sQuote = Chr(34)  'quotation mark
End Sub

'program termination
Private Sub Form_Unload(Cancel As Integer)
  Unload frmAdd
  End
End Sub

'construct class header into a string
Public Function ConstructHeader() As String
  Dim sHeader As String
  
  sHeader = ""
  sHeader = sHeader & "**************************************************************************" & vbCrLf
  sHeader = sHeader & "*  Class Name: " & g_uClassHeader.ClassName & vbCrLf
  sHeader = sHeader & "*  Programmer: " & g_uClassHeader.Author & vbCrLf
  sHeader = sHeader & "*  Date: " & g_uClassHeader.Date & vbCrLf
  sHeader = sHeader & "*  Purpose: " & g_uClassHeader.Purpose & vbCrLf
  sHeader = sHeader & "**************************************************************************" & vbCrLf
  sHeader = sHeader & "Option Explicit" & vbCrLf & vbCrLf
  ConstructHeader = sHeader
End Function

'constructs variable declaration
Public Function ConstructDeclaration(nIndex As Integer) As String
  Dim sOut As String
  Dim sVar As String  'name of private variable with prefix
  Dim sType As String
  Dim sScope As String
  Dim sProp As String 'name of property
  Dim sComment As String
  Dim sInitial As String 'initial value
  sOut = ""
  
  'get variable name
  sProp = g_uClassContent(nIndex).Name
  
  'get data type
  Select Case g_uClassContent(nIndex).DataType
    Case G_INTEGER:
      sType = "Integer"
      sVar = "m_n" & sProp
    Case G_SINGLE:
      sType = "Single"
      sVar = "m_n" & sProp
    Case G_DOUBLE:
      sType = "Double"
      sVar = "m_n" & sProp
    Case G_BOOLEAN:
      sType = "Boolean"
      sVar = "m_b" & sProp
    Case G_BYTE:
      sType = "Byte"
      sVar = "m_by" & sProp
    Case G_STRING:
      sType = "String"
      sVar = "m_s" & sProp
  End Select
  
  'get scope
  Select Case g_uClassContent(nIndex).Scope
    Case G_PRIVATE:
      sScope = "Private"
    Case G_PUBLIC:
      sScope = "Public"
  End Select
  
  sOut = "  " & sScope & " " & sVar & " As " & sType & vbCrLf
  ConstructDeclaration = sOut
End Function

'constructs initialization code into a string
Public Function ConstructInitialization(nIndex As Integer) As String
  Dim sOut As String
  Dim sVar As String  'name of private variable with prefix
  Dim sType As String
  Dim sScope As String
  Dim sProp As String 'name of property
  Dim sComment As String
  Dim sInitial As String 'initial value
  sOut = ""
  
  'get variable name
  sProp = g_uClassContent(nIndex).Name
  
  'get data type
  Select Case g_uClassContent(nIndex).DataType
    Case G_INTEGER:
      sType = "Integer"
      sVar = "m_n" & sProp
    Case G_SINGLE:
      sType = "Single"
      sVar = "m_n" & sProp
    Case G_DOUBLE:
      sType = "Double"
      sVar = "m_n" & sProp
    Case G_BOOLEAN:
      sType = "Boolean"
      sVar = "m_b" & sProp
    Case G_BYTE:
      sType = "Byte"
      sVar = "m_by" & sProp
    Case G_STRING:
      sType = "String"
      sVar = "m_s" & sProp
  End Select
  
  'get scope
  Select Case g_uClassContent(nIndex).Scope
    Case G_PRIVATE:
      sScope = "Private"
    Case G_PUBLIC:
      sScope = "Public"
  End Select
     
  sComment = g_uClassContent(nIndex).Comment
  sInitial = g_uClassContent(nIndex).InitialValue
  
  If Len(sComment) < 1 Then
    sOut = "  " & sVar & " = " & sInitial & vbCrLf
  Else
    sOut = "  " & sVar & " = " & sInitial & " '" & sComment & vbCrLf
  End If
  ConstructInitialization = sOut
End Function

'constructs properties into a string
Public Function ConstructProperty(nIndex As Integer) As String
  Dim sOut As String
  Dim sVar As String  'name of private variable with prefix
  Dim sType As String
  Dim sScope As String
  Dim sProp As String 'name of property
  Dim sArg As String 'name of arg being passed to Let()
  Dim sComment As String
  sOut = ""
  
  'get variable name
  sProp = g_uClassContent(nIndex).Name
  
  'get data type
  Select Case g_uClassContent(nIndex).DataType
    Case G_INTEGER:
      sType = "Integer"
      sVar = "m_n" & sProp
      sArg = "nNewValue"
    Case G_SINGLE:
      sType = "Single"
      sVar = "m_n" & sProp
      sArg = "nNewValue"
    Case G_DOUBLE:
      sType = "Double"
      sVar = "m_n" & sProp
      sArg = "nNewValue"
    Case G_BOOLEAN:
      sType = "Boolean"
      sVar = "m_b" & sProp
      sArg = "bNewValue"
    Case G_BYTE:
      sType = "Byte"
      sVar = "m_by" & sProp
      sArg = "byNewValue"
    Case G_STRING:
      sType = "String"
      sVar = "m_s" & sProp
      sArg = "sNewValue"
  End Select
  
  'get scope
  Select Case g_uClassContent(nIndex).Scope
    Case G_PRIVATE:
      sScope = "Private"
    Case G_PUBLIC:
      sScope = "Public"
  End Select
     
  sComment = g_uClassContent(nIndex).Comment
  
  'abstract property
  If g_uClassContent(nIndex).Abstract = True Then     'abstract
    sOut = sOut & "'" & sComment & vbCrLf
    
    'you can read this value
    If g_uClassContent(nIndex).ReadValue = True Then
      sOut = sOut & sScope & " Function Get" & sProp & "( ) As " & sType & vbCrLf
      sOut = sOut & "   Get" & sProp & " = " & sVar & vbCrLf
      sOut = sOut & "End Function" & vbCrLf & vbCrLf
    End If
    
    'you can change this value
    If g_uClassContent(nIndex).WriteValue = True Then
      sOut = sOut & sScope & " Sub Set" & sProp & "( )" & vbCrLf
      sOut = sOut & "   'do something with " & sVar & " = " & vbCrLf
      sOut = sOut & "End Sub" & vbCrLf & vbCrLf
    End If

  Else  'not abstract
    sOut = sOut & "'" & sComment & vbCrLf
    
    'you can read this
    If g_uClassContent(nIndex).ReadValue = True Then
      sOut = sOut & sScope & " Property Get " & sProp & "( ) As " & sType & vbCrLf
      sOut = sOut & "  " & sProp & " = " & sVar & vbCrLf
      sOut = sOut & "End Property" & vbCrLf & vbCrLf
    End If
    
    'you can write this
    If g_uClassContent(nIndex).WriteValue = True Then
      sOut = sOut & sScope & " Property Let " & sProp & "(ByVal " & sArg & " As " & sType & ")" & vbCrLf
      sOut = sOut & "  " & sVar & " = " & sArg & vbCrLf
      sOut = sOut & "End Property" & vbCrLf & vbCrLf
    End If
  End If
  
  ConstructProperty = sOut
End Function



Private Sub lstProp_Click()
  Dim i As Integer
  
  For i = 0 To lstProp.ListCount - 1
    If lstProp.Selected(i) = True Then
      txtSample.Text = ConstructProperty(i + 1)
      If InStr(1, lstProp.List(i), "[Deleted]") Then 'this is delete
        cmdDeleteProp.Caption = "Undelete Property"
      Else
        cmdDeleteProp.Caption = "Delete Property"
      End If
      
    End If
  Next i
End Sub

Private Sub mnuUtilCtoOut_Click()
  frmCtoOut.Show
End Sub
