VERSION 5.00
Begin VB.UserControl ErrorControl 
   BackColor       =   &H008080FF&
   ClientHeight    =   540
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   585
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   540
   ScaleWidth      =   585
   ToolboxBitmap   =   "ErrorControl.ctx":0000
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ERR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "ErrorControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***********************************************************
' ErrorControl - Written by Chuck Bolin
' Date: 3.24.03
' Control receives filename and text.  Any error can
' be processed by this control. Information is written to
' file called 'filename'. Text is written to file.
' Also, this control maintains a data log.
' Properties:
'   .Text         'stores any string
'   .Filename     'stores filename to write data to
' Methods:
'   .LogData      'write a string to file
'   .ProcessError 'writes err number and description to file
'   .StartProgram 'writes program start info to file
'   .EndProgram   'writes program end info to file
'***********************************************************
Option Explicit

'variables
Private mstrFilename As String
Private mstrText As String
Private mlngErrorNum As Integer
Private mstrErrorDescription As String

'FILENAME property
Public Property Get Filename() As String
  Filename = mstrFilename
End Property

Public Property Let Filename(ByVal vNewValue As String)
  mstrFilename = vNewValue
End Property

'TEXT property
Public Property Get Text() As String
  Text = mstrText
End Property

Public Property Let Text(ByVal vNewValue As String)
  mstrText = vNewValue
End Property

'handles error message, receives error number and description
Public Sub ProcessError(lngError As Long, strDescription As String)
  Dim nFile As Integer
  Dim strOut As String
  
  'formatting of error message
  strOut = ""
  strOut = strOut & "**************************************************" & vbCrLf
  strOut = strOut & "                  E R R O R ! ! !" & vbCrLf
  strOut = strOut & " If this problem reoccurs, record the following  " & vbCrLf
  strOut = strOut & " information and give it to the programmer. " & vbCrLf
  strOut = strOut & "**************************************************" & vbCrLf
  strOut = strOut & "Date: " & Date & " Time: " & Time & vbCrLf
  strOut = strOut & "In code: " & mstrText & vbCrLf
  strOut = strOut & "Error No.: " & lngError & vbCrLf
  strOut = strOut & "Error Description: " & strDescription & vbCrLf
  strOut = strOut & "**************************************************" & vbCrLf
  strOut = strOut & "**************************************************" & vbCrLf
  
  nFile = FreeFile
  Open mstrFilename For Append As nFile
    Print #nFile, strOut
  Close #nFile
  
  MsgBox strOut
  
End Sub

'used for datalogging
Public Sub LogData(strText As String)
  Dim nFile As Integer
  nFile = FreeFile
  Open mstrFilename For Append As nFile
    Print #nFile, strText
  Close #nFile
End Sub

'prints to data file
Public Sub StartProgram()
  Dim nFile As Integer
  nFile = FreeFile
  Open mstrFilename For Append As nFile
    Print #nFile, ""
    Print #nFile, "**************************************************"
    Print #nFile, "Start of Program: " & Date & ", " & Time
    Print #nFile, "**************************************************"
    Print #nFile, ""
  Close #nFile
End Sub

'
Public Sub EndProgram()
  Dim nFile As Integer
  nFile = FreeFile
  Open mstrFilename For Append As nFile
    Print #nFile, ""
    Print #nFile, "**************************************************"
    Print #nFile, "End of Program: " & Date & ", " & Time
    Print #nFile, "**************************************************"
    Print #nFile, ""
  Close #nFile
  LogData "Log file cleared " & Date & ", " & Time
End Sub


Private Sub UserControl_Initialize()
  On Error GoTo MyError
  
  'default log file
  mstrFilename = App.Path & "\logfile.txt"
  Exit Sub
MyError:
  MsgBox App.Path & "   " & Err.Description & vbCrLf & _
  App.Path & "\logfile.txt  does not exit!" & vbCrLf & _
  "Verify correct file path and filename. Try again!"
  Exit Sub
End Sub

Public Sub ClearLog()
  Dim nFile As Integer
  nFile = FreeFile
  Open mstrFilename For Output As nFile
  Close nFile
End Sub

'*********************************************************************
'Following is an example showing how to use this error control
'*********************************************************************
'Private Sub CreateError_Click()
'  Dim x As Integer
'  On Error GoTo myerror
'    e.LogData "Command1_Click"
'    e.Text = "100"
    
'    For x = 1 To 100
'     Err.Raise CLng(x)
'    Next x
    
'    e.Text = "110"
'    e.LogData e.Text
    
'  Exit Sub
  
'myerror:
 ' e.ProcessError Err.Number, Err.Description
 ' Resume Next
'End Sub

'Private Sub Clear_Click()
'  e.ClearLog
'End Sub

'Private Sub Form_Load()
'  e.Filename = App.Path & "\log.txt"
'  e.StartProgram
'End Sub

'Private Sub Form_Unload(Cancel As Integer)
'  e.EndProgram
'End Sub



