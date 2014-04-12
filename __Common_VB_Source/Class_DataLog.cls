VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class_DataLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'======================================================================
' Class_DataLog.cls - Written by Chuck Bolin, June 5, 2005
' Allows for data logging various events in program.
' Version 0.01 - 6/6/2005 - CB - Initial Class setup
'
'
'
'**********************************************************************
' Sample Use of Class
' -------------------
'  Public cBeta As New Class_DataLog 'used for beta testing
'
'  Private Sub Form_Load()
'    cBeta.SetFilePath App.Path     'define file path
'    cBeta.SetFileName "beta.txt"   'define file name
'    cBeta.Enable                   'enable data logging
'    cBeta.LogDateTime              'logs data and time
'  End Sub
'
'  Private Sub Command1_Click()
'    cBeta.Log Text1                'logs content of text1 to file
'  End Sub
'
'  Private Sub Command2_Click()
'    cBeta.ClearFile                'deletes file, starts another
'  End Sub
'
'  Private Sub Command3_Click()
'    cBeta.LogDivider Text1         'bunch these ***************
'  End Sub
'  Private Sub Command4_Click()
'    On Error GoTo MyError
'
'    Err.Raise 70  'used to test error handling
'
'    Exit Sub
'MyError:
'    cBeta.LogError "command4_click"
'  End Sub
'======================================================================
Option Explicit

'private variables
Private m_sFileName As String 'just the filename
Private m_sFilePath As String 'just the path
Private m_sFile As String     'combined path and filename
Private m_bEnable As Boolean  'true means to data log, false no data log
Private m_nMaxFileSize As Long 'sets maximum file size before auto delete

'begins here
Private Sub Class_Initialize()
  m_bEnable = False
  m_sFileName = "" 'nothing
  m_sFilePath = "" 'nothing
  m_sFile = ""     'nothing
  m_nMaxFileSize = 1000000 '1 million bytes
End Sub

'user defined filepath to log file
'*****************************************
Public Sub SetFilePath(sInput As String)
  
  sInput = Trim(sInput)
  If Len(sInput) < 1 Then Exit Sub
  m_bEnable = False
  
  'doesn't mean path is correct
  m_sFilePath = sInput
  
  'path may end in \
  If Right(m_sFilePath, 1) = "\" Then
    m_sFilePath = Left(m_sFilePath, Len(m_sFilePath) - 1)
  End If
  
  'created combine file path name
  m_sFile = m_sFilePath & "\" & m_sFileName
  
End Sub

'user defined filename for log file
'*****************************************
Public Sub SetFileName(sInput As String)
  sInput = Trim(sInput)
  If Len(sInput) < 1 Then Exit Sub
  m_bEnable = False
  
  'doesn't mean filename is correct or not already in use
  m_sFileName = sInput
  m_sFile = m_sFilePath & "\" & m_sFileName
End Sub

'enables use of data log features
'*****************************************
Public Sub Enable()
  If Len(m_sFile) < 1 Then Exit Sub

  m_bEnable = True
End Sub

'diables use of data log features
'*****************************************
Public Sub Disable()
  m_bEnable = False
End Sub

'writes a string
'*****************************************
Public Sub Log(sInput As String)
  Dim nFile As Integer
  
  'verify datalogging enabled
  If m_bEnable = False Then Exit Sub
  
  'verify content
  sInput = Trim(sInput)
  If Len(sInput) < 1 Then Exit Sub
  
  'writes (appends) string to file
  nFile = FreeFile
  Open m_sFile For Append As nFile
    Print #nFile, sInput
  Close nFile

  AutoDeleteFile
End Sub

'writes data time stamp
'*****************************************
Public Sub LogDateTime()
  Dim nFile As Integer
  
  If m_bEnable = False Then Exit Sub
  
  'writes (appends) data and time to file
  nFile = FreeFile
  Open m_sFile For Append As nFile
    Print #1, "=========================================="
    Print #nFile, "  Date: " & Date & "  Time: " & Time
    Print #1, "=========================================="
  Close nFile

  AutoDeleteFile
End Sub

'writes time stamp
'*****************************************
Public Sub LogTime()
  Dim nFile As Integer
  
  If m_bEnable = False Then Exit Sub
  
  'writes (appends) data and time to file
  nFile = FreeFile
  Open m_sFile For Append As nFile
    Print #1, "=========================================="
    Print #nFile, "  Time: " & Time
    Print #1, "=========================================="
  Close nFile

  AutoDeleteFile
End Sub

'writes err.* information and an optional
'string for user
'*****************************************
Public Sub LogError(Optional sInput As String)
  Dim nFile As Integer
  
  If m_bEnable = False Then Exit Sub
  
  'writes error information
  nFile = FreeFile
  Open m_sFile For Append As nFile
    Print #1, "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
    Print #nFile, "  Date: " & Date & "  Time: " & Time
    Print #nFile, "  Error Number: " & Err.Number
    Print #nFile, "  Error Description: " & Err.Description
    If Len(sInput) > 0 Then
      Print #nFile, "  User Message: " & sInput
    End If
    Print #1, "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
  Close nFile

  AutoDeleteFile
End Sub

'sets maximum file size
'*****************************************
Public Sub SetMaxFileSize(nMax As Long)
  If nMax < 20000 Then Exit Sub 'minimum is 20KB
  If nMax > 20000000 Then Exit Sub 'max is 20MB
  m_nMaxFileSize = nMax
End Sub

'auto deletes log file
'*****************************************
Private Sub AutoDeleteFile()
  Dim nLength As Long
  
  nLength = FileLen(m_sFile)
  If nLength > m_nMaxFileSize Then
    Kill (m_sFile)
    LogDateTime
    Log "File auto deleted because it exceeded " & CStr(nLength) & " bytes!"
    Log "New file created!"
  End If
  
End Sub

'user clears data log file
'*****************************************
Public Sub ClearFile()
  
  Kill (m_sFile)
  LogDateTime
  Log "File deleted by user!"
  Log "New file created!"
  
End Sub

'writes a blank line
'*****************************************
Public Sub LogBlankLine(Optional nLines As Integer)
  Dim nFile As Integer
  Dim i As Integer
  
  If m_bEnable = False Then Exit Sub
  
  'writes (appends) data and time to file
  nFile = FreeFile
  Open m_sFile For Append As nFile
    If nLines = 0 Then
      Print #1, ""
    ElseIf nLines > 0 Then
      For i = 1 To nLines
        Print #1, ""
      Next i
    End If
  Close nFile

  AutoDeleteFile
End Sub

'writes a string of 42 characters used to
'divide text in log file
'*****************************************
Public Sub LogDivider(Optional sChar As String)
  Dim nFile As Integer
  
  If m_bEnable = False Then Exit Sub
  
  'writes sChar to file
  nFile = FreeFile
  Open m_sFile For Append As nFile
    
    'default character to print 42x
    If Len(sChar) < 1 Then
      Print #1, "=========================================="
      
    'user defined char
    Else
      Print #1, String(42, Left(sChar, 1))
    End If
  Close nFile

  AutoDeleteFile
End Sub

