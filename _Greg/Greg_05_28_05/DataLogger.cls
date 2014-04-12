VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_sFilename As String
Private m_bEnabled As Boolean

'read-only
Public Property Get FileName() As String
  FileName = m_sFilename
  m_bEnabled = True
End Property

'writes
'Public Property Let FileName(ByVal vNewValue As Variant)
'  m_sFilename = vNewValue
'End Property

Private Sub Class_Initialize()
  m_sFilename = App.Path & "\datalog.txt"
End Sub

'writes string data to file
Public Function WriteData(sOut As String)
  If m_bEnabled = False Then Exit Function
  Dim nFile As Integer
  nFile = FreeFile
  
  Open m_sFilename For Append As #nFile
    Print #nFile, sOut
  Close #nFile
End Function

'returns file size
Public Function GetDataSize() As Long
  Dim nFile As Integer
  nFile = FreeFile
  
  Open m_sFilename For Append As #nFile
    GetDataSize = LOF(nFile)
  Close #nFile
End Function

'clears data file
Public Function ClearData()
  Dim nFile As Integer
  nFile = FreeFile
  
  Open m_sFilename For Output As #nFile
    
  Close #nFile
  
End Function

'reads
Public Property Get Enable() As Boolean
  Enable = m_bEnabled
End Property

Public Property Let Enable(ByVal vNewValue As Boolean)
  m_bEnabled = vNewValue
End Property

Public Function DateTimeStamp()
  If m_bEnabled = False Then Exit Function
  Dim nFile As Integer
  nFile = FreeFile
  
  Open m_sFilename For Append As #nFile
    Print #nFile, "Date: " & Date & "  Time: " & Time
  Close #nFile

End Function
