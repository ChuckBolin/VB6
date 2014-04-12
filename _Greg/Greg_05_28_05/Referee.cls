VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Referee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type PENALTY_DESCRIPTION
  Name As String
  Condition As String
  IfTrue As String
  IfFalse As String
  Delay As String
End Type

Private pen(100) As PENALTY_DESCRIPTION

'************************************
'Loads script file for penalties
Public Function LoadPenaltyFile() As Boolean
  Dim sFilename As String
  Dim i As Integer
  Dim nFile As Integer 'file handle
  Dim sInput As String
  Dim nCount As Integer ' counts number of penalities
  Dim sPenalty() As String 'store each line parsed into 5 strings
  
  LoadPenaltyFile = False
  
  'verify correctness of filepath
  sFilename = App.Path & "\penalty.txt"
  If Len(dir(sFilename)) < 1 Then Exit Function
  
  'read file and load pen() array
  nFile = FreeFile
  nCount = 0 'start at index 0
  Open sFilename For Input As #nFile
    Do
      Line Input #nFile, sInput
      If Len(Trim(sInput)) > 0 And InStr(1, sInput, ",") > 0 Then
        sPenalty = Split(sInput, ",")
        pen(nCount).Name = sPenalty(0)
        pen(nCount).Condition = sPenalty(1)
        pen(nCount).IfTrue = sPenalty(2)
        pen(nCount).IfFalse = sPenalty(3)
        pen(nCount).Delay = sPenalty(4)
        nCount = nCount + 1
      End If
    Loop Until EOF(nFile)
  Close #nFile
  nCount = nCount - 1
  
  'display info in array pen()
  For i = 0 To nCount
   MsgBox pen(i).Name & vbCrLf & pen(i).Condition & vbCrLf & pen(i).IfTrue & vbCrLf & pen(i).IfFalse & vbCrLf & pen(i).Delay
   
  Next i
  
  'must be at least one penalty in file
  If nCount < 0 Then Exit Function
    
  LoadPenaltyFile = True
End Function

'************************************
Private Sub Class_Initialize()
  'Dim bRet As Boolean
  
  'bRet = LoadPenaltyFile()
  'MsgBox bRet
  'If bRet = False Then
  '  MsgBox "Unable to load penalty.txt file" & vbCrLf & "Terminating program"
  '  End
  'End If
End Sub
