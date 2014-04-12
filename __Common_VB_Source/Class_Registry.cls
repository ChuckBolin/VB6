VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class_Registry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=======================================================================
' Class_Registry.cls - Written by Chuck Bolin, June 6, 2005
' Allows application data to be written/read to/from a registry data
' in the app.path folder.  This IS NOT the Windows Registry file.
' This file is Registry.dat.  It is comprised of keys and data which
' will be saved as a comma delimited line in the file. This class
' provides necessary interface for the Registry.dat file.
' Version 0.01 6/6/05 - CB - Initial class setup.
'
'***********************************************************************
' Sample use of class:
'    Public cReg As New Class_Registry
'
'    Private Sub Command1_Click()
'      Dim bReturn As Boolean
'      bReturn = cReg.AddEntry(Text1, Text2)
'    End Sub
'
'    Private Sub Command2_Click()
'      Text2 = cReg.GetValue(Text1)
'    End Sub
'
'    Private Sub Command3_Click()
'      Dim bReturn As Boolean
'      bReturn = cReg.ModifyEntry(Text1, Text2)
'    End Sub
'
'    Private Sub Command4_Click()
'      Dim bReturn As Boolean
'      bReturn = cReg.DeleteEntry(Text1)
'    End Sub
'=======================================================================
Option Explicit

'stores information
Private Type KEY_DATA
  Key As String
  Data As String
End Type

Private m_uData() As KEY_DATA

'create array with one empty row
'************************************
Private Sub Class_Initialize()
  ReDim m_uData(0)
End Sub

'writes key name and data to file
'************************************
Public Function AddEntry(sKey As String, vValue As Variant) As Boolean
  sKey = Trim(sKey)
  vValue = Trim(vValue)
  AddEntry = False
  
  'exit if invalid key or value
  If Len(sKey) < 1 Or Len(vValue) < 1 Then Exit Function
  
  'add key if it doesnpt exist
  If KeyExists(sKey) = False Then
    SaveToFile sKey & "," & vValue
  End If
  AddEntry = True
End Function

'returns data associated with requested key
'********************************************
Public Function GetValue(sKey As String) As Variant
  Dim nFile As Integer
  Dim sInput As String
  Dim nPos As String 'position of comma
  Dim sLeft As String 'left side of comma
  Dim sRight As String 'right side of comma
  Dim sReturn As String
  
  GetValue = ""
  
  'exit if no key entered
  sKey = Trim(sKey)
  If Len(sKey) < 1 Then Exit Function
  
  'read through file and look for requested key
  nFile = FreeFile
  sReturn = Dir(App.Path & "\registry.dat")
  If Len(sReturn) < 1 Then Exit Function
  
  Open App.Path & "\registry.dat" For Input As nFile
    Do
      Line Input #nFile, sInput
      If Len(sInput) > 0 Then
        nPos = InStr(1, sInput, ",")
        If nPos > 0 Then
          sLeft = Trim(Left(sInput, nPos - 1))
          sRight = Trim(Mid(sInput, nPos + 1))
          
          'there is a match for this key
          If LCase(sKey) = LCase(sLeft) Then
            GetValue = CVar(sRight)
            Close #nFile
            Exit Function
          End If
        End If
      End If
    Loop Until EOF(nFile)
  Close #nFile
  
End Function

'modifies existing key
'********************************
Public Function ModifyEntry(sKey As String, vData As Variant) As Boolean
  Dim nFile As Integer
  Dim sInput As String
  Dim nPos As String 'position of comma
  Dim sLeft As String 'left side of comma
  Dim sRight As String 'right side of comma
  Dim sReturn As String
  Dim sNew As String
  Dim bFound As Boolean
  
  ModifyEntry = False
  bFound = False
  If KeyExists(sKey) = False Then Exit Function
  
  'exit if no key entered
  sKey = Trim(sKey)
  If Len(sKey) < 1 Then Exit Function
  
  'read through file and look for requested key
  nFile = FreeFile
  sReturn = Dir(App.Path & "\registry.dat")
  If Len(sReturn) < 1 Then Exit Function
  
  'read each line and save to sNew, ignore key of choice
  Open App.Path & "\registry.dat" For Input As nFile
    Do
      Line Input #nFile, sInput
      If Len(sInput) > 0 Then
        nPos = InStr(1, sInput, ",")
        If nPos > 0 Then
          sLeft = Trim(Left(sInput, nPos - 1))
          sRight = Trim(Mid(sInput, nPos + 1))
          
          'there is a match for this key
          If LCase(sKey) = LCase(sLeft) Then
            bFound = True
            sNew = sNew & sKey & "," & vData & vbCrLf
          Else
            sNew = sNew & sInput & vbCrLf
          End If
        End If
      End If
    Loop Until EOF(nFile)
  Close #nFile
  
  If bFound = True Then
    nFile = FreeFile
    Open App.Path & "\registry.dat" For Output As nFile
      Print #nFile, sNew
    Close nFile
  End If
End Function

'deletes key value
'***********************************
Public Function DeleteEntry(sKey As String) As Boolean
  Dim nFile As Integer
  Dim sInput As String
  Dim nPos As String 'position of comma
  Dim sLeft As String 'left side of comma
  Dim sRight As String 'right side of comma
  Dim sReturn As String
  Dim sNew As String
  Dim bFound As Boolean
  
  DeleteEntry = False
  bFound = False
  If KeyExists(sKey) = False Then Exit Function
  
  'exit if no key entered
  sKey = Trim(sKey)
  If Len(sKey) < 1 Then Exit Function
  
  'read through file and look for requested key
  nFile = FreeFile
  sReturn = Dir(App.Path & "\registry.dat")
  If Len(sReturn) < 1 Then Exit Function
  
  'read each line and save to sNew, ignore key of choice
  Open App.Path & "\registry.dat" For Input As nFile
    Do
      Line Input #nFile, sInput
      If Len(sInput) > 0 Then
        nPos = InStr(1, sInput, ",")
        If nPos > 0 Then
          sLeft = Trim(Left(sInput, nPos - 1))
          sRight = Trim(Mid(sInput, nPos + 1))
          
          'there is a match for this key
          If LCase(sKey) = LCase(sLeft) Then
            bFound = True
            'do nothing else
          Else
            sNew = sNew & sInput & vbCrLf
          End If
        End If
      End If
    Loop Until EOF(nFile)
  Close #nFile
  
  If bFound = True Then
    nFile = FreeFile
    Open App.Path & "\registry.dat" For Output As nFile
      Print #nFile, sNew
    Close nFile
  End If
End Function


'writes string to file
'************************************
Private Sub SaveToFile(sInput)
  Dim nFile As Integer
  
  sInput = Trim(sInput)
  
  'exit if nothing to write or missing a comma
  If Len(sInput) < 1 Then Exit Sub
  If InStr(1, sInput, ",") < 1 Then Exit Sub
  
  nFile = FreeFile
  Open App.Path & "\registry.dat" For Append As #nFile
    Print #nFile, sInput
  Close #nFile

End Sub

'returns true if key already exists
Private Function KeyExists(sKey As String) As Boolean
  Dim nFile As Integer
  Dim sInput As String
  Dim nPos As String 'position of comma
  Dim sLeft As String 'left side of comma
  Dim sRight As String 'right side of comma
  Dim sReturn As String
  
  KeyExists = False
  
  'exit if no key entered
  sKey = Trim(sKey)
  If Len(sKey) < 1 Then Exit Function
  
  'read through file and look for requested key
  nFile = FreeFile
  sReturn = Dir(App.Path & "\registry.dat")
  If Len(sReturn) < 1 Then Exit Function
  
  Open App.Path & "\registry.dat" For Input As nFile
    Do
      Line Input #nFile, sInput
      If Len(sInput) > 0 Then
        nPos = InStr(1, sInput, ",")
        If nPos > 0 Then
          sLeft = Trim(Left(sInput, nPos - 1))
          sRight = Trim(Mid(sInput, nPos + 1))
          
          'there is a match for this key
          If LCase(sKey) = LCase(sLeft) Then
            KeyExists = True
            Close #nFile
            Exit Function
          End If
        End If
      End If
    Loop Until EOF(nFile)
  Close #nFile
End Function

