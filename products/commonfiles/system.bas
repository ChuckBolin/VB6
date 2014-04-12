Attribute VB_Name = "System"
'********************************************************
'S Y S T E M . B A S - Written by Chuck Bolin
'Purpose:  Provides data logged information reguarding
'errors and normal events.  Used primarily for debugging
'and beta testing software.
'Date:  11/08/99
'Rev A: 2/29/00 - Incorporates additional features.
'Rev B:
'********************************************************
Option Explicit

'required system variables and constants
Public zblnRecordEventOn As Boolean 'this flag allows event recording
Public zstrEventFilename As String 'stores path and file of event log
Const zintFileHandle = 255
Public zstrFlag As String 'stores a value that indicates position in program
Public zvarDummy As Variant 'used for PushEvents when no variable required
Public zstrText As String
Public zstrEventLog As String

'version number and other important information
Public zstrVersion As String 'version of program
Public zstrProgram As String 'name of program

'***********************************************
' S T A R T E V E N T R E C O R D
' Purpose:  Sets required variables for event
' recording.
'***********************************************
Public Sub StartEventRecord()
  zblnRecordEventOn = True
  zstrEventFilename = App.Path & "\eventlog.dat"
  zvarDummy = "Yes"
End Sub

'***********************************************
' C L E A R E V E N T L O G
' Purpose:  Clears Event.dat file
'***********************************************
Public Sub ClearEventLog()
  On Error GoTo ErrorHelpEvent
  Dim lngReturn As Long
  
  'prompt user before deleting
  lngReturn = MsgBox("Are you sure?", vbYesNo, "Clear the Event Log")
  If lngReturn = vbNo Then
    PushEvent "ClearEventLog.... aborted!", Time
    Exit Sub
  End If
  
  'okay to delete
  PushEvent "System - ClearEventLog " & Date, Time
  Kill App.Path & "\" & "eventlog.dat"
  PushEvent "Event Log File Deleted: " & Date, Time
  Exit Sub
  
ErrorHelpEvent:

End Sub

'***********************************************
' R E A D E V E N T L O G
' Purpose:  Clears Event.dat file
'***********************************************
Public Sub ReadEventLog()
  PushEvent "System - ReadEventLog ", Time
  Dim lngReturn As Long
  On Error GoTo ErrorHelpEvent:
  
  'delete file if it is larger than 20Kbytes
  Open App.Path & "\eventlog.dat" For Append As #1
  If LOF(1) > 200000 Then
    Close #1
    Kill App.Path & "\eventlog.dat"
    PushEvent "EVENTLOG.DAT file automatically deleted...", Time
  Else
    Close #1
  End If
    
  lngReturn = Shell(mstrWordpadPath & " " & App.Path & "\eventlog.dat", vbMaximizedFocus)
  Exit Sub
  
ErrorHelpEvent:
  PushError "mnuHelpEvent_Click", ""
  lngReturn = Shell("notepad.exe " & App.Path & "\eventlog.dat", vbMaximizedFocus)
End Sub

'***********************************************
' P U S H E V E N T
' Purpose:  This procedure writes text and a
' variable value to the eventlog.dat file.
'***********************************************
Public Sub PushEvent(strText As String, varValue As Variant)
  If zblnRecordEventOn = False Then Exit Sub
  
  'writes info to file
  Open zstrEventFilename For Append As zintFileHandle
     Print #zintFileHandle, strText & "  " & CStr(varValue)
  Close zintFileHandle
  
  'normal sub termination
  Exit Sub
  
'errors go here
ErrorHandler:
  Close
  Resume Next

End Sub

'***********************************************
' P U S H V A R I A B L E
' Purpose:  This procedure writes text and a
' variable value to the eventlog.dat file.
'***********************************************
Public Sub PushVariable(strText As String, varValue As Variant)
  If zblnRecordEventOn = False Then Exit Sub
  
  'writes info to file
  Open zstrEventFilename For Append As zintFileHandle
     Print #zintFileHandle, strText & "  " & CStr(varValue)
  Close zintFileHandle
  
  'normal sub termination
  Exit Sub
  
'errors go here
ErrorHandler:
  Close
  Resume Next

End Sub

'***********************************************
' P U S H E R R O R
' Purpose:  This procedure writes errors, err.num
' and err.description to eventlod.dat file.
'***********************************************
Public Sub PushError(strText As String, varValue As Variant)
  If zblnRecordEventOn = False Then Exit Sub
  
  'writes info to file
  Open zstrEventFilename For Append As zintFileHandle
    Print #zintFileHandle, " "
    Print #zintFileHandle, "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
    Print #zintFileHandle, "!!!!!!!!!!!!!!!! E R R O R  H A S  O C C U R E D !!!!!!!!!!!!!!!"
    Print #zintFileHandle, "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
    Print #zintFileHandle, "Error Number: " & Err.Number
    Print #zintFileHandle, "Error Description: " & Err.Description
    Print #zintFileHandle, strText & ": " & CStr(varValue)
    Print #zintFileHandle, "Location in code: " & zstrFlag
    Print #zintFileHandle, "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
    Print #zintFileHandle, " "
  Close zintFileHandle
End Sub
