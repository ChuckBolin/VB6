Attribute VB_Name = "Faults"
'*********************************************************************
' FAULTS.BAS
' Describes all faults
'*********************************************************************
Option Explicit

Public Const MAX_NUMBER_FAULTS = 10

Public m_nNumberOfFaults As Integer

'*************************************************** GetNumberOfFaults
Public Function GetNumberOfFaults() As Integer
  GetNumberOfFaults = m_nNumberOfFaults
End Function

'*************************************************** ClearFault
'0, clears all
Public Sub ClearFault(nVal As Integer)
  Dim i As Integer
  Dim nFault As Integer
  
  If nFault < 1 Then Exit Sub
  If nFault > MAX_ELECTRICAL_COMPONENTS Then Exit Sub
  
  'clear all faults
  If nVal = 0 Then
    For i = 1 To MAX_ELECTRICAL_COMPONENTS
      f(i) = False
    Next i
    m_nNumberOfFaults = 0
  'clears only requested fault
  Else
    f(nVal) = False 'clears fault
    m_nNumberOfFaults = m_nNumberOfFaults - 1
  End If

End Sub


'*************************************************** AddFault
Public Sub AddFault()
  Dim nFault As Integer
  
PickAnother:
  nFault = (Rnd * MAX_ELECTRICAL_COMPONENTS) Mod 23 'MAX_ELECTRICAL_COMPONENTS
  
  'boundary control
  If nFault < 1 Then nFault = 1
  If nFault > MAX_ELECTRICAL_COMPONENTS Then nFault = MAX_ELECTRICAL_COMPONENTS
  
  'these are unmeaningful faults
  If nFault = 1 Or nFault = 5 Or nFault = 9 Or nFault = 13 Then GoTo PickAnother
  If nFault = 19 Or nFault = 20 Then GoTo PickAnother
  
  'don't pick fault if already exists
  If f(nFault) = True Then GoTo PickAnother
  
  'MsgBox nFault
  f(nFault) = True
  m_nNumberOfFaults = m_nNumberOfFaults + 1
  'MsgBox "Fault No.: " & nFault & vbCrLf & "Description: " & GetFaultDescription(nFault)
End Sub

'****************************************************** GetFaultDescription
'Returns a description of the fault
Public Function GetFaultDescription(nFault As Integer) As String
  GetFaultDescription = ""
  Dim sOut As String
  
  Select Case nFault
    Case 2
      sOut = "Main Disconnect Q0, contacts 1 and 2 are open."
    Case 3
      sOut = "Main Disconnect Q0, contacts 3 and 4 are open."
    Case 4
      sOut = "Main Disconnect Q0, contacts 5 and 6 are open."
    Case 6
      sOut = "Main Fuse F1_A is open."
    Case 7
      sOut = "Main Fuse F1_B is open."
    Case 8
      sOut = "Main Fuse F1_C is open."
    Case 10
      sOut = "Motor #1 Fuse F3_A is open."
    Case 11
      sOut = "Motor #1 Fuse F3_B is open."
    Case 12
      sOut = "Motor #1 Fuse F3_C is open."
    Case 14
      sOut = "Motor #2 Fuse F4_A is open."
    Case 15
      sOut = "Motor #2 Fuse F4_B is open."
    Case 16
      sOut = "Motor #2 Fuse F4_C is open."
    Case 17
      sOut = "Transformer Fuse F2_A is open."
    Case 18
      sOut = "Transformer Fuse F2_B is open."
    Case 21
      sOut = "Transformer H1 to H2 coil is open."
    Case 22
      sOut = "Transformer H2 to H3 jumper is open."
    Case 23
      sOut = "Transformer H3 to H4 coil is open."
    
      
  End Select
  GetFaultDescription = sOut
  
End Function

