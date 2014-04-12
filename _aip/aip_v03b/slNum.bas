Attribute VB_Name = "slNum"
'******************************************************************************************
' slNum.bas - Written by ChuckB (cbolin@dycon.com)
' Super Long (SL) Numbers
' Long binary numbers are manipulated by various logic/math operations but
' are expressed as strings.
'******************************************************************************************
Option Explicit

'******************************************************
'SL_XOR
'******************************************************
Public Function SL_XOR(sNum1 As String, sNum2 As String, sResult As String) As Boolean
  Dim x As Integer
  
  SL_XOR = False: sResult = ""
  If SL_Valid(sNum1, sNum2) = False Then Exit Function
  
  'compare bit by bit
  For x = 1 To Len(sNum1)
    If CInt(Mid(sNum1, x, 1)) Xor CInt(Mid(sNum2, x, 1)) Then
      sResult = sResult & "1"
    Else
      sResult = sResult & "0"
    End If
  Next x
  SL_XOR = True
End Function

'******************************************************
'SL_AND
'******************************************************
Public Function SL_AND(sNum1 As String, sNum2 As String, sResult As String) As Boolean
  Dim x As Integer
  
  SL_AND = False: sResult = ""
  If SL_Valid(sNum1, sNum2) = False Then Exit Function
  
  'compare bit by bit
  For x = 1 To Len(sNum1)
    If CInt(Mid(sNum1, x, 1)) And CInt(Mid(sNum2, x, 1)) Then
      sResult = sResult & "1"
    Else
      sResult = sResult & "0"
    End If
  Next x
  SL_AND = True

End Function

'******************************************************
'SL_OR
'******************************************************
Public Function SL_OR(sNum1 As String, sNum2 As String, sResult As String) As Boolean
  Dim x As Integer
  
  SL_OR = False: sResult = ""
  If SL_Valid(sNum1, sNum2) = False Then Exit Function
  
  'compare bit by bit
  For x = 1 To Len(sNum1)
    If CInt(Mid(sNum1, x, 1)) Or CInt(Mid(sNum2, x, 1)) Then
      sResult = sResult & "1"
    Else
      sResult = sResult & "0"
    End If
  Next x
  SL_OR = True

End Function

'******************************************************
'SL_NAND
'******************************************************
Public Function SL_NAND(sNum1 As String, sNum2 As String, sResult As String) As Boolean
  Dim x As Integer
  
  SL_NAND = False: sResult = ""
  If SL_Valid(sNum1, sNum2) = False Then Exit Function
  
  'compare bit by bit
  For x = 1 To Len(sNum1)
    If CInt(Mid(sNum1, x, 1)) And CInt(Mid(sNum2, x, 1)) Then
      sResult = sResult & "0"
    Else
      sResult = sResult & "1"
    End If
  Next x
  SL_NAND = True

End Function

'******************************************************
'SL_NOR
'******************************************************
Public Function SL_NOR(sNum1 As String, sNum2 As String, sResult As String) As Boolean

  Dim x As Integer
  
  SL_NOR = False: sResult = ""
  If SL_Valid(sNum1, sNum2) = False Then Exit Function
  
  'compare bit by bit
  For x = 1 To Len(sNum1)
    If CInt(Mid(sNum1, x, 1)) Or CInt(Mid(sNum2, x, 1)) Then
      sResult = sResult & "0"
    Else
      sResult = sResult & "1"
    End If
  Next x
  SL_NOR = True

End Function

'******************************************************
'SL_Add
'This is still buggy...got tired of debugging.
'Please help. :-)
'******************************************************
Public Function SL_ADD(sNum1 As String, sNum2 As String, sResult As String) As Boolean
  Dim x, y As Integer
  Dim nCarry As Integer
  Dim nNum1 As Integer
  Dim nNum2 As Integer
  Dim nSum As Integer
    
  SL_ADD = False: sResult = ""
  If SL_Valid(sNum1, sNum2) = False Then Exit Function
  
  'compare bit by bit
  For x = Len(sNum1) To 1 Step -1
    nNum1 = CInt(Mid(sNum1, x, 1))
    nNum2 = CInt(Mid(sNum2, x, 1))
    nSum = nNum1 + nNum2 + nCarry
    
    If nSum = 3 Then
      sResult = "1" & sResult
      nCarry = 1
    ElseIf nSum = 2 Then
      sResult = "0" & sResult
      nCarry = 1
    ElseIf nSum = 1 Then
      sResult = "1" & sResult
      nCarry = 0
    ElseIf nSum = 0 Then
      sResult = "0" & sResult
      nCarry = 0
    End If
  Next x
  SL_ADD = True

End Function

'******************************************************
'SL_GetBit
'******************************************************
Public Function SL_GetBit(sNum1 As String, nBit As Integer) As Integer
  If nBit < 1 Or nBit > Len(sNum1) Then Exit Function
  SL_GetBit = CInt(Mid(sNum1, nBit, 1))
End Function

'******************************************************
'SL_SetBit
'******************************************************
Public Function SL_SetBit(sNum1 As String, nBit As Integer) As Boolean
  SL_SetBit = False
  If nBit < 1 Or nBit > Len(sNum1) Then Exit Function
  If Mid(sNum1, nBit, 1) = "0" Then Mid(sNum1, nBit, 1) = "1"
  SL_SetBit = True
End Function

'******************************************************
'SL_ShiftLeft
'******************************************************
Public Function SL_ShiftLeft(sNum1 As String, nBits As Integer) As Boolean
  Dim x As Integer
  SL_ShiftLeft = False
  If nBits > Len(sNum1) Then Exit Function
  If nBits < 1 Then Exit Function
  sNum1 = Mid(sNum1, nBits + 1) & String(nBits, "0")
  
  SL_ShiftLeft = True

End Function

'******************************************************
'SL_ShiftRight
'******************************************************
Public Function SL_ShiftRight(sNum1 As String, nBits As Integer) As Boolean
  Dim x As Integer
  
  SL_ShiftRight = False
  If nBits > Len(sNum1) Then Exit Function
  If nBits < 1 Then Exit Function
  sNum1 = String(nBits, "0") & Mid(sNum1, 1, Len(sNum1) - nBits)
  SL_ShiftRight = True
End Function

'******************************************************
'SL_RotateLeft
'******************************************************
Public Function SL_RotateLeft(sNum1 As String, nBits As Integer) As Boolean
  Dim x As Integer
  
  SL_RotateLeft = False
  If nBits > Len(sNum1) Then Exit Function
  If nBits < 1 Then Exit Function
  sNum1 = Mid(sNum1, nBits + 1) & Left(sNum1, nBits)
  
  SL_RotateLeft = True

End Function

'******************************************************
'SL_RotateRight
'******************************************************
Public Function SL_RotateRight(sNum1 As String, nBits As Integer) As Boolean
  Dim x As Integer
  
  SL_RotateRight = False
  If nBits > Len(sNum1) Then Exit Function
  If nBits < 1 Then Exit Function
  sNum1 = Right(sNum1, nBits) & Mid(sNum1, 1, Len(sNum1) - nBits)
  SL_RotateRight = True
End Function

'******************************************************
'SL_Valid
'******************************************************
Public Function SL_Valid(sNum1 As String, sNum2 As String) As Boolean
  Dim x As Integer
  
  SL_Valid = False
  
  'lengths must match
  If Len(sNum1) <> Len(sNum2) Then Exit Function

  'must contain either 0 or 1
  For x = 1 To Len(sNum1)
    If Mid(sNum1, x, 1) <> "1" And Mid(sNum1, x, 1) <> "0" Then Exit Function
    If Mid(sNum2, x, 1) <> "1" And Mid(sNum2, x, 1) <> "0" Then Exit Function
  Next x

  SL_Valid = True
End Function
