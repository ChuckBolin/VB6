Attribute VB_Name = "TranslatorC"
Option Explicit

'******************************************
' Removes all comments and leading/lagging
' whitespaces.
'******************************************
Public Function RemoveComments(sInput As String) As String
 
  Dim nComment1 As Integer  'tracks position of comments /* and //
  Dim nComment2 As Integer  'tracks position of comment */
  Dim nCrLf As Integer        'position of vbCrLf
  Dim i, j As Integer
  Dim nCt As Integer  'counts /*
  Dim nCt2 As Integer 'counts */
  Dim sOut As String
  Dim sLines() As String
  Dim nLC As Integer
  Dim nBegin As Integer
      
  'count open comments /*
  For i = 1 To Len(sInput)
    If Mid(sInput, i, 2) = "/*" Then
      nCt = nCt + 1
    ElseIf Mid(sInput, i, 2) = "*/" Then
      nCt2 = nCt2 + 1
    End If
  Next i
  
  If nCt <> nCt2 Then  'error with comments /* and */
    sOut = "ERROR!" & vbCrLf
    sOut = sOut & "...Unequal number of /* and */. Verify comments" & vbCrLf
    RemoveComments = sOut
    Exit Function
  End If
    
  'strips out comments /* and */
  nComment1 = 0: nComment2 = 0: sOut = "": nBegin = 1
  If nCt > 0 Then  'comments have been found in sInput
    For i = 1 To Len(sInput)
      If Mid(sInput, i, 2) = "/*" Then
        nLC = nLC + 1
        If nLC = 1 Then
          nComment1 = i 'opening comment found
          If nComment1 > nBegin Then
            sOut = sOut & Mid(sInput, nBegin, nComment1 - nBegin - 1)
          End If
        End If
      ElseIf Mid(sInput, i, 2) = "*/" Then
        nLC = nLC - 1
        If nLC = 0 Then   'closing comment found
          nComment2 = i
          'sOut = sOut & Mid(sInput, nComment2 + 2)
          nBegin = nComment2
        ElseIf nLC < 1 Then
          sOut = "ERROR!" & vbCrLf
          sOut = sOut & ".../* and */ out of order. Verify comments" & vbCrLf
          RemoveComments = sOut
        End If
      End If
    Next i
  Else
    sOut = sInput
  End If
  
  MsgBox sOut
  If Len(sOut) > 0 Then sInput = sOut  'keep modified code
  'RemoveComments = sOut
  'Exit Function
    
  'strips out comments /* and */
  'If nCt > 0 Then
  '  For j = 1 To nCt
  '    nComment1 = InStr(1, sInput, "/*")
  '    If nComment1 > 0 Then nComment2 = InStr(nComment1 + 1, sInput, "*/")
  '    If nComment2 > nComment1 And nComment1 > 0 Then
  '      sInput = Left(sInput, nComment1 - 1) & Mid(sInput, nComment2 + 2)
  '    End If
  '  Next j
  'End If
  
  'verifies content to code after removal of /*...*/ comments
  sLines = Split(sInput, vbCrLf)
  If UBound(sLines) < 0 Then
    sOut = "ERROR!" & vbCrLf
    sOut = sOut & "...No code after any /*...*/ comments have been removed" & vbCrLf
    RemoveComments = sOut
    Exit Function
  End If
  
  'strips out comments //
  sOut = ""
  For i = 0 To UBound(sLines) - 1
    j = InStr(1, sLines(i), "//")
    If j > 0 Then
      sLines(i) = Left(sLines(i), j - 1)
    End If
    sLines(i) = Trim(sLines(i))  'delete leading/trailing comments
    sOut = sOut & sLines(i) & vbCrLf 'rebuild string w/o // comments
  Next i
  MsgBox sOut
  'nComment1 = 0
  'For i = 1 To Len(sInput)
  '  nComment1 = InStr(nComment1 + 1, sInput, "//")
  '  nCrLf = InStr(nComment1 + 1, sInput, vbCrLf)
   ' If nComment1 > 0 And nCrLf > 0 Then
   '   sOut = Left(sInput, nComment1 - 1)
   '   sOut = sOut & Mid(sInput, nCrLf)
   '   sInput = sOut
   ' End If
  'Next i
    
  'strips out leading/trailing white spaces
  'sLines = Split(sInput, vbCrLf)
  
  'If UBound(sLines) < 0 Then
  
  'sOut = ""
  'For i = 0 To UBound(sLines) - 1
   ' sLines(i) = Trim(sLines(i))
   ' sOut = sOut & sLines(i) & vbCrLf
  'Next i
    
  RemoveComments = sOut
End Function

