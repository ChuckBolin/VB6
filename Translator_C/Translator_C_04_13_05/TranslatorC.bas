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
  If nCt > 0 Then
    For j = 1 To nCt
      nComment1 = InStr(1, sInput, "/*")
      If nComment1 > 0 Then nComment2 = InStr(nComment1 + 1, sInput, "*/")
      If nComment2 > nComment1 And nComment1 > 0 Then
        sInput = Left(sInput, nComment1 - 1) & Mid(sInput, nComment2 + 2)
      End If
    Next j
  End If
  
  'strips out comments //
  nComment1 = 0
  For i = 1 To Len(sInput)
    nComment1 = InStr(nComment1 + 1, sInput, "//")
    nCrLf = InStr(nComment1 + 1, sInput, vbCrLf)
    If nComment1 > 0 And nCrLf > 0 Then
      sOut = Left(sInput, nComment1 - 1)
      sOut = sOut & Mid(sInput, nCrLf)
      sInput = sOut
    End If
  Next i
    
  'strips out leading/trailing white spaces
  sLines = Split(sInput, vbCrLf)
  
  If UBound(sLines) < 0 Then
    sOut = "ERROR!" & vbCrLf
    sOut = sOut & "...No code" & vbCrLf
    RemoveComments = sOut
    Exit Function
  End If
  
  sOut = ""
  For i = 0 To UBound(sLines) - 1
    sLines(i) = Trim(sLines(i))
    sOut = sOut & sLines(i) & vbCrLf
  Next i
    
  RemoveComments = sOut
End Function

