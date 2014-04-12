Attribute VB_Name = "Language"
Option Explicit

Private Type INPUT_PATTERN
   Pattern As String
   Response As String
End Type

'stores question and statement patterns along with response
Private totalPatterns As Long
Private lang() As INPUT_PATTERN

'initialization routine
Public Sub loadHumanInputVariables()
  Dim count As Long
  Dim sInput As String
  
  If Dir(App.Path & "\data\data1.dat") = "" Then Exit Sub
  
  Open App.Path & "\data\data1.dat" For Input As #1
    Do
      Line Input #1, sInput
      sInput = Trim(sInput)
      If LCase(Left(sInput, 2)) = "p:" Then
        totalPatterns = totalPatterns + 1
        ReDim Preserve lang(totalPatterns)
        lang(totalPatterns).Pattern = LCase(LTrim(Mid(sInput, 3)))
      ElseIf LCase(Left(sInput, 2)) = "r:" Then
        lang(totalPatterns).Response = LTrim(Mid(sInput, 3))
      End If
    Loop Until EOF(1)
  Close #1
  
  
End Sub

'********************************************************
' The user input is evaluated for some keywords. Based
' upon this a user response is generated
'********************************************************
Public Function processHumanInput(sInput As String) As String
  Dim i As Long
  
  processHumanInput = "Sorry, I don't understand."
  
  For i = 1 To totalPatterns
    If findPatternMatch(i, sInput) = True Then
      processHumanInput = lang(i).Response
      Exit Function
    End If
  Next i
  
End Function

'********************************************************
' Returns true if a pattern is found that matches.
'********************************************************
Private Function findPatternMatch(num As Long, s As String) As Boolean
  Dim pat() As String
  Dim i As Integer
  Dim nPos As Integer
  Dim nLast As Integer
  
  findPatternMatch = False
  
  If num < 1 Then Exit Function
  If num > totalPatterns Then Exit Function
  
  'string 's' must include all pattern words and be in
  'the correct sequence
  s = LCase(s)
  pat() = Split(lang(num).Pattern, ",")
  For i = 0 To UBound(pat) - 1
    pat(i) = LCase(Trim(pat(i))) 'get rid of leading/trailing spaces
  Next i
  
  'add error checking to array
  nLast = 1
  For i = 0 To UBound(pat) - 1
    nPos = InStr(nLast, s, pat(i))
    If nPos < 1 Then Exit Function
    nLast = nPos + 1
  Next i
  
  'MsgBox s & vbCrLf & lang(num).Pattern
  'alright...these items exists
  findPatternMatch = True
End Function



'********************************************************
' Patterns and responses are loaded here.
'********************************************************
Private Sub loadHumanInputPatterns()
  lang(1).Pattern = "what, your, last, name"
  lang(1).Response = "My last name is " & clepProfile.getLastName & "."
  lang(2).Pattern = "what, your, first, name"
  lang(2).Response = "My name is " & clepProfile.getFirstName & "."
  lang(3).Pattern = "what, your, name"
  lang(3).Response = "My name is " & clepProfile.getFirstName & "."
  lang(4).Pattern = "when, you, born"
  lang(4).Response = "I was born on " & CStr(clepProfile.getBirthMonth) & "/" & CStr(clepProfile.getBirthDay) & "/" & CStr(clepProfile.getBirthYear) & "."
  lang(5).Pattern = "when, your, birthday"
  lang(5).Response = "I was born on " & CStr(clepProfile.getBirthMonth) & "/" & CStr(clepProfile.getBirthDay) & "/" & CStr(clepProfile.getBirthYear) & "."
  lang(6).Pattern = "how, old, you"
  lang(6).Response = "I am " & CStr(Year(Date) - clepProfile.getBirthYear) & " years old."
  lang(7).Pattern = "who, are, you"
  lang(7).Response = "I am CLEP, an artificial intelligence program."
  lang(8).Pattern = "what,language, you, written"
  lang(8).Response = "I am written in Visual Basic 6.0."
  lang(9).Pattern = "what, are, you"
  lang(9).Response = "I am a computer program."
  
  
  
End Sub

