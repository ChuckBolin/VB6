Attribute VB_Name = "AIP"
Option Explicit

Public rules As String 'really, really big string

Public Sub parseQuestion(s As String)
  Dim nPos As Integer
  
  If InStr(1, LCase(s), "<r01>") Then
    nPos = InStr(1, LCase(s), "<r01>")
    frmMain.txtResponse.Text = getRule("<r01")
  End If
End Sub

Public Sub loadRules()
  Dim sInput As String
  
  rules = ""
  
  Open App.Path & "\data\rules.txt" For Input As 1
    Do
      Line Input #1, sInput
      rules = rules & sInput & vbCrLf
    Loop Until EOF(1)
  Close #1

End Sub



Private Function getRule(sRule) As String
  Dim nBegin As Long
  Dim nEnd As Long
  
  nBegin = InStr(1, LCase(rules), LCase(sRule))
  nEnd = InStr(nBegin, LCase(rules), "</r")
  getRule = Mid(rules, nBegin, nEnd - nBegin)

End Function
