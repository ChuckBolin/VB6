Attribute VB_Name = "Global"
Option Explicit

'constants used to ID reasons for filtering
Public Const FILTER_OK = 0
Public Const FILTER_TOO_SHORT = 1
Public Const FILTER_TOO_LONG = 2
Public Const FILTER_BAD_WORDS = 4

'stores all email critical information
Public Type EMAIL_DATA
  subject As String
  from As String
  messagebody As String
  delete_code As Integer 'see filter constants above
End Type

'global arrays
Public em() As EMAIL_DATA 'stores all data required for filtering
Public word() As String          'stores list of SPAM words

'global variables
Public gintEmailToReview As Integer  'this is the number to be reviewed by frmReview
Public gintTotalEmails As Integer 'total number of emails to be downloaded
Public gstrString As String 'holds a string globally for passing between forms

'********************************************************************
' C L E A N U P   S T R I N G
' Removes punctuations, spaces, etc.
'********************************************************************
Public Function CleanupString(strIn As String) As String
  Dim strTemp As String
  Dim X As Integer
  
  'remove all spaces
  For X = 1 To Len(strIn)
    If Mid(strIn, X, 1) <> " " Then strTemp = strTemp & Mid(strIn, X, 1)
  Next X
  strIn = strTemp
   
  'convert to lower case
  strIn = LCase(strIn)
  
  'replace
  '| or ! or 1 with i
  '0       with o
  '@      with a
  '$       with s
  strTemp = ""
  For X = 1 To Len(strIn)
    If Mid(strIn, X, 1) = "!" Or Mid(strIn, X, 1) = "|" Or Mid(strIn, X, 1) = "1" Then
      strTemp = strTemp & "i"
    ElseIf Mid(strIn, X, 1) = "0" Then
      strTemp = strTemp & "o"
    ElseIf Mid(strIn, X, 1) = "@" Then
      strTemp = strTemp & "a"
    ElseIf Mid(strIn, X, 1) = "$" Then
      strTemp = strTemp & "s"
    Else
      strTemp = strTemp & Mid(strIn, X, 1)
    End If
  Next X
  strIn = strTemp
  
    'delete all remaining pronunciation  marks
  strTemp = ""
  
  For X = 1 To Len(strIn)
     If Asc(Mid(strIn, X, 1)) >= 48 And Asc(Mid(strIn, X, 1)) <= 57 Then
       strTemp = strTemp & Mid(strIn, X, 1)
     ElseIf Asc(Mid(strIn, X, 1)) >= 97 And Asc(Mid(strIn, X, 1)) <= 122 Then
       strTemp = strTemp & Mid(strIn, X, 1)
      Else
      End If
  Next X
  CleanupString = strTemp

End Function


'********************************************************************
' F I L T E R    S U B J E C T
'analyses subject for filtering state
'********************************************************************
Public Function FilterSubject(intNum As Integer) As String
  Dim intCode As Integer
  Dim strS, strTemp As String
  Dim strSubject As String
  Dim X, z As Integer
   
  'load original subject string
  intCode = FILTER_OK
  em(intNum).delete_code = intCode
  strSubject = em(intNum).subject
  strS = CleanupString(strSubject)
  
  'determines code
  'empty subjects
  If Len(strS) < 1 Then
    intCode = intCode + FILTER_TOO_SHORT
  End If
  
  'longer subjects
  If Len(strS) > 50 Then
    intCode = intCode + FILTER_TOO_LONG
  End If
  'search for spamwords in subject
  For z = 1 To UBound(word)
    If Len(word(z)) > 0 Then
      If InStr(1, strS, word(z)) > 0 Then
        intCode = intCode + FILTER_BAD_WORDS
        frmMain.Text3.Text = word(z)
        'Text3.Text = word(z)
        Exit For
      End If
    End If
  Next z
  em(intNum).delete_code = intCode 'save filter code...0 is okay
  'Text2.Text = strS
  FilterSubject = CStr(intCode)
End Function


