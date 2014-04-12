Attribute VB_Name = "Global"
'**************************************************************************
' G L O B A L . B A S  - December 2003
' Public variables, constants, arrays and subs/functions
'**************************************************************************
Option Explicit

'constants used to ID reasons for filtering
Public Const FILTER_OK = 0
Public Const FILTER_TOO_SHORT = 1
Public Const FILTER_TOO_LONG = 2
Public Const FILTER_BAD_WORDS = 4
Public Const FILTER_TOO_MANY_CONSONANTS = 8

'stores all email critical information
Public Type EMAIL_DATA
  subject As String
  from As String
  cc As String
  messagebody As String
  delete_code As Integer 'see filter constants above
  bytes_total As Long
End Type

'variables to store global filter information
Public gblnMaxSubLen As Boolean  'true if max subject len is enabled
Public gblnMinSubLen As Boolean   'true if min subject len is enabled
Public gintMaxSubLen As Integer     'max subject length
Public gintMinSubLen As Integer      'min subject length
Public gblnSubPhrases As Boolean  'true if subject phrases is enabled
Public gblnSubConsonants As Boolean 'true if subject consonants are enabled
Public gintMaxSubConsonants As Integer 'max allowable consecutive consonants in a string
'global arrays
Public em() As EMAIL_DATA 'stores all data required for filtering
Public word() As String          'stores list of SPAM words

'global variables
Public gintEmailToReview As Integer  'this is the number to be reviewed by frmReview
Public gintTotalEmails As Integer       'total number of emails to be downloaded
Public gstrString As String                 'holds a string globally for passing between forms
Public gstrProgram As String            'name of program
Public gstrVersion As String             'version of program
Public gstrDate As String                 'date of last program change
Public gblnDownloadComplete  As Boolean 'true if complete

'********************************************************************
' C L E A N U P   S T R I N G
' Removes punctuations, spaces, etc.
'********************************************************************
Public Sub LoadGlobalVariables()
  gblnMaxSubLen = True
  gblnMinSubLen = True
  gintMaxSubLen = 50
  gintMinSubLen = 1
  gblnSubPhrases = True
  gblnSubConsonants = True
  gintMaxSubConsonants = 6
  gstrProgram = "Mail Checker"
  gstrVersion = "v0.1c"
  gstrDate = "December 16, 2003"
  gblnDownloadComplete = False
End Sub

'********************************************************************
' C L E A N U P   S T R I N G
' Removes punctuations, spaces, etc.
'********************************************************************
Public Function CleanupString(strIn As String) As String
  Dim strTemp As String
  Dim x As Integer
  
  'remove all spaces
  For x = 1 To Len(strIn)
    If Mid(strIn, x, 1) <> " " Then strTemp = strTemp & Mid(strIn, x, 1)
  Next x
  strIn = strTemp
   
  'convert to lower case
  strIn = LCase(strIn)
  
  'replace
  '| or ! or 1 with i
  '0       with o
  '@      with a
  '$       with s
  strTemp = ""
  For x = 1 To Len(strIn)
    If Mid(strIn, x, 1) = "!" Or Mid(strIn, x, 1) = "|" Or Mid(strIn, x, 1) = "1" Then
      strTemp = strTemp & "i"
    ElseIf Mid(strIn, x, 1) = "3" Then
      strTemp = strTemp & "e"
    ElseIf Mid(strIn, x, 1) = "4" Then
      strTemp = strTemp & "g"
    ElseIf Mid(strIn, x, 1) = "0" Then
      strTemp = strTemp & "o"
    ElseIf Mid(strIn, x, 1) = "@" Then
      strTemp = strTemp & "a"
    ElseIf Mid(strIn, x, 1) = "$" Then
      strTemp = strTemp & "s"
    Else
      strTemp = strTemp & Mid(strIn, x, 1)
    End If
  Next x
  strIn = strTemp
  
  'delete all remaining pronunciation  marks
  strTemp = ""
  
  'this converts tags inside email to nothing.  I have noticed tags </  > being placed into subjects recently
  'Add TAG filter here later.......................................................
    
  For x = 1 To Len(strIn)
     If Asc(Mid(strIn, x, 1)) >= 48 And Asc(Mid(strIn, x, 1)) <= 57 Then
       strTemp = strTemp & Mid(strIn, x, 1)
     ElseIf Asc(Mid(strIn, x, 1)) >= 97 And Asc(Mid(strIn, x, 1)) <= 122 Then
       strTemp = strTemp & Mid(strIn, x, 1)
      Else
      End If
  Next x
  CleanupString = strTemp

End Function

'********************************************************************
' F I L T E R    S U B J E C T
'analyses subject for filtering state
'********************************************************************
Public Function FilterSubject(strSubject) As Integer
  Dim intCode As Integer
  Dim strS, strTemp As String
  'Dim strSubject As String
  Dim x, z As Integer
   
  'load original subject string
  intCode = FILTER_OK
  'em(intNum).delete_code = intCode
  'strSubject = em(intNum).subject
  strS = CleanupString(CStr(strSubject))
  
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
       ' frmMain.Text3.Text = word(z)
        'Text3.Text = word(z)
        Exit For
      End If
    End If
  Next z
  
  'search for a larg string of consonants
  If ConsecConsCount(CStr(strS)) > gintMaxSubConsonants Then intCode = FILTER_TOO_MANY_CONSONANTS
    
  'em(intNum).delete_code = intCode 'save filter code...0 is okay
  'Text2.Text = strS
  FilterSubject = intCode
End Function

'********************************************************************
' C O N S E C    C O N S    C O U N T
' Many emails have strings of 10 to 20 characters that are
' just a random pattern of numbers. Most of these strings
' are predominantly consonants.  This function returns the
' max number of consonants in a sequence. More than
' six normally indicates spam
'********************************************************************
Public Function ConsecConsCount(strS As String) As Integer
  Dim x As Integer
  Dim ct As Integer
  
  If Len(strS) < 1 Then ConsecConsCount = 0: Exit Function
  For x = 1 To Len(strS)
    If CharType(Mid(strS, x, 1)) = 1 Then  'if its a consonant
      ct = ct + 1
    Else
      ct = 0
    End If
  Next x
End Function

'********************************************************************
' C H A R   T Y P E
' Types:  0=Nothing (ASC 0-32), 1=Consonant, 2=Vowel,
' 3=Everything else such as punctuation and other marks
'********************************************************************
Public Function CharType(strS As String) As Integer
  If Len(strS) < 1 Then CharType = 0: Exit Function  'eliminate empty character
  If Asc(strS) < 33 Then CharType = 0: Exit Function 'space key and lower in ascii value
  CharType = 3
  strS = LCase(strS) 'convert to lowercase
  If Asc(strS) >= 97 And Asc(strS) <= 122 Then CharType = 1  'is it a letter?
  If strS = "a" Or strS = "e" Or strS = "i" Or strS = "o" Or strS = "u" Then CharType = 2 'is it a vowel?
End Function
