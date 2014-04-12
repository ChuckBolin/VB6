Attribute VB_Name = "Global"
'**************************************************************************
' G L O B A L . B A S  - December 2003
' Public variables, constants, arrays and subs/functions
'**************************************************************************
Option Explicit

'constants used to ID reasons for filtering
Public Const FILTER_OK = 0
Public Const FILTER_SUB_TOO_SHORT = 1
Public Const FILTER_SUB_TOO_LONG = 2
Public Const FILTER_SUB_BAD_WORDS = 4
Public Const FILTER_SUB_TOO_MANY_CONSONANTS = 8
Public Const FILTER_MSG_BAD_WORDS = 64
Public Const FILTER_MSG_TOO_MANY_CONSONANTS = 128

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
Public gblnSubConsonants As Boolean 'true if subject consonants are enabled
Public gblnSubPhrases As Boolean  'true if subject phrases is enabled
Public gblnMsgConsonants As Boolean 'true if message body consonants are enabled
Public gblnMsgphrases As Boolean

Public gintMaxSubLen As Integer     'max subject length
Public gintMinSubLen As Integer      'min subject length
Public gintMaxSubConsonants As Integer 'max allowable consecutive consonants in a string
Public gintMaxMsgConsonants As Integer 'max allowable consecutive consonants in the message string

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
  gblnSubPhrases = True
  gblnSubConsonants = True
  gblnMsgConsonants = True
  gblnMsgphrases = True
  
  gintMaxSubLen = 50
  gintMinSubLen = 1
  gintMaxSubConsonants = 6
  gintMaxMsgConsonants = 6
  
  gstrProgram = "Mail Checker"
  gstrVersion = "v0.1d"
  gstrDate = "December 17, 2003"
  gblnDownloadComplete = False
End Sub

'********************************************************************
' C L E A N U P   S T R I N G
' Removes punctuations, spaces, etc.
'********************************************************************
Public Function CleanupString(strIn As String) As String
  Dim strTemp As String
  Dim X As Long
  
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
    ElseIf Mid(strIn, X, 1) = "3" Then
      strTemp = strTemp & "e"
    ElseIf Mid(strIn, X, 1) = "4" Then
      strTemp = strTemp & "g"
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
  
  'this converts tags inside email to nothing.  I have noticed tags </  > being placed into subjects recently
  'Add TAG filter here later.......................................................
    
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
' C L E A N U P   T A G S
' Removes text inside of HTML tags.
'********************************************************************
Public Function CleanupTag(strIn As String) As String
  Dim strTemp As String
  Dim X As Long
  
  'remove all spaces
  For X = 1 To Len(strIn)
   
   
   
  Next X
  
  'strIn = strTemp
   
  CleanupTag = strIn
End Function

'********************************************************************
' F I L T E R    S U B J E C T
'analyses subject for filtering state
'********************************************************************
Public Function FilterSubject(strSubject As String) As Integer
  Dim intCode As Integer
  Dim strS, strTemp As String
  'Dim strSubject As String
  Dim X, z As Integer
   
  'load original subject string
  intCode = FILTER_OK
  'em(intNum).delete_code = intCode
  'strSubject = em(intNum).subject
  strS = CleanupString(strSubject)
  
  'determines code
  'empty subjects
  If Len(strS) < gintMinSubLen And gblnMinSubLen = True Then
    intCode = intCode + FILTER_SUB_TOO_SHORT
  End If
  
  'longer subjects
  If Len(strS) > gintMaxSubLen And gblnMaxSubLen = True Then
    intCode = intCode + FILTER_SUB_TOO_LONG
  End If
  'search for spamwords in subject
  If gblnSubPhrases = True Then
    For z = 1 To UBound(word)
      If Len(word(z)) > 0 Then
        If InStr(1, strS, word(z)) > 0 Then
          intCode = intCode + FILTER_SUB_BAD_WORDS
          Exit For
        End If
      End If
    Next z
  End If
  
  'search for a larg string of consonants
  If ConsecConsCount(CStr(strS)) > gintMaxSubConsonants And gblnSubConsonants = True Then
    intCode = intCode + FILTER_SUB_TOO_MANY_CONSONANTS
  End If
  
  'returns code corresponding to reason that the email has been flagged as spam
  FilterSubject = intCode
End Function

'********************************************************************
' F I L T E R    M E S S A  G E
'analyses message body for filtering state
'********************************************************************
Public Function FilterMessage(strMessage As String, intNum As Integer) As Integer
  Dim intCode As Integer
  Dim strS, strTemp As String
  Dim X, z As Integer
   
  'load original subject string
  intCode = intNum
  
  strS = CleanupTag(strMessage) 'removes HTML tags </ > from message
  strS = CleanupString(CStr(strS))         'normal cleanup
  
  'search for spamwords and phrases in body
  If gblnMsgphrases = True Then
    For z = 1 To UBound(word)
      If Len(word(z)) > 0 Then
        If InStr(1, strS, word(z)) > 0 Then
          intCode = intCode + FILTER_MSG_BAD_WORDS
          Exit For
        End If
      End If
    Next z
  End If
  
  'search for a larg string of consonants...uses uncleaned string
  If ConsecConsCount(CStr(strMessage)) > gintMaxMsgConsonants And gblnMsgConsonants = True Then
    intCode = intCode + FILTER_MSG_TOO_MANY_CONSONANTS
  End If
  
  'returns code corresponding to reason that the email has been flagged as spam
  FilterMessage = intCode
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
  Dim X As Long
  Dim ct As Integer
  Dim intMost As Integer
  Dim strFrag As String
  
  strFrag = "" 'URLs in subject defeat this filter..so ignore httpwww
  
  intMost = 0
  If Len(strS) < 1 Then ConsecConsCount = 0: Exit Function
  For X = 1 To Len(strS)
    If CharType(Mid(strS, X, 1)) = 1 Then  'if its a consonant
      ct = ct + 1
      strFrag = strFrag & Mid(strS, X, 1)
      If InStr(1, strFrag, "httpwww") > 0 Then 'url fragment not found...therefore keep
        ct = 0
        strFrag = ""
      End If
    Else
      If ct > 0 Then
        If ct > intMost Then intMost = ct 'save count if it is the highest so far
      End If
      ct = 0
      strFrag = ""
    End If
  Next X

  ConsecConsCount = intMost
  
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
