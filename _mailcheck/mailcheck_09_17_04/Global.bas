Attribute VB_Name = "Global"
'**************************************************************************
' G L O B A L . B A S  - December 2003
' Public variables, constants, arrays and subs/functions
'**************************************************************************
Option Explicit

'stores all email critical information
Public Type EMAIL_DATA
  Subject As String
  From As String
  MessageTo As String
  CC As String
  BCC As String
  MessageID As String
  SendDate As String
  Sender As String
  ReturnPath As String
  Size As String
  Comments As String
  Encrypted As String
  InReplyTo As String
  Received As String
  References As String
  MessageBody As String
  delete_code As Integer 'see filter constants above
  Bytes_total As Long
  Delete As Boolean 'true if email is to be deleted
  Dead As Boolean 'true if dead
  Score As Single
  'sub_word As String 'words that are spam words in subject
  'msg_word As String 'words that are spam words in message body
End Type

'used to store scores...constants or particular emails
Public Type EMAIL_SCORE
  EmptySubject As Single
  SubjectXAscii As Single
  SubjectNumbers As Single
  SubjectAddon As Single
  SubjectConsecCount As Single
  Attract As Single
  Degree As Single
  Finance As Single
  Porn As Single
  Misc As Single
  Health As Single
  Holiday As Single
  HWSW As Single
  BodyText As Single
  ReceivedUnknown As Single
  ToMissingUserName As Single
  ToMissingDomain As Single
  DomainCount As Single
  DateMissing As Single
  FromMissing As Single
  CountryCode As Single
  Friend As Single
  SpamMinimum As Single
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
Public g_uScore As EMAIL_SCORE

'global arrays
Public em() As EMAIL_DATA 'stores all data required for filtering
Public word() As String          'stores list of SPAM words

'8 categories of spam words that ranked in the top 8
Public g_sHealth() As String      'stores health related words
Public g_sFinance() As String     'financial, get rich
Public g_sPorn() As String        'porn words
Public g_sMisc() As String        'misc words not in other cats, like 'septic tank'
Public g_sHWSW() As String        'hardware/software
Public g_sAttract() As String     'attract a partner
Public g_sDegree() As String      'get a degree
Public g_sHoliday() As String     'vacation, holiday
Public g_sFriend() As String      'these items are most trusted

'friendly domain endings and country codes: i.e. .com, .net, .uk, etc.
Public g_sCountryCode() As String  'friendly country codes
Public g_sBodyText() As String 'unfriendly text in message body

'global variables
Public gintEmailToReview As Integer  'this is the number to be reviewed by frmReview
Public gintTotalEmails As Integer       'total number of emails to be downloaded
Public gstrString As String                 'holds a string globally for passing between forms
Public gstrProgram As String            'name of program
Public gstrVersion As String             'version of program
Public gstrDate As String                 'date of last program change
Public gblnDownloadComplete  As Boolean 'true if complete
Public gstrBadSubWord As String     'stores offending word from spam list found in subject
Public gstrBadMsgWord As String    '...found in message
Public gblnDumpToFile As Boolean   'true means dump when downloading
Public g_sUserName As String 'txtUserName
Public g_sUserDomain As String 'after @ in txtUserName
Public g_nScoreCutoff As Integer 'anything above this value will be deleted

'********************************************************************
' L O A D   G L O B A L  V A R I A B L E S
'********************************************************************
Public Sub LoadGlobalVariables()
  gblnMaxSubLen = True
  gblnMinSubLen = True
  gblnSubPhrases = True
  gblnSubConsonants = True
  gblnMsgConsonants = False
  gblnMsgphrases = True
  
  gintMaxSubLen = 50
  gintMinSubLen = 1
  gintMaxSubConsonants = 6
  gintMaxMsgConsonants = 10
  
  gstrProgram = "Mail Checker"
  gstrVersion = "v0.1f"
  gstrDate = "September 5, 2004"
  gblnDownloadComplete = False
  
  '***************************************************** Weighting for filter
  '   Change these to alter your score
  '   Determine how many points to add to overall score
  g_uScore.BodyText = 5
  g_uScore.CountryCode = 10
  g_uScore.DateMissing = 20  'forged date if missing
  g_uScore.EmptySubject = 20
  g_uScore.FromMissing = 40  'can't trust
  g_uScore.ReceivedUnknown = 20
  g_uScore.SubjectAddon = 15
  g_uScore.Attract = 5
  g_uScore.Degree = 5
  g_uScore.Finance = 5
  g_uScore.Health = 5
  g_uScore.Holiday = 5
  g_uScore.HWSW = 5
  g_uScore.Misc = 5
  g_uScore.SubjectNumbers = 5
  g_uScore.Porn = 5
  g_uScore.SubjectXAscii = 20 'extended ascii more than 50%
  g_uScore.SubjectConsecCount = 3
  g_uScore.DomainCount = 10  'too many of the same domain
  g_uScore.ToMissingDomain = 20
  g_uScore.ToMissingUserName = 20
  g_uScore.Friend = -20     'most trusted
  g_uScore.SpamMinimum = 25   'needs this many points to be spam
  
  '************************************ reading files into arrays
  Dim nFile As Integer
  Dim sFileName As String
  Dim sIn As String
  
  'load health array from health.txt file
  nFile = FreeFile
  ReDim g_sHealth(0)
  sFileName = Dir(App.Path & "\health.txt")
  If sFileName <> "" Then
    Open App.Path & "\health.txt" For Input As #nFile
      Do
        Line Input #nFile, sIn
        
        'length must be > 0 and 1st char not an apostrophe
        If Len(sIn) > 0 And Left(sIn, 1) <> "'" Then
          ReDim Preserve g_sHealth(UBound(g_sHealth) + 1)
          g_sHealth(UBound(g_sHealth)) = LCase(sIn)
        End If
      Loop Until EOF(nFile)
    Close #nFile
  Else
    MsgBox "Health.txt is not available!"
  End If
  
  'load finance array from finance.txt file
  nFile = FreeFile
  ReDim g_sFinance(0)
  sFileName = Dir(App.Path & "\finance.txt")
  If sFileName <> "" Then
    Open App.Path & "\finance.txt" For Input As #nFile
      Do
        Line Input #nFile, sIn
        If Len(sIn) > 0 And Left(sIn, 1) <> "'" Then
          ReDim Preserve g_sFinance(UBound(g_sFinance) + 1)
          g_sFinance(UBound(g_sFinance)) = LCase(sIn)
        End If
      Loop Until EOF(nFile)
    Close #nFile
  Else
    MsgBox "Finance.txt is not available!"
  End If
  
  
  'load porn array from porn.txt file
  nFile = FreeFile
  ReDim g_sPorn(0)
  sFileName = Dir(App.Path & "\porn.txt")
  If sFileName <> "" Then
    Open App.Path & "\porn.txt" For Input As #nFile
      Do
        Line Input #nFile, sIn
        If Len(sIn) > 0 And Left(sIn, 1) <> "'" Then
          ReDim Preserve g_sPorn(UBound(g_sPorn) + 1)
          g_sPorn(UBound(g_sPorn)) = LCase(sIn)
        End If
      Loop Until EOF(nFile)
    Close #nFile
  Else
    MsgBox "Porn.txt is not available!"
  End If
  
  'load misc array from misc.txt file
  nFile = FreeFile
  ReDim g_sMisc(0)
  sFileName = Dir(App.Path & "\misc.txt")
  If sFileName <> "" Then
    Open App.Path & "\misc.txt" For Input As #nFile
      Do
        Line Input #nFile, sIn
        If Len(sIn) > 0 And Left(sIn, 1) <> "'" Then
          ReDim Preserve g_sMisc(UBound(g_sMisc) + 1)
          g_sMisc(UBound(g_sMisc)) = LCase(sIn)
        End If
      Loop Until EOF(nFile)
    Close #nFile
  Else
    MsgBox "Misc.txt is not available!"
  End If
  
  'load hwsw array from hwsw.txt file
  nFile = FreeFile
  ReDim g_sHWSW(0)
  sFileName = Dir(App.Path & "\hwsw.txt")
  If sFileName <> "" Then
    Open App.Path & "\hwsw.txt" For Input As #nFile
      Do
        Line Input #nFile, sIn
        If Len(sIn) > 0 And Left(sIn, 1) <> "'" Then
          ReDim Preserve g_sHWSW(UBound(g_sHWSW) + 1)
          g_sHWSW(UBound(g_sHWSW)) = LCase(sIn)
        End If
      Loop Until EOF(nFile)
    Close #nFile
  Else
    MsgBox "hwsw.txt is not available!"
  End If
  
  'load attract array from attract.txt file
  nFile = FreeFile
  ReDim g_sAttract(0)
  sFileName = Dir(App.Path & "\attract.txt")
  If sFileName <> "" Then
    Open App.Path & "\attract.txt" For Input As #nFile
      Do
        Line Input #nFile, sIn
        If Len(sIn) > 0 And Left(sIn, 1) <> "'" Then
          ReDim Preserve g_sAttract(UBound(g_sAttract) + 1)
          g_sAttract(UBound(g_sAttract)) = LCase(sIn)
        End If
      Loop Until EOF(nFile)
    Close #nFile
  Else
    MsgBox "attract.txt is not available!"
  End If
  
  'load degree array from degree.txt file
  nFile = FreeFile
  ReDim g_sDegree(0)
  sFileName = Dir(App.Path & "\degree.txt")
  If sFileName <> "" Then
    Open App.Path & "\degree.txt" For Input As #nFile
      Do
        Line Input #nFile, sIn
        If Len(sIn) > 0 And Left(sIn, 1) <> "'" Then
          ReDim Preserve g_sDegree(UBound(g_sDegree) + 1)
          g_sDegree(UBound(g_sDegree)) = LCase(sIn)
        End If
      Loop Until EOF(nFile)
    Close #nFile
  Else
    MsgBox "degree.txt is not available!"
  End If
  
  'load holiday array from holiday.txt file
  nFile = FreeFile
  ReDim g_sHoliday(0)
  sFileName = Dir(App.Path & "\holiday.txt")
  If sFileName <> "" Then
    Open App.Path & "\holiday.txt" For Input As #nFile
      Do
        Line Input #nFile, sIn
        If Len(sIn) > 0 And Left(sIn, 1) <> "'" Then
          ReDim Preserve g_sHoliday(UBound(g_sHoliday) + 1)
          g_sHoliday(UBound(g_sHoliday)) = LCase(sIn)
        End If
      Loop Until EOF(nFile)
    Close #nFile
  Else
    MsgBox "holiday.txt is not available!"
  End If
  
  'load country code array from friendcountrycode.txt file
  nFile = FreeFile
  ReDim g_sCountryCode(0)
  sFileName = Dir(App.Path & "\friendcountrycode.txt")
  If sFileName <> "" Then
  Open App.Path & "\friendcountrycode.txt" For Input As #nFile
    Do
      Line Input #nFile, sIn
      If Len(sIn) > 4 And Left(sIn, 1) <> "'" Then
        ReDim Preserve g_sCountryCode(UBound(g_sCountryCode) + 1)
        g_sCountryCode(UBound(g_sCountryCode)) = RTrim(LCase(Left(sIn, 4)))  'grab 4 characters
      End If
    Loop Until EOF(nFile)
  Close #nFile
    Else
    MsgBox "friendcountrycode.txt is not available!"
  End If
  
  'load body text array from unfriendbody.txt file
  nFile = FreeFile
  ReDim g_sBodyText(0)
  Open App.Path & "\unfriendlybody.txt" For Input As #nFile
    Do
      Line Input #nFile, sIn
      If Len(sIn) > 0 And Left(sIn, 1) <> "'" Then
        ReDim Preserve g_sBodyText(UBound(g_sBodyText) + 1)
        g_sBodyText(UBound(g_sBodyText)) = RTrim(LCase(sIn))
      End If
    Loop Until EOF(nFile)
  Close #nFile

  'load friendly text array from friendbody.txt file
  nFile = FreeFile
  ReDim g_sFriend(0)
  Open App.Path & "\friendlybody.txt" For Input As #nFile
    Do
      Line Input #nFile, sIn
      If Len(sIn) > 0 And Left(sIn, 1) <> "'" Then
        ReDim Preserve g_sFriend(UBound(g_sFriend) + 1)
        g_sFriend(UBound(g_sFriend)) = RTrim(LCase(sIn))
      End If
    Loop Until EOF(nFile)
  Close #nFile

End Sub

'****************************************** LookForEmptySubject +2
Public Function LookForEmptySubject(strIn As String) As Single
  LookForEmptySubject = 0
  If Len(strIn) < 1 Then LookForEmptySubject = 2
End Function

'****************************************** LookForSubjectAddOn +2
Public Function LookForSubjectAddOn(strIn As String) As Single
 LookForSubjectAddOn = 0
 Dim nPos As Integer
 nPos = InStr(1, strIn, Space(4))
 If nPos > 0 Then
  If nPos + 3 < Len(strIn) Then
    LookForSubjectAddOn = g_uScore.SubjectAddon
  End If
 End If
  
End Function

'****************************************** LookForSubjectXAscii +2
Public Function LookForSubjectXAscii(strIn As String) As Single
  LookForSubjectXAscii = 0
  
  If Len(strIn) < 1 Then Exit Function
  
  Dim i As Integer, nReg As Integer, nX As Integer
  For i = 1 To Len(strIn)
    If Asc(Mid(strIn, i, 1)) > 127 Then
      nX = nX + 1
    Else
      nReg = nReg + 1
    End If
  Next i
  
  If nX > nReg Then LookForSubjectXAscii = g_uScore.SubjectXAscii
End Function

'****************************************** LookForSubjectConsecCon
Public Function LookForSubjectConsecCon(sIn As String) As Single
  LookForSubjectConsecCon = 0
  If Len(sIn) < 1 Then Exit Function
  
  If ConsecConsCount(sIn) > 5 Then LookForSubjectConsecCon = g_uScore.SubjectConsecCount
End Function

'****************************************** LookForSubjectNumbers +2
'subjects normally don't have too many number in
'it. The limit here is adjusted to 5.
Public Function LookForSubjectNumbers(strIn As String) As Single
  LookForSubjectNumbers = 0
  
  If Len(strIn) < 1 Then Exit Function
  
  Dim i As Integer, nLetters As Integer, nNumbers As Integer
  For i = 1 To Len(strIn)
    
    ' 0 through 9 test
    If Asc(Mid(strIn, i, 1)) > 47 And Asc(Mid(strIn, i, 1)) < 58 Then
      nNumbers = nNumbers + 1
    End If
  Next i
  
  If nNumbers > 5 Then LookForSubjectNumbers = g_uScore.SubjectNumbers
End Function

'**************************************** LookForHealth +2
'health related words
Public Function LookForHealth(strIn As String) As Single
  LookForHealth = 0
  Dim sTemp As String, nFound As Long, i As Integer

  sTemp = LCase(RemoveSpaces(strIn))
  
  For i = 1 To UBound(g_sHealth)
    If InStr(1, g_sHealth(i), "*") < 1 Then   'no wildcards in word
      If InStr(1, sTemp, g_sHealth(i)) > 0 Then
        LookForHealth = LookForHealth + g_uScore.Health
      End If
    Else                                ' evaluate wild card *
      If GetPatternInStr(sTemp, g_sHealth(i), "*") > 0 Then
        LookForHealth = LookForHealth + g_uScore.Health
      End If
    End If
  Next i
End Function

'**************************************** LookForFinance +2
Public Function LookForFinance(strIn As String) As Single
  LookForFinance = 0
  ' remove spaces
  Dim sTemp As String, nFound As Long
  
  sTemp = LCase(RemoveSpaces(strIn))
  
  Dim i As Integer
  
  For i = 1 To UBound(g_sFinance)
    If InStr(1, g_sFinance(i), "*") < 1 Then
      If InStr(1, sTemp, g_sFinance(i)) > 0 Then
        LookForFinance = LookForFinance + g_uScore.Finance
        'Exit For
      End If
    Else  ' evaluate wild card *
      If GetPatternInStr(sTemp, g_sFinance(i), "*") > 0 Then
        LookForFinance = LookForFinance + g_uScore.Finance
      End If
    End If
  Next i
End Function

'**************************************** LookForPorn +2
Public Function LookForPorn(strIn As String) As Single
  LookForPorn = 0
  ' remove spaces
  Dim sTemp As String, nFound As Long
  
  sTemp = LCase(RemoveSpaces(strIn))
  
  Dim i As Integer
  
  For i = 1 To UBound(g_sPorn)
    If InStr(1, g_sPorn(i), "*") < 1 Then
      If InStr(1, sTemp, g_sPorn(i)) > 0 Then
        LookForPorn = LookForPorn + g_uScore.Porn
        'Exit For
      End If
    Else  ' evaluate wild card *
      If GetPatternInStr(sTemp, g_sPorn(i), "*") > 0 Then
        LookForPorn = LookForPorn + g_uScore.Porn
      End If
    End If
  Next i
End Function


'**************************************** LookForMisc +2
Public Function LookForMisc(strIn As String) As Single
  LookForMisc = 0
  ' remove spaces
  Dim sTemp As String, nFound As Long
  
  sTemp = LCase(RemoveSpaces(strIn))
  
  Dim i As Integer
  
  For i = 1 To UBound(g_sMisc)
    If InStr(1, g_sMisc(i), "*") < 1 Then
      If InStr(1, sTemp, g_sMisc(i)) > 0 Then
        LookForMisc = LookForMisc + g_uScore.Misc
        'Exit For
      End If
    Else  ' evaluate wild card *
      If GetPatternInStr(sTemp, g_sMisc(i), "*") > 0 Then
        LookForMisc = LookForMisc + g_uScore.Misc
      End If
    End If
  Next i
End Function


'**************************************** LookForHWSW +2
Public Function LookForHWSW(strIn As String) As Single
  LookForHWSW = 0
  ' remove spaces
  Dim sTemp As String, nFound As Long
  
  sTemp = LCase(RemoveSpaces(strIn))
  
  Dim i As Integer
  
  For i = 1 To UBound(g_sHWSW)
    If InStr(1, g_sHWSW(i), "*") < 1 Then
      If InStr(1, sTemp, g_sHWSW(i)) > 0 Then
        LookForHWSW = LookForHWSW + g_uScore.HWSW
        'Exit For
      End If
    Else  ' evaluate wild card *
      If GetPatternInStr(sTemp, g_sHWSW(i), "*") > 0 Then
        LookForHWSW = LookForHWSW + g_uScore.HWSW
      End If
    End If
  Next i
End Function

'**************************************** LookForDegree +2
Public Function LookForDegree(strIn As String) As Single
  LookForDegree = 0
  ' remove spaces
  Dim sTemp As String, nFound As Long
  
  sTemp = LCase(RemoveSpaces(strIn))
  
  Dim i As Integer
  
  For i = 1 To UBound(g_sDegree)
    If InStr(1, g_sDegree(i), "*") < 1 Then
      If InStr(1, sTemp, g_sDegree(i)) > 0 Then
        LookForDegree = LookForDegree + g_uScore.Degree
        'Exit For
      End If
    Else  ' evaluate wild card *
      If GetPatternInStr(sTemp, g_sDegree(i), "*") > 0 Then
        LookForDegree = LookForDegree + g_uScore.Degree
      End If
    End If
  Next i
End Function

'**************************************** LookForAttract +2
Public Function LookForAttract(strIn As String) As Single
  LookForAttract = 0
  ' remove spaces
  Dim sTemp As String, nFound As Long
  
  sTemp = LCase(RemoveSpaces(strIn))
  
  Dim i As Integer
  
  For i = 1 To UBound(g_sAttract)
    If InStr(1, g_sAttract(i), "*") < 1 Then
      If InStr(1, sTemp, g_sAttract(i)) > 0 Then
        LookForAttract = LookForAttract + g_uScore.Attract
        'Exit For
      End If
    Else  ' evaluate wild card *
      If GetPatternInStr(sTemp, g_sFinance(i), "*") > 0 Then
        LookForAttract = LookForAttract + g_uScore.Attract
      End If
    End If
  Next i
End Function


'**************************************** LookForHoliday +2
Public Function LookForHoliday(strIn As String) As Single
  LookForHoliday = 0
  ' remove spaces
  Dim sTemp As String, nFound As Long
  
  sTemp = LCase(RemoveSpaces(strIn))
  
  Dim i As Integer
  
  For i = 1 To UBound(g_sHoliday)
    If InStr(1, g_sHoliday(i), "*") < 1 Then
      If InStr(1, sTemp, g_sHoliday(i)) > 0 Then
        LookForHoliday = LookForHoliday + g_uScore.Holiday
        'Exit For
      End If
    Else  ' evaluate wild card *
      If GetPatternInStr(sTemp, g_sHoliday(i), "*") > 0 Then
        LookForHoliday = LookForHoliday + g_uScore.Holiday
      End If
    End If
  Next i
End Function

'****************************************** LookForReceivedUnknown +2
Public Function LookForReceivedUnknown(sIn As String) As Single
  LookForReceivedUnknown = 0
  Dim sTemp As String
  sTemp = LCase(RemoveSpaces(sIn))
  If InStr(1, sTemp, "fromunknown") Then LookForReceivedUnknown = g_uScore.ReceivedUnknown
  If InStr(1, sTemp, "fromunnobody") Then LookForReceivedUnknown = g_uScore.ReceivedUnknown
End Function

'****************************************** LookForDomainCount +1 each
Public Function LookForDomainCount(sIn As String) As Single
  LookForDomainCount = 0
  If Len(sIn) < 1 Then Exit Function
  If Len(sIn) < Len(g_sUserDomain) Then Exit Function
  
  Dim sTemp As String
  Dim i As Integer, nCt As Integer, j As Integer
    
  sTemp = LCase(RemoveSpaces(sIn))
  For i = 1 To Len(sTemp)
    If Mid(sTemp, i, Len(g_sUserDomain)) = LCase(g_sUserDomain) Then
      nCt = nCt + 1
    End If
  Next i
  
  If nCt > 1 Then
    LookForDomainCount = nCt * g_uScore.DomainCount
  End If
  
End Function

'****************************************** LookForMissingDomain +2
Public Function LookForMissingDomain(sIn As String) As Single
  LookForMissingDomain = 0
  If Len(sIn) < 1 Then Exit Function
  If Len(sIn) < Len(g_sUserDomain) Then Exit Function
  
  Dim sTemp As String
  Dim i As Integer, nCt As Integer, j As Integer
    
  sTemp = LCase(RemoveSpaces(sIn))
  nCt = 0
  For i = 1 To Len(sTemp)
    If Mid(sTemp, i, Len(g_sUserDomain)) = LCase(g_sUserDomain) Then
      nCt = nCt + 1
    End If
  Next i
  
  'domain not found in this string
  If nCt < 1 Then
    LookForMissingDomain = g_uScore.ToMissingDomain
  End If
  
End Function

'****************************************** LookForMissingUserName +2
Public Function LookForMissingUserName(sIn As String) As Single
  LookForMissingUserName = 0
  If Len(sIn) < 1 Then Exit Function
  If Len(sIn) < Len(g_sUserName) Then Exit Function
  
  Dim sTemp As String
  Dim i As Integer, nCt As Integer, j As Integer
    
  sTemp = LCase(RemoveSpaces(sIn))
  nCt = 0
  For i = 1 To Len(sTemp)
    If Mid(sTemp, i, Len(g_sUserName)) = LCase(g_sUserName) Then
      nCt = nCt + 1
    End If
  Next i
  
  'domain not found in this string
  If nCt < 1 Then
    LookForMissingUserName = g_uScore.ToMissingUserName
  End If
  
End Function

'***************************************** LookForMissingDate +2
Public Function LookForMissingDate(sIn As String)
  LookForMissingDate = 0
  If Len(sIn) < 1 Then LookForMissingDate = g_uScore.DateMissing
End Function

'***************************************** LookForMissingFromAddress +2
Public Function LookForMissingFromAddress(sIn As String)
  LookForMissingFromAddress = 0
  If InStr(1, sIn, "@") < 1 Then LookForMissingFromAddress = g_uScore.FromMissing
End Function

'**************************************** LookForAttract +2
'must find <skdjfksdf.com>
Public Function LookForCountryCode(strIn As String) As Single
  LookForCountryCode = 0
  If Len(strIn) < 1 Then Exit Function 'missing strIn
  LookForCountryCode = g_uScore.CountryCode 'assume bad unless good code found
  
  ' remove spaces
  Dim sTemp As String, i As Integer, sCode As String
  Dim a As Integer, b As Integer, c As Integer 'positions
  
  sTemp = LCase(RemoveSpaces(strIn))
  
  a = GetPatternInStr(sTemp, ".**>", "*")
  b = GetPatternInStr(sTemp, ".***>", "*")
  If a > 0 Then
    sCode = Mid(sTemp, a + 1, 2)
  ElseIf b > 0 Then
    sCode = Mid(sTemp, b + 1, 3)
  Else
    LookForCountryCode = 0  'missing <  . >
    Exit Function
  End If
    
  
  'sCode contains extracted country code
  'now lets go through list
  For i = 1 To UBound(g_sCountryCode)
    'MsgBox sCode & " : " & g_sCountryCode(i)
    If sCode = g_sCountryCode(i) Then
      LookForCountryCode = 0 'match found
    End If
  Next i
  
  
End Function

'**************************************** LookForBodyText
Public Function LookForBodyText(strIn As String) As Single
  LookForBodyText = 0
  If Len(strIn) < 1 Then Exit Function 'missing strIn
  
  ' remove spaces
  Dim sTemp As String, i As Long, j As Long
  
  sTemp = LCase(RemoveSpaces(strIn))
  
  For i = 1 To UBound(g_sBodyText)
    If InStr(1, g_sBodyText(i), "*") < 1 Then
      If InStr(1, sTemp, g_sBodyText(i)) > 0 Then
        LookForBodyText = LookForBodyText + g_uScore.BodyText
      End If
    Else  ' evaluate wild card *
      If GetPatternInStr(sTemp, g_sBodyText(i), "*") > 0 Then
        LookForBodyText = LookForBodyText + g_uScore.BodyText
      End If
    End If
  Next i
  
End Function

'**************************************** LookForFriendText
Public Function LookForFriendText(strIn As String) As Single
  LookForFriendText = 0
  If Len(strIn) < 1 Then Exit Function 'missing strIn
  
  ' remove spaces
  Dim sTemp As String, i As Long
  
  sTemp = LCase(RemoveSpaces(strIn))
  
  'now lets go through list
  For i = 1 To UBound(g_sFriend)
    If InStr(1, sTemp, g_sFriend(i)) Then
      LookForFriendText = LookForFriendText + g_uScore.Friend 'match found
    End If
  Next i
  
End Function


'******************************** GetPatternInStr
Public Function GetPatternInStr(strIn As String, sPattern As String, sWildCard As String) As Long
  GetPatternInStr = 0
  
  If Len(strIn) < 1 Or Len(sPattern) < 1 Or Len(sWildCard) < 1 Then Exit Function
  ' only single character wild cards allowed
  If Len(sWildCard) > 1 Then sWildCard = Left(sWildCard, 1)
  
  Dim i As Long, j As Integer, nLet As Integer, nCount As Integer

  ReDim sPat(Len(sPattern) - 1) As String ' array for sPattern
  
  'load pattern array with each letter from sPattern
  For i = 1 To Len(sPattern)
    sPat(i - 1) = Mid(sPattern, i, 1)
    If sPat(i - 1) <> sWildCard Then nLet = nLet + 1 'count non-wildcard letters
    'frmMain.txtViewAll = frmMain.txtViewAll & "Pattern: " & sPat(i - 1) & vbCrLf
  Next i
  
  'go through strIn
  For i = 1 To Len(strIn)
      nCount = 0
      For j = 0 To UBound(sPat)
        If sPat(j) <> sWildCard And sPat(j) <> "" Then ' ignore all wildcards
          If Mid(strIn, i + j, 1) = sPat(j) Then
            nCount = nCount + 1
          End If
        End If
      Next j
      
      If nCount = nLet Then
        GetPatternInStr = i
        Exit For
      End If
  Next i

End Function

'******************************** RemoveNonAlphaNumeric
'keep only letters and numbers
Public Function RemoveNonAlphaNumeric(strIn As String) As String
  Dim i As Long
  Dim strTemp As String, sOut As String
  
  strTemp = LCase(RemoveSpaces(strIn))
  sOut = ""
  'remove all spaces
  For i = 1 To Len(strTemp)
    
    'numeric
    If Asc(Mid(strTemp, i, 1)) > 47 And Asc(Mid(strTemp, i, 1)) < 58 Then
      sOut = sOut & Mid(strTemp, i, 1)
      
    'alpha - lower and upper case
    ElseIf Asc(Mid(strTemp, i, 1)) > 64 And Asc(Mid(strTemp, i, 1)) < 123 Then
      sOut = sOut & Mid(strTemp, i, 1)
    End If
  Next i
    
  RemoveNonAlphaNumeric = sOut
End Function



'******************************** RemoveSpaces
Public Function RemoveSpaces(strIn As String) As String
  Dim i As Long
  Dim strTemp As String
  
  'remove all spaces
  For i = 1 To Len(strIn)
    If Mid(strIn, i, 1) <> " " Then strTemp = strTemp & Mid(strIn, i, 1)
  Next i
  RemoveSpaces = strTemp
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

