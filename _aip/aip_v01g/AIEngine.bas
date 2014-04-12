Attribute VB_Name = "AIEngineMod"
Option Explicit

'***************************************************************************
'AI ENGINE( )
'This sub returns the index or cell number within the grid
'related to its desired move. It returns a ZERO if no move is valid.
'MainSubs PLAYCOORDINATOR is responsible for calling
'this FUNCTION when it is the program's turn to play. The
'AIEngine relies on GETFACTBASEDMOVE to pick the best
'offensive/defensive play or use GETRANDOMMOVE if AI
'cannot produce a useable value.
'***************************************************************************
Public Function AIEngine() As Integer
  Dim nPlay As Integer
  Dim sMsg As String
  
  On Error GoTo MyError
  
  'AIP uses knowledge of program, teacher's positions within the grid as well as knowledge base to computer best move
   nPlay = GetFactBasedMove   'this is AI
  If nPlay = 0 Then nPlay = GetRandomMove  'this is guessing, the second best thing to AI ;-)
  
  'returns this play
  AIEngine = nPlay
  
  Exit Function

MyError:
  gsForm = "AIEngineMod"
  gsProcedure = "AIEngine"
  ErrorHandler
End Function

'*************************************************************************
' G E T  F A C T  B A S E D  M O V E
'Considers only ABS patterns with wins >0
'This function returns the best logical move for the program,
'either to play or block. In the event is has no clue then it
'returns 0 and AIP depends upon a random move.
'NOTE: This function is used in the program to simply update
'the display in the txtKB.text box.  I was too lazy to write just
'a display function. See PlayCoordinator to see what I mean
'*************************************************************************
Public Function GetFactBasedMove() As Integer
  Dim u, v, w, x, y As Integer 'used for looping
  Dim nProgramTotal As Integer 'stores total number of still playable and possible winning patterns for the program
  Dim nTeacherTotal As Integer ' stores total number of still playable and possible winning patterns for the teacher
  ReDim uProgram(nProgramTotal) As pattern2  'stores possible wins for the program from uABS based on current grid play
  ReDim uTeacher(nTeacherTotal) As pattern2  'stores possible wins for the teacher from uABS based on current grid play
  ReDim nProgramVSum(gnTotalCells) As Integer 'stores V sum (explained below)
  ReDim nTeacherVSum(gnTotalCells) As Integer
  Dim sWord As String 'stores string equivalent of binary number
  Dim sBin As String
  Dim lCombo As Long
  Dim sStatus As String
  Dim nChoice As Integer 'stores current cell favorite to play
  Dim nWins As Integer 'stores best historical wins
  Dim nSum As Integer 'stores V Sum
  Dim nBit As Integer 'stores selected bit or cell in selected bit pattern
  On Error GoTo MyError
  GetFactBasedMove = 0 'default value meaning that AIP will select random move
  gsLocation = "10"
  
  'diplays global variables for debug
  glAllCells = glProgram Or glTeacher                   'this shows a bit '1' if cell is occupied by teacher or program
  glFreeCells = glAllCells Xor (2 ^ gnTotalCells) - 1  'XOR allows the cells to be inverted. '1' means cell is free
  
  'formats display of data
  sBin = "": GetBinaryString sBin, glAllCells: sWord = sWord & "All Used: " & sBin & vbCrLf
  sBin = "": GetBinaryString sBin, glFreeCells: sWord = sWord & "Free:     " & sBin & vbCrLf
  sBin = "": GetBinaryString sBin, glProgram: sWord = sWord & "Prog:    " & sBin & vbCrLf
  sBin = "": GetBinaryString sBin, glTeacher: sWord = sWord & "Teach:  " & sBin & vbCrLf
  sStatus = sWord
  
  'must read all playable wins into both arrays.  To do this, the uABS winning pattern must be AND'ed together with
  'the combined position of the cells with program symbols and free cells.  For teacher it must be the combined
  'position of the cells with teacher symbols and free cells
  For x = 1 To UBound(uABS)
    'get all patterns that contain the same bits as the free cells and the program cells
    lCombo = glFreeCells Or glProgram
    If (uABS(x).word And lCombo) = uABS(x).word Then    'potential to win with this pattern...add it to array
        nProgramTotal = nProgramTotal + 1
        ReDim Preserve uProgram(nProgramTotal) As pattern2
        uProgram(nProgramTotal).word = uABS(x).word And glFreeCells  'now remove cells already play..no longer needed
        uProgram(nProgramTotal).wins = uABS(x).wins
    End If
    
    'get all patterns that contain the same bits as the free cells and the program cells
    lCombo = glFreeCells Or glTeacher
    If (uABS(x).word And lCombo) = uABS(x).word Then 'teacher can win with this
        nTeacherTotal = nTeacherTotal + 1
        ReDim Preserve uTeacher(nTeacherTotal) As pattern2
        uTeacher(nTeacherTotal).word = uABS(x).word And glFreeCells 'remove cells alreadyed played...not needed
        uTeacher(nTeacherTotal).wins = uABS(x).wins
    End If
  Next x
  

  '**************************************************************************************************************************************
  '   M O S T    I M P O R T A N T    P A R T     O F    A I    P R O G R A M    F O L L O W S ! ! !
  '**************************************************************************************************************************************
  'now winning patterns have been reduced from complete list based upon which cells have already been played
  'and the cells that are still empty or free. Since the plays that have already occurred are no longer necessary, the
  'two arrays show only bits that are still free and will lead to a win if played.
  'These remaining patterns must have two operations done to them.
  'From the patterns, a horizontal (H) sum will be produced.  A horizontal sum of 1 means a win occurs the very next play,
  ' a 2 means a win is in two plays and a 3 means 3 plays, etc.
  'From the patterns, a vertical (V) sum of bits in the same column of one or more patterns is summed up.  This
  'calculates which cell when played produces the highest probability of a win.  Therefore the bigger number is
  'better.
  'Now the H and V sums are calculated for both teacher and program.  From  this data, AIP is able to make the
  'best offensive/defensive plays.  This is the heart of decision making and thinking by AI.  Poor coding here
  'produces a dumb AI program...really cool and thoughtful programming produces a really intelligent program
  'that plays as well as most people learning a game for the first time.
  '
  'cells
  ' 9 8 7 6 5 4 3 2 1
  '============
  ' 0 0 0 1 1 1 0 0 0  =  3
  ' 1 0 0 0 0 0 1 0 0  =  2
  ' 1 0 0 0 0 0 0 0 0  =  1
  '-------------------------
  ' 2 0 0 1 1 1 1 0 0
  '
  'The three rows of binary numbers above have H sums and V sums.  The H sum =1 means this is the play
  'that will win the game.  The H sum = 3 means this is 3 plays away from a win.
  'Cell 9 has a V sum of 2, meaning playing this cell satisifies the patterns for two possible wins.
  'AIP will choose the row with H sum =1 because that is a definite winner.  AIP also produces the above table
  'for the teacher.  If AIP's lowest H sum=2 and the lowest H sum for the teacher is 1, then AIP will play to block
  'the win.
  'So let's get started...
  '*************************************************************************************************************************************
  gsLocation = "100"
  
  'let's calculate H sum for each pattern in both arrays
  For x = 1 To UBound(uProgram)
    uProgram(x).sum = GetHSum(uProgram(x).word)
  Next x
  For x = 1 To UBound(uTeacher)
    uTeacher(x).sum = GetHSum(uTeacher(x).word)
  Next x
  
  'lets calculate V sum and place into arrays
  ReDim nProgramVSum(gnTotalCells) As Integer
  ReDim nTeacherVSum(gnTotalCells) As Integer
  For x = 1 To UBound(uProgram)  'scroll through winning patterns for the program
    For y = 1 To gnTotalCells  'scan each bit..starting with LSB to MSB to develop V sum
      If 2 ^ (y - 1) And uProgram(x).word Then nProgramVSum(y) = nProgramVSum(y) + 1
    Next y
  Next x
  For x = 1 To UBound(uTeacher)  'scroll through winning patterns for the teacher
    For y = 1 To gnTotalCells  'scan each bit..starting with LSB to MSB to develop V sum
      If 2 ^ (y - 1) And uTeacher(x).word Then nTeacherVSum(y) = nTeacherVSum(y) + 1
    Next y
  Next x
  
  '**************************************************************************************************************
  ' K E Y    D A T A
  ' AIP now has the following information
  ' uProgram array stores all patterns and H sums for program to win
  ' uTeacher array stores all patterns and H sums for teacher to win
  ' nProgramVSum array stores all V sums for program
  ' nTeacherVSum array stores all V sums for teacher
  ' uProgram(x).wins and uTeacher(x).wins stores all wins with this pattern
  '
  ' Remember H sums tells AIP how many moves to win game (lowest number pattern is best)
  ' V sums tell which cells are used in the most patterns (highest number pattern is best)
  ' .wins tells how many times this pattern was won with (highest number pattern is bettern)
  '*************************************************************************************************************
  gsLocation = "200"
  'MsgBox UBound(uProgram) & "  " & UBound(uTeacher)
  
  'RULE 1:
  'if H sum exists for the program then this is the win...go for it
  For x = 0 To UBound(uProgram)
    If uProgram(x).sum = 1 Then
      For y = gnTotalCells To 1 Step -1
        If uProgram(x).word And (2 ^ (y - 1)) Then
          GetFactBasedMove = y
          Exit Function
        End If
      Next y
    End If
  Next x
 
  gsLocation = "250"
  'RULE 2:
  'so...AIP didn't win, can teacher win next? If so, play to block the human
  For x = 0 To UBound(uTeacher)
    If uTeacher(x).sum = 1 Then
      For y = gnTotalCells To 1 Step -1
        If uTeacher(x).word And (2 ^ (y - 1)) Then
          GetFactBasedMove = y
          Exit Function
        End If
      Next y
    End If
  Next x

  gsLocation = "260"
  'RULE 3:
  'no next time winners, so starting with a sum of 2, lets find who can win next and then what plays should be made
  'based upon V sum and historical wins for specific patterns
  nChoice = 0: nWins = 0
  For u = 2 To gnTotalCells
    For v = 0 To UBound(uProgram)
      If uProgram(v).sum = u Then  'looking for H sums equal to the smallest number...probably lots of patterns with
                                                 'the same H sum, therefore, pick the one with the highest historical wins
        If uProgram(v).wins > nWins Then
          nWins = uProgram(v).wins  'remember, highest no. of wins is best overall
          nChoice = v                       'stores bit pattern index of array..this is not the game choice yet...more to come
        End If
      End If
    Next v
    If nWins > 0 Then Exit For           'okay, found a winning bit pattern with nWins...lets break FOR..NEXT
  Next u
  
  'now let's scan binary bit pattern from LSB to MSB..if a '1' is found note the V Sum value.  Continue through all
  'bits. Goal: find the bit in the selected bit pattern with the highest V sum value. This is program best choice.
  nSum = 0: nBit = 0
  For x = gnTotalCells To 1 Step -1
    If uProgram(nChoice).word And (2 ^ (x - 1)) Then    'found bit in selected bit pattern
      If nProgramVSum(x) > nSum Then
        nSum = nProgramVSum(x)
        nBit = x
      End If
    End If
 Next x
 
 If nBit > 0 Then GetFactBasedMove = nBit
  'okay, nSum is the maximum V sum found and it corresponds to bit 'nBit' of selected pattern 'nChoice'. Before
 'playing let's make sure Teacher does not have a surprise awaiting
 

  '************************************************************************************************************************************
  ' D I S P L A Y S   A L L   K N O W L E D G E
  '************************************************************************************************************************************
  gsLocation = "350"
  'program knowledge
  gsKB = ""
  If UBound(uProgram) < 1 Then GoTo skipprogram
  gsKB = gsKB & String(gnTotalCells * 3, "*") & vbCrLf
  gsKB = gsKB & "Program cells to win..." & vbCrLf
  gsKB = gsKB & String(gnTotalCells * 3, "*") & vbCrLf
  For x = gnTotalCells To 1 Step -1
    gsKB = gsKB & CStr(x) & String(3 - Len(CStr(x)), " ")
  Next x
  gsKB = gsKB & vbCrLf   'indicates bit positions
  gsKB = gsKB & String(gnTotalCells * 3, "*") & vbCrLf
  For x = 1 To UBound(uProgram)
    GetBinaryString sWord, uProgram(x).word
    gsKB = gsKB & sWord & " = " & CStr(uProgram(x).sum) & " (" & uProgram(x).wins & ")" & vbCrLf
  Next x
  gsKB = gsKB & String(gnTotalCells * 3, "*") & vbCrLf
  For x = gnTotalCells To 1 Step -1
    gsKB = gsKB & CStr(nProgramVSum(x)) & String(3 - Len(CStr(nProgramVSum(x))), " ")
  Next x
  gsKB = gsKB & vbCrLf & vbCrLf
skipprogram:
  
  gsLocation = "400"
  'knowledge of teacher
  If UBound(uTeacher) < 1 Then GoTo skipteacher
  gsKB = gsKB & String(gnTotalCells * 3, "*") & vbCrLf
  gsKB = gsKB & "Teacher cells to win..." & vbCrLf
  gsKB = gsKB & String(gnTotalCells * 3, "*") & vbCrLf
  For x = gnTotalCells To 1 Step -1
    gsKB = gsKB & CStr(x) & String(3 - Len(CStr(x)), " ")
  Next x
  gsKB = gsKB & vbCrLf   'indicates bit positions
  gsKB = gsKB & String(gnTotalCells * 3, "*") & vbCrLf
  For x = 1 To UBound(uTeacher)
    GetBinaryString sWord, uTeacher(x).word
    gsKB = gsKB & sWord & " = " & CStr(uTeacher(x).sum) & " (" & uTeacher(x).wins & ")" & vbCrLf
  Next x
  gsKB = gsKB & String(gnTotalCells * 3, "*") & vbCrLf
  For x = gnTotalCells To 1 Step -1
    gsKB = gsKB & CStr(nTeacherVSum(x)) & String(3 - Len(CStr(nTeacherVSum(x))), " ")
  Next x
  gsKB = gsKB & vbCrLf & vbCrLf
skipteacher:

  'displays the data
  gsLocation = "500"
  If frmMain.tabInfo.SelectedItem.Index = 3 Then
    If UBound(uProgram) < 1 Then
      frmMain.txtKB.Text = "No data loaded!"
    Else
      frmMain.txtKB.Text = "'"
      frmMain.ShowKnowledge
      frmMain.txtKB.Text = frmMain.txtKB.Text & gsKB & sStatus
    End If
  End If
  
  gsLocation = "1000"
  Exit Function
MyError:
  gsForm = "AIEngineMod"
  gsProcedure = "GetFactBasedMove"
  ErrorHandler
End Function


'*************************************************************************
'P I C K   R A N D O M   MOVE
'Returns a valid random cell that has not yet 'been chosen.
'This is AIP's alternative when thinking doesn't "cut the mustard",
'an American colloquial expression meaning thinking "just
'won't work!"
'*************************************************************************
Public Function GetRandomMove() As Integer
  Dim lWord As Long
  Dim x As Integer
  Dim nFreeCount As Integer
  ReDim nCells(gnTotalCells) As Integer
  Dim nRandom As Integer
  Dim sMsg As String
  
  On Error GoTo MyError
  
  'find all cells that are empty and place into array nCells()
  'glFreeCells = InvertBits(glAllCells)
  For x = 1 To gnTotalCells
    If ReadBit(glAllCells, x) = False Then
      nFreeCount = nFreeCount + 1
      nCells(nFreeCount) = x
    End If
  Next x
  
  'exit if no free cells remaining
  If nFreeCount = 0 Then
    GetRandomMove = 0
    MsgBox "No more choices..."
    Exit Function
  End If
  
  'select random free cell
  Randomize Timer
  nRandom = (Rnd * nFreeCount) + 1
  If nRandom > nFreeCount Then nRandom = nFreeCount
  GetRandomMove = nCells(nRandom)
    
  Exit Function
MyError:
  gsForm = "AIEngineMod"
  gsProcedure = "PickRandomCell"
  ErrorHandler
End Function


