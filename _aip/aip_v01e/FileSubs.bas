Attribute VB_Name = "FileSubs"
Option Explicit

Public Function WriteFile(sFileName As String) As Boolean
  Dim nFile As Integer 'file handle
  Dim x, y As Integer
  Dim bFound As Boolean
  Dim nWins As Integer
  Dim lWord As Long
  Dim sABS As String
  
  Dim sFile As String
  
  On Error GoTo MyError:
 
  'extract filename from complete path and filename
  For x = Len(sFileName) To 1 Step -1
    If InStr(x, sFileName, "\") Then
      sFile = Mid(sFileName, x + 1)
      Exit For
    End If
  Next x
 
  nFile = FreeFile
  Open sFileName For Output As nFile
    Print #nFile, "'**********************************"
    Print #nFile, "'" & sFile
    Print #nFile, "'Date:" & Date
    Print #nFile, "'**********************************"
    Print #nFile, ""
    Print #nFile, "'system variables"
    Print #nFile, "GAME.NAME=" & gsGameName
    Print #nFile, "GAME.TYPE=" & CStr(gnGameType)
    Print #nFile, "GAME.ROWS=" & CStr(gnRows)
    Print #nFile, "GAME.COLS=" & CStr(gnCols)
    Print #nFile, "GAME.TEACHER.SYMBOL=" & gsTeacherSymbol
    Print #nFile, "GAME.TEACHER.VALUE=" & CStr(gnTeacherValue)
    Print #nFile, "GAME.PROGRAM.SYMBOL=" & gsProgramSymbol
    Print #nFile, "GAME.PROGRAM.VALUE=" & CStr(gnProgramValue)
    Print #nFile, "GAME.GOFIRST=" & CStr(gnGoFirst)
    Print #nFile, "GAME.CELL.COLOR=" & CStr(glCellColor)
    Print #nFile, "GAME.CELL.SELECTEDCOLOR=" & CStr(glCellSelectedColor)
    Print #nFile, ""
    Print #nFile, "'win-loss history"
    Print #nFile, "GAME.PROGRAM.WIN=" & CStr(gnProgramWins)
    Print #nFile, "GAME.PROGRAM.LOSS=" & CStr(gnProgramLosses)
    Print #nFile, "GAME.PROGRAM.TIE=" & CStr(gnProgramTies)
    Print #nFile, ""
    Print #nFile, "'winning patterns"
    For x = 1 To UBound(uABS)
      sABS = ""
      GetABSString uABS(x).word, sABS, uABS(x).wins
      Print #nFile, sABS
    Next x
  
  Close nFile


  Exit Function

MyError:
  gsForm = "Module FileSubs"
  gsProcedure = "WriteFile"
  ErrorHandler
End Function

'reads game variable data from file
Public Function ReadFile(sFileName As String) As Boolean
  Dim sIn As String 'holds line read from file
  Dim nFile As Integer 'file handle
  Dim x, y As Integer
  Dim lWord As Long
  Dim bFound As Boolean
  Dim nWins As Integer
  Dim sNew As String
  
  On Error GoTo MyError:
  ReadFile = False
  gbFilenameExists = False
  gnABSTotal = 0 'zeros number of winning patterns in array
  ReDim uABS(0) As Pattern
  
  'file exists, open and read one line at a time
  nFile = FreeFile 'gets next file handle
  Open sFileName For Input As nFile
    Do
    Line Input #1, sIn
    sIn = UCase(LTrim(sIn)) 'forces string to uppercase for parsing next step
                                        'trims left whitespaces
    'RemoveWhitespace sIn, sNew
    'sIn = sNew
    
    'reads only lines that are not empty
    If Len(sIn) > 0 Then
      
      'reads lines
      If Left(sIn, 1) = "'" Or Left(sIn, 2) = "//" Then
        'do nothing...comment found (VB or C format symbol)
      Else  'important variable here
        
        'system variable located here
        If InStr(1, sIn, "GAME.") Then
    
          x = 0: x = InStr(1, sIn, "=") 'gets positions of '=' sign
          If InStr(1, sIn, "GAME.NAME") Then gsGameName = Mid(sIn, x + 1)
          If InStr(1, sIn, "GAME.TYPE") Then gnGameType = CInt(Mid(sIn, x + 1))
          If InStr(1, sIn, "GAME.ROWS") Then gnRows = CInt(Mid(sIn, x + 1))
          If InStr(1, sIn, "GAME.COLS") Then gnCols = CInt(Mid(sIn, x + 1))
          If InStr(1, sIn, "GAME.TEACHER.SYMBOL") Then gsTeacherSymbol = Mid(sIn, x + 1)
          If InStr(1, sIn, "GAME.TEACHER.VALUE") Then gnTeacherValue = CInt(Mid(sIn, x + 1))
          If InStr(1, sIn, "GAME.PROGRAM.SYMBOL") Then gsProgramSymbol = Mid(sIn, x + 1)
          If InStr(1, sIn, "GAME.PROGRAM.VALUE") Then gnProgramValue = CInt(Mid(sIn, x + 1))
          If InStr(1, sIn, "GAME.GOFIRST") Then gnGoFirst = CInt(Mid(sIn, x + 1))
          If InStr(1, sIn, "GAME.PROGRAM.WIN") Then gnProgramWins = CInt(Mid(sIn, x + 1))
          If InStr(1, sIn, "GAME.PROGRAM.LOSS") Then gnProgramLosses = CInt(Mid(sIn, x + 1))
          If InStr(1, sIn, "GAME.PROGRAM.TIE") Then gnProgramTies = CInt(Mid(sIn, x + 1))
          If InStr(1, sIn, "GAME.CELL.COLOR") Then glCellColor = CLng(Mid(sIn, x + 1))
          If InStr(1, sIn, "GAME.CELL.SELECTEDCOLOR") Then glCellSelectedColor = CLng(Mid(sIn, x + 1))
        End If
        
        'reads absolute win patterns and loads into uABS( )
        If InStr(1, sIn, "ABS(") Then
          lWord = ConvertABS2Bits(sIn)
          nWins = GetABSWins(sIn)
          
          'additional win patterns after initial winning pattern
          If gnABSTotal > 0 Then
            bFound = False
            For y = 1 To gnABSTotal
               If uABS(y).word = lWord Then 'pattern found in uABS()
                 uABS(y).wins = uABS(y).wins + nWins
                 bFound = True
                 Exit For
              End If
            Next y
            If bFound = False Then 'pattern not found in array, must add
              gnABSTotal = gnABSTotal + 1
              ReDim Preserve uABS(gnABSTotal) As Pattern
              uABS(gnABSTotal).word = lWord
              uABS(y).wins = uABS(y).wins + nWins
            End If
          End If
          
          'very first win pattern to be loaded
          If gnABSTotal = 0 Then
            gnABSTotal = gnABSTotal + 1
            ReDim uABS(gnABSTotal) As Pattern
            uABS(gnABSTotal).word = lWord
            uABS(gnABSTotal).wins = nWins
          End If
        End If
    End If
  End If
  Loop Until EOF(nFile)
  
  Close nFile
  ReadFile = True
  gbFilenameExists = True
  Exit Function

MyError:
  gsForm = "Module FileSubs"
  gsProcedure = "ReadFile"
  ErrorHandler
End Function

