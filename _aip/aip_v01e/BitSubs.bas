Attribute VB_Name = "BitSubs"
'*************************************************************************************************************
' AIS - (Artificial Intelligence Scripting)      Game variables and winning patterns are stored in
'       a collection of KB (knowledge base) files in a readable format using scripting.  For example:
'       ABS(1,1);(2,2);(3,3)=2    ABS means absolute.  The win was constructed from playing a
'       piece in three cells indentified by the three coordinate pairs. In fact, the '=2' means this
'       resulted in two wins thus far.
'       Games start be reading this file and end by writing to this file. One file for each game type.
'
'BIT Pattern - AIS is inadequate for the program.  Instead all wins and status of pieces in a
'       grid are stored in a 32 bit long integer.  The MSD bit is a sign bit and is therefore not used.
'       For example:
'       A 3x3 grid has a total of 9 cells. Therefore, a 9 bit number is used to represent a win.
'       If the winning pattern has a piece in a cell, the corresponding bit is equal to a '1'.
'
' BitSubs - Contains all subs/functions required to convert AIS to Bits and convert Bits to AIS.  A winning
'       ABS pattern is converted to a REL (relative) position. REL pos is defined by moving
'       the winning pattern as far to the left and up as possible.
'       The REL pattern is then rotated 3 times and flipped 4 times to produce 7 more potential
'       ABS patterns stored in Bit format.
'****************************************************************************************************************

'**************************************************************
'K E Y   V A R I A B L E S
'**************************************************************
'create module specific variables
Option Explicit

'the following variables are required for this module
'user defined variable storing the winning pattern and the number of wins
Public Type Pattern
  word  As Long '32 bit can hold a 5x5 array
  wins As Integer 'number of wins attributed to this pattern
End Type

'this array stores winning patterns -actual and derived
Public uABS() As Pattern 'stores all winning or derived potential winning patterns
Public gnABSTotal As Integer 'stores total number of ABS values in uABS

Public glTeacher As Long  'stores moves made by Teacher pieces
Public glProgram As Long 'stores moves made by Program pieces
Public glAllCells As Long  'stores combined Teacher and Program pieces
Public glAllCellsInverted As Long 'this is an inverted copy of uAllCells

'*************************************************************
' S U B S
'*************************************************************

'*************************************************************
' U P D A T E A B S ( )
'*************************************************************
'passes a long value add to uABS array
Public Sub UpdateABS(lWord As Long)
  Dim x, y As Integer
  Dim bFound As Boolean
  
  'additional win patterns after initial winning pattern
  If gnABSTotal > 0 Then
      bFound = False
      For y = 1 To gnABSTotal
        If uABS(y).word = lWord Then 'pattern found in uABS()
          uABS(y).wins = uABS(y).wins + 1
          bFound = True
          Exit For
        End If
      Next y
      If bFound = False Then 'pattern not found in array, must add
        gnABSTotal = gnABSTotal + 1
        ReDim Preserve uABS(gnABSTotal) As Pattern
        uABS(gnABSTotal).word = lWord
        uABS(y).wins = 1
      End If
  End If
  
    'very first win pattern to be loaded
    If gnABSTotal = 0 Then
      gnABSTotal = gnABSTotal + 1
      ReDim uABS(gnABSTotal) As Pattern
      uABS(gnABSTotal).word = lWord
      uABS(gnABSTotal).wins = 1
    End If
  
End Sub

'*****************************************************************************
'GETABSSTRING
' Given a long variable with bits set, get the equivalent ABS string
'*****************************************************************************
Public Sub GetABSString(lNum As Long, sABS As String, nWins As Integer)
  Dim x As Integer
  Dim nRow As Integer
  Dim nCol As Integer
 
  On Error GoTo MyError
  
  'constructs ABS string from long integer
  sABS = "ABS"
  For x = 1 To gnTotalCells
    If (lNum And 2 ^ (x - 1)) Then
      nRow = GetCellRow(x)
      nCol = GetCellCol(x)
      sABS = sABS & "(" & CStr(nRow) & "," & CStr(nCol) & ");"
    End If
   Next x
    
   'removes last
   sABS = Left(sABS, Len(sABS) - 1) & "=" & CStr(nWins)
    
  Exit Sub
MyError:
  gsForm = "Subs"
  gsProcedure = "GetABSString"
  ErrorHandler
End Sub

Public Sub RemoveWhitespace(sOld As String, sNew As String)
  Dim x As Integer
  If Len(sOld) < 1 Then Exit Sub
  
  For x = 1 To Len(sOld)
    If Mid(sOld, x, 1) <> Chr(32) And Mid(sOld, x, 1) <> Chr(9) Then
      sNew = Mid(sOld, x, 1)
    End If
  Next x
  
End Sub

'*************************************************************
' F U N C T I O N S
' 1) GetRelativePattern - Converts absolute to relative
' 2) GetRotatedPattern - Creates 3 rotated patterns
' 3) GetFlippedPattern  - Creates 4 flipped patterns
' 4) ConvertAIS2Bits     - Converts AIS to bits
' 5) GetABSRow
' 6) GetABSCol
' 7) CountChar             - Returns # of a character
' 8) GetCellRow
' 9) GetCellCol
'10) GetABSWins
'11) GetCellIndex
'*************************************************************

'This normalizes a winning ABS pattern by converting it to a
'relative pattern. This relative pattern is used by program to
'recognize similarities among patterns (i.e. horizontal lines)
Public Function GetRelativePattern(lPattern As Long) As Long

End Function


'This rotates a pattern by 90, 180 or 270 degrees clockwise and
'returns the result. Usually this is called 3 times per REL pattern
Public Function GetRotatedPattern(lPattern As Long, nAngle As Integer) As Long

End Function

'This takes a relative pattern and flips it in four angles of 0, 90, 180 and 270
'clockwise.  Usually called 4 times per REL pattern
Public Function GetFlippedPattern(lPattern As Long, nAngle As Integer) As Long

End Function

'*************************************************************************
'ConvertABS2Bits
'Converts an AIS string (i.e. ABS(x,x);(y,y)=z to a long
'*************************************************************************
Public Function ConvertABS2Bits(sABS As String) As Long
  Dim y As Integer
  Dim nIndex As Integer
  Dim nRow As Integer
  Dim nCol As Integer
  Dim lWord As Long 'stores ABS number in binary format matching pattern
  
  On Error GoTo MyError:
  ConvertABS2Bits = 0
  
  'extract bit values
  For y = 1 To CountChar(sABS, "(")
    nRow = GetABSRow(sABS, y)
    nCol = GetABSCol(sABS, y)
    nIndex = GetCellIndex(nRow, nCol)
    lWord = lWord + (2 ^ (nIndex - 1))
  Next y
  
  'returns this value
  ConvertABS2Bits = lWord
  Exit Function
  
MyError:
  gsForm = "Module BitSubs"
  gsProcedure = "ConvertABS2Bits"
  ErrorHandler

End Function

'**********************************************************************************
'GetABSRow
'Given an ABS(  );(  );(  )=x string, returns row from the pair indicated by
'nIndex.  The first pair is nIndex=1
'**********************************************************************************
Public Function GetABSRow(sABS As String, nIndex As Integer) As Integer
  Dim x, y, z As Integer
  Dim nRow As Integer
  Dim nNumParen As Integer
  Dim lWord As Long 'stores ABS number in binary format matching pattern
  
  On Error GoTo MyError:
  GetABSRow = 0
  
  'make sure that the index is within legal range
  nNumParen = CountChar(sABS, "(")
  If nIndex < 1 Or nIndex > nNumParen Then Exit Function
  
  'searchs for correct '('
  y = 1 'initial search position
  For x = 1 To CountChar(sABS, "(")
    y = InStr(y, sABS, "(")  'tracks position of '(' character
    'MsgBox "Row " & nIndex & " " & y
    
    If x = nIndex Then        'correct '(' found
      z = InStr(y, sABS, ",") 'looks for position of comma
      
      'row value is now between positions y and z, or (   ,
      nRow = CInt(Mid(sABS, y + 1, z - y - 1))
      GetABSRow = nRow
      
    End If
    y = y + 1
  Next x
  
  Exit Function
  
MyError:
  gsForm = "Module BitSubs"
  gsProcedure = "GetABSRow"
  ErrorHandler
End Function

'Given ABS string, returns column.
Public Function GetABSCol(sABS As String, nIndex As Integer) As Integer
  Dim x, y, z As Integer
  Dim nCol As Integer
  Dim nNumParen As Integer
  Dim lWord As Long 'stores ABS number in binary format matching pattern
  
  On Error GoTo MyError:
  GetABSCol = 0
  
  'make sure that the index is within legal range
  nNumParen = CountChar(sABS, "(")
  If nIndex < 1 Or nIndex > nNumParen Then Exit Function
  
  'searchs for correct '('
  y = 1 'initial search position
  For x = 1 To CountChar(sABS, ",")
    y = InStr(y, sABS, ",")  'tracks position of comma
    If x = nIndex Then        'correct ',' found
      z = InStr(y, sABS, ")") 'looks for position of ')'
      
      'col value is now between positions y and z, or ,   )
      nCol = CInt(Mid(sABS, y + 1, z - y - 1))
      GetABSCol = nCol
    End If
    y = y + 1
  Next x
  
  Exit Function
  
MyError:
  gsForm = "Module BitSubs"
  gsProcedure = "GetABSCol"
  ErrorHandler

End Function

'*****************************************************************************
'COUNTCHAR( )
'counts the number of occurences of a particular
'character found in a string and returns the value
'*****************************************************************************
Public Function CountChar(sIn As String, sChar As String) As Integer
  Dim x As Integer
  Dim nCount As Integer
  On Error GoTo MyError
  
  If Len(sIn) < 1 Then Exit Function
  If Len(sChar) < 1 Then Exit Function
  sIn = UCase(sIn)
  sChar = UCase(sChar)
  For x = 1 To Len(sIn)
    If Mid(sIn, x, 1) = sChar Then nCount = nCount + 1
  Next x
  CountChar = nCount
  Exit Function
  
MyError:
  gsForm = "Module BitSubs"
  gsProcedure = "CountChar"
  ErrorHandler
End Function

'*******************************************************************************
'GetCellRow( )
'Gets a value corresponding to the cell position
'*******************************************************************************
Public Function GetCellRow(nIndex As Integer) As Integer
  Dim nNum As Integer
  
  On Error GoTo MyError:
  nNum = nIndex \ gnRows
  If nIndex > gnRows * nNum Then nNum = nNum + 1
  GetCellRow = nNum
    
  Exit Function
MyError:
  gsForm = "BitSubs"
  gsProcedure = "GetCellRow"
  ErrorHandler
End Function

'*******************************************************************************
'GetCellCol( )
'Gets a value corresponding to the cell position
'*******************************************************************************
Public Function GetCellCol(nIndex As Integer) As Integer

  Dim nNum As Integer
  Dim nRow As Integer
  
  On Error GoTo MyError:
  nRow = nIndex \ gnRows
  If nIndex > gnRows * nRow Then nRow = nRow + 1
  GetCellCol = nIndex - (nRow - 1) * 3
    
  Exit Function
MyError:
  gsForm = "BitSubs"
  gsProcedure = "GetCellCol"
  ErrorHandler
End Function

Public Function GetABSWins(sABS As String) As Integer

  Dim y As Integer
  Dim nWins As Integer
  
  On Error GoTo MyError:
  GetABSWins = 0
  
  'finds position of '=' in string sABS
  y = InStr(1, sABS, "=")
  If y < 1 Then Exit Function
  
  'everything to the left of '=' is wins total
  nWins = CInt(Mid(sABS, y + 1))
  GetABSWins = nWins
  Exit Function
  
MyError:
  gsForm = "Module BitSubs"
  gsProcedure = "GetABSWins"
  ErrorHandler

End Function


Public Function GetCellIndex(nRow As Integer, nCol As Integer) As Integer
  GetCellIndex = 0
  If nRow < 1 Or nRow > gnRows Then Exit Function
  If nCol < 1 Or nCol > gnCols Then Exit Function
  GetCellIndex = (nRow - 1) * gnCols + nCol
End Function



'Sets a bit in a long integer. Note: The MSB is bit 1
Public Function SetBit(lWord As Long, nBit As Integer) As Long
  Dim lNum As Long
  
  If nBit < 1 Or nBit > 31 Then Exit Function
  lNum = lWord
  lNum = lNum Xor (2 ^ (nBit - 1))
  SetBit = lNum
 ' MsgBox "Setbit: " & lWord & " " & lNum
End Function

'Clears a bit in a long integer to a '0'. The MSB is bit 1
Public Function ClearBit(lWord As Long, nBit As Integer) As Long
  If nBit < 1 Or nBit > 31 Then Exit Function
  lWord = lWord Xor (2 ^ (nBit - 1))
  ClearBit = lWord
End Function

Public Function ReadBit(lWord As Long, nBit As Integer) As Boolean
  ReadBit = False
  If nBit < 1 Or nBit > 31 Then Exit Function
  If lWord And 2 ^ (nBit - 1) Then ReadBit = True
End Function

'inverts all bits within long number up to gnTotalCells
Public Function InvertBits(lWord As Long) As Long
  Dim x As Integer
  Dim lReturn As Long
  
  For x = 1 To gnTotalCells
    'MsgBox ReadBit(lWord, x) & "  " & lWord
    If ReadBit(lWord, x) = True Then
      lReturn = ClearBit(lWord, x)
    Else
      lReturn = SetBit(lWord, x)
    End If
  Next x
  InvertBits = lReturn
  
End Function
