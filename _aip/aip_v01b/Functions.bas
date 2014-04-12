Attribute VB_Name = "Functions"
Option Explicit

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
  gsForm = "Module Functions"
  gsProcedure = "CountChar"
  ErrorHandler
End Function

'*******************************************************************************
'GetBitW( )
'Returns the boolean state of a particular bit in a long number.
'*******************************************************************************
Public Function GetBitW(lNum As Long, nBit As Integer) As Boolean
  On Error GoTo MyError:
  If nBit < 0 Or nBit > 30 Then Exit Function
  If lNum < 0 Or lNum > 2147483647 Then Exit Function
  GetBitW = (2 ^ nBit) And lNum
  Exit Function
MyError:
  gsForm = "Functions"
  gsProcedure = "GetBitW"
  ErrorHandler
End Function

'*******************************************************************************
'SetBitW( )
'Sets a bit within a long number if not sent. Returns the result.
'*******************************************************************************
Public Function SetBitW(lNum As Long, nBit As Integer) As Long
  On Error GoTo MyError:
  If nBit < 0 Or nBit > 30 Then Exit Function
  If lNum < 0 Or lNum > 2147483647 Then Exit Function
  SetBitW = (2 ^ nBit) Or lNum
  Exit Function
MyError:
  gsForm = "Functions"
  gsProcedure = "SetBitW"
  ErrorHandler
End Function

'*******************************************************************************
'SetPosW( )
'Sets a bit corresponding to the position within the grid. Returns new
'long value
'*******************************************************************************
Public Function SetPosW(lNum As Long, nPos As Integer) As Long
  On Error GoTo MyError:
  If nPos < 0 Or nPos > 31 Then Exit Function
  If lNum < 0 Or lNum > 2147483647 Then Exit Function
  SetPosW = (2 ^ (32 - nPos)) Or lNum
  
  Exit Function
MyError:
  gsForm = "Functions"
  gsProcedure = "SetPosW"
  ErrorHandler
End Function

'*******************************************************************************
'GetPosW( )
'Gets a bit indicating if the desired position in grid is empty..
'*******************************************************************************
Public Function GetPosW(lNum As Long, nPos As Integer) As Boolean
  On Error GoTo MyError:
  If nPos < 0 Or nPos > 31 Then Exit Function
  If lNum < 0 Or lNum > 2147483647 Then Exit Function
  GetPosW = (2 ^ (32 - nPos)) Or lNum
  
  Exit Function
MyError:
  gsForm = "Functions"
  gsProcedure = "GetPosW"
  ErrorHandler
End Function

'*******************************************************************************
'GetCellPos( )
'Gets a value corresponding to the cell position
'*******************************************************************************
Public Function GetCellPos(nRow As Integer, nCol As Integer) As Integer

  On Error GoTo MyError:
  GetCellPos = 0
  If nRow < 1 Or nCol < 1 Or nRow > gnRows Or nCol > gnCols Then Exit Function
  GetCellPos = ((nRow - 1) * gnCols) + nCol
    
  Exit Function
MyError:
  gsForm = "Functions"
  gsProcedure = "GetCellPos"
  ErrorHandler
End Function


