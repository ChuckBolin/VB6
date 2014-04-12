Attribute VB_Name = "AIEngineMod"
Option Explicit

'*************************************************************************
'AI ENGINE( )
'This sub returns the index within the grid of its move.
'It returns a ZERO if no move is valid.
'*************************************************************************
Public Function AIEngine() As Integer

  AIEngine = PickRandomCell
  
End Function


'Returns a valid random cell that has not yet
'been chosen.
Public Function PickRandomCell() As Integer '<<<<<<<<<<<<<<<<<here
  Dim lWord As Long
  Dim x As Integer
  Dim nFreeCount As Integer
  ReDim nCells(gnTotalCells) As Integer
  Dim nRandom As Integer
  Dim sMsg As String
  
  'find all cells that are empty and place into array nCells()
  'glAllCellsInverted = InvertBits(glAllCells)
  For x = 1 To gnTotalCells
    If ReadBit(glAllCells, x) = False Then
      nFreeCount = nFreeCount + 1
      nCells(nFreeCount) = x
    End If
  Next x
  
  'exit if no free cells remaining
  If nFreeCount = 0 Then
    PickRandomCell = 0
    MsgBox "No more choices..."
    Exit Function
  End If
  
  'select random free cell
  Randomize Timer
  nRandom = (Rnd * nFreeCount) + 1
  If nRandom > nFreeCount Then nRandom = nFreeCount
  PickRandomCell = nCells(nRandom)
  
  'sMsg = "Free count: " & nFreeCount & vbCrLf
  'For x = 1 To nFreeCount
  '  sMsg = sMsg & nCells(x) & vbCrLf
  'Next x
  'MsgBox sMsg
End Function
