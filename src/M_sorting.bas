Attribute VB_Name = "M_sorting"
' Module: M_Sorting
Option Explicit

' Bubble Sort to order moves (Best moves first)
' OPTIMIZED: Calculates scores once, then sorts.
Public Sub SortMoves(ByRef moves As Collection, pColor As Integer)
    Dim n As Integer
    n = moves.count
    If n < 2 Then Exit Sub
    
    ' Arrays to hold moves and their scores
    Dim mArr() As Long
    Dim sArr() As Integer
    Dim i As Integer, j As Integer
    Dim tempMove As Long, tempScore As Integer
    
    ReDim mArr(1 To n)
    ReDim sArr(1 To n)
    
    ' 1. Convert Collection to Array AND Pre-calculate Scores
    ' This is much faster than calculating inside the sort loop
    For i = 1 To n
        mArr(i) = moves(i)
        sArr(i) = ScoreMove(mArr(i)) ' No pColor needed
    Next i
    
    ' 2. Bubble Sort (Sorts both arrays based on Score)
    For i = 1 To n - 1
        For j = i + 1 To n
            ' If Move J has a higher score than Move I, swap them
            If sArr(j) > sArr(i) Then
                ' Swap Scores
                tempScore = sArr(i)
                sArr(i) = sArr(j)
                sArr(j) = tempScore
                
                ' Swap Moves
                tempMove = mArr(i)
                mArr(i) = mArr(j)
                mArr(j) = tempMove
            End If
        Next j
    Next i
    
    ' 3. Put sorted moves back into Collection
    Set moves = New Collection
    For i = 1 To n
        moves.Add mArr(i)
    Next i
End Sub

' Returns a Score for a move (MVV-LVA logic)
' High score = Capture high value piece with low value piece
Public Function ScoreMove(moveInt As Long) As Integer
    ' Use On Error Resume Next for speed in tight loops,
    ' but strictly we assume Board exists here.
    
    Dim startSq As Integer, endSq As Integer
    Dim victim As Integer, attacker As Integer
    Dim vVal As Integer, aVal As Integer
    
    startSq = moveInt \ 1000
    endSq = moveInt Mod 1000
    
    ' Safety Check
    If endSq < 0 Or endSq > 119 Then
        ScoreMove = 0
        Exit Function
    End If
    
    victim = Board(endSq)
    
    ' If not a capture, return 0 (Sorted last)
    If victim = 0 Then
        ScoreMove = 0
        Exit Function
    End If
    
    attacker = Board(startSq)
    
    ' Get Values
    vVal = GetPieceValue(victim)
    aVal = GetPieceValue(attacker)
    
    ' MVV-LVA Formula:
    ' Most Valuable Victim - Least Valuable Attacker
    ' Example: Pawn(1) eats Queen(9) -> 100 + 90 - 1 = 189 (High priority)
    ' Example: Queen(9) eats Pawn(1) -> 100 + 10 - 9 = 101 (Low priority)
    ScoreMove = 100 + (vVal * 10) - aVal
End Function

' Helper to get values for sorting
Private Function GetPieceValue(piece As Integer) As Integer
    Select Case piece
        Case 1, 7:   GetPieceValue = 1      ' Pawn
        Case 2, 8:   GetPieceValue = 3      ' Knight
        Case 3, 9:   GetPieceValue = 3      ' Bishop
        Case 4, 10:  GetPieceValue = 5      ' Rook
        Case 5, 11:  GetPieceValue = 9      ' Queen
        Case 6, 12:  GetPieceValue = 100    ' King (Shouldn't happen often, but highest priority)
        Case Else:   GetPieceValue = 0
    End Select
End Function
