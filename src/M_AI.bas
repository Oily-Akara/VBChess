Attribute VB_Name = "M_AI"
' MODULE: M_AI
Option Explicit

Private Const MAX_DEPTH As Integer = 5
Private Const TIME_LIMIT As Double = 5
Private Const CHECK_FREQ As Long = 2000

Private AiStartTime As Double
Private StopSearch As Boolean
Private NodesCount As Long

Private MoveBuffer(20, 150) As Long
Private MoveCount(20) As Long
Public pstPawn(119) As Integer
Public pstKnight(119) As Integer
Public pstBishop(119) As Integer
Public pstRook(119) As Integer
Public pstQueen(119) As Integer
Public pstKingMid(119) As Integer
Public pstKingEnd(119) As Integer

' THE BRAIN (Iterative Deepening + Negamax)
Public Sub MakeComputerMove()
    On Error Resume Next
    If UBound(Board) <> 119 Then
        MsgBox "Memory lost. Resetting..."
        InitBoard
        Exit Sub
    End If
    On Error GoTo 0

    ' PREPARE AI SEARCH
    Dim moves As Collection
    Set moves = GetAllLegalMoves(2)
    
    If moves.count = 0 Then
        CheckGameStatus
        Range("K2").Value = "GAME OVER"
        Exit Sub
    End If
    
    ' CRITICAL: Filter out moves that hang pieces
    Dim safeMoves As Collection
    Set safeMoves = FilterHangingMoves(moves, 2)
    
    ' If all moves hang something, use the full move list (forced)
    If safeMoves.count = 0 Then
        Set safeMoves = moves
        Range("K3").Value = "WARNING: All moves lose material"
    End If
    
    SortMoves safeMoves, 2
    
    Dim bestMove As Long
    Dim currentDepthMove As Long
    Dim score As Long
    Dim alpha As Long, beta As Long
    Dim depth As Integer
    
    ' INITIALIZE TIMER
    AiStartTime = Timer
    StopSearch = False
    NodesCount = 0
    bestMove = safeMoves(1)
    
    ' ITERATIVE DEEPENING LOOP
    For depth = 1 To MAX_DEPTH
        alpha = -30000
        beta = 30000
        Dim bestDepthScore As Long: bestDepthScore = -30000
        Dim moveItem As Variant
        
        For Each moveItem In safeMoves
            Dim moveInt As Long: moveInt = moveItem
            Dim startSq As Integer: startSq = moveInt \ 1000
            Dim endSq As Integer: endSq = moveInt Mod 1000
            Dim cap As Integer
            
            ' Make move
            cap = Board(endSq)
            Board(endSq) = Board(startSq)
            Board(startSq) = EMPTY_SQ
            
            ' Search
            score = -Negamax(depth - 1, -beta, -alpha, 1)
            
            ' Unmake move
            Board(startSq) = Board(endSq)
            Board(endSq) = cap
            
            ' TIMEOUT CHECK
            If StopSearch Then Exit For
            
            If score > bestDepthScore Then
                bestDepthScore = score
                currentDepthMove = moveInt
            End If
            If score > alpha Then alpha = score
        Next moveItem
        
        If StopSearch Then
            Range("K3").Value = "Timeout! Depth " & (depth - 1)
            Exit For
        Else
            bestMove = currentDepthMove
            Range("K3").Value = "Depth " & depth & " | Score: " & bestDepthScore
        End If
        DoEvents
    Next depth

    ' EXECUTE BEST MOVE
    Dim finalStart As Integer, finalEnd As Integer
    Dim pMoving As Integer
    
    finalStart = bestMove \ 1000
    finalEnd = bestMove Mod 1000
    pMoving = Board(finalStart)
    
    ' Update Board
    Board(finalEnd) = Board(finalStart)
    Board(finalStart) = EMPTY_SQ
    
    ' Castling Logic (Black Short Castle)
    If pMoving = 12 And (finalEnd - finalStart) = 2 Then
        Board(26) = Board(28)
        Board(28) = EMPTY_SQ
    End If
    
    ' Castling Logic (Black Long Castle)
    If pMoving = 12 And (finalStart - finalEnd) = 2 Then
        Board(24) = Board(21)
        Board(21) = EMPTY_SQ
    End If
    
    ' Castling Logic (White Short Castle)
    If pMoving = 6 And (finalEnd - finalStart) = 2 Then
        Board(96) = Board(98)
        Board(98) = EMPTY_SQ
    End If
    
    ' Castling Logic (White Long Castle)
    If pMoving = 6 And (finalStart - finalEnd) = 2 Then
        Board(94) = Board(91)
        Board(91) = EMPTY_SQ
    End If
    
    ' Promotion (Auto-Queen)
    If pMoving = 7 And finalEnd >= 91 Then Board(finalEnd) = 11
    If pMoving = 1 And finalEnd <= 28 Then Board(finalEnd) = 5
    
    ' Update History
    moveHistory = moveHistory & finalStart & finalEnd & " "
    
    Turn = 1
    Range("K2").Value = "Turn: White"
    
    RenderBoard
    HighlightChecks
    CheckGameStatus
End Sub

' CRITICAL SAFETY: Filter moves that hang pieces
Function FilterHangingMoves(allMoves As Collection, pColor As Integer) As Collection
    Dim safeMoves As New Collection
    Dim moveItem As Variant
    
    For Each moveItem In allMoves
        Dim moveInt As Long: moveInt = moveItem
        Dim startSq As Integer: startSq = moveInt \ 1000
        Dim endSq As Integer: endSq = moveInt Mod 1000
        
        ' Make move
        Dim cap As Integer
        Dim movingPiece As Integer
        movingPiece = Board(startSq)
        cap = Board(endSq)
        Board(endSq) = Board(startSq)
        Board(startSq) = EMPTY_SQ
        
        ' Check if the piece that just moved is now attacked and undefended
        Dim isSafe As Boolean
        isSafe = True
        
        If IsSquareAttacked(endSq, 3 - pColor) Then
            ' The piece is attacked - check if it's defended
            Dim pieceValue As Integer
            pieceValue = GetPieceValue(movingPiece)
            
            ' If we captured something of equal or greater value, it's okay
            Dim captureValue As Integer
            captureValue = GetPieceValue(cap)
            
            If captureValue < pieceValue Then
                ' We might be hanging - check if defended
                If Not IsSquareDefended(endSq, pColor) Then
                    isSafe = False ' HANGING!
                End If
            End If
        End If
        
        ' Unmake move
        Board(startSq) = Board(endSq)
        Board(endSq) = cap
        
        If isSafe Then
            safeMoves.Add moveInt
        End If
    Next moveItem
    
    Set FilterHangingMoves = safeMoves
End Function



' Check if a square is defended by friendly pieces
Function IsSquareDefended(sq As Integer, byColor As Integer) As Boolean
    Dim defenderSq As Integer
    
    For defenderSq = 21 To 98
        If Board(defenderSq) <> EMPTY_SQ And Board(defenderSq) <> OFF_BOARD Then
            Dim defenderColor As Integer
            If Board(defenderSq) <= 6 Then defenderColor = 1 Else defenderColor = 2
            
            If defenderColor = byColor Then
                ' Temporarily remove the piece at sq to test if defender can reach it
                Dim tempPiece As Integer
                tempPiece = Board(sq)
                Board(sq) = EMPTY_SQ
                
                Dim canDefend As Boolean
                canDefend = IsLegalMove(defenderSq, sq)
                
                Board(sq) = tempPiece
                
                If canDefend Then
                    IsSquareDefended = True
                    Exit Function
                End If
            End If
        End If
    Next defenderSq
    
    IsSquareDefended = False
End Function

' Get material value of a piece
Function GetPieceValue(piece As Integer) As Integer
    Select Case piece
        Case 1, 7: GetPieceValue = 100  ' Pawn
        Case 2, 8: GetPieceValue = 320  ' Knight
        Case 3, 9: GetPieceValue = 330  ' Bishop
        Case 4, 10: GetPieceValue = 500 ' Rook
        Case 5, 11: GetPieceValue = 900 ' Queen
        Case 6, 12: GetPieceValue = 0   ' King (don't count in trades)
        Case Else: GetPieceValue = 0
    End Select
End Function

' THE ALGORITHM (Negamax + Quiescence)
Function Negamax(depth As Integer, alpha As Long, beta As Long, pColor As Integer) As Long
    NodesCount = NodesCount + 1
    
    If (NodesCount Mod CHECK_FREQ = 0) Then
        DoEvents
        If Timer - AiStartTime > TIME_LIMIT Then
            StopSearch = True
        End If
    End If
    
    If StopSearch Then
        Negamax = 0
        Exit Function
    End If
    
    If depth <= 0 Then
        Negamax = Quiescence(alpha, beta, pColor)
        Exit Function
    End If
    
    Dim moves As Collection
    Set moves = GetAllLegalMoves(pColor)
    
    If moves.count = 0 Then
        Negamax = -20000 + depth
        Exit Function
    End If
    
    SortMoves moves, pColor
    
    Dim moveInt As Long
    Dim startSq As Integer, endSq As Integer
    Dim cap As Integer
    Dim score As Long
    Dim moveItem As Variant
    
    For Each moveItem In moves
        moveInt = moveItem
        startSq = moveInt \ 1000
        endSq = moveInt Mod 1000
        
        cap = Board(endSq)
        Board(endSq) = Board(startSq)
        Board(startSq) = EMPTY_SQ
        
        score = -Negamax(depth - 1, -beta, -alpha, 3 - pColor)
        
        Board(startSq) = Board(endSq)
        Board(endSq) = cap
        
        If StopSearch Then Exit Function
        
        If score >= beta Then
            Negamax = beta
            Exit Function
        End If
        If score > alpha Then alpha = score
    Next moveItem
    
    Negamax = alpha
End Function

Function Quiescence(alpha As Long, beta As Long, pColor As Integer) As Long
    NodesCount = NodesCount + 1
    If (NodesCount Mod CHECK_FREQ = 0) Then
        DoEvents
        If Timer - AiStartTime > TIME_LIMIT Then StopSearch = True
    End If
    If StopSearch Then
        Quiescence = 0
        Exit Function
    End If
    
    Dim stand_pat As Long
    stand_pat = EvaluatePosition()
    
    If stand_pat >= beta Then
        Quiescence = beta
        Exit Function
    End If
    If alpha < stand_pat Then alpha = stand_pat
    
    Dim moves As Collection
    Set moves = GetAllCaptures(pColor)
    SortMoves moves, pColor
    
    Dim moveInt As Long
    Dim startSq As Integer, endSq As Integer
    Dim cap As Integer
    Dim score As Long
    Dim moveItem As Variant
    
    For Each moveItem In moves
        moveInt = moveItem
        startSq = moveInt \ 1000
        endSq = moveInt Mod 1000
        
        cap = Board(endSq)
        Board(endSq) = Board(startSq)
        Board(startSq) = EMPTY_SQ
        
        score = -Quiescence(-beta, -alpha, 3 - pColor)
        
        Board(startSq) = Board(endSq)
        Board(endSq) = cap
        
        If StopSearch Then Exit Function
        
        If score >= beta Then
            Quiescence = beta
            Exit Function
        End If
        If score > alpha Then alpha = score
    Next moveItem
    
    Quiescence = alpha
End Function

Function GetAllLegalMoves(pColor As Integer) As Collection
    Dim moves As New Collection
    Dim tempMoves As New Collection
    Dim startSq As Integer, endSq As Integer
    Dim piece As Integer, Target As Integer
    
    For startSq = 21 To 98
        piece = Board(startSq)
        If piece <> OFF_BOARD And piece <> EMPTY_SQ Then
            If (piece <= 6 And pColor = 1) Or (piece > 6 And pColor = 2) Then
                For endSq = 21 To 98
                    Target = Board(endSq)
                    If Target <> OFF_BOARD Then
                        If IsLegalMove(startSq, endSq) Then
                            tempMoves.Add (CLng(startSq) * 1000) + endSq
                        End If
                    End If
                Next endSq
            End If
        End If
    Next startSq
    
    Dim moveItem As Variant
    For Each moveItem In tempMoves
        Dim moveInt As Long: moveInt = moveItem
        startSq = moveInt \ 1000
        endSq = moveInt Mod 1000
        
        Dim cap As Integer
        cap = Board(endSq)
        Board(endSq) = Board(startSq)
        Board(startSq) = EMPTY_SQ
        
        If Not IsKingInCheck(pColor) Then
            moves.Add moveInt
        End If
        
        Board(startSq) = Board(endSq)
        Board(endSq) = cap
    Next moveItem
    
    Set GetAllLegalMoves = moves
End Function

Function GetAllCaptures(pColor As Integer) As Collection
    Dim moves As New Collection
    Dim startSq As Integer, endSq As Integer
    Dim piece As Integer, Target As Integer
    
    For startSq = 21 To 98
        piece = Board(startSq)
        If piece <> OFF_BOARD And piece <> EMPTY_SQ Then
            If (piece <= 6 And pColor = 1) Or (piece > 6 And pColor = 2) Then
                For endSq = 21 To 98
                    Target = Board(endSq)
                    If Target <> OFF_BOARD And Target <> EMPTY_SQ Then
                        Dim targetColor As Integer
                        If Target <= 6 Then targetColor = 1 Else targetColor = 2
                        If targetColor <> pColor Then
                            If IsLegalMove(startSq, endSq) Then
                                moves.Add (CLng(startSq) * 1000) + endSq
                            End If
                        End If
                    End If
                Next endSq
            End If
        End If
    Next startSq
    Set GetAllCaptures = moves
End Function

Public Function EvaluatePosition() As Long
    Dim i As Integer, p As Integer
    Dim score As Long
    Dim mirrorSq As Integer
    
    score = 0
    
    For i = 21 To 98
        p = Board(i)
        If p <> EMPTY_SQ And p <> OFF_BOARD Then
            mirrorSq = 119 - i
            
            Select Case p
                Case 1: score = score + 100 + pstPawn(i)
                Case 2: score = score + 320 + pstKnight(i)
                Case 3: score = score + 330 + pstBishop(i)
                Case 4: score = score + 500 + pstRook(i)
                Case 5: score = score + 900 + pstQueen(i)
                Case 6: score = score + 20000 + pstKingMid(i)
                Case 7: score = score - (100 + pstPawn(mirrorSq))
                Case 8: score = score - (320 + pstKnight(mirrorSq))
                Case 9: score = score - (330 + pstBishop(mirrorSq))
                Case 10: score = score - (500 + pstRook(mirrorSq))
                Case 11: score = score - (900 + pstQueen(mirrorSq))
                Case 12: score = score - (20000 + pstKingMid(mirrorSq))
            End Select
        End If
    Next i
    
    EvaluatePosition = score
End Function

Public Function IsKingInCheck(kingColor As Integer) As Boolean
    Dim kingSq As Integer
    Dim i As Integer
    
    For i = 21 To 98
        If kingColor = 1 And Board(i) = 6 Then kingSq = i: Exit For
        If kingColor = 2 And Board(i) = 12 Then kingSq = i: Exit For
    Next i
    
    If kingSq = 0 Then
        IsKingInCheck = True
        Exit Function
    End If
    
    Dim attackSq As Integer
    For attackSq = 21 To 98
        If Board(attackSq) <> EMPTY_SQ And Board(attackSq) <> OFF_BOARD Then
            Dim attackColor As Integer
            If Board(attackSq) <= 6 Then attackColor = 1 Else attackColor = 2
            
            If attackColor <> kingColor Then
                If IsLegalMove(attackSq, kingSq) Then
                    IsKingInCheck = True
                    Exit Function
                End If
            End If
        End If
    Next attackSq
    
    IsKingInCheck = False
End Function
