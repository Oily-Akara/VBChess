Attribute VB_Name = "M_AI"
' MODULE: M_AI
Option Explicit


Private Const MAX_DEPTH As Integer = 4
Private Const TIME_LIMIT As Double = 5
Private Const CHECK_FREQ As Long = 2000


Private AiStartTime As Double
Private StopSearch As Boolean
Private NodesCount As Long

Private MoveBuffer(20, 150) As Long
Private moveCount(20) As Long
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

    
    Dim bookMove As Long
    
    bookMove = GetBookMove(Turn)
    
    If bookMove > 0 Then
        Dim bStart As Integer, bEnd As Integer
        bStart = bookMove \ 1000
        bEnd = bookMove Mod 1000
        
      
        Board(bEnd) = Board(bStart)
        Board(bStart) = EMPTY_SQ
     
        moveHistory = moveHistory & bStart & bEnd & " "
        
      
        If (Board(bEnd) = 12 Or Board(bEnd) = 6) And Abs(bStart - bEnd) = 2 Then
 
             If bEnd = 27 Then Board(26) = Board(28): Board(28) = 0
          
        End If
        
        Turn = 1
        Range("K2").Value = "Turn: White"
        Range("K3").Value = "Opening Book: KID Theory"
        RenderBoard
        HighlightChecks
        Exit Sub
    End If

    ' 2. PREPARE AI SEARCH
    Dim moves As Collection
    Set moves = GetAllLegalMoves(2)
    
    If moves.count = 0 Then
        CheckGameStatus
        Range("K2").Value = "GAME OVER"
        Exit Sub
    End If
    
    SortMoves moves, 2
    
    Dim bestMove As Long
    Dim currentDepthMove As Long
    Dim score As Long
    Dim alpha As Long, beta As Long
    Dim depth As Integer
    
    ' - 3. INITIALIZE TIMER -
    AiStartTime = Timer
    StopSearch = False
    NodesCount = 0
    bestMove = moves(1) ' Default fallback
    
    ' - 4. ITERATIVE DEEPENING LOOP -
    For depth = 1 To MAX_DEPTH
        alpha = -30000
        beta = 30000
        Dim bestDepthScore As Long: bestDepthScore = -30000
        Dim moveItem As Variant
        
        For Each moveItem In moves
            Dim moveInt As Long: moveInt = moveItem
            Dim startSq As Integer: startSq = moveInt \ 1000
            Dim endSq As Integer: endSq = moveInt Mod 1000
            Dim cap As Integer
            
            ' Make
            cap = Board(endSq)
            Board(endSq) = Board(startSq)
            Board(startSq) = EMPTY_SQ
            
            ' Search
            score = -Negamax(depth - 1, -beta, -alpha, 1)
            
            ' Unmake
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

       ' 5. EXECUTE BEST MOVE
    Dim finalStart As Integer, finalEnd As Integer
    Dim pMoving As Integer
    
    finalStart = bestMove \ 1000
    finalEnd = bestMove Mod 1000
    pMoving = Board(finalStart)
    
    ' Update Board
    Board(finalEnd) = Board(finalStart)
    Board(finalStart) = EMPTY_SQ
    
   
    
    ' Castling Logic (Black Short Castle)
    ' If King moves 2 steps right (e8->g8 which is 25->27)
    If pMoving = 12 And (finalEnd - finalStart) = 2 Then
        ' Move the Rook from h8(28) to f8(26)
        Board(26) = Board(28)
        Board(28) = EMPTY_SQ
    End If
    
    ' Castling Logic (White Short Castle - just in case)
    If pMoving = 6 And (finalEnd - finalStart) = 2 Then
        Board(96) = Board(98)
        Board(98) = EMPTY_SQ
    End If
    
    ' Promotion (Simplest form: Auto-Queen)
    If pMoving = 7 And finalEnd > 90 Then Board(finalEnd) = 11 ' Black promotes
    If pMoving = 1 And finalEnd < 29 Then Board(finalEnd) = 5  ' White promotes
    
    ' -------------------------------------------------------

    ' Update History
    moveHistory = moveHistory & finalStart & finalEnd & " "
    
    Turn = 1
    Range("K2").Value = "Turn: White"
    
    RenderBoard
    HighlightChecks
    CheckGameStatus
End Sub


' THE ALGORITHM (Negamax + Quiescence)
Function Negamax(depth As Integer, alpha As Long, beta As Long, pColor As Integer) As Long
    ' - FAILSAFE: CHECK TIME -
    NodesCount = NodesCount + 1
    
    ' Only check system clock every X nodes to save CPU cycles
    If (NodesCount Mod CHECK_FREQ = 0) Then
        DoEvents ' Prevent "Not Responding"
        If Timer - AiStartTime > TIME_LIMIT Then
            StopSearch = True
        End If
    End If
    
    ' If flag is raised, return a dummy score to unwind recursion fast
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
    
    ' Sorting is heavy, only do it if we have time
    If depth > 1 Then SortMoves moves, pColor
    
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
        
        ' Fast exit if timeout occurred inside recursion
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
    ' FAILSAFE
    NodesCount = NodesCount + 1
    If (NodesCount Mod CHECK_FREQ = 0) Then
        DoEvents
        If Timer - AiStartTime > TIME_LIMIT Then StopSearch = True
    End If
    If StopSearch Then
        Quiescence = 0
        Exit Function
    End If
    '
    
    Dim stand_pat As Long
    stand_pat = EvaluatePosition()
    
    If stand_pat >= beta Then
        Quiescence = beta
        Exit Function
    End If
    If alpha < stand_pat Then alpha = stand_pat
    
    Dim moves As Collection
    Set moves = GetAllCaptures(pColor)
    ' Sorting captures is crucial for speed
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

'
' 3. MOVE GENERATORS (INTEGERS ONLY)
'
Function GetAllLegalMoves(pColor As Integer) As Collection
    Dim moves As New Collection
    Dim tempMoves As New Collection
    Dim startSq As Integer, endSq As Integer
    Dim piece As Integer, target As Integer
    
    ' PASS 1: Generate pseudo-legal moves (captures first)
    For startSq = 21 To 98
        piece = Board(startSq)
        If piece <> OFF_BOARD And piece <> EMPTY_SQ Then
            If (piece <= 6 And pColor = 1) Or (piece > 6 And pColor = 2) Then
                For endSq = 21 To 98
                    target = Board(endSq)
                    If target <> OFF_BOARD Then
                        If IsLegalMove(startSq, endSq) Then
                            tempMoves.Add (CLng(startSq) * 1000) + endSq
                        End If
                    End If
                Next endSq
            End If
        End If
    Next startSq
    
    ' PASS 2: Filter out moves that leave King in check
    Dim moveItem As Variant
    For Each moveItem In tempMoves
        Dim moveInt As Long: moveInt = moveItem
        startSq = moveInt \ 1000
        endSq = moveInt Mod 1000
        
        ' Make move
        Dim cap As Integer
        cap = Board(endSq)
        Board(endSq) = Board(startSq)
        Board(startSq) = EMPTY_SQ
        
        ' Check if King is safe
        If Not IsKingInCheck(pColor) Then
            moves.Add moveInt ' Legal!
        End If
        
        ' Unmake move
        Board(startSq) = Board(endSq)
        Board(endSq) = cap
    Next moveItem
    
    Set GetAllLegalMoves = moves
End Function

Function GetAllCaptures(pColor As Integer) As Collection
    Dim moves As New Collection
    Dim startSq As Integer, endSq As Integer
    Dim piece As Integer
    Dim target As Integer
    
    For startSq = 21 To 98
        piece = Board(startSq)
        If piece <> OFF_BOARD And piece <> EMPTY_SQ Then
            If (piece <= 6 And pColor = 1) Or (piece > 6 And pColor = 2) Then
                For endSq = 21 To 98
                    target = Board(endSq)
                    If target <> OFF_BOARD And target <> EMPTY_SQ Then
                        Dim targetColor As Integer
                        If target <= 6 Then targetColor = 1 Else targetColor = 2
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



' 4. EVALUATION FUNCTION (PST Optimized)
Public Function EvaluatePosition() As Long
    Dim i As Integer, p As Integer
    Dim score As Long
    Dim mirrorSq As Integer ' To flip table for Black
    
    score = 0
    
    For i = 21 To 98
        p = Board(i)
        If p <> EMPTY_SQ And p <> OFF_BOARD Then
            
            ' 120-Board Mirroring: 119 - i gives the opponent's perspective
            mirrorSq = 119 - i
            
            Select Case p
                ' --- WHITE PIECES (Add Score) ---
                Case 1: ' Pawn
                    score = score + 100 + pstPawn(i)
                Case 2: ' Knight
                    score = score + 320 + pstKnight(i)
                Case 3: ' Bishop
                    score = score + 330 + pstBishop(i)
                Case 4: ' Rook
                    score = score + 500 + pstRook(i)
                Case 5: ' Queen
                    score = score + 900 + pstQueen(i)
                Case 6: ' King
                    score = score + 20000 + pstKingMid(i)
                
                ' --- BLACK PIECES (Subtract Score) ---
                ' We use the same PST tables, but flipped (mirrorSq)
                Case 7: ' Pawn
                    score = score - (100 + pstPawn(mirrorSq))
                Case 8: ' Knight
                    score = score - (320 + pstKnight(mirrorSq))
                Case 9: ' Bishop
                    score = score - (330 + pstBishop(mirrorSq))
                Case 10: ' Rook
                    score = score - (500 + pstRook(mirrorSq))
                Case 11: ' Queen
                    score = score - (900 + pstQueen(mirrorSq))
                Case 12: ' King
                    score = score - (20000 + pstKingMid(mirrorSq))
            End Select
            
        End If
    Next i
    
    EvaluatePosition = score
End Function
Public Function IsKingInCheck(kingColor As Integer) As Boolean
    Dim kingSq As Integer
    Dim i As Integer
    
    ' Find the king
    For i = 21 To 98
        If kingColor = 1 And Board(i) = 6 Then kingSq = i: Exit For
        If kingColor = 2 And Board(i) = 12 Then kingSq = i: Exit For
    Next i
    
    If kingSq = 0 Then
        IsKingInCheck = True ' King missing = checkmate
        Exit Function
    End If
    
    ' Check if any enemy piece attacks the king
    Dim attackSq As Integer
    For attackSq = 21 To 98
        If Board(attackSq) <> EMPTY_SQ And Board(attackSq) <> OFF_BOARD Then
            ' Enemy piece?
            Dim attackColor As Integer
            If Board(attackSq) <= 6 Then attackColor = 1 Else attackColor = 2
            
            If attackColor <> kingColor Then
                ' Can this piece move to king square?
                If IsLegalMove(attackSq, kingSq) Then
                    IsKingInCheck = True
                    Exit Function
                End If
            End If
        End If
    Next attackSq
    
    IsKingInCheck = False
End Function
