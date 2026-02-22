Attribute VB_Name = "M_AI"
' MODULE: M_AI (IMPROVED)
Option Explicit

Private Const MAX_DEPTH As Integer = 4
Private Const TIME_LIMIT As Double = 10
Private Const CHECK_FREQ As Long = 200

' Opening phase constants
Private Const PHASE_OPENING As Integer = 0
Private Const PHASE_MIDDLEGAME As Integer = 1
Private Const PHASE_ENDGAME As Integer = 2

' Book move bonus — large enough to prefer book when equal, small enough
' that winning a piece (300+) or avoiding losing one (-300) overrides it
Private Const BOOK_BONUS As Integer = 60

' If we've deviated from book this many plies, stop consulting it
Private Const MAX_BOOK_PLY As Integer = 20

Private AiStartTime As Double
Public StopSearch As Boolean
Private NodesCount As Long
Private aiColor As Integer

' Track opening state
Private BookMovesPlayed As Integer
Private OpponentDeviated As Boolean

Public pstPawn(119) As Integer
Public pstKnight(119) As Integer
Public pstBishop(119) As Integer
Public pstRook(119) As Integer
Public pstQueen(119) As Integer
Public pstKingMid(119) As Integer
Public pstKingEnd(119) As Integer

Public Sub ForceStopAI()
    StopSearch = True
    Range("K3").Value = "Search Aborted."
End Sub

' =============================================================================
' GAME PHASE DETECTION
' Returns PHASE_OPENING, PHASE_MIDDLEGAME, or PHASE_ENDGAME
' Based on material remaining and move count — not just move number
' =============================================================================
Private Function GetGamePhase() As Integer
    Dim totalMaterial As Long
    Dim i As Integer, p As Integer
    
    totalMaterial = 0
    For i = 21 To 98
        p = Board(i)
        If p <> EMPTY_SQ And p <> OFF_BOARD Then
            Select Case p
                Case 1, 7:  totalMaterial = totalMaterial + 100   ' Pawns
                Case 2, 8:  totalMaterial = totalMaterial + 320   ' Knights
                Case 3, 9:  totalMaterial = totalMaterial + 330   ' Bishops
                Case 4, 10: totalMaterial = totalMaterial + 500   ' Rooks
                Case 5, 11: totalMaterial = totalMaterial + 900   ' Queens
            End Select
        End If
    Next i
    
    ' Opening: lots of pieces, few moves played
    If totalMaterial > 5500 And BookMovesPlayed < MAX_BOOK_PLY Then
        GetGamePhase = PHASE_OPENING
    ' Endgame: queens gone or very little material
    ElseIf totalMaterial < 2800 Then
        GetGamePhase = PHASE_ENDGAME
    Else
        GetGamePhase = PHASE_MIDDLEGAME
    End If
End Function

' =============================================================================
' OPPONENT DEVIATION DETECTOR
' Check if Black has deviated from KID-compatible responses to our London
' or if White has deviated from KID-targetable play as Black
' =============================================================================
Private Function DetectOpponentDeviation(pColor As Integer) As Boolean
    If pColor = 1 Then
        ' We are White (London). Has Black deviated from normal responses?
        ' Signs of deviation: Black fianchettoed queenside, played early c5,
        ' or played a Dutch (...f5) setup instead
        
        ' Black played early c5 with no d5 — could be Benoni/Sicilian-like
        If Board(43) = 7 And Board(44) = EMPTY_SQ And Board(45) = EMPTY_SQ Then
            DetectOpponentDeviation = True
            Exit Function
        End If
        
        ' Black fianchettoed queenside (b6 + Bb7) — Queen's Indian deviation
        If Board(46) = 7 And Board(39) = 9 Then
            DetectOpponentDeviation = True
            Exit Function
        End If
        
        ' Black played Dutch: f5 pawn pushed early
        If Board(46) = 7 And Board(44) = EMPTY_SQ Then
            DetectOpponentDeviation = True
            Exit Function
        End If
        
    ElseIf pColor = 2 Then
        ' We are Black (KID). Has White deviated from d4+c4 systems?
        ' If White hasn't played d4 or c4, KID book is irrelevant
        
        ' White didn't play d4 at all — must be 1.e4 or flank game
        If Board(64) <> 1 Then
            DetectOpponentDeviation = True
            Exit Function
        End If
        
        ' White went Colle-style (e3 before Nf3, no c4)
        If Board(64) = 1 And Board(73) = EMPTY_SQ And Board(75) = 1 Then
            DetectOpponentDeviation = True
            Exit Function
        End If
        
        ' White played f3 early (Sämisch vs KID) — book handles this but flag it
        ' so the engine gives book less weight and calculates more carefully
        If Board(76) = EMPTY_SQ And Board(76 - 10) = 1 Then ' f3 played
            DetectOpponentDeviation = True
            Exit Function
        End If
    End If
    
    DetectOpponentDeviation = False
End Function

' =============================================================================
' COUNT PIECES for contextual awareness
' =============================================================================
Private Function CountPieces(pColor As Integer) As Integer
    Dim i As Integer, count As Integer
    count = 0
    For i = 21 To 98
        If Board(i) <> EMPTY_SQ And Board(i) <> OFF_BOARD Then
            If pColor = 1 And Board(i) <= 6 Then count = count + 1
            If pColor = 2 And Board(i) > 6 And Board(i) <= 12 Then count = count + 1
        End If
    Next i
    CountPieces = count
End Function

' =============================================================================
' MAIN AI ENTRY POINT
' =============================================================================
Public Sub MakeComputerMove()
    On Error Resume Next
    If UBound(Board) <> 119 Then
        MsgBox "Memory lost. Resetting..."
        InitBoard
        Exit Sub
    End If
    On Error GoTo 0
    
    aiColor = Turn
    
    Dim moves As Collection
    Set moves = GetAllLegalMoves(aiColor)
    
    If moves.count = 0 Then
        CheckGameStatus
        Range("K2").Value = "GAME OVER"
        Exit Sub
    End If
    
    ' --- PHASE DETECTION ---
    Dim gamePhase As Integer
    gamePhase = GetGamePhase()
    
    ' --- DEVIATION DETECTION ---
    ' Only check for deviation in opening phase
    If gamePhase = PHASE_OPENING And Not OpponentDeviated Then
        OpponentDeviated = DetectOpponentDeviation(aiColor)
        If OpponentDeviated Then
            Range("K3").Value = "Opponent deviated — going tactical!"
        End If
    End If
    
    ' Filter moves that hang pieces
    Dim safeMoves As Collection
    Set safeMoves = FilterHangingMoves(moves, aiColor)
    
    If safeMoves.count = 0 Then
        Set safeMoves = moves
        Range("K3").Value = "WARNING: All moves lose material"
    End If
    
    SortMoves safeMoves, aiColor
    
    ' --- BOOK MOVE LOGIC ---
    ' Only use book if: opening phase, book ply limit not exceeded,
    ' opponent hasn't deviated in a way that makes our system irrelevant
    Dim targetBookMove As Long
    targetBookMove = 0
    
    Dim useBook As Boolean
    useBook = (gamePhase = PHASE_OPENING) And _
              (BookMovesPlayed < MAX_BOOK_PLY) And _
              Not (OpponentDeviated And Not CanAdaptBookToDeviation(aiColor))
    
    If useBook Then
        targetBookMove = GetDesiredOpeningMove(aiColor, safeMoves)
    End If
    
    ' --- TACTICAL SITUATION ASSESSMENT ---
    ' Reduce book bonus if we're in a sharp position (pieces hanging, checks available)
    Dim effectiveBookBonus As Integer
    effectiveBookBonus = BOOK_BONUS
    
    If IsKingInCheck(3 - aiColor) Then
        ' We're giving check or opponent is in check — be tactical, reduce book weight
        effectiveBookBonus = BOOK_BONUS \ 3
    End If
    
    If CountHangingPieces(3 - aiColor) > 0 Then
        ' Opponent has hanging pieces — don't follow book, take them!
        effectiveBookBonus = 0
        targetBookMove = 0
        Range("K3").Value = "Tactical opportunity — ignoring book!"
    End If
    
    ' --- SEARCH ---
    Dim bestMove As Long
    Dim currentDepthMove As Long
    Dim score As Long
    Dim alpha As Long, beta As Long
    Dim depth As Integer
    
    AiStartTime = Timer
    StopSearch = False
    NodesCount = 0
    bestMove = safeMoves(1)
    
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
            
            cap = Board(endSq)
            Board(endSq) = Board(startSq)
            Board(startSq) = EMPTY_SQ
            
            score = -Negamax(depth - 1, -beta, -alpha, 3 - aiColor)
            
            ' Apply book bonus — but only if tactically safe to do so
            ' (effectiveBookBonus of 0 means we've detected a tactic to exploit)
            If moveInt = targetBookMove And effectiveBookBonus > 0 Then
                ' Additional safety: don't apply bonus if this book move loses material
                ' i.e., if the raw score is very negative, the book move is a blunder
                If score > -150 Then
                    score = score + effectiveBookBonus
                End If
            End If
            
            Board(startSq) = Board(endSq)
            Board(endSq) = cap
            
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
            
            ' Status reporting
            Dim statusMsg As String
            If bestMove = targetBookMove And targetBookMove <> 0 Then
                statusMsg = "D" & depth & " | Book: " & PhaseLabel(gamePhase)
            ElseIf targetBookMove <> 0 Then
                statusMsg = "D" & depth & " | Tactical Override!"
            ElseIf gamePhase = PHASE_OPENING And OpponentDeviated Then
                statusMsg = "D" & depth & " | Adapting to deviation"
            Else
                statusMsg = "D" & depth & " | " & PhaseLabel(gamePhase)
            End If
            Range("K3").Value = statusMsg
        End If
        DoEvents
    Next depth
    
    ' Track book ply count
    If bestMove = targetBookMove And targetBookMove <> 0 Then
        BookMovesPlayed = BookMovesPlayed + 1
    End If
    
    ' --- EXECUTE MOVE ---
    Dim finalStart As Integer, finalEnd As Integer
    Dim pMoving As Integer
    Dim isCapture As Boolean
    
    finalStart = bestMove \ 1000
    finalEnd = bestMove Mod 1000
    pMoving = Board(finalStart)
    isCapture = (Board(finalEnd) <> EMPTY_SQ)
    
    ' En Passant
    If (pMoving = 1 Or pMoving = 7) And finalEnd = EnPassantTarget And Board(finalEnd) = EMPTY_SQ Then
        Dim victimSq As Integer
        If aiColor = 1 Then victimSq = finalEnd + 10 Else victimSq = finalEnd - 10
        Board(victimSq) = EMPTY_SQ
        isCapture = True
    End If
    
    Board(finalEnd) = Board(finalStart)
    Board(finalStart) = EMPTY_SQ
    
    ' Castling rook updates
    If pMoving = 12 And (finalEnd - finalStart) = 2 Then Board(26) = Board(28): Board(28) = EMPTY_SQ
    If pMoving = 12 And (finalStart - finalEnd) = 2 Then Board(24) = Board(21): Board(21) = EMPTY_SQ
    If pMoving = 6 And (finalEnd - finalStart) = 2 Then Board(96) = Board(98): Board(98) = EMPTY_SQ
    If pMoving = 6 And (finalStart - finalEnd) = 2 Then Board(94) = Board(91): Board(91) = EMPTY_SQ
    
    ' Promotion (auto-queen)
    If pMoving = 7 And finalEnd >= 91 Then Board(finalEnd) = 11
    If pMoving = 1 And finalEnd <= 28 Then Board(finalEnd) = 5
    
    ' Flag updates
    If pMoving = 6 Then W_KingMoved = True
    If pMoving = 12 Then B_KingMoved = True
    If finalStart = 91 Then W_RookLMoved = True
    If finalStart = 98 Then W_RookRMoved = True
    If finalStart = 21 Then B_RookLMoved = True
    If finalStart = 28 Then B_RookRMoved = True
    
    EnPassantTarget = 0
    If (pMoving = 1 Or pMoving = 7) And Abs(finalEnd - finalStart) = 20 Then
        EnPassantTarget = (finalStart + finalEnd) / 2
    End If
    
    moveHistory = moveHistory & finalStart & finalEnd & " "
    RecordMove finalStart, finalEnd, isCapture
    
    ' --- HIGHLIGHT AI MOVE ---
    ' Store the computer's move coordinates into our visual variables
    LastMoveStart = finalStart
    LastMoveEnd = finalEnd
    
    Turn = 3 - aiColor
    Range("K5").Value = "Turn: " & IIf(Turn = 1, "White", "Black")
    
    ' --- RENDER WITH HIGHLIGHTS ---
    RenderBoard
    ApplyCheckerboard ' Paints the background colors so the AI's move glows
    HighlightChecks
    CheckGameStatus
End Sub

Private Function PhaseLabel(phase As Integer) As String
    Select Case phase
        Case PHASE_OPENING:    PhaseLabel = "Opening"
        Case PHASE_MIDDLEGAME: PhaseLabel = "Middlegame"
        Case PHASE_ENDGAME:    PhaseLabel = "Endgame"
    End Select
End Function

' =============================================================================
' COUNT HANGING PIECES for the opponent (undefended pieces we can take)
' =============================================================================
Private Function CountHangingPieces(targetColor As Integer) As Integer
    Dim i As Integer, count As Integer
    count = 0
    For i = 21 To 98
        Dim p As Integer: p = Board(i)
        If p <> EMPTY_SQ And p <> OFF_BOARD Then
            Dim pColor As Integer
            If p <= 6 Then pColor = 1 Else pColor = 2
            If pColor = targetColor Then
                ' Is this piece attacked by us AND not defended by them?
                If IsSquareAttacked(i, 3 - targetColor) And Not IsSquareDefended(i, targetColor) Then
                    count = count + 1
                End If
            End If
        End If
    Next i
    CountHangingPieces = count
End Function

' =============================================================================
' CAN WE ADAPT OUR BOOK TO THE OPPONENT'S DEVIATION?
' Some deviations (e.g. Black plays KID against our London) are fine.
' Others (e.g. 1.e4 e5 — we're playing Black's KID moves) are not.
' =============================================================================
Private Function CanAdaptBookToDeviation(pColor As Integer) As Boolean
    If pColor = 1 Then
        ' London can adapt to most Black setups — it's a system opening
        ' Only truly incompatible if Black plays something we can't reach with Bf4
        CanAdaptBookToDeviation = True
    ElseIf pColor = 2 Then
        ' KID as Black requires White to have played d4 + c4
        ' If White played 1.e4 or 1.Nf3 without d4, KID is impossible
        If Board(64) = 1 Then
            CanAdaptBookToDeviation = True  ' d4 was played, KID still reachable
        Else
            CanAdaptBookToDeviation = False ' Wrong opening — ignore book entirely
        End If
    End If
End Function

' =============================================================================
' OPENING BOOK — London (White) and KID (Black)
' Priority order matters: most important setup moves come first.
' Each condition checks whether the desired piece is ALREADY on its target square.
' If it is, we skip to the next priority. If not, we try to play it.
' =============================================================================
Function GetDesiredOpeningMove(pColor As Integer, validMoves As Collection) As Long
    GetDesiredOpeningMove = 0
    
    If pColor = 1 Then
        ' -----------------------------------------------------------------------
        ' LONDON SYSTEM (White)
        ' Priority: d4 first, then Bf4 BEFORE e3 (key London move order!),
        ' then Nf3, e3, Nbd2, c3 support, Bd3 (the key attacking bishop), h3
        ' -----------------------------------------------------------------------
        
        ' 1. d4 — must be first move
        If Board(64) <> 1 And IsMoveInCollection(84064, validMoves) Then
            GetDesiredOpeningMove = 84064: Exit Function
        End If
        
        ' 2. Bf4 BEFORE e3 — the London move order (critical!)
        '    Only play Bf4 if e3 hasn't closed the diagonal yet
        If Board(66) <> 3 And Board(75) = EMPTY_SQ And IsMoveInCollection(93066, validMoves) Then
            GetDesiredOpeningMove = 93066: Exit Function
        End If
        
        ' 3. Nf3 — develop the knight
        If Board(76) <> 2 And IsMoveInCollection(97076, validMoves) Then
            GetDesiredOpeningMove = 97076: Exit Function
        End If
        
        ' 4. e3 — only AFTER Bf4 is developed (don't lock bishop in!)
        If Board(75) <> 1 And Board(66) = 3 And IsMoveInCollection(85075, validMoves) Then
            GetDesiredOpeningMove = 85075: Exit Function
        End If
        
        ' 5. Bd3 — the key attacking piece (aims at h7)
        '    Only go to d3 if f4 is occupied (our bishop) and e3 pawn supports
        If Board(74) <> 3 And Board(66) = 3 And Board(75) = 1 And IsMoveInCollection(96074, validMoves) Then
            GetDesiredOpeningMove = 96074: Exit Function
        End If
        
        ' 6. Nbd2 — flexible, supports e4 or goes to e5/f1
        If Board(84) <> 2 And IsMoveInCollection(92084, validMoves) Then
            GetDesiredOpeningMove = 92084: Exit Function
        End If
        
        ' 7. c3 — solidify the center triangle
        If Board(73) <> 1 And IsMoveInCollection(83073, validMoves) Then
            GetDesiredOpeningMove = 83073: Exit Function
        End If
        
        ' 8. h3 — prevent ...Ng4 hitting Bf4 or Bd3
        '    Only play this if we haven't castled yet (early prevention)
        If Board(78) <> 1 And Board(67) = EMPTY_SQ And IsMoveInCollection(88078, validMoves) Then
            GetDesiredOpeningMove = 88078: Exit Function
        End If
        
        ' 9. Castle kingside — safety
        If Not W_KingMoved And IsMoveInCollection(95097, validMoves) Then
            GetDesiredOpeningMove = 95097: Exit Function
        End If
        
    ElseIf pColor = 2 Then
        ' -----------------------------------------------------------------------
        ' KING'S INDIAN DEFENSE (Black)
        ' Priority: Nf6 first, then g6, Bg7, d6, 0-0, THEN e5
        ' e5 only after castled — don't push e5 without king safety!
        ' -----------------------------------------------------------------------
        
        ' 1. Nf6 — develop knight, fight for center
        If Board(46) <> 8 And IsMoveInCollection(27046, validMoves) Then
            GetDesiredOpeningMove = 27046: Exit Function
        End If
        
        ' 2. g6 — prepare fianchetto
        If Board(47) <> 7 And IsMoveInCollection(37047, validMoves) Then
            GetDesiredOpeningMove = 37047: Exit Function
        End If
        
        ' 3. Bg7 — the crown jewel bishop
        If Board(37) <> 9 And Board(47) = 7 And IsMoveInCollection(26037, validMoves) Then
            GetDesiredOpeningMove = 26037: Exit Function
        End If
        
        ' 4. d6 — pawn structure (only after Bg7 is developed)
        If Board(44) <> 7 And Board(37) = 9 And IsMoveInCollection(34044, validMoves) Then
            GetDesiredOpeningMove = 34044: Exit Function
        End If
        
        ' 5. Castle kingside — MUST happen before e5!
        '    King safety is paramount in KID
        If Board(27) <> 12 And Board(37) = 9 And Board(46) = 8 And IsMoveInCollection(25027, validMoves) Then
            GetDesiredOpeningMove = 25027: Exit Function
        End If
        
        ' 6. e5 — ONLY after castled and d6 is set up
        '    The central strike — the heart of KID strategy
        If Board(55) <> 7 And Board(27) = 12 And Board(44) = 7 And IsMoveInCollection(35055, validMoves) Then
            GetDesiredOpeningMove = 35055: Exit Function
        End If
        
        ' 7. Nbd7 — flexible knight development (after e5 creates room)
        '    Avoids Nc6 which can be hit by d5 in some lines
        If Board(34) <> 8 And Board(27) = 12 And IsMoveInCollection(22034, validMoves) Then
            GetDesiredOpeningMove = 22034: Exit Function
        End If
    End If
End Function

Function IsMoveInCollection(targetMove As Long, col As Collection) As Boolean
    Dim m As Variant
    For Each m In col
        If CLng(m) = targetMove Then
            IsMoveInCollection = True
            Exit Function
        End If
    Next m
    IsMoveInCollection = False
End Function

' =============================================================================
' NEGAMAX with Alpha-Beta Pruning
' =============================================================================
Function Negamax(depth As Integer, alpha As Long, beta As Long, pColor As Integer) As Long
    NodesCount = NodesCount + 1
    
    If (NodesCount Mod CHECK_FREQ = 0) Then
        DoEvents
        If Timer - AiStartTime > TIME_LIMIT Then StopSearch = True
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
        If IsKingInCheck(pColor) Then
            Negamax = -20000 + depth
        Else
            Negamax = 0
        End If
        Exit Function
    End If
    
    SortMoves moves, pColor
    
    Dim moveInt As Long, startSq As Integer, endSq As Integer, cap As Integer
    Dim score As Long, moveItem As Variant
    
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

' =============================================================================
' QUIESCENCE SEARCH — Resolves captures before returning static eval
' =============================================================================
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
    If pColor = 2 Then stand_pat = -stand_pat
    
    If stand_pat >= beta Then
        Quiescence = beta
        Exit Function
    End If
    If alpha < stand_pat Then alpha = stand_pat
    
    Dim moves As Collection
    Set moves = GetAllCaptures(pColor)
    SortMoves moves, pColor
    
    Dim moveInt As Long, startSq As Integer, endSq As Integer, cap As Integer
    Dim score As Long, moveItem As Variant
    
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

' =============================================================================
' MOVE GENERATION
' =============================================================================
Function GetAllLegalMoves(pColor As Integer) As Collection
    Dim moves As New Collection
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
                            moves.Add (CLng(startSq) * 1000) + endSq
                        End If
                    End If
                Next endSq
            End If
        End If
    Next startSq
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

' =============================================================================
' KING SAFETY
' =============================================================================
Public Function IsKingInCheck(kingColor As Integer) As Boolean
    Dim kingSq As Integer, i As Integer, kVal As Integer
    kVal = IIf(kingColor = 1, 6, 12)
    
    For i = 21 To 98
        If Board(i) = kVal Then
            kingSq = i
            Exit For
        End If
    Next i
    
    If kingSq = 0 Then
        IsKingInCheck = True
        Exit Function
    End If
    IsKingInCheck = IsSquareAttacked(kingSq, 3 - kingColor)
End Function

' =============================================================================
' PIECE SAFETY FILTERS
' =============================================================================
Function FilterHangingMoves(allMoves As Collection, pColor As Integer) As Collection
    Dim safeMoves As New Collection, moveItem As Variant
    
    For Each moveItem In allMoves
        Dim moveInt As Long: moveInt = moveItem
        Dim startSq As Integer: startSq = moveInt \ 1000
        Dim endSq As Integer: endSq = moveInt Mod 1000
        
        Dim cap As Integer, movingPiece As Integer
        movingPiece = Board(startSq)
        cap = Board(endSq)
        Board(endSq) = Board(startSq)
        Board(startSq) = EMPTY_SQ
        
        Dim isSafe As Boolean: isSafe = True
        If IsSquareAttacked(endSq, 3 - pColor) Then
            Dim pieceValue As Integer, captureValue As Integer
            pieceValue = GetPieceValue(movingPiece)
            captureValue = GetPieceValue(cap)
            
            If captureValue < pieceValue Then
                If Not IsSquareDefended(endSq, pColor) Then isSafe = False
            End If
        End If
        
        Board(startSq) = Board(endSq)
        Board(endSq) = cap
        
        If isSafe Then safeMoves.Add moveInt
    Next moveItem
    Set FilterHangingMoves = safeMoves
End Function

Function IsSquareDefended(sq As Integer, byColor As Integer) As Boolean
    Dim defenderSq As Integer
    For defenderSq = 21 To 98
        If Board(defenderSq) <> EMPTY_SQ And Board(defenderSq) <> OFF_BOARD Then
            Dim defenderColor As Integer
            If Board(defenderSq) <= 6 Then defenderColor = 1 Else defenderColor = 2
            
            If defenderColor = byColor Then
                Dim tempPiece As Integer: tempPiece = Board(sq)
                Board(sq) = EMPTY_SQ
                Dim canDefend As Boolean: canDefend = IsLegalMove(defenderSq, sq)
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

' =============================================================================
' STATIC EVALUATION
' Material + PST + Contextual positional bonuses (phase-aware)
' =============================================================================
Public Function EvaluatePosition() As Long
    Dim i As Integer, p As Integer, score As Long, mirrorSq As Integer
    score = 0
    
    Dim phase As Integer
    phase = GetGamePhase()
    
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
                Case 6:
                    If phase = PHASE_ENDGAME Then
                        score = score + 20000 + pstKingEnd(i)
                    Else
                        score = score + 20000 + pstKingMid(i)
                    End If
                Case 7: score = score - (100 + pstPawn(mirrorSq))
                Case 8: score = score - (320 + pstKnight(mirrorSq))
                Case 9: score = score - (330 + pstBishop(mirrorSq))
                Case 10: score = score - (500 + pstRook(mirrorSq))
                Case 11: score = score - (900 + pstQueen(mirrorSq))
                Case 12:
                    If phase = PHASE_ENDGAME Then
                        score = score - (20000 + pstKingEnd(mirrorSq))
                    Else
                        score = score - (20000 + pstKingMid(mirrorSq))
                    End If
            End Select
        End If
    Next i
    
    ' =========================================================================
    ' CONTEXTUAL POSITIONAL BONUSES
    ' These only fire in opening/middlegame — not endgame where they're irrelevant
    ' =========================================================================
    
    If phase <> PHASE_ENDGAME Then
    
        ' --- WHITE: London System Themes ---
        
        ' London Triangle: d4+e3+c3 — solid center base (+30)
        If Board(64) = 1 And Board(75) = 1 And Board(73) = 1 Then score = score + 30
        
        ' Bishop coordination: Bf4 + Bd3 in place — classic London battery (+25)
        If Board(66) = 3 And Board(74) = 3 Then score = score + 25
        
        ' Bf4 developed before e3 — correct move order (+10)
        ' (Bf4 is on f4, e3 pawn in place — only bonus if Bf4 came first which we can't
        '  track directly, but we reward the formation itself)
        If Board(66) = 3 Then score = score + 10
        
        ' Ne5 outpost — heart of London attack (+45)
        ' Only bonus if it's actually our knight (piece=2, not 8)
        If Board(65) = 2 Then score = score + 45
        
        ' London Spike: f4+g4 pawns advanced — kingside attack theme (+20)
        If Board(66) = 1 And Board(67) = 1 Then score = score + 20
        
        ' Rook on a-file after queen trade — London queenside theme (+15)
        If Board(91) = 4 Or Board(81) = 4 Then score = score + 15
        
        ' h3 prevention of ...Ng4 is worth a small bonus when Bf4 is out (+8)
        If Board(78) = 1 And Board(66) = 3 Then score = score + 8
        
        ' PENALTY: Bishop locked behind e3 before Bf4 developed — wrong move order (-20)
        If Board(75) = 1 And Board(66) = EMPTY_SQ And Board(91) = EMPTY_SQ Then
            ' e3 played but Bf4 not yet out AND bishop not yet moved from c1 (91)
            ' This means we played e3 before developing Bf4 — punish it
            score = score - 20
        End If
        
        ' --- BLACK: King's Indian Defense Themes ---
        
        ' Bg7 in place — crown jewel (+40)
        If Board(37) = 9 Then score = score - 40
        
        ' Full KID setup: Nf6 + g6 + Bg7 + d6 + castled (+35 bonus stack)
        If Board(46) = 8 And Board(47) = 7 And Board(37) = 9 And Board(44) = 7 Then
            score = score - 35
        End If
        
        ' Castled king — safety before launching attack (-20 bonus for Black)
        If Board(27) = 12 Then score = score - 20
        
        ' e5 pushed AFTER castling — correct KID strategy (-20)
        If Board(55) = 7 And Board(27) = 12 Then score = score - 20
        
        ' PENALTY: e5 pushed before castling — premature (-15)
        If Board(55) = 7 And Board(27) <> 12 Then score = score + 15
        
        ' KID Nf4 outpost — extremely powerful knight (-40)
        If Board(66) = 8 Then score = score - 40
        
        ' Nh5 — in transit to f4, or targeting f4 (-15)
        If Board(58) = 8 Then score = score - 15
        
        ' f5 storm — kingside attack underway (-25)
        If Board(56) = 7 And Board(27) = 12 Then score = score - 25
        
        ' g5 storm — full kingside pawn advance (-20)
        If Board(57) = 7 And Board(56) = 7 Then score = score - 20
        
        ' PENALTY: Bg7 not developed by move 6+ (if g6 is played but no Bg7) (+15)
        If Board(47) = 7 And Board(37) <> 9 And BookMovesPlayed > 2 Then
            score = score + 15
        End If
        
    End If
    
    ' =========================================================================
    ' ENDGAME-SPECIFIC EVAL
    ' =========================================================================
    If phase = PHASE_ENDGAME Then
        ' Passed pawns become much more valuable in endgames
        Dim file As Integer
        For file = 1 To 8
            ' Check for passed pawns (simplified: no enemy pawn ahead on same file)
            ' This is a rough heuristic but better than nothing
            Dim rank As Integer
            For rank = 3 To 8
                Dim sq As Integer: sq = (10 - rank) * 10 + file + 20
                If sq >= 21 And sq <= 98 Then
                    If Board(sq) = 1 Then  ' White pawn — check if passed
                        Dim blocked As Boolean: blocked = False
                        Dim checkRank As Integer
                        For checkRank = rank + 1 To 8
                            Dim checkSq As Integer: checkSq = (10 - checkRank) * 10 + file + 20
                            If checkSq >= 21 And checkSq <= 98 Then
                                If Board(checkSq) = 7 Then blocked = True
                            End If
                        Next checkRank
                        If Not blocked Then score = score + (rank - 2) * 15
                    End If
                End If
            Next rank
        Next file
    End If
    
    EvaluatePosition = score
End Function


