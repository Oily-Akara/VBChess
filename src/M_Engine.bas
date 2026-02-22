Attribute VB_Name = "M_Engine"
' Module: M_Engine
Option Explicit
Public SelectedSq As Integer ' Stores which square clicked first
Public Turn As Integer ' 1 for White, 2 for Black
Public PositionHistory(1 To 500) As String
Public PositionCount As Integer

' HISTORY VARIABLES
' 120-square board based PST arrays
Public pstPawn(119) As Integer
Public pstKnight(119) As Integer
Public pstBishop(119) As Integer
Public pstRook(119) As Integer
Public pstQueen(119) As Integer
Public pstKingMid(119) As Integer
Public pstKingEnd(119) As Integer

' Search Heuristics
Public KillerMoves(20, 2) As Long   ' Store killer moves for up to depth 20
Public HistoryMoves(119, 119) As Long ' History heuristic

' Move Buffers (To replace Collections inside recursion)
Private MoveBuffer(20, 100) As Long ' Depth 20, Max 100 moves
Private MoveScores(20, 100) As Long

' HIGHLIGHTING VARIABLES (FANCY VISUALS)
Public LastMoveStart As Integer
Public LastMoveEnd As Integer

' 1. Track if Kings have moved
Public W_KingMoved As Boolean
Public B_KingMoved As Boolean

' 2. Track if Rooks have moved
Public W_RookLMoved As Boolean, W_RookRMoved As Boolean
Public B_RookLMoved As Boolean, B_RookRMoved As Boolean

' 3. Track En Passant (Stores the index of the "Ghost" square, or 0 if none)
Public EnPassantTarget As Integer

'  CONSTANTS
Public Const EMPTY_SQ As Integer = 0
Public Const OFF_BOARD As Integer = 99

' White Pieces (1-6)
Public Const W_PAWN As Integer = 1
Public Const W_KNIGHT As Integer = 2
Public Const W_BISHOP As Integer = 3
Public Const W_ROOK As Integer = 4
Public Const W_QUEEN As Integer = 5
Public Const W_KING As Integer = 6

' Black Pieces (7-12)
Public Const B_PAWN As Integer = 7
Public Const B_KNIGHT As Integer = 8
Public Const B_BISHOP As Integer = 9
Public Const B_ROOK As Integer = 10
Public Const B_QUEEN As Integer = 11
Public Const B_KING As Integer = 12
Public moveHistory As String

Public Board(0 To 119) As Integer

'  MAIN SETUP SUB
Public Sub InitBoard()
    Dim i As Integer
    Dim r As Integer, c As Integer

    ' 1. Reset Arrays
    For i = 0 To 119
        Board(i) = OFF_BOARD
    Next i
    
    ' 2. Clear Center
    For r = 0 To 7
        For c = 1 To 8
            Board(20 + (r * 10) + c) = EMPTY_SQ
        Next c
    Next r

    ' 3. Reset History & Flags
    moveHistory = ""
    W_KingMoved = False: B_KingMoved = False
    EnPassantTarget = 0
    PositionCount = 0
    
    ' Reset Highlights
    LastMoveStart = 0
    LastMoveEnd = 0
    SelectedSq = 0
    
    ' 4. Initialize Move Notation Array
    ReDim MoveNotation(1 To 200)
    MoveCount = 0
    
    ' 4. CLEAR SEARCH MEMORY
    Dim d As Integer, m As Integer
    For d = 0 To 20
        For m = 0 To 2
            KillerMoves(d, m) = 0
        Next m
    Next d
    For r = 0 To 119
        For c = 0 To 119
            HistoryMoves(r, c) = 0
        Next c
    Next r
   
    ' 5. Init Tables
    InitPST
     SetupSidePanel ' Builds the notation table
    UpdateTurnUI   ' Draws the turn indicator
    
    ' 6. Setup & Render
    SetupStandardPosition
    RenderBoard
    ApplyCheckerboard
    
    ' 7. Reset Turn
    Turn = 1
End Sub
Private Sub InitPST()
    Dim i As Integer
    ' Fill Arrays with 0
    For i = 0 To 119
        pstPawn(i) = 0: pstKnight(i) = 0: pstBishop(i) = 0
    Next i
    
    ' We map 0-63 (8x8) values to the 120 board
    ' These values are mirrored. For White, use index. For Black, use mirror index.
    
    Dim sq As Integer, file As Integer, rank As Integer
    Dim val As Integer
    
    ' SIMPLE KID-FOCUSED TABLES (Values in Centipawns)
    
    ' KNIGHTS: Like center, but especially f5/c5 for attacking
    Dim kTable As Variant
    kTable = Array(-50, -40, -30, -30, -30, -30, -40, -50, _
                   -40, -20, 0, 0, 0, 0, -20, -40, _
                   -30, 0, 10, 15, 15, 10, 0, -30, _
                   -30, 5, 15, 20, 20, 15, 5, -30, _
                   -30, 0, 15, 20, 20, 15, 0, -30, _
                   -30, 5, 10, 15, 15, 10, 5, -30, _
                   -40, -20, 0, 5, 5, 0, -20, -40, _
                   -50, -40, -30, -30, -30, -30, -40, -50)
                   
    ' PAWNS: Encourage advancement and center control
    Dim pTable As Variant
    pTable = Array(0, 0, 0, 0, 0, 0, 0, 0, _
                   50, 50, 50, 50, 50, 50, 50, 50, _
                   10, 10, 20, 30, 30, 20, 10, 10, _
                   5, 5, 10, 25, 25, 10, 5, 5, _
                   0, 0, 0, 20, 20, 0, 0, 0, _
                   5, -5, -10, 0, 0, -10, -5, 5, _
                   5, 10, 10, -20, -20, 10, 10, 5, _
                   0, 0, 0, 0, 0, 0, 0, 0)

    ' MAPPING LOOP
    Dim r8 As Integer, c8 As Integer
    Dim board120 As Integer
    
    For r8 = 0 To 7
        For c8 = 0 To 7
            board120 = 21 + (r8 * 10) + c8
            
            ' Apply values
            pstKnight(board120) = kTable((r8 * 8) + c8)
            pstPawn(board120) = pTable((r8 * 8) + c8)
            
            ' Bishop: Slight bonus for Fianchetto squares (g2/b2 for white, g7/b7 for black)
            If board120 = 32 Or board120 = 37 Then pstBishop(board120) = 20 ' Fianchetto spots
            
            ' Center control for others
            If c8 >= 3 And c8 <= 4 And r8 >= 3 And r8 <= 4 Then
                pstBishop(board120) = pstBishop(board120) + 10
                pstQueen(board120) = 10
            End If
        Next c8
    Next r8
End Sub
' PLACES PIECES FOR START OF GAME
Private Sub SetupStandardPosition()
    Dim i As Integer
    
    ' --- BLACK PIECES (Top Rows 21-28) ---
    ' Order: Rook(10), Knight(8), Bishop(9), Queen(11), King(12), Bishop(9), Knight(8), Rook(10)
    Board(21) = 10: Board(22) = 8: Board(23) = 9: Board(24) = 11
    Board(25) = 12: Board(26) = 9: Board(27) = 8: Board(28) = 10
    
    ' --- BLACK PAWNS (Row 31-38) ---
    For i = 31 To 38
        Board(i) = 7
    Next i
    
    ' --- WHITE PAWNS (Row 81-88) ---
    For i = 81 To 88
        Board(i) = 1
    Next i
    
    ' --- WHITE PIECES (Bottom Rows 91-98) ---
    ' Order: Rook(4), Knight(2), Bishop(3), Queen(5), King(6), Bishop(3), Knight(2), Rook(4)
    Board(91) = 4: Board(92) = 2: Board(93) = 3: Board(94) = 5
    Board(95) = 6: Board(96) = 3: Board(97) = 2: Board(98) = 4
    
    ' Set Castling Rights (Flags)
    W_KingMoved = False
    B_KingMoved = False
    W_RookLMoved = False: W_RookRMoved = False
    B_RookLMoved = False: B_RookRMoved = False
End Sub



Function Negamax(depth As Integer, alpha As Long, beta As Long, pColor As Integer) As Long
    
    ' 1. TIMEOUT CHECK (Every 2048 nodes)
    NodesCount = NodesCount + 1
    If (NodesCount And 2047) = 0 Then
        DoEvents
        If Timer - AiStartTime > TIME_LIMIT Then StopSearch = True
    End If
    If StopSearch Then Negamax = 0: Exit Function

    ' 2. BASE CASE
    If depth <= 0 Then
        Negamax = Quiescence(alpha, beta, pColor)
        Exit Function
    End If
    
    ' 3. MOVE GENERATION
    ' Convert Collection to Array for Speed & Sorting
    Dim mColl As Collection
    Set mColl = GetAllLegalMoves(pColor)
    
    Dim count As Integer
    count = mColl.count
    If count = 0 Then
        Negamax = -20000 + depth ' Checkmate detection
        Exit Function
    End If
    
    ' Load moves into local scope buffers
    Dim i As Integer
    Dim mVal As Long
    Dim sSq As Integer, eSq As Integer
    Dim pieceVal As Integer
    
    For i = 1 To count
        mVal = mColl(i)
        MoveBuffer(depth, i) = mVal
        
        ' 4. SCORING MOVES (Move Ordering)
        ' This is where we gain ELO. We guess which move is best before searching.
        
        sSq = mVal \ 1000
        eSq = mVal Mod 1000
        MoveScores(depth, i) = 0
        
        ' Priority 1: Captures (MVV-LVA)
        If Board(eSq) <> EMPTY_SQ Then
             ' Victim Value * 10 - Attacker Value
             ' (Taking a Queen with Pawn is huge, Taking Pawn with Queen is low)
             Dim vicVal As Integer, attVal As Integer
             vicVal = GetPieceValue(Board(eSq))
             attVal = GetPieceValue(Board(sSq))
             MoveScores(depth, i) = 10000 + (vicVal * 10) - attVal
             
        ' Priority 2: Killer Moves (Moves that caused cutoffs recently)
        ElseIf mVal = KillerMoves(depth, 1) Then
             MoveScores(depth, i) = 9000
        ElseIf mVal = KillerMoves(depth, 2) Then
             MoveScores(depth, i) = 8000
             
        ' Priority 3: History Heuristic (Good quiet moves)
        Else
             MoveScores(depth, i) = HistoryMoves(sSq, eSq)
        End If
    Next i
    
    ' 5. THE LOOP (Selection Sort)
    Dim bestScore As Long, bestIndex As Integer
    Dim cap As Integer, score As Long
    Dim bestMoveFound As Long
    
    
    For i = 1 To count
        
        ' Find best remaining move in list
        bestScore = -100000
        bestIndex = -1
        
        Dim j As Integer
        For j = 1 To count
            If MoveScores(depth, j) > bestScore Then
                bestScore = MoveScores(depth, j)
                bestIndex = j
            End If
        Next j
        
        ' Mark as visited (set score super low)
        MoveScores(depth, bestIndex) = -200000
        mVal = MoveBuffer(depth, bestIndex)
        
        ' Decode
        sSq = mVal \ 1000
        eSq = mVal Mod 1000
        cap = Board(eSq)
        
        ' MAKE
        Board(eSq) = Board(sSq)
        Board(sSq) = EMPTY_SQ
        
        ' RECURSE
        score = -Negamax(depth - 1, -beta, -alpha, 3 - pColor)
        
        ' UNMAKE
        Board(sSq) = Board(eSq)
        Board(eSq) = cap
        
        If StopSearch Then Exit Function
        
        If score >= beta Then
            ' FAIL HIGH (Beta Cutoff)
            ' This move is too good, opponent won't allow it.
            
            ' Store Killer Move (if not capture)
            If cap = EMPTY_SQ Then
                KillerMoves(depth, 2) = KillerMoves(depth, 1)
                KillerMoves(depth, 1) = mVal
                ' Update History
                HistoryMoves(sSq, eSq) = HistoryMoves(sSq, eSq) + (depth * depth)
            End If
            
            Negamax = beta
            Exit Function
        End If
        
        If score > alpha Then
            alpha = score
            bestMoveFound = mVal
            ' For Alpha-Node (Best Move), we might boost history too
            If cap = EMPTY_SQ Then HistoryMoves(sSq, eSq) = HistoryMoves(sSq, eSq) + depth
        End If
        
    Next i
    
    Negamax = alpha
End Function

' HELPER FOR SCORING
Function GetPieceValue(p As Integer) As Integer
    Select Case p
        Case 1, 7: GetPieceValue = 100
        Case 2, 8: GetPieceValue = 320
        Case 3, 9: GetPieceValue = 330
        Case 4, 10: GetPieceValue = 500
        Case 5, 11: GetPieceValue = 900
        Case 6, 12: GetPieceValue = 20000
        Case Else: GetPieceValue = 0
    End Select
End Function

'  VISUAL RENDERER
Public Sub RenderBoard()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Chess")
    
    Application.ScreenUpdating = False
    
    ' Clear Previous Render & Reset Font colors (Clears move indicator dots too)
    ws.Range("B2:I9").ClearContents
    ws.Range("B2:I9").Font.ColorIndex = xlAutomatic
    
    Dim row As Integer, col As Integer
    Dim piece As Integer
    Dim index As Integer
    Dim symbol As String
    
    ' Loop through visual rows 0 to 7 (Top to Bottom)
    For row = 0 To 7
        For col = 0 To 7
            
            ' Flipped view mapping for Black
            If HumanColor = 2 Then
                index = 21 + ((7 - row) * 10) + (7 - col)
            Else
                index = 21 + (row * 10) + col
            End If
            
            piece = Board(index)
            
            Select Case piece
                Case W_KING: symbol = ChrW(&H2654)
                Case W_QUEEN: symbol = ChrW(&H2655)
                Case W_ROOK: symbol = ChrW(&H2656)
                Case W_BISHOP: symbol = ChrW(&H2657)
                Case W_KNIGHT: symbol = ChrW(&H2658)
                Case W_PAWN: symbol = ChrW(&H2659)
                Case B_KING: symbol = ChrW(&H265A)
                Case B_QUEEN: symbol = ChrW(&H265B)
                Case B_ROOK: symbol = ChrW(&H265C)
                Case B_BISHOP: symbol = ChrW(&H265D)
                Case B_KNIGHT: symbol = ChrW(&H265E)
                Case B_PAWN: symbol = ChrW(&H265F)
                Case Else: symbol = ""
            End Select
            
            With ws.Cells(row + 2, col + 2)
                If symbol <> "" Then
                    .Value = symbol
                    .Font.Name = "Segoe UI Symbol"
                    .Font.Size = 24
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                End If
            End With
        Next col
    Next row
    
    Application.ScreenUpdating = True
End Sub

'  CLICK HANDLER
Public Sub HandleClick(r As Integer, c As Integer)
    ' Check if board is initialized (Board(0) == 99 means active)
    If Board(0) <> 99 Then
        ' THIS FIXES THE MENU: Now accepts clicks on rows 5, 6, 7, and 8
        If c = 11 And (r >= 5 And r <= 8) Then
            HandleMenuClick r, c
            Exit Sub
        Else
            MsgBox "Please select a game mode first!"
            ShowGameMenu
            Exit Sub
        End If
    End If
    
    ' Check if clicking control buttons
    If c = 11 And r >= 8 And r <= 10 Then
        HandleMenuClick r, c
        Exit Sub
    End If
    
    ' Regular game click logic
    Dim clickedIndex As Integer
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Chess")
    
    ' NEW FLIP LOGIC: Inverse the click coordinates if playing as Black
    Dim bRow As Integer, bCol As Integer
    If HumanColor = 2 Then
        bRow = 7 - (r - 2)
        bCol = 7 - (c - 2)
    Else
        bRow = (r - 2)
        bCol = (c - 2)
    End If
    
    clickedIndex = 21 + (bRow * 10) + bCol
    
    ' PART A: SELECTING A PIECE
    If SelectedSq = 0 Then
        If Board(clickedIndex) = 0 Then Exit Sub
        
        Dim pieceColor As Integer
        If Board(clickedIndex) <= 6 Then pieceColor = 1 Else pieceColor = 2
        If CurrentGameMode = MODE_PVAI Then
            If pieceColor <> HumanColor Then
                MsgBox "That is the computer's piece!"
                Exit Sub
            End If
        End If
        
        If pieceColor <> Turn Then
            MsgBox "It's not your turn!"
            Exit Sub
        End If
        
        ' Select and highlight
        SelectedSq = clickedIndex
        ApplyCheckerboard
        HighlightLegalMoves SelectedSq
        
    ' PART B: MOVING THE PIECE
    Else
        ' 1. Friendly Fire / Reselect Check
        Dim targetPiece As Integer
        targetPiece = Board(clickedIndex)
        If targetPiece <> 0 Then
            Dim targetColor As Integer
            If targetPiece <= 6 Then targetColor = 1 Else targetColor = 2
            If targetColor = Turn Then
                SelectedSq = clickedIndex
                RenderBoard ' Clear previous legal move dots
                ApplyCheckerboard ' Recalculate highlighting
                HighlightLegalMoves SelectedSq
                Exit Sub
            End If
        End If
        
        ' 2. Legal Move Check
        If IsLegalMove(SelectedSq, clickedIndex) = False Then
            MsgBox "Invalid Move!"
            SelectedSq = 0
            RenderBoard
            ApplyCheckerboard
            Exit Sub
        End If
        
        ' 3. EXECUTE MOVE
        Dim pieceMoved As Integer
        Dim captureFlag As Boolean
        pieceMoved = Board(SelectedSq)
        captureFlag = (Board(clickedIndex) <> EMPTY_SQ)
        
        ' RECORD THE HIGHLIGHT FOR THE MOVE MADE
        LastMoveStart = SelectedSq
        LastMoveEnd = clickedIndex
        
        ' UPDATE HISTORY
        moveHistory = moveHistory & SelectedSq & clickedIndex & " "
        
        ' A. Castling
        If GetType(pieceMoved) = W_KING And Abs(clickedIndex - SelectedSq) = 2 Then
            Board(clickedIndex) = Board(SelectedSq)
            Board(SelectedSq) = 0
            
            Dim rRow As Integer, rookFrom As Integer, rookTo As Integer
            rRow = GetRow(clickedIndex)
            If clickedIndex > SelectedSq Then ' Kingside
                rookFrom = (rRow + 2) * 10 + 8: rookTo = (rRow + 2) * 10 + 6
            Else ' Queenside
                rookFrom = (rRow + 2) * 10 + 1: rookTo = (rRow + 2) * 10 + 4
            End If
            Board(rookTo) = Board(rookFrom)
            Board(rookFrom) = 0
        
        ' B. En Passant
        ElseIf GetType(pieceMoved) = W_PAWN And clickedIndex = EnPassantTarget And Board(clickedIndex) = 0 Then
            Board(clickedIndex) = Board(SelectedSq)
            Board(SelectedSq) = 0
            Dim victimSq As Integer
            If Turn = 1 Then victimSq = clickedIndex + 10 Else victimSq = clickedIndex - 10
            Board(victimSq) = 0
            captureFlag = True
            
        ' C. Normal Move
        Else
            Board(clickedIndex) = Board(SelectedSq)
            Board(SelectedSq) = 0
        End If
        
        HandlePromotion clickedIndex
        RecordMove SelectedSq, clickedIndex, captureFlag
        
        ' 4. UPDATE FLAGS
        If pieceMoved = W_KING Then W_KingMoved = True
        If pieceMoved = B_KING Then B_KingMoved = True
        If SelectedSq = 91 Then W_RookLMoved = True
        If SelectedSq = 98 Then W_RookRMoved = True
        If SelectedSq = 21 Then B_RookLMoved = True
        If SelectedSq = 28 Then B_RookRMoved = True
        
        ' 5. SET EN PASSANT
        EnPassantTarget = 0
        If GetType(pieceMoved) = W_PAWN And Abs(clickedIndex - SelectedSq) = 20 Then
            EnPassantTarget = (SelectedSq + clickedIndex) / 2
        End If
        
        SelectedSq = 0
        
       ' 6. SWITCH TURN
        If Turn = 1 Then Turn = 2 Else Turn = 1
        
        ' --- AUTO-FLIP FOR PVP ---
        ' If playing 2-player mode, spin the board around to face the new player
        If CurrentGameMode = MODE_PVP Then
            HumanColor = Turn
        End If
        
        ' Render graphics and turn indicator BEFORE computer moves
        UpdateTurnUI
        RenderBoard
        ApplyCheckerboard
        HighlightChecks
        CheckGameStatus
        
        ' 7. TRIGGER COMPUTER MOVE (If applicable)
        If CurrentGameMode = MODE_PVAI And Turn <> HumanColor Then
            DoEvents ' Allows the "Thinking" UI to render
            MakeComputerMove
            
            ' Once the AI finishes its move, it will have switched the Turn back to the Human.
            ' We must update the UI one more time to reflect this!
            UpdateTurnUI
            RenderBoard
            ApplyCheckerboard
            HighlightChecks
            CheckGameStatus
        End If
    End If
End Sub

Function IsLegalMove(startSq As Integer, endSq As Integer) As Boolean
    Dim pVal As Integer, pType As Integer, pColor As Integer
    Dim r1 As Integer, c1 As Integer
    Dim r2 As Integer, c2 As Integer
    Dim dRow As Integer, dCol As Integer
    Dim targetContent As Integer
    Dim targetColor As Integer ' < New Variable
    
    pVal = Board(startSq)
    targetContent = Board(endSq)
    
    '  FIX: PREVENT CAPTURING OWN PIECES
    If targetContent <> 0 And targetContent <> OFF_BOARD Then
        ' Determine color of piece at start
        If pVal <= 6 Then pColor = 1 Else pColor = 2
        
        ' Determine color of piece at destination
        If targetContent <= 6 Then targetColor = 1 Else targetColor = 2
        
        ' If colors match, it is ILLEGAL (cannot move onto own piece)
        If pColor = targetColor Then
            IsLegalMove = False
            Exit Function
        End If
    End If
    pType = GetType(pVal)
    
    ' Determine Color (1=White, 2=Black)
    If pVal <= 6 Then pColor = 1 Else pColor = 2
    
    ' Get Coordinates & Deltas
    r1 = GetRow(startSq): c1 = GetCol(startSq)
    r2 = GetRow(endSq): c2 = GetCol(endSq)
    dRow = Abs(r2 - r1)
    dCol = Abs(c2 - c1)
    
    IsLegalMove = False ' Default to false
    
    Select Case pType
    
        Case W_KNIGHT
            ' L-Shape: (2,1) or (1,2). Knights Jump (No collision check needed)
            If (dRow = 2 And dCol = 1) Or (dRow = 1 And dCol = 2) Then IsLegalMove = True
            
        Case W_ROOK
            ' Straight line + Path Clear
            If (dRow = 0 Or dCol = 0) Then
                If IsPathClear(startSq, endSq) Then IsLegalMove = True
            End If
            
        Case W_BISHOP
            ' Diagonal + Path Clear
            If (dRow = dCol) Then
                If IsPathClear(startSq, endSq) Then IsLegalMove = True
            End If
            
        Case W_QUEEN
            ' Straight OR Diagonal + Path Clear
            If (dRow = 0 Or dCol = 0) Or (dRow = dCol) Then
                If IsPathClear(startSq, endSq) Then IsLegalMove = True
            End If
            
       Case W_KING
            ' Standard Move (1 square)
            If dRow <= 1 And dCol <= 1 Then IsLegalMove = True
            
            '  CASTLING LOGIC
            ' 1. Must be moving 2 squares sideways, 0 vertical
            If dRow = 0 And dCol = 2 Then
                Dim kMoved As Boolean, rMoved As Boolean
                Dim y As Integer ' The Array Row Index
                
                ' Get correct flags/row
                If pColor = 1 Then
                    kMoved = W_KingMoved
                    y = 9 ' White Row (Index 90s)
                Else
                    kMoved = B_KingMoved
                    y = 2 ' Black Row (Index 20s)
                End If
                
                If kMoved = False Then
                    ' KINGSIDE CASTLING (Targeting Col G / Index 6)
                    If c2 = 6 Then
                        If pColor = 1 Then rMoved = W_RookRMoved Else rMoved = B_RookRMoved
                        
                        ' A. Check Path Empty (F=Index 6, G=Index 7)
                        If rMoved = False And Board(y * 10 + 6) = 0 And Board(y * 10 + 7) = 0 Then
                            ' B. Check Safety: Current(E), Middle(F), Dest(G) must not be attacked
                            ' E = Col 5, F = Col 6, G = Col 7
                            If Not IsSquareAttacked(startSq, 3 - pColor) And _
                               Not IsSquareAttacked(y * 10 + 6, 3 - pColor) And _
                               Not IsSquareAttacked(y * 10 + 7, 3 - pColor) Then
                                IsLegalMove = True
                            End If
                        End If
                        
                    ' QUEENSIDE CASTLING (Targeting Col C / Index 2)
                    ElseIf c2 = 2 Then
                        If pColor = 1 Then rMoved = W_RookLMoved Else rMoved = B_RookLMoved
                        
                        ' A. Check Path Empty (B=Index 2, C=Index 3, D=Index 4)
                        If rMoved = False And Board(y * 10 + 2) = 0 And Board(y * 10 + 3) = 0 And Board(y * 10 + 4) = 0 Then
                            ' B. Check Safety: Current(E), Middle(D), Dest(C) must not be attacked
                            ' E=5, D=4, C=3
                            If Not IsSquareAttacked(startSq, 3 - pColor) And _
                               Not IsSquareAttacked(y * 10 + 4, 3 - pColor) And _
                               Not IsSquareAttacked(y * 10 + 3, 3 - pColor) Then
                                IsLegalMove = True
                            End If
                        End If
                    End If
                End If
            End If

        Case W_PAWN
            Dim direction As Integer, startRow As Integer
            If pColor = 1 Then
                direction = -1: startRow = 6
            Else
                direction = 1: startRow = 1
            End If
            
            ' 1. Forward Push
            If dCol = 0 And (r2 - r1) = direction Then
                If targetContent = 0 Then IsLegalMove = True
            End If
            
            ' 2. Double Push
            If dCol = 0 And (r2 - r1) = (2 * direction) Then
                If r1 = startRow And targetContent = 0 Then
                    If IsPathClear(startSq, endSq) Then IsLegalMove = True
                End If
            End If
            
            ' 3. Capture (Normal OR En Passant)
            If dCol = 1 And (r2 - r1) = direction Then
                ' Normal Capture
                If targetContent <> 0 Then IsLegalMove = True
                
                ' En Passant Capture
                ' Logic: The destination square is empty, BUT it matches the EnPassantTarget
                If targetContent = 0 And endSq = EnPassantTarget Then IsLegalMove = True
            End If
            
    End Select
    ' If the move is geometrically valid, we must ensure it doesn't leave King in check.
    If IsLegalMove = True Then
        
        ' 1. Make the move hypothetically
        Dim tempStart As Integer, tempEnd As Integer
        tempStart = Board(startSq)
        tempEnd = Board(endSq)
        
        Board(endSq) = tempStart
        Board(startSq) = EMPTY_SQ
        
        ' 2. Find My King
        Dim myKingIndex As Integer
        Dim i As Integer
        Dim kVal As Integer
        If pColor = 1 Then kVal = W_KING Else kVal = B_KING
        
        ' Search the board for the King (inefficient but safe for now)
        For i = 0 To 119
            If Board(i) = kVal Then
                myKingIndex = i
                Exit For
            End If
        Next i
        
        ' 3. Is King Attacked by Enemy?
        Dim enemyColor As Integer
        If pColor = 1 Then enemyColor = 2 Else enemyColor = 1
        
        If IsSquareAttacked(myKingIndex, enemyColor) Then
            IsLegalMove = False ' VETO: Move is illegal because King is exposed
        End If
        
        ' 4. Unmake the move (Restore board)
        Board(startSq) = tempStart
        Board(endSq) = tempEnd
        
    End If
End Function

' ==========================================
'  MOVE NOTATION & RECORDING
' ==========================================
Public Sub RecordMove(startSq As Integer, endSq As Integer, isCapture As Boolean)
    Dim pieceMoved As Integer
    Dim pType As Integer
    Dim moveText As String
    Dim startAlg As String, endAlg As String
    
    ' 1. Increment total move count (plies)
    MoveCount = MoveCount + 1
    
    ' 2. Get the piece that just moved
    pieceMoved = Board(endSq)
    pType = GetType(pieceMoved)
    
    ' Convert board indexes to "e2", "e4", etc. using your existing function
    startAlg = IndexToAlgebraic(startSq)
    endAlg = IndexToAlgebraic(endSq)
    
    ' 3. Format Algebraic Notation
    If pType = W_KING And Abs(endSq - startSq) = 2 Then
        ' Handle Castling Notation
        If endSq > startSq Then
            moveText = "O-O"
        Else
            moveText = "O-O-O"
        End If
    Else
        ' Get Piece Letter for standard moves
        Dim pLetter As String
        Select Case pType
            Case W_KNIGHT: pLetter = "N"
            Case W_BISHOP: pLetter = "B"
            Case W_ROOK:   pLetter = "R"
            Case W_QUEEN:  pLetter = "Q"
            Case W_KING:   pLetter = "K"
            Case W_PAWN:   pLetter = ""
        End Select
        
        If isCapture Then
            ' If a pawn captures, we show its starting file (e.g., exd5)
            If pType = W_PAWN Then pLetter = Left(startAlg, 1)
            moveText = pLetter & "x" & endAlg
        Else
            moveText = pLetter & endAlg
        End If
    End If
    
    ' 4. Store in engine memory
    If MoveCount <= UBound(MoveNotation) Then
        MoveNotation(MoveCount) = moveText
    End If
    
    ' 5. SEND TO THE NEW FANCY UI PANEL!
    LogMoveToUI moveText, MoveCount
End Sub

Public Sub ApplyCheckerboard()
    Dim r As Integer, c As Integer
    Dim ws As Worksheet
    Dim vRow As Integer, vCol As Integer
    Dim index As Integer
    Dim isLight As Boolean
    
    Set ws = ThisWorkbook.Sheets("Chess")
    
    For r = 2 To 9
        For c = 2 To 9
            ' Determine 120-array index to check against Highlights
            If HumanColor = 2 Then
                vRow = 7 - (r - 2)
                vCol = 7 - (c - 2)
            Else
                vRow = (r - 2)
                vCol = (c - 2)
            End If
            index = 21 + (vRow * 10) + vCol
            
            isLight = ((r + c) Mod 2 = 0)
            
            ' 1. Last Move Highlighting (Yellowish-Green)
            If index = LastMoveStart Or index = LastMoveEnd Then
                If isLight Then
                    ws.Cells(r, c).Interior.color = RGB(246, 246, 105)
                Else
                    ws.Cells(r, c).Interior.color = RGB(186, 202, 67)
                End If
                
            ' 2. Selected Piece Highlighting
            ElseIf index = SelectedSq Then
                If isLight Then
                    ws.Cells(r, c).Interior.color = RGB(214, 214, 150)
                Else
                    ws.Cells(r, c).Interior.color = RGB(150, 170, 100)
                End If
                
            ' 3. Standard Checkerboard (Lichess/Chess.com standard)
            Else
                If isLight Then
                    ws.Cells(r, c).Interior.color = RGB(238, 238, 210) ' Cream
                Else
                    ws.Cells(r, c).Interior.color = RGB(118, 150, 86)  ' Green
                End If
            End If
        Next c
    Next r
End Sub

Public Sub HighlightLegalMoves(startSq As Integer)
    Dim endSq As Integer
    Dim vRow As Integer, vCol As Integer
    Dim r As Integer, c As Integer
    Dim isCapture As Boolean
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Chess")
    
    Application.ScreenUpdating = False
    
    For endSq = 21 To 98
        If Board(endSq) <> OFF_BOARD Then
            If IsLegalMove(startSq, endSq) Then
                ' Calculate Visual Coordinates
                vRow = GetRow(endSq)
                vCol = GetCol(endSq)
                
                If HumanColor = 2 Then
                    r = (7 - vRow) + 2
                    c = (7 - vCol) + 2
                Else
                    r = vRow + 2
                    c = vCol + 2
                End If
                
                ' Check if square involves taking an enemy
                isCapture = (Board(endSq) <> EMPTY_SQ)
                
                ' Extra logic for visually confirming En Passant Captures
                If Board(endSq) = EMPTY_SQ And endSq = EnPassantTarget And GetType(Board(startSq)) = W_PAWN Then
                    isCapture = True
                End If
                
                ' -- DRAW HIGHLIGHTS --
                If isCapture Then
                    ' Red target background for Captures
                    If (r + c) Mod 2 = 0 Then
                        ws.Cells(r, c).Interior.color = RGB(225, 125, 115) ' Light Square Red
                    Else
                        ws.Cells(r, c).Interior.color = RGB(210, 100, 80) ' Dark Square Red
                    End If
                Else
                    ' Empty Square: Draw a Lichess-style dark dot
                    With ws.Cells(r, c)
                        .Value = ChrW(&H25CF) ' Big Dot character
                        .Font.Name = "Arial"
                        .Font.color = RGB(80, 100, 60) ' Semi-transparent green appearance
                        .Font.Size = 14
                    End With
                End If
            End If
        End If
    Next endSq
    
    Application.ScreenUpdating = True
End Sub

Function IsPathClear(startSq As Integer, endSq As Integer) As Boolean
    Dim diff As Integer
    Dim step As Integer
    Dim curr As Integer
    
    diff = endSq - startSq

    If diff Mod 10 = 0 Then
        ' Vertical Move
        step = 10 * Sgn(diff)
        
    ElseIf diff Mod 11 = 0 Then
        ' Diagonal
        step = 11 * Sgn(diff)
        
    ElseIf diff Mod 9 = 0 Then
        ' Diagonal
        step = 9 * Sgn(diff)
        
    ElseIf (diff \ 10) = 0 Then
        ' Horizontal Move
        ' We check this LAST because 9 \ 10 is 0, which would confuse a diagonal for a horizontal
        step = 1 * Sgn(diff)
        
    Else
        ' Should not happen for valid Rook/Bishop moves
        IsPathClear = False
        Exit Function
    End If
    
    ' Loop from the square AFTER start, up to the square BEFORE end
    curr = startSq + step
    While curr <> endSq
        If Board(curr) <> 0 Then
            IsPathClear = False
            Exit Function
        End If
        curr = curr + step
    Wend
    
    IsPathClear = True
End Function


' Returns TRUE if the square 'sq' is being attacked by 'attackerColor'
Function IsSquareAttacked(sq As Integer, attackerColor As Integer) As Boolean
    Dim r As Integer, c As Integer
    Dim step As Integer, curr As Integer
    Dim piece As Integer, pType As Integer, pColor As Integer
    Dim dir As Variant
    Dim i As Integer
    
    ' 1. Check Knight Attacks (8 jumps)
    Dim kJumps As Variant
    kJumps = Array(8, 12, 19, 21, -8, -12, -19, -21)
    
    For i = LBound(kJumps) To UBound(kJumps)
        curr = sq + kJumps(i)
        
        '  SAFETY CHECK 1: Ensure index is within array limits
        If curr >= 0 And curr <= 119 Then
            piece = Board(curr)
            If piece <> OFF_BOARD And piece <> EMPTY_SQ Then
                If piece <= 6 Then pColor = 1 Else pColor = 2
                pType = GetType(piece)
                
                If pColor = attackerColor And pType = W_KNIGHT Then
                    IsSquareAttacked = True
                    Exit Function
                End If
            End If
        End If
    Next i
    
    ' 2. Check Rays (Rooks/Bishops/Queens)
    Dim rayDirs As Variant
    rayDirs = Array(-10, 10, 1, -1, -9, -11, 11, 9)
    
    For i = LBound(rayDirs) To UBound(rayDirs)
        step = rayDirs(i)
        curr = sq + step
        
        ' Loop while inside valid array bounds
        Do While curr >= 0 And curr <= 119
            If Board(curr) = OFF_BOARD Then Exit Do
            
            piece = Board(curr)
            
            If piece <> EMPTY_SQ Then
                If piece <= 6 Then pColor = 1 Else pColor = 2
                pType = GetType(piece)
                
                If pColor = attackerColor Then
                    If i <= 3 Then ' Straight Ray
                        If pType = W_ROOK Or pType = W_QUEEN Then
                            IsSquareAttacked = True
                            Exit Function
                        End If
                    Else ' Diagonal Ray
                        If pType = W_BISHOP Or pType = W_QUEEN Then
                            IsSquareAttacked = True
                            Exit Function
                        End If
                    End If
                End If
                Exit Do ' Stop ray at first piece
            End If
            
            curr = curr + step
        Loop
    Next i
    
    ' 3. Check Pawns
    Dim pDir As Integer
    If attackerColor = 1 Then pDir = 10 Else pDir = -10
    
    Dim pSq As Variant
    pSq = Array(sq + pDir - 1, sq + pDir + 1)
    
    For i = 0 To 1
        curr = pSq(i)
        '  SAFETY CHECK 2: Pawn lookups can also go out of bounds
        If curr >= 0 And curr <= 119 Then
            piece = Board(curr)
            If piece <> OFF_BOARD And piece <> EMPTY_SQ Then
                If piece <= 6 Then pColor = 1 Else pColor = 2
                If pColor = attackerColor And GetType(piece) = W_PAWN Then
                    IsSquareAttacked = True
                    Exit Function
                End If
            End If
        End If
    Next i
    
    ' 4. Check King (Adjacent)
    Dim kDirs As Variant
    kDirs = Array(-10, -9, -11, -1, 1, 9, 10, 11)
    For i = LBound(kDirs) To UBound(kDirs)
        curr = sq + kDirs(i)
        If curr >= 0 And curr <= 119 Then
            piece = Board(curr)
            If piece <> OFF_BOARD And piece <> EMPTY_SQ Then
                If piece <= 6 Then pColor = 1 Else pColor = 2
                If pColor = attackerColor And GetType(piece) = W_KING Then
                    IsSquareAttacked = True
                    Exit Function
                End If
            End If
        End If
    Next i
    
    IsSquareAttacked = False
End Function


Public Sub HighlightChecks()
    Dim i As Integer, piece As Integer
    Dim r As Integer, c As Integer
    Dim vRow As Integer, vCol As Integer
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Chess")
    
    ' Scan the board for Kings
    For i = 21 To 98
        piece = Board(i)
        
        If (piece = W_KING And IsSquareAttacked(i, 2)) Or (piece = B_KING And IsSquareAttacked(i, 1)) Then
            vRow = GetRow(i)
            vCol = GetCol(i)
            
            ' Flipped mapping for red Check highlights
            If HumanColor = 2 Then
                r = (7 - vRow) + 2
                c = (7 - vCol) + 2
            Else
                r = vRow + 2
                c = vCol + 2
            End If
            
            ws.Cells(r, c).Interior.color = RGB(255, 100, 100) ' Soft Red Warning
        End If
    Next i
End Sub

Function HasLegalMoves(colorToCheck As Integer) As Boolean
    Dim startSq As Integer, endSq As Integer
    Dim piece As Integer, pColor As Integer
    
    ' Loop through every square on the board (Source)
    For startSq = 21 To 98
        piece = Board(startSq)
        
        ' Is there a piece here?
        If piece <> OFF_BOARD And piece <> EMPTY_SQ Then
            
            ' Is it My piece?
            If piece <= 6 Then pColor = 1 Else pColor = 2
            
            If pColor = colorToCheck Then
                
                ' Try moving it to every other square (Destination)
                ' (Note: This is brute force, but simplest for now)
                For endSq = 21 To 98
                    
                    ' Optimization: Don't check OFF_BOARD squares
                    If Board(endSq) <> OFF_BOARD Then
                        
                        ' Is this specific move legal? (Includes King Safety check)
                        If IsLegalMove(startSq, endSq) = True Then
                            ' We found a move! The game is alive.
                            HasLegalMoves = True
                            Exit Function
                        End If
                        
                    End If
                Next endSq
                
            End If
        End If
    Next startSq
    
    ' If we get here, we checked every piece and found 0 moves.
    HasLegalMoves = False
End Function

Public Sub CheckGameStatus()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Chess")
    
    ' 1. Record current board state for Threefold Repetition
    RecordPosition
    
    ' 2. Check Draw Rules
    If IsInsufficientMaterial() Then
        ShowGameOverMenu "DRAW!", "Insufficient Material"
        Board(0) = 0 ' Ends game
        Exit Sub
    End If
    
    If IsThreefoldRepetition() Then
        ShowGameOverMenu "DRAW!", "Threefold Repetition"
        Board(0) = 0 ' Ends game
        Exit Sub
    End If
    
    ' 3. Check for Checkmate or Stalemate
    If HasLegalMoves(Turn) = False Then
        
        Dim kIndex As Integer, i As Integer, kVal As Integer
        If Turn = 1 Then kVal = W_KING Else kVal = B_KING
        
        For i = 21 To 98
            If Board(i) = kVal Then
                kIndex = i
                Exit For
            End If
        Next i
        
        Dim enemyColor As Integer: enemyColor = 3 - Turn
        
        If IsSquareAttacked(kIndex, enemyColor) Then
            ShowGameOverMenu "CHECKMATE!", IIf(Turn = 2, "White", "Black") & " Wins!"
        Else
            ShowGameOverMenu "STALEMATE!", "Game is a Draw."
        End If
        
        Board(0) = 0 ' Turn off game flag so menu buttons work again
    End If
End Sub

Public Sub HandlePromotion(sq As Integer)
    Dim row As Integer
    Dim piece As Integer
    Dim pColor As Integer
    
    piece = Board(sq)
    row = GetRow(sq) ' Returns Visual Row 0 to 7
    
    ' 1. Check if it is a Pawn
    If GetType(piece) <> W_PAWN Then Exit Sub
    
    ' 2. Check if it is on the Back Rank (Row 0 for White, Row 7 for Black)
    If row <> 0 And row <> 7 Then Exit Sub
    
    ' 3. Determine Color
    If piece <= 6 Then pColor = 1 Else pColor = 2
    
    ' 4. AI LOGIC (Auto-Queen)
    ' If it is the AI's turn (Turn = 2) or we are running a simulation, just pick Queen.
    If pColor = 2 Then
        Board(sq) = B_QUEEN
        Exit Sub
    End If
    
    ' 5. HUMAN LOGIC (Input Box)
    Dim choice As String
    Dim newPiece As Integer
    
    ' Loop until valid input
    Do
        choice = InputBox("Promote to? Enter Q, R, B, or N", "Pawn Promotion", "Q")
        choice = UCase(Left(choice, 1))
    Loop Until InStr("QRBN", choice) > 0
    
    ' Convert Letter to Piece Constant
    Select Case choice
        Case "R": newPiece = W_ROOK
        Case "B": newPiece = W_BISHOP
        Case "N": newPiece = W_KNIGHT
        Case "Q": newPiece = W_QUEEN
    End Select
    
  
    If pColor = 1 Then
        Board(sq) = newPiece
    Else
        Board(sq) = newPiece + 6
    End If
End Sub

' --- DRAW RULE HELPER FUNCTIONS ---

Private Sub RecordPosition()
    Dim i As Integer, s As String
    ' Create a string representation of the board + whose turn it is
    For i = 21 To 98
        If Board(i) <> OFF_BOARD Then s = s & Board(i) & ","
    Next i
    s = s & Turn
    
    PositionCount = PositionCount + 1
    If PositionCount <= 500 Then
        PositionHistory(PositionCount) = s
    End If
End Sub

Private Function IsThreefoldRepetition() As Boolean
    If PositionCount < 5 Then Exit Function ' Needs at least 5 turns to repeat 3 times
    
    Dim currentPos As String
    currentPos = PositionHistory(PositionCount)
    
    Dim matchCount As Integer
    matchCount = 0
    
    Dim i As Integer
    For i = 1 To PositionCount
        If PositionHistory(i) = currentPos Then
            matchCount = matchCount + 1
        End If
    Next i
    
    If matchCount >= 3 Then IsThreefoldRepetition = True Else IsThreefoldRepetition = False
End Function

Private Function IsInsufficientMaterial() As Boolean
    Dim i As Integer, p As Integer
    Dim knightCount As Integer, bishopCount As Integer
    Dim majorPiece As Boolean
    
    For i = 21 To 98
        p = Board(i)
        If p <> EMPTY_SQ And p <> OFF_BOARD Then
            Select Case p
                Case 1, 7, 4, 10, 5, 11 ' Pawns, Rooks, Queens
                    majorPiece = True
                    Exit For
                Case 2, 8 ' Knights
                    knightCount = knightCount + 1
                Case 3, 9 ' Bishops
                    bishopCount = bishopCount + 1
            End Select
        End If
    Next i
    
    ' If there's a Pawn, Rook, or Queen, checkmate is still possible
    If majorPiece Then
        IsInsufficientMaterial = False
    ' If just Kings, King+Knight, or King+Bishop remain, checkmate is impossible
    ElseIf (knightCount + bishopCount) <= 1 Then
        IsInsufficientMaterial = True
    Else
        IsInsufficientMaterial = False
    End If
End Function

' --- HELPER FUNCTIONS ---

' Simplify piece value to Type (1=Pawn, 2=Knight... regardless of color)
Public Function GetType(pieceVal As Integer) As Integer
    If pieceVal > 6 Then
        GetType = pieceVal - 6
    Else
        GetType = pieceVal
    End If
End Function

' Get the Row (0-7) from the 10x12 Index
Public Function GetRow(idx As Integer) As Integer
    GetRow = (idx \ 10) - 2
End Function

' Get the Col (0-7) from the 10x12 Index
Public Function GetCol(idx As Integer) As Integer
    GetCol = (idx Mod 10) - 1
End Function

