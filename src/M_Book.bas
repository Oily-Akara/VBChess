Attribute VB_Name = "M_Book"


Public Function GetBookMove(turnColor As Integer) As Long
    If turnColor = 1 Then
        GetBookMove = 0
        Exit Function ' We only play as Black
    End If
    
    ' Parse the move history to understand what White played
    Dim moveHist As String
    moveHist = Trim(moveHistory) ' Use the global variable from M_AI
    
    ' Count moves by counting spaces + 1
    Dim MoveCount As Integer
    If moveHist = "" Then
        MoveCount = 0
    Else
        Dim moves() As String
        moves = Split(moveHist)
        MoveCount = UBound(moves) + 1
    End If
    
    ' Get Black's move number (we move on odd indices: 1, 3, 5...)
    Dim blackMoveNum As Integer
    blackMoveNum = (MoveCount + 1) \ 2
    
    ' =
    ' KING'S INDIAN DEFENSE MAIN LINE VS 1.d4
    ' =
    
    ' MOVE 1: After 1.d4, play 1...Nf6
    If blackMoveNum = 1 Then
        If Board(27) = 8 And Board(46) = 0 Then
            GetBookMove = 27046 ' Nf6
            Exit Function
        End If
    End If
    
    ' MOVE 2: After 2.c4 (or most moves), play 2...g6
    If blackMoveNum = 2 Then
        If Board(37) = 7 And Board(47) = 0 Then
            GetBookMove = 37047 ' g6
            Exit Function
        End If
    End If
    
    ' MOVE 3: After 3.Nc3 (or g3, Nf3), play 3...Bg7
    If blackMoveNum = 3 Then
        If Board(26) = 9 And Board(37) = 0 And Board(47) = 7 Then
            GetBookMove = 26037 ' Bg7
            Exit Function
        End If
    End If
    
    ' MOVE 4: Play 4...d6 (universal KID setup)
    If blackMoveNum = 4 Then
        If Board(34) = 7 And Board(44) = 0 Then
            GetBookMove = 34044 ' d6
            Exit Function
        End If
    End If
    
    ' MOVE 5: Play 5...0-0 (castle kingside)
    If blackMoveNum = 5 Then
        If Board(25) = 12 And Board(28) = 10 Then
            If Board(26) = 0 And Board(27) = 0 Then
                GetBookMove = 25027 ' O-O
                Exit Function
            End If
        End If
    End If
    
    ' =
    ' MOVE 6+: BRANCHING BASED ON WHITE'S SYSTEM
    ' =
    
    If blackMoveNum = 6 Then
        ' Check if White played e4 (Classical/Saemisch)
        If Board(75) = 1 Then ' White pawn on e4
            ' Classical Main Line: 6...e5
            If Board(35) = 7 And Board(55) = 0 Then
                GetBookMove = 35055 ' e5
                Exit Function
            End If
        Else
            ' Against other systems (Fianchetto, London, etc): 6...Nbd7
            If Board(22) = 8 And Board(34) = 0 Then
                GetBookMove = 22034 ' Nbd7
                Exit Function
            End If
        End If
    End If
    
    If blackMoveNum = 7 Then
        ' If White played Nf3 and e4: Classical Main Line
        If Board(75) = 1 Then ' e4 on board
            ' Play 7...Nc6 (Gligoric System)
            If Board(22) = 8 And Board(43) = 0 Then
                GetBookMove = 22043 ' Nc6
                Exit Function
            ' Alternative: 7...Na6 (Modern Treatment)
            ElseIf Board(22) = 8 And Board(31) = 0 Then
                GetBookMove = 22031 ' Na6
                Exit Function
            ' Fallback: 7...Nbd7
            ElseIf Board(22) = 8 And Board(34) = 0 Then
                GetBookMove = 22034 ' Nbd7
                Exit Function
            End If
        Else
            ' Against Fianchetto/Catalan: 7...c5 (counterplay)
            If Board(33) = 7 And Board(53) = 0 Then
                GetBookMove = 33053 ' c5
                Exit Function
            End If
        End If
    End If
    
    If blackMoveNum = 8 Then
        ' If playing Classical Main Line (e5 already played)
        If Board(55) = 7 Then ' Black pawn on e5
            ' If White has d5 (closed center)
            If Board(64) = 1 Then ' White pawn on d5
                ' Play 8...Ne8 (preparing f5 break)
                If Board(46) = 8 And Board(25) = 0 Then
                    GetBookMove = 46025 ' Ne8
                    Exit Function
                End If
            Else
                ' Center still fluid: 8...exd4 or 8...Bg4
                ' Simplified: play 8...Bg4 (pin the knight)
                If Board(26) = 9 And Board(46) = 0 Then
                    GetBookMove = 26046 ' Bg4
                    Exit Function
                End If
            End If
        Else
            ' Against other systems: 8...e5 (universal break)
            If Board(35) = 7 And Board(55) = 0 Then
                GetBookMove = 35055 ' e5
                Exit Function
            End If
        End If
    End If
    
    If blackMoveNum = 9 Then
        ' Mar Del Plata Attack: 9...f5 (THE KING'S INDIAN PAWN STORM!)
        If Board(64) = 1 And Board(55) = 7 Then ' Closed center (d5 + e5)
            If Board(36) = 7 And Board(56) = 0 Then
                GetBookMove = 36056 ' f5!
                Exit Function
            End If
        Else
            ' Open/Semi-open: Develop with 9...Nd7
            If Board(46) = 8 And Board(34) = 0 Then
                GetBookMove = 46034 ' Nd7
                Exit Function
            End If
        End If
    End If
    
    If blackMoveNum = 10 Then
        ' After f5, continue with 10...Nf6 (typical maneuver)
        If Board(25) = 8 And Board(46) = 0 Then
            GetBookMove = 25046 ' Nf6 (from e8)
            Exit Function
        ' Or 10...c6 (support the d6 pawn)
        ElseIf Board(33) = 7 And Board(43) = 0 Then
            GetBookMove = 33043 ' c6
            Exit Function
        End If
    End If
    
    ' =
    ' MOVES 11-15: MIDDLEGAME PREPARATION
    ' =
    
    If blackMoveNum >= 11 And blackMoveNum <= 15 Then
        ' Priority 1: Complete f5 break if not done
        If Board(36) = 7 And Board(56) = 0 And Board(55) = 7 Then
            GetBookMove = 36056 ' f5
            Exit Function
        End If
        
        ' Priority 2: Play c5 if not done (counterplay)
        If Board(33) = 7 And Board(53) = 0 Then
            GetBookMove = 33053 ' c5
            Exit Function
        End If
        
        ' Priority 3: Develop Queen to e8 (supports f5)
        If Board(24) = 11 And Board(25) = 0 And Board(27) = 12 Then
            GetBookMove = 24025 ' Qe8
            Exit Function
        End If
        
        ' Priority 4: Advance g5-g4 attack
        If Board(47) = 7 And Board(57) = 0 And Board(56) = 7 Then
            GetBookMove = 47057 ' g5
            Exit Function
        End If
        
        ' Priority 5: Play a6 + b5 queenside expansion
        If Board(31) = 7 And Board(41) = 0 Then
            GetBookMove = 31041 ' a6
            Exit Function
        End If
        
        If Board(32) = 7 And Board(52) = 0 And Board(41) = 7 Then
            GetBookMove = 32052 ' b5
            Exit Function
        End If
        
        ' Priority 6: Rook to f8 (support f-pawn)
        If Board(28) = 10 And Board(26) = 0 Then
            GetBookMove = 28026 ' Rf8
            Exit Function
        End If
    End If
    
    ' =
    ' ANTI-SYSTEMS REPERTOIRE (Moves 4-8)
    ' =
    
    ' VS LONDON SYSTEM (Bf4 early): Accelerate ...e5
    If blackMoveNum >= 4 And blackMoveNum <= 6 Then
        ' Check if White has Bf4
        If Board(76) = 3 Or Board(66) = 3 Then ' Bishop on f4 or c4
            ' Play ...e5 immediately to challenge
            If Board(35) = 7 And Board(55) = 0 Then
                GetBookMove = 35055 ' e5!
                Exit Function
            End If
        End If
    End If
    
    ' VS FIANCHETTO (g3 + Bg2): Play ...c5 early
    If blackMoveNum >= 5 And blackMoveNum <= 7 Then
        ' Check if White has g3/Bg2 setup
        If Board(87) = 1 Or Board(77) = 3 Then ' g3 pawn or Bg2
            If Board(33) = 7 And Board(53) = 0 Then
                GetBookMove = 33053 ' c5 (immediate counterplay)
                Exit Function
            End If
        End If
    End If
    
    ' VS SAEMISCH (f3): Prepare ...e5 and ...exf3
    If blackMoveNum >= 6 And blackMoveNum <= 8 Then
        ' Check if White played f3
        If Board(86) = 1 Then ' f3 pawn
            ' Play ...e5 to challenge
            If Board(35) = 7 And Board(55) = 0 Then
                GetBookMove = 35055 ' e5
                Exit Function
            End If
            ' Then prepare ...c5
            If Board(33) = 7 And Board(53) = 0 And Board(55) = 7 Then
                GetBookMove = 33053 ' c5
                Exit Function
            End If
        End If
    End If
    
    ' =
    ' TACTICAL MOTIFS (Moves 10+)
    ' =
    
    If blackMoveNum >= 10 Then
        ' MOTIF 1: If White's e4 pawn is undefended, consider ...fxe4
        If Board(75) = 1 And Board(56) = 7 Then ' f5 pawn can capture
            ' Check if e4 is really hanging (simplified check)
            If Board(74) = 0 And Board(76) = 0 Then ' No pieces defending
                GetBookMove = 56075 ' fxe4!
                Exit Function
            End If
        End If
        
        ' MOTIF 2: If White castled kingside, advance g5-g4-g3
        If Board(98) = 6 Then ' White King on g1 (castled)
            If Board(47) = 7 And Board(57) = 0 Then
                GetBookMove = 47057 ' g5
                Exit Function
            End If
            If Board(57) = 7 And Board(67) = 0 Then
                GetBookMove = 57067 ' g4
                Exit Function
            End If
        End If
        
        ' MOTIF 3: Trade dark-squared bishops if White has one on g5
        If Board(57) = 3 Then ' White Bg5
            If Board(37) = 9 And Board(46) = 0 Then
                GetBookMove = 37046 ' Bxf6 ideas
                Exit Function
            End If
        End If
    End If
    
  
    ' ENDGAME TRANSITIONS (Moves 15+)
   
    
    If blackMoveNum >= 15 Then
        ' If trades occurred, activate King
        If Board(27) = 12 Then ' King still on g8
            ' Move to f7 for activity
            If Board(36) = 0 Then
                GetBookMove = 27036 ' Kf7
                Exit Function
            End If
        End If
        
        ' Push passed pawns
        If Board(52) = 7 And Board(62) = 0 Then
            GetBookMove = 52062 ' b4
            Exit Function
        End If
        
        If Board(53) = 7 And Board(63) = 0 Then
            GetBookMove = 53063 ' c4
            Exit Function
        End If
    End If
    
 
    ' EXIT BOOK - LET ENGINE SEARCH
  
    GetBookMove = 0
End Function
