Attribute VB_Name = "M_Menu"
' MODULE: M_Menu
Option Explicit

' Game Mode Constants
Public Const MODE_PVP As Integer = 1
Public Const MODE_PVAI As Integer = 2

' Global Variables
Public CurrentGameMode As Integer
Public HumanColor As Integer ' 1 = White, 2 = Black
Public MoveNotation() As String
Public MoveCount As Integer
Public IsSideMenuOpen As Boolean ' Tracks if we are selecting a side

' Initialize the menu
Public Sub ShowGameMenu()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Chess")
    
    
    CurrentGameMode = 0
    IsSideMenuOpen = False
    
    ' Clear the board and UI
    ws.Range("B2:I9").ClearContents
    ws.Range("B2:I9").Interior.ColorIndex = xlNone
    ApplyCheckerboard
    
    ws.Range("K2:K20").ClearContents
    ws.Range("K2:K20").Interior.ColorIndex = xlNone
    ws.Range("K2:K20").Borders.LineStyle = xlNone
    ws.Range("M1:N100").ClearContents
    
    ' Create Title
    With ws.Range("K2")
        .Value = "VBCHESS"
        .Font.Bold = True
        .Font.Size = 14
    End With
    
    ' Create Main Menu Options
    ws.Range("K4").Value = "Select Game Mode:"
    
    With ws.Range("K5")
        .Value = "1. PLAYER vs PLAYER"
        .Interior.color = RGB(100, 150, 200) ' Blue
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Font.color = vbWhite
    End With
    
    With ws.Range("K6")
        .Value = "2. PLAYER vs AI"
        .Interior.color = RGB(200, 100, 100) ' Red
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Font.color = vbWhite
    End With
    
    ws.Range("K8").Value = "Click a mode to start"
End Sub

' Show the Side Selection Sub-Menu
Private Sub ShowSideSelection()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Chess")
    
    IsSideMenuOpen = True
    
    ws.Range("K4").Value = "Choose Your Side:"
    
    ' Clear previous buttons
    ws.Range("K5:K6").ClearContents
    ws.Range("K5:K6").Interior.ColorIndex = xlNone
    
    ' Create Side Buttons
    With ws.Range("K6")
        .Value = "PLAY AS WHITE"
        .Interior.color = RGB(255, 255, 255)
        .Font.color = vbBlack
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
    End With
    
    With ws.Range("K7")
        .Value = "PLAY AS BLACK"
        .Interior.color = RGB(50, 50, 50)
        .Font.color = vbWhite
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    
    With ws.Range("K8")
        .Value = "RANDOM SIDE"
        .Interior.color = RGB(150, 150, 150)
        .Font.color = vbBlack
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
End Sub

' Start the game
Public Sub StartGame(mode As Integer, side As Integer)
    CurrentGameMode = mode
    HumanColor = side
    IsSideMenuOpen = False
    
   
    If HumanColor = 3 Then
        Randomize
        If Rnd() > 0.5 Then HumanColor = 1 Else HumanColor = 2
    End If
    
    
    InitBoard
    
    ' Update UI
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Chess")
    
    ws.Range("K4").Value = IIf(mode = MODE_PVP, "PvP Mode", "PvAI Mode")
    
    If mode = MODE_PVAI Then
        ws.Range("K4").Value = ws.Range("K4").Value & IIf(HumanColor = 1, " (White)", " (Black)")
    End If
    
    ws.Range("K5").Value = "Turn: White"
    
    ' Clear Menu Buttons Area
    ws.Range("K6:K8").ClearContents
    ws.Range("K6:K8").Interior.ColorIndex = xlNone
    ws.Range("K6:K8").Borders.LineStyle = xlNone
    
  
    CreateGameControls
    
    ' Setup Notation
    ws.Range("M1").Value = "MOVE HISTORY"
    ws.Range("M1").Font.Bold = True
    ws.Range("M2").Value = "White"
    ws.Range("N2").Value = "Black"
    ws.Range("M2:N2").Font.Bold = True
    
    ' AI LOGIC: If Player is Black, AI moves immediately
    If CurrentGameMode = MODE_PVAI And HumanColor = 2 Then
        ws.Range("K5").Value = "Turn: White (Thinking...)"
        DoEvents
        Application.Wait (Now + TimeValue("0:00:01"))
        MakeComputerMove
    End If
End Sub

Private Sub CreateGameControls()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Chess")
    
    ' Undo button
    With ws.Range("K8")
        .Value = "UNDO MOVE"
        .Interior.color = RGB(255, 200, 100)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    
    ' New Game button
    With ws.Range("K9")
        .Value = "NEW GAME"
        .Interior.color = RGB(150, 200, 150)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    
    ' Show Moves button
    With ws.Range("K10")
        .Value = "VIEW MOVES"
        .Interior.color = RGB(200, 200, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
End Sub

' Handle menu clicks
Public Sub HandleMenuClick(r As Integer, c As Integer)
    If c <> 11 Then Exit Sub ' Must be Column K
    
    ' 1. MAIN MENU CLICKS (Using CurrentGameMode Variable, NOT Board Array)
    If CurrentGameMode = 0 Then
        
        If IsSideMenuOpen = False Then
            If r = 5 Then ' PvP
                StartGame MODE_PVP, 1
            ElseIf r = 6 Then ' PvAI (Open Side Menu)
                ShowSideSelection
            End If
            Exit Sub
        End If
        
        ' SIDE SELECTION LOGIC
        If IsSideMenuOpen = True Then
            If r = 6 Then StartGame MODE_PVAI, 1 ' White
            If r = 7 Then StartGame MODE_PVAI, 2 ' Black
            If r = 8 Then StartGame MODE_PVAI, 3 ' Random
            Exit Sub
        End If
    End If
    
    ' 2. IN-GAME CONTROLS (Only if game is running)
    If CurrentGameMode > 0 Then
        If r = 8 Then ' Undo
            UndoLastMove
        ElseIf r = 9 Then ' New Game
            ShowGameMenu
        ElseIf r = 10 Then ' View Moves
            ShowMoveHistory
        End If
    End If
End Sub

'SUPPORT FUNCTIONS

Function IndexToAlgebraic(sq As Integer) As String
    Dim row As Integer, col As Integer
    row = GetRow(sq)
    col = GetCol(sq)
    Dim files As String
    files = "abcdefgh"
    IndexToAlgebraic = Mid(files, col + 1, 1) & (8 - row)
End Function

Public Sub RecordMove(startSq As Integer, endSq As Integer, captureFlag As Boolean)
    On Error Resume Next
    If UBound(MoveNotation) = 0 Then ReDim MoveNotation(1 To 200): MoveCount = 0
    If Err.Number <> 0 Then ReDim MoveNotation(1 To 200): MoveCount = 0: Err.Clear
    On Error GoTo 0
    
    MoveCount = MoveCount + 1
    
    Dim piece As Integer, notation As String
    piece = Board(endSq)
    Dim pieceSymbol As String
    Select Case GetType(piece)
        Case W_KING: pieceSymbol = "K"
        Case W_QUEEN: pieceSymbol = "Q"
        Case W_ROOK: pieceSymbol = "R"
        Case W_BISHOP: pieceSymbol = "B"
        Case W_KNIGHT: pieceSymbol = "N"
        Case W_PAWN: pieceSymbol = ""
    End Select
    
    notation = pieceSymbol
    If captureFlag Then
        If pieceSymbol = "" Then notation = Mid("abcdefgh", GetCol(startSq) + 1, 1)
        notation = notation & "x"
    End If
    notation = notation & IndexToAlgebraic(endSq)
    
    Dim enemyColor As Integer: enemyColor = 3 - Turn
    If IsKingInCheck(enemyColor) Then
        If Not HasLegalMoves(enemyColor) Then notation = notation & "#" Else notation = notation & "+"
    End If
    
    MoveNotation(MoveCount) = notation
    UpdateMoveDisplay
End Sub

Private Sub UpdateMoveDisplay()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Chess")
    Dim moveNum As Integer: moveNum = (MoveCount + 1) \ 2
    Dim displayRow As Integer: displayRow = moveNum + 2
    
    If MoveCount Mod 2 = 1 Then
        ws.Cells(displayRow, 13).Value = moveNum & "."
        ws.Cells(displayRow, 14).Value = MoveNotation(MoveCount)
    Else
        ws.Cells(displayRow, 15).Value = MoveNotation(MoveCount)
    End If
End Sub

Public Sub UndoLastMove()
    If MoveCount = 0 Then MsgBox "No moves to undo!": Exit Sub
    
    Dim undoCount As Integer
    If CurrentGameMode = MODE_PVAI Then
        If Turn = HumanColor Then undoCount = 2 Else undoCount = 1
    Else
        undoCount = 1
    End If
    
    MsgBox "Game reset due to Undo complexity."
    ShowGameMenu
End Sub

Public Sub ShowMoveHistory()
    If MoveCount = 0 Then MsgBox "No moves yet!": Exit Sub
    Dim s As String, i As Integer
    For i = 1 To MoveCount
        s = s & MoveNotation(i) & " "
        If i Mod 2 = 0 Then s = s & vbCrLf
    Next i
    MsgBox s
End Sub

Private Sub ApplyCheckerboard()
    Dim r As Integer, c As Integer, ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Chess")
    For r = 2 To 9
        For c = 2 To 9
            If (r + c) Mod 2 <> 0 Then ws.Cells(r, c).Interior.color = RGB(118, 150, 86)
            If (r + c) Mod 2 = 0 Then ws.Cells(r, c).Interior.color = RGB(238, 238, 210)
        Next c
    Next r
End Sub
