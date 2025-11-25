Attribute VB_Name = "M_Engine"
' Module: M_Engine
Option Explicit
Public SelectedSq As Integer ' Stores which square clicked first

' --- CONSTANTS ---
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


' The 10x12 Board Array
' Index 21 is Top-Left (A8), Index 98 is Bottom-Right (H1)
Public Board(0 To 119) As Integer

' --- MAIN SETUP SUB ---
Public Sub InitBoard()
    Dim i As Integer
    Dim row As Integer, col As Integer

    ' 1. Fill entire array with OFF_BOARD (The Border)
    For i = 0 To 119
        Board(i) = OFF_BOARD
    Next i

    ' 2. Clear the center play area (Indices 21 to 98, excluding borders)
    For row = 0 To 7
        For col = 1 To 8
            Board(20 + (row * 10) + col) = EMPTY_SQ
        Next col
    Next row

    ' 3. Place the pieces in the Array
    SetupStandardPosition
    
    ' 4. Draw the array to the Excel Sheet
    RenderBoard
    
    MsgBox "Game Reset!"
End Sub

Private Sub SetupStandardPosition()
    Dim col As Integer
    
    ' BLACK PIECES (Top of sheet, indices 21-28)
    Board(21) = B_ROOK: Board(22) = B_KNIGHT: Board(23) = B_BISHOP
    Board(24) = B_QUEEN: Board(25) = B_KING: Board(26) = B_BISHOP
    Board(27) = B_KNIGHT: Board(28) = B_ROOK
    
    ' BLACK PAWNS (Indices 31-38)
    For col = 1 To 8: Board(30 + col) = B_PAWN: Next col

    ' WHITE PIECES (Bottom of sheet, indices 91-98)
    Board(91) = W_ROOK: Board(92) = W_KNIGHT: Board(93) = W_BISHOP
    Board(94) = W_QUEEN: Board(95) = W_KING: Board(96) = W_BISHOP
    Board(97) = W_KNIGHT: Board(98) = W_ROOK
    
    ' WHITE PAWNS (Indices 81-88)
    For col = 1 To 8: Board(80 + col) = W_PAWN: Next col
End Sub

' --- VISUAL RENDERER ---
Public Sub RenderBoard()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Chess")
    
    Application.ScreenUpdating = False
    
    ' Clear old board
    ws.Range("B2:I9").ClearContents
    
    Dim row As Integer, col As Integer
    Dim piece As Integer
    Dim index As Integer
    Dim symbol As String
    
    ' Loop through visual rows 0 to 7 (Top to Bottom)
    For row = 0 To 7
        For col = 0 To 7
            ' Calculate the 10x12 array index
            ' 21 is the top-left corner (A8)
            index = 21 + (row * 10) + col
            
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
            
            ' Draw to Excel (Offset by 2 because board starts at B2)
            With ws.Cells(row + 2, col + 2)
                .Value = symbol
                .Font.Name = "Segoe UI Symbol" ' Good font for chess pieces
                .Font.Size = 24
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
        Next col
    Next row
    
    Application.ScreenUpdating = True
End Sub

' --- CLICK HANDLER ---
Public Sub HandleClick(r As Integer, c As Integer)
    Dim clickedIndex As Integer
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Chess")
    
    ' 1. Convert Excel Row/Col (2-9) to Array Index
    ' Math: 21 is top-left.
    clickedIndex = 21 + ((r - 2) * 10) + (c - 2)
    
    ' 2. DECISION: Select or Move?
    
    ' If we haven't selected anything yet (0)...
    If SelectedSq = 0 Then
        ' Only select if there is a piece there (Not empty)
        If Board(clickedIndex) <> 0 Then
            SelectedSq = clickedIndex
            ' Color the cell Yellow to show selection
            ws.Cells(r, c).Interior.Color = vbYellow
        End If
        
    ' If we ALREADY selected a piece...
    Else
        ' 3. MOVE THE PIECE
        ' Copy piece from Old Square to New Square
        Board(clickedIndex) = Board(SelectedSq)
        ' Delete piece from Old Square
        Board(SelectedSq) = 0
        
        ' 4. CLEAN UP
        SelectedSq = 0 ' Forget selection
        
        ' Reset Colors (Green/Cream)
        Dim i As Integer, j As Integer
        For i = 2 To 9
            For j = 2 To 9
                If (i + j) Mod 2 <> 0 Then
                    ws.Cells(i, j).Interior.Color = RGB(118, 150, 86)
                Else
                    ws.Cells(i, j).Interior.Color = RGB(238, 238, 210)
                End If
            Next j
        Next i
        
        ' Redraw the pieces
        RenderBoard
    End If
End Sub
