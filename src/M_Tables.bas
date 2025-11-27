Attribute VB_Name = "M_Tables"
' --- MISSING VARIABLES ---
Public KnightTable(0 To 119) As Integer

' --- MISSING SUB: INIT TABLES ---
Public Sub InitTables()
    Dim i As Integer, r As Integer, c As Integer
    
    ' Initialize Knight Position Table (10x12 Mailbox format)
    ' Higher numbers = Center of board. Lower = Edges.
    ' We fill everything with 0 first
    For i = 0 To 119: KnightTable(i) = 0: Next i
    
    ' Define a 8x8 table for the visual board
    Dim tempTable(0 To 7, 0 To 7) As Integer
    
    ' Fill 8x8 logic (Standard Piece-Square Table logic)
    ' -50 for corners, 20 for center
    tempTable(0, 0) = -50: tempTable(0, 1) = -40: tempTable(0, 2) = -30: tempTable(0, 3) = -30
    tempTable(0, 4) = -30: tempTable(0, 5) = -30: tempTable(0, 6) = -40: tempTable(0, 7) = -50
    
    tempTable(1, 0) = -40: tempTable(1, 1) = -20: tempTable(1, 2) = 0: tempTable(1, 3) = 0
    tempTable(1, 4) = 0: tempTable(1, 5) = 0: tempTable(1, 6) = -20: tempTable(1, 7) = -40
    
    tempTable(2, 0) = -30: tempTable(2, 1) = 0: tempTable(2, 2) = 10: tempTable(2, 3) = 15
    tempTable(2, 4) = 15: tempTable(2, 5) = 10: tempTable(2, 6) = 0: tempTable(2, 7) = -30
    
    tempTable(3, 0) = -30: tempTable(3, 1) = 5: tempTable(3, 2) = 15: tempTable(3, 3) = 20
    tempTable(3, 4) = 20: tempTable(3, 5) = 15: tempTable(3, 6) = 5: tempTable(3, 7) = -30
    
    tempTable(4, 0) = -30: tempTable(4, 1) = 0: tempTable(4, 2) = 15: tempTable(4, 3) = 20
    tempTable(4, 4) = 20: tempTable(4, 5) = 15: tempTable(4, 6) = 0: tempTable(4, 7) = -30
    
    tempTable(5, 0) = -30: tempTable(5, 1) = 0: tempTable(5, 2) = 10: tempTable(5, 3) = 15
    tempTable(5, 4) = 15: tempTable(5, 5) = 10: tempTable(5, 6) = 0: tempTable(5, 7) = -30
    
    tempTable(6, 0) = -40: tempTable(6, 1) = -20: tempTable(6, 2) = 0: tempTable(6, 3) = 5
    tempTable(6, 4) = 5: tempTable(6, 5) = 0: tempTable(6, 6) = -20: tempTable(6, 7) = -40
    
    tempTable(7, 0) = -50: tempTable(7, 1) = -40: tempTable(7, 2) = -30: tempTable(7, 3) = -30
    tempTable(7, 4) = -30: tempTable(7, 5) = -30: tempTable(7, 6) = -40: tempTable(7, 7) = -50

    ' Map 8x8 temp table to the 10x12 KnightTable
    For r = 0 To 7
        For c = 0 To 7
            ' Convert Row/Col to 10x12 Index (Start at 21)
            ' 21 is Row 0, Col 0
            i = 21 + (r * 10) + c
            KnightTable(i) = tempTable(r, c)
        Next c
    Next r
End Sub

