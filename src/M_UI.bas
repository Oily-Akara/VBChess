Attribute VB_Name = "M_UI"
' Module: M_UI
' Handles the Graphical Menus, Turn Tracking, and Notation Panels
Option Explicit

' ==========================================
'  MAIN MENU OVERLAY
' ==========================================
Public Sub ShowFancyMenu()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Chess")
    
    RemoveFancyMenu
    RemoveGameOverMenu
    
    ' 1. Main Background Overlay
    Dim bg As Shape
    Set bg = ws.Shapes.AddShape(msoShapeRectangle, ws.Range("B2").Left, ws.Range("B2").Top, ws.Range("B2:I9").Width, ws.Range("B2:I9").Height)
    bg.Name = "MenuBackground"
    bg.Fill.ForeColor.RGB = RGB(25, 25, 30)
    bg.Fill.Transparency = 0.15
    bg.Line.Visible = msoFalse
    
    ' 2. Title Text
    Dim title As Shape
    Set title = ws.Shapes.AddShape(msoShapeRectangle, ws.Range("C3").Left, ws.Range("C3").Top, ws.Range("C3:H4").Width, 50)
    title.Name = "MenuTitle"
    title.Fill.Visible = msoFalse
    title.Line.Visible = msoFalse
    With title.TextFrame2.TextRange
        .Text = ChrW(&H265A) & " VBCHESS " & ChrW(&H265A)
        .Font.Name = "Arial Black"
        .Font.Size = 34
        .Font.Fill.ForeColor.RGB = RGB(255, 215, 0) ' Gold Text
        .ParagraphFormat.Alignment = msoAlignCenter
    End With
    title.Shadow.Type = msoShadow21
    
    ' 3. Buttons
    CreateMenuButton ws, "BtnPlayWhite", "Play as White vs AI", ws.Range("D5").Top, RGB(70, 130, 180), "Menu_PlayWhite"
    CreateMenuButton ws, "BtnPlayBlack", "Play as Black vs AI", ws.Range("D6").Top + 5, RGB(178, 34, 34), "Menu_PlayBlack"
    CreateMenuButton ws, "BtnTwoPlayer", "Pass & Play (2 Player)", ws.Range("D7").Top + 10, RGB(46, 139, 87), "Menu_TwoPlayer"
End Sub

Public Sub RemoveFancyMenu()
    Dim ws As Worksheet, shp As Shape
    Set ws = ThisWorkbook.Sheets("Chess")
    For Each shp In ws.Shapes
        If Left(shp.Name, 4) = "Menu" Or Left(shp.Name, 3) = "Btn" Then shp.Delete
    Next shp
End Sub

' ==========================================
'  GAME OVER / CHECKMATE OVERLAY
' ==========================================
Public Sub ShowGameOverMenu(MainText As String, SubText As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Chess")
    
    RemoveGameOverMenu
    
    ' Dark out the board
    Dim bg As Shape
    Set bg = ws.Shapes.AddShape(msoShapeRectangle, ws.Range("B2").Left, ws.Range("B2").Top, ws.Range("B2:I9").Width, ws.Range("B2:I9").Height)
    bg.Name = "GO_Background"
    bg.Fill.ForeColor.RGB = RGB(10, 10, 10)
    bg.Fill.Transparency = 0.4
    bg.Line.Visible = msoFalse
    
    ' Status Box
    Dim box As Shape
    Set box = ws.Shapes.AddShape(msoShapeRoundedRectangle, ws.Range("C4").Left, ws.Range("C4").Top, ws.Range("C4:H6").Width, ws.Range("C4:H6").Height)
    box.Name = "GO_Box"
    box.Fill.ForeColor.RGB = RGB(40, 40, 45)
    box.Line.ForeColor.RGB = RGB(255, 215, 0)
    box.Line.Weight = 3
    box.Adjustments.Item(1) = 0.1
    box.Shadow.Type = msoShadow21
    
    ' Main Title (e.g., CHECKMATE!)
    Dim title As Shape
    Set title = ws.Shapes.AddShape(msoShapeRectangle, ws.Range("C4").Left, ws.Range("C4").Top + 15, ws.Range("C4:H4").Width, 40)
    title.Name = "GO_Title"
    title.Fill.Visible = msoFalse
    title.Line.Visible = msoFalse
    With title.TextFrame2.TextRange
        .Text = MainText
        .Font.Name = "Arial Black"
        .Font.Size = 28
        .Font.Fill.ForeColor.RGB = RGB(255, 80, 80) ' Red Alert
        .ParagraphFormat.Alignment = msoAlignCenter
    End With
    
    ' Sub Title (e.g., White Wins!)
    Dim subtitle As Shape
    Set subtitle = ws.Shapes.AddShape(msoShapeRectangle, ws.Range("C5").Left, ws.Range("C5").Top, ws.Range("C4:H4").Width, 30)
    subtitle.Name = "GO_Sub"
    subtitle.Fill.Visible = msoFalse
    subtitle.Line.Visible = msoFalse
    With subtitle.TextFrame2.TextRange
        .Text = SubText
        .Font.Name = "Segoe UI"
        .Font.Size = 18
        .Font.Fill.ForeColor.RGB = RGB(240, 240, 240)
        .ParagraphFormat.Alignment = msoAlignCenter
    End With
    
    ' Buttons
    CreateMenuButton ws, "Btn_PlayAgain", "Play Again", ws.Range("D6").Top - 10, RGB(46, 139, 87), "Action_PlayAgain"
    CreateMenuButton ws, "Btn_MainMenu", "Main Menu", ws.Range("F6").Top - 10, RGB(70, 130, 180), "Action_MainMenu"
    
    ' Fix Button Widths for Game Over screen
    ws.Shapes("Btn_PlayAgain").Width = ws.Range("D1:E1").Width - 10
    ws.Shapes("Btn_MainMenu").Width = ws.Range("F1:G1").Width - 10
    ws.Shapes("Btn_MainMenu").Left = ws.Range("F1").Left + 5
End Sub

Public Sub RemoveGameOverMenu()
    Dim ws As Worksheet, shp As Shape
    Set ws = ThisWorkbook.Sheets("Chess")
    For Each shp In ws.Shapes
        If Left(shp.Name, 3) = "GO_" Or Left(shp.Name, 4) = "Btn_" Then shp.Delete
    Next shp
End Sub

' ==========================================
'  TURN INDICATOR & NOTATION PANEL
' ==========================================
Public Sub UpdateTurnUI()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Chess")
    
    On Error Resume Next
    Dim turnShp As Shape
    Set turnShp = ws.Shapes("TurnIndicator")
    On Error GoTo 0
    
    ' Create shape if it doesn't exist
    If turnShp Is Nothing Then
        Set turnShp = ws.Shapes.AddShape(msoShapeRoundedRectangle, ws.Range("L2").Left, ws.Range("L2").Top, ws.Range("L2:N2").Width, 35)
        turnShp.Name = "TurnIndicator"
        turnShp.Line.Visible = msoFalse
        turnShp.Shadow.Type = msoShadow21
    Else
        ' Ensure the shape matches the new column widths dynamically
        turnShp.Left = ws.Range("L2").Left
        turnShp.Top = ws.Range("L2").Top
        turnShp.Width = ws.Range("L2:N2").Width
        turnShp.Height = 35
    End If
    
    ' Style based on Turn
    With turnShp.TextFrame2.TextRange
        .Font.Name = "Segoe UI"
        .Font.Size = 13 ' Reduced size slightly so it fits inside the box
        .Font.Bold = msoTrue
        .ParagraphFormat.Alignment = msoAlignCenter
        
        If Turn = 1 Then
            .Text = "WHITE TO MOVE"
            turnShp.Fill.ForeColor.RGB = RGB(245, 245, 245) ' White
            .Font.Fill.ForeColor.RGB = RGB(30, 30, 30) ' Dark Text
        Else
            .Text = "BLACK TO MOVE"
            turnShp.Fill.ForeColor.RGB = RGB(35, 35, 40) ' Black
            .Font.Fill.ForeColor.RGB = RGB(245, 245, 245) ' Light Text
        End If
    End With
End Sub

' ==========================================
'  TURN INDICATOR & NOTATION PANEL
' ==========================================
Public Sub SetupSidePanel()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Chess")
    
    ' 1. Clear everything on the right side
    ws.Range("J1:Z100").Clear
    
    ' Clear any old generated side buttons so they don't stack up
    Dim shp As Shape
    For Each shp In ws.Shapes
        If shp.Name = "Btn_SideMainMenu" Then shp.Delete
    Next shp
    
    ' 2. Setup precise column widths
    ws.Columns("J:K").ColumnWidth = 2   ' Spacers
    ws.Columns("L").ColumnWidth = 5     ' Move Number (#)
    ws.Columns("M").ColumnWidth = 10    ' White Move
    ws.Columns("N").ColumnWidth = 10    ' Black Move
    
    ' 3. DRAW A PERMANENT "MAIN MENU" BUTTON
    ' Now you can safely delete the old gray form button!
    CreateSideButton ws, "Btn_SideMainMenu", "Return to Main Menu", ws.Range("L4").Top, RGB(70, 130, 180), "Action_MainMenu"
    
    ' 4. Setup Notation Table Headers (Shifted down to row 6)
    ws.Range("L6:N6").Merge
    ws.Range("L6").Value = "MATCH HISTORY"
    ws.Range("L6").Font.Name = "Segoe UI"
    ws.Range("L6").Font.Size = 10
    ws.Range("L6").Font.Bold = True
    ws.Range("L6").HorizontalAlignment = xlCenter
    ws.Range("L6").Interior.color = RGB(40, 40, 50)
    ws.Range("L6").Font.color = RGB(255, 255, 255)
    
    ws.Range("L7").Value = "#"
    ws.Range("M7").Value = "White"
    ws.Range("N7").Value = "Black"
    
    With ws.Range("L7:N7")
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.color = RGB(220, 220, 225)
        .Borders(xlEdgeBottom).Weight = xlThin
    End With
End Sub

Public Sub LogMoveToUI(moveText As String, MoveNumber As Integer)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Chess")
    
    Dim rowNum As Integer
    ' Starts putting moves on row 8 (below the headers)
    rowNum = 7 + ((MoveNumber + 1) \ 2)
    
    ' Print Row Number
    ws.Cells(rowNum, "L").Value = ((MoveNumber + 1) \ 2) & "."
    ws.Cells(rowNum, "L").Font.color = RGB(150, 150, 150)
    ws.Cells(rowNum, "L").Font.Name = "Segoe UI"
    ws.Cells(rowNum, "L").HorizontalAlignment = xlCenter
    
    ' Uses MoveNumber to figure out if it was White (Odd) or Black (Even)
    If MoveNumber Mod 2 <> 0 Then
        ws.Cells(rowNum, "M").Value = moveText
        ws.Cells(rowNum, "M").Font.Name = "Segoe UI"
        ws.Cells(rowNum, "M").HorizontalAlignment = xlCenter
    Else
        ws.Cells(rowNum, "N").Value = moveText
        ws.Cells(rowNum, "N").Font.Name = "Segoe UI"
        ws.Cells(rowNum, "N").HorizontalAlignment = xlCenter
    End If
End Sub

' Helper for the smaller side-panel button
Private Sub CreateSideButton(ws As Worksheet, btnName As String, btnText As String, topPos As Single, btnColor As Long, macroName As String)
    Dim btn As Shape
    Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, ws.Range("L4").Left, topPos, ws.Range("L4:N4").Width, 25)
    btn.Name = btnName
    btn.Fill.ForeColor.RGB = btnColor
    btn.Line.Visible = msoFalse
    btn.Shadow.Type = msoShadow21
    With btn.TextFrame2.TextRange
        .Text = btnText
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Font.Bold = msoTrue
        .Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .ParagraphFormat.Alignment = msoAlignCenter
    End With
    btn.OnAction = macroName
End Sub



' ==========================================
'  HELPER: DRAW BUTTON
' ==========================================
Private Sub CreateMenuButton(ws As Worksheet, btnName As String, btnText As String, topPos As Single, btnColor As Long, macroName As String)
    Dim btn As Shape
    Dim leftPos As Single, btnWidth As Single
    
    leftPos = ws.Range("D1").Left
    btnWidth = ws.Range("D1:G1").Width
    
    Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos, topPos, btnWidth, 35)
    btn.Name = btnName
    btn.Fill.TwoColorGradient msoGradientHorizontal, 1
    btn.Fill.ForeColor.RGB = btnColor
    btn.Fill.BackColor.RGB = RGB(30, 30, 30)
    btn.Line.Visible = msoFalse
    btn.Adjustments.Item(1) = 0.5
    btn.Shadow.Type = msoShadow21
    btn.Shadow.Transparency = 0.6
    
    With btn.TextFrame2.TextRange
        .Text = btnText
        .Font.Name = "Segoe UI"
        .Font.Size = 12
        .Font.Bold = msoTrue
        .Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .ParagraphFormat.Alignment = msoAlignCenter
    End With
    
    btn.OnAction = macroName
End Sub

' ==========================================
'  CLICK HANDLERS / MACROS
' ==========================================
Public Sub Menu_PlayWhite()
    RemoveFancyMenu
    CurrentGameMode = MODE_PVAI
    HumanColor = 1
    InitBoard
End Sub

Public Sub Menu_PlayBlack()
    RemoveFancyMenu
    CurrentGameMode = MODE_PVAI
    HumanColor = 2
    InitBoard
    DoEvents
    MakeComputerMove
End Sub

Public Sub Menu_TwoPlayer()
    RemoveFancyMenu
    CurrentGameMode = MODE_PVP
    HumanColor = 1
    InitBoard
End Sub

Public Sub Action_PlayAgain()
    RemoveGameOverMenu
    InitBoard
    If CurrentGameMode = MODE_PVAI And HumanColor = 2 Then
        DoEvents
        MakeComputerMove
    End If
End Sub

Public Sub Action_MainMenu()
    RemoveGameOverMenu
    InitBoard
    ShowFancyMenu
End Sub
