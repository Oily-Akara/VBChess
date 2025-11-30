VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    ' 1. If user selects more than one cell, do nothing
    If Target.count > 1 Then Exit Sub
    
    ' 2. DEFINE ZONES
    Dim isBoard As Boolean
    Dim isMenu As Boolean
    
    ' Check if click is on Board (B2:I9)
    isBoard = Not Intersect(Target, Range("B2:I9")) Is Nothing
    
    ' Check if click is on Menu Buttons (K5 to K10)
    isMenu = Not Intersect(Target, Range("K5:K10")) Is Nothing
    
    ' 3. VALIDATION
    ' If the click is NOT the board AND NOT the menu, stop here.
    If isBoard = False And isMenu = False Then Exit Sub
    
    ' 4. Run the logic
    ' (Make sure HandleClick is Public in your Module)
    HandleClick Target.row, Target.Column
    
    ' 5. Move selection away so you can click the same button again
    Application.EnableEvents = False
    Range("O1").Select ' Moving it to O1 keeps it out of the way
    Application.EnableEvents = True
End Sub
