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
    ' 1. If click is outside the board, do nothing
    If Intersect(Target, Range("B2:I9")) Is Nothing Then Exit Sub
    
    ' 2. If user selects more than one cell, do nothing
    If Target.Count > 1 Then Exit Sub
    
    ' 3. Run the logic in M_Engine
    HandleClick Target.row, Target.Column
    
    ' 4. Move selection to A1 so the blue highlight disappears
    Application.EnableEvents = False
    Range("A1").Select
    Application.EnableEvents = True
End Sub
