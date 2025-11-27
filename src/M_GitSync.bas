Attribute VB_Name = "M_GitSync"
' Module: M_GitSync
Option Explicit

Public Sub ExportSourceCode()
    Dim component As Object
    Dim exportPath As String
    Dim ext As String
    Dim wb As Workbook
    
    Set wb = ThisWorkbook
    ' Create a "src" folder in the same directory as the Excel file
    exportPath = wb.Path & "\src\"
    
    If dir(exportPath, vbDirectory) = "" Then
        MkDir exportPath
    End If
    
    ' Loop through all VBA components (Modules, Forms, Classes)
    For Each component In wb.VBProject.VBComponents
        Select Case component.Type
            Case 1: ext = ".bas" ' Standard Module
            Case 2: ext = ".cls" ' Class Module
            Case 3: ext = ".frm" ' UserForm
            Case 100: ext = ".cls" ' Worksheet/Workbook code
            Case Else: ext = ".txt"
        End Select
        
        ' Export the component to the src folder
        component.Export exportPath & component.Name & ext
    Next component
    
    MsgBox "Source code exported to " & exportPath
End Sub
