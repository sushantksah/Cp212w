Attribute VB_Name = "Module1"
Sub Assignment1()
Attribute Assignment1.VB_Description = "CP212: Assignment 1 - Sushant Sah"
Attribute Assignment1.VB_ProcData.VB_Invoke_Func = "m\n14"
'
' Assignment1 Macro
' CP212: Assignment 1 - Sushant Sah
'
' Keyboard Shortcut: Ctrl+m
'
    Sheets.Add After:=ActiveSheet
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Regional Report"
    Range("A1").Select
    Selection.Style = "Title"
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "Name"
    Range("B3").Select
    ActiveCell.FormulaR1C1 = "District"
    Range("C3").Select
    ActiveCell.FormulaR1C1 = "Sales Total"
    Range("A10").Select
    ActiveCell.FormulaR1C1 = "Done!"
    Range("A10").Select
    Selection.Copy
    Application.CutCopyMode = False
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    Sheets.Add After:=ActiveSheet
    Sheets("Sheet1").Select
    ActiveSheet.Unprotect
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Regional Report"
    Range("A1").Select
End Sub
